//! `CellStore`: a compact, lazily materialising replacement for the
//! `Worksheet._cells` dictionary of loaded worksheets.
//!
//! openpyxl keeps one `Cell` object (plus a `StyleArray`, a key tuple and the
//! value object) per cell: ~400 bytes. Following the approach of Rust
//! spreadsheet readers such as calamine, cells read from a file are kept in
//! struct-of-arrays form, grouped by row (16 bytes per cell, no hash index:
//! lookups are binary searches in the row directory and in the row). Formula
//! text is de-duplicated into relative templates (`=A2+B2` and `=A3+B3`
//! share one template). The Python `Cell` object is only created when a cell
//! is accessed, and then cached, so identity semantics are those of a dict
//! (`ws["A1"] is ws["A1"]`).
//!
//! Cells added after loading live in a small insertion-ordered "general"
//! area. Together the store behaves exactly like the dict it replaces:
//! insertion order, overwriting keeps the position, deleting and re-inserting
//! moves the key to the end, KeyError/RuntimeError semantics. Keys that are
//! not plain `(int, int)` tuples switch the store into "dict mode" (all cells
//! are materialised into a real dict which then handles everything).

use std::collections::HashMap;
use std::hash::{BuildHasherDefault, Hasher};

use pyo3::exceptions::{PyKeyError, PyRuntimeError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple, PyType};
use pyo3::{ffi, PyTraverseError, PyVisit};

use crate::dates::Epoch;
use crate::ftemplate::{key_hash, Template};
use crate::sheet::{date_to_py, shared_to_py, CellRec, DType, Shared, Val, DT};
use crate::slots::SlotClass;

thread_local! {
    /// scratch buffer for rendering formulae
    static SCRATCH: std::cell::RefCell<String> = std::cell::RefCell::new(String::with_capacity(256));
}

#[derive(Default)]
pub struct FxHasher {
    hash: u64,
}

impl Hasher for FxHasher {
    fn write(&mut self, bytes: &[u8]) {
        for chunk in bytes.chunks(8) {
            let mut b = [0u8; 8];
            b[..chunk.len()].copy_from_slice(chunk);
            self.write_u64(u64::from_le_bytes(b));
        }
    }
    #[inline]
    fn write_u64(&mut self, i: u64) {
        self.hash = (self.hash.rotate_left(5) ^ i).wrapping_mul(0x51_7c_c1_b7_27_22_0a_95);
    }
    fn finish(&self) -> u64 {
        // murmur3 fmix64: all input bits affect the low bits used by the
        // hash table (cell keys are `row << 32 | column`)
        let mut h = self.hash;
        h ^= h >> 33;
        h = h.wrapping_mul(0xff51_afd7_ed55_8ccd);
        h ^= h >> 33;
        h = h.wrapping_mul(0xc4ce_b9fe_1a85_ec53);
        h ^= h >> 33;
        h
    }
}

pub type FxMap<K, V> = HashMap<K, V, BuildHasherDefault<FxHasher>>;

pub const K_DEAD: u8 = 0;
pub const K_OBJ: u8 = 1;
pub const K_NONE: u8 = 2;
pub const K_INT: u8 = 3;
pub const K_FLOAT: u8 = 4;
pub const K_BOOL: u8 = 5;
pub const K_SST: u8 = 6;
pub const K_TEXT: u8 = 7;
pub const K_SHARED: u8 = 8;
pub const K_SHARED_NODEST: u8 = 9;
pub const K_DATE_INT: u8 = 10;
pub const K_DATE_FLOAT: u8 = 11;
pub const K_PYVAL: u8 = 12;
/// formula from a de-duplicated template (payload = template id)
pub const K_FTPL: u8 = 13;

/// style value meaning `_style = None` (cells created by `Worksheet.cell`)
pub const NO_STYLE: u32 = u32::MAX;

#[inline]
pub fn pack(row: i64, col: i64) -> Option<u64> {
    if (0..=u32::MAX as i64).contains(&row) && (0..=u32::MAX as i64).contains(&col) {
        Some(((row as u64) << 32) | col as u64)
    } else {
        None
    }
}

#[inline]
pub fn unpack(k: u64) -> (i64, i64) {
    ((k >> 32) as i64, (k & 0xffff_ffff) as i64)
}

/// A cell record (general area / transfer form)
#[derive(Clone, Copy)]
pub struct Entry {
    pub key: u64,
    pub style: u32,
    pub kind: u8,
    pub dtype: u8,
    pub payload: u64,
}

/// Packed 10 byte cell record of the base area.
#[repr(C, packed)]
#[derive(Clone, Copy)]
struct Rec {
    col: u16,
    kind: u8,
    dtype: u8,
    /// style index; WIDE_STYLE = look up `Base::wide_styles`
    style: u16,
    /// small value, or index into a side table (see `Base::resolve`)
    payload: u32,
}

const WIDE_STYLE: u16 = u16::MAX;
/// ints stored inline in `payload` (as i32)
const K_INT32: u8 = 100;
const CHUNK_BITS: usize = 16;
const CHUNK: usize = 1 << CHUNK_BITS;

/// Append-only storage in fixed size chunks: growing never reallocates or
/// copies existing data (no 2x capacity peaks).
struct ChunkVec<T: Copy> {
    chunks: Vec<Vec<T>>,
    len: usize,
}

impl<T: Copy> Default for ChunkVec<T> {
    fn default() -> Self {
        ChunkVec {
            chunks: Vec::new(),
            len: 0,
        }
    }
}

impl<T: Copy> ChunkVec<T> {
    #[inline]
    fn push(&mut self, v: T) {
        if self.len & (CHUNK - 1) == 0 {
            self.chunks.push(Vec::with_capacity(CHUNK));
        }
        self.chunks.last_mut().unwrap().push(v);
        self.len += 1;
    }
    #[inline]
    fn get(&self, i: usize) -> T {
        self.chunks[i >> CHUNK_BITS][i & (CHUNK - 1)]
    }
    #[inline]
    fn get_mut(&mut self, i: usize) -> &mut T {
        &mut self.chunks[i >> CHUNK_BITS][i & (CHUNK - 1)]
    }
    fn shrink(&mut self) {
        if let Some(last) = self.chunks.last_mut() {
            last.shrink_to_fit();
        }
        self.chunks.shrink_to_fit();
    }
}

/// Cells loaded from the file, in file order, which must be row-major
/// (strictly increasing rows, strictly increasing columns within a row).
#[derive(Default)]
struct Base {
    rows: Vec<u32>,
    /// start index of each row in `recs` (len = rows.len() + 1)
    starts: Vec<u32>,
    recs: ChunkVec<Rec>,
    /// 64 bit payloads (large ints, floats, dates, text descriptors)
    wide: ChunkVec<u64>,
    /// style ids that do not fit in u16
    wide_styles: HashMap<u32, u32>,
    /// text descriptor -> side slot
    text_slots: FxMap<u64, u32>,
    live: usize,
}

impl Base {
    fn len(&self) -> usize {
        self.recs.len
    }

    fn row_of(&self, i: usize) -> u32 {
        let r = self.starts.partition_point(|&s| (s as usize) <= i) - 1;
        self.rows[r]
    }

    #[inline]
    fn kind(&self, i: usize) -> u8 {
        let k = self.recs.get(i).kind;
        if k == K_INT32 {
            K_INT
        } else {
            k
        }
    }

    #[inline]
    fn col(&self, i: usize) -> u16 {
        self.recs.get(i).col
    }

    fn find(&self, key: u64) -> Option<usize> {
        let (row, col) = unpack(key);
        if col > u16::MAX as i64 {
            return None;
        }
        let r = self.rows.binary_search(&(row as u32)).ok()?;
        let (s, e) = (self.starts[r] as usize, self.starts[r + 1] as usize);
        // binary search on columns within the row
        let (mut lo, mut hi) = (s, e);
        let col = col as u16;
        while lo < hi {
            let mid = (lo + hi) / 2;
            let c = self.recs.get(mid).col;
            if c < col {
                lo = mid + 1;
            } else if c > col {
                hi = mid;
            } else {
                return if self.recs.get(mid).kind == K_DEAD {
                    None
                } else {
                    Some(mid)
                };
            }
        }
        None
    }

    /// Encode an entry's (kind, payload) into the packed form.
    fn encode(&mut self, kind: u8, payload: u64) -> (u8, u32) {
        match kind {
            K_INT => {
                let v = payload as i64;
                if v >= i32::MIN as i64 && v <= i32::MAX as i64 {
                    (K_INT32, v as i32 as u32)
                } else {
                    self.wide.push(payload);
                    (K_INT, (self.wide.len - 1) as u32)
                }
            }
            K_TEXT => {
                // texts are interned: share the side slot too
                if let Some(&slot) = self.text_slots.get(&payload) {
                    return (kind, slot);
                }
                self.wide.push(payload);
                let slot = (self.wide.len - 1) as u32;
                self.text_slots.insert(payload, slot);
                (kind, slot)
            }
            K_FLOAT | K_DATE_INT | K_DATE_FLOAT => {
                self.wide.push(payload);
                (kind, (self.wide.len - 1) as u32)
            }
            _ => (kind, payload as u32),
        }
    }

    /// try to append in order; false if the order would be violated
    fn try_push(&mut self, e: &Entry) -> bool {
        let (row, col) = unpack(e.key);
        if row > u32::MAX as i64 || col > u16::MAX as i64 || self.wide.len >= u32::MAX as usize {
            return false;
        }
        let (row, col) = (row as u32, col as u16);
        match self.rows.last() {
            Some(&last) if last == row => {
                if self.recs.get(self.recs.len - 1).col >= col {
                    return false;
                }
            }
            Some(&last) if last > row => return false,
            _ => {
                self.rows.push(row);
                if self.starts.is_empty() {
                    self.starts.push(0);
                }
                self.starts.push(self.recs.len as u32);
            }
        }
        let (kind, payload) = self.encode(e.kind, e.payload);
        let style = if e.style < WIDE_STYLE as u32 {
            e.style as u16
        } else {
            self.wide_styles.insert(self.recs.len as u32, e.style);
            WIDE_STYLE
        };
        self.recs.push(Rec {
            col,
            kind,
            dtype: e.dtype,
            style,
            payload,
        });
        *self.starts.last_mut().unwrap() = self.recs.len as u32;
        self.live += 1;
        true
    }

    /// Resolve the 64 bit payload of record i.
    #[inline]
    fn payload(&self, i: usize) -> u64 {
        let r = self.recs.get(i);
        match r.kind {
            K_INT32 => (r.payload as i32) as i64 as u64,
            K_INT | K_FLOAT | K_DATE_INT | K_DATE_FLOAT | K_TEXT => {
                self.wide.get(r.payload as usize)
            }
            _ => r.payload as u64,
        }
    }

    fn style(&self, i: usize) -> u32 {
        let s = self.recs.get(i).style;
        if s == WIDE_STYLE {
            self.wide_styles[&(i as u32)]
        } else {
            s as u32
        }
    }

    fn entry(&self, i: usize, row: u32) -> Entry {
        let r = self.recs.get(i);
        Entry {
            key: ((row as u64) << 32) | r.col as u64,
            style: self.style(i),
            kind: self.kind(i),
            dtype: r.dtype,
            payload: self.payload(i),
        }
    }

    /// Replace kind/payload (payload must fit u32 for kinds stored inline:
    /// K_OBJ / K_PYVAL / K_DEAD indexes)
    fn set_kind_payload(&mut self, i: usize, kind: u8, payload: u64) {
        let (k, p) = self.encode(kind, payload);
        let r = self.recs.get_mut(i);
        r.kind = k;
        r.payload = p;
    }

    /// Replace everything but the position (duplicate coordinate in a file)
    fn set_entry(&mut self, i: usize, e: &Entry) {
        let (k, p) = self.encode(e.kind, e.payload);
        let style = if e.style < WIDE_STYLE as u32 {
            e.style as u16
        } else {
            self.wide_styles.insert(i as u32, e.style);
            WIDE_STYLE
        };
        let r = self.recs.get_mut(i);
        r.kind = k;
        r.payload = p;
        r.dtype = e.dtype;
        r.style = style;
    }

    fn kill(&mut self, i: usize) {
        self.recs.get_mut(i).kind = K_DEAD;
        self.live -= 1;
    }

    fn shrink(&mut self) {
        self.rows.shrink_to_fit();
        self.starts.shrink_to_fit();
        self.recs.shrink();
        self.wide.shrink();
        self.text_slots = FxMap::default();
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Loc {
    /// index in the base area and the row of that cell
    Base(usize, u32),
    Gen(usize),
}

/// A value read from a lazy entry, without creating Python objects when
/// possible (used by the writer).
pub enum LazyValue<'a> {
    None,
    Int(i64),
    Float(f64),
    Bool(bool),
    Str(&'a str),
    Owned(String),
    Py(Bound<'a, PyAny>),
}

#[pyclass(subclass, module = "openrsxl._openrsxl")]
pub struct CellStore {
    ws: Py<PyAny>,
    cell_styles: Py<PyAny>,
    shared_strings: Py<PyAny>,
    epoch: Option<Epoch>,
    masters: Vec<Shared>,
    slots: Option<SlotClass>,
    array_new: Option<ffi::newfunc>,
    typecode: Option<Py<PyString>>,
    dt: Option<DT>,
    dtypes_extra: Vec<Py<PyAny>>,
    base: Base,
    /// general area: insertion ordered, tombstones for deleted entries
    gen: Vec<Entry>,
    gen_index: FxMap<u64, u32>,
    gen_live: usize,
    objs: Vec<Option<Py<PyAny>>>,
    free_objs: Vec<u32>,
    arena: Vec<u8>,
    templates: Vec<Template>,
    /// hash of a template key -> template id (u32::MAX: seen once)
    template_ids: FxMap<u64, u32>,
    /// hash of a text -> arena payload (interning, while loading)
    text_ids: FxMap<u64, u64>,
    version: u64,
    dict: Option<Py<PyDict>>,
    loading: bool,
    /// cached `('i', StyleArray)` constructor arguments per style index
    style_args: Vec<Option<Py<PyTuple>>>,
    /// last row int object handed out
    row_obj: Option<(i64, Py<PyAny>)>,
}

impl CellStore {
    fn init_runtime(&mut self, py: Python<'_>, cell_cls: &Bound<'_, PyType>) -> PyResult<()> {
        self.slots = Some(SlotClass::new(
            cell_cls,
            &[
                "parent",
                "_style",
                "row",
                "column",
                "_value",
                "_hyperlink",
                "data_type",
                "_comment",
            ],
        )?);
        let array_type = py.import("array")?.getattr("array")?;
        self.array_new = unsafe { (*(array_type.as_ptr() as *mut ffi::PyTypeObject)).tp_new };
        self.typecode = Some(PyString::new(py, "i").unbind());
        self.dt = Some(DT::new(py));
        Ok(())
    }

    pub fn in_dict_mode(&self) -> bool {
        self.dict.is_some()
    }

    pub fn live(&self) -> usize {
        self.base.live + self.gen_live
    }

    // ------------------------------------------------------------ access

    pub fn entry(&self, loc: Loc) -> Entry {
        match loc {
            Loc::Base(i, row) => self.base.entry(i, row),
            Loc::Gen(i) => self.gen[i],
        }
    }

    fn set_kind_payload(&mut self, loc: Loc, kind: u8, payload: u64) {
        match loc {
            Loc::Base(i, _) => self.base.set_kind_payload(i, kind, payload),
            Loc::Gen(i) => {
                self.gen[i].kind = kind;
                self.gen[i].payload = payload;
            }
        }
    }

    pub fn find_key(&self, key: u64) -> Option<Loc> {
        if let Some(i) = self.base.find(key) {
            return Some(Loc::Base(i, unpack(key).0 as u32));
        }
        self.gen_index.get(&key).map(|&i| Loc::Gen(i as usize))
    }

    /// All live cells in dict (insertion) order.
    pub fn locs(&self) -> Vec<Loc> {
        let mut out = Vec::with_capacity(self.live());
        for ri in 0..self.base.rows.len() {
            let row = self.base.rows[ri];
            for i in self.base.starts[ri] as usize..self.base.starts[ri + 1] as usize {
                if self.base.kind(i) != K_DEAD {
                    out.push(Loc::Base(i, row));
                }
            }
        }
        for (i, e) in self.gen.iter().enumerate() {
            if e.kind != K_DEAD {
                out.push(Loc::Gen(i));
            }
        }
        out
    }

    /// Rows of the base area (sorted) and their record ranges.
    pub fn base_rows(&self) -> (&[u32], &[u32]) {
        (&self.base.rows, &self.base.starts)
    }

    /// Live cell of the base area at index i (with its row), if any.
    #[inline]
    pub fn base_cell(&self, i: usize, row: u32) -> Option<(u64, Loc)> {
        if self.base.kind(i) == K_DEAD {
            None
        } else {
            Some((
                ((row as u64) << 32) | self.base.col(i) as u64,
                Loc::Base(i, row),
            ))
        }
    }

    /// Live cells of the general area sorted by key.
    pub fn gen_sorted(&self) -> Vec<(u64, Loc)> {
        let mut gen: Vec<(u64, Loc)> = self
            .gen
            .iter()
            .enumerate()
            .filter(|(_, e)| e.kind != K_DEAD)
            .map(|(i, e)| (e.key, Loc::Gen(i)))
            .collect();
        gen.sort_unstable_by_key(|x| x.0);
        gen
    }

    fn put_obj(&mut self, o: Py<PyAny>) -> u64 {
        if let Some(i) = self.free_objs.pop() {
            self.objs[i as usize] = Some(o);
            i as u64
        } else {
            self.objs.push(Some(o));
            (self.objs.len() - 1) as u64
        }
    }

    fn drop_payload(&mut self, e: &Entry) {
        if e.kind == K_OBJ || e.kind == K_PYVAL {
            self.objs[e.payload as usize] = None;
            self.free_objs.push(e.payload as u32);
        }
    }

    /// Code of a data type: the standard ones (DT_NAMES) or an interned
    /// extra one (eg. "str"); NativeFallback when the codes run out.
    fn dtype_code(&mut self, py: Python<'_>, d: DType) -> PyResult<u8> {
        match d {
            DType::Code(c) => Ok(c),
            DType::Py(o) => {
                let b = o.bind(py);
                for (i, e) in self.dtypes_extra.iter().enumerate() {
                    if e.bind(py).eq(b)? {
                        return Ok((7 + i) as u8);
                    }
                }
                if self.dtypes_extra.len() >= (u8::MAX as usize) - 7 {
                    return Err(crate::NativeFallback::new_err("too many data types"));
                }
                self.dtypes_extra.push(o);
                Ok((6 + self.dtypes_extra.len()) as u8)
            }
        }
    }

    pub fn dtype_obj(&self, py: Python<'_>, code: u8) -> Py<PyAny> {
        if code <= 6 {
            self.dt.as_ref().unwrap().code(py, code)
        } else {
            self.dtypes_extra[(code - 7) as usize].clone_ref(py)
        }
    }

    pub fn dtype_str(&self, py: Python<'_>, code: u8) -> Option<String> {
        if code <= 6 {
            Some(crate::sheet::DT_NAMES[code as usize].to_string())
        } else {
            self.dtypes_extra[(code - 7) as usize]
                .bind(py)
                .cast::<PyString>()
                .ok()?
                .to_str()
                .ok()
                .map(|s| s.to_string())
        }
    }

    pub fn text(&self, payload: u64) -> &str {
        let off = (payload >> 24) as usize;
        let len = (payload & 0xff_ffff) as usize;
        unsafe { std::str::from_utf8_unchecked(&self.arena[off..off + len]) }
    }

    fn store_text(&mut self, py: Python<'_>, s: &str) -> (u8, u64) {
        // (offset in the top 40 bits of the payload; u64 arithmetic: 32-bit targets)
        if s.len() < (1 << 24) && (self.arena.len() as u64) < (1u64 << 40) {
            // intern repeated texts (verified byte for byte)
            let mut h = FxHasher::default();
            h.write(s.as_bytes());
            h.write_u64(s.len() as u64);
            let hash = h.finish();
            if let Some(&payload) = self.text_ids.get(&hash) {
                if self.text(payload) == s {
                    return (K_TEXT, payload);
                }
            }
            let off = self.arena.len() as u64;
            self.arena.extend_from_slice(s.as_bytes());
            let payload = (off << 24) | s.len() as u64;
            self.text_ids.insert(hash, payload);
            (K_TEXT, payload)
        } else {
            let o = PyString::new(py, s).into_any().unbind();
            (K_PYVAL, self.put_obj(o))
        }
    }

    /// Try to store formula `f` (at row, col) as a de-duplicated template.
    /// A template is created on the second occurrence of a pattern (unique
    /// formulae stay plain text, which is smaller); a template is only used
    /// for a cell if rendering it there reproduces `f` exactly.
    fn template_for(&mut self, f: &str, row: i64, col: i64) -> Option<u32> {
        if f.len() < 4 {
            return None;
        }
        let hash = key_hash(f, row, col)?;
        match self.template_ids.get(&hash) {
            Some(&id) if id != u32::MAX => {
                if self.templates[id as usize].matches(f, row, col) {
                    Some(id)
                } else {
                    None
                }
            }
            Some(_) => {
                let t = Template::build(f, row, col)?;
                if !t.matches(f, row, col) {
                    return None;
                }
                let id = self.templates.len() as u32;
                self.templates.push(t);
                self.template_ids.insert(hash, id);
                Some(id)
            }
            None => {
                self.template_ids.insert(hash, u32::MAX);
                None
            }
        }
    }

    /// Append a parsed cell (from the native reader).
    pub fn push_rec(&mut self, py: Python<'_>, c: CellRec, style: u32) -> PyResult<()> {
        let key = match pack(c.row, c.column) {
            Some(k) => k,
            None => return Err(crate::NativeFallback::new_err("coordinate out of range")),
        };
        let dtype = self.dtype_code(py, c.dtype)?;
        let (kind, payload) = match c.val {
            Val::None => (K_NONE, 0),
            Val::Int(i) => (K_INT, i as u64),
            Val::Float(f) => (K_FLOAT, f.to_bits()),
            Val::Bool(b) => (K_BOOL, b as u64),
            Val::Sst(i) => (K_SST, i as u64),
            Val::Text(s) => {
                let tpl = if dtype == crate::sheet::DT_F {
                    self.template_for(&s, c.row, c.column)
                } else {
                    None
                };
                match tpl {
                    Some(id) => (K_FTPL, id as u64),
                    None => self.store_text(py, &s),
                }
            }
            Val::Shared { master, dest } => {
                if dest.is_some() && dest == Some((c.row, c.column)) {
                    (K_SHARED, master as u64)
                } else if dest.is_none() {
                    (K_SHARED_NODEST, master as u64)
                } else {
                    let o = shared_to_py(py, &self.masters[master as usize], dest)?;
                    (K_PYVAL, self.put_obj(o))
                }
            }
            Val::DateInt(i) => (K_DATE_INT, i as u64),
            Val::DateFloat(f) => (K_DATE_FLOAT, f.to_bits()),
            Val::Py(o) => (K_PYVAL, self.put_obj(o)),
        };
        let e = Entry {
            key,
            style,
            kind,
            dtype,
            payload,
        };
        self.push_entry_loading(e);
        Ok(())
    }

    /// Insert during loading: dict assignment semantics (`_cells[k] = c`).
    fn push_entry_loading(&mut self, e: Entry) {
        if self.gen.is_empty() && self.base.try_push(&e) {
            return;
        }
        // out of order (or duplicate): general area
        match self.find_key(e.key) {
            Some(loc) => {
                // duplicate key: keeps its position, value replaced
                let old = self.entry(loc);
                self.drop_payload(&old);
                match loc {
                    Loc::Base(i, _) => self.base.set_entry(i, &e),
                    Loc::Gen(i) => self.gen[i] = e,
                }
            }
            None => {
                self.gen_index.insert(e.key, self.gen.len() as u32);
                self.gen.push(e);
                self.gen_live += 1;
            }
        }
    }

    pub fn finish_loading(&mut self, masters: Vec<Shared>) {
        self.masters = masters;
        self.base.shrink();
        self.arena.shrink_to_fit();
        self.template_ids = FxMap::default(); // only needed while loading
        self.text_ids = FxMap::default();
        self.templates.shrink_to_fit();
        self.loading = false;
    }

    /// Value of a cell as a Python object (does not materialise the cell)
    pub fn value_obj(&self, py: Python<'_>, loc: Loc) -> PyResult<Py<PyAny>> {
        let e = self.entry(loc);
        Ok(match e.kind {
            K_OBJ => self.objs[e.payload as usize]
                .as_ref()
                .unwrap()
                .bind(py)
                .getattr("value")?
                .unbind(),
            K_NONE => py.None(),
            K_INT => (e.payload as i64).into_pyobject(py)?.into_any().unbind(),
            K_FLOAT => PyFloat::new(py, f64::from_bits(e.payload))
                .into_any()
                .unbind(),
            K_BOOL => PyBool::new(py, e.payload != 0)
                .to_owned()
                .into_any()
                .unbind(),
            K_SST => self
                .shared_strings
                .bind(py)
                .get_item(e.payload as usize)?
                .unbind(),
            K_TEXT => PyString::new(py, self.text(e.payload)).into_any().unbind(),
            K_SHARED | K_SHARED_NODEST | K_FTPL => SCRATCH.with(|b| {
                let mut b = b.borrow_mut();
                b.clear();
                let (r, c) = unpack(e.key);
                match e.kind {
                    K_FTPL => {
                        let ok = self.templates[e.payload as usize].render_into(r, c, &mut b);
                        debug_assert!(ok, "verified");
                    }
                    _ => match &self.masters[e.payload as usize] {
                        Shared::Native(t) => {
                            let dest = if e.kind == K_SHARED {
                                Some((r, c))
                            } else {
                                None
                            };
                            t.translate_into(dest, &mut b)
                                .expect("validated at load time");
                        }
                        Shared::Py(_) => unreachable!(),
                    },
                }
                PyString::new(py, &b).into_any().unbind()
            }),
            K_DATE_INT => date_to_py(
                py,
                crate::dates::from_excel_int(e.payload as i64, self.epoch.as_ref().unwrap()),
            )?,
            K_DATE_FLOAT => date_to_py(
                py,
                crate::dates::from_excel_float(
                    f64::from_bits(e.payload),
                    self.epoch.as_ref().unwrap(),
                ),
            )?,
            K_PYVAL => self.objs[e.payload as usize]
                .as_ref()
                .unwrap()
                .clone_ref(py),
            _ => unreachable!(),
        })
    }

    /// Lazy value without Python objects where possible (writer).
    pub fn lazy_value<'a>(&'a self, py: Python<'a>, loc: Loc) -> PyResult<LazyValue<'a>> {
        self.lazy_value_e(py, loc, self.entry(loc))
    }

    /// `lazy_value` with the entry already at hand.
    pub fn lazy_value_e<'a>(
        &'a self,
        py: Python<'a>,
        loc: Loc,
        e: Entry,
    ) -> PyResult<LazyValue<'a>> {
        Ok(match e.kind {
            K_NONE => LazyValue::None,
            K_INT => LazyValue::Int(e.payload as i64),
            K_FLOAT => LazyValue::Float(f64::from_bits(e.payload)),
            K_BOOL => LazyValue::Bool(e.payload != 0),
            K_TEXT => LazyValue::Str(self.text(e.payload)),
            K_SHARED | K_SHARED_NODEST => {
                let dest = if e.kind == K_SHARED {
                    Some(unpack(e.key))
                } else {
                    None
                };
                match &self.masters[e.payload as usize] {
                    Shared::Native(t) => LazyValue::Owned(t.translate(dest).expect("validated")),
                    Shared::Py(_) => unreachable!(),
                }
            }
            K_FTPL => {
                let (r, c) = unpack(e.key);
                LazyValue::Owned(
                    self.templates[e.payload as usize]
                        .render(r, c)
                        .expect("verified"),
                )
            }
            _ => LazyValue::Py(self.value_obj(py, loc)?.into_bound(py)),
        })
    }

    /// The from_excel result of a native date cell.
    pub fn excel_date(&self, loc: Loc) -> Option<crate::dates::ExcelDate> {
        let e = self.entry(loc);
        let ep = self.epoch.as_ref()?;
        match e.kind {
            K_DATE_INT => Some(crate::dates::from_excel_int(e.payload as i64, ep)),
            K_DATE_FLOAT => Some(crate::dates::from_excel_float(
                f64::from_bits(e.payload),
                ep,
            )),
            _ => None,
        }
    }

    pub fn obj_payload(&self, payload: u64) -> Option<&Py<PyAny>> {
        self.objs.get(payload as usize).and_then(|o| o.as_ref())
    }

    pub fn obj_at(&self, loc: Loc) -> Option<&Py<PyAny>> {
        let e = self.entry(loc);
        if e.kind == K_OBJ {
            self.objs[e.payload as usize].as_ref()
        } else {
            None
        }
    }

    pub fn cell_styles<'py>(&self, py: Python<'py>) -> Bound<'py, PyAny> {
        self.cell_styles.bind(py).clone()
    }

    /// Create (and cache) the Cell object for `loc`.
    pub fn materialize(&mut self, py: Python<'_>, loc: Loc) -> PyResult<Py<PyAny>> {
        let e = self.entry(loc);
        if e.kind == K_OBJ {
            return Ok(self.objs[e.payload as usize]
                .as_ref()
                .unwrap()
                .clone_ref(py));
        }
        let value = self.value_obj(py, loc)?;
        let (row, col) = unpack(e.key);
        let style_copy: Py<PyAny> = if e.style == NO_STYLE {
            py.None()
        } else {
            // StyleArray(style_array) == array.__new__(StyleArray, 'i', style);
            // the argument tuple is cached per style index
            let idx = e.style as usize;
            if idx >= self.style_args.len() {
                self.style_args.resize_with(idx + 1, || None);
            }
            if self.style_args[idx].is_none() {
                let styles = self.cell_styles.bind(py);
                let style = match styles.cast::<PyList>() {
                    Ok(l) if idx < l.len() => l.get_item(idx)?,
                    _ => styles.get_item(idx)?,
                };
                self.style_args[idx] = Some(
                    PyTuple::new(
                        py,
                        [self.typecode.as_ref().unwrap().bind(py).as_any(), &style],
                    )?
                    .unbind(),
                );
            }
            let args = self.style_args[idx].as_ref().unwrap().bind(py);
            let style = args.get_item(1)?;
            unsafe {
                let o = (self.array_new.unwrap())(
                    ffi::Py_TYPE(style.as_ptr()),
                    args.as_ptr(),
                    std::ptr::null_mut(),
                );
                if o.is_null() {
                    return Err(PyErr::fetch(py));
                }
                Bound::from_owned_ptr(py, o).unbind()
            }
        };
        let dtype = self.dtype_obj(py, e.dtype);
        let none = py.None();
        // one int object per row number, as openpyxl's binding shares them
        let row_obj = match &self.row_obj {
            Some((r, o)) if *r == row => o.clone_ref(py),
            _ => {
                let o = row.into_pyobject(py)?.into_any().unbind();
                self.row_obj = Some((row, o.clone_ref(py)));
                o
            }
        };
        let vals = unsafe {
            [
                self.ws.clone_ref(py).into_ptr(),
                style_copy.into_ptr(),
                row_obj.into_ptr(),
                ffi::PyLong_FromLongLong(col),
                value.into_ptr(),
                none.clone_ref(py).into_ptr(),
                dtype.into_ptr(),
                none.into_ptr(),
            ]
        };
        let cell = self.slots.as_ref().unwrap().create(py, &vals)?.unbind();
        if e.kind == K_PYVAL {
            self.objs[e.payload as usize] = None;
            self.free_objs.push(e.payload as u32);
        }
        let slot = self.put_obj(cell.clone_ref(py));
        self.set_kind_payload(loc, K_OBJ, slot);
        Ok(cell)
    }

    fn exact_key(key: &Bound<'_, PyAny>) -> Option<u64> {
        if !key.is_exact_instance_of::<PyTuple>() {
            return None;
        }
        let t = key.cast::<PyTuple>().ok()?;
        if t.len() != 2 {
            return None;
        }
        let a = t.get_item(0).ok()?;
        let b = t.get_item(1).ok()?;
        if !a.is_exact_instance_of::<PyInt>() || !b.is_exact_instance_of::<PyInt>() {
            return None;
        }
        pack(a.extract::<i64>().ok()?, b.extract::<i64>().ok()?)
    }

    /// Normalise keys equal to an integer pair (eg. (1.0, True)) as a dict
    /// lookup would.
    fn equivalent_key(py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Option<u64>> {
        let Ok(t) = key.cast::<PyTuple>() else {
            return Ok(None);
        };
        if t.len() != 2 {
            return Ok(None);
        }
        let int = py.import("builtins")?.getattr("int")?;
        let mut out = [0i64; 2];
        for (i, slot) in out.iter_mut().enumerate() {
            let el = t.get_item(i)?;
            let as_int = match int.call1((&el,)) {
                Ok(v) => v,
                Err(_) => return Ok(None),
            };
            if !el.eq(&as_int).unwrap_or(false) || el.hash().ok() != as_int.hash().ok() {
                return Ok(None);
            }
            match as_int.extract::<i64>() {
                Ok(v) => *slot = v,
                Err(_) => return Ok(None),
            }
        }
        Ok(pack(out[0], out[1]))
    }

    fn find(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Option<Loc>> {
        let k = match Self::exact_key(key) {
            Some(k) => Some(k),
            None => Self::equivalent_key(py, key)?,
        };
        Ok(k.and_then(|k| self.find_key(k)))
    }

    fn enter_dict_mode(&mut self, py: Python<'_>) -> PyResult<()> {
        if self.dict.is_some() {
            return Ok(());
        }
        let d = PyDict::new(py);
        for loc in self.locs() {
            let obj = self.materialize(py, loc)?;
            let (r, c) = unpack(self.entry(loc).key);
            d.set_item(PyTuple::new(py, [r, c])?, obj)?;
        }
        self.reset();
        self.dict = Some(d.unbind());
        Ok(())
    }

    fn reset(&mut self) {
        self.base = Base::default();
        self.gen = Vec::new();
        self.gen_index = FxMap::default();
        self.gen_live = 0;
        self.objs = Vec::new();
        self.free_objs = Vec::new();
        self.arena = Vec::new();
        self.version += 1;
    }

    fn compact_gen(&mut self) {
        let dead = self.gen.len() - self.gen_live;
        if dead < 1024 || dead < self.gen_live {
            return;
        }
        let mut gen = Vec::with_capacity(self.gen_live);
        let mut index = FxMap::default();
        for e in self.gen.iter() {
            if e.kind != K_DEAD {
                index.insert(e.key, gen.len() as u32);
                gen.push(*e);
            }
        }
        self.gen = gen;
        self.gen_index = index;
    }

    fn insert_new(&mut self, e: Entry) -> Loc {
        self.gen_index.insert(e.key, self.gen.len() as u32);
        self.gen.push(e);
        self.gen_live += 1;
        self.version += 1;
        Loc::Gen(self.gen.len() - 1)
    }

    fn remove(&mut self, loc: Loc) {
        let e = self.entry(loc);
        self.drop_payload(&e);
        match loc {
            Loc::Base(i, _) => {
                self.base.kill(i);
            }
            Loc::Gen(i) => {
                self.gen[i].kind = K_DEAD;
                self.gen_index.remove(&e.key);
                self.gen_live -= 1;
            }
        }
        self.version += 1;
        self.compact_gen();
    }
}

/// Pauses Python's cyclic garbage collector until dropped (restoring the
/// previous state: a GC disabled by the user stays disabled).
struct GcPause {
    was_enabled: bool,
}

impl GcPause {
    fn new() -> Self {
        GcPause {
            was_enabled: unsafe { ffi::PyGC_Disable() } == 1,
        }
    }
}

impl Drop for GcPause {
    fn drop(&mut self) {
        if self.was_enabled {
            unsafe {
                ffi::PyGC_Enable();
            }
        }
    }
}

#[pyclass]
pub struct KeyIter {
    store: Py<CellStore>,
    /// position: base cells first, then general cells
    pos: usize,
    len: usize,
    version: u64,
}

#[pymethods]
impl KeyIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let s = self.store.borrow(py);
        if s.version != self.version {
            if s.live() != self.len {
                return Err(PyRuntimeError::new_err(
                    "dictionary changed size during iteration",
                ));
            }
            return Err(PyRuntimeError::new_err(
                "dictionary keys changed during iteration",
            ));
        }
        let nb = s.base.len();
        while self.pos < nb {
            let i = self.pos;
            self.pos += 1;
            if s.base.kind(i) != K_DEAD {
                let row = s.base.row_of(i) as i64;
                return Ok(Some(
                    PyTuple::new(py, [row, s.base.col(i) as i64])?
                        .into_any()
                        .unbind(),
                ));
            }
        }
        while self.pos - nb < s.gen.len() {
            let e = s.gen[self.pos - nb];
            self.pos += 1;
            if e.kind != K_DEAD {
                let (r, c) = unpack(e.key);
                return Ok(Some(PyTuple::new(py, [r, c])?.into_any().unbind()));
            }
        }
        Ok(None)
    }
}

#[pymethods]
impl CellStore {
    #[new]
    #[pyo3(signature = (*_args, **_kwargs))]
    fn py_new(
        py: Python<'_>,
        _args: &Bound<'_, PyTuple>,
        _kwargs: Option<&Bound<'_, PyDict>>,
    ) -> Self {
        CellStore {
            ws: py.None(),
            cell_styles: py.None(),
            shared_strings: py.None(),
            epoch: None,
            masters: Vec::new(),
            slots: None,
            array_new: None,
            typecode: None,
            dt: None,
            dtypes_extra: Vec::new(),
            base: Base::default(),
            gen: Vec::new(),
            gen_index: FxMap::default(),
            gen_live: 0,
            objs: Vec::new(),
            free_objs: Vec::new(),
            arena: Vec::new(),
            templates: Vec::new(),
            template_ids: FxMap::default(),
            text_ids: FxMap::default(),
            version: 0,
            dict: None,
            loading: false,
            style_args: Vec::new(),
            row_obj: None,
        }
    }

    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        visit.call(&self.ws)?;
        visit.call(&self.cell_styles)?;
        visit.call(&self.shared_strings)?;
        for o in self.objs.iter().flatten() {
            visit.call(o)?;
        }
        if let Some(d) = &self.dict {
            visit.call(d)?;
        }
        for m in &self.masters {
            if let Shared::Py(o) = m {
                visit.call(o)?;
            }
        }
        for o in &self.dtypes_extra {
            visit.call(o)?;
        }
        for o in self.style_args.iter().flatten() {
            visit.call(o)?;
        }
        Ok(())
    }

    fn __clear__(&mut self) {
        self.style_args.clear();
        self.row_obj = None;
        self.objs.clear();
        self.dict = None;
        self.masters.clear();
        self.dtypes_extra.clear();
        self.base = Base::default();
        self.gen.clear();
        self.gen_index.clear();
        self.gen_live = 0;
    }

    fn __len__(&self, py: Python<'_>) -> usize {
        match &self.dict {
            Some(d) => d.bind(py).len(),
            None => self.live(),
        }
    }

    fn __contains__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<bool> {
        if let Some(d) = &self.dict {
            return d.bind(py).contains(key);
        }
        if let Some(k) = Self::exact_key(key) {
            return Ok(self.find_key(k).is_some());
        }
        key.hash()?;
        Ok(self.find(py, key)?.is_some())
    }

    fn __getitem__(&mut self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        if let Some(d) = &self.dict {
            return match d.bind(py).get_item(key)? {
                Some(v) => Ok(v.unbind()),
                None => Err(PyKeyError::new_err((key.clone().unbind(),))),
            };
        }
        let loc = match Self::exact_key(key) {
            Some(k) => self.find_key(k),
            None => {
                key.hash()?;
                self.find(py, key)?
            }
        };
        match loc {
            Some(l) => self.materialize(py, l),
            None => Err(PyKeyError::new_err((key.clone().unbind(),))),
        }
    }

    #[pyo3(signature = (key, default=None))]
    fn get(
        &mut self,
        py: Python<'_>,
        key: &Bound<'_, PyAny>,
        default: Option<Py<PyAny>>,
    ) -> PyResult<Py<PyAny>> {
        if let Some(d) = &self.dict {
            return Ok(match d.bind(py).get_item(key)? {
                Some(v) => v.unbind(),
                None => default.unwrap_or_else(|| py.None()),
            });
        }
        let loc = match Self::exact_key(key) {
            Some(k) => self.find_key(k),
            None => {
                key.hash()?;
                self.find(py, key)?
            }
        };
        match loc {
            Some(l) => self.materialize(py, l),
            None => Ok(default.unwrap_or_else(|| py.None())),
        }
    }

    fn __setitem__(
        &mut self,
        py: Python<'_>,
        key: &Bound<'_, PyAny>,
        value: Py<PyAny>,
    ) -> PyResult<()> {
        if let Some(d) = &self.dict {
            return d.bind(py).set_item(key, value);
        }
        key.hash()?;
        let exact = Self::exact_key(key);
        let k = match exact {
            Some(k) => Some(k),
            None => Self::equivalent_key(py, key)?,
        };
        if let Some(k) = k {
            if let Some(loc) = self.find_key(k) {
                // existing key: keeps its position (and its key object)
                let old = self.entry(loc);
                self.drop_payload(&old);
                let slot = self.put_obj(value);
                self.set_kind_payload(loc, K_OBJ, slot);
                return Ok(());
            }
            if exact.is_some() {
                let slot = self.put_obj(value);
                self.insert_new(Entry {
                    key: k,
                    style: NO_STYLE,
                    kind: K_OBJ,
                    dtype: 0,
                    payload: slot,
                });
                return Ok(());
            }
        }
        // a new key that is not a plain (int, int) pair
        self.enter_dict_mode(py)?;
        self.dict.as_ref().unwrap().bind(py).set_item(key, value)
    }

    fn __delitem__(&mut self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<()> {
        if let Some(d) = &self.dict {
            let d = d.bind(py);
            if !d.contains(key)? {
                return Err(PyKeyError::new_err((key.clone().unbind(),)));
            }
            return d.del_item(key);
        }
        key.hash()?;
        match self.find(py, key)? {
            Some(loc) => {
                self.remove(loc);
                Ok(())
            }
            None => Err(PyKeyError::new_err((key.clone().unbind(),))),
        }
    }

    fn __iter__(slf: PyRef<'_, Self>, py: Python<'_>) -> PyResult<Py<PyAny>> {
        if let Some(d) = &slf.dict {
            return Ok(d.bind(py).try_iter()?.into_any().unbind());
        }
        let len = slf.live();
        let version = slf.version;
        let store: Py<CellStore> = slf.into();
        Ok(Py::new(
            py,
            KeyIter {
                store,
                pos: 0,
                len,
                version,
            },
        )?
        .into_any())
    }

    fn __reversed__(&mut self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        if let Some(d) = &self.dict {
            return Ok(py
                .import("builtins")?
                .getattr("reversed")?
                .call1((d.bind(py),))?
                .unbind());
        }
        let l = PyList::empty(py);
        for loc in self.locs().into_iter().rev() {
            let (r, c) = unpack(self.entry(loc).key);
            l.append(PyTuple::new(py, [r, c])?)?;
        }
        Ok(l.try_iter()?.into_any().unbind())
    }

    fn clear(&mut self) {
        self.reset();
        self.dict = None;
    }

    /// dict.popitem(): remove and return the last inserted (key, value)
    fn popitem(&mut self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        if let Some(d) = &self.dict {
            return Ok(d.bind(py).call_method0("popitem")?.unbind());
        }
        let Some(&loc) = self.locs().last() else {
            return Err(PyKeyError::new_err("popitem(): dictionary is empty"));
        };
        let v = self.materialize(py, loc)?;
        let (r, c) = unpack(self.entry(loc).key);
        self.remove(loc);
        Ok(
            PyTuple::new(py, [PyTuple::new(py, [r, c])?.into_any().unbind(), v])?
                .into_any()
                .unbind(),
        )
    }

    // ---------------------------------------------------------------
    // fast paths used by openrsxl's Worksheet

    /// (min_row, max_row, min_col, max_col) of the keys, or None if empty or
    /// in dict mode.
    fn _bounds(&self) -> Option<(i64, i64, i64, i64)> {
        if self.dict.is_some() {
            return None;
        }
        let mut b: Option<(i64, i64, i64, i64)> = None;
        let mut upd = |r: i64, c: i64| {
            b = Some(match b {
                None => (r, r, c, c),
                Some((a, bb, cc, d)) => (a.min(r), bb.max(r), cc.min(c), d.max(c)),
            });
        };
        for ri in 0..self.base.rows.len() {
            let row = self.base.rows[ri] as i64;
            for i in self.base.starts[ri] as usize..self.base.starts[ri + 1] as usize {
                if self.base.kind(i) != K_DEAD {
                    upd(row, self.base.col(i) as i64);
                }
            }
        }
        for e in &self.gen {
            if e.kind != K_DEAD {
                let (r, c) = unpack(e.key);
                upd(r, c);
            }
        }
        b
    }

    fn _in_dict_mode(&self) -> bool {
        self.dict.is_some()
    }

    /// One row of `Worksheet._cells_by_row`: creates missing cells exactly
    /// like `Worksheet.cell(row, column)` and returns a tuple of cells (or of
    /// values when values_only), plus whether cells were created.
    fn _row(
        &mut self,
        py: Python<'_>,
        row: i64,
        min_col: i64,
        max_col: i64,
        values_only: bool,
    ) -> PyResult<(Py<PyAny>, bool)> {
        let mut out: Vec<Py<PyAny>> = Vec::with_capacity((max_col - min_col + 1).max(0) as usize);
        let mut created = false;
        // the row's records in the base area (sorted by column)
        let (mut bi, bend) = match (row <= u32::MAX as i64)
            .then(|| self.base.rows.binary_search(&(row as u32)).ok())
            .flatten()
        {
            Some(r) => (
                self.base.starts[r] as usize,
                self.base.starts[r + 1] as usize,
            ),
            None => (0, 0),
        };
        let gen_empty = self.gen_live == 0;
        // creating many cell objects would trigger a young generation GC pass
        // every few hundred allocations; pause the cyclic GC for the row
        // (state restored on exit, also on error)
        let _gc = if values_only {
            None
        } else {
            Some(GcPause::new())
        };
        for col in min_col..=max_col {
            let key = pack(row, col).expect("checked by caller");
            while bi < bend && (self.base.col(bi) as i64) < col {
                bi += 1;
            }
            let base_hit =
                if bi < bend && self.base.col(bi) as i64 == col && self.base.kind(bi) != K_DEAD {
                    Some(Loc::Base(bi, row as u32))
                } else {
                    None
                };
            let loc = match base_hit {
                Some(l) => l,
                None => match if gen_empty {
                    None
                } else {
                    self.gen_index.get(&key).map(|&i| Loc::Gen(i as usize))
                } {
                    Some(l) => l,
                    None => {
                        created = true;
                        self.insert_new(Entry {
                            key,
                            style: NO_STYLE,
                            kind: K_NONE,
                            dtype: crate::sheet::DT_N,
                            payload: 0,
                        })
                    }
                },
            };
            if values_only {
                out.push(self.value_obj(py, loc)?);
            } else {
                out.push(self.materialize(py, loc)?);
            }
        }
        Ok((PyTuple::new(py, out)?.into_any().unbind(), created))
    }

    /// Diagnostics: (cells, materialised cells, formula templates, text bytes)
    fn _stats(&self) -> (usize, usize, usize, usize) {
        let mat = self
            .locs()
            .into_iter()
            .filter(|&l| self.entry(l).kind == K_OBJ)
            .count();
        (self.live(), mat, self.templates.len(), self.arena.len())
    }
}

/// Bind an empty store to `ws` (called from bind_cells).
pub fn setup_store(
    py: Python<'_>,
    store: &mut CellStore,
    ws: &Bound<'_, PyAny>,
    cell_styles: &Bound<'_, PyAny>,
    shared_strings: &Bound<'_, PyAny>,
    epoch: Option<Epoch>,
    cell_cls: &Bound<'_, PyType>,
) -> PyResult<()> {
    store.reset();
    store.dict = None;
    store.ws = ws.clone().unbind();
    store.cell_styles = cell_styles.clone().unbind();
    store.shared_strings = shared_strings.clone().unbind();
    store.epoch = epoch;
    store.loading = true;
    store.init_runtime(py, cell_cls)
}

/// `ws.parent._cell_styles[style_id]` as done by openpyxl's bind_cells:
/// returns the (non-negative) list index, or raises the same exception.
pub fn style_index(
    cell_styles: &Bound<'_, PyAny>,
    n: i64,
    s: &crate::sheet::StyleId,
) -> PyResult<u32> {
    use crate::sheet::StyleId;
    match s {
        StyleId::Int(i) => {
            let j = if *i < 0 { *i + n } else { *i };
            if j < 0 || j >= n {
                cell_styles.get_item(*i)?; // raises IndexError
                return Err(crate::NativeFallback::new_err("style index"));
            }
            Ok(j as u32)
        }
        StyleId::Obj(o) => {
            cell_styles.get_item(o)?; // raises (eg. TypeError for "")
            Err(crate::NativeFallback::new_err("unusual style id"))
        }
    }
}
