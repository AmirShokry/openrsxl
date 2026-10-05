//! Native worksheet reader reproducing
//! `openpyxl.worksheet._reader.WorkSheetParser` exactly.

use std::collections::{HashMap, HashSet};

use pyo3::exceptions::{PyIndexError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDateTime, PyDict, PyFloat, PyInt, PyList, PyString, PyTime};

use crate::formula::{FormulaError, Translator};
use crate::text::{forbid_none, parse_text, skip_checked, Forbid, TextCtx, TextOutcome};
use crate::utils::{fast_coordinate, fast_float, fast_int};
use crate::xml::{Ev, Reader, Source, XmlError, NS_MAIN, NS_NONE};

/// Tags acted upon by openpyxl's `WorkSheetParser.parse` dispatch loop.
const DISPATCH: &[&[u8]] = &[
    b"col",
    b"sheetProtection",
    b"extLst",
    b"conditionalFormatting",
    b"legacyDrawing",
    b"rowBreaks",
    b"colBreaks",
    b"customSheetViews",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"autoFilter",
    b"dataValidations",
    b"sheetPr",
    b"sheetViews",
    b"sheetFormatPr",
    b"scenarios",
    b"tableParts",
    b"hyperlinks",
    b"mergeCells",
    b"row",
];

pub fn forbid_dispatch(ns: u32, local: &[u8]) -> bool {
    ns == NS_MAIN && DISPATCH.contains(&local)
}

pub enum SheetError {
    Xml(XmlError),
    Py(PyErr),
}

impl From<XmlError> for SheetError {
    fn from(e: XmlError) -> Self {
        SheetError::Xml(e)
    }
}

impl From<PyErr> for SheetError {
    fn from(e: PyErr) -> Self {
        SheetError::Py(e)
    }
}

pub type SResult<T> = Result<T, SheetError>;

/// Python objects used by the parser.
pub struct Helpers {
    pub int: Py<PyAny>,
    pub float: Py<PyAny>,
    pub coordinate_to_tuple: Py<PyAny>,
    pub row_index: Py<PyAny>,
    pub from_excel: Py<PyAny>,
    pub from_iso8601: Py<PyAny>,
    pub date_warning: Py<PyAny>,
    pub array_formula: Py<PyAny>,
    pub data_table_formula: Py<PyAny>,
    pub translator: Py<PyAny>,
    pub tokenizer_error: Py<PyAny>,
    pub translator_error: Py<PyAny>,
    pub text_content: Py<PyAny>,
    pub rich_inline: Py<PyAny>,
    pub validate_rpr: Py<PyAny>,
    pub text_dangerous: HashSet<Vec<u8>>,
    pub rich_dangerous: HashSet<Vec<u8>>,
}

impl Helpers {
    pub fn from_dict(d: &Bound<'_, PyDict>) -> PyResult<Self> {
        let g = |k: &str| -> PyResult<Py<PyAny>> {
            d.get_item(k)?
                .map(|v| v.unbind())
                .ok_or_else(|| PyTypeError::new_err(format!("missing helper {}", k)))
        };
        let names = |k: &str| -> PyResult<HashSet<Vec<u8>>> {
            let v = g(k)?;
            Python::attach(|py| {
                let mut out = HashSet::new();
                for item in v.bind(py).try_iter()? {
                    let s: String = item?.extract()?;
                    out.insert(s.into_bytes());
                }
                Ok(out)
            })
        };
        Ok(Helpers {
            int: g("int")?,
            float: g("float")?,
            coordinate_to_tuple: g("coordinate_to_tuple")?,
            row_index: g("row_index")?,
            from_excel: g("from_excel")?,
            from_iso8601: g("from_ISO8601")?,
            date_warning: g("date_warning")?,
            array_formula: g("ArrayFormula")?,
            data_table_formula: g("DataTableFormula")?,
            translator: g("Translator")?,
            tokenizer_error: g("TokenizerError")?,
            translator_error: g("TranslatorError")?,
            text_content: g("text_content")?,
            rich_inline: g("rich_inline")?,
            validate_rpr: g("validate_rpr")?,
            text_dangerous: names("text_dangerous")?,
            rich_dangerous: names("rich_dangerous")?,
        })
    }
}

pub fn ns_list<'py>(py: Python<'py>, ns: &[(String, String)]) -> PyResult<Bound<'py, PyList>> {
    let l = PyList::empty(py);
    for (p, u) in ns {
        l.append((p.as_str(), u.as_str()))?;
    }
    Ok(l)
}

/// Settings coming from the workbook / parser.
pub struct Settings {
    /// openrsxl.extended formula_and_value: read formula and cached value
    pub dual: bool,
    /// native from_excel parameters, None if the epoch is unusual
    pub fast_epoch: Option<crate::dates::Epoch>,
    pub shared_strings: Py<PyAny>,
    /// len(shared_strings) if it is an exact list (fast index path)
    pub shared_len: usize,
    pub data_only: bool,
    pub rich_text: bool,
    pub epoch: Py<PyAny>,
    pub date_formats: HashSet<i64>,
    pub timedelta_formats: HashSet<i64>,
}

fn int_set(obj: &Bound<'_, PyAny>) -> PyResult<HashSet<i64>> {
    let mut s = HashSet::new();
    for item in obj.try_iter()? {
        let item = item?;
        if let Ok(v) = item.extract::<i64>() {
            if item.is_instance_of::<PyInt>() && !item.is_instance_of::<PyBool>() {
                s.insert(v);
            }
        }
    }
    Ok(s)
}

impl Settings {
    pub fn from_parser(p: &Bound<'_, PyAny>) -> PyResult<Self> {
        let py = p.py();
        let epoch = p.getattr("epoch")?;
        let dtmod = py.import("datetime")?;
        let windows = py
            .import("openrsxl.utils.datetime")?
            .getattr("WINDOWS_EPOCH")?;
        let fast_epoch = if epoch.get_type().is(&dtmod.getattr("datetime")?)
            && epoch.getattr("tzinfo")?.is_none()
        {
            let g = |k: &str| -> PyResult<i64> { epoch.getattr(k)?.extract::<i64>() };
            Some(crate::dates::Epoch::new(
                g("year")?,
                g("month")?,
                g("day")?,
                g("hour")?,
                g("minute")?,
                g("second")?,
                g("microsecond")?,
                epoch.eq(&windows)?,
            ))
        } else {
            None
        };
        Ok(Settings {
            dual: match p.getattr("formula_and_value") {
                Ok(v) => v.is_truthy()?,
                Err(_) => false,
            },
            fast_epoch,
            shared_len: {
                let ss = p.getattr("shared_strings")?;
                if ss.is_exact_instance_of::<PyList>() {
                    ss.len()?
                } else {
                    0
                }
            },
            shared_strings: p.getattr("shared_strings")?.unbind(),
            data_only: p.getattr("data_only")?.is_truthy()?,
            rich_text: p.getattr("rich_text")?.is_truthy()?,
            epoch: p.getattr("epoch")?.unbind(),
            date_formats: int_set(&p.getattr("date_formats")?)?,
            timedelta_formats: int_set(&p.getattr("timedelta_formats")?)?,
        })
    }
}

#[derive(Clone, Copy)]
pub enum Num {
    Int(i64),
    Float(f64),
    Other,
}

pub enum StyleId {
    Int(i64),
    Obj(Py<PyAny>),
}

/// A cell value in compact (not yet Python) form.
pub enum Val {
    None,
    Int(i64),
    Float(f64),
    Bool(bool),
    /// index into the shared string list (0 <= idx < len)
    Sst(u32),
    Text(String),
    /// dependent of a native shared formula master
    Shared {
        master: u32,
        dest: Option<(i64, i64)>,
    },
    /// date serials converted natively (from_excel succeeded at load time)
    DateInt(i64),
    DateFloat(f64),
    Py(Py<PyAny>),
}

/// data_type codes (see DT_NAMES)
pub const DT_N: u8 = 0;
pub const DT_S: u8 = 1;
pub const DT_F: u8 = 2;
pub const DT_B: u8 = 3;
pub const DT_E: u8 = 4;
pub const DT_D: u8 = 5;
pub const DT_INLINE: u8 = 6;
pub const DT_NAMES: [&str; 7] = ["n", "s", "f", "b", "e", "d", "inlineStr"];

pub enum DType {
    Code(u8),
    Py(Py<PyAny>),
}

pub struct CellRec {
    pub row: i64,
    pub column: i64,
    pub val: Val,
    pub dtype: DType,
    pub style_id: StyleId,
    /// formula_and_value: the formula when it is not the primary value
    pub formula: Option<Val>,
    /// formula_and_value: the cached value when it is not the primary value
    pub cached: Option<Val>,
    /// the primary value is the formula (data_only=False, formula cell)
    pub formula_primary: bool,
}

#[derive(Clone, Copy)]
struct CellFlags {
    t_present: bool,
    coord_present: bool,
    f_seen: bool,
    f_type_present: bool,
    f_ref_present: bool,
    f_si_present: bool,
}

/// Reusable text buffers for parse_cell
#[derive(Default)]
pub struct CellBufs {
    t: String,
    coord: String,
    s: String,
    v: String,
    f_type: String,
    f_ref: String,
    f_si: String,
    f_text: String,
}

impl CellBufs {
    fn clear(&mut self) {
        self.t.clear();
        self.coord.clear();
        self.s.clear();
        self.v.clear();
        self.f_type.clear();
        self.f_ref.clear();
        self.f_si.clear();
        self.f_text.clear();
    }
}

pub enum Shared {
    Native(Translator),
    Py(Py<PyAny>),
}

/// Interned data type strings
pub struct DT {
    pub n: Py<PyAny>,
    pub s: Py<PyAny>,
    pub f: Py<PyAny>,
    pub b: Py<PyAny>,
    pub e: Py<PyAny>,
    pub d: Py<PyAny>,
    pub inline: Py<PyAny>,
}

impl DT {
    pub fn code(&self, py: Python<'_>, c: u8) -> Py<PyAny> {
        match c {
            DT_N => self.n.clone_ref(py),
            DT_S => self.s.clone_ref(py),
            DT_F => self.f.clone_ref(py),
            DT_B => self.b.clone_ref(py),
            DT_E => self.e.clone_ref(py),
            DT_D => self.d.clone_ref(py),
            _ => self.inline.clone_ref(py),
        }
    }

    pub fn new(py: Python<'_>) -> Self {
        let i = |s: &str| PyString::intern(py, s).into_any().unbind();
        DT {
            n: i("n"),
            s: i("s"),
            f: i("f"),
            b: i("b"),
            e: i("e"),
            d: i("d"),
            inline: i("inlineStr"),
        }
    }
}

pub enum Item {
    Row(i64, Vec<CellRec>),
    /// a top level element (raw snippet + namespaces) to replay in Python
    Element(Vec<u8>, Vec<(String, String)>),
    Eof,
}

pub struct SheetParser<S: Source> {
    pub r: Reader<S>,
    pub h: Helpers,
    pub st: Settings,
    pub dt: DT,
    text_ctx: TextCtx,
    /// shared formula `si` -> index in `masters`
    shared: crate::store::FxMap<Vec<u8>, u32>,
    /// master of shared formulae without `si` attribute
    shared_none: Option<u32>,
    bufs: CellBufs,
    /// shared formula masters, referenced by Val::Shared
    pub masters: Vec<Shared>,
    row_counter: i64,
    col_counter: i64,
    in_sheet_data: bool,
    /// actions to replay in Python (warnings) - list of tuples
    pub actions: Py<PyList>,
    /// row dimension attributes: str(row) -> dict
    pub row_dimensions: Py<PyDict>,
}

fn text_of(r: &Reader<impl Source>, a: &crate::xml::Attr) -> String {
    r.attr_value(a).into_owned()
}

impl<S: Source> SheetParser<S> {
    pub fn new(
        py: Python<'_>,
        src: S,
        h: Helpers,
        st: Settings,
        actions: Py<PyList>,
        row_dimensions: Py<PyDict>,
    ) -> Self {
        let text_ctx = TextCtx {
            text_dangerous: h.text_dangerous.clone(),
            rich_dangerous: h.rich_dangerous.clone(),
            rpr_cache: HashMap::new(),
        };
        SheetParser {
            r: Reader::new(src),
            h,
            st,
            dt: DT::new(py),
            text_ctx,
            shared: Default::default(),
            shared_none: None,
            bufs: CellBufs::default(),
            masters: Vec::new(),
            row_counter: 0,
            col_counter: 0,
            in_sheet_data: false,
            actions,
            row_dimensions,
        }
    }

    /// Advance to the next row or top-level element.
    pub fn next_item(&mut self, py: Python<'_>) -> SResult<Item> {
        loop {
            match self.r.next()? {
                Ev::Start | Ev::Empty => {
                    let d = self.r.depth();
                    if d == 1 {
                        continue;
                    }
                    if self.in_sheet_data && d == 3 {
                        if self.r.is(NS_MAIN, b"row") {
                            return self.parse_row(py);
                        }
                        return self.capture_checked();
                    }
                    if d == 2 {
                        if self.r.is(NS_MAIN, b"sheetData") {
                            if self.r.ev == Ev::Start {
                                self.in_sheet_data = true;
                            } else {
                                self.r.next()?;
                            }
                            continue;
                        }
                        return self.capture_checked();
                    }
                    // deeper elements are handled by the branches above
                    return Err(SheetError::Xml(XmlError::Unsupported("unexpected nesting")));
                }
                Ev::End => {
                    if self.in_sheet_data && self.r.depth() == 1 {
                        self.in_sheet_data = false;
                    }
                }
                Ev::Text => {}
                Ev::Eof => return Ok(Item::Eof),
            }
        }
    }

    fn capture_checked(&mut self) -> SResult<Item> {
        self.r.begin_capture();
        skip_checked(&mut self.r, forbid_none)?;
        let snippet = self.r.end_capture();
        Ok(Item::Element(snippet, self.r.in_scope_namespaces()))
    }

    fn py_call1(&self, py: Python<'_>, f: &Py<PyAny>, arg: &str) -> PyResult<Py<PyAny>> {
        Ok(f.bind(py).call1((arg,))?.unbind())
    }

    fn parse_row(&mut self, py: Python<'_>) -> SResult<Item> {
        // attributes
        let mut need_dims = false;
        let mut r_attr: Option<String> = None;
        for a in &self.r.attrs {
            if a.ns == NS_NONE {
                let l = self.r.slice(a.local);
                if l == b"r" {
                    r_attr = Some(text_of(&self.r, a));
                } else if l != b"spans" {
                    need_dims = true;
                }
            }
        }
        if let Some(rv) = &r_attr {
            self.row_counter = match fast_int(rv.as_bytes()) {
                Some(v) => v,
                None => self
                    .h
                    .row_index
                    .bind(py)
                    .call1((rv.as_str(),))?
                    .extract::<i64>()?,
            };
        } else {
            // i64 overflow: Python's int keeps counting (fallback)
            self.row_counter = self
                .row_counter
                .checked_add(1)
                .ok_or(SheetError::Xml(XmlError::Unsupported("row index")))?;
        }
        self.col_counter = 0;
        if need_dims {
            let attrs = PyDict::new(py);
            for a in &self.r.attrs {
                attrs.set_item(self.r.attr_key(a), self.r.attr_value(a).as_ref())?;
            }
            self.row_dimensions
                .bind(py)
                .set_item(self.row_counter.to_string(), attrs)?;
        }
        let mut cells = Vec::new();
        if self.r.ev == Ev::Empty {
            self.r.next()?;
            return Ok(Item::Row(self.row_counter, cells));
        }
        let depth = self.r.depth();
        let fast = self.r.default_ns_is_main() && !self.st.rich_text;
        let mut fc = crate::xml::FastCell::default();
        loop {
            if fast && self.r.fast_cell(&mut fc) {
                let c = self.parse_cell_fast(py, &fc)?;
                cells.push(c);
                continue;
            }
            match self.r.next()? {
                Ev::Start | Ev::Empty => {
                    if forbid_dispatch(self.r.name_ns, self.r.local_name()) {
                        return Err(SheetError::Xml(XmlError::Unsupported(
                            "dispatch tag inside row",
                        )));
                    }
                    let c = self.parse_cell(py)?;
                    cells.push(c);
                }
                Ev::End => {
                    if self.r.depth() < depth {
                        break;
                    }
                }
                Ev::Text => {}
                Ev::Eof => return Err(SheetError::Xml(XmlError::Malformed("eof in row"))),
            }
        }
        Ok(Item::Row(self.row_counter, cells))
    }

    fn coordinate(&self, py: Python<'_>, coord: &str) -> PyResult<(i64, i64)> {
        match fast_coordinate(coord.as_bytes()) {
            Some(t) => Ok(t),
            None => self
                .h
                .coordinate_to_tuple
                .bind(py)
                .call1((coord,))?
                .extract::<(i64, i64)>(),
        }
    }

    fn py_int(&self, py: Python<'_>, s: &str) -> PyResult<Py<PyAny>> {
        match fast_int(s.as_bytes()) {
            Some(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
            None => self.py_call1(py, &self.h.int, s),
        }
    }

    fn cast_number(&self, py: Python<'_>, s: &str) -> PyResult<(Py<PyAny>, Num)> {
        let b = s.as_bytes();
        if b.iter().any(|&c| c == b'.' || c == b'E' || c == b'e') {
            match fast_float(b) {
                Some(f) => Ok((PyFloat::new(py, f).into_any().unbind(), Num::Float(f))),
                None => {
                    let o = self.py_call1(py, &self.h.float, s)?;
                    let f = o.bind(py).extract::<f64>().ok();
                    Ok((o, f.map(Num::Float).unwrap_or(Num::Other)))
                }
            }
        } else {
            match fast_int(b) {
                Some(v) => Ok((v.into_pyobject(py)?.into_any().unbind(), Num::Int(v))),
                None => Ok((self.py_call1(py, &self.h.int, s)?, Num::Other)),
            }
        }
    }

    fn parse_cell(&mut self, py: Python<'_>) -> SResult<CellRec> {
        // reusable buffers (no allocation per cell)
        let mut b = std::mem::take(&mut self.bufs);
        b.clear();
        let res = self.parse_cell_inner(py, &mut b);
        self.bufs = b;
        res
    }

    fn parse_cell_inner(&mut self, py: Python<'_>, b: &mut CellBufs) -> SResult<CellRec> {
        let (mut t_present, mut coord_present, mut s_present) = (false, false, false);
        for a in &self.r.attrs {
            if a.ns == NS_NONE {
                match self.r.slice(a.local) {
                    b"t" => {
                        t_present = true;
                        b.t.push_str(&self.r.attr_value(a));
                    }
                    b"r" => {
                        coord_present = true;
                        b.coord.push_str(&self.r.attr_value(a));
                    }
                    b"s" => {
                        s_present = true;
                        b.s.push_str(&self.r.attr_value(a));
                    }
                    _ => {}
                }
            }
        }
        let style_id = self.style_of(py, s_present, b)?;
        let data_type_raw: &str = if t_present { b.t.as_str() } else { "n" };

        // children
        let mut v_seen = false;
        let mut f_seen = false;
        let (mut f_type_present, mut f_ref_present, mut f_si_present) = (false, false, false);
        let mut f_attrib: Option<Bound<'_, PyDict>> = None;
        let mut is_seen = false;
        let mut is_value: Option<Val> = None;
        let want_is = data_type_raw == "inlineStr";
        if self.r.ev == Ev::Start {
            let depth = self.r.depth();
            loop {
                match self.r.next()? {
                    Ev::Start | Ev::Empty => {
                        let ns = self.r.name_ns;
                        let local = self.r.local_name();
                        if forbid_dispatch(ns, local) {
                            return Err(SheetError::Xml(XmlError::Unsupported(
                                "dispatch tag inside cell",
                            )));
                        }
                        if ns == NS_MAIN && local == b"v" && !v_seen {
                            v_seen = true;
                            crate::text::element_text_into(&mut self.r, forbid_dispatch, &mut b.v)?;
                        } else if ns == NS_MAIN && local == b"f" && !f_seen {
                            f_seen = true;
                            for a in &self.r.attrs {
                                if a.ns == NS_NONE {
                                    match self.r.slice(a.local) {
                                        b"t" => {
                                            f_type_present = true;
                                            b.f_type.push_str(&self.r.attr_value(a));
                                        }
                                        b"ref" => {
                                            f_ref_present = true;
                                            b.f_ref.push_str(&self.r.attr_value(a));
                                        }
                                        b"si" => {
                                            f_si_present = true;
                                            b.f_si.push_str(&self.r.attr_value(a));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            if f_type_present && b.f_type == "dataTable" && !self.st.data_only {
                                let d = PyDict::new(py);
                                for a in &self.r.attrs {
                                    d.set_item(self.r.attr_key(a), self.r.attr_value(a).as_ref())?;
                                }
                                f_attrib = Some(d);
                            }
                            crate::text::element_text_into(
                                &mut self.r,
                                forbid_dispatch,
                                &mut b.f_text,
                            )?;
                        } else if ns == NS_MAIN && local == b"is" && !is_seen {
                            is_seen = true;
                            if want_is {
                                is_value = Some(self.inline_string(py)?);
                            } else {
                                skip_checked(&mut self.r, forbid_dispatch)?;
                            }
                        } else {
                            skip_checked(&mut self.r, forbid_dispatch)?;
                        }
                    }
                    Ev::End => {
                        if self.r.depth() < depth {
                            break;
                        }
                    }
                    Ev::Text => {}
                    Ev::Eof => return Err(SheetError::Xml(XmlError::Malformed("eof in cell"))),
                }
            }
        } else {
            self.r.next()?; // End of empty cell
        }

        let flags = CellFlags {
            t_present,
            coord_present,
            f_seen,
            f_type_present,
            f_ref_present,
            f_si_present,
        };
        self.interpret(py, b, style_id, flags, f_attrib, is_value)
    }

    fn style_of(&self, py: Python<'_>, s_present: bool, b: &CellBufs) -> PyResult<StyleId> {
        Ok(if !s_present {
            StyleId::Int(0)
        } else if b.s.is_empty() {
            StyleId::Obj(PyString::new(py, "").into_any().unbind())
        } else {
            match fast_int(b.s.as_bytes()) {
                Some(v) => StyleId::Int(v),
                None => {
                    let v = self.h.int.bind(py).call1((b.s.as_str(),))?;
                    match v.extract::<i64>() {
                        Ok(i) => StyleId::Int(i),
                        Err(_) => StyleId::Obj(v.unbind()),
                    }
                }
            }
        })
    }

    /// The simple cell scanned by `Reader::fast_cell`.
    fn parse_cell_fast(&mut self, py: Python<'_>, fc: &crate::xml::FastCell) -> SResult<CellRec> {
        let mut b = std::mem::take(&mut self.bufs);
        b.clear();
        let res = (|| -> SResult<CellRec> {
            let buf = &self.r.buf;
            let take = |r: Option<crate::xml::Range>, dst: &mut String| -> bool {
                match r {
                    Some(r) => {
                        // validated as ASCII / UTF-8 by the fast scanner
                        dst.push_str(unsafe {
                            std::str::from_utf8_unchecked(&buf[r.start..r.end])
                        });
                        true
                    }
                    None => false,
                }
            };
            let t_present = take(fc.t, &mut b.t);
            let coord_present = take(fc.r, &mut b.coord);
            let s_present = take(fc.s, &mut b.s);
            let f_type_present = take(fc.f_t, &mut b.f_type);
            let f_ref_present = take(fc.f_ref, &mut b.f_ref);
            let f_si_present = take(fc.f_si, &mut b.f_si);
            take(fc.f_text, &mut b.f_text);
            let data_type_raw = if t_present { b.t.as_str() } else { "n" };
            if data_type_raw != "inlineStr" {
                take(fc.v, &mut b.v);
            }
            let style_id = self.style_of(py, s_present, &b)?;
            let flags = CellFlags {
                t_present,
                coord_present,
                f_seen: fc.f,
                f_type_present,
                f_ref_present,
                f_si_present,
            };
            self.interpret(py, &mut b, style_id, flags, None, None)
        })();
        self.bufs = b;
        res
    }

    fn interpret(
        &mut self,
        py: Python<'_>,
        b: &mut CellBufs,
        style_id: StyleId,
        flags: CellFlags,
        f_attrib: Option<Bound<'_, PyDict>>,
        is_value: Option<Val>,
    ) -> SResult<CellRec> {
        let CellFlags {
            t_present,
            coord_present,
            f_seen,
            f_type_present,
            f_ref_present,
            f_si_present,
        } = flags;
        let data_type_raw: &str = if t_present { b.t.as_str() } else { "n" };
        // `element.findtext(VALUE_TAG, None) or None`
        let value: Option<String> = if data_type_raw == "inlineStr" || b.v.is_empty() {
            None
        } else {
            Some(std::mem::take(&mut b.v))
        };
        let coordinate: Option<&str> = if coord_present {
            Some(b.coord.as_str())
        } else {
            None
        };

        // coordinates
        let (row, column) = match coordinate {
            Some(c) if !c.is_empty() => {
                let (r, c) = self.coordinate(py, c)?;
                self.col_counter = c;
                (r, c)
            }
            _ => {
                self.col_counter = self
                    .col_counter
                    .checked_add(1)
                    .ok_or(SheetError::Xml(XmlError::Unsupported("column index")))?;
                (self.row_counter, self.col_counter)
            }
        };

        let primary_formula = !self.st.data_only && f_seen;
        let mut fval: Option<Val> = None;
        if f_seen && (primary_formula || self.st.dual) {
            fval = Some(self.parse_formula(
                py,
                if f_type_present {
                    Some(b.f_type.as_str())
                } else {
                    None
                },
                if f_ref_present {
                    Some(b.f_ref.as_str())
                } else {
                    None
                },
                if f_si_present {
                    Some(b.f_si.as_str())
                } else {
                    None
                },
                f_attrib,
                b.f_text.as_str(),
                coordinate,
                match coordinate {
                    Some(c) if !c.is_empty() => Some((row, column)),
                    _ => None,
                },
            )?);
        }
        let quiet = primary_formula;
        let vres = if !primary_formula || self.st.dual {
            let style_int = match &style_id {
                StyleId::Int(i) => Some(*i),
                _ => None,
            };
            Some(self.value_branch(
                py,
                value,
                data_type_raw,
                style_int,
                coordinate,
                is_value,
                quiet,
            )?)
        } else {
            None
        };
        let (val, dtype, formula, cached) = if primary_formula {
            (fval.unwrap(), DType::Code(DT_F), None, vres.map(|x| x.0))
        } else {
            let (v, d) = vres.unwrap();
            (v, d, fval, None)
        };
        Ok(CellRec {
            row,
            column,
            val,
            dtype,
            style_id,
            formula,
            cached,
            formula_primary: primary_formula,
        })
    }

    /// The value of a cell as `data_only=True` (or a cell without formula)
    /// reads it. `quiet`: do not record warnings (secondary value of
    /// formula_and_value).
    #[allow(clippy::too_many_arguments)]
    fn value_branch(
        &mut self,
        py: Python<'_>,
        mut value: Option<String>,
        data_type_raw: &str,
        style_int: Option<i64>,
        coordinate: Option<&str>,
        is_value: Option<Val>,
        quiet: bool,
    ) -> SResult<(Val, DType)> {
        let (val, dtype): (Val, DType);
        if let Some(v) = value.take() {
            match data_type_raw {
                "n" => {
                    let (num, kind) = self.cast_number(py, &v)?;
                    let is_date = matches!(style_int, Some(i) if self.st.date_formats.contains(&i));
                    let td = is_date
                        && matches!(style_int, Some(i) if self.st.timedelta_formats.contains(&i));
                    let fast = if is_date && !td {
                        match (&self.st.fast_epoch, kind) {
                            (Some(ep), Num::Int(i)) => Some(crate::dates::from_excel_int(i, ep)),
                            (Some(ep), Num::Float(f)) => {
                                Some(crate::dates::from_excel_float(f, ep))
                            }
                            _ => None,
                        }
                    } else {
                        None
                    };
                    if let Some(res) = fast {
                        use crate::dates::ExcelDate;
                        match res {
                            ExcelDate::Error => {
                                let msg = self
                                    .h
                                    .date_warning
                                    .bind(py)
                                    .call1((coordinate, num.bind(py)))?;
                                if !quiet {
                                    self.actions.bind(py).append(("warn", msg))?;
                                }
                                dtype = DType::Code(DT_E);
                                val = Val::Text("#VALUE!".to_string());
                            }
                            _ => {
                                dtype = DType::Code(DT_D);
                                val = match kind {
                                    Num::Int(i) => Val::DateInt(i),
                                    Num::Float(f) => Val::DateFloat(f),
                                    Num::Other => unreachable!(),
                                };
                            }
                        }
                    } else if is_date {
                        let kw = PyDict::new(py);
                        kw.set_item("timedelta", td)?;
                        match self
                            .h
                            .from_excel
                            .bind(py)
                            .call((num.bind(py), self.st.epoch.bind(py)), Some(&kw))
                        {
                            Ok(d) => {
                                dtype = DType::Code(DT_D);
                                val = Val::Py(d.unbind());
                            }
                            Err(e) => {
                                if e.is_instance_of::<pyo3::exceptions::PyOverflowError>(py)
                                    || e.is_instance_of::<pyo3::exceptions::PyValueError>(py)
                                {
                                    let msg = self
                                        .h
                                        .date_warning
                                        .bind(py)
                                        .call1((coordinate, num.bind(py)))?;
                                    if !quiet {
                                        self.actions.bind(py).append(("warn", msg))?;
                                    }
                                    dtype = DType::Code(DT_E);
                                    val = Val::Text("#VALUE!".to_string());
                                } else {
                                    return Err(e.into());
                                }
                            }
                        }
                    } else {
                        dtype = DType::Code(DT_N);
                        val = match kind {
                            Num::Int(i) => Val::Int(i),
                            Num::Float(f) => Val::Float(f),
                            Num::Other => Val::Py(num),
                        };
                    }
                }
                "s" => {
                    let len = self.st.shared_len;
                    val = match fast_int(v.as_bytes()) {
                        Some(i) if i >= 0 && (i as usize) < len => Val::Sst(i as u32),
                        _ => {
                            let idx = self.py_int(py, &v)?;
                            let ss = self.st.shared_strings.bind(py);
                            Val::Py(ss.get_item(idx.bind(py))?.unbind())
                        }
                    };
                    dtype = DType::Code(DT_S);
                }
                "b" => {
                    let b = match fast_int(v.as_bytes()) {
                        Some(i) => i != 0,
                        None => self.py_int(py, &v)?.bind(py).is_truthy()?,
                    };
                    val = Val::Bool(b);
                    dtype = DType::Code(DT_B);
                }
                "str" => {
                    val = Val::Text(v);
                    dtype = DType::Code(DT_S);
                }
                "d" => {
                    val = Val::Py(self.h.from_iso8601.bind(py).call1((v.as_str(),))?.unbind());
                    dtype = DType::Code(DT_D);
                }
                other => {
                    dtype = dtype_of(py, other);
                    val = Val::Text(v);
                }
            }
        } else if data_type_raw == "inlineStr" {
            match is_value {
                Some(v) => {
                    dtype = DType::Code(DT_S);
                    val = v;
                }
                None => {
                    dtype = DType::Code(DT_INLINE);
                    val = Val::None;
                }
            }
        } else {
            dtype = dtype_of(py, data_type_raw);
            val = Val::None;
        }
        Ok((val, dtype))
    }

    /// Convert a compact value to the Python object openpyxl would hold.
    pub fn val_to_py(&self, py: Python<'_>, v: Val) -> PyResult<Py<PyAny>> {
        to_py(
            py,
            v,
            self.st.shared_strings.bind(py),
            &self.masters,
            self.st.fast_epoch.as_ref(),
        )
    }

    fn inline_string(&mut self, py: Python<'_>) -> SResult<Val> {
        if self.st.rich_text {
            self.r.begin_capture();
            skip_checked(&mut self.r, forbid_dispatch)?;
            let snippet = self.r.end_capture();
            let ns = ns_list(py, &self.r.in_scope_namespaces())?;
            return Ok(Val::Py(
                self.h
                    .rich_inline
                    .bind(py)
                    .call1((PyBytes::new(py, &snippet), ns))?
                    .unbind(),
            ));
        }
        let validate = self.h.validate_rpr.clone_ref(py);
        let mut vf = |snip: &[u8], ns: Vec<(String, String)>| -> PyResult<bool> {
            let l = ns_list(py, &ns)?;
            validate
                .bind(py)
                .call1((PyBytes::new(py, snip), l))?
                .is_truthy()
        };
        let forbid: Forbid = forbid_dispatch;
        match parse_text(&mut self.r, &mut self.text_ctx, forbid, &mut vf)? {
            Err(e) => Err(e.into()),
            Ok(TextOutcome::Content(s)) => Ok(Val::Text(s)),
            Ok(TextOutcome::Fallback(snippet, ns)) => {
                let l = ns_list(py, &ns)?;
                Ok(Val::Py(
                    self.h
                        .text_content
                        .bind(py)
                        .call1((PyBytes::new(py, &snippet), l))?
                        .unbind(),
                ))
            }
        }
    }

    fn formula_err(&self, py: Python<'_>, e: FormulaError) -> PyErr {
        match e {
            FormulaError::Tokenizer(m) => match self.h.tokenizer_error.bind(py).call1((m,)) {
                Ok(exc) => PyErr::from_value(exc),
                Err(e) => e,
            },
            FormulaError::Translator(m) => match self.h.translator_error.bind(py).call1((m,)) {
                Ok(exc) => PyErr::from_value(exc),
                Err(e) => e,
            },
            FormulaError::PopEmpty => PyIndexError::new_err("pop from empty list"),
            FormulaError::NeedPython => PyTypeError::new_err("internal: NeedPython"),
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn parse_formula(
        &mut self,
        py: Python<'_>,
        f_type: Option<&str>,
        f_ref: Option<&str>,
        f_si: Option<&str>,
        f_attrib: Option<Bound<'_, PyDict>>,
        f_text: &str,
        coordinate: Option<&str>,
        coord_rc: Option<(i64, i64)>,
    ) -> PyResult<Val> {
        let make_value = || {
            let mut v = String::with_capacity(f_text.len() + 1);
            v.push('=');
            v.push_str(f_text);
            v
        };
        match f_type {
            Some("array") => {
                let kw = PyDict::new(py);
                kw.set_item("ref", f_ref)?;
                kw.set_item("text", make_value())?;
                Ok(Val::Py(
                    self.h.array_formula.bind(py).call((), Some(&kw))?.unbind(),
                ))
            }
            Some("shared") => {
                let known = match f_si {
                    Some(si) => self.shared.get(si.as_bytes()).copied(),
                    None => self.shared_none,
                };
                if let Some(mi) = known {
                    match &self.masters[mi as usize] {
                        Shared::Native(t) => {
                            // the coordinate was already parsed by parse_cell
                            let dest = coord_rc;
                            // translation errors are raised now, like openpyxl
                            match t.check(dest) {
                                Ok(()) => Ok(Val::Shared { master: mi, dest }),
                                Err(e) => Err(self.formula_err(py, e)),
                            }
                        }
                        Shared::Py(obj) => {
                            let s: Py<PyAny> = obj
                                .bind(py)
                                .call_method1("translate_formula", (coordinate,))?
                                .unbind();
                            Ok(Val::Py(s))
                        }
                    }
                } else {
                    let value = make_value();
                    if value != "=" {
                        let native = match coordinate {
                            Some(c) => {
                                let (row, col) = self.coordinate(py, c)?;
                                match Translator::new(&value, row, col) {
                                    Ok(t) => Some(Shared::Native(t)),
                                    Err(FormulaError::NeedPython) => None,
                                    Err(e) => return Err(self.formula_err(py, e)),
                                }
                            }
                            None => None,
                        };
                        let entry = match native {
                            Some(n) => n,
                            None => Shared::Py(
                                self.h
                                    .translator
                                    .bind(py)
                                    .call1((value.as_str(), coordinate))?
                                    .unbind(),
                            ),
                        };
                        self.masters.push(entry);
                        let mi = (self.masters.len() - 1) as u32;
                        match f_si {
                            Some(si) => {
                                self.shared.insert(si.as_bytes().to_vec(), mi);
                            }
                            None => self.shared_none = Some(mi),
                        }
                    }
                    Ok(Val::Text(value))
                }
            }
            Some("dataTable") => {
                let kw = f_attrib.unwrap_or_else(|| PyDict::new(py));
                Ok(Val::Py(
                    self.h
                        .data_table_formula
                        .bind(py)
                        .call((), Some(&kw))?
                        .unbind(),
                ))
            }
            _ => Ok(Val::Text(make_value())),
        }
    }
}

fn dtype_of(py: Python<'_>, s: &str) -> DType {
    match DT_NAMES.iter().position(|n| *n == s) {
        Some(i) => DType::Code(i as u8),
        None => DType::Py(PyString::new(py, s).into_any().unbind()),
    }
}

/// Convert a compact value into the Python object openpyxl would hold.
pub fn to_py(
    py: Python<'_>,
    v: Val,
    shared_strings: &Bound<'_, PyAny>,
    masters: &[Shared],
    epoch: Option<&crate::dates::Epoch>,
) -> PyResult<Py<PyAny>> {
    Ok(match v {
        Val::None => py.None(),
        Val::Int(i) => i.into_pyobject(py)?.into_any().unbind(),
        Val::Float(f) => PyFloat::new(py, f).into_any().unbind(),
        Val::Bool(b) => PyBool::new(py, b).to_owned().into_any().unbind(),
        Val::Sst(i) => shared_strings.get_item(i as usize)?.unbind(),
        Val::Text(s) => PyString::new(py, &s).into_any().unbind(),
        Val::Shared { master, dest } => shared_to_py(py, &masters[master as usize], dest)?,
        Val::DateInt(i) => date_to_py(py, crate::dates::from_excel_int(i, epoch.unwrap()))?,
        Val::DateFloat(f) => date_to_py(py, crate::dates::from_excel_float(f, epoch.unwrap()))?,
        Val::Py(o) => o,
    })
}

pub fn shared_to_py(py: Python<'_>, m: &Shared, dest: Option<(i64, i64)>) -> PyResult<Py<PyAny>> {
    match m {
        Shared::Native(t) => match t.translate(dest) {
            Ok(s) => Ok(PyString::new(py, &s).into_any().unbind()),
            Err(_) => unreachable!("validated at load time"),
        },
        Shared::Py(_) => unreachable!("python translators produce values directly"),
    }
}

pub fn date_to_py(py: Python<'_>, d: crate::dates::ExcelDate) -> PyResult<Py<PyAny>> {
    use crate::dates::ExcelDate;
    Ok(match d {
        ExcelDate::DateTime {
            y,
            m,
            d,
            h,
            mi,
            s,
            us,
        } => PyDateTime::new(py, y, m, d, h, mi, s, us, None)?
            .into_any()
            .unbind(),
        ExcelDate::Time { h, mi, s, us } => {
            PyTime::new(py, h, mi, s, us, None)?.into_any().unbind()
        }
        ExcelDate::Error => unreachable!("validated at load time"),
    })
}

/// Convert an XmlError into the Python exception used to request the pure
/// Python fallback.
pub fn fallback_err(_py: Python<'_>, e: SheetError) -> PyErr {
    match e {
        SheetError::Py(e) => e,
        SheetError::Xml(XmlError::Source(e)) => e,
        SheetError::Xml(XmlError::Malformed(m)) => {
            crate::NativeFallback::new_err(format!("malformed: {}", m))
        }
        SheetError::Xml(XmlError::Unsupported(m)) => {
            crate::NativeFallback::new_err(format!("unsupported: {}", m))
        }
    }
}

/// Build a style id Python object
pub fn style_obj(py: Python<'_>, s: &StyleId) -> Py<PyAny> {
    match s {
        StyleId::Int(i) => i.into_pyobject(py).unwrap().into_any().unbind(),
        StyleId::Obj(o) => o.clone_ref(py),
    }
}

pub fn dtype_to_py(py: Python<'_>, dt: &DT, d: DType) -> Py<PyAny> {
    match d {
        DType::Code(c) => dt.code(py, c),
        DType::Py(o) => o,
    }
}
