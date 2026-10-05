//! Native `<sheetData>` writer reproducing `openpyxl.worksheet._writer`
//! (`rows`, `write_row`) and `openpyxl.cell._writer` (`write_cell`) byte for
//! byte, for both serialisation back-ends openpyxl can use:
//!
//! * ElementTree + et_xmlfile (`lxml=False`): output is text, written through
//!   the same text-mode write function et_xmlfile uses (so newline
//!   translation and encoding are identical);
//! * lxml (`lxml=True`): output is ASCII bytes with character references.
//!
//! Cells whose values are not plain str/int/float/bool/None (rich text,
//! array formulae, numpy/Decimal values, strings with characters that need
//! special treatment ...) are rendered by the original Python code.

use std::borrow::Cow;
use std::collections::HashMap;

use pyo3::buffer::PyBuffer;
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple};

use crate::sheet::DT_NAMES;
use crate::store::{
    unpack, CellStore, LazyValue, Loc, K_DATE_FLOAT, K_DATE_INT, K_NONE, K_PYVAL, NO_STYLE,
};
use crate::utils::{format_g16, push_column_letter, py_strip};

const FLUSH_AT: usize = 1 << 20;

struct Names {
    value: Py<PyString>,
    style: Py<PyString>,
    comment: Py<PyString>,
    data_type: Py<PyString>,
    hyperlink: Py<PyString>,
    row: Py<PyString>,
    column: Py<PyString>,
}

#[pyclass]
pub struct SheetWriter {
    ws: Py<PyAny>,
    write: Py<PyAny>,
    lxml: bool,
    buf: Vec<u8>,
    names: Names,
    cell_styles: Py<PyAny>,
    style_cache: HashMap<[i32; 9], Py<PyAny>>,
    style_array_type: Py<PyAny>,
    hyperlinks: Py<PyAny>,
    add_comment: Py<PyAny>,
    fallback: Py<PyAny>,
    date_value: Py<PyAny>,
    date_value_raw: Py<PyAny>,
    safe_string: Py<PyAny>,
    /// current workbook epoch (for native to_excel), None if unusual
    epoch_now: Option<crate::dates::Epoch>,
    iso_dates: bool,
}

#[inline]
fn is_illegal_ctrl(c: char) -> bool {
    (c as u32) < 0x20 && c != '\t' && c != '\n' && c != '\r'
}

impl SheetWriter {
    fn esc_text(&mut self, s: &str) {
        let buf = &mut self.buf;
        if self.lxml {
            let b = s.as_bytes();
            if !b
                .iter()
                .any(|&c| c == b'&' || c == b'<' || c == b'>' || c == b'\r' || c >= 0x80)
            {
                buf.extend_from_slice(b);
                return;
            }
            for c in s.chars() {
                match c {
                    '&' => buf.extend_from_slice(b"&amp;"),
                    '<' => buf.extend_from_slice(b"&lt;"),
                    '>' => buf.extend_from_slice(b"&gt;"),
                    '\r' => buf.extend_from_slice(b"&#13;"),
                    c if (c as u32) < 0x80 => buf.push(c as u8),
                    c => buf.extend_from_slice(format!("&#{};", c as u32).as_bytes()),
                }
            }
        } else {
            let b = s.as_bytes();
            if memchr::memchr3(b'&', b'<', b'>', b).is_none() {
                buf.extend_from_slice(b);
                return;
            }
            let mut last = 0;
            for (i, &c) in b.iter().enumerate() {
                let rep: &[u8] = match c {
                    b'&' => b"&amp;",
                    b'<' => b"&lt;",
                    b'>' => b"&gt;",
                    _ => continue,
                };
                buf.extend_from_slice(&b[last..i]);
                buf.extend_from_slice(rep);
                last = i + 1;
            }
            buf.extend_from_slice(&b[last..]);
        }
    }

    fn esc_attr(&mut self, s: &str) {
        let buf = &mut self.buf;
        for c in s.chars() {
            match c {
                '&' => buf.extend_from_slice(b"&amp;"),
                '<' => buf.extend_from_slice(b"&lt;"),
                '>' => buf.extend_from_slice(b"&gt;"),
                '"' => buf.extend_from_slice(b"&quot;"),
                '\n' => buf.extend_from_slice(b"&#10;"),
                '\r' => buf.extend_from_slice(b"&#13;"),
                '\t' => buf.extend_from_slice(if self.lxml { b"&#9;" } else { b"&#09;" }),
                c if (c as u32) < 0x80 || !self.lxml => {
                    let mut tmp = [0u8; 4];
                    buf.extend_from_slice(c.encode_utf8(&mut tmp).as_bytes());
                }
                c => buf.extend_from_slice(format!("&#{};", c as u32).as_bytes()),
            }
        }
    }

    fn flush_if_needed(&mut self, py: Python<'_>) -> PyResult<()> {
        if self.buf.len() >= FLUSH_AT {
            self.flush_buf(py)?;
        }
        Ok(())
    }

    fn flush_buf(&mut self, py: Python<'_>) -> PyResult<()> {
        if self.buf.is_empty() {
            return Ok(());
        }
        if self.lxml {
            self.write.bind(py).call1((PyBytes::new(py, &self.buf),))?;
        } else {
            // SAFETY: only valid utf-8 is ever appended
            let s = unsafe { std::str::from_utf8_unchecked(&self.buf) };
            self.write.bind(py).call1((PyString::new(py, s),))?;
        }
        self.buf.clear();
        Ok(())
    }

    /// Append the output of the Python fallback renderer.
    fn push_fallback(
        &mut self,
        py: Python<'_>,
        cell: &Bound<'_, PyAny>,
        styled: bool,
    ) -> PyResult<()> {
        let out = self
            .fallback
            .bind(py)
            .call1((self.ws.bind(py), cell, styled))?;
        if self.lxml {
            let b = out.cast::<PyBytes>()?;
            self.buf.extend_from_slice(b.as_bytes());
        } else {
            let s = out.cast::<PyString>()?;
            match s.to_str() {
                Ok(v) => self.buf.extend_from_slice(v.as_bytes()),
                Err(_) => {
                    // contains surrogates: flush and pass the str through
                    self.flush_buf(py)?;
                    self.write.bind(py).call1((s,))?;
                }
            }
        }
        Ok(())
    }

    fn style_id(&mut self, py: Python<'_>, style: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        if style.get_type().is(self.style_array_type.bind(py)) {
            if let Ok(b) = PyBuffer::<i32>::get(style) {
                if b.item_count() == 9 {
                    let v = b.to_vec(py)?;
                    let key: [i32; 9] = v.try_into().unwrap();
                    if let Some(id) = self.style_cache.get(&key) {
                        return Ok(id.clone_ref(py));
                    }
                    let id = self
                        .cell_styles
                        .bind(py)
                        .call_method1("add", (style,))?
                        .unbind();
                    self.style_cache.insert(key, id.clone_ref(py));
                    return Ok(id);
                }
            }
        }
        Ok(self
            .cell_styles
            .bind(py)
            .call_method1("add", (style,))?
            .unbind())
    }

    fn has_style(&self, py: Python<'_>, style: &Bound<'_, PyAny>) -> PyResult<bool> {
        if style.is_none() {
            return Ok(false);
        }
        if style.get_type().is(self.style_array_type.bind(py)) {
            if let Ok(b) = PyBuffer::<i32>::get(style) {
                let v = b.to_vec(py)?;
                return Ok(v.iter().any(|&x| x != 0));
            }
        }
        let any = py.import("builtins")?.getattr("any")?;
        any.call1((style,))?.is_truthy()
    }

    fn start_c(&mut self, row: i64, col: i64, style_id: Option<&str>, t: Option<&str>) {
        self.buf.extend_from_slice(b"<c r=\"");
        push_column_letter(&mut self.buf, col);
        let mut tmp = [0u8; 24];
        self.buf
            .extend_from_slice(crate::ftemplate::fmt_i64(row, &mut tmp).as_bytes());
        self.buf.push(b'"');
        if let Some(s) = style_id {
            self.buf.extend_from_slice(b" s=\"");
            self.esc_attr(s);
            self.buf.push(b'"');
        }
        if let Some(t) = t {
            self.buf.extend_from_slice(b" t=\"");
            self.esc_attr(t);
            self.buf.push(b'"');
        }
    }

    fn empty_close(&mut self) {
        if self.lxml {
            // lxml_write_cell uses `with xf.element("c", attributes): return`
            self.buf.extend_from_slice(b"></c>");
        } else {
            self.buf.extend_from_slice(b" />");
        }
    }

    fn empty_v(&mut self) {
        if self.lxml {
            self.buf.extend_from_slice(b"<v></v>");
        } else {
            self.buf.extend_from_slice(b"<v />");
        }
    }

    /// write_row: handles comments, skipping and write_cell for one cell.
    fn write_one(&mut self, py: Python<'_>, cell: &Bound<'_, PyAny>) -> PyResult<()> {
        let n = &self.names;
        let comment = cell.getattr(n.comment.bind(py))?;
        if !comment.is_none() {
            self.add_comment.bind(py).call1((self.ws.bind(py), cell))?;
        }
        let value = cell.getattr(n.value.bind(py))?;
        let style = cell.getattr(n.style.bind(py))?;
        let styled = self.has_style(py, &style)?;
        if value.is_none() && !styled && !comment.is_truthy()? {
            return Ok(());
        }
        self.write_cell(py, cell, value, style, styled)
    }

    fn write_cell(
        &mut self,
        py: Python<'_>,
        cell: &Bound<'_, PyAny>,
        value: Bound<'_, PyAny>,
        style: Bound<'_, PyAny>,
        styled: bool,
    ) -> PyResult<()> {
        let n = &self.names;
        let data_type_obj = cell.getattr(n.data_type.bind(py))?;
        let row_obj = cell.getattr(n.row.bind(py))?;
        let col_obj = cell.getattr(n.column.bind(py))?;
        let (row, col) = match (exact_int(&row_obj), exact_int(&col_obj)) {
            (Some(r), Some(c)) if (1..=18278).contains(&c) => (r, c),
            _ => return self.push_fallback(py, cell, styled),
        };
        let data_type: String = match data_type_obj.cast::<PyString>() {
            Ok(s) => match s.to_str() {
                Ok(v) => v.to_string(),
                Err(_) => return self.push_fallback(py, cell, styled),
            },
            Err(_) => return self.push_fallback(py, cell, styled),
        };

        // style id first, like _set_attributes (adding to _cell_styles is idempotent)
        let sid: Option<String> = if styled {
            let id = self.style_id(py, &style)?;
            Some(id.bind(py).str()?.to_string())
        } else {
            None
        };
        // classify the value; anything unusual goes to the Python renderer
        let mut t_attr: Option<String> = match data_type.as_str() {
            "s" => Some("inlineStr".to_string()),
            "f" => None,
            other => Some(other.to_string()),
        };
        let v = if data_type == "d" {
            // date conversion through the original helpers (timezone check,
            // iso_dates, to_excel)
            let res = self.date_value.bind(py).call1((self.ws.bind(py), cell))?;
            let (dv, iso): (Bound<'_, PyAny>, bool) = res.extract()?;
            if !iso {
                t_attr = Some("n".to_string());
            }
            match classify(py, &dv, &self.safe_string)? {
                Some(x) => owned(x),
                None => return self.push_fallback(py, cell, styled),
            }
        } else {
            match classify(py, &value, &self.safe_string)? {
                Some(x) => x,
                None => return self.push_fallback(py, cell, styled),
            }
        };
        let link = cell.getattr(self.names.hyperlink.bind(py))?;
        if !self.emit(
            py,
            row,
            col,
            &data_type,
            t_attr.as_deref(),
            v,
            sid.as_deref(),
            Some(link),
        )? {
            return self.push_fallback(py, cell, styled);
        }
        Ok(())
    }

    /// Emit a `<c>` element; returns false (without side effects) if the
    /// value needs the Python renderer.
    #[allow(clippy::too_many_arguments)]
    fn emit(
        &mut self,
        py: Python<'_>,
        row: i64,
        col: i64,
        data_type: &str,
        t_attr: Option<&str>,
        v: Classified<'_>,
        sid: Option<&str>,
        link: Option<Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        enum V<'b> {
            None,
            Str(Cow<'b, str>),
            Num(Cow<'b, str>),
        }
        // value kinds acceptable per data type
        let v = match (data_type, v) {
            (_, Classified::None) => V::None,
            ("s", Classified::Str(s)) | ("f", Classified::Str(s)) => V::Str(s),
            ("s", _) | ("f", _) => return Ok(false),
            (_, Classified::Str(s)) if s.is_empty() => V::None,
            (_, Classified::Str(s)) => V::Num(s),
            (_, Classified::Num(s)) => V::Num(s),
        };
        if let V::Str(s) = &v {
            if s.chars().any(is_illegal_ctrl) {
                return Ok(false);
            }
        }
        if let V::Num(s) = &v {
            if s.chars().any(is_illegal_ctrl) {
                return Ok(false);
            }
        }
        // hyperlink
        if let Some(link) = link {
            if link.is_truthy()? {
                self.hyperlinks.bind(py).call_method1("append", (link,))?;
            }
        }
        self.start_c(row, col, sid, t_attr);
        match v {
            V::None => self.empty_close(),
            V::Str(s) if s.is_empty() => self.empty_close(),
            V::Str(s) => {
                self.buf.push(b'>');
                if data_type == "f" {
                    let mut it = s.chars();
                    it.next();
                    let rest = it.as_str();
                    if rest.is_empty() && !self.lxml {
                        // ElementTree: empty text -> short empty element
                        self.buf.extend_from_slice(b"<f />");
                    } else {
                        self.buf.extend_from_slice(b"<f>");
                        self.esc_text(rest);
                        self.buf.extend_from_slice(b"</f>");
                    }
                    self.empty_v();
                } else {
                    self.buf.extend_from_slice(b"<is><t");
                    let stripped = py_strip(&s);
                    let preserve = if self.lxml {
                        s != stripped
                    } else {
                        !stripped.is_empty() && s != stripped
                    };
                    if preserve {
                        self.buf.extend_from_slice(b" xml:space=\"preserve\"");
                    }
                    self.buf.push(b'>');
                    self.esc_text(&s);
                    self.buf.extend_from_slice(b"</t></is>");
                }
                self.buf.extend_from_slice(b"</c>");
            }
            V::Num(s) => {
                self.buf.push(b'>');
                if s.is_empty() {
                    self.empty_v();
                } else {
                    self.buf.extend_from_slice(b"<v>");
                    self.esc_text(&s);
                    self.buf.extend_from_slice(b"</v>");
                }
                self.buf.extend_from_slice(b"</c>");
            }
        }
        Ok(true)
    }

    /// `<row r=...>` with the row dimension's attributes; `row` is
    /// `f"{row_idx}"` (any int, as in openpyxl's write_row)
    fn write_row_start(
        &mut self,
        py: Python<'_>,
        row: &str,
        dims: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<()> {
        self.buf.extend_from_slice(b"<row");
        match dims {
            None => {
                self.buf.extend_from_slice(b" r=\"");
                self.esc_attr(row);
                self.buf.push(b'"');
            }
            Some(rd) => {
                let attrs = PyDict::new(py);
                attrs.set_item("r", row)?;
                attrs.call_method1("update", (rd,))?;
                for (k, v) in attrs.iter() {
                    let k: String = k.extract()?;
                    let v: String = v.extract()?;
                    self.buf.push(b' ');
                    self.esc_attr(&k);
                    self.buf.extend_from_slice(b"=\"");
                    self.esc_attr(&v);
                    self.buf.push(b'"');
                }
            }
        }
        self.buf.push(b'>');
        Ok(())
    }
}

enum Classified<'a> {
    None,
    Str(Cow<'a, str>),
    Num(Cow<'a, str>),
}

/// "%.16g" % value for an int value (converted to float first, as Python)
fn g16_int(i: i64) -> Cow<'static, str> {
    // exactly representable as a float, and < 1e16: printed with all digits
    if i.unsigned_abs() <= (1u64 << 53) {
        Cow::Owned(i.to_string())
    } else {
        Cow::Owned(format_g16(i as f64))
    }
}

/// "%.16g" % value for a finite float
fn g16_float(f: f64) -> Cow<'static, str> {
    if f == f.trunc() && f.abs() < 1e16 && !(f == 0.0 && f.is_sign_negative()) {
        Cow::Owned((f as i64).to_string())
    } else {
        Cow::Owned(format_g16(f))
    }
}

fn exact_int(o: &Bound<'_, PyAny>) -> Option<i64> {
    if o.is_exact_instance_of::<PyInt>() {
        o.extract::<i64>().ok()
    } else {
        None
    }
}

/// Plain values only: None, exact str, exact int, bool, exact float.
fn classify<'a>(
    py: Python<'_>,
    v: &'a Bound<'_, PyAny>,
    safe_string: &Py<PyAny>,
) -> PyResult<Option<Classified<'a>>> {
    let _ = safe_string;
    let _ = py;
    if v.is_none() {
        return Ok(Some(Classified::None));
    }
    if v.is_exact_instance_of::<PyString>() {
        return Ok(match v.cast::<PyString>()?.to_str() {
            Ok(s) => Some(Classified::Str(Cow::Borrowed(s))),
            Err(_) => None,
        });
    }
    if v.is_exact_instance_of::<PyBool>() {
        let b = v.is_truthy()?;
        return Ok(Some(Classified::Num(Cow::Borrowed(if b {
            "1"
        } else {
            "0"
        }))));
    }
    if v.is_exact_instance_of::<PyInt>() {
        return Ok(match v.extract::<i64>() {
            Ok(i) => {
                // "%.16g" % int converts to float first
                Some(Classified::Num(g16_int(i)))
            }
            Err(_) => None,
        });
    }
    if v.is_exact_instance_of::<PyFloat>() {
        let f: f64 = v.extract()?;
        if f.is_nan() || f.is_infinite() {
            return Ok(Some(Classified::Num(Cow::Borrowed(""))));
        }
        return Ok(Some(Classified::Num(g16_float(f))));
    }
    Ok(None)
}

impl SheetWriter {
    /// `<sheetData>` content for a CellStore, without materialising cells.
    fn write_store(
        &mut self,
        py: Python<'_>,
        store: &Bound<'_, CellStore>,
        dims: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let mut dim_rows: Vec<i64> = Vec::new();
        let mut generic_dims = false;
        for k in dims.call_method0("keys")?.try_iter()? {
            match exact_int(&k?) {
                Some(r) => dim_rows.push(r),
                None => generic_dims = true,
            }
        }
        if generic_dims {
            return Err(crate::NativeFallback::new_err("row dimension keys"));
        }
        dim_rows.sort_unstable();
        dim_rows.dedup();
        let cell_styles = store.borrow().cell_styles(py);
        let mut style_info: Vec<Option<(bool, Option<String>)>> = Vec::new();
        // rows of the base area, cells added after loading, row dimensions:
        // a three way merge in (row, column) order
        let gen = store.borrow().gen_sorted();
        let (base_rows, base_starts): (Vec<u32>, Vec<u32>) = {
            let st = store.borrow();
            let (r, s) = st.base_rows();
            (r.to_vec(), s.to_vec())
        };
        let (mut ri, mut gi, mut di) = (0usize, 0usize, 0usize);
        let mut row_cells: Vec<(u64, Loc)> = Vec::new();
        loop {
            let cands = [
                base_rows.get(ri).map(|&r| r as i64),
                gen.get(gi).map(|g| unpack(g.0).0),
                dim_rows.get(di).copied(),
            ];
            let Some(row_idx) = cands.iter().flatten().min().copied() else {
                break;
            };
            row_cells.clear();
            {
                let st = store.borrow();
                if base_rows.get(ri).map(|&r| r as i64) == Some(row_idx) {
                    for i in base_starts[ri] as usize..base_starts[ri + 1] as usize {
                        if let Some(c) = st.base_cell(i, row_idx as u32) {
                            row_cells.push(c);
                        }
                    }
                    ri += 1;
                }
            }
            let nbase = row_cells.len();
            while gi < gen.len() && unpack(gen[gi].0).0 == row_idx {
                row_cells.push(gen[gi]);
                gi += 1;
            }
            if row_cells.len() > nbase {
                row_cells.sort_by_key(|x| x.0);
            }
            let mut is_dim_row = false;
            while di < dim_rows.len() && dim_rows[di] == row_idx {
                di += 1;
                is_dim_row = true;
            }
            if row_cells.is_empty() && !is_dim_row {
                // a stored row whose cells were all deleted / moved
                continue;
            }
            let rd = if is_dim_row {
                Some(dims.call_method1("get", (row_idx,))?)
            } else {
                None
            };
            self.write_row_start(
                py,
                &row_idx.to_string(),
                rd.as_ref().filter(|r| !r.is_none()),
            )?;
            for &(_, loc) in &row_cells {
                let (e, obj) = {
                    let st = store.borrow();
                    let e = st.entry(loc);
                    let obj = if e.kind == crate::store::K_OBJ {
                        st.obj_payload(e.payload).map(|o| o.clone_ref(py))
                    } else {
                        None
                    };
                    (e, obj)
                };
                match obj {
                    Some(cell) => self.write_one(py, cell.bind(py))?,
                    None => self.write_lazy(py, store, loc, e, &cell_styles, &mut style_info)?,
                }
            }
            self.buf.extend_from_slice(b"</row>");
            self.flush_if_needed(py)?;
        }
        Ok(())
    }

    fn write_lazy(
        &mut self,
        py: Python<'_>,
        store: &Bound<'_, CellStore>,
        loc: Loc,
        e: crate::store::Entry,
        cell_styles: &Bound<'_, PyAny>,
        style_info: &mut Vec<Option<(bool, Option<String>)>>,
    ) -> PyResult<()> {
        let (row, col) = unpack(e.key);
        // has_style / style_id, computed once per style index
        let (styled, sid): (bool, Option<&str>) = if e.style == NO_STYLE {
            (false, None)
        } else {
            let idx = e.style as usize;
            if idx >= style_info.len() {
                style_info.resize(idx + 1, None);
            }
            if style_info[idx].is_none() {
                let stored = cell_styles.get_item(idx)?;
                let styled = self.has_style(py, &stored)?;
                let sid = if styled {
                    // a copy, like the `_style` of a Cell: IndexedList.add()
                    // finds the stored array itself by identity even when its
                    // hash is stale (openpyxl renumbers custom number formats
                    // in place after hashing), but appends an equal copy
                    let style = self.style_array_type.bind(py).call1((&stored,))?;
                    Some(self.style_id(py, &style)?.bind(py).str()?.to_string())
                } else {
                    None
                };
                style_info[idx] = Some((styled, sid));
            }
            let (styled, sid) = style_info[idx].as_ref().unwrap();
            (*styled, sid.as_deref())
        };
        let is_none = match e.kind {
            K_NONE => true,
            K_PYVAL => store.borrow().value_obj(py, loc)?.is_none(py),
            _ => false,
        };
        if is_none && !styled {
            return Ok(());
        }
        let materialize_and_write = |w: &mut Self| -> PyResult<()> {
            let cell = store.borrow_mut().materialize(py, loc)?;
            w.write_one(py, cell.bind(py))
        };
        if !(1..=18278).contains(&col) {
            return materialize_and_write(self);
        }
        let custom_dtype;
        let data_type: &str = if (e.dtype as usize) < DT_NAMES.len() {
            DT_NAMES[e.dtype as usize]
        } else {
            match store.borrow().dtype_str(py, e.dtype) {
                Some(d) => {
                    custom_dtype = d;
                    custom_dtype.as_str()
                }
                None => return materialize_and_write(self),
            }
        };
        let mut t_attr: Option<&str> = match data_type {
            "s" => Some("inlineStr"),
            "f" => None,
            other => Some(other),
        };
        if data_type == "d" {
            // native to_excel for dates read from the file (same epoch rules)
            let native = match (e.kind, &self.epoch_now) {
                (K_DATE_INT | K_DATE_FLOAT, Some(now)) if !self.iso_dates => {
                    let st = store.borrow();
                    st.excel_date(loc)
                        .and_then(|d| crate::dates::to_excel(&d, now))
                }
                _ => None,
            };
            let classified = match native {
                Some(v) => {
                    t_attr = Some("n");
                    Classified::Num(g16_float(v))
                }
                None => {
                    let value = store.borrow().value_obj(py, loc)?;
                    let res = self
                        .date_value_raw
                        .bind(py)
                        .call1((self.ws.bind(py), value))?;
                    let (dv, iso): (Bound<'_, PyAny>, bool) = res.extract()?;
                    if !iso {
                        t_attr = Some("n");
                    }
                    match classify(py, &dv, &self.safe_string)? {
                        Some(c) => owned(c),
                        None => return materialize_and_write(self),
                    }
                }
            };
            if !self.emit(py, row, col, data_type, t_attr, classified, sid, None)? {
                return materialize_and_write(self);
            }
            return Ok(());
        }
        let ok = {
            let st = store.borrow();
            let lv = st.lazy_value_e(py, loc, e)?;
            let pyval;
            let classified = match &lv {
                LazyValue::None => Some(Classified::None),
                LazyValue::Int(i) => Some(Classified::Num(g16_int(*i))),
                LazyValue::Float(f) => Some(Classified::Num(if f.is_finite() {
                    g16_float(*f)
                } else {
                    Cow::Borrowed("")
                })),
                LazyValue::Bool(b) => {
                    Some(Classified::Num(Cow::Borrowed(if *b { "1" } else { "0" })))
                }
                LazyValue::Str(s) => Some(Classified::Str(Cow::Borrowed(s))),
                LazyValue::Owned(s) => Some(Classified::Str(Cow::Borrowed(s.as_str()))),
                LazyValue::Py(o) => {
                    pyval = o.clone();
                    classify(py, &pyval, &self.safe_string)?
                }
            };
            match classified {
                Some(c) => self.emit(py, row, col, data_type, t_attr, c, sid, None)?,
                None => false,
            }
        };
        if !ok {
            return materialize_and_write(self);
        }
        Ok(())
    }
}

fn owned(c: Classified<'_>) -> Classified<'static> {
    match c {
        Classified::None => Classified::None,
        Classified::Str(s) => Classified::Str(Cow::Owned(s.into_owned())),
        Classified::Num(s) => Classified::Num(Cow::Owned(s.into_owned())),
    }
}

#[pymethods]
impl SheetWriter {
    #[new]
    fn new(
        py: Python<'_>,
        ws: &Bound<'_, PyAny>,
        write: Py<PyAny>,
        lxml: bool,
        helpers: &Bound<'_, PyDict>,
    ) -> PyResult<Self> {
        let g = |k: &str| -> PyResult<Py<PyAny>> {
            helpers
                .get_item(k)?
                .map(|v| v.unbind())
                .ok_or_else(|| PyTypeError::new_err(format!("missing helper {}", k)))
        };
        let i = |s: &str| PyString::intern(py, s).unbind();
        Ok(SheetWriter {
            ws: ws.clone().unbind(),
            write,
            lxml,
            buf: Vec::with_capacity(FLUSH_AT + 4096),
            names: Names {
                value: i("_value"),
                style: i("_style"),
                comment: i("_comment"),
                data_type: i("data_type"),
                hyperlink: i("hyperlink"),
                row: i("row"),
                column: i("column"),
            },
            cell_styles: ws.getattr("parent")?.getattr("_cell_styles")?.unbind(),
            style_cache: HashMap::new(),
            style_array_type: g("StyleArray")?,
            hyperlinks: ws.getattr("_hyperlinks")?.unbind(),
            add_comment: g("add_comment")?,
            fallback: g("write_cell")?,
            date_value: g("date_value")?,
            date_value_raw: g("date_value_raw")?,
            epoch_now: {
                let wb = ws.getattr("parent")?;
                let epoch = wb.getattr("epoch")?;
                let dtmod = py.import("datetime")?;
                let windows = py
                    .import("openrsxl.utils.datetime")?
                    .getattr("WINDOWS_EPOCH")?;
                if epoch.get_type().is(&dtmod.getattr("datetime")?)
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
                }
            },
            iso_dates: ws.getattr("parent")?.getattr("iso_dates")?.is_truthy()?,
            safe_string: g("safe_string")?,
        })
    }

    /// Write the whole `<sheetData>` element for `ws` (WorksheetWriter.write_rows).
    /// `rows_fallback` is `WorksheetWriter.rows` used when cell keys are not
    /// plain integer tuples.
    fn write_sheet_data(
        &mut self,
        py: Python<'_>,
        rows_fallback: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self.ws.clone_ref(py);
        let ws = ws.bind(py);
        let cells = ws.getattr("_cells")?;
        let dims = ws.getattr("row_dimensions")?;
        self.buf.extend_from_slice(b"<sheetData>");

        if let Ok(store) = cells.cast::<CellStore>() {
            if !store.borrow().in_dict_mode() {
                self.write_store(py, store, &dims)?;
                self.buf.extend_from_slice(b"</sheetData>");
                return self.flush_buf(py);
            }
        }

        // collect (row, col, cell)
        let mut entries: Vec<(i64, i64, Bound<'_, PyAny>)> = Vec::new();
        let mut ok = true;
        if let Ok(d) = cells.cast::<PyDict>() {
            entries.reserve(d.len());
            for (k, v) in d.iter() {
                let t = match k.cast::<PyTuple>() {
                    Ok(t) if t.len() == 2 => t,
                    _ => {
                        ok = false;
                        break;
                    }
                };
                match (exact_int(&t.get_item(0)?), exact_int(&t.get_item(1)?)) {
                    (Some(r), Some(c)) => entries.push((r, c, v)),
                    _ => {
                        ok = false;
                        break;
                    }
                }
            }
        } else {
            ok = false;
        }
        let mut dim_rows: Vec<i64> = Vec::new();
        if ok {
            for k in dims.call_method0("keys")?.try_iter()? {
                match exact_int(&k?) {
                    Some(r) => dim_rows.push(r),
                    None => {
                        ok = false;
                        break;
                    }
                }
            }
        }
        if ok {
            entries.sort_unstable_by_key(|e| (e.0, e.1));
            dim_rows.sort_unstable();
            let mut di = 0;
            let mut i = 0;
            while i < entries.len() || di < dim_rows.len() {
                let next_cell_row = entries.get(i).map(|e| e.0);
                let next_dim_row = dim_rows.get(di).copied();
                let row_idx = match (next_cell_row, next_dim_row) {
                    (Some(a), Some(b)) => a.min(b),
                    (Some(a), None) => a,
                    (None, Some(b)) => b,
                    (None, None) => unreachable!(),
                };
                while di < dim_rows.len() && dim_rows[di] == row_idx {
                    di += 1;
                }
                let rd = dims.call_method1("get", (row_idx,))?;
                self.write_row_start(
                    py,
                    &row_idx.to_string(),
                    if rd.is_none() { None } else { Some(&rd) },
                )?;
                while i < entries.len() && entries[i].0 == row_idx {
                    let cell = entries[i].2.clone();
                    self.write_one(py, &cell)?;
                    i += 1;
                }
                self.buf.extend_from_slice(b"</row>");
                self.flush_if_needed(py)?;
            }
        } else {
            for item in rows_fallback.call0()?.try_iter()? {
                let item = item?;
                let (row_idx, row): (Bound<'_, PyAny>, Bound<'_, PyAny>) = item.extract()?;
                // any int (beyond 64 bits too): f"{row_idx}" as openpyxl
                let ri = row_idx.str()?.to_string();
                let rd = dims.call_method1("get", (row_idx,))?;
                self.write_row_start(py, &ri, if rd.is_none() { None } else { Some(&rd) })?;
                for cell in row.try_iter()? {
                    self.write_one(py, &cell?)?;
                }
                self.buf.extend_from_slice(b"</row>");
                self.flush_if_needed(py)?;
            }
        }
        self.buf.extend_from_slice(b"</sheetData>");
        self.flush_buf(py)
    }

    /// Write one `<row>` from an iterable of cells (write-only worksheets).
    fn write_row(
        &mut self,
        py: Python<'_>,
        cells: &Bound<'_, PyAny>,
        row_idx: i64,
    ) -> PyResult<()> {
        let dims = self.ws.bind(py).getattr("row_dimensions")?;
        let rd = dims.call_method1("get", (row_idx,))?;
        self.write_row_start(
            py,
            &row_idx.to_string(),
            if rd.is_none() { None } else { Some(&rd) },
        )?;
        for cell in cells.try_iter()? {
            self.write_one(py, &cell?)?;
        }
        self.buf.extend_from_slice(b"</row>");
        self.flush_if_needed(py)
    }

    fn flush(&mut self, py: Python<'_>) -> PyResult<()> {
        self.flush_buf(py)
    }
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<SheetWriter>()?;
    let _ = PyList::empty(m.py());
    Ok(())
}
