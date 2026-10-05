//! openrsxl native engine.
//!
//! Python-facing functions used by the openrsxl package. Every function here
//! mirrors a specific piece of openpyxl's implementation; see the Python call
//! sites for the exact correspondence.

use pyo3::exceptions::PyException;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyTuple, PyType};
use pyo3::{create_exception, ffi};

mod dates;
mod formula;
mod ftemplate;
mod prescan;
mod pysource;
mod sheet;
mod slots;
mod store;
mod strings;
mod text;
mod utils;
mod writer;
mod xml;

use pysource::PyFileSource;
use sheet::{fallback_err, style_obj, Helpers, Item, Settings, SheetParser};
use slots::SlotClass;

create_exception!(_openrsxl, NativeFallback, PyException, "The native engine cannot reproduce openpyxl's behaviour for this input; the pure Python implementation must be used.");

/// read_string_table(source, helpers) -> list
#[pyfunction]
fn read_string_table(
    py: Python<'_>,
    source: Py<PyAny>,
    helpers: &Bound<'_, PyDict>,
) -> PyResult<Py<PyList>> {
    let h = Helpers::from_dict(helpers)?;
    let si_content = helpers.get_item("si_content")?.expect("si_content helper");
    strings::read_strings(py, PyFileSource::new(source), &h, &si_content)
        .map_err(|e| fallback_err(py, e))
}

/// Read all cells of a worksheet (full / editable mode) into `store`.
///
/// Mirrors `WorksheetReader.bind_cells` + `WorkSheetParser.parse`. Cells are
/// kept in compact form in the `CellStore` (calamine style) and become
/// `Cell` objects only when accessed. Other elements are appended to
/// `actions` as `("element", bytes, namespaces)` and warnings as
/// `("warn", message)` for the caller to replay in document order.
#[pyfunction]
#[allow(clippy::too_many_arguments)]
fn bind_cells(
    py: Python<'_>,
    source: Py<PyAny>,
    parser: &Bound<'_, PyAny>,
    ws: &Bound<'_, PyAny>,
    store: &Bound<'_, store::CellStore>,
    cell_styles: &Bound<'_, PyAny>,
    cell_cls: &Bound<'_, PyType>,
    helpers: &Bound<'_, PyDict>,
    actions: &Bound<'_, PyList>,
) -> PyResult<()> {
    let h = Helpers::from_dict(helpers)?;
    let st = Settings::from_parser(parser)?;
    let row_dims = parser.getattr("row_dimensions")?.cast_into::<PyDict>()?;
    let shared_strings = st.shared_strings.clone_ref(py);
    let epoch = st.fast_epoch.clone();
    let mut p = SheetParser::new(
        py,
        PyFileSource::new(source),
        h,
        st,
        actions.clone().unbind(),
        row_dims.unbind(),
    );
    let mut s = store.borrow_mut();
    store::setup_store(
        py,
        &mut s,
        ws,
        cell_styles,
        shared_strings.bind(py),
        epoch,
        cell_cls,
    )?;
    // _cell_styles does not change while the worksheet is read
    let n_styles = cell_styles.len()? as i64;

    #[cfg(feature = "prof")]
    let (mut t_parse, mut t_push) = (std::time::Duration::ZERO, std::time::Duration::ZERO);
    loop {
        #[cfg(feature = "prof")]
        let t0 = std::time::Instant::now();
        let item = p.next_item(py).map_err(|e| fallback_err(py, e))?;
        #[cfg(feature = "prof")]
        {
            t_parse += t0.elapsed();
        }
        match item {
            Item::Eof => break,
            Item::Element(snippet, ns) => {
                actions.append((
                    "element",
                    PyBytes::new(py, &snippet),
                    sheet::ns_list(py, &ns)?,
                ))?;
            }
            Item::Row(_, recs) => {
                #[cfg(feature = "prof")]
                let t1 = std::time::Instant::now();
                for c in recs {
                    // style = ws.parent._cell_styles[style_id] must succeed
                    let style = store::style_index(cell_styles, n_styles, &c.style_id)?;
                    s.push_rec(py, c, style)?;
                }
                #[cfg(feature = "prof")]
                {
                    t_push += t1.elapsed();
                }
            }
        }
    }
    #[cfg(feature = "prof")]
    eprintln!("bind_cells: parse {:?} push {:?}", t_parse, t_push);
    s.finish_loading(std::mem::take(&mut p.masters));
    Ok(())
}

/// Size of the blocks in which ElementTree.iterparse feeds expat. Events
/// (rows, elements) only become visible once the whole block containing
/// their end has been parsed without error: a well-formedness error anywhere
/// in a block hides all events completed in that block.
const FEED_BLOCK: u64 = 16 * 1024;

/// Streaming row reader used by ReadOnlyWorksheet.
#[pyclass]
struct RowReader {
    parser: SheetParser<PyFileSource>,
    #[pyo3(get)]
    actions: Py<PyList>,
    slots: Option<SlotClass>,
    slots_cls: Option<Py<PyType>>,
    /// parsed items with the absolute offset of their end
    queue: std::collections::VecDeque<(Item, u64)>,
    done: bool,
    failed: Option<PyErr>,
}

impl RowReader {
    /// Next item whose feed block has been completely parsed.
    fn next_ready(&mut self, py: Python<'_>) -> PyResult<Item> {
        loop {
            if let Some(err) = self.failed.take() {
                // everything not yet released is discarded; the caller falls
                // back to the Python implementation
                self.queue.clear();
                return Err(err);
            }
            if let Some((_, end)) = self.queue.front() {
                let block_end = (end / FEED_BLOCK + 1) * FEED_BLOCK;
                if self.done || self.parser.r.abs_pos() >= block_end {
                    return Ok(self.queue.pop_front().unwrap().0);
                }
            } else if self.done {
                return Ok(Item::Eof);
            }
            match self.parser.next_item(py) {
                Ok(Item::Eof) => self.done = true,
                Ok(item) => {
                    let end = self.parser.r.abs_pos();
                    self.queue.push_back((item, end));
                }
                Err(e) => {
                    let e = fallback_err(py, e);
                    // any error (malformed XML or a value openpyxl would choke
                    // on): let the Python implementation take over so that the
                    // exception and its timing are exactly openpyxl's
                    self.failed = Some(if e.is_instance_of::<NativeFallback>(py) {
                        e
                    } else {
                        NativeFallback::new_err(format!("native reader error: {}", e))
                    });
                }
            }
        }
    }
}

#[pymethods]
impl RowReader {
    #[new]
    fn new(
        py: Python<'_>,
        source: Py<PyAny>,
        parser: &Bound<'_, PyAny>,
        helpers: &Bound<'_, PyDict>,
    ) -> PyResult<Self> {
        let h = Helpers::from_dict(helpers)?;
        let st = Settings::from_parser(parser)?;
        let actions = PyList::empty(py).unbind();
        let row_dims = parser.getattr("row_dimensions")?.cast_into::<PyDict>()?;
        let p = SheetParser::new(
            py,
            PyFileSource::new(source),
            h,
            st,
            actions.clone_ref(py),
            row_dims.unbind(),
        );
        Ok(RowReader {
            parser: p,
            actions,
            slots: None,
            slots_cls: None,
            queue: Default::default(),
            done: false,
            failed: None,
        })
    }

    /// Next raw row: (idx, [(row, column, value, data_type, style_id), ...])
    /// or None at the end of the document.
    fn next_raw(&mut self, py: Python<'_>) -> PyResult<Option<(i64, Py<PyList>)>> {
        loop {
            match self.next_ready(py)? {
                Item::Eof => return Ok(None),
                Item::Element(snippet, ns) => {
                    self.actions.bind(py).append((
                        "element",
                        PyBytes::new(py, &snippet),
                        sheet::ns_list(py, &ns)?,
                    ))?;
                }
                Item::Row(idx, recs) => {
                    let l = PyList::empty(py);
                    for c in recs {
                        let v = self.parser.val_to_py(py, c.val)?;
                        let d = sheet::dtype_to_py(py, &self.parser.dt, c.dtype);
                        l.append((c.row, c.column, v, d, style_obj(py, &c.style_id)))?;
                    }
                    return Ok(Some((idx, l.unbind())));
                }
            }
        }
    }

    /// Mirrors ReadOnlyWorksheet._get_row for the next row in the source.
    /// Returns None at the end, else (idx, tuple_of_cells_or_values).
    #[allow(clippy::too_many_arguments)]
    #[pyo3(signature = (ws, min_col, max_col, values_only, cell_cls, filler, extra_none=None))]
    fn next_row(
        &mut self,
        py: Python<'_>,
        ws: &Bound<'_, PyAny>,
        min_col: i64,
        max_col: Option<i64>,
        values_only: bool,
        cell_cls: &Bound<'_, PyType>,
        filler: &Bound<'_, PyAny>,
        extra_none: Option<Vec<String>>,
    ) -> PyResult<Option<(i64, Py<PyAny>)>> {
        let recs = loop {
            match self.next_ready(py)? {
                Item::Eof => return Ok(None),
                Item::Element(snippet, ns) => {
                    self.actions.bind(py).append((
                        "element",
                        PyBytes::new(py, &snippet),
                        sheet::ns_list(py, &ns)?,
                    ))?;
                }
                Item::Row(idx, recs) => break (idx, recs),
            }
        };
        let (idx, recs) = recs;
        if recs.is_empty() && max_col.unwrap_or(0) == 0 {
            return Ok(Some((idx, PyTuple::empty(py).into_any().unbind())));
        }
        let max_col = match max_col {
            Some(m) if m != 0 => m,
            _ => recs.last().unwrap().column,
        };
        let width = max_col + 1 - min_col;
        let width = if width < 0 { 0 } else { width as usize };
        // openrsxl.extended: dual values and extra slots initialised to None
        let dual = self.parser.st.dual;
        let extra = extra_none.unwrap_or_default();
        if !values_only
            && (self.slots.is_none()
                || !self
                    .slots_cls
                    .as_ref()
                    .map(|c| c.bind(py).is(cell_cls))
                    .unwrap_or(false))
        {
            let mut names: Vec<&str> = vec![
                "parent",
                "row",
                "column",
                "_value",
                "data_type",
                "_style_id",
            ];
            if dual {
                names.push("formula");
                names.push("cached_value");
            }
            for n in &extra {
                names.push(n.as_str());
            }
            self.slots = Some(SlotClass::new(cell_cls, &names)?);
            self.slots_cls = Some(cell_cls.clone().unbind());
        }
        let mut items: Vec<Py<PyAny>> = (0..width).map(|_| filler.clone().unbind()).collect();
        let mut last_row: Option<(i64, Bound<'_, PyAny>)> = None;
        for c in recs {
            let counter = c.column;
            if min_col <= counter && counter <= max_col {
                let i = (counter - min_col) as usize;
                if values_only {
                    items[i] = self.parser.val_to_py(py, c.val)?;
                } else {
                    let value = self.parser.val_to_py(py, c.val)?;
                    let dtype = sheet::dtype_to_py(py, &self.parser.dt, c.dtype);
                    let slots = self.slots.as_ref().unwrap();
                    let row_obj = match &last_row {
                        Some((r, o)) if *r == c.row => o.clone(),
                        _ => {
                            let o = c.row.into_pyobject(py)?.into_any();
                            last_row = Some((c.row, o.clone()));
                            o
                        }
                    };
                    let mut vals: Vec<*mut ffi::PyObject> = Vec::with_capacity(8 + extra.len());
                    unsafe {
                        vals.push(ws.clone().into_ptr());
                        vals.push(row_obj.into_ptr());
                        vals.push(ffi::PyLong_FromLongLong(c.column));
                        if dual {
                            // formula: the data_only=False value of formula cells
                            let formula = if c.formula_primary {
                                value.clone_ref(py)
                            } else {
                                match c.formula {
                                    Some(f) => self.parser.val_to_py(py, f)?,
                                    None => py.None(),
                                }
                            };
                            // cached_value: the data_only=True value
                            let cached = match c.cached {
                                Some(v) => self.parser.val_to_py(py, v)?,
                                None => value.clone_ref(py),
                            };
                            vals.push(value.into_ptr());
                            vals.push(dtype.into_ptr());
                            vals.push(style_obj(py, &c.style_id).into_ptr());
                            vals.push(formula.into_ptr());
                            vals.push(cached.into_ptr());
                        } else {
                            vals.push(value.into_ptr());
                            vals.push(dtype.into_ptr());
                            vals.push(style_obj(py, &c.style_id).into_ptr());
                        }
                        for _ in &extra {
                            vals.push(py.None().into_ptr());
                        }
                    }
                    items[i] = slots.create(py, &vals)?.unbind();
                }
            }
        }
        Ok(Some((idx, PyTuple::new(py, items)?.into_any().unbind())))
    }
}

/// Find the `dimension` element like `WorkSheetParser.parse_dimensions`.
/// Returns (snippet, namespaces) or None.
#[pyfunction]
#[allow(clippy::type_complexity)]
fn find_dimension(
    py: Python<'_>,
    source: Py<PyAny>,
) -> PyResult<Option<(Py<PyBytes>, Py<PyList>)>> {
    use xml::{Ev, Reader, NS_MAIN};
    let mut r = Reader::new(PyFileSource::new(source));
    let res: Result<Option<(Vec<u8>, Vec<(String, String)>)>, xml::XmlError> = (|| loop {
        match r.next()? {
            Ev::Start | Ev::Empty => {
                if r.is(NS_MAIN, b"dimension") {
                    r.begin_capture();
                    text::skip_checked(&mut r, |ns, l| {
                        ns == NS_MAIN && (l == b"sheetData" || l == b"dimension")
                    })?;
                    let snip = r.end_capture();
                    return Ok(Some((snip, r.in_scope_namespaces())));
                }
            }
            Ev::End => {
                if r.is(NS_MAIN, b"sheetData") {
                    return Ok(None);
                }
            }
            Ev::Text => {}
            Ev::Eof => return Ok(None),
        }
    })();
    match res {
        Ok(None) => Ok(None),
        Ok(Some((s, ns))) => Ok(Some((
            PyBytes::new(py, &s).unbind(),
            sheet::ns_list(py, &ns)?.unbind(),
        ))),
        Err(e) => Err(fallback_err(py, sheet::SheetError::Xml(e))),
    }
}

#[pymodule]
fn _openrsxl(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("NativeFallback", m.py().get_type::<NativeFallback>())?;
    m.add_function(wrap_pyfunction!(read_string_table, m)?)?;
    m.add_function(wrap_pyfunction!(bind_cells, m)?)?;
    m.add_function(wrap_pyfunction!(find_dimension, m)?)?;
    m.add_function(wrap_pyfunction!(prescan::prescan_sheet, m)?)?;
    m.add_class::<RowReader>()?;
    m.add_class::<store::CellStore>()?;
    writer::register(m)?;
    Ok(())
}
