//! Fast pre-scan of a worksheet for openrsxl.extended's streaming reader.
//!
//! Collects, without materialising any cell:
//! * the top level elements before and after `<sheetData>` (raw snippets,
//!   parsed by openpyxl's own handlers in Python): merged cells, hyperlinks,
//!   conditional formatting, data validation, sheet properties, ...;
//! * optionally the attributes of `<row>` elements (row dimensions), with
//!   the row numbering rules of `WorkSheetParser.parse_row`;
//! * optionally the `s` attribute of a set of cells (corners of merged
//!   ranges), with the coordinate rules of `WorkSheetParser.parse_cell`.
//!
//! `<sheetData>` is skipped with byte searches (no tokenizing): anything the
//! raw scan cannot handle safely raises NativeFallback and the caller uses
//! the original Python parser instead.

use std::collections::{HashMap, HashSet};

use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};

use crate::pysource::PyFileSource;
use crate::sheet::{fallback_err, ns_list, SheetError};
use crate::text::{forbid_none, skip_checked};
use crate::utils::{fast_coordinate, fast_int};
use crate::xml::{
    decode_attr, validate_chars, validate_refs, Ev, Range, Reader, XmlError, NS_MAIN,
};

struct RawAttr {
    prefix: Range,
    local: Range,
    value: String,
}

fn is_ws(b: u8) -> bool {
    b == b' ' || b == b'\t' || b == b'\n' || b == b'\r'
}

/// Parse the attributes of a raw start tag (content after `<`, without the
/// closing `>`/`/>`).
fn parse_attrs(buf: &[u8], r: Range) -> Result<Vec<RawAttr>, XmlError> {
    let b = &buf[..r.end];
    let mut i = r.start;
    // element name
    while i < r.end && !is_ws(b[i]) {
        i += 1;
    }
    let mut out = Vec::new();
    loop {
        let ws_start = i;
        while i < r.end && is_ws(b[i]) {
            i += 1;
        }
        if i >= r.end {
            break;
        }
        if i == ws_start {
            return Err(XmlError::Malformed("missing whitespace"));
        }
        let ns = i;
        let mut colon = None;
        while i < r.end && b[i] != b'=' && !is_ws(b[i]) {
            if b[i] == b':' {
                if colon.is_some() {
                    return Err(XmlError::Unsupported("qname"));
                }
                colon = Some(i);
            } else if !(b[i].is_ascii_alphanumeric()
                || b[i] == b'_'
                || b[i] == b'-'
                || b[i] == b'.')
            {
                return Err(XmlError::Unsupported("attribute name"));
            }
            i += 1;
        }
        let ne = i;
        while i < r.end && is_ws(b[i]) {
            i += 1;
        }
        if i >= r.end || b[i] != b'=' {
            return Err(XmlError::Malformed("expected ="));
        }
        i += 1;
        while i < r.end && is_ws(b[i]) {
            i += 1;
        }
        if i >= r.end || (b[i] != b'"' && b[i] != b'\'') {
            return Err(XmlError::Malformed("expected quote"));
        }
        let q = b[i];
        let vs = i + 1;
        let ve = match b[vs..r.end].iter().position(|&c| c == q) {
            Some(o) => vs + o,
            None => return Err(XmlError::Malformed("unterminated value")),
        };
        let raw = &b[vs..ve];
        if raw.contains(&b'<') {
            return Err(XmlError::Malformed("< in attribute"));
        }
        // the tokenizer's own checks: characters, UTF-8 and every reference
        // (decode_attr relies on them)
        validate_chars(raw)?;
        if raw.contains(&b'&') {
            validate_refs(raw)?;
        }
        let value = decode_attr(raw).into_owned();
        let (prefix, local) = match colon {
            Some(c) => (
                Range { start: ns, end: c },
                Range {
                    start: c + 1,
                    end: ne,
                },
            ),
            None => (Range { start: ns, end: ns }, Range { start: ns, end: ne }),
        };
        let pfx = &b[prefix.start..prefix.end];
        if pfx == b"xmlns" || (colon.is_none() && &b[ns..ne] == b"xmlns") {
            return Err(XmlError::Unsupported("namespace declaration"));
        }
        out.push(RawAttr {
            prefix,
            local,
            value,
        });
        i = ve + 1;
    }
    Ok(out)
}

/// prescan_sheet(source, want_rows, corners, helpers) -> dict
#[pyfunction]
#[pyo3(signature = (source, want_rows, corners, helpers))]
pub fn prescan_sheet(
    py: Python<'_>,
    source: Py<PyAny>,
    want_rows: bool,
    corners: Option<&Bound<'_, PyAny>>,
    helpers: &Bound<'_, PyDict>,
) -> PyResult<Py<PyDict>> {
    let row_index = helpers
        .get_item("row_index")?
        .expect("row_index helper")
        .unbind();
    let mut wanted: HashSet<(i64, i64)> = HashSet::new();
    if let Some(c) = corners {
        for item in c.try_iter()? {
            wanted.insert(item?.extract::<(i64, i64)>()?);
        }
    }
    let want_cells = !wanted.is_empty();
    let wanted_rows: HashSet<i64> = wanted.iter().map(|x| x.0).collect();
    let mut r = Reader::new(PyFileSource::new(source));
    let head = PyList::empty(py);
    let tail = PyList::empty(py);
    let rows = PyList::empty(py);
    let mut found: HashMap<(i64, i64), Option<String>> = HashMap::new();
    let mut seen_data = false;
    let res: Result<(), SheetError> = (|| {
        loop {
            match r.next()? {
                Ev::Start | Ev::Empty => {
                    let d = r.depth();
                    if d == 1 {
                        continue;
                    }
                    if d == 2 && r.is(NS_MAIN, b"sheetData") && !seen_data {
                        seen_data = true;
                        if r.ev == Ev::Empty {
                            r.next()?;
                            continue;
                        }
                        // unprefixed name in the default (main) namespace only
                        let raw_name = &r.buf[r.tag.start + 1..];
                        let plain = raw_name.starts_with(b"sheetData")
                            && !raw_name[9].is_ascii_alphanumeric()
                            && raw_name[9] != b':';
                        if !plain || r.resolve_prefix(b"") != Some(NS_MAIN) {
                            return Err(SheetError::Xml(XmlError::Unsupported(
                                "prefixed sheetData",
                            )));
                        }
                        let mut row_counter: i64 = 0;
                        let mut col_counter: i64 = 0;
                        let mut pyerr: Option<PyErr> = None;
                        let ok = r.raw_skip(
                            b"sheetData",
                            want_rows || want_cells,
                            want_cells,
                            &mut |rd, is_row, range| {
                                if is_row && !want_rows {
                                    // fast path: only the row number is needed
                                    let t = &rd.buf[range.start..range.end];
                                    if t.len() > 7 && t.starts_with(b"row r=\"") {
                                        if let Some(e) = t[7..].iter().position(|&b| b == b'"') {
                                            if let Some(n) = fast_int(&t[7..7 + e]) {
                                                row_counter = n;
                                                col_counter = 0;
                                                return Ok(());
                                            }
                                        }
                                    }
                                }
                                if !is_row {
                                    // fast path: `<c r="A1"...` in a row without
                                    // requested cells needs no attribute parsing
                                    let t = &rd.buf[range.start..range.end];
                                    if t.len() > 5 && t.starts_with(b"c r=\"") {
                                        if let Some(e) = t[5..].iter().position(|&b| b == b'"') {
                                            if let Some((rr, cc)) = fast_coordinate(&t[5..5 + e]) {
                                                if !wanted_rows.contains(&rr) {
                                                    col_counter = cc;
                                                    return Ok(());
                                                }
                                            }
                                        }
                                    }
                                }
                                let attrs = parse_attrs(&rd.buf, range)?;
                                if is_row {
                                    let mut r_attr: Option<&str> = None;
                                    let mut need = false;
                                    for a in &attrs {
                                        if a.prefix.start == a.prefix.end {
                                            let l = &rd.buf[a.local.start..a.local.end];
                                            if l == b"r" {
                                                r_attr = Some(&a.value);
                                            } else if l != b"spans" {
                                                need = true;
                                            }
                                        }
                                    }
                                    row_counter = match r_attr {
                                        Some(v) => match fast_int(v.as_bytes()) {
                                            Some(n) => n,
                                            None => match row_index.bind(py).call1((v,)) {
                                                // beyond 64 bits: the Python scan
                                                Ok(o) => o.extract::<i64>().map_err(|_| {
                                                    XmlError::Unsupported("row index")
                                                })?,
                                                Err(e) => {
                                                    // openpyxl's own error (eg. int("x"))
                                                    pyerr = Some(e);
                                                    return Err(XmlError::Unsupported("row index"));
                                                }
                                            },
                                        },
                                        None => row_counter
                                            .checked_add(1)
                                            .ok_or(XmlError::Unsupported("row index"))?,
                                    };
                                    col_counter = 0;
                                    if need && want_rows {
                                        let l = PyList::empty(py);
                                        for a in &attrs {
                                            let local = unsafe {
                                                std::str::from_utf8_unchecked(
                                                    &rd.buf[a.local.start..a.local.end],
                                                )
                                            };
                                            let key = if a.prefix.start == a.prefix.end {
                                                local.to_string()
                                            } else {
                                                let id = rd
                                                    .resolve_prefix(
                                                        &rd.buf[a.prefix.start..a.prefix.end],
                                                    )
                                                    .ok_or(XmlError::Malformed("unbound prefix"))?;
                                                format!("{{{}}}{}", rd.uri(id), local)
                                            };
                                            l.append((key, a.value.as_str()))
                                                .map_err(|_| XmlError::Unsupported("append"))?;
                                        }
                                        rows.append((row_counter, l))
                                            .map_err(|_| XmlError::Unsupported("append"))?;
                                    }
                                } else {
                                    let mut coord: Option<&str> = None;
                                    let mut s_attr: Option<&str> = None;
                                    for a in &attrs {
                                        if a.prefix.start == a.prefix.end {
                                            match &rd.buf[a.local.start..a.local.end] {
                                                b"r" => coord = Some(&a.value),
                                                b"s" => s_attr = Some(&a.value),
                                                _ => {}
                                            }
                                        }
                                    }
                                    let (row, col) = match coord {
                                        Some(c) if !c.is_empty() => {
                                            let (rr, cc) = fast_coordinate(c.as_bytes())
                                                .ok_or(XmlError::Unsupported("coordinate"))?;
                                            col_counter = cc;
                                            (rr, cc)
                                        }
                                        _ => {
                                            col_counter = col_counter
                                                .checked_add(1)
                                                .ok_or(XmlError::Unsupported("column index"))?;
                                            (row_counter, col_counter)
                                        }
                                    };
                                    if wanted.contains(&(row, col)) {
                                        found.insert((row, col), s_attr.map(|s| s.to_string()));
                                    }
                                }
                                Ok(())
                            },
                        );
                        if let Some(e) = pyerr {
                            return Err(SheetError::Py(e));
                        }
                        if !ok? {
                            return Err(SheetError::Xml(XmlError::Unsupported(
                                "comment/cdata/pi in sheetData",
                            )));
                        }
                        continue;
                    }
                    // a top level element (children of other elements are
                    // inside the captured snippets)
                    r.begin_capture();
                    skip_checked(&mut r, forbid_none)?;
                    let snippet = r.end_capture();
                    let ns = ns_list(py, &r.in_scope_namespaces())?;
                    let item = (PyBytes::new(py, &snippet), ns);
                    if seen_data {
                        tail.append(item)?
                    } else {
                        head.append(item)?
                    }
                }
                Ev::End | Ev::Text => {}
                Ev::Eof => return Ok(()),
            }
        }
    })();
    res.map_err(|e| fallback_err(py, e))?;
    let out = PyDict::new(py);
    out.set_item("head", head)?;
    out.set_item("tail", tail)?;
    out.set_item("rows", rows)?;
    let corners_out = PyDict::new(py);
    for ((row, col), s) in found {
        corners_out.set_item((row, col), s)?;
    }
    out.set_item("corners", corners_out)?;
    Ok(out.unbind())
}
