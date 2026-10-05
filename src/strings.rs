//! Shared string table reader (`openpyxl.reader.strings.read_string_table`).

use std::collections::HashMap;

use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyList, PyString};

use crate::sheet::{ns_list, Helpers, SResult, SheetError};
use crate::text::{forbid_si, parse_text, TextCtx, TextOutcome};
use crate::xml::{Ev, Reader, Source, XmlError, NS_MAIN};

pub fn read_strings<S: Source>(
    py: Python<'_>,
    src: S,
    h: &Helpers,
    si_content: &Bound<'_, PyAny>,
) -> SResult<Py<PyList>> {
    let mut r = Reader::new(src);
    let out = PyList::empty(py);
    let mut ctx = TextCtx {
        text_dangerous: h.text_dangerous.clone(),
        rich_dangerous: h.rich_dangerous.clone(),
        rpr_cache: HashMap::new(),
    };
    let validate = h.validate_rpr.clone_ref(py);
    let mut vf = |snip: &[u8], ns: Vec<(String, String)>| -> PyResult<bool> {
        let l = ns_list(py, &ns)?;
        validate
            .bind(py)
            .call1((PyBytes::new(py, snip), l))?
            .is_truthy()
    };
    loop {
        match r.next()? {
            Ev::Start | Ev::Empty => {
                if r.is(NS_MAIN, b"si") {
                    if r.depth() != 2 {
                        return Err(SheetError::Xml(XmlError::Unsupported("nested si")));
                    }
                    match parse_text(&mut r, &mut ctx, forbid_si, &mut vf)? {
                        Err(e) => return Err(e.into()),
                        Ok(TextOutcome::Content(s)) => {
                            let s = if s.contains("x005F_") {
                                s.replace("x005F_", "")
                            } else {
                                s
                            };
                            out.append(PyString::new(py, &s))?;
                        }
                        Ok(TextOutcome::Fallback(snippet, ns)) => {
                            let l = ns_list(py, &ns)?;
                            out.append(si_content.call1((PyBytes::new(py, &snippet), l))?)?;
                        }
                    }
                }
            }
            Ev::End | Ev::Text => {}
            Ev::Eof => break,
        }
    }
    Ok(out.unbind())
}
