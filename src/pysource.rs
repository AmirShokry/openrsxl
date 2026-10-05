//! Byte sources backed by Python objects.

use pyo3::prelude::*;
use pyo3::types::PyBytes;

use crate::xml::{Source, XResult, XmlError};

const CHUNK: usize = 1 << 16; // 64 KiB, like expat feeding

/// Streams from a Python binary file-like object (anything with `read(n)`).
pub struct PyFileSource {
    obj: Py<PyAny>,
}

impl PyFileSource {
    pub fn new(obj: Py<PyAny>) -> Self {
        PyFileSource { obj }
    }
}

impl Source for PyFileSource {
    fn fill(&mut self, buf: &mut Vec<u8>) -> XResult<bool> {
        Python::attach(|py| {
            let data = self
                .obj
                .bind(py)
                .call_method1("read", (CHUNK,))
                .map_err(XmlError::Source)?;
            let bytes = data
                .cast::<PyBytes>()
                .map_err(|e| XmlError::Source(PyErr::from(e)))?;
            let b = bytes.as_bytes();
            if b.is_empty() {
                return Ok(false);
            }
            buf.extend_from_slice(b);
            Ok(true)
        })
    }
}
