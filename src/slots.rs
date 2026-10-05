//! Fast construction of instances of Python classes using `__slots__`.
//!
//! Objects are allocated with the type's allocator (no `__init__` call) and
//! their slots are written directly through the member descriptor offsets.
//! This produces objects that are indistinguishable from objects created by
//! the regular constructor followed by attribute assignment.

use pyo3::exceptions::PyTypeError;
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::types::PyType;

pub struct SlotClass {
    tp: Py<PyType>,
    offsets: Vec<isize>,
}

impl SlotClass {
    /// `names` are slot names; every one must be an object member descriptor
    /// reachable through the class.
    pub fn new(cls: &Bound<'_, PyType>, names: &[&str]) -> PyResult<Self> {
        let py = cls.py();
        let member_type = py.import("types")?.getattr("MemberDescriptorType")?;
        let mut offsets = Vec::with_capacity(names.len());
        for name in names {
            let desc = cls.getattr(*name)?;
            if !desc.get_type().is(&member_type) {
                return Err(PyTypeError::new_err(format!("{} is not a slot", name)));
            }
            let off = unsafe {
                let d = desc.as_ptr() as *mut ffi::PyMemberDescrObject;
                // CPython stores a `PyMemberDef *` here in every version;
                // pyo3-ffi declares it as `*mut PyGetSetDef` before 3.11
                let m = (*d).d_member.cast::<ffi::PyMemberDef>();
                if (*m).type_code != ffi::Py_T_OBJECT_EX {
                    return Err(PyTypeError::new_err(format!(
                        "{} has unexpected slot type",
                        name
                    )));
                }
                (*m).offset
            };
            offsets.push(off);
        }
        Ok(SlotClass {
            tp: cls.clone().unbind(),
            offsets,
        })
    }

    /// Allocate a new instance and fill the slots with the given (owned)
    /// values, in the order of the names given to `new`.
    #[inline]
    pub fn create<'py>(
        &self,
        py: Python<'py>,
        values: &[*mut ffi::PyObject],
    ) -> PyResult<Bound<'py, PyAny>> {
        debug_assert_eq!(values.len(), self.offsets.len());
        unsafe {
            let tp = self.tp.as_ptr() as *mut ffi::PyTypeObject;
            let alloc = (*tp).tp_alloc.unwrap_or(ffi::PyType_GenericAlloc);
            let obj = alloc(tp, 0);
            if obj.is_null() {
                for v in values {
                    ffi::Py_XDECREF(*v);
                }
                return Err(PyErr::fetch(py));
            }
            for (off, v) in self.offsets.iter().zip(values) {
                let p = (obj as *mut u8).offset(*off) as *mut *mut ffi::PyObject;
                *p = *v;
            }
            Ok(Bound::from_owned_ptr(py, obj))
        }
    }
}
