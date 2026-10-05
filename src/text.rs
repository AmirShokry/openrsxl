//! Native implementation of `Text.from_tree(node).content` (used for shared
//! strings and inline strings) with exact fallback to Python for anything
//! unusual.

use std::collections::{HashMap, HashSet};

use crate::xml::{Ev, Reader, Source, XResult, XmlError, NS_MAIN, NS_NONE};

/// Predicate deciding whether an element (ns, local) found while skipping a
/// subtree would have been acted on by openpyxl's iterparse loop. If so the
/// native reader cannot reproduce openpyxl's behaviour.
pub type Forbid = fn(u32, &[u8]) -> bool;

pub fn forbid_none(_: u32, _: &[u8]) -> bool {
    false
}

pub fn forbid_si(ns: u32, local: &[u8]) -> bool {
    ns == NS_MAIN && local == b"si"
}

/// Skip the remainder of the current element (after Start/Empty).
pub fn skip_checked<S: Source>(r: &mut Reader<S>, forbid: Forbid) -> XResult<()> {
    if r.ev == Ev::Empty {
        r.next()?;
        return Ok(());
    }
    let depth = r.depth();
    loop {
        match r.next()? {
            Ev::Start | Ev::Empty => {
                if forbid(r.name_ns, r.local_name()) {
                    return Err(XmlError::Unsupported(
                        "nested element handled by dispatcher",
                    ));
                }
            }
            Ev::End => {
                if r.depth() < depth {
                    return Ok(());
                }
            }
            Ev::Text => {}
            Ev::Eof => return Err(XmlError::Malformed("unexpected eof")),
        }
    }
}

/// ElementTree `.text` of the current element, then skip to its end.
pub fn element_text_checked<S: Source>(
    r: &mut Reader<S>,
    forbid: Forbid,
) -> XResult<Option<String>> {
    if r.ev == Ev::Empty {
        r.next()?;
        return Ok(None);
    }
    let depth = r.depth();
    let mut text: Option<String> = None;
    let mut collecting = true;
    loop {
        match r.next()? {
            Ev::Text => {
                if collecting {
                    let t = r.text_str();
                    match text.as_mut() {
                        Some(s) => s.push_str(&t),
                        None => text = Some(t.into_owned()),
                    }
                }
            }
            Ev::Start | Ev::Empty => {
                collecting = false;
                if forbid(r.name_ns, r.local_name()) {
                    return Err(XmlError::Unsupported(
                        "nested element handled by dispatcher",
                    ));
                }
            }
            Ev::End => {
                if r.depth() < depth {
                    return Ok(text);
                }
            }
            Ev::Eof => return Err(XmlError::Malformed("unexpected eof")),
        }
    }
}

/// Like `element_text_checked` but appends to `out` (no allocation once the
/// buffer has grown). Empty text and missing text are not distinguished.
pub fn element_text_into<S: Source>(
    r: &mut Reader<S>,
    forbid: Forbid,
    out: &mut String,
) -> XResult<()> {
    if r.ev == Ev::Empty {
        r.next()?;
        return Ok(());
    }
    let depth = r.depth();
    let mut collecting = true;
    loop {
        match r.next()? {
            Ev::Text => {
                if collecting {
                    out.push_str(&r.text_str());
                }
            }
            Ev::Start | Ev::Empty => {
                collecting = false;
                if forbid(r.name_ns, r.local_name()) {
                    return Err(XmlError::Unsupported(
                        "nested element handled by dispatcher",
                    ));
                }
            }
            Ev::End => {
                if r.depth() < depth {
                    return Ok(());
                }
            }
            Ev::Eof => return Err(XmlError::Malformed("unexpected eof")),
        }
    }
}

pub struct TextCtx {
    /// attribute names of `Text` (other than t/r/rPh/phoneticPr) - a child
    /// with such a local name makes `Text.from_tree` misbehave
    pub text_dangerous: HashSet<Vec<u8>>,
    /// same for `RichText` (other than t/rPr)
    pub rich_dangerous: HashSet<Vec<u8>>,
    /// rPr snippets already validated (true = valid)
    pub rpr_cache: HashMap<Vec<u8>, bool>,
}

pub enum TextOutcome {
    Content(String),
    /// openpyxl's Python implementation must handle this element: the raw
    /// snippet plus the namespace declarations in scope.
    Fallback(Vec<u8>, Vec<(String, String)>),
}

/// Parse the current element (`si` or `is`) as `Text.from_tree(el).content`.
/// `validate_rpr` returns whether `InlineFont.from_tree` accepts the given
/// rPr snippet.
#[allow(clippy::type_complexity)]
pub fn parse_text<S: Source>(
    r: &mut Reader<S>,
    ctx: &mut TextCtx,
    forbid: Forbid,
    validate_rpr: &mut dyn FnMut(&[u8], Vec<(String, String)>) -> pyo3::PyResult<bool>,
) -> XResult<Result<TextOutcome, pyo3::PyErr>> {
    let mut fallback = r.attrs.iter().any(|a| a.ns == NS_NONE);
    r.begin_capture();
    let mut plain: Option<String> = None;
    let mut runs: Vec<Option<String>> = Vec::new();
    let mut pyerr: Option<pyo3::PyErr> = None;
    if r.ev == Ev::Start {
        let depth = r.depth();
        loop {
            match r.next()? {
                Ev::Start | Ev::Empty => {
                    if forbid(r.name_ns, r.local_name()) {
                        return Err(XmlError::Unsupported(
                            "nested element handled by dispatcher",
                        ));
                    }
                    let local = r.local_name().to_vec();
                    match local.as_slice() {
                        b"t" => plain = element_text_checked(r, forbid)?,
                        b"r" => {
                            if r.attrs.iter().any(|a| a.ns == NS_NONE) {
                                fallback = true;
                            }
                            let mut rt: Option<String> = None;
                            if r.ev == Ev::Start {
                                let rdepth = r.depth();
                                loop {
                                    match r.next()? {
                                        Ev::Start | Ev::Empty => {
                                            if forbid(r.name_ns, r.local_name()) {
                                                return Err(XmlError::Unsupported(
                                                    "nested element handled by dispatcher",
                                                ));
                                            }
                                            let l = r.local_name().to_vec();
                                            match l.as_slice() {
                                                b"t" => rt = element_text_checked(r, forbid)?,
                                                b"rPr" => {
                                                    let start = r.tag.start;
                                                    skip_checked(r, forbid)?;
                                                    let snippet = r.buf[start..r.tag.end].to_vec();
                                                    if !fallback && pyerr.is_none() {
                                                        let ok = match ctx.rpr_cache.get(&snippet) {
                                                            Some(v) => *v,
                                                            None => {
                                                                let ns = r.in_scope_namespaces();
                                                                match validate_rpr(&snippet, ns) {
                                                                    Ok(v) => {
                                                                        ctx.rpr_cache
                                                                            .insert(snippet, v);
                                                                        v
                                                                    }
                                                                    Err(e) => {
                                                                        pyerr = Some(e);
                                                                        true
                                                                    }
                                                                }
                                                            }
                                                        };
                                                        if !ok {
                                                            fallback = true;
                                                        }
                                                    }
                                                }
                                                other => {
                                                    if ctx.rich_dangerous.contains(other) {
                                                        fallback = true;
                                                    }
                                                    skip_checked(r, forbid)?;
                                                }
                                            }
                                        }
                                        Ev::End => {
                                            if r.depth() < rdepth {
                                                break;
                                            }
                                        }
                                        Ev::Text => {}
                                        Ev::Eof => {
                                            return Err(XmlError::Malformed("unexpected eof"))
                                        }
                                    }
                                }
                            } else {
                                r.next()?; // End of empty <r/>
                            }
                            runs.push(rt);
                        }
                        b"rPh" | b"phoneticPr" => {
                            fallback = true;
                            skip_checked(r, forbid)?;
                        }
                        other => {
                            if ctx.text_dangerous.contains(other) {
                                fallback = true;
                            }
                            skip_checked(r, forbid)?;
                        }
                    }
                }
                Ev::End => {
                    if r.depth() < depth {
                        break;
                    }
                }
                Ev::Text => {}
                Ev::Eof => return Err(XmlError::Malformed("unexpected eof")),
            }
        }
    } else {
        r.next()?; // End of empty element
    }
    if let Some(e) = pyerr {
        r.end_capture();
        return Ok(Err(e));
    }
    if fallback {
        let snippet = r.end_capture();
        return Ok(Ok(TextOutcome::Fallback(snippet, r.in_scope_namespaces())));
    }
    r.cancel_capture();
    let mut s = plain.unwrap_or_default();
    for t in runs.into_iter().flatten() {
        s.push_str(&t);
    }
    Ok(Ok(TextOutcome::Content(s)))
}
