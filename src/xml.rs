//! A small, strict, streaming XML 1.0 tokenizer with namespace support.
//!
//! The goal is *not* to be a general purpose XML library but to reproduce
//! exactly what Python's `xml.etree.ElementTree` (expat) reports for the
//! documents found inside xlsx packages:
//!
//! * line-ending normalisation (`\r\n` and `\r` -> `\n`) in text,
//! * attribute value normalisation (whitespace characters -> space),
//! * the five predefined entities and character references,
//! * namespace resolution (`{uri}local` names),
//! * CDATA sections are text, comments and processing instructions are
//!   invisible.
//!
//! Anything this tokenizer is not 100% sure about (DOCTYPE declarations,
//! non UTF-8 encodings, non-ASCII names, any well-formedness problem) is
//! reported as an error.  Callers then fall back to the original Python
//! implementation, which guarantees identical behaviour (including the exact
//! exception raised for malformed documents).

use std::borrow::Cow;
use std::collections::HashMap;

use memchr::{memchr, memchr2, memchr3, memmem};

pub const NS_NONE: u32 = 0;
pub const NS_XML: u32 = 1;
pub const NS_MAIN: u32 = 2;

pub const XML_NS_URI: &str = "http://www.w3.org/XML/1998/namespace";
pub const MAIN_NS_URI: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[derive(Debug)]
pub enum XmlError {
    /// The document is (or might be) not well-formed.
    Malformed(&'static str),
    /// The document uses a feature we do not handle natively.
    Unsupported(&'static str),
    /// Reading from the underlying source failed.
    Source(pyo3::PyErr),
}

pub type XResult<T> = Result<T, XmlError>;

/// Provider of raw bytes.
pub trait Source {
    /// Append more bytes to `buf`. Returns `false` at end of input.
    fn fill(&mut self, buf: &mut Vec<u8>) -> XResult<bool>;
}

/// The whole document is already in memory.
#[allow(dead_code)]
pub struct SliceSource<'a> {
    data: &'a [u8],
    done: bool,
}

#[allow(dead_code)]
impl<'a> SliceSource<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        SliceSource { data, done: false }
    }
}

impl<'a> Source for SliceSource<'a> {
    fn fill(&mut self, buf: &mut Vec<u8>) -> XResult<bool> {
        if self.done {
            return Ok(false);
        }
        self.done = true;
        buf.extend_from_slice(self.data);
        Ok(!self.data.is_empty())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Ev {
    Start,
    Empty,
    End,
    Text,
    Eof,
}

#[derive(Clone, Copy, Debug)]
pub struct Range {
    pub start: usize,
    pub end: usize,
}

#[derive(Clone, Debug)]
pub struct Attr {
    pub ns: u32,
    /// raw qualified name
    pub qname: Range,
    /// local part of the name
    pub local: Range,
    /// raw (undecoded) value, without quotes
    pub value: Range,
    pub needs_decode: bool,
}

struct OpenElem {
    /// qualified name, stored in `Reader::names`
    name_start: u32,
    name_len: u32,
    ns_bindings: u32,
}

pub struct Reader<S: Source> {
    src: S,
    pub buf: Vec<u8>,
    pos: usize,
    eof: bool,
    started: bool,
    stack: Vec<OpenElem>,
    /// qualified names of the open elements, concatenated
    names: Vec<u8>,
    root_closed: bool,
    seen_root: bool,
    // namespaces
    uris: Vec<String>,
    uri_ids: HashMap<Vec<u8>, u32>,
    /// (prefix, uri id); prefix empty = default namespace
    bindings: Vec<(Vec<u8>, u32)>,
    // current event data
    pub ev: Ev,
    pub name_ns: u32,
    pub name_local: Range,
    pub tag: Range,
    pub attrs: Vec<Attr>,
    pub text: Range,
    pub text_cdata: bool,
    pub text_needs_decode: bool,
    capture_start: Option<usize>,
    pending_end: bool,
    /// number of bytes drained from the front of `buf`
    base_offset: u64,
    scratch_names: Vec<(u32, Range)>,
}

const COMPACT_THRESHOLD: usize = 1 << 16;

#[inline]
fn is_ws(b: u8) -> bool {
    b == b' ' || b == b'\t' || b == b'\n' || b == b'\r'
}

#[inline]
fn is_name_start(b: u8) -> bool {
    b.is_ascii_alphabetic() || b == b'_' || b == b':'
}

#[inline]
fn is_name_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b':' || b == b'-' || b == b'.'
}

/// Validate a run of character data (text, attribute value, comment ...).
/// Returns whether it contains '&' (needs entity decoding) or '\r'.
pub(crate) fn validate_chars(data: &[u8]) -> XResult<()> {
    // fast ASCII path
    let mut non_ascii = false;
    for &b in data {
        if b < 0x20 {
            if b != b'\t' && b != b'\n' && b != b'\r' {
                return Err(XmlError::Malformed("invalid character"));
            }
        } else if b >= 0x80 {
            non_ascii = true;
        }
    }
    if non_ascii {
        let s = std::str::from_utf8(data).map_err(|_| XmlError::Malformed("invalid utf-8"))?;
        // U+FFFE and U+FFFF are not XML characters
        if memmem::find(data, b"\xEF\xBF\xBE").is_some()
            || memmem::find(data, b"\xEF\xBF\xBF").is_some()
        {
            return Err(XmlError::Malformed("invalid character"));
        }
        let _ = s;
    }
    Ok(())
}

#[inline]
fn is_xml_char(c: u32) -> bool {
    c == 0x9
        || c == 0xA
        || c == 0xD
        || (0x20..=0xD7FF).contains(&c)
        || (0xE000..=0xFFFD).contains(&c)
        || (0x10000..=0x10FFFF).contains(&c)
}

/// Check the entity references in `data` (which contains at least one '&').
pub(crate) fn validate_refs(data: &[u8]) -> XResult<()> {
    let mut i = 0;
    while let Some(off) = memchr(b'&', &data[i..]) {
        let start = i + off + 1;
        let semi = memchr(b';', &data[start..]).ok_or(XmlError::Malformed("bad reference"))?;
        let name = &data[start..start + semi];
        resolve_ref(name)?;
        i = start + semi + 1;
    }
    Ok(())
}

fn resolve_ref(name: &[u8]) -> XResult<char> {
    match name {
        b"lt" => Ok('<'),
        b"gt" => Ok('>'),
        b"amp" => Ok('&'),
        b"apos" => Ok('\''),
        b"quot" => Ok('"'),
        _ => {
            if name.len() >= 2 && name[0] == b'#' {
                let (digits, radix) = if name[1] == b'x' {
                    (&name[2..], 16)
                } else {
                    (&name[1..], 10)
                };
                if digits.is_empty() || digits.len() > 10 {
                    return Err(XmlError::Malformed("bad char ref"));
                }
                let mut v: u32 = 0;
                for &d in digits {
                    let dv = match d {
                        b'0'..=b'9' => (d - b'0') as u32,
                        b'a'..=b'f' if radix == 16 => (d - b'a' + 10) as u32,
                        b'A'..=b'F' if radix == 16 => (d - b'A' + 10) as u32,
                        _ => return Err(XmlError::Malformed("bad char ref")),
                    };
                    v = v
                        .checked_mul(radix)
                        .and_then(|v| v.checked_add(dv))
                        .ok_or(XmlError::Malformed("bad char ref"))?;
                }
                if !is_xml_char(v) {
                    return Err(XmlError::Malformed("invalid char ref"));
                }
                char::from_u32(v).ok_or(XmlError::Malformed("invalid char ref"))
            } else {
                // undefined entity (no DTD support)
                Err(XmlError::Malformed("undefined entity"))
            }
        }
    }
}

/// Decode character data: normalise line endings and resolve references.
/// The data must have been validated.
pub fn decode_text(raw: &[u8], cdata: bool) -> Cow<'_, str> {
    let has_amp = !cdata && memchr(b'&', raw).is_some();
    let has_cr = memchr(b'\r', raw).is_some();
    if !has_amp && !has_cr {
        // SAFETY: validated as utf-8 during tokenization
        return Cow::Borrowed(unsafe { std::str::from_utf8_unchecked(raw) });
    }
    let mut out = String::with_capacity(raw.len());
    let mut i = 0;
    let n = raw.len();
    while i < n {
        let next = if cdata {
            memchr(b'\r', &raw[i..])
        } else {
            memchr2(b'\r', b'&', &raw[i..])
        };
        match next {
            None => {
                out.push_str(unsafe { std::str::from_utf8_unchecked(&raw[i..]) });
                break;
            }
            Some(off) => {
                out.push_str(unsafe { std::str::from_utf8_unchecked(&raw[i..i + off]) });
                let j = i + off;
                if raw[j] == b'\r' {
                    out.push('\n');
                    i = j + 1;
                    if i < n && raw[i] == b'\n' {
                        i += 1;
                    }
                } else {
                    let semi = memchr(b';', &raw[j + 1..]).unwrap();
                    let c = resolve_ref(&raw[j + 1..j + 1 + semi]).unwrap();
                    out.push(c);
                    i = j + 1 + semi + 1;
                }
            }
        }
    }
    Cow::Owned(out)
}

/// Decode an attribute value: normalise whitespace and resolve references.
pub fn decode_attr(raw: &[u8]) -> Cow<'_, str> {
    let needs = raw
        .iter()
        .any(|&b| b == b'&' || b == b'\r' || b == b'\n' || b == b'\t');
    if !needs {
        return Cow::Borrowed(unsafe { std::str::from_utf8_unchecked(raw) });
    }
    let mut out = String::with_capacity(raw.len());
    let mut i = 0;
    let n = raw.len();
    while i < n {
        let b = raw[i];
        match b {
            b'\r' => {
                out.push(' ');
                i += 1;
                if i < n && raw[i] == b'\n' {
                    i += 1;
                }
            }
            b'\n' | b'\t' => {
                out.push(' ');
                i += 1;
            }
            b'&' => {
                let semi = memchr(b';', &raw[i + 1..]).unwrap();
                let c = resolve_ref(&raw[i + 1..i + 1 + semi]).unwrap();
                out.push(c);
                i = i + 1 + semi + 1;
            }
            _ => {
                // copy a run of ordinary bytes
                let mut j = i + 1;
                while j < n && !matches!(raw[j], b'\r' | b'\n' | b'\t' | b'&') {
                    j += 1;
                }
                out.push_str(unsafe { std::str::from_utf8_unchecked(&raw[i..j]) });
                i = j;
            }
        }
    }
    Cow::Owned(out)
}

impl<S: Source> Reader<S> {
    pub fn new(src: S) -> Self {
        let mut uri_ids = HashMap::new();
        uri_ids.insert(Vec::new(), NS_NONE);
        uri_ids.insert(XML_NS_URI.as_bytes().to_vec(), NS_XML);
        uri_ids.insert(MAIN_NS_URI.as_bytes().to_vec(), NS_MAIN);
        Reader {
            src,
            buf: Vec::with_capacity(1 << 16),
            pos: 0,
            eof: false,
            started: false,
            stack: Vec::new(),
            names: Vec::with_capacity(256),
            root_closed: false,
            seen_root: false,
            uris: vec![
                String::new(),
                XML_NS_URI.to_string(),
                MAIN_NS_URI.to_string(),
            ],
            uri_ids,
            bindings: Vec::new(),
            ev: Ev::Eof,
            name_ns: 0,
            name_local: Range { start: 0, end: 0 },
            tag: Range { start: 0, end: 0 },
            attrs: Vec::new(),
            text: Range { start: 0, end: 0 },
            text_cdata: false,
            text_needs_decode: false,
            capture_start: None,
            pending_end: false,
            base_offset: 0,
            scratch_names: Vec::new(),
        }
    }

    /// Absolute offset (in the decoded stream) of the parse position.
    pub fn abs_pos(&self) -> u64 {
        self.base_offset + self.pos as u64
    }

    pub fn depth(&self) -> usize {
        self.stack.len()
    }

    #[inline]
    pub fn slice(&self, r: Range) -> &[u8] {
        &self.buf[r.start..r.end]
    }

    #[inline]
    pub fn local_name(&self) -> &[u8] {
        &self.buf[self.name_local.start..self.name_local.end]
    }

    #[inline]
    pub fn is(&self, ns: u32, local: &[u8]) -> bool {
        self.name_ns == ns && self.local_name() == local
    }

    /// Current in-scope namespace declarations (prefix, uri) - later entries
    /// override earlier ones.
    pub fn in_scope_namespaces(&self) -> Vec<(String, String)> {
        let mut out: Vec<(String, String)> = Vec::new();
        for (p, id) in &self.bindings {
            let p = String::from_utf8_lossy(p).into_owned();
            let u = self.uris[*id as usize].clone();
            if let Some(e) = out.iter_mut().find(|(q, _)| *q == p) {
                e.1 = u;
            } else {
                out.push((p, u));
            }
        }
        out
    }

    /// Text of the current Text event, decoded.
    pub fn text_str(&self) -> Cow<'_, str> {
        let raw = &self.buf[self.text.start..self.text.end];
        if !self.text_needs_decode {
            return Cow::Borrowed(unsafe { std::str::from_utf8_unchecked(raw) });
        }
        decode_text(raw, self.text_cdata)
    }

    pub fn attr_value(&self, a: &Attr) -> Cow<'_, str> {
        let raw = &self.buf[a.value.start..a.value.end];
        if !a.needs_decode {
            return Cow::Borrowed(unsafe { std::str::from_utf8_unchecked(raw) });
        }
        decode_attr(raw)
    }

    /// ElementTree style key for an attribute: `local` or `{uri}local`
    pub fn attr_key(&self, a: &Attr) -> String {
        let local = unsafe { std::str::from_utf8_unchecked(&self.buf[a.local.start..a.local.end]) };
        if a.ns == NS_NONE {
            local.to_string()
        } else {
            format!("{{{}}}{}", self.uris[a.ns as usize], local)
        }
    }

    /// Start capturing raw bytes at the current start tag.
    pub fn begin_capture(&mut self) {
        self.capture_start = Some(self.tag.start);
    }

    /// Stop capturing; returns the bytes from the captured start tag to the
    /// end of the current event.
    pub fn end_capture(&mut self) -> Vec<u8> {
        let start = self.capture_start.take().expect("capture not started");
        self.buf[start..self.tag.end].to_vec()
    }

    /// Stop capturing without copying.
    pub fn cancel_capture(&mut self) {
        self.capture_start = None;
    }

    fn more(&mut self) -> XResult<bool> {
        if self.eof {
            return Ok(false);
        }
        // compact
        let keep = match self.capture_start {
            Some(c) => c.min(self.pos),
            None => self.pos,
        };
        if keep >= COMPACT_THRESHOLD && keep * 2 >= self.buf.len() {
            self.buf.drain(..keep);
            self.base_offset += keep as u64;
            self.pos -= keep;
            if let Some(c) = self.capture_start.as_mut() {
                *c -= keep;
            }
            self.tag.start = self.tag.start.saturating_sub(keep);
            self.tag.end = self.tag.end.saturating_sub(keep);
        }
        let got = self.src.fill(&mut self.buf)?;
        if !got {
            self.eof = true;
        }
        Ok(got)
    }

    /// Ensure that at least `n` bytes are available from `self.pos`
    fn ensure(&mut self, n: usize) -> XResult<bool> {
        while self.buf.len() - self.pos < n {
            if !self.more()? {
                return Ok(false);
            }
        }
        Ok(true)
    }

    /// Find `needle` at or after `from` (index >= self.pos), reading more
    /// input as required. Indices are relative to the (possibly compacted)
    /// buffer after the call.
    fn find_seq_from(&mut self, from: usize, needle: &[u8]) -> XResult<Option<usize>> {
        let mut scan = from;
        loop {
            if let Some(off) = memmem::find(&self.buf[scan..], needle) {
                return Ok(Some(scan + off));
            }
            let len = self.buf.len();
            let pos_before = self.pos;
            if !self.more()? {
                return Ok(None);
            }
            let shift = pos_before - self.pos;
            let from_shifted = from - shift;
            scan = (len - shift)
                .saturating_sub(needle.len() - 1)
                .max(from_shifted);
        }
    }

    /// Advance to the next event.
    pub fn next(&mut self) -> XResult<Ev> {
        if self.pending_end {
            // the End half of an empty element tag
            self.pending_end = false;
            self.pop_element();
            self.ev = Ev::End;
            return Ok(Ev::End);
        }
        if !self.started {
            self.started = true;
            self.prolog()?;
        }
        loop {
            if self.pos >= self.buf.len() && !self.more()? {
                return self.at_eof();
            }
            if self.buf[self.pos] == b'<' {
                if !self.ensure(2)? {
                    return Err(XmlError::Malformed("unclosed token"));
                }
                match self.buf[self.pos + 1] {
                    b'/' => return self.end_tag(),
                    b'!' => {
                        if self.markup_decl()? {
                            return Ok(Ev::Text);
                        }
                    }
                    b'?' => self.pi()?,
                    _ => return self.start_tag(),
                }
            } else {
                // character data: self.pos stays at the start of the text
                // while scanning so that compaction keeps it
                let end = self.scan_text_end()?;
                let start = self.pos;
                if self.stack.is_empty() {
                    // outside the root element only whitespace is allowed
                    // (handled here, in the loop: no recursion however many
                    // comments / PIs surround the root)
                    if !self.buf[start..end].iter().all(|&b| is_ws(b)) {
                        return Err(XmlError::Malformed("junk outside root"));
                    }
                    self.pos = end;
                    continue;
                }
                return self.emit_text(start, end);
            }
        }
    }

    /// Scan for the end of character data starting at `self.pos`, reading
    /// input as required.
    fn scan_text_end(&mut self) -> XResult<usize> {
        let mut scan = self.pos;
        loop {
            if let Some(off) = memchr(b'<', &self.buf[scan..]) {
                return Ok(scan + off);
            }
            let len = self.buf.len();
            let before = self.pos;
            if !self.more()? {
                return Ok(self.buf.len());
            }
            let shift = before - self.pos;
            scan = len - shift;
        }
    }

    fn emit_text(&mut self, start: usize, end: usize) -> XResult<Ev> {
        let data = &self.buf[start..end];
        validate_chars(data)?;
        let has_amp = memchr(b'&', data).is_some();
        if has_amp {
            validate_refs(data)?;
        }
        if memmem::find(data, b"]]>").is_some() {
            return Err(XmlError::Malformed("]]> in text"));
        }
        self.text_needs_decode = has_amp || memchr(b'\r', data).is_some();
        self.text = Range { start, end };
        self.text_cdata = false;
        self.pos = end;
        self.ev = Ev::Text;
        Ok(Ev::Text)
    }

    fn at_eof(&mut self) -> XResult<Ev> {
        if !self.stack.is_empty() {
            return Err(XmlError::Malformed("unclosed element"));
        }
        if !self.seen_root {
            return Err(XmlError::Malformed("no element found"));
        }
        self.ev = Ev::Eof;
        Ok(Ev::Eof)
    }

    /// Handle BOM and XML declaration
    fn prolog(&mut self) -> XResult<()> {
        self.ensure(4)?;
        let b = &self.buf[self.pos..];
        if b.starts_with(b"\xEF\xBB\xBF") {
            self.pos += 3;
        } else if b.starts_with(b"\xFE\xFF")
            || b.starts_with(b"\xFF\xFE")
            || b.first() == Some(&0)
            || b.get(1) == Some(&0)
        {
            return Err(XmlError::Unsupported("utf-16"));
        }
        self.ensure(6)?;
        if self.buf[self.pos..].starts_with(b"<?xml") {
            let after = self.buf.get(self.pos + 5).copied();
            if after.map(is_ws).unwrap_or(false) || after == Some(b'?') {
                let end = self
                    .find_seq_from(self.pos, b"?>")?
                    .ok_or(XmlError::Malformed("unclosed decl"))?;
                let decl = self.buf[self.pos + 5..end].to_vec();
                self.check_decl(&decl)?;
                self.pos = end + 2;
            }
        }
        Ok(())
    }

    fn check_decl(&self, decl: &[u8]) -> XResult<()> {
        // parse pseudo attributes
        let mut i = 0;
        let n = decl.len();
        let mut seen_version = false;
        let mut idx = 0;
        loop {
            let ws_start = i;
            while i < n && is_ws(decl[i]) {
                i += 1;
            }
            if i >= n {
                break;
            }
            if i == ws_start {
                return Err(XmlError::Malformed("bad decl"));
            }
            let ns = i;
            while i < n && decl[i].is_ascii_alphabetic() {
                i += 1;
            }
            let name = &decl[ns..i];
            while i < n && is_ws(decl[i]) {
                i += 1;
            }
            if i >= n || decl[i] != b'=' {
                return Err(XmlError::Malformed("bad decl"));
            }
            i += 1;
            while i < n && is_ws(decl[i]) {
                i += 1;
            }
            if i >= n || (decl[i] != b'"' && decl[i] != b'\'') {
                return Err(XmlError::Malformed("bad decl"));
            }
            let q = decl[i];
            i += 1;
            let vs = i;
            while i < n && decl[i] != q {
                i += 1;
            }
            if i >= n {
                return Err(XmlError::Malformed("bad decl"));
            }
            let val = &decl[vs..i];
            i += 1;
            match (name, idx) {
                (b"version", 0) => {
                    if val != b"1.0" {
                        return Err(XmlError::Unsupported("xml version"));
                    }
                    seen_version = true;
                }
                (b"encoding", 1) => {
                    let v = val.to_ascii_lowercase();
                    if v != b"utf-8" {
                        return Err(XmlError::Unsupported("encoding"));
                    }
                }
                (b"standalone", 1) | (b"standalone", 2) => {
                    if val != b"yes" && val != b"no" {
                        return Err(XmlError::Malformed("bad standalone"));
                    }
                }
                _ => return Err(XmlError::Malformed("bad decl")),
            }
            idx += 1;
            if name == b"encoding" {
                idx = 2;
            }
        }
        if !seen_version {
            return Err(XmlError::Malformed("bad decl"));
        }
        Ok(())
    }

    /// `<!` constructs. Returns true if a Text event (CDATA) was produced.
    fn markup_decl(&mut self) -> XResult<bool> {
        self.ensure(9)?;
        let b = &self.buf[self.pos..];
        if b.starts_with(b"<!--") {
            let end = self
                .find_seq_from(self.pos + 4, b"-->")?
                .ok_or(XmlError::Malformed("unclosed comment"))?;
            let body = &self.buf[self.pos + 4..end];
            if memmem::find(body, b"--").is_some() || body.last() == Some(&b'-') {
                return Err(XmlError::Malformed("-- in comment"));
            }
            validate_chars(body)?;
            self.pos = end + 3;
            Ok(false)
        } else if b.starts_with(b"<![CDATA[") {
            if self.stack.is_empty() {
                return Err(XmlError::Malformed("cdata outside root"));
            }
            let start = self.pos + 9;
            let end = self
                .find_seq_from(start, b"]]>")?
                .ok_or(XmlError::Malformed("unclosed cdata"))?;
            let body = &self.buf[start..end];
            validate_chars(body)?;
            self.text = Range { start, end };
            self.text_cdata = true;
            self.text_needs_decode = memchr(b'\r', body).is_some();
            self.pos = end + 3;
            self.ev = Ev::Text;
            Ok(true)
        } else if b.starts_with(b"<!DOCTYPE") {
            Err(XmlError::Unsupported("doctype"))
        } else {
            Err(XmlError::Malformed("bad markup declaration"))
        }
    }

    fn pi(&mut self) -> XResult<()> {
        let end = self
            .find_seq_from(self.pos + 2, b"?>")?
            .ok_or(XmlError::Malformed("unclosed pi"))?;
        let body = &self.buf[self.pos + 2..end];
        let mut i = 0;
        while i < body.len() && is_name_char(body[i]) {
            i += 1;
        }
        let target = &body[..i];
        if target.is_empty() || !is_name_start(target[0]) {
            return Err(XmlError::Malformed("bad pi"));
        }
        if target.eq_ignore_ascii_case(b"xml") {
            return Err(XmlError::Malformed("misplaced xml declaration"));
        }
        if memchr(b':', target).is_some() {
            return Err(XmlError::Malformed("colon in pi target"));
        }
        if i < body.len() && !is_ws(body[i]) {
            return Err(XmlError::Malformed("bad pi"));
        }
        validate_chars(body)?;
        self.pos = end + 2;
        Ok(())
    }

    fn intern_uri(&mut self, uri: &str) -> u32 {
        if let Some(&id) = self.uri_ids.get(uri.as_bytes()) {
            return id;
        }
        let id = self.uris.len() as u32;
        self.uris.push(uri.to_string());
        self.uri_ids.insert(uri.as_bytes().to_vec(), id);
        id
    }

    fn lookup_prefix(&self, prefix: &[u8]) -> Option<u32> {
        if prefix == b"xml" {
            return Some(NS_XML);
        }
        for (p, id) in self.bindings.iter().rev() {
            if p.as_slice() == prefix {
                return Some(*id);
            }
        }
        if prefix.is_empty() {
            return Some(NS_NONE);
        }
        None
    }

    /// Parse a qualified name at `i` within the buffer; returns end index.
    fn scan_name(&self, i: usize, end: usize) -> XResult<usize> {
        let b = &self.buf;
        if i >= end {
            return Err(XmlError::Malformed("expected name"));
        }
        if b[i] >= 0x80 {
            return Err(XmlError::Unsupported("non-ascii name"));
        }
        if !is_name_start(b[i]) {
            return Err(XmlError::Malformed("bad name"));
        }
        let mut j = i + 1;
        while j < end {
            let c = b[j];
            if c >= 0x80 {
                return Err(XmlError::Unsupported("non-ascii name"));
            }
            if !is_name_char(c) {
                break;
            }
            j += 1;
        }
        Ok(j)
    }

    /// split qname into (prefix range, local range); validates colons
    fn split_qname(&self, s: usize, e: usize) -> XResult<(Range, Range)> {
        let q = &self.buf[s..e];
        match memchr(b':', q) {
            None => Ok((Range { start: s, end: s }, Range { start: s, end: e })),
            Some(c) => {
                if c == 0 || c + 1 == q.len() || memchr(b':', &q[c + 1..]).is_some() {
                    return Err(XmlError::Unsupported("bad qname"));
                }
                if !is_name_start(q[c + 1]) || q[c + 1] == b':' {
                    return Err(XmlError::Malformed("bad qname"));
                }
                Ok((
                    Range {
                        start: s,
                        end: s + c,
                    },
                    Range {
                        start: s + c + 1,
                        end: e,
                    },
                ))
            }
        }
    }

    fn find_tag_end(&mut self) -> XResult<usize> {
        // find '>' outside of quotes starting from pos+1
        let mut i = self.pos + 1;
        let mut quote: u8 = 0;
        loop {
            let found = if quote == 0 {
                memchr3(b'>', b'"', b'\'', &self.buf[i..]).map(|o| i + o)
            } else {
                memchr(quote, &self.buf[i..]).map(|o| i + o)
            };
            match found {
                Some(j) => {
                    let c = self.buf[j];
                    if quote != 0 {
                        quote = 0;
                        i = j + 1;
                    } else if c == b'>' {
                        return Ok(j);
                    } else {
                        quote = c;
                        i = j + 1;
                    }
                }
                None => {
                    let len = self.buf.len();
                    let before = self.pos;
                    if !self.more()? {
                        return Err(XmlError::Malformed("unclosed tag"));
                    }
                    let shift = before - self.pos;
                    i = len - shift;
                    let _ = i;
                    i = len - shift;
                }
            }
        }
    }

    fn start_tag(&mut self) -> XResult<Ev> {
        if self.root_closed {
            return Err(XmlError::Malformed("junk after document element"));
        }
        let gt = self.find_tag_end()?;
        let start = self.pos;
        let mut content_end = gt;
        let empty = self.buf[gt - 1] == b'/' && gt - 1 > start;
        if empty {
            content_end = gt - 1;
        }
        let name_end = self.scan_name(start + 1, content_end)?;
        let (prefix, local) = self.split_qname(start + 1, name_end)?;
        self.attrs.clear();
        let mut i = name_end;
        let mut new_bindings: Vec<(Vec<u8>, u32)> = Vec::new();
        loop {
            let ws_start = i;
            while i < content_end && is_ws(self.buf[i]) {
                i += 1;
            }
            if i >= content_end {
                break;
            }
            if i == ws_start {
                return Err(XmlError::Malformed("missing whitespace between attributes"));
            }
            let an_end = self.scan_name(i, content_end)?;
            let an_s = i;
            i = an_end;
            while i < content_end && is_ws(self.buf[i]) {
                i += 1;
            }
            if i >= content_end || self.buf[i] != b'=' {
                return Err(XmlError::Malformed("expected ="));
            }
            i += 1;
            while i < content_end && is_ws(self.buf[i]) {
                i += 1;
            }
            if i >= content_end {
                return Err(XmlError::Malformed("expected quote"));
            }
            let q = self.buf[i];
            if q != b'"' && q != b'\'' {
                return Err(XmlError::Malformed("expected quote"));
            }
            let vs = i + 1;
            let ve = match memchr(q, &self.buf[vs..content_end]) {
                Some(o) => vs + o,
                None => return Err(XmlError::Malformed("unterminated attribute")),
            };
            i = ve + 1;
            let val = &self.buf[vs..ve];
            if memchr(b'<', val).is_some() {
                return Err(XmlError::Malformed("< in attribute"));
            }
            validate_chars(val)?;
            let has_amp = memchr(b'&', val).is_some();
            if has_amp {
                validate_refs(val)?;
            }
            let needs_decode =
                has_amp || val.iter().any(|&b| b == b'\r' || b == b'\n' || b == b'\t');
            // duplicate qname check
            let qn = &self.buf[an_s..an_end];
            for a in &self.attrs {
                if &self.buf[a.qname.start..a.qname.end] == qn {
                    return Err(XmlError::Malformed("duplicate attribute"));
                }
            }
            let (ap, al) = self.split_qname(an_s, an_end)?;
            let pfx = &self.buf[ap.start..ap.end];
            let is_xmlns = (ap.start == ap.end && qn == b"xmlns") || pfx == b"xmlns";
            if is_xmlns {
                let uri = if needs_decode {
                    decode_attr(val).into_owned()
                } else {
                    unsafe { std::str::from_utf8_unchecked(val) }.to_string()
                };
                let p: Vec<u8> = if pfx == b"xmlns" {
                    self.buf[al.start..al.end].to_vec()
                } else {
                    Vec::new()
                };
                if p == b"xmlns" {
                    return Err(XmlError::Malformed("xmlns prefix"));
                }
                if p == b"xml" {
                    if uri != XML_NS_URI {
                        return Err(XmlError::Malformed("xml prefix"));
                    }
                } else if uri == XML_NS_URI || uri == "http://www.w3.org/2000/xmlns/" {
                    return Err(XmlError::Malformed("reserved namespace"));
                }
                if !p.is_empty() && uri.is_empty() {
                    return Err(XmlError::Malformed("empty prefixed namespace"));
                }
                let id = self.intern_uri(&uri);
                new_bindings.push((p, id));
                // xmlns attributes are kept to detect duplicates but marked
                self.attrs.push(Attr {
                    ns: u32::MAX,
                    qname: Range {
                        start: an_s,
                        end: an_end,
                    },
                    local: al,
                    value: Range { start: vs, end: ve },
                    needs_decode,
                });
            } else {
                // namespace resolved after all bindings are known; store prefix in ns temporarily
                self.attrs.push(Attr {
                    ns: u32::MAX - 1,
                    qname: Range {
                        start: an_s,
                        end: an_end,
                    },
                    local: al,
                    value: Range { start: vs, end: ve },
                    needs_decode,
                });
            }
        }
        // push bindings
        let nb = new_bindings.len();
        self.bindings.extend(new_bindings);
        // resolve element name
        let ens = match self.lookup_prefix(&self.buf[prefix.start..prefix.end]) {
            Some(id) => id,
            None => return Err(XmlError::Malformed("unbound prefix")),
        };
        // resolve attributes
        self.attrs.retain(|a| a.ns != u32::MAX);
        let mut names = std::mem::take(&mut self.scratch_names);
        names.clear();
        for k in 0..self.attrs.len() {
            let a = &self.attrs[k];
            let q = a.qname;
            let l = a.local;
            let ns = if l.start == q.start {
                NS_NONE
            } else {
                let p = &self.buf[q.start..l.start - 1];
                if p.is_empty() {
                    NS_NONE
                } else {
                    match self.lookup_prefix(p) {
                        Some(id) if id != NS_NONE => id,
                        _ => return Err(XmlError::Malformed("unbound prefix")),
                    }
                }
            };
            self.attrs[k].ns = ns;
            // expanded-name duplicate check
            for (ons, ol) in names.iter() {
                if *ons == ns && self.buf[ol.start..ol.end] == self.buf[l.start..l.end] {
                    return Err(XmlError::Malformed("duplicate attribute"));
                }
            }
            names.push((ns, l));
        }
        self.scratch_names = names;
        self.name_ns = ens;
        self.name_local = local;
        self.tag = Range { start, end: gt + 1 };
        let name_start = self.names.len() as u32;
        self.names.extend_from_slice(&self.buf[start + 1..name_end]);
        self.stack.push(OpenElem {
            name_start,
            name_len: (name_end - start - 1) as u32,
            ns_bindings: nb as u32,
        });
        self.seen_root = true;
        self.pos = gt + 1;
        if empty {
            self.pending_end = true;
            self.ev = Ev::Empty;
            Ok(Ev::Empty)
        } else {
            self.ev = Ev::Start;
            Ok(Ev::Start)
        }
    }

    fn pop_element(&mut self) {
        let e = self.stack.pop().unwrap();
        self.names.truncate(e.name_start as usize);
        let n = self.bindings.len() - e.ns_bindings as usize;
        self.bindings.truncate(n);
        if self.stack.is_empty() {
            self.root_closed = true;
        }
    }

    fn end_tag(&mut self) -> XResult<Ev> {
        let gt_rel = loop {
            if let Some(o) = memchr(b'>', &self.buf[self.pos..]) {
                break o;
            }
            if !self.more()? {
                return Err(XmlError::Malformed("unclosed end tag"));
            }
        };
        let start = self.pos;
        let gt = start + gt_rel;
        let name_end = self.scan_name(start + 2, gt)?;
        let mut i = name_end;
        while i < gt && is_ws(self.buf[i]) {
            i += 1;
        }
        if i != gt {
            return Err(XmlError::Malformed("bad end tag"));
        }
        let (p, local) = self.split_qname(start + 2, name_end)?;
        match self.stack.last() {
            Some(top)
                if self.names
                    [top.name_start as usize..(top.name_start + top.name_len) as usize]
                    == self.buf[start + 2..name_end] => {}
            _ => return Err(XmlError::Malformed("mismatched tag")),
        }
        // resolve name before popping bindings
        self.name_ns = self
            .lookup_prefix(&self.buf[p.start..p.end])
            .unwrap_or(NS_NONE);
        self.name_local = local;
        self.tag = Range { start, end: gt + 1 };
        self.pos = gt + 1;
        self.pop_element();
        self.ev = Ev::End;
        Ok(Ev::End)
    }
}

/// Result of `Reader::fast_cell`: byte ranges (in `Reader::buf`) of the
/// parts of a simple `<c>` element. Valid until the next call to `next`.
#[derive(Default)]
pub struct FastCell {
    pub r: Option<Range>,
    pub s: Option<Range>,
    pub t: Option<Range>,
    /// the `<f>` child: present, its t / ref / si attributes and its text
    pub f: bool,
    pub f_t: Option<Range>,
    pub f_ref: Option<Range>,
    pub f_si: Option<Range>,
    pub f_text: Option<Range>,
    /// text of the `<v>` child (None: no child or empty)
    pub v: Option<Range>,
}

#[inline]
fn fast_text_ok(b: &[u8]) -> bool {
    // character data accepted by the fast path: no markup, no references, no
    // CR (needs normalisation), no forbidden characters, valid UTF-8
    let mut non_ascii = false;
    for &c in b {
        if c < 0x20 {
            if c != b'\t' && c != b'\n' {
                return false;
            }
        } else if c == b'&' || c == b'<' || c == b'>' {
            // '>' only matters as part of "]]>"; reject conservatively
            return false;
        } else if c >= 0x80 {
            non_ascii = true;
        }
    }
    !non_ascii
        || (std::str::from_utf8(b).is_ok()
            && memmem::find(b, b"\xEF\xBF\xBE").is_none()
            && memmem::find(b, b"\xEF\xBF\xBF").is_none())
}

impl<S: Source> Reader<S> {
    /// True if unprefixed element names currently resolve to the
    /// spreadsheetml main namespace.
    pub fn default_ns_is_main(&self) -> bool {
        self.lookup_prefix(b"") == Some(NS_MAIN)
    }

    /// Parse attributes of a fast-path tag at `i` up to `>` or `/>`.
    /// Accepts only unprefixed ASCII names, double quoted values without
    /// references / whitespace normalisation. Calls `found(name, value)`.
    /// Returns (index after the tag, empty element).
    #[inline]
    fn fast_attrs(
        &self,
        mut i: usize,
        found: &mut dyn FnMut(&[u8], Range) -> bool,
    ) -> Option<(usize, bool)> {
        let b = &self.buf;
        let n = b.len();
        // names seen, to reject duplicates (cells have few attributes)
        let mut seen: [(usize, usize); 8] = [(0, 0); 8];
        let mut nseen = 0;
        loop {
            if i >= n {
                return None;
            }
            match b[i] {
                b'>' => return Some((i + 1, false)),
                b'/' => {
                    return if i + 1 < n && b[i + 1] == b'>' {
                        Some((i + 2, true))
                    } else {
                        None
                    };
                }
                b' ' => {}
                _ => return None,
            }
            // one or more spaces
            while i < n && b[i] == b' ' {
                i += 1;
            }
            if i >= n {
                return None;
            }
            if b[i] == b'>' || b[i] == b'/' {
                continue;
            }
            let ns = i;
            while i < n && b[i].is_ascii_alphabetic() {
                i += 1;
            }
            if i == ns || i + 1 >= n || b[i] != b'=' || b[i + 1] != b'"' {
                return None;
            }
            let name = (ns, i);
            if &b[ns..i] == b"xmlns" {
                return None;
            }
            for k in 0..nseen {
                if b[seen[k].0..seen[k].1] == b[ns..i] {
                    return None;
                }
            }
            if nseen == seen.len() {
                return None;
            }
            seen[nseen] = name;
            nseen += 1;
            let vs = i + 2;
            let ve = vs + memchr(b'"', &b[vs..])?;
            let val = &b[vs..ve];
            if val
                .iter()
                .any(|&c| c < 0x20 || c == b'<' || c == b'&' || c >= 0x80)
            {
                return None;
            }
            if !found(&b[ns..name.1], Range { start: vs, end: ve }) {
                return None;
            }
            i = ve + 1;
        }
    }

    /// Fast path for a simple `<c>` element starting at the current position
    /// (`<c r=".." s=".." t="..">[<f ..>..</f>][<v>..</v>]</c>`). On success
    /// the reader is positioned after the element and `out` describes it; on
    /// failure nothing is consumed (use the generic path).
    pub fn fast_cell(&mut self, out: &mut FastCell) -> bool {
        let b = &self.buf;
        let n = b.len();
        let mut i = self.pos;
        if i + 3 > n
            || &b[i..i + 2] != b"<c"
            || !(b[i + 2] == b' ' || b[i + 2] == b'>' || b[i + 2] == b'/')
        {
            return false;
        }
        *out = FastCell::default();
        let mut ok = true;
        let res = {
            let mut found = |name: &[u8], r: Range| -> bool {
                match name {
                    b"r" => out.r = Some(r),
                    b"s" => out.s = Some(r),
                    b"t" => out.t = Some(r),
                    _ => {}
                }
                true
            };
            self.fast_attrs(i + 2, &mut found)
        };
        let Some((after, empty)) = res else {
            return false;
        };
        i = after;
        if !empty {
            let b = &self.buf;
            loop {
                if i + 3 > n || b[i] != b'<' {
                    return false;
                }
                if b[i + 1] == b'/' {
                    // </c>
                    if i + 4 <= n && &b[i..i + 4] == b"</c>" {
                        i += 4;
                        break;
                    }
                    return false;
                }
                let child = b[i + 1];
                if !(child == b'v' || child == b'f')
                    || !(b[i + 2] == b'>' || b[i + 2] == b' ' || b[i + 2] == b'/')
                {
                    return false;
                }
                if child == b'v' {
                    if out.v.is_some() || !ok {
                        return false;
                    }
                    if b[i + 2] == b'/' {
                        if i + 4 <= n && b[i + 3] == b'>' {
                            i += 4;
                            ok = true;
                            out.v = Some(Range { start: i, end: i });
                            continue;
                        }
                        return false;
                    }
                    if b[i + 2] != b'>' {
                        return false; // attributes on <v>: generic path
                    }
                    let ts = i + 3;
                    let Some(off) = memchr(b'<', &b[ts..]) else {
                        return false;
                    };
                    let te = ts + off;
                    if te + 4 > n || &b[te..te + 4] != b"</v>" || !fast_text_ok(&b[ts..te]) {
                        return false;
                    }
                    out.v = Some(Range { start: ts, end: te });
                    i = te + 4;
                } else {
                    if out.f {
                        return false;
                    }
                    out.f = true;
                    let res = {
                        let mut found = |name: &[u8], r: Range| -> bool {
                            match name {
                                b"t" => out.f_t = Some(r),
                                b"ref" => out.f_ref = Some(r),
                                b"si" => out.f_si = Some(r),
                                _ => return false, // other attributes: generic path
                            }
                            true
                        };
                        self.fast_attrs(i + 2, &mut found)
                    };
                    let Some((after, fempty)) = res else {
                        return false;
                    };
                    let b = &self.buf;
                    i = after;
                    if !fempty {
                        let ts = i;
                        let Some(off) = memchr(b'<', &b[ts..]) else {
                            return false;
                        };
                        let te = ts + off;
                        if te + 4 > n || &b[te..te + 4] != b"</f>" || !fast_text_ok(&b[ts..te]) {
                            return false;
                        }
                        out.f_text = Some(Range { start: ts, end: te });
                        i = te + 4;
                    }
                }
            }
        }
        // a dataTable formula needs the full attribute dict: generic path
        if let Some(t) = out.f_t {
            if &self.buf[t.start..t.end] == b"dataTable" {
                return false;
            }
        }
        if out.v.map(|r| r.start == r.end).unwrap_or(false) {
            out.v = None;
        }
        self.tag = Range {
            start: self.pos,
            end: i,
        };
        self.pos = i;
        true
    }
}

impl<S: Source> Reader<S> {
    /// Resolve a namespace prefix in the current scope (pub for raw scans).
    pub fn uri(&self, id: u32) -> &str {
        &self.uris[id as usize]
    }

    pub fn resolve_prefix(&self, prefix: &[u8]) -> Option<u32> {
        self.lookup_prefix(prefix)
    }

    /// Raw scan of the content of the element whose Start event was just
    /// returned (eg. `sheetData`) without tokenizing it: byte searches for
    /// the closing tag `</{qname}>`. `<row ...>` / `<c ...>` start tags are
    /// reported to `f` (is_row, range of the tag content without `<` and
    /// `>`/`/>`) when requested. Returns Ok(false), without guarantees about
    /// the position, if the content contains constructs a raw scan cannot
    /// handle (comments, CDATA sections, processing instructions); on
    /// Ok(true) the reader is positioned on the closing tag.
    pub fn raw_skip(
        &mut self,
        qname: &[u8],
        want_rows: bool,
        want_cells: bool,
        f: &mut dyn FnMut(&Self, bool, Range) -> XResult<()>,
    ) -> XResult<bool> {
        let mut scan = self.pos;
        'outer: loop {
            let off = match memchr(b'<', &self.buf[scan..]) {
                Some(o) => o,
                None => {
                    self.pos = self.buf.len();
                    if !self.more()? {
                        return Err(XmlError::Malformed("eof in element"));
                    }
                    scan = self.pos;
                    continue;
                }
            };
            let i = scan + off;
            // enough bytes to classify the tag (refill keeps data from pos)
            if self.buf.len() - i < qname.len() + 3 {
                self.pos = i;
                if !self.more()? {
                    return Err(XmlError::Malformed("eof in element"));
                }
                scan = self.pos;
                continue;
            }
            let b = &self.buf;
            match b[i + 1] {
                b'/' => {
                    let rest = &b[i + 2..];
                    if rest.starts_with(qname) {
                        let k = rest[qname.len()];
                        if k == b'>' || is_ws(k) {
                            self.pos = i;
                            return Ok(true);
                        }
                    }
                    scan = i + 2;
                }
                b'!' | b'?' => return Ok(false),
                c => {
                    let is_row = want_rows
                        && b[i + 1..].starts_with(b"row")
                        && (is_ws(b[i + 4]) || b[i + 4] == b'>' || b[i + 4] == b'/');
                    let is_cell = !is_row
                        && want_cells
                        && c == b'c'
                        && (is_ws(b[i + 2]) || b[i + 2] == b'>' || b[i + 2] == b'/');
                    if !(is_row || is_cell) {
                        scan = i + 1;
                        continue;
                    }
                    // the whole start tag must be in the buffer
                    let gt = match self.tag_end_in_buffer(i) {
                        Some(g) => g,
                        None => {
                            self.pos = i;
                            if !self.more()? {
                                return Err(XmlError::Malformed("eof in tag"));
                            }
                            scan = self.pos; // rescan from the tag start
                            continue 'outer;
                        }
                    };
                    let end = if self.buf[gt - 1] == b'/' { gt - 1 } else { gt };
                    f(self, is_row, Range { start: i + 1, end })?;
                    scan = gt + 1;
                }
            }
        }
    }

    /// Index of the '>' ending the tag starting at `start` (quote aware), if
    /// it is in the buffer.
    fn tag_end_in_buffer(&self, start: usize) -> Option<usize> {
        let b = &self.buf;
        let mut i = start;
        let mut quote = 0u8;
        loop {
            let found = if quote == 0 {
                memchr3(b'>', b'"', b'\'', &b[i..])
            } else {
                memchr(quote, &b[i..])
            }?;
            let j = i + found;
            let c = b[j];
            if quote != 0 {
                quote = 0;
            } else if c == b'>' {
                return Some(j);
            } else {
                quote = c;
            }
            i = j + 1;
        }
    }
}
