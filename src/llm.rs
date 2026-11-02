//! openrsxl.extended.llm: lossless, LLM readable encoding of worksheets.
//!
//! The grammar is specified in docs/LLM_ENCODING.md (and in the legend that
//! `openrsxl.extended.llm` puts in front of every document); the Python
//! decoder (`openrsxl/extended/llm/_decode.py`) is its reference reader.
//!
//! Cells are streamed in row-major order through a `RowAssembler` (rows are
//! completed, sorted by column, de-duplicated like full mode's `ws._cells`)
//! into the consumers of the current pass:
//!
//! * pass 1 - analysis (counts, bounding box, token estimates of every
//!   values encoding), the inverted index of values (bounded: abandoned
//!   above `index_limit` bytes), formula fill groups and the style map;
//! * pass 2 - the values grid (`grid` and `hybrid` encodings), written
//!   row by row to a sink so that memory does not grow with the sheet.
//!
//! Losslessness never depends on a heuristic: a formula range is emitted
//! only after rendering its template at every cell reproduced the original
//! text byte for byte (`ftemplate`), and a string is quoted whenever its
//! unquoted form would read as anything else.

use std::collections::BTreeMap;
use std::sync::Arc;

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyDate, PyDateAccess, PyDateTime, PyDelta, PyDeltaAccess, PyDict, PyFloat, PyInt,
    PyList, PyString, PyTime, PyTimeAccess, PyTzInfoAccess,
};

use crate::dates::ExcelDate;
use crate::ftemplate::{col_letters, key_hash, Template};
use crate::pysource::PyFileSource;
use crate::sheet::{
    fallback_err, DType, Helpers, Item, Settings, Shared, SheetParser, StyleId, Val, DT_E,
};
use crate::store::FxMap;

// ---------------------------------------------------------------------------
// scalar tokens

/// Kind of a value, for the analysis.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Number,
    Text,
    Bool,
    Error,
    Date,
}

const KIND_NAMES: [&str; 5] = ["number", "text", "bool", "error", "date"];

fn kind_index(k: Kind) -> usize {
    match k {
        Kind::Number => 0,
        Kind::Text => 1,
        Kind::Bool => 2,
        Kind::Error => 3,
        Kind::Date => 4,
    }
}

/// `repr(float)` of a finite float (shortest round-trip digits, Python's
/// choice between fixed and exponent notation).
pub fn push_float(out: &mut String, v: f64) {
    if v == 0.0 {
        out.push_str(if v.is_sign_negative() { "-0.0" } else { "0.0" });
        return;
    }
    let s = format!("{:e}", v);
    let (mant, exp) = s.split_once('e').unwrap();
    let exp: i32 = exp.parse().unwrap();
    let (neg, mant) = match mant.strip_prefix('-') {
        Some(m) => (true, m),
        None => (false, mant),
    };
    let digits: Vec<u8> = mant.bytes().filter(|&c| c != b'.').collect();
    let nd = digits.len() as i32;
    if neg {
        out.push('-');
    }
    let decpt = exp + 1;
    let ds = |a: i32, b: i32| -> &str {
        // digits[a..b] as str
        unsafe { std::str::from_utf8_unchecked(&digits[a as usize..b as usize]) }
    };
    if -4 < decpt && decpt <= 16 {
        if decpt <= 0 {
            out.push_str("0.");
            for _ in 0..(-decpt) {
                out.push('0');
            }
            out.push_str(ds(0, nd));
        } else if decpt >= nd {
            out.push_str(ds(0, nd));
            for _ in 0..(decpt - nd) {
                out.push('0');
            }
            out.push_str(".0");
        } else {
            out.push_str(ds(0, decpt));
            out.push('.');
            out.push_str(ds(decpt, nd));
        }
    } else {
        out.push_str(ds(0, 1));
        if nd > 1 {
            out.push('.');
            out.push_str(ds(1, nd));
        }
        out.push('e');
        let e = decpt - 1;
        out.push(if e < 0 { '-' } else { '+' });
        let ea = e.unsigned_abs();
        if ea < 10 {
            out.push('0');
        }
        out.push_str(&ea.to_string());
    }
}

fn is_digit(c: u8) -> bool {
    c.is_ascii_digit()
}

/// `[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?` (unquoted numbers)
pub fn looks_number(s: &[u8]) -> bool {
    let n = s.len();
    let mut i = 0;
    if i < n && (s[i] == b'+' || s[i] == b'-') {
        i += 1;
    }
    let ds = i;
    while i < n && is_digit(s[i]) {
        i += 1;
    }
    let int_digits = i - ds;
    let mut frac_digits = 0;
    if i < n && s[i] == b'.' {
        i += 1;
        let fs = i;
        while i < n && is_digit(s[i]) {
            i += 1;
        }
        frac_digits = i - fs;
    }
    if int_digits == 0 && frac_digits == 0 {
        return false;
    }
    if i < n && (s[i] == b'e' || s[i] == b'E') {
        i += 1;
        if i < n && (s[i] == b'+' || s[i] == b'-') {
            i += 1;
        }
        let es = i;
        while i < n && is_digit(s[i]) {
            i += 1;
        }
        if i == es {
            return false;
        }
    }
    i == n
}

fn digits_at(s: &[u8], i: usize, k: usize) -> bool {
    i + k <= s.len() && s[i..i + k].iter().all(|&c| is_digit(c))
}

/// `\d{2}:\d{2}:\d{2}(\.\d{1,6})?` starting at `i`, to the end
fn looks_time_at(s: &[u8], i: usize) -> bool {
    if !(digits_at(s, i, 2)
        && s.get(i + 2) == Some(&b':')
        && digits_at(s, i + 3, 2)
        && s.get(i + 5) == Some(&b':')
        && digits_at(s, i + 6, 2))
    {
        return false;
    }
    let j = i + 8;
    if j == s.len() {
        return true;
    }
    if s[j] != b'.' {
        return false;
    }
    let k = s.len() - j - 1;
    (1..=6).contains(&k) && digits_at(s, j + 1, k)
}

/// `\d{4}-\d{2}-\d{2}( <time>)?` or `<time>` (unquoted dates and times)
pub fn looks_date(s: &[u8]) -> bool {
    if looks_time_at(s, 0) {
        return true;
    }
    if !(digits_at(s, 0, 4)
        && s.get(4) == Some(&b'-')
        && digits_at(s, 5, 2)
        && s.get(7) == Some(&b'-')
        && digits_at(s, 8, 2))
    {
        return false;
    }
    if s.len() == 10 {
        return true;
    }
    s.get(10) == Some(&b' ') && looks_time_at(s, 11)
}

/// `\*\d+$` (run suffix)
fn has_run_suffix(s: &[u8]) -> bool {
    let n = s.len();
    let mut i = n;
    while i > 0 && is_digit(s[i - 1]) {
        i -= 1;
    }
    i < n && i > 0 && s[i - 1] == b'*'
}

/// `^[A-Z]{1,3}=` (column jump prefix)
fn has_jump_prefix(s: &[u8]) -> bool {
    let mut i = 0;
    while i < s.len() && i < 4 && s[i].is_ascii_uppercase() {
        i += 1;
    }
    (1..=3).contains(&i) && s.get(i) == Some(&b'=')
}

/// Can `s` be written as an unquoted string token?
pub fn str_is_plain(s: &str) -> bool {
    let b = s.as_bytes();
    if b.is_empty() {
        return false;
    }
    if matches!(b[0], b'"' | b'#' | b'@' | b'=') {
        return false;
    }
    let first = s.chars().next().unwrap();
    let last = s.chars().next_back().unwrap();
    if first.is_whitespace() || last.is_whitespace() {
        return false;
    }
    for ch in s.chars() {
        if ch == '|' || ch.is_control() || ch == '\u{2028}' || ch == '\u{2029}' || ch == '\u{85}' {
            return false;
        }
    }
    if s == "TRUE" || s == "FALSE" || s == "^" {
        return false;
    }
    !(looks_number(b) || looks_date(b) || has_run_suffix(b) || has_jump_prefix(b))
}

/// JSON string literal (what Python's json.loads reads back)
pub fn push_json(out: &mut String, s: &str) {
    out.push('"');
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20
                || c == '\u{7f}'
                || c == '\u{2028}'
                || c == '\u{2029}'
                || c == '\u{85}' =>
            {
                out.push_str(&format!("\\u{:04x}", c as u32));
            }
            c => out.push(c),
        }
    }
    out.push('"');
}

pub fn push_str_token(out: &mut String, s: &str) {
    if str_is_plain(s) {
        out.push_str(s);
    } else {
        push_json(out, s);
    }
}

/// An error value: plain (`#N/A`) when it reads back as that error
pub fn push_error_token(out: &mut String, s: &str) {
    let plain = s.starts_with('#')
        && s.len() > 1
        && !has_run_suffix(s.as_bytes())
        && !s
            .chars()
            .any(|c| c == '|' || c.is_whitespace() || c.is_control());
    if plain {
        out.push_str(s);
    } else {
        out.push_str("@err:");
        push_json(out, s);
    }
}

fn push_2(out: &mut String, v: u32) {
    if v < 10 {
        out.push('0');
    }
    out.push_str(&v.to_string());
}

fn push_time(out: &mut String, h: u32, mi: u32, s: u32, us: u32) {
    push_2(out, h);
    out.push(':');
    push_2(out, mi);
    out.push(':');
    push_2(out, s);
    if us != 0 {
        out.push_str(&format!(".{:06}", us));
    }
}

fn push_ymd(out: &mut String, y: i32, m: u32, d: u32) {
    out.push_str(&format!("{:04}-{:02}-{:02}", y, m, d));
}

#[allow(clippy::too_many_arguments)]
/// A datetime token: always with its time, also at midnight
/// (`2025-04-06 00:00:00`: openpyxl reads every date cell as a datetime, and
/// a reader must see the value as openpyxl gives it). A bare `YYYY-MM-DD`
/// (written before 0.4.3 for midnight) still decodes as a midnight datetime.
pub fn push_datetime(out: &mut String, y: i32, m: u32, d: u32, h: u32, mi: u32, s: u32, us: u32) {
    push_ymd(out, y, m, d);
    out.push(' ');
    push_time(out, h, mi, s, us);
}

fn push_excel_date(out: &mut String, d: ExcelDate) -> Kind {
    match d {
        ExcelDate::DateTime {
            y,
            m,
            d,
            h,
            mi,
            s,
            us,
        } => {
            push_datetime(
                out, y, m as u32, d as u32, h as u32, mi as u32, s as u32, us,
            );
            Kind::Date
        }
        ExcelDate::Time { h, mi, s, us } => {
            push_time(out, h as u32, mi as u32, s as u32, us);
            Kind::Date
        }
        ExcelDate::Error => unreachable!("validated at load time"),
    }
}

/// Token of a Python value (None: no value).
pub fn py_value_token(
    v: &Bound<'_, PyAny>,
    error: bool,
    out: &mut String,
) -> PyResult<Option<Kind>> {
    let py = v.py();
    if v.is_none() {
        return Ok(None);
    }
    if v.is_instance_of::<PyBool>() {
        out.push_str(if v.is_truthy()? { "TRUE" } else { "FALSE" });
        return Ok(Some(Kind::Bool));
    }
    if v.is_instance_of::<PyInt>() {
        match v.extract::<i64>() {
            Ok(i) => out.push_str(&i.to_string()),
            Err(_) => out.push_str(&v.str()?.to_cow()?),
        }
        return Ok(Some(Kind::Number));
    }
    if v.is_instance_of::<PyFloat>() {
        let f: f64 = v.extract()?;
        if f.is_finite() {
            push_float(out, f);
        } else {
            out.push_str("@float:");
            out.push_str(if f.is_nan() {
                "nan"
            } else if f > 0.0 {
                "inf"
            } else {
                "-inf"
            });
        }
        return Ok(Some(Kind::Number));
    }
    if v.is_instance_of::<PyString>() {
        let s = v.cast::<PyString>()?;
        match s.to_str() {
            Ok(t) => {
                if error {
                    push_error_token(out, t);
                    return Ok(Some(Kind::Error));
                }
                push_str_token(out, t);
            }
            Err(_) => {
                // lone surrogates: Python's JSON escapes them
                let j: String = py.import("json")?.call_method1("dumps", (v,))?.extract()?;
                if error {
                    out.push_str("@err:");
                    out.push_str(&j);
                    return Ok(Some(Kind::Error));
                }
                out.push_str(&j);
            }
        }
        return Ok(Some(Kind::Text));
    }
    if v.is_instance_of::<PyDateTime>() {
        let d = v.cast::<PyDateTime>()?;
        if d.get_tzinfo().is_some() {
            out.push_str("@dt:");
            out.push_str(&d.call_method0("isoformat")?.extract::<String>()?);
        } else {
            push_datetime(
                out,
                d.get_year(),
                d.get_month() as u32,
                d.get_day() as u32,
                d.get_hour() as u32,
                d.get_minute() as u32,
                d.get_second() as u32,
                d.get_microsecond(),
            );
        }
        return Ok(Some(Kind::Date));
    }
    if v.is_instance_of::<PyDate>() {
        let d = v.cast::<PyDate>()?;
        out.push_str("@date:");
        push_ymd(out, d.get_year(), d.get_month() as u32, d.get_day() as u32);
        return Ok(Some(Kind::Date));
    }
    if v.is_instance_of::<PyTime>() {
        let t = v.cast::<PyTime>()?;
        if t.get_tzinfo().is_some() {
            out.push_str("@time:");
            out.push_str(&t.call_method0("isoformat")?.extract::<String>()?);
        } else {
            push_time(
                out,
                t.get_hour() as u32,
                t.get_minute() as u32,
                t.get_second() as u32,
                t.get_microsecond(),
            );
        }
        return Ok(Some(Kind::Date));
    }
    if v.is_instance_of::<PyDelta>() {
        let d = v.cast::<PyDelta>()?;
        let total: i128 = d.get_days() as i128 * 86_400_000_000
            + d.get_seconds() as i128 * 1_000_000
            + d.get_microseconds() as i128;
        out.push_str("@td:");
        if total < 0 {
            out.push('-');
        }
        let a = total.unsigned_abs();
        let us = (a % 1_000_000) as u32;
        let secs = a / 1_000_000;
        out.push_str(&(secs / 3600).to_string());
        out.push(':');
        push_2(out, ((secs / 60) % 60) as u32);
        out.push(':');
        push_2(out, (secs % 60) as u32);
        if us != 0 {
            out.push_str(&format!(".{:06}", us));
        }
        return Ok(Some(Kind::Date));
    }
    Err(PyValueError::new_err(format!(
        "openrsxl.extended.llm: cannot encode a value of type {}",
        v.get_type().name()?
    )))
}

// ---------------------------------------------------------------------------
// token estimate (o200k / cl100k-like BPE, calibrated on tiktoken)

pub fn est_tokens(s: &str) -> u64 {
    let b = s.as_bytes();
    let n = b.len();
    let mut i = 0;
    let mut t: u64 = 0;
    while i < n {
        let c = b[i];
        if c >= 0x80 {
            // one token per non-ASCII character (CJK) / per two (accents)
            let ch_len = match c {
                0xF0..=0xFF => 4,
                0xE0..=0xEF => 3,
                _ => 2,
            };
            t += if ch_len >= 3 { 1 } else { 0 };
            if ch_len == 2 {
                // latin letters with accents mostly merge with their word
                t += ((i == 0) || !b[i - 1].is_ascii_alphabetic()) as u64;
            }
            i += ch_len;
            continue;
        }
        if c.is_ascii_alphabetic() {
            let s0 = i;
            while i < n && b[i].is_ascii_alphabetic() {
                i += 1;
            }
            t += ((i - s0) as u64).div_ceil(6);
            continue;
        }
        if c.is_ascii_digit() {
            let s0 = i;
            while i < n && b[i].is_ascii_digit() {
                i += 1;
            }
            t += ((i - s0) as u64).div_ceil(3);
            continue;
        }
        if c == b' ' {
            // a single space joins the next word
            if i + 1 < n && b[i + 1].is_ascii_alphanumeric() {
                i += 1;
                continue;
            }
            let s0 = i;
            while i < n && b[i] == b' ' {
                i += 1;
            }
            t += ((i - s0) as u64).div_ceil(4);
            continue;
        }
        // punctuation: runs of the same character merge (`|||`)
        let s0 = i;
        while i < n && b[i] == c {
            i += 1;
        }
        t += ((i - s0) as u64).div_ceil(4);
    }
    t
}

// ---------------------------------------------------------------------------
// rectangles

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Rect {
    pub r0: u32,
    pub c0: u32,
    pub r1: u32,
    pub c1: u32,
}

/// A horizontal run of cells with the same key in one row.
#[derive(Clone, Copy)]
pub struct Run {
    pub c0: u32,
    pub c1: u32,
    pub key: u32,
    /// never merged with the runs of other rows
    pub lock: bool,
}

/// Streaming rectangle cover: runs of consecutive rows with the same span
/// and key are merged into rectangles. Exact (the rectangles of a key cover
/// exactly its cells), memory proportional to the number of rectangles.
#[derive(Default)]
pub struct Merger {
    /// (key, rect) in creation order: row-major order of the top-left cells
    pub rects: Vec<(u32, Rect)>,
    prev: Vec<(Run, usize)>,
    cur: Vec<(Run, usize)>,
    prev_row: Option<u32>,
}

impl Merger {
    pub fn row(&mut self, row: u32, runs: &[Run]) {
        self.cur.clear();
        let contiguous = self.prev_row.map(|p| row == p + 1).unwrap_or(false);
        let mut j = 0;
        for run in runs {
            let mut found = None;
            if contiguous && !run.lock {
                while j < self.prev.len() && self.prev[j].0.c0 < run.c0 {
                    j += 1;
                }
                if let Some((p, idx)) = self.prev.get(j) {
                    if p.c0 == run.c0 && p.c1 == run.c1 && p.key == run.key && !p.lock {
                        found = Some(*idx);
                    }
                }
            }
            let idx = match found {
                Some(i) => {
                    self.rects[i].1.r1 = row;
                    i
                }
                None => {
                    self.rects.push((
                        run.key,
                        Rect {
                            r0: row,
                            c0: run.c0,
                            r1: row,
                            c1: run.c1,
                        },
                    ));
                    self.rects.len() - 1
                }
            };
            self.cur.push((*run, idx));
        }
        std::mem::swap(&mut self.prev, &mut self.cur);
        self.prev_row = Some(row);
    }

    /// Indices of the rectangles grouped by key (keys ascending), each group
    /// in row-major order.
    pub fn by_key(&self) -> Vec<usize> {
        let mut idx: Vec<usize> = (0..self.rects.len()).collect();
        idx.sort_by_key(|&i| self.rects[i].0); // stable
        idx
    }
}

pub fn push_coord(out: &mut String, row: u32, col: u32) {
    let mut tmp = [0u8; 24];
    out.push_str(col_letters(col as i64, &mut tmp));
    out.push_str(&row.to_string());
}

pub fn push_rect(out: &mut String, r: &Rect) {
    push_coord(out, r.r0, r.c0);
    if r.r1 != r.r0 || r.c1 != r.c0 {
        out.push(':');
        push_coord(out, r.r1, r.c1);
    }
}

fn col_name(col: u32) -> String {
    let mut tmp = [0u8; 24];
    col_letters(col as i64, &mut tmp).to_string()
}

/// A rectangle (`step == 1`) or the columns c0..c1 of every `step`-th row
/// from r0 to r1 (`A6:Q98/4`: rows 6, 10, ..., 98).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Span {
    pub r0: u32,
    pub c0: u32,
    pub r1: u32,
    pub c1: u32,
    pub step: u32,
}

impl Span {
    fn rect(r: &Rect) -> Self {
        Span {
            r0: r.r0,
            c0: r.c0,
            r1: r.r1,
            c1: r.c1,
            step: 1,
        }
    }

    /// number of rows covered
    pub fn rows(&self) -> u32 {
        (self.r1 - self.r0) / self.step + 1
    }

    pub fn covers_row(&self, row: u32) -> bool {
        self.r0 <= row && row <= self.r1 && (row - self.r0).is_multiple_of(self.step)
    }
}

pub fn push_span(out: &mut String, s: &Span) {
    push_coord(out, s.r0, s.c0);
    if s.r1 != s.r0 || s.c1 != s.c0 {
        out.push(':');
        push_coord(out, s.r1, s.c1);
    }
    if s.step > 1 {
        out.push('/');
        out.push_str(&s.step.to_string());
    }
}

/// The rectangles of one key (row-major order) as spans: single-row
/// rectangles of the same columns repeating every `d` rows (at least 3
/// times) become one strided span. Exact; sorted by top-left cell.
pub fn spans_of(rects: &[Rect]) -> Vec<Span> {
    let mut out = Vec::with_capacity(rects.len());
    let mut rows: FxMap<(u32, u32), Vec<u32>> = Default::default();
    let mut order: Vec<(u32, u32)> = Vec::new();
    for r in rects {
        if r.r0 == r.r1 {
            let e = rows.entry((r.c0, r.c1)).or_default();
            if e.is_empty() {
                order.push((r.c0, r.c1));
            }
            e.push(r.r0);
        } else {
            out.push(Span::rect(r));
        }
    }
    for (c0, c1) in order {
        let rs = &rows[&(c0, c1)];
        let n = rs.len();
        let mut i = 0;
        while i < n {
            if i + 2 < n {
                let d = rs[i + 1] - rs[i];
                if d > 1 && rs[i + 2] - rs[i + 1] == d {
                    let mut j = i + 2;
                    while j + 1 < n && rs[j + 1] - rs[j] == d {
                        j += 1;
                    }
                    out.push(Span {
                        r0: rs[i],
                        c0,
                        r1: rs[j],
                        c1,
                        step: d,
                    });
                    i = j + 1;
                    continue;
                }
            }
            out.push(Span {
                r0: rs[i],
                c0,
                r1: rs[i],
                c1,
                step: 1,
            });
            i += 1;
        }
    }
    out.sort_unstable_by_key(|s| (s.r0, s.c0));
    out
}

/// The rectangles of a merger grouped by key: [(key, rects)] in key order.
fn group_by_key(m: &Merger) -> Vec<(u32, Vec<Rect>)> {
    let mut out: Vec<(u32, Vec<Rect>)> = Vec::new();
    for i in m.by_key() {
        let (key, r) = m.rects[i];
        match out.last_mut() {
            Some((k, v)) if *k == key => v.push(r),
            _ => out.push((key, vec![r])),
        }
    }
    out
}

/// Spans of every key of a merger: [(key, spans)] in key order.
fn spans_by_key(m: &Merger) -> Vec<(u32, Vec<Span>)> {
    let order = m.by_key();
    let mut out = Vec::new();
    let mut k = 0;
    let mut group: Vec<Rect> = Vec::new();
    while k < order.len() {
        let key = m.rects[order[k]].0;
        group.clear();
        while k < order.len() && m.rects[order[k]].0 == key {
            group.push(m.rects[order[k]].1);
            k += 1;
        }
        out.push((key, spans_of(&group)));
    }
    out
}

fn push_spans(out: &mut String, spans: &[Span]) {
    for (i, s) in spans.iter().enumerate() {
        if i > 0 {
            out.push(',');
        }
        push_span(out, s);
    }
}

// ---------------------------------------------------------------------------
// cells

/// A formula as written in the formulas section (after the cell / range).
pub enum FIn {
    /// a plain formula: its text (starts with '=' unless malformed)
    Text(String),
    /// `[array REF]` + text
    Array(Option<String>, String),
    /// `[dataTable ATTRS]`
    Table(String),
}

/// A value token, shared (a shared string's token is made once per sheet)
#[derive(Clone)]
pub struct Tok {
    pub s: Arc<str>,
    pub kind: Kind,
    /// estimated LLM tokens
    pub cost: u32,
    /// kept alive by the shared string cache for the whole pass (its
    /// address identifies it)
    pub shared: bool,
}

impl Tok {
    fn new(s: String, kind: Kind) -> Self {
        let cost = est_tokens(&s).min(u32::MAX as u64) as u32;
        Tok {
            s: s.into(),
            kind,
            cost,
            shared: false,
        }
    }

    /// same token (identical pointers first: long repeated strings)
    fn same(&self, other: &Tok) -> bool {
        Arc::ptr_eq(&self.s, &other.s) || self.s == other.s
    }
}

pub struct CellIn {
    pub col: u32,
    pub value: Option<Tok>,
    pub formula: Option<FIn>,
    /// None: not a valid index of the workbook's cell styles
    pub style: Option<u32>,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Enc {
    Auto,
    Grid,
    Index,
    Hybrid,
}

pub struct Opts {
    values: bool,
    formulas: bool,
    styles: bool,
    encoding: Enc,
    /// minimum number of empty cells replaced by a column jump
    jump_gap: u32,
    /// `value*N` runs
    runs: bool,
    /// grid rows: a line is labelled `ROW:` unless it is the row after the
    /// previous line; rows that are multiples of `row_anchor` always are
    row_anchor: u32,
    /// `^` = the value of the cell above (same column, previous line) on
    /// unlabelled lines
    ditto: bool,
    /// memory limit of the inverted index while analysing (bytes)
    index_limit: usize,
    /// automatic choice: index if cheaper than ratio * grid
    index_ratio: f64,
    hybrid_ratio: f64,
    /// hybrid: minimum height of a block
    min_block_rows: u32,
    /// cells left out of every cell section (hidden rows / columns, outside
    /// the print area); None: every cell
    filter: Option<Filter>,
}

/// Cells to leave out: `skip_rows` (sorted, disjoint ranges), `skip_cols`
/// (by column, empty: none), outside `areas` (empty: everywhere).
pub struct Filter {
    skip_rows: Vec<(u32, u32)>,
    skip_cols: Vec<bool>,
    areas: Vec<Rect>,
}

impl Filter {
    #[inline]
    fn keeps(&self, row: u32, col: u32) -> bool {
        if !self.skip_cols.is_empty() && self.skip_cols.get(col as usize).copied().unwrap_or(false)
        {
            return false;
        }
        if !self.skip_rows.is_empty() {
            let i = self.skip_rows.partition_point(|r| r.1 < row);
            if i < self.skip_rows.len() && self.skip_rows[i].0 <= row {
                return false;
            }
        }
        self.areas.is_empty()
            || self
                .areas
                .iter()
                .any(|a| a.r0 <= row && row <= a.r1 && a.c0 <= col && col <= a.c1)
    }

    fn from_dict(d: &Bound<'_, PyDict>) -> PyResult<Option<Self>> {
        let get = |k: &str| -> PyResult<Vec<Vec<i64>>> {
            Ok(match d.get_item(k)? {
                Some(v) if !v.is_none() => v.extract()?,
                _ => Vec::new(),
            })
        };
        let clamp = |v: i64, hi: u32| v.clamp(0, hi as i64) as u32;
        let mut skip_rows: Vec<(u32, u32)> = Vec::new();
        for r in get("skip_rows")? {
            if r.len() != 2 {
                return Err(PyValueError::new_err("skip_rows: (first, last) pairs"));
            }
            let (a, b) = (clamp(r[0], u32::MAX), clamp(r[1], u32::MAX));
            if a <= b {
                skip_rows.push((a, b));
            }
        }
        skip_rows.sort_unstable();
        let mut merged: Vec<(u32, u32)> = Vec::with_capacity(skip_rows.len());
        for (a, b) in skip_rows {
            match merged.last_mut() {
                Some(last) if a <= last.1.saturating_add(1) => last.1 = last.1.max(b),
                _ => merged.push((a, b)),
            }
        }
        let mut skip_cols = Vec::new();
        for c in get("skip_cols")? {
            if c.len() != 2 {
                return Err(PyValueError::new_err("skip_cols: (first, last) pairs"));
            }
            let (a, b) = (clamp(c[0], MAX_COL), clamp(c[1], MAX_COL));
            if a <= b {
                if skip_cols.is_empty() {
                    skip_cols = vec![false; MAX_COL as usize + 1];
                }
                for x in a..=b {
                    skip_cols[x as usize] = true;
                }
            }
        }
        let mut areas = Vec::new();
        let raw_areas = d.get_item("areas")?;
        let has_areas = matches!(&raw_areas, Some(v) if !v.is_none());
        for a in get("areas")? {
            if a.len() != 4 {
                return Err(PyValueError::new_err(
                    "areas: (row1, col1, row2, col2) tuples",
                ));
            }
            areas.push(Rect {
                r0: clamp(a[0], u32::MAX),
                c0: clamp(a[1], MAX_COL),
                r1: clamp(a[2], u32::MAX),
                c1: clamp(a[3], MAX_COL),
            });
        }
        if has_areas && areas.is_empty() {
            // an empty area list: nothing is kept
            areas.push(Rect {
                r0: 1,
                c0: 1,
                r1: 0,
                c1: 0,
            });
        }
        if merged.is_empty() && skip_cols.is_empty() && areas.is_empty() {
            return Ok(None);
        }
        Ok(Some(Filter {
            skip_rows: merged,
            skip_cols,
            areas,
        }))
    }
}

impl Opts {
    fn from_dict(d: &Bound<'_, PyDict>) -> PyResult<Self> {
        let get_bool = |k: &str, dflt: bool| -> PyResult<bool> {
            Ok(match d.get_item(k)? {
                Some(v) => v.is_truthy()?,
                None => dflt,
            })
        };
        let get_f = |k: &str, dflt: f64| -> PyResult<f64> {
            Ok(match d.get_item(k)? {
                Some(v) => v.extract()?,
                None => dflt,
            })
        };
        let get_u = |k: &str, dflt: u64| -> PyResult<u64> {
            Ok(match d.get_item(k)? {
                Some(v) => v.extract()?,
                None => dflt,
            })
        };
        let encoding = match d.get_item("encoding")? {
            Some(v) => match v.extract::<String>()?.as_str() {
                "auto" => Enc::Auto,
                "grid" => Enc::Grid,
                "index" => Enc::Index,
                "hybrid" => Enc::Hybrid,
                other => {
                    return Err(PyValueError::new_err(format!(
                        "unknown encoding {:?}",
                        other
                    )))
                }
            },
            None => Enc::Auto,
        };
        Ok(Opts {
            values: get_bool("values", true)?,
            formulas: get_bool("formulas", true)?,
            styles: get_bool("styles", true)?,
            encoding,
            // 0: every field is written COL=v (cite_cells)
            jump_gap: get_u("jump_gap", 4)? as u32,
            runs: get_bool("runs", true)?,
            row_anchor: get_u("row_anchor", 10)?.clamp(1, u32::MAX as u64) as u32,
            ditto: get_bool("ditto", true)?,
            index_limit: get_u("index_limit", 64 << 20)? as usize,
            index_ratio: get_f("index_ratio", 0.85)?,
            hybrid_ratio: get_f("hybrid_ratio", 0.92)?,
            min_block_rows: get_u("min_block_rows", 3)?.max(2) as u32,
            filter: Filter::from_dict(d)?,
        })
    }
}

/// Template flags: directions in which a fill would not be Excel's.
#[derive(Clone, Copy, Default)]
struct FillFlags {
    /// whole-column ranges, structured / external references, columns
    /// beyond XFD: filling right would shift them in Excel
    h_unsafe: bool,
    /// whole-row ranges: filling down would shift them in Excel
    v_unsafe: bool,
}

/// Classify a formula: references Excel shifts but `ftemplate` does not.
fn fill_flags(f: &str) -> FillFlags {
    let b = f.as_bytes();
    let n = b.len();
    let mut fl = FillFlags::default();
    let ident = |c: u8| c.is_ascii_alphanumeric() || c == b'_' || c == b'.' || c == b'$';
    let mut i = 0;
    while i < n {
        let c = b[i];
        if c == b'"' || c == b'\'' {
            let q = c;
            i += 1;
            while i < n {
                if b[i] == q {
                    if i + 1 < n && b[i + 1] == q {
                        i += 2;
                        continue;
                    }
                    break;
                }
                i += 1;
            }
            i += 1;
            continue;
        }
        if c == b'[' || c == b']' {
            // structured references (Table1[Col], [#This Row]) and external
            // workbooks ([1]Sheet1!A1): filled down they do not change
            // (ftemplate's reading), filled right Excel shifts the columns
            fl.h_unsafe = true;
        }
        if (i == 0 || !ident(b[i - 1])) && (c == b'$' || c.is_ascii_alphanumeric()) {
            // a token: letters / digits / $ / ':' parts
            let s0 = i;
            let mut j = i;
            while j < n && (ident(b[j]) || b[j] == b':') {
                j += 1;
            }
            let tok = &b[s0..j];
            if let Some(p) = tok.iter().position(|&x| x == b':') {
                let (l, r) = (&tok[..p], &tok[p + 1..]);
                let strip =
                    |x: &[u8]| -> Vec<u8> { x.iter().copied().filter(|&c| c != b'$').collect() };
                let (l, r) = (strip(l), strip(r));
                let letters = |x: &[u8]| {
                    !x.is_empty() && x.len() <= 3 && x.iter().all(|c| c.is_ascii_alphabetic())
                };
                let digits = |x: &[u8]| !x.is_empty() && x.iter().all(|c| c.is_ascii_digit());
                if letters(&l) && letters(&r) {
                    fl.h_unsafe = true;
                }
                if digits(&l) && digits(&r) {
                    fl.v_unsafe = true;
                }
            }
            // columns beyond XFD look like references to ftemplate only
            for part in tok.split(|&x| x == b':' || x == b'!') {
                let p: Vec<u8> = part.iter().copied().filter(|&c| c != b'$').collect();
                let letters = p.iter().take_while(|c| c.is_ascii_uppercase()).count();
                if letters == 3
                    && p.len() > 3
                    && p[3].is_ascii_digit()
                    && &p[..3] > b"XFD".as_slice()
                {
                    fl.h_unsafe = true;
                    fl.v_unsafe = true;
                }
            }
            i = j;
            continue;
        }
        i += 1;
    }
    fl
}

struct Templates {
    by_hash: FxMap<u64, Vec<u32>>,
    list: Vec<(Template, FillFlags)>,
}

impl Templates {
    fn new() -> Self {
        Templates {
            by_hash: Default::default(),
            list: Vec::new(),
        }
    }

    fn id(&mut self, f: &str, row: u32, col: u32) -> (u32, FillFlags) {
        let (r, c) = (row as i64, col as i64);
        let h = key_hash(f, r, c).unwrap_or(0);
        if let Some(ids) = self.by_hash.get(&h) {
            for &id in ids {
                let (t, fl) = &self.list[id as usize];
                if t.matches(f, r, c) {
                    return (id, *fl);
                }
            }
        }
        let t = Template::build(f, r, c).expect("template");
        let fl = if t.matches(f, r, c) {
            fill_flags(f)
        } else {
            // cannot happen (ftemplate renders what it scanned); never group
            FillFlags {
                h_unsafe: true,
                v_unsafe: true,
            }
        };
        self.list.push((t, fl));
        let id = (self.list.len() - 1) as u32;
        self.by_hash.entry(h).or_default().push(id);
        (id, fl)
    }
}

/// A formula entry that is not part of a fill group.
struct Single {
    row: u32,
    col: u32,
    line: String,
}

#[derive(Default)]
struct Stats {
    /// a cell / a value was seen (rows can be 0)
    seen: bool,
    vseen: bool,
    cells: u64,
    kinds: [u64; 5],
    rows: u64,
    vmin_row: u32,
    vmax_row: u32,
    vmin_col: u32,
    vmax_col: u32,
    /// bounding box of every cell (values, formulas, styles)
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
    formulas: u64,
    styled: u64,
    grid_est: u64,
    value_tokens: u64,
    /// last row with values (grid row labels) and its values (ditto marks)
    last_vrow: Option<u32>,
    prev_cells: Vec<(u32, Tok)>,
}

/// Pass-1 inverted index of values (bounded memory).
struct Index {
    keys: FxMap<Arc<str>, u32>,
    /// shared tokens by address (no hashing of long strings per cell)
    by_ptr: FxMap<usize, u32>,
    tokens: Vec<Arc<str>>,
    costs: Vec<u32>,
    merger: Merger,
    bytes: usize,
    abandoned: bool,
}

impl Index {
    fn new() -> Self {
        Index {
            keys: Default::default(),
            by_ptr: Default::default(),
            tokens: Vec::new(),
            costs: Vec::new(),
            merger: Merger::default(),
            bytes: 0,
            abandoned: false,
        }
    }

    fn key(&mut self, tok: &Tok) -> u32 {
        let ptr = Arc::as_ptr(&tok.s) as *const u8 as usize;
        if tok.shared {
            if let Some(&k) = self.by_ptr.get(&ptr) {
                return k;
            }
        }
        let k = match self.keys.get(&*tok.s) {
            Some(&k) => k,
            None => {
                let k = self.tokens.len() as u32;
                // one shared allocation + map and vector entries
                self.bytes += tok.s.len() + 64;
                self.keys.insert(tok.s.clone(), k);
                self.costs.push(tok.cost);
                self.tokens.push(tok.s.clone());
                k
            }
        };
        if tok.shared {
            self.by_ptr.insert(ptr, k);
        }
        k
    }

    fn abandon(&mut self) {
        self.abandoned = true;
        self.keys = Default::default();
        self.by_ptr = Default::default();
        self.tokens = Vec::new();
        self.costs = Vec::new();
        self.merger = Merger::default();
    }
}

/// Assembles complete rows from a cell stream.
struct Rows {
    row: u32,
    cells: Vec<CellIn>,
    /// buffered mode (unordered input): everything, emitted at the end
    buffered: Option<BTreeMap<u32, Vec<CellIn>>>,
    last_emitted: i64,
}

const MAX_COL: u32 = 18278;

/// Does a grid line of `row` need its `ROW:` label? (first line, a row that
/// is not the one after the previous line, or an anchor row)
fn needs_label(prev: Option<u32>, row: u32, anchor: u32) -> bool {
    prev.map(|p| p as u64 + 1 != row as u64).unwrap_or(true) || row.is_multiple_of(anchor)
}

/// `^[0-9]+:` - an unlabelled line starting like this would read as labelled
fn looks_labelled(s: &[u8]) -> bool {
    let n = s.iter().take_while(|c| c.is_ascii_digit()).count();
    n > 0 && s.get(n) == Some(&b':')
}

/// Sort a row by column, keeping the last cell of each position (later
/// cells replace earlier ones in full mode's ws._cells)
fn normalise_row(cells: &mut Vec<CellIn>) {
    let sorted = cells.windows(2).all(|w| w[0].col < w[1].col);
    if sorted {
        return;
    }
    // stable sort, then keep the last of equal columns
    let mut v: Vec<(usize, CellIn)> = cells.drain(..).enumerate().collect();
    v.sort_by_key(|(i, c)| (c.col, *i));
    let mut out: Vec<CellIn> = Vec::with_capacity(v.len());
    for (_, c) in v {
        if let Some(last) = out.last_mut() {
            if last.col == c.col {
                *last = c;
                continue;
            }
        }
        out.push(c);
    }
    *cells = out;
}

#[pyclass]
pub struct SheetEncoder {
    opts: Opts,
    pass: u8,
    rows: Rows,
    stats: Stats,
    index: Index,
    templates: Templates,
    fmerger: Merger,
    singles: Vec<Single>,
    smerger: Merger,
    used_styles: std::collections::BTreeSet<u32>,
    bad_style: Option<(u32, u32)>,
    /// chosen values encoding (after pass 1)
    chosen: Enc,
    reason: String,
    est: [Option<u64>; 3],
    /// hybrid blocks (indices into index.merger.rects), and their text
    blocks_text: String,
    nblocks: u64,
    /// covered[col]: the active hybrid blocks of each column
    covered: Vec<Vec<Span>>,
    block_starts: Vec<Span>,
    /// inverted index of the values as spans, by key (index encoding)
    index_spans: Vec<(u32, Vec<Span>)>,
    block_pos: usize,
    grid_start: u32,
    // output of pass 2
    out: String,
    sink: Option<Py<PyAny>>,
    sink_json: bool,
    out_tokens: u64,
    out_bytes: u64,
    /// row of the last grid line written, and its fields (ditto marks)
    last_line_row: Option<u32>,
    prev_line: Vec<(u32, Tok)>,
    /// tokens of the shared strings this sheet uses (sparse: a sheet that
    /// uses string 600000 does not pay for the others), up to SST_CACHE bytes
    sst_tokens: FxMap<u32, Tok>,
    sst_bytes: usize,
    unordered: bool,
    /// cells of pass 1 left out by the filter
    skipped: u64,
}

/// Memory for the shared-string token cache of one sheet. Strings beyond it
/// are converted again at each use (same tokens, only slower).
const SST_CACHE: usize = 32 << 20;

impl SheetEncoder {
    fn reset_pass(&mut self, pass: u8) {
        self.pass = pass;
        self.rows = Rows {
            row: 0,
            cells: Vec::new(),
            buffered: if self.unordered {
                Some(BTreeMap::new())
            } else {
                None
            },
            last_emitted: -1,
        };
        if pass == 1 {
            self.stats = Stats::default();
            self.index = Index::new();
            if self.opts.encoding == Enc::Grid || !self.opts.values {
                self.index.abandoned = true;
            }
            self.templates = Templates::new();
            self.fmerger = Merger::default();
            self.singles = Vec::new();
            self.smerger = Merger::default();
            self.used_styles = Default::default();
            self.bad_style = None;
            self.skipped = 0;
        } else {
            self.out = String::new();
            self.out_tokens = 0;
            self.out_bytes = 0;
            self.last_line_row = None;
            self.prev_line = Vec::new();
            self.block_pos = 0;
            self.covered = Vec::new();
        }
    }

    fn push_cell(&mut self, row: i64, cell: CellIn) -> PyResult<()> {
        // row 0 is a valid position for openpyxl (r="A0" / <row r="0">)
        if !(0..=u32::MAX as i64).contains(&row) || !(1..=MAX_COL).contains(&cell.col) {
            return Err(PyValueError::new_err(format!(
                "openrsxl.extended.llm: cell position out of range (row {}, column {})",
                row, cell.col
            )));
        }
        let row = row as u32;
        if let Some(f) = &self.opts.filter {
            if !f.keeps(row, cell.col) {
                if self.pass == 1 {
                    self.skipped += 1;
                }
                return Ok(());
            }
        }
        if let Some(buf) = self.rows.buffered.as_mut() {
            buf.entry(row).or_default().push(cell);
            return Ok(());
        }
        if row != self.rows.row {
            if row < self.rows.row || row as i64 <= self.rows.last_emitted {
                return Err(unordered_err());
            }
            self.flush_row()?;
            self.rows.row = row;
        }
        self.rows.cells.push(cell);
        Ok(())
    }

    fn flush_row(&mut self) -> PyResult<()> {
        if self.rows.cells.is_empty() {
            return Ok(());
        }
        let row = self.rows.row;
        let mut cells = std::mem::take(&mut self.rows.cells);
        normalise_row(&mut cells);
        self.rows.last_emitted = row as i64;
        if self.pass == 1 {
            self.analyse_row(row, &cells)?;
        } else {
            self.write_row(row, &cells)?;
        }
        cells.clear();
        self.rows.cells = cells;
        Ok(())
    }

    fn finish_pass(&mut self) -> PyResult<()> {
        if let Some(buf) = self.rows.buffered.take() {
            for (row, cells) in buf {
                self.rows.row = row;
                self.rows.cells = cells;
                self.flush_row()?;
            }
            self.rows.buffered = Some(BTreeMap::new());
        } else {
            self.flush_row()?;
        }
        Ok(())
    }

    // -- pass 1 --------------------------------------------------------------
    fn analyse_row(&mut self, row: u32, cells: &[CellIn]) -> PyResult<()> {
        {
            let st = &mut self.stats;
            let (c_first, c_last) = (cells[0].col, cells[cells.len() - 1].col);
            if !st.seen {
                st.seen = true;
                st.min_row = row;
                st.min_col = c_first;
            }
            st.max_row = row;
            st.min_col = st.min_col.min(c_first);
            st.max_col = st.max_col.max(c_last);
        }

        // values: counts, grid cost, inverted index
        if self.opts.values {
            let mut any = false;
            let mut line_cost: u64 = 0;
            let mut prev_col: u32 = 0;
            // previous field: (ditto, token)
            let mut prev_tok: Option<(bool, &Tok)> = None;
            let mut run_len: u64 = 0;
            let mut runs: Vec<Run> = Vec::new();
            let track = !self.index.abandoned;
            let ditto_ok =
                self.opts.ditto && !needs_label(self.stats.last_vrow, row, self.opts.row_anchor);
            let above = std::mem::take(&mut self.stats.prev_cells);
            let mut this_row: Vec<(u32, Tok)> = Vec::new();
            let mut aj = 0;
            for c in cells {
                let Some(tok) = &c.value else {
                    continue;
                };
                while aj < above.len() && above[aj].0 < c.col {
                    aj += 1;
                }
                let ditto = ditto_ok
                    && tok.cost > 1
                    && aj < above.len()
                    && above[aj].0 == c.col
                    && above[aj].1.same(tok);
                if self.opts.ditto {
                    this_row.push((c.col, tok.clone()));
                }
                let kind = &tok.kind;
                let st = &mut self.stats;
                if !any {
                    any = true;
                    st.rows += 1;
                    if !st.vseen {
                        st.vseen = true;
                        st.vmin_row = row;
                        st.vmin_col = c.col;
                    }
                    st.vmax_row = row;
                    st.vmin_col = st.vmin_col.min(c.col);
                }
                st.vmax_col = st.vmax_col.max(c.col);
                st.cells += 1;
                st.kinds[kind_index(*kind)] += 1;
                st.value_tokens += tok.cost as u64;
                // grid cost: separator + token (1 for a ditto mark), runs,
                // empties / jumps
                let cost = if ditto { 1 } else { tok.cost as u64 };
                let consecutive = prev_col != 0 && c.col == prev_col + 1;
                let same = prev_tok
                    .map(|(d, p)| d == ditto && (ditto || p.same(tok)))
                    .unwrap_or(false);
                if consecutive && self.opts.runs && same {
                    run_len += 1;
                    if run_len == 3 {
                        // "x|x|x" -> "x*3"
                        line_cost = line_cost.saturating_sub(2 * cost + 2) + 2;
                    }
                    if run_len < 3 {
                        line_cost += cost + 1;
                    }
                } else {
                    run_len = 1;
                    let gap = if prev_col == 0 {
                        0
                    } else {
                        c.col - prev_col - 1
                    };
                    if gap >= self.opts.jump_gap {
                        line_cost += 3;
                    } else {
                        line_cost += (gap as u64).div_ceil(4);
                    }
                    line_cost += cost + 1;
                }
                prev_col = c.col;
                prev_tok = Some((ditto, tok));
                if track {
                    let key = self.index.key(tok);
                    match runs.last_mut() {
                        Some(r) if r.key == key && r.c1 + 1 == c.col => r.c1 = c.col,
                        _ => runs.push(Run {
                            c0: c.col,
                            c1: c.col,
                            key,
                            lock: false,
                        }),
                    }
                }
            }
            if any {
                // row label "123:" (when needed) + newline
                if needs_label(self.stats.last_vrow, row, self.opts.row_anchor) {
                    line_cost += est_tokens(&row.to_string()) + 1;
                }
                line_cost += 1;
                self.stats.last_vrow = Some(row);
                self.stats.prev_cells = this_row;
                self.stats.grid_est += line_cost;
            } else {
                self.stats.prev_cells = above;
            }
            if track {
                self.index.merger.row(row, &runs);
                if self.opts.encoding == Enc::Auto
                    && self.index.bytes + self.index.merger.rects.len() * 20 > self.opts.index_limit
                {
                    self.index.abandon();
                }
            }
        }

        // formulas: fill groups
        if self.opts.formulas {
            let mut runs: Vec<Run> = Vec::new();
            let mut prev: Option<(u32, u32, bool)> = None; // (col, tid, h_unsafe)
            for c in cells {
                match &c.formula {
                    None => {}
                    Some(FIn::Text(f))
                        if f.starts_with('=') && !f.contains('\n') && !f.contains('\r') =>
                    {
                        self.stats.formulas += 1;
                        let (tid, fl) = self.templates.id(f, row, c.col);
                        let extend = matches!(prev, Some((pc, pt, false)) if pc + 1 == c.col && pt == tid)
                            && !fl.h_unsafe;
                        if extend {
                            runs.last_mut().unwrap().c1 = c.col;
                        } else {
                            runs.push(Run {
                                c0: c.col,
                                c1: c.col,
                                key: tid,
                                lock: fl.v_unsafe,
                            });
                        }
                        prev = Some((c.col, tid, fl.h_unsafe));
                        continue;
                    }
                    Some(f) => {
                        self.stats.formulas += 1;
                        let mut line = String::new();
                        push_coord(&mut line, row, c.col);
                        match f {
                            FIn::Text(t) => {
                                line.push('~');
                                push_json(&mut line, t);
                            }
                            FIn::Array(r, t) => {
                                line.push_str("[array");
                                if let Some(r) = r {
                                    line.push(' ');
                                    line.push_str(r);
                                }
                                line.push(']');
                                push_formula_body(&mut line, t);
                            }
                            FIn::Table(attrs) => {
                                line.push_str("[dataTable");
                                if !attrs.is_empty() {
                                    line.push(' ');
                                    line.push_str(attrs);
                                }
                                line.push(']');
                            }
                        }
                        self.singles.push(Single {
                            row,
                            col: c.col,
                            line,
                        });
                    }
                }
                prev = None;
            }
            self.fmerger.row(row, &runs);
        }

        // styles: map of non-default styles
        if self.opts.styles {
            let mut runs: Vec<Run> = Vec::new();
            for c in cells {
                let s = match c.style {
                    Some(s) => s,
                    None => {
                        if self.bad_style.is_none() {
                            self.bad_style = Some((row, c.col));
                        }
                        continue;
                    }
                };
                self.used_styles.insert(s);
                if s == 0 {
                    continue;
                }
                self.stats.styled += 1;
                match runs.last_mut() {
                    Some(r) if r.key == s && r.c1 + 1 == c.col => r.c1 = c.col,
                    _ => runs.push(Run {
                        c0: c.col,
                        c1: c.col,
                        key: s,
                        lock: false,
                    }),
                }
            }
            self.smerger.row(row, &runs);
        }
        Ok(())
    }

    /// Choose the values encoding from the pass-1 estimates.
    fn decide(&mut self) {
        let grid = self.stats.grid_est;
        let mut index_cost = None;
        let mut hybrid_cost = None;
        self.blocks_text.clear();
        self.block_starts.clear();
        self.nblocks = 0;
        self.index_spans = Vec::new();
        if !self.index.abandoned && self.opts.values {
            self.index_spans = spans_by_key(&self.index.merger);
            // the merger's rectangles are not needed any more
            self.index.merger = Merger::default();
            // index: one line per value: spans ',' ... '=' token '\n'
            let mut cost: u64 = 0;
            let mut s = String::new();
            for (key, spans) in &self.index_spans {
                for sp in spans {
                    s.clear();
                    push_span(&mut s, sp);
                    cost += est_tokens(&s) + 1; // + ',' or '='
                }
                cost += self.index.costs[*key as usize] as u64 + 1;
            }
            index_cost = Some(cost);
            // hybrid: spans of identical values that are cheaper as a range
            // line than in the grid
            let mut gain_total: u64 = 0;
            let mut chosen: Vec<(usize, usize)> = Vec::new();
            for (gi, (key, spans)) in self.index_spans.iter().enumerate() {
                let tc = self.index.costs[*key as usize] as u64;
                for (si, sp) in spans.iter().enumerate() {
                    let rows = sp.rows();
                    if (sp.step == 1 && rows < self.opts.min_block_rows) || rows < 3 {
                        continue;
                    }
                    let w = (sp.c1 - sp.c0 + 1) as u64;
                    let per_row = |tc: u64| {
                        if w >= 3 && self.opts.runs {
                            tc + 3
                        } else {
                            w * (tc + 1)
                        }
                    };
                    // in the grid the rows of a block below its first are
                    // ditto marks (one token per cell), except anchor rows
                    let in_grid = if self.opts.ditto && sp.step == 1 && tc > 1 {
                        let rows = rows as u64;
                        let anchors = 1 + rows / self.opts.row_anchor as u64;
                        let full = anchors.min(rows);
                        full * per_row(tc) + (rows - full) * per_row(1)
                    } else {
                        rows as u64 * per_row(tc)
                    };
                    s.clear();
                    push_span(&mut s, sp);
                    let as_block = est_tokens(&s) + tc + 2;
                    if in_grid > as_block + 2 {
                        gain_total += in_grid - as_block;
                        chosen.push((gi, si));
                    }
                }
            }
            if !chosen.is_empty() {
                hybrid_cost = Some(grid.saturating_sub(gain_total));
                // block lines: grouped by value like the index (chosen is in
                // key order already)
                let mut k = 0;
                while k < chosen.len() {
                    let gi = chosen[k].0;
                    let mut first = true;
                    while k < chosen.len() && chosen[k].0 == gi {
                        let sp = self.index_spans[gi].1[chosen[k].1];
                        if !first {
                            self.blocks_text.push(',');
                        }
                        first = false;
                        push_span(&mut self.blocks_text, &sp);
                        self.block_starts.push(sp);
                        self.nblocks += 1;
                        k += 1;
                    }
                    self.blocks_text.push('=');
                    self.blocks_text
                        .push_str(&self.index.tokens[self.index_spans[gi].0 as usize]);
                    self.blocks_text.push('\n');
                }
                self.block_starts.sort_by_key(|r| (r.r0, r.c0));
            }
        }
        self.est = [Some(grid), index_cost, hybrid_cost];
        let (enc, reason) = match self.opts.encoding {
            Enc::Grid => (Enc::Grid, "requested".to_string()),
            Enc::Index => (Enc::Index, "requested".to_string()),
            Enc::Hybrid => (
                if hybrid_cost.is_some() {
                    Enc::Hybrid
                } else {
                    Enc::Grid
                },
                if hybrid_cost.is_some() {
                    "requested".to_string()
                } else {
                    "requested; no repeated blocks: grid".to_string()
                },
            ),
            Enc::Auto => {
                let g = grid as f64;
                let i = index_cost.map(|x| x as f64);
                let h = hybrid_cost.map(|x| x as f64);
                let density = self.density();
                if self.stats.cells == 0 {
                    (Enc::Grid, "no values".to_string())
                } else if let Some(i) =
                    i.filter(|&i| i < g * self.opts.index_ratio && h.map(|h| i < h).unwrap_or(true))
                {
                    (
                        Enc::Index,
                        format!(
                            "{} sheet (density {:.2}) with repeated values: inverted index ~{:.0}% of grid",
                            if density < 0.3 { "sparse" } else { "repetitive" },
                            density,
                            100.0 * i / g.max(1.0)
                        ),
                    )
                } else if let Some(h) = h.filter(|&h| h < g * self.opts.hybrid_ratio) {
                    (
                        Enc::Hybrid,
                        format!(
                            "dense sheet (density {:.2}) with blocks of repeated values: hybrid ~{:.0}% of grid",
                            density,
                            100.0 * h / g.max(1.0)
                        ),
                    )
                } else {
                    (
                        Enc::Grid,
                        if self.index.abandoned {
                            format!("density {:.2}, mostly distinct values: grid", density)
                        } else {
                            format!("density {:.2}: grid is the cheapest readable form", density)
                        },
                    )
                }
            }
        };
        self.chosen = enc;
        self.reason = reason;
        if enc != Enc::Hybrid {
            self.blocks_text.clear();
            self.block_starts.clear();
            self.nblocks = 0;
        }
        if enc != Enc::Index {
            self.index_spans = Vec::new();
        }
        self.grid_start = if self.stats.vmin_col == 0 {
            1
        } else {
            self.stats.vmin_col
        };
    }

    fn density(&self) -> f64 {
        let st = &self.stats;
        if st.cells == 0 {
            return 0.0;
        }
        let area = (st.vmax_row - st.vmin_row + 1) as f64 * (st.vmax_col - st.vmin_col + 1) as f64;
        st.cells as f64 / area
    }

    fn index_text(&self) -> String {
        let mut out = String::new();
        for (key, spans) in &self.index_spans {
            push_spans(&mut out, spans);
            out.push('=');
            out.push_str(&self.index.tokens[*key as usize]);
            out.push('\n');
        }
        out
    }

    fn formulas_text(&self) -> String {
        // one line per template (all its fills; strided when periodic) and
        // per single entry, in row-major order of their first cell
        let mut lines: Vec<(u32, u32, String)> = Vec::new();
        for (tid, rects) in group_by_key(&self.fmerger) {
            let (t, fl) = &self.templates.list[tid as usize];
            let spans: Vec<Vec<Span>> = if fl.h_unsafe || fl.v_unsafe {
                // a fill is only Excel's in the directions it was found
                rects.iter().map(|r| vec![Span::rect(r)]).collect()
            } else {
                vec![spans_of(&rects)]
            };
            for group in spans {
                let a = group[0];
                let mut line = String::new();
                push_spans(&mut line, &group);
                let f = t
                    .render(a.r0 as i64, a.c0 as i64)
                    .expect("verified template");
                push_formula_body(&mut line, &f);
                lines.push((a.r0, a.c0, line));
            }
        }
        for s in &self.singles {
            lines.push((s.row, s.col, s.line.clone()));
        }
        lines.sort_unstable_by_key(|a| (a.0, a.1));
        let mut out = String::new();
        for (_, _, l) in lines {
            out.push_str(&l);
            out.push('\n');
        }
        out
    }

    fn style_map_text(&self) -> String {
        let mut out = String::new();
        for (key, spans) in spans_by_key(&self.smerger) {
            push_spans(&mut out, &spans);
            out.push_str("=s");
            out.push_str(&key.to_string());
            out.push('\n');
        }
        out
    }

    // -- pass 2 --------------------------------------------------------------
    fn write_row(&mut self, row: u32, cells: &[CellIn]) -> PyResult<()> {
        // activate hybrid blocks starting at or before this row
        while self.block_pos < self.block_starts.len()
            && self.block_starts[self.block_pos].r0 <= row
        {
            let sp = self.block_starts[self.block_pos];
            if self.covered.len() <= sp.c1 as usize {
                self.covered.resize(sp.c1 as usize + 1, Vec::new());
            }
            for c in sp.c0..=sp.c1 {
                let v = &mut self.covered[c as usize];
                v.retain(|s| s.r1 >= row);
                v.push(sp);
            }
            self.block_pos += 1;
        }
        let covered = &self.covered;
        let is_covered = |col: u32| {
            (col as usize) < covered.len()
                && covered[col as usize].iter().any(|s| s.covers_row(row))
        };
        // the fields: (column, token, ditto). A ditto mark `^` = the same
        // token as the previous line's field of this column, used only on
        // lines of the row right after it that are not anchors (an anchor
        // line is self-contained: nothing is ever more than 9 lines up)
        let ditto_ok =
            self.opts.ditto && !needs_label(self.last_line_row, row, self.opts.row_anchor);
        let above = std::mem::take(&mut self.prev_line);
        let mut aj = 0;
        let mut fields: Vec<(u32, &Tok, bool)> = Vec::with_capacity(cells.len());
        for c in cells {
            let Some(t) = &c.value else {
                continue;
            };
            if is_covered(c.col) {
                continue;
            }
            while aj < above.len() && above[aj].0 < c.col {
                aj += 1;
            }
            let d = ditto_ok
                && t.cost > 1
                && aj < above.len()
                && above[aj].0 == c.col
                && above[aj].1.same(t);
            fields.push((c.col, t, d));
        }
        if fields.is_empty() {
            self.prev_line = above;
            return Ok(());
        }
        let out = &mut self.out;
        let line_start = out.len();
        let mut pos = self.grid_start; // column of the next field
        let mut first = true;
        let mut i = 0;
        let n = fields.len();
        while i < n {
            let (col, tok, d) = fields[i];
            // run of identical fields (tokens, or ditto marks) in consecutive
            // columns
            let mut j = i + 1;
            if self.opts.runs {
                while j < n
                    && fields[j].0 == fields[j - 1].0 + 1
                    && fields[j].2 == d
                    && (d || fields[j].1.same(tok))
                {
                    j += 1;
                }
            }
            let mut count = (j - i) as u32;
            let cost = if d { 1 } else { tok.cost as u64 };
            let run = count >= 3 && cost * count as u64 + count as u64 - 1 > cost + 2;
            if !run {
                count = 1;
            }
            let gap = col - pos;
            if !first {
                out.push('|');
            }
            if gap >= self.opts.jump_gap {
                out.push_str(&col_name(col));
                out.push('=');
            } else {
                for _ in 0..gap {
                    out.push('|');
                }
            }
            if d {
                out.push('^');
            } else {
                out.push_str(&tok.s);
            }
            if run {
                out.push('*');
                out.push_str(&count.to_string());
            }
            pos = col + count;
            first = false;
            i += count as usize;
        }
        if self.opts.ditto {
            self.prev_line = fields.iter().map(|&(c, t, _)| (c, t.clone())).collect();
        }
        // the label: unless this is the row after the previous line (and the
        // line cannot be mistaken for a labelled one)
        if needs_label(self.last_line_row, row, self.opts.row_anchor)
            || looks_labelled(&out.as_bytes()[line_start..])
        {
            out.insert_str(line_start, &format!("{}:", row));
        }
        self.last_line_row = Some(row);
        out.push('\n');
        if self.sink.is_some() && self.out.len() >= 1 << 20 {
            self.flush_sink()?;
        }
        Ok(())
    }

    fn flush_sink(&mut self) -> PyResult<()> {
        if self.out.is_empty() {
            return Ok(());
        }
        self.out_tokens += est_tokens(&self.out);
        self.out_bytes += self.out.len() as u64;
        if let Some(sink) = &self.sink {
            Python::attach(|py| -> PyResult<()> {
                if self.sink_json {
                    let mut j = String::with_capacity(self.out.len() + 16);
                    push_json(&mut j, &self.out);
                    // without the surrounding quotes
                    sink.bind(py).call1((&j[1..j.len() - 1],))?;
                } else {
                    sink.bind(py).call1((self.out.as_str(),))?;
                }
                Ok(())
            })?;
            self.out.clear();
        }
        Ok(())
    }

    // -- native cell conversion -----------------------------------------------
    fn val_token<S: crate::xml::Source>(
        &mut self,
        py: Python<'_>,
        p: &SheetParser<S>,
        v: Val,
        error: bool,
    ) -> PyResult<Option<Tok>> {
        let mut out = String::new();
        let kind = match v {
            Val::None => return Ok(None),
            Val::Int(i) => {
                out.push_str(&i.to_string());
                Kind::Number
            }
            Val::Float(f) => {
                if f.is_finite() {
                    push_float(&mut out, f);
                } else {
                    return py_tok(&PyFloat::new(py, f).into_any(), error);
                }
                Kind::Number
            }
            Val::Bool(b) => {
                out.push_str(if b { "TRUE" } else { "FALSE" });
                Kind::Bool
            }
            Val::Sst(i) => {
                if !error {
                    if let Some(t) = self.sst_tokens.get(&i) {
                        return Ok(Some(t.clone()));
                    }
                }
                let obj = p.st.shared_strings.bind(py).get_item(i as usize)?;
                let r = py_tok(&obj, error)?;
                if let Some(t) = r {
                    if error {
                        return Ok(Some(t));
                    }
                    let mut t = t;
                    // (map entry + Arc header)
                    let cost = t.s.len() + 48;
                    if self.sst_bytes + cost <= SST_CACHE {
                        t.shared = true;
                        self.sst_bytes += cost;
                        self.sst_tokens.insert(i, t.clone());
                    }
                    return Ok(Some(t));
                }
                return Ok(None);
            }
            Val::Text(s) => {
                if error {
                    push_error_token(&mut out, &s);
                    Kind::Error
                } else {
                    push_str_token(&mut out, &s);
                    Kind::Text
                }
            }
            Val::Shared { master, dest } => {
                // never a value (formulas are handled by formula_of)
                let s = shared_text(p, master, dest)?;
                push_str_token(&mut out, &s);
                Kind::Text
            }
            Val::DateInt(i) => push_excel_date(
                &mut out,
                crate::dates::from_excel_int(i, p.st.fast_epoch.as_ref().unwrap()),
            ),
            Val::DateFloat(f) => push_excel_date(
                &mut out,
                crate::dates::from_excel_float(f, p.st.fast_epoch.as_ref().unwrap()),
            ),
            Val::Py(o) => return py_tok(o.bind(py), error),
        };
        Ok(Some(Tok::new(out, kind)))
    }

    fn formula_of<S: crate::xml::Source>(
        &self,
        py: Python<'_>,
        p: &SheetParser<S>,
        v: Val,
        info: &Bound<'_, PyAny>,
    ) -> PyResult<Option<FIn>> {
        Ok(match v {
            Val::None => None,
            Val::Text(s) => Some(FIn::Text(s)),
            Val::Shared { master, dest } => Some(FIn::Text(shared_text(p, master, dest)?)),
            Val::Py(o) => py_formula(o.bind(py), info)?,
            // a formula is text or a formula object
            _ => None,
        })
    }
}

fn shared_text<S: crate::xml::Source>(
    p: &SheetParser<S>,
    master: u32,
    dest: Option<(i64, i64)>,
) -> PyResult<String> {
    match &p.masters[master as usize] {
        Shared::Native(t) => t
            .translate(dest)
            .map_err(|_| PyValueError::new_err("shared formula translation")),
        Shared::Py(_) => unreachable!("python translators produce values directly"),
    }
}

fn py_tok(v: &Bound<'_, PyAny>, error: bool) -> PyResult<Option<Tok>> {
    let mut out = String::new();
    Ok(py_value_token(v, error, &mut out)?.map(|k| Tok::new(out, k)))
}

/// A formula object (str, ArrayFormula, DataTableFormula) through the
/// Python helper `formula_info(obj) -> ("text", s) | ("array", ref, text) |
/// ("table", attrs)`.
fn py_formula(o: &Bound<'_, PyAny>, info: &Bound<'_, PyAny>) -> PyResult<Option<FIn>> {
    if o.is_none() {
        return Ok(None);
    }
    if let Ok(s) = o.cast::<PyString>() {
        return Ok(Some(FIn::Text(s.to_cow()?.into_owned())));
    }
    let t = info.call1((o,))?;
    let kind: String = t.get_item(0)?.extract()?;
    Ok(Some(match kind.as_str() {
        "array" => FIn::Array(t.get_item(1)?.extract()?, t.get_item(2)?.extract()?),
        "table" => FIn::Table(t.get_item(1)?.extract()?),
        _ => FIn::Text(t.get_item(1)?.extract()?),
    }))
}

/// `=body` (the formula without its '=') or `~"json"` when the text cannot
/// stand on one line or does not start with '='.
fn push_formula_body(out: &mut String, f: &str) {
    if f.starts_with('=') && !f.contains('\n') && !f.contains('\r') {
        out.push_str(f);
    } else {
        out.push('~');
        push_json(out, f);
    }
}

pub fn unordered_err() -> PyErr {
    PyValueError::new_err("openrsxl.extended.llm: cells out of order")
}

fn style_u32(s: &StyleId, n_styles: Option<u32>) -> Option<u32> {
    match s {
        StyleId::Int(i) if *i >= 0 && *i <= u32::MAX as i64 => {
            let i = *i as u32;
            match n_styles {
                Some(n) if i >= n => None,
                _ => Some(i),
            }
        }
        _ => None,
    }
}

#[pymethods]
impl SheetEncoder {
    #[new]
    fn new(options: &Bound<'_, PyDict>) -> PyResult<Self> {
        let opts = Opts::from_dict(options)?;
        let mut e = SheetEncoder {
            opts,
            pass: 1,
            rows: Rows {
                row: 0,
                cells: Vec::new(),
                buffered: None,
                last_emitted: -1,
            },
            stats: Stats::default(),
            index: Index::new(),
            templates: Templates::new(),
            fmerger: Merger::default(),
            singles: Vec::new(),
            smerger: Merger::default(),
            used_styles: Default::default(),
            bad_style: None,
            chosen: Enc::Grid,
            reason: String::new(),
            est: [None; 3],
            blocks_text: String::new(),
            nblocks: 0,
            covered: Vec::new(),
            block_starts: Vec::new(),
            index_spans: Vec::new(),
            block_pos: 0,
            grid_start: 1,
            out: String::new(),
            sink: None,
            sink_json: false,
            out_tokens: 0,
            out_bytes: 0,
            last_line_row: None,
            prev_line: Vec::new(),
            sst_tokens: Default::default(),
            sst_bytes: 0,
            unordered: false,
            skipped: 0,
        };
        e.reset_pass(1);
        Ok(e)
    }

    /// Start pass 1 or 2 (from the beginning of the sheet). `buffered`:
    /// accept cells in any order (kept in memory until `end_pass`).
    #[pyo3(signature = (n, buffered=None))]
    fn begin_pass(&mut self, n: u8, buffered: Option<bool>) -> PyResult<()> {
        if let Some(b) = buffered {
            self.unordered = b;
        }
        if n == 2 && self.pass != 2 && self.est[0].is_none() {
            return Err(PyValueError::new_err("pass 1 first"));
        }
        self.reset_pass(n);
        Ok(())
    }

    /// Sink for pass 2: a callable receiving chunks of the values text
    /// (JSON-escaped, without quotes, if `json`). None: kept in memory.
    #[pyo3(signature = (sink, json=false))]
    fn set_sink(&mut self, sink: Option<Py<PyAny>>, json: bool) {
        self.sink = sink;
        self.sink_json = json;
    }

    /// Does the chosen encoding need pass 2 (the grid)?
    #[getter]
    fn needs_pass2(&self) -> bool {
        self.opts.values && self.stats.cells > 0 && matches!(self.chosen, Enc::Grid | Enc::Hybrid)
    }

    /// Read the whole sheet with the native parser (raises NativeFallback
    /// when the Python implementation must be used instead).
    #[pyo3(signature = (source, parser, helpers, formula_info, n_styles=None))]
    fn feed_native(
        &mut self,
        py: Python<'_>,
        source: Py<PyAny>,
        parser: &Bound<'_, PyAny>,
        helpers: &Bound<'_, PyDict>,
        formula_info: &Bound<'_, PyAny>,
        n_styles: Option<u32>,
    ) -> PyResult<()> {
        let h = Helpers::from_dict(helpers)?;
        let st = Settings::from_parser(parser)?;
        let actions = PyList::empty(py).unbind();
        let row_dims = PyDict::new(py).unbind();
        let mut p = SheetParser::new(py, PyFileSource::new(source), h, st, actions, row_dims);
        let want_values = self.opts.values;
        let want_formulas = self.opts.formulas;
        loop {
            let item = p.next_item(py).map_err(|e| fallback_err(py, e))?;
            match item {
                Item::Eof => break,
                Item::Element(..) => {}
                Item::Row(_, recs) => {
                    for c in recs {
                        let (value, formula) = if c.formula_primary {
                            let v = match (want_values, c.cached) {
                                (true, Some(v)) => self.val_token(py, &p, v, c.cached_error)?,
                                _ => None,
                            };
                            let f = if want_formulas {
                                self.formula_of(py, &p, c.val, formula_info)?
                            } else {
                                None
                            };
                            (v, f)
                        } else {
                            let err = matches!(c.dtype, DType::Code(DT_E));
                            let v = if want_values {
                                self.val_token(py, &p, c.val, err)?
                            } else {
                                None
                            };
                            (v, None)
                        };
                        if !(1..=MAX_COL as i64).contains(&c.column) {
                            return Err(PyValueError::new_err(format!(
                                "openrsxl.extended.llm: column {} out of range",
                                c.column
                            )));
                        }
                        let cell = CellIn {
                            col: c.column as u32,
                            value,
                            formula,
                            style: style_u32(&c.style_id, n_styles),
                        };
                        self.push_cell(c.row, cell)?;
                    }
                }
            }
        }
        Ok(())
    }

    /// Feed cells read by Python: [(row, column, value, data_type, formula,
    /// style_id), ...] - value as read with data_only=True, formula as read
    /// with data_only=False (None for other cells).
    #[pyo3(signature = (cells, formula_info, n_styles=None))]
    fn feed(
        &mut self,
        cells: &Bound<'_, PyList>,
        formula_info: &Bound<'_, PyAny>,
        n_styles: Option<u32>,
    ) -> PyResult<()> {
        for item in cells.iter() {
            let row: i64 = item.get_item(0)?.extract()?;
            let col: i64 = item.get_item(1)?.extract()?;
            let value = item.get_item(2)?;
            let dtype = item.get_item(3)?;
            let formula = item.get_item(4)?;
            let style = item.get_item(5)?;
            if !(1..=MAX_COL as i64).contains(&col) {
                return Err(PyValueError::new_err(format!(
                    "openrsxl.extended.llm: column {} out of range",
                    col
                )));
            }
            let err = dtype.extract::<String>().map(|d| d == "e").unwrap_or(false);
            let v = if self.opts.values {
                py_tok(&value, err)?
            } else {
                None
            };
            let f = if self.opts.formulas {
                py_formula(&formula, formula_info)?
            } else {
                None
            };
            let s = match style.extract::<i64>() {
                Ok(i) if !style.is_instance_of::<PyBool>() => style_u32(&StyleId::Int(i), n_styles),
                _ => None,
            };
            self.push_cell(
                row,
                CellIn {
                    col: col as u32,
                    value: v,
                    formula: f,
                    style: s,
                },
            )?;
        }
        Ok(())
    }

    /// End the current pass. Pass 1 returns the analysis.
    fn end_pass(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyDict>>> {
        self.finish_pass()?;
        if self.pass == 2 {
            if self.sink.is_some() {
                self.flush_sink()?;
            } else {
                self.out_tokens = est_tokens(&self.out);
                self.out_bytes = self.out.len() as u64;
            }
            return Ok(None);
        }
        if self.opts.styles {
            if let Some((r, c)) = self.bad_style {
                let mut coord = String::new();
                push_coord(&mut coord, r, c);
                return Err(PyValueError::new_err(format!(
                    "openrsxl.extended.llm: cell {} has an invalid style id",
                    coord
                )));
            }
        }
        self.decide();
        let st = &self.stats;
        let d = PyDict::new(py);
        d.set_item("cells", st.cells)?;
        d.set_item("rows", st.rows)?;
        let kinds = PyDict::new(py);
        for (i, name) in KIND_NAMES.iter().enumerate() {
            if st.kinds[i] > 0 {
                kinds.set_item(*name, st.kinds[i])?;
            }
        }
        d.set_item("kinds", kinds)?;
        if st.seen {
            let mut dim = String::new();
            push_rect(
                &mut dim,
                &Rect {
                    r0: st.min_row,
                    c0: st.min_col,
                    r1: st.max_row,
                    c1: st.max_col,
                },
            );
            d.set_item("dimension", dim)?;
        } else {
            d.set_item("dimension", py.None())?;
        }
        if st.cells > 0 {
            let mut vr = String::new();
            push_rect(
                &mut vr,
                &Rect {
                    r0: st.vmin_row,
                    c0: st.vmin_col,
                    r1: st.vmax_row,
                    c1: st.vmax_col,
                },
            );
            d.set_item("values_range", vr)?;
            d.set_item("density", (self.density() * 10000.0).round() / 10000.0)?;
        } else {
            d.set_item("values_range", py.None())?;
            d.set_item("density", 0.0)?;
        }
        if !self.index.abandoned {
            d.set_item("distinct_values", self.index.tokens.len())?;
            d.set_item(
                "distinct_ratio",
                if st.cells > 0 {
                    ((self.index.tokens.len() as f64 / st.cells as f64) * 10000.0).round() / 10000.0
                } else {
                    0.0
                },
            )?;
            d.set_item("value_rectangles", self.index.merger.rects.len())?;
        } else {
            d.set_item("distinct_values", py.None())?;
            d.set_item("distinct_ratio", py.None())?;
            d.set_item("value_rectangles", py.None())?;
        }
        d.set_item(
            "avg_value_tokens",
            if st.cells > 0 {
                ((st.value_tokens as f64 / st.cells as f64) * 100.0).round() / 100.0
            } else {
                0.0
            },
        )?;
        let est = PyDict::new(py);
        let names = ["grid", "index", "hybrid"];
        for (i, n) in names.iter().enumerate() {
            match self.est[i] {
                Some(v) if self.opts.values => est.set_item(*n, v)?,
                _ => est.set_item(*n, py.None())?,
            }
        }
        d.set_item("estimated_tokens", est)?;
        d.set_item(
            "encoding",
            match self.chosen {
                Enc::Grid => "grid",
                Enc::Index => "index",
                Enc::Hybrid => "hybrid",
                Enc::Auto => "auto",
            },
        )?;
        d.set_item("reason", &self.reason)?;
        d.set_item("grid_start", col_name(self.grid_start))?;
        d.set_item("blocks", self.nblocks)?;
        d.set_item("formulas", st.formulas)?;
        d.set_item("formula_templates", self.templates.list.len())?;
        d.set_item(
            "formula_groups",
            self.fmerger.rects.len() + self.singles.len(),
        )?;
        d.set_item("styled_cells", st.styled)?;
        d.set_item("skipped_cells", self.skipped)?;
        d.set_item("style_ranges", self.smerger.rects.len())?;
        d.set_item(
            "used_styles",
            self.used_styles.iter().copied().collect::<Vec<u32>>(),
        )?;
        Ok(Some(d.unbind()))
    }

    /// The sections computed by pass 1: {"values": str | None (None: the
    /// grid of pass 2), "blocks": str (hybrid blocks, written before the
    /// grid), "formulas": str, "style_map": str}
    fn sections(&mut self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        let d = PyDict::new(py);
        if self.opts.values {
            match self.chosen {
                Enc::Index => d.set_item("values", self.index_text())?,
                _ if self.stats.cells == 0 => d.set_item("values", "")?,
                _ => d.set_item("values", py.None())?,
            }
            d.set_item("blocks", &self.blocks_text)?;
        }
        if self.opts.formulas {
            d.set_item("formulas", self.formulas_text())?;
        }
        if self.opts.styles {
            d.set_item("style_map", self.style_map_text())?;
        }
        Ok(d.unbind())
    }

    /// Release the memory of pass 1 (after `sections`); pass 2 still works.
    fn release(&mut self) {
        self.index = Index::new();
        self.index.abandoned = true;
        self.templates = Templates::new();
        self.fmerger = Merger::default();
        self.singles = Vec::new();
        self.smerger = Merger::default();
        self.index_spans = Vec::new();
        self.sst_tokens = Default::default();
        self.sst_bytes = 0;
    }

    /// The grid lines of pass 2 kept in memory (without hybrid blocks).
    fn take_grid(&mut self) -> String {
        self.sst_tokens = Default::default();
        self.sst_bytes = 0;
        std::mem::take(&mut self.out)
    }

    /// (estimated tokens, bytes) of the grid written by pass 2
    #[getter]
    fn grid_size(&self) -> (u64, u64) {
        (self.out_tokens, self.out_bytes)
    }
}

// ---------------------------------------------------------------------------
// helpers exposed for tests and the Python side

/// The token of a value (as written in a values section).
#[pyfunction]
#[pyo3(signature = (value, error=false))]
pub fn llm_token(value: &Bound<'_, PyAny>, error: bool) -> PyResult<Option<String>> {
    let mut out = String::new();
    Ok(py_value_token(value, error, &mut out)?.map(|_| out))
}

/// Estimated number of LLM tokens of a text.
#[pyfunction]
pub fn llm_estimate_tokens(text: &str) -> u64 {
    est_tokens(text)
}

/// Render the formula `f` written for (row, col) at (row2, col2) with the
/// fill semantics of the formulas section (None: off the sheet).
#[pyfunction]
pub fn llm_fill_formula(f: &str, row: i64, col: i64, row2: i64, col2: i64) -> Option<String> {
    Template::build(f, row, col)?.render(row2, col2)
}

/// (h_unsafe, v_unsafe) of a formula (tests)
#[pyfunction]
pub fn llm_fill_flags(f: &str) -> (bool, bool) {
    let fl = fill_flags(f);
    (fl.h_unsafe, fl.v_unsafe)
}

pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<SheetEncoder>()?;
    m.add_function(wrap_pyfunction!(llm_token, m)?)?;
    m.add_function(wrap_pyfunction!(llm_estimate_tokens, m)?)?;
    m.add_function(wrap_pyfunction!(llm_fill_formula, m)?)?;
    m.add_function(wrap_pyfunction!(llm_fill_flags, m)?)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn floats() {
        let f = |v: f64| {
            let mut s = String::new();
            push_float(&mut s, v);
            s
        };
        assert_eq!(f(1.0), "1.0");
        assert_eq!(f(0.1), "0.1");
        assert_eq!(f(1e16), "1e+16");
        assert_eq!(f(1e15), "1000000000000000.0");
        assert_eq!(f(1e-5), "1e-05");
        assert_eq!(f(0.0001), "0.0001");
        assert_eq!(f(-2.5e-7), "-2.5e-07");
        assert_eq!(f(123456789012345680.0), "1.2345678901234568e+17");
        assert_eq!(f(0.30000000000000004), "0.30000000000000004");
        assert_eq!(f(-0.0), "-0.0");
        assert_eq!(f(1.5e300), "1.5e+300");
    }

    #[test]
    fn plain_strings() {
        for s in [
            "Alice", "New York", "a=b", "x*", "10:30", "2024-1-5", "1,000", "TRUEx", "e5",
        ] {
            assert!(str_is_plain(s), "{}", s);
        }
        for s in [
            "",
            " a",
            "a ",
            "12",
            "-3.5",
            "1e5",
            ".5",
            "5.",
            "TRUE",
            "#N/A",
            "@x",
            "=A1",
            "\"q",
            "a|b",
            "a\nb",
            "2024-01-05",
            "10:30:00",
            "2024-01-05 10:30:00.5",
            "x*3",
            "B=1",
            "ABC=x",
        ] {
            assert!(!str_is_plain(s), "{}", s);
        }
    }

    #[test]
    fn merger() {
        let mut m = Merger::default();
        let run = |c0, c1, key| Run {
            c0,
            c1,
            key,
            lock: false,
        };
        m.row(1, &[run(1, 3, 0)]);
        m.row(2, &[run(1, 1, 1), run(2, 3, 0)]);
        m.row(3, &[run(1, 1, 1), run(2, 3, 0)]);
        m.row(5, &[run(1, 1, 1)]);
        let r: Vec<(u32, Rect)> = m.rects.clone();
        assert_eq!(r.len(), 4);
        assert_eq!(
            r[0],
            (
                0,
                Rect {
                    r0: 1,
                    c0: 1,
                    r1: 1,
                    c1: 3
                }
            )
        );
        assert_eq!(
            r[1],
            (
                1,
                Rect {
                    r0: 2,
                    c0: 1,
                    r1: 3,
                    c1: 1
                }
            )
        );
        assert_eq!(
            r[2],
            (
                0,
                Rect {
                    r0: 2,
                    c0: 2,
                    r1: 3,
                    c1: 3
                }
            )
        );
        assert_eq!(
            r[3],
            (
                1,
                Rect {
                    r0: 5,
                    c0: 1,
                    r1: 5,
                    c1: 1
                }
            )
        );
    }

    #[test]
    fn flags() {
        let f = |s: &str| {
            let x = fill_flags(s);
            (x.h_unsafe, x.v_unsafe)
        };
        assert_eq!(f("=A1+B2"), (false, false));
        assert_eq!(f("=SUM(A:A)"), (true, false));
        assert_eq!(f("=SUM($A:$C)"), (true, false));
        assert_eq!(f("=SUM(1:1)"), (false, true));
        assert_eq!(f("=SUM(A1:B5)"), (false, false));
        assert_eq!(f("=Table1[Col]"), (true, false));
        assert_eq!(f("=t[[#This Row],[A]]+IF(ISNA(B1),#N/A,1)"), (true, false));
        assert_eq!(f("=A1#"), (false, false));
        assert_eq!(f("=\"A:A\"&B1"), (false, false));
        assert_eq!(f("=XFE1"), (true, true));
        assert_eq!(f("=Sheet1!A1:B2"), (false, false));
    }
}
