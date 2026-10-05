//! Position independent formula templates, used to de-duplicate the text of
//! plain (non shared) formulae: `=A2*B2` in C2 and `=A3*B3` in C3 are the
//! same template.
//!
//! The scanner only needs to be *consistent*, not to understand formulae: a
//! template is used for a cell only after verifying that rendering it at that
//! cell reproduces the original text byte for byte, otherwise the raw text is
//! stored. Correctness therefore never depends on the scanner (nor on hash
//! collisions).

use std::hash::Hasher;

use crate::store::FxHasher;
use crate::utils::col_index_from_letters;

#[derive(Debug)]
pub enum Part {
    Text(Box<str>),
    /// column: absolute ($) or offset from the cell's column
    Col {
        abs: bool,
        v: i64,
    },
    /// row: absolute ($) or offset from the cell's row
    Row {
        abs: bool,
        v: i64,
    },
}

pub struct Template {
    parts: Box<[Part]>,
}

#[inline]
fn is_ident(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b'.' || b == b'$'
}

/// One reference found by the scanner.
struct Ref {
    start: usize,
    end: usize,
    col_abs: bool,
    col: i64,
    row_abs: bool,
    row: i64,
}

enum Item<'a> {
    Text(&'a str),
    Ref(Ref),
}

/// Scan `f` and report literal spans / references through the callback.
/// Returns whether at least one reference was found.
#[inline]
fn scan_with<'a>(f: &'a str, mut cb: impl FnMut(Item<'a>)) -> bool {
    let b = f.as_bytes();
    let n = b.len();
    let mut lit_start = 0;
    let mut i = 0;
    let mut found = false;
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
        if (c == b'$' || c.is_ascii_uppercase()) && (i == 0 || !is_ident(b[i - 1])) {
            let mut j = i;
            let col_abs = c == b'$';
            if col_abs {
                j += 1;
            }
            let ls = j;
            while j < n && j - ls < 4 && b[j].is_ascii_uppercase() {
                j += 1;
            }
            let le = j;
            let row_abs = j < n && b[j] == b'$';
            if row_abs {
                j += 1;
            }
            let ds = j;
            let mut row: i64 = 0;
            while j < n && j - ds < 8 && b[j].is_ascii_digit() {
                row = row * 10 + (b[j] - b'0') as i64;
                j += 1;
            }
            let de = j;
            if (1..=3).contains(&(le - ls))
                && (1..=7).contains(&(de - ds))
                && b[ds] != b'0'
                && (j >= n || !(is_ident(b[j]) || b[j] == b'(' || b[j] == b'!'))
            {
                if lit_start < i {
                    cb(Item::Text(&f[lit_start..i]));
                }
                let col = col_index_from_letters(&b[ls..le]).unwrap();
                cb(Item::Ref(Ref {
                    start: i,
                    end: j,
                    col_abs,
                    col,
                    row_abs,
                    row,
                }));
                found = true;
                i = j;
                lit_start = i;
                continue;
            }
        }
        i += 1;
    }
    if lit_start < n {
        cb(Item::Text(&f[lit_start..]));
    }
    found
}

/// Hash of the position independent form of `f` at (row, col), or None if
/// `f` contains no reference. Allocation free.
pub fn key_hash(f: &str, row: i64, col: i64) -> Option<u64> {
    let mut h = FxHasher::default();
    let found = scan_with(f, |it| match it {
        Item::Text(t) => {
            h.write_u64(0);
            h.write(t.as_bytes());
            h.write_u64(t.len() as u64);
        }
        Item::Ref(r) => {
            h.write_u64(1 + r.col_abs as u64);
            h.write_u64((if r.col_abs { r.col } else { r.col - col }) as u64);
            h.write_u64(3 + r.row_abs as u64);
            h.write_u64((if r.row_abs { r.row } else { r.row - row }) as u64);
        }
    });
    let _ = found; // formulae without references are de-duplicated by text
    Some(h.finish())
}

impl Template {
    /// Build the template of `f` at (row, col).
    pub fn build(f: &str, row: i64, col: i64) -> Option<Template> {
        let mut out: Vec<Part> = Vec::new();
        let found = scan_with(f, |it| match it {
            Item::Text(t) => out.push(Part::Text(t.into())),
            Item::Ref(r) => {
                let _ = (r.start, r.end);
                out.push(Part::Col {
                    abs: r.col_abs,
                    v: if r.col_abs { r.col } else { r.col - col },
                });
                out.push(Part::Row {
                    abs: r.row_abs,
                    v: if r.row_abs { r.row } else { r.row - row },
                });
            }
        });
        let _ = found;
        Some(Template {
            parts: out.into_boxed_slice(),
        })
    }

    /// Render the template for the cell at (row, col); None if a reference
    /// would fall outside the sheet.
    pub fn render(&self, row: i64, col: i64) -> Option<String> {
        let mut out = String::with_capacity(32);
        if self.render_into(row, col, &mut out) {
            Some(out)
        } else {
            None
        }
    }

    /// Render appending to `out`; false if a reference falls off the sheet.
    pub fn render_into(&self, row: i64, col: i64, out: &mut String) -> bool {
        let mut tmp = [0u8; 24];
        for p in self.parts.iter() {
            match p {
                Part::Text(t) => out.push_str(t),
                Part::Col { abs, v } => {
                    let c = if *abs { *v } else { col + v };
                    if !(1..=18278).contains(&c) {
                        return false;
                    }
                    if *abs {
                        out.push('$');
                    }
                    out.push_str(col_letters(c, &mut tmp));
                }
                Part::Row { abs, v } => {
                    let r = if *abs { *v } else { row + v };
                    if r < 1 {
                        return false;
                    }
                    if *abs {
                        out.push('$');
                    }
                    out.push_str(fmt_i64(r, &mut tmp));
                }
            }
        }
        true
    }

    /// Does rendering at (row, col) give exactly `f`? (no allocation)
    pub fn matches(&self, f: &str, row: i64, col: i64) -> bool {
        let mut rest = f;
        let mut tmp = [0u8; 24];
        for p in self.parts.iter() {
            let piece: &str = match p {
                Part::Text(t) => t,
                Part::Col { abs, v } => {
                    let c = if *abs { *v } else { col + v };
                    if !(1..=18278).contains(&c) {
                        return false;
                    }
                    if *abs {
                        match rest.strip_prefix('$') {
                            Some(r) => rest = r,
                            None => return false,
                        }
                    }
                    col_letters(c, &mut tmp)
                }
                Part::Row { abs, v } => {
                    let r = if *abs { *v } else { row + v };
                    if r < 1 {
                        return false;
                    }
                    if *abs {
                        match rest.strip_prefix('$') {
                            Some(x) => rest = x,
                            None => return false,
                        }
                    }
                    fmt_i64(r, &mut tmp)
                }
            };
            match rest.strip_prefix(piece) {
                Some(r) => rest = r,
                None => return false,
            }
        }
        rest.is_empty()
    }
}

pub fn col_letters(c: i64, buf: &mut [u8; 24]) -> &str {
    let mut tmp = [0u8; 3];
    let mut n = 0;
    let mut x = c;
    while x > 0 {
        tmp[n] = b'A' + ((x - 1) % 26) as u8;
        n += 1;
        x = (x - 1) / 26;
    }
    for i in 0..n {
        buf[i] = tmp[n - 1 - i];
    }
    unsafe { std::str::from_utf8_unchecked(&buf[..n]) }
}

pub fn fmt_i64(mut v: i64, buf: &mut [u8; 24]) -> &str {
    let mut i = buf.len();
    let neg = v < 0;
    if v == 0 {
        i -= 1;
        buf[i] = b'0';
    }
    while v != 0 {
        i -= 1;
        buf[i] = b'0' + (v % 10).unsigned_abs() as u8;
        v /= 10;
    }
    if neg {
        i -= 1;
        buf[i] = b'-';
    }
    unsafe { std::str::from_utf8_unchecked(&buf[i..]) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn roundtrip() {
        for (f, r, c) in [
            ("=A2*B2", 2, 3),
            ("=SUM($A$1:A5)+'S h'!B2", 5, 4),
            ("=\"A1\"&C3", 3, 1),
            ("=IF(XFD1,1,2)", 1, 1),
            ("=AB12+ZZ$3", 12, 30),
        ] {
            let t = Template::build(f, r, c).unwrap();
            assert!(t.matches(f, r, c), "{}", f);
            assert_eq!(t.render(r, c).unwrap(), f);
        }
        assert_eq!(key_hash("=A2*B2", 2, 3), key_hash("=A3*B3", 3, 3));
        assert_ne!(key_hash("=A2*B2", 2, 3), key_hash("=A2*B2", 3, 3));
        let t = Template::build("=A2*B2", 2, 3).unwrap();
        assert!(t.matches("=A3*B3", 3, 3));
        assert!(!t.matches("=A3*B4", 3, 3));
    }
}
