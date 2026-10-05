//! Small helpers reproducing Python / openpyxl semantics exactly.

/// `column_index_from_string` for ASCII letters (case-insensitive).
pub fn col_index_from_letters(b: &[u8]) -> Option<i64> {
    if b.is_empty() || b.len() > 3 {
        return None;
    }
    let mut idx: i64 = 0;
    for &c in b {
        let u = c.to_ascii_uppercase();
        if !u.is_ascii_uppercase() {
            return None;
        }
        idx = idx * 26 + (u - b'A' + 1) as i64;
    }
    if idx > 0 && idx < 18279 {
        Some(idx)
    } else {
        None
    }
}

/// `get_column_letter`
pub fn column_letter(mut idx: i64) -> Option<String> {
    if !(1..=18278).contains(&idx) {
        return None;
    }
    let mut out = [0u8; 3];
    let mut n = 0;
    while idx > 0 {
        let rem = ((idx - 1) % 26) as u8;
        out[n] = b'A' + rem;
        n += 1;
        idx = (idx - 1) / 26;
    }
    out[..n].reverse();
    Some(String::from_utf8(out[..n].to_vec()).unwrap())
}

pub fn push_column_letter(out: &mut Vec<u8>, mut idx: i64) {
    let mut tmp = [0u8; 8];
    let mut n = 0;
    while idx > 0 && n < 8 {
        let rem = ((idx - 1) % 26) as u8;
        tmp[n] = b'A' + rem;
        n += 1;
        idx = (idx - 1) / 26;
    }
    for i in (0..n).rev() {
        out.push(tmp[i]);
    }
}

/// Fast path of `coordinate_to_tuple` for canonical coordinates such as
/// "AB12" (uppercase letters, no leading zero). Returns None if the
/// coordinate is not canonical (the Python implementation must be used).
pub fn fast_coordinate(s: &[u8]) -> Option<(i64, i64)> {
    let mut i = 0;
    while i < s.len() && s[i].is_ascii_uppercase() {
        i += 1;
    }
    if i == 0 || i > 3 || i == s.len() {
        return None;
    }
    let digits = &s[i..];
    if digits[0] == b'0' || digits.len() > 9 || !digits.iter().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let mut row: i64 = 0;
    for &d in digits {
        row = row * 10 + (d - b'0') as i64;
    }
    let col = col_index_from_letters(&s[..i])?;
    Some((row, col))
}

/// Fast path for Python's `int(s)`: plain optional '-' followed by ASCII
/// digits (no whitespace, '+', underscores ...), fitting in i64.
pub fn fast_int(s: &[u8]) -> Option<i64> {
    let (neg, d) = match s.first() {
        Some(b'-') => (true, &s[1..]),
        _ => (false, s),
    };
    if d.is_empty() || d.len() > 18 || !d.iter().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let mut v: i64 = 0;
    for &c in d {
        v = v * 10 + (c - b'0') as i64;
    }
    Some(if neg { -v } else { v })
}

/// Fast path for Python's `float(s)` on the common decimal syntax.
pub fn fast_float(s: &[u8]) -> Option<f64> {
    // [-]digits[.digits][(e|E)[+-]digits]  or  [-].digits...
    let mut i = 0;
    let n = s.len();
    if i < n && (s[i] == b'-' || s[i] == b'+') {
        i += 1;
    }
    let ds = i;
    while i < n && s[i].is_ascii_digit() {
        i += 1;
    }
    let mut ndig = i - ds;
    if i < n && s[i] == b'.' {
        i += 1;
        let fs = i;
        while i < n && s[i].is_ascii_digit() {
            i += 1;
        }
        ndig += i - fs;
    }
    if ndig == 0 {
        return None;
    }
    if i < n && (s[i] == b'e' || s[i] == b'E') {
        i += 1;
        if i < n && (s[i] == b'-' || s[i] == b'+') {
            i += 1;
        }
        let es = i;
        while i < n && s[i].is_ascii_digit() {
            i += 1;
        }
        if i == es {
            return None;
        }
    }
    if i != n {
        return None;
    }
    // SAFETY: ASCII checked above
    unsafe { std::str::from_utf8_unchecked(s) }
        .parse::<f64>()
        .ok()
}

/// Python's `str.isspace()` for a single character
pub fn py_isspace(c: char) -> bool {
    matches!(
        c,
        '\t' | '\n'
            | '\x0b'
            | '\x0c'
            | '\r'
            | '\x1c'
            | '\x1d'
            | '\x1e'
            | '\x1f'
            | ' '
            | '\u{85}'
            | '\u{a0}'
            | '\u{1680}'
            | '\u{2000}'
            ..='\u{200a}' | '\u{2028}' | '\u{2029}' | '\u{202f}' | '\u{205f}' | '\u{3000}'
    )
}

/// Python's `str.strip()`
pub fn py_strip(s: &str) -> &str {
    s.trim_matches(py_isspace)
}

/// Would Python's `float(s)` succeed?  None if non-ASCII (ask Python).
pub fn py_float_parses(s: &str) -> Option<bool> {
    if !s.is_ascii() {
        return None;
    }
    let t = py_strip(s).as_bytes();
    let mut i = 0;
    let n = t.len();
    if i < n && (t[i] == b'+' || t[i] == b'-') {
        i += 1;
    }
    let rest = &t[i..];
    let lower: Vec<u8> = rest.to_ascii_lowercase();
    if lower == b"inf" || lower == b"infinity" || lower == b"nan" {
        return Some(true);
    }
    // digitpart := digit (['_'] digit)*
    fn digitpart(b: &[u8], mut i: usize) -> Option<usize> {
        if i >= b.len() || !b[i].is_ascii_digit() {
            return None;
        }
        i += 1;
        loop {
            if i < b.len() && b[i].is_ascii_digit() {
                i += 1;
            } else if i + 1 < b.len() && b[i] == b'_' && b[i + 1].is_ascii_digit() {
                i += 2;
            } else {
                return Some(i);
            }
        }
    }
    let b = rest;
    let mut j = 0;
    let mut have_digits = false;
    if let Some(e) = digitpart(b, j) {
        j = e;
        have_digits = true;
    }
    if j < b.len() && b[j] == b'.' {
        j += 1;
        if let Some(e) = digitpart(b, j) {
            j = e;
            have_digits = true;
        }
    }
    if !have_digits {
        return Some(false);
    }
    if j < b.len() && (b[j] == b'e' || b[j] == b'E') {
        j += 1;
        if j < b.len() && (b[j] == b'+' || b[j] == b'-') {
            j += 1;
        }
        match digitpart(b, j) {
            Some(e) => j = e,
            None => return Some(false),
        }
    }
    Some(j == b.len())
}

/// Python's `"%.16g" % value` for a finite float.
pub fn format_g16(v: f64) -> String {
    if v == 0.0 {
        return if v.is_sign_negative() {
            "-0".into()
        } else {
            "0".into()
        };
    }
    // 16 significant digits in scientific notation
    let s = format!("{:.15e}", v);
    let (mant, exp) = s.split_once('e').unwrap();
    let exp: i32 = exp.parse().unwrap();
    if !(-4..16).contains(&exp) {
        let mut m = mant.to_string();
        if m.contains('.') {
            while m.ends_with('0') {
                m.pop();
            }
            if m.ends_with('.') {
                m.pop();
            }
        }
        let sign = if exp < 0 { '-' } else { '+' };
        let ea = exp.unsigned_abs();
        if ea < 10 {
            format!("{}e{}0{}", m, sign, ea)
        } else {
            format!("{}e{}{}", m, sign, ea)
        }
    } else {
        let prec = (15 - exp).max(0) as usize;
        let mut m = format!("{:.*}", prec, v);
        if m.contains('.') {
            while m.ends_with('0') {
                m.pop();
            }
            if m.ends_with('.') {
                m.pop();
            }
        }
        m
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn letters() {
        assert_eq!(column_letter(1).unwrap(), "A");
        assert_eq!(column_letter(26).unwrap(), "Z");
        assert_eq!(column_letter(27).unwrap(), "AA");
        assert_eq!(column_letter(702).unwrap(), "ZZ");
        assert_eq!(column_letter(703).unwrap(), "AAA");
        assert_eq!(column_letter(18278).unwrap(), "ZZZ");
        for i in 1..=18278 {
            assert_eq!(
                col_index_from_letters(column_letter(i).unwrap().as_bytes()),
                Some(i)
            );
        }
    }

    #[test]
    fn g16() {
        assert_eq!(format_g16(1.5), "1.5");
        assert_eq!(format_g16(1e20), "1e+20");
        assert_eq!(format_g16(0.1), "0.1");
        assert_eq!(format_g16(1e-5), "1e-05");
        assert_eq!(format_g16(123456789012345678.0), "1.234567890123457e+17");
        assert_eq!(format_g16(43832.12783564815), "43832.12783564815");
    }
}
