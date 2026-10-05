//! Exact port of `openpyxl.formula.tokenizer.Tokenizer` and
//! `openpyxl.formula.translate.Translator` (the parts used to expand shared
//! formulae while reading worksheets).
//!
//! Strings are handled as `Vec<char>` so that offsets reported in error
//! messages are code point offsets, exactly like Python's.

use crate::utils::{col_index_from_letters, column_letter, py_float_parses};

#[derive(Debug)]
pub enum FormulaError {
    Tokenizer(String),
    Translator(String),
    /// `IndexError: pop from empty list`
    PopEmpty,
    /// The input needs the Python implementation (eg. non-ASCII number check)
    NeedPython,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum TType {
    Literal,
    Operand,
    Func,
    Array,
    Paren,
    Sep,
    OpPre,
    OpIn,
    OpPost,
    WSpace,
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum SubType {
    None,
    Text,
    Number,
    Logical,
    Error,
    Range,
    Open,
    Close,
    Arg,
    Row,
}

#[derive(Clone, Debug)]
pub struct Token {
    pub value: String,
    pub ttype: TType,
    pub subtype: SubType,
}

const ERROR_CODES: [&str; 8] = [
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
];
const TOKEN_ENDERS: &str = ",;}) +-*/^&=><%";

fn make_operand(value: String) -> Result<Token, FormulaError> {
    let subtype = if value.starts_with('"') {
        SubType::Text
    } else if value.starts_with('#') {
        SubType::Error
    } else if value == "TRUE" || value == "FALSE" {
        SubType::Logical
    } else {
        match py_float_parses(&value) {
            Some(true) => SubType::Number,
            Some(false) => SubType::Range,
            None => return Err(FormulaError::NeedPython),
        }
    };
    Ok(Token {
        value,
        ttype: TType::Operand,
        subtype,
    })
}

fn make_subexp(value: &str, func: bool) -> Token {
    let ttype = if func {
        TType::Func
    } else if "{}".contains(value) {
        TType::Array
    } else if "()".contains(value) {
        TType::Paren
    } else {
        TType::Func
    };
    let subtype = if ")}".contains(value) {
        SubType::Close
    } else {
        SubType::Open
    };
    Token {
        value: value.to_string(),
        ttype,
        subtype,
    }
}

struct Tokenizer<'a> {
    formula: &'a [char],
    formula_str: &'a str,
    items: Vec<Token>,
    stack: Vec<Token>,
    offset: usize,
    token: Vec<String>,
}

fn sn_match(s: &str) -> bool {
    // ^[1-9](\.[0-9]+)?[Ee]$   ($ may also match before a trailing newline)
    let b = s.as_bytes();
    let b = if b.last() == Some(&b'\n') {
        &b[..b.len() - 1]
    } else {
        b
    };
    // also allow exact match without stripping
    fn m(b: &[u8]) -> bool {
        if b.len() < 2 {
            return false;
        }
        if !(b'1'..=b'9').contains(&b[0]) {
            return false;
        }
        let last = b[b.len() - 1];
        if last != b'E' && last != b'e' {
            return false;
        }
        let mid = &b[1..b.len() - 1];
        if mid.is_empty() {
            return true;
        }
        mid[0] == b'.' && mid.len() >= 2 && mid[1..].iter().all(|c| c.is_ascii_digit())
    }
    m(b) || m(s.as_bytes())
}

impl<'a> Tokenizer<'a> {
    fn err_unexpected(&self) -> FormulaError {
        FormulaError::Tokenizer(format!(
            "Unexpected character at position {} in '{}'",
            self.offset, self.formula_str
        ))
    }

    fn assert_empty_token(&self, can_follow: Option<&str>) -> Result<(), FormulaError> {
        if let Some(last) = self.token.last() {
            let ok = match can_follow {
                // `x in "<str>"` is a substring test
                Some(cf) => cf.contains(last.as_str()),
                None => false,
            };
            if !ok {
                return Err(self.err_unexpected());
            }
        }
        Ok(())
    }

    fn save_token(&mut self) -> Result<(), FormulaError> {
        if !self.token.is_empty() {
            let v: String = self.token.concat();
            self.token.clear();
            self.items.push(make_operand(v)?);
        }
        Ok(())
    }

    fn parse(&mut self) -> Result<(), FormulaError> {
        let f = self.formula;
        if f.is_empty() {
            return Ok(());
        } else if f[0] == '=' {
            self.offset += 1;
        } else {
            self.items.push(Token {
                value: self.formula_str.to_string(),
                ttype: TType::Literal,
                subtype: SubType::None,
            });
            return Ok(());
        }
        while self.offset < f.len() {
            if self.check_scientific_notation() {
                continue;
            }
            let c = f[self.offset];
            if TOKEN_ENDERS.contains(c) {
                self.save_token()?;
            }
            let consumed = match c {
                '"' | '\'' => Some(self.parse_string()?),
                '[' => Some(self.parse_brackets()?),
                '#' => Some(self.parse_error()?),
                ' ' | '\n' => Some(self.parse_whitespace()),
                '+' | '-' | '*' | '/' | '^' | '&' | '=' | '>' | '<' | '%' => {
                    Some(self.parse_operator())
                }
                '{' | '(' => Some(self.parse_opener()?),
                ')' | '}' => Some(self.parse_closer()?),
                ';' | ',' => Some(self.parse_separator()),
                _ => None,
            };
            match consumed {
                Some(n) => self.offset += n,
                None => {
                    self.token.push(c.to_string());
                    self.offset += 1;
                }
            }
        }
        self.save_token()
    }

    fn check_scientific_notation(&mut self) -> bool {
        let c = self.formula[self.offset];
        if (c == '+' || c == '-') && !self.token.is_empty() && sn_match(&self.token.concat()) {
            self.token.push(c.to_string());
            self.offset += 1;
            return true;
        }
        false
    }

    fn parse_string(&mut self) -> Result<usize, FormulaError> {
        self.assert_empty_token(Some(":"))?;
        let f = self.formula;
        let delim = f[self.offset];
        // emulate '"(?:[^"]*"")*[^"]*"(?!")'
        let mut i = self.offset + 1;
        let end = loop {
            match f[i..].iter().position(|&c| c == delim) {
                None => {
                    let subtype = if delim == '"' { "string" } else { "link" };
                    return Err(FormulaError::Tokenizer(format!(
                        "Reached end of formula while parsing {} in {}",
                        subtype, self.formula_str
                    )));
                }
                Some(p) => {
                    let j = i + p;
                    if j + 1 < f.len() && f[j + 1] == delim {
                        i = j + 2;
                    } else {
                        break j + 1;
                    }
                }
            }
        };
        let m: String = f[self.offset..end].iter().collect();
        let n = end - self.offset;
        if delim == '"' {
            self.items.push(make_operand(m)?);
        } else {
            self.token.push(m);
        }
        Ok(n)
    }

    fn parse_brackets(&mut self) -> Result<usize, FormulaError> {
        let f = self.formula;
        let mut count: i64 = 0;
        for (idx, &c) in f[self.offset..].iter().enumerate() {
            if c == '[' {
                count += 1;
            } else if c == ']' {
                count -= 1;
            } else {
                continue;
            }
            if count == 0 {
                let outer_right = idx + 1;
                self.token
                    .push(f[self.offset..self.offset + outer_right].iter().collect());
                return Ok(outer_right);
            }
        }
        Err(FormulaError::Tokenizer(format!(
            "Encountered unmatched '[' in {}",
            self.formula_str
        )))
    }

    fn parse_error(&mut self) -> Result<usize, FormulaError> {
        self.assert_empty_token(Some("!"))?;
        let sub = &self.formula[self.offset..];
        for err in ERROR_CODES.iter() {
            let ec: Vec<char> = err.chars().collect();
            if sub.len() >= ec.len() && sub[..ec.len()] == ec[..] {
                let v = format!("{}{}", self.token.concat(), err);
                self.items.push(make_operand(v)?);
                self.token.clear();
                return Ok(ec.len());
            }
        }
        Err(FormulaError::Tokenizer(format!(
            "Invalid error code at position {} in '{}'",
            self.offset, self.formula_str
        )))
    }

    fn parse_whitespace(&mut self) -> usize {
        let f = self.formula;
        self.items.push(Token {
            value: f[self.offset].to_string(),
            ttype: TType::WSpace,
            subtype: SubType::None,
        });
        let mut n = 0;
        while self.offset + n < f.len() && (f[self.offset + n] == ' ' || f[self.offset + n] == '\n')
        {
            n += 1;
        }
        n
    }

    fn parse_operator(&mut self) -> usize {
        let f = self.formula;
        if self.offset + 2 <= f.len() {
            let two: String = f[self.offset..self.offset + 2].iter().collect();
            if two == ">=" || two == "<=" || two == "<>" {
                self.items.push(Token {
                    value: two,
                    ttype: TType::OpIn,
                    subtype: SubType::None,
                });
                return 2;
            }
        }
        let c = f[self.offset];
        let token = if c == '%' {
            Token {
                value: "%".into(),
                ttype: TType::OpPost,
                subtype: SubType::None,
            }
        } else if "*/^&=><".contains(c) {
            Token {
                value: c.to_string(),
                ttype: TType::OpIn,
                subtype: SubType::None,
            }
        } else if self.items.is_empty() {
            Token {
                value: c.to_string(),
                ttype: TType::OpPre,
                subtype: SubType::None,
            }
        } else {
            let prev = self.items.iter().rev().find(|i| i.ttype != TType::WSpace);
            let is_infix = match prev {
                Some(p) => {
                    p.subtype == SubType::Close
                        || p.ttype == TType::OpPost
                        || p.ttype == TType::Operand
                }
                None => false,
            };
            Token {
                value: c.to_string(),
                ttype: if is_infix { TType::OpIn } else { TType::OpPre },
                subtype: SubType::None,
            }
        };
        self.items.push(token);
        1
    }

    fn parse_opener(&mut self) -> Result<usize, FormulaError> {
        let c = self.formula[self.offset];
        let token = if c == '{' {
            self.assert_empty_token(None)?;
            make_subexp("{", false)
        } else if !self.token.is_empty() {
            let v = format!("{}(", self.token.concat());
            self.token.clear();
            make_subexp(&v, false)
        } else {
            make_subexp("(", false)
        };
        self.items.push(token.clone());
        self.stack.push(token);
        Ok(1)
    }

    fn parse_closer(&mut self) -> Result<usize, FormulaError> {
        let top = self.stack.pop().ok_or(FormulaError::PopEmpty)?;
        let value = if top.ttype == TType::Array { "}" } else { ")" };
        let token = make_subexp(value, top.ttype == TType::Func);
        if !token.value.starts_with(self.formula[self.offset]) {
            return Err(FormulaError::Tokenizer(format!(
                "Mismatched ( and {{ pair in '{}'",
                self.formula_str
            )));
        }
        self.items.push(token);
        Ok(1)
    }

    fn parse_separator(&mut self) -> usize {
        let c = self.formula[self.offset];
        let token = if c == ';' {
            Token {
                value: ";".into(),
                ttype: TType::Sep,
                subtype: SubType::Row,
            }
        } else {
            match self.stack.last() {
                None => Token {
                    value: ",".into(),
                    ttype: TType::OpIn,
                    subtype: SubType::None,
                },
                Some(t) if t.ttype == TType::Paren => Token {
                    value: ",".into(),
                    ttype: TType::OpIn,
                    subtype: SubType::None,
                },
                Some(_) => Token {
                    value: ",".into(),
                    ttype: TType::Sep,
                    subtype: SubType::Arg,
                },
            }
        };
        self.items.push(token);
        1
    }
}

pub fn tokenize(formula: &str) -> Result<Vec<Token>, FormulaError> {
    let chars: Vec<char> = formula.chars().collect();
    let mut t = Tokenizer {
        formula: &chars,
        formula_str: formula,
        items: Vec::new(),
        stack: Vec::new(),
        offset: 0,
        token: Vec::new(),
    };
    t.parse()?;
    Ok(t.items)
}

/// A shared formula master, ready to be translated to other cells.
pub struct Translator {
    tokens: Vec<Token>,
    /// the formula pre-compiled into literal text and relative references
    template: Vec<Seg>,
    pub row: i64,
    pub col: i64,
}

/// Pre-compiled piece of a translated formula
enum Seg {
    Text(String),
    RelRow(i64),
    RelCol(i64),
}

fn push_text(out: &mut Vec<Seg>, s: &str) {
    if let Some(Seg::Text(t)) = out.last_mut() {
        t.push_str(s);
    } else {
        out.push(Seg::Text(s.to_string()));
    }
}

fn compile_row(out: &mut Vec<Seg>, row_str: &str) {
    if row_str.starts_with('$') {
        push_text(out, row_str);
    } else {
        out.push(Seg::RelRow(row_str.parse::<i64>().unwrap()));
    }
}

fn compile_col(out: &mut Vec<Seg>, col_str: &str) {
    if col_str.starts_with('$') {
        push_text(out, col_str);
    } else {
        // always 1-3 ASCII letters here, hence valid
        out.push(Seg::RelCol(
            col_index_from_letters(col_str.as_bytes()).unwrap(),
        ));
    }
}

/// Compile `Translator.translate_range` for a range token into segments.
fn compile_range(out: &mut Vec<Seg>, range_str: &str) {
    let (ws_part, range_str) = match range_str.rsplit_once('!') {
        Some((sheet, r)) => (format!("{}!", sheet), r),
        None => (String::new(), range_str),
    };
    if let Some((a, b)) = match_range(range_str, match_row_part) {
        push_text(out, &ws_part);
        compile_row(out, a);
        push_text(out, ":");
        compile_row(out, b);
        return;
    }
    if let Some((a, b)) = match_range(range_str, match_col_part) {
        push_text(out, &ws_part);
        compile_col(out, a);
        push_text(out, ":");
        compile_col(out, b);
        return;
    }
    if range_str.contains(':') {
        push_text(out, &ws_part);
        for (i, piece) in range_str.split(':').enumerate() {
            if i > 0 {
                push_text(out, ":");
            }
            compile_range(out, piece);
        }
        return;
    }
    match match_cell(range_str) {
        None => push_text(out, range_str),
        Some((c, r)) => {
            push_text(out, &ws_part);
            compile_col(out, c);
            compile_row(out, r);
        }
    }
}

fn out_of_range() -> FormulaError {
    FormulaError::Translator("Formula out of range".into())
}

/// Python's `re.match(...$)`: `$` matches at the end or before a final "\n".
fn strip_final_newline(s: &str) -> [&str; 2] {
    if let Some(stripped) = s.strip_suffix('\n') {
        [s, stripped]
    } else {
        [s, s]
    }
}

fn match_row_part(s: &str) -> bool {
    // \$?[1-9][0-9]{0,6}
    let b = s.strip_prefix('$').unwrap_or(s).as_bytes();
    !b.is_empty()
        && b.len() <= 7
        && (b'1'..=b'9').contains(&b[0])
        && b.iter().all(|c| c.is_ascii_digit())
}

fn match_col_part(s: &str) -> bool {
    let b = s.strip_prefix('$').unwrap_or(s).as_bytes();
    !b.is_empty() && b.len() <= 3 && b.iter().all(|c| c.is_ascii_alphabetic())
}

fn match_range(s: &str, part: fn(&str) -> bool) -> Option<(&str, &str)> {
    for cand in strip_final_newline(s) {
        if let Some((a, b)) = cand.split_once(':') {
            if part(a) && part(b) {
                return Some((a, b));
            }
        }
    }
    None
}

fn match_cell(s: &str) -> Option<(&str, &str)> {
    for cand in strip_final_newline(s) {
        let b = cand.as_bytes();
        let mut i = 0;
        if i < b.len() && b[i] == b'$' {
            i += 1;
        }
        let ls = i;
        while i < b.len() && b[i].is_ascii_alphabetic() {
            i += 1;
        }
        if i == ls || i - ls > 3 {
            continue;
        }
        let (col, row) = cand.split_at(i);
        if match_row_part(row) {
            return Some((col, row));
        }
    }
    None
}

impl Translator {
    pub fn new(formula: &str, row: i64, col: i64) -> Result<Self, FormulaError> {
        let tokens = tokenize(formula)?;
        let mut template = Vec::new();
        if !tokens.is_empty() && tokens[0].ttype != TType::Literal {
            push_text(&mut template, "=");
            for t in &tokens {
                if t.ttype == TType::Operand && t.subtype == SubType::Range {
                    compile_range(&mut template, &t.value);
                } else {
                    push_text(&mut template, &t.value);
                }
            }
        }
        Ok(Translator {
            tokens,
            template,
            row,
            col,
        })
    }

    fn translate_row(row_str: &str, rdelta: i64) -> Result<String, FormulaError> {
        if row_str.starts_with('$') {
            return Ok(row_str.to_string());
        }
        let new_row = row_str.parse::<i64>().unwrap() + rdelta;
        if new_row <= 0 {
            return Err(out_of_range());
        }
        Ok(new_row.to_string())
    }

    fn translate_col(col_str: &str, cdelta: i64) -> Result<String, FormulaError> {
        if col_str.starts_with('$') {
            return Ok(col_str.to_string());
        }
        let idx = col_index_from_letters(col_str.as_bytes()).ok_or_else(out_of_range)? + cdelta;
        column_letter(idx).ok_or_else(out_of_range)
    }

    fn translate_range(range_str: &str, rdelta: i64, cdelta: i64) -> Result<String, FormulaError> {
        let (ws_part, range_str) = match range_str.rsplit_once('!') {
            Some((sheet, r)) => (format!("{}!", sheet), r),
            None => (String::new(), range_str),
        };
        if let Some((a, b)) = match_range(range_str, match_row_part) {
            return Ok(format!(
                "{}{}:{}",
                ws_part,
                Self::translate_row(a, rdelta)?,
                Self::translate_row(b, rdelta)?
            ));
        }
        if let Some((a, b)) = match_range(range_str, match_col_part) {
            return Ok(format!(
                "{}{}:{}",
                ws_part,
                Self::translate_col(a, cdelta)?,
                Self::translate_col(b, cdelta)?
            ));
        }
        if range_str.contains(':') {
            let mut parts = Vec::new();
            for piece in range_str.split(':') {
                parts.push(Self::translate_range(piece, rdelta, cdelta)?);
            }
            return Ok(format!("{}{}", ws_part, parts.join(":")));
        }
        match match_cell(range_str) {
            None => Ok(range_str.to_string()),
            Some((c, r)) => Ok(format!(
                "{}{}{}",
                ws_part,
                Self::translate_col(c, cdelta)?,
                Self::translate_row(r, rdelta)?
            )),
        }
    }

    /// translate_formula(dest) where dest has already been converted to
    /// (row, col)
    pub fn translate(&self, dest: Option<(i64, i64)>) -> Result<String, FormulaError> {
        let mut out = String::with_capacity(64);
        self.translate_into(dest, &mut out)?;
        Ok(out)
    }

    /// `translate` appending to `out` (no allocation besides `out` growth).
    pub fn translate_into(
        &self,
        dest: Option<(i64, i64)>,
        out: &mut String,
    ) -> Result<(), FormulaError> {
        if self.tokens.is_empty() {
            return Ok(());
        }
        if self.tokens[0].ttype == TType::Literal {
            out.push_str(&self.tokens[0].value);
            return Ok(());
        }
        let (rd, cd) = match dest {
            Some((r, c)) => (r - self.row, c - self.col),
            None => (0, 0),
        };
        let mut tmp = [0u8; 24];
        for seg in &self.template {
            match seg {
                Seg::Text(t) => out.push_str(t),
                Seg::RelRow(r) => {
                    let n = r + rd;
                    if n <= 0 {
                        return Err(out_of_range());
                    }
                    out.push_str(crate::ftemplate::fmt_i64(n, &mut tmp));
                }
                Seg::RelCol(c) => {
                    let n = c + cd;
                    if !(1..=18278).contains(&n) {
                        return Err(out_of_range());
                    }
                    out.push_str(crate::ftemplate::col_letters(n, &mut tmp));
                }
            }
        }
        Ok(())
    }

    /// Position independent encoding of the compiled formula: two formulae
    /// with equal canonical forms differ only by their relative references
    /// (eg. `=A2+B2` in C2 and `=A3+B3` in C3). None for literals.
    pub fn canonical(&self) -> Option<Vec<u8>> {
        if self.tokens.is_empty() || self.tokens[0].ttype == TType::Literal {
            return None;
        }
        let mut out = Vec::with_capacity(64);
        for seg in &self.template {
            match seg {
                Seg::Text(t) => {
                    out.push(0);
                    out.extend_from_slice(&(t.len() as u32).to_le_bytes());
                    out.extend_from_slice(t.as_bytes());
                }
                Seg::RelRow(r) => {
                    out.push(1);
                    out.extend_from_slice(&(r - self.row).to_le_bytes());
                }
                Seg::RelCol(c) => {
                    out.push(2);
                    out.extend_from_slice(&(c - self.col).to_le_bytes());
                }
            }
        }
        Some(out)
    }

    /// Would `translate(dest)` succeed? (no allocation)
    pub fn check(&self, dest: Option<(i64, i64)>) -> Result<(), FormulaError> {
        if self.tokens.is_empty() || self.tokens[0].ttype == TType::Literal {
            return Ok(());
        }
        let (rd, cd) = match dest {
            Some((r, c)) => (r - self.row, c - self.col),
            None => (0, 0),
        };
        for seg in &self.template {
            match seg {
                Seg::Text(_) => {}
                Seg::RelRow(r) => {
                    if r + rd <= 0 {
                        return Err(out_of_range());
                    }
                }
                Seg::RelCol(c) => {
                    if !(1..=18278).contains(&(c + cd)) {
                        return Err(out_of_range());
                    }
                }
            }
        }
        Ok(())
    }

    /// Reference implementation (direct port, used by tests).
    #[allow(dead_code)]
    pub fn translate_slow(&self, dest: Option<(i64, i64)>) -> Result<String, FormulaError> {
        if self.tokens.is_empty() {
            return Ok(String::new());
        }
        if self.tokens[0].ttype == TType::Literal {
            return Ok(self.tokens[0].value.clone());
        }
        let (rd, cd) = match dest {
            Some((r, c)) => (r - self.row, c - self.col),
            None => (0, 0),
        };
        let mut out = String::with_capacity(64);
        out.push('=');
        for t in &self.tokens {
            if t.ttype == TType::Operand && t.subtype == SubType::Range {
                out.push_str(&Self::translate_range(&t.value, rd, cd)?);
            } else {
                out.push_str(&t.value);
            }
        }
        Ok(out)
    }
}
