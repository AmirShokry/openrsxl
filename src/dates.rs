//! Exact port of `openpyxl.utils.datetime.from_excel` (non-timedelta case).
//!
//! Python semantics reproduced:
//! * `divmod(value, 1)` for floats (CPython's float_divmod),
//! * `round()` = round half to even,
//! * `datetime + timedelta` range checks (results outside year 1..9999 are
//!   errors, reported to the caller who emits openpyxl's warning).

/// Result of the conversion.
pub enum ExcelDate {
    DateTime {
        y: i32,
        m: u8,
        d: u8,
        h: u8,
        mi: u8,
        s: u8,
        us: u32,
    },
    Time {
        h: u8,
        mi: u8,
        s: u8,
        us: u32,
    },
    /// OverflowError / ValueError in Python
    Error,
}

/// CPython float_divmod(vx, 1.0)
fn py_divmod1(vx: f64) -> (f64, f64) {
    let wx = 1.0f64;
    let mut m = vx % wx;
    let mut div = (vx - m) / wx;
    if m != 0.0 {
        if (wx < 0.0) != (m < 0.0) {
            m += wx;
            div -= 1.0;
        }
    } else {
        m = 0.0f64.copysign(wx);
    }
    let floordiv = if div != 0.0 {
        let mut f = div.floor();
        if div - f > 0.5 {
            f += 1.0;
        }
        f
    } else {
        0.0f64.copysign(vx / wx)
    };
    (floordiv, m)
}

/// days since 0001-01-01 (proleptic Gregorian, day 0 = 0001-01-01)
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468 + 719162 // shift unix epoch based -> 0001-01-01 based
}

fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z - 719162 + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

const MAX_DAY: i64 = 3652058; // 9999-12-31 relative to 0001-01-01

#[derive(Clone)]
pub struct Epoch {
    pub days: i64,
    /// microseconds within the day
    pub us: i64,
    pub windows: bool,
}

impl Epoch {
    #[allow(clippy::too_many_arguments)] // the fields of a datetime
    pub fn new(y: i64, m: i64, d: i64, h: i64, mi: i64, s: i64, us: i64, windows: bool) -> Self {
        Epoch {
            days: days_from_civil(y, m, d),
            us: ((h * 60 + mi) * 60 + s) * 1_000_000 + us,
            windows,
        }
    }
}

fn finish(epoch: &Epoch, day: i64, ms: i64) -> ExcelDate {
    // epoch + timedelta(days=day) + timedelta(milliseconds=ms)
    // (two additions, each must stay within range)
    let total_us_day = |days: i64, us: i64| -> Option<(i64, i64)> {
        let extra = us.div_euclid(86_400_000_000);
        let us = us.rem_euclid(86_400_000_000);
        let days = days.checked_add(extra)?;
        if !(0..=MAX_DAY).contains(&days) {
            return None;
        }
        Some((days, us))
    };
    let Some(step1) = epoch.days.checked_add(day) else {
        return ExcelDate::Error;
    };
    let Some((d1, u1)) = total_us_day(step1, epoch.us) else {
        return ExcelDate::Error;
    };
    let Some((d2, u2)) = total_us_day(d1, u1 + ms * 1000) else {
        return ExcelDate::Error;
    };
    let (y, m, d) = civil_from_days(d2);
    let secs = u2 / 1_000_000;
    ExcelDate::DateTime {
        y: y as i32,
        m: m as u8,
        d: d as u8,
        h: (secs / 3600) as u8,
        mi: ((secs / 60) % 60) as u8,
        s: (secs % 60) as u8,
        us: (u2 % 1_000_000) as u32,
    }
}

fn time_of(ms: i64) -> ExcelDate {
    // days_to_time(timedelta(milliseconds=ms)) with ms < 1 day
    let secs = ms / 1000;
    ExcelDate::Time {
        h: (secs / 3600) as u8,
        mi: ((secs / 60) % 60) as u8,
        s: (secs % 60) as u8,
        us: ((ms % 1000) * 1000) as u32,
    }
}

/// from_excel for a Python int value
pub fn from_excel_int(value: i64, epoch: &Epoch) -> ExcelDate {
    // divmod(int, 1) == (value, 0); diff == 0
    if value == 0 {
        return time_of(0);
    }
    let mut day = value;
    if 0 < value && value < 60 && epoch.windows {
        day += 1;
    }
    if day.abs() > 999_999_999 {
        return ExcelDate::Error;
    }
    finish(epoch, day, 0)
}

/// from_excel for a Python float value
pub fn from_excel_float(value: f64, epoch: &Epoch) -> ExcelDate {
    if !value.is_finite() {
        return ExcelDate::Error;
    }
    let (mut day, fraction) = py_divmod1(value);
    let ms_f = (fraction * 86400.0 * 1000.0).round_ties_even();
    if !ms_f.is_finite() {
        return ExcelDate::Error;
    }
    let ms = ms_f as i64;
    // timedelta(milliseconds=ms).days == 0  <=>  ms < 86_400_000
    if (0.0..1.0).contains(&value) && ms < 86_400_000 {
        return time_of(ms);
    }
    if 0.0 < value && value < 60.0 && epoch.windows {
        day += 1.0;
    }
    if day.abs() > 999_999_999.0 {
        return ExcelDate::Error;
    }
    finish(epoch, day as i64, ms)
}

/// `openpyxl.utils.datetime.to_excel` for a value produced by from_excel
/// (naive datetime or time), with the given epoch.
pub fn to_excel(d: &ExcelDate, epoch: &Epoch) -> Option<f64> {
    // time_to_days: ((h*3600) + (m*60) + s + us/10**6) / 86400
    let time_to_days = |h: u8, mi: u8, s: u8, us: u32| -> f64 {
        let secs = (h as i64 * 3600 + mi as i64 * 60 + s as i64) as f64 + (us as f64 / 1e6);
        secs / 86400.0
    };
    match *d {
        ExcelDate::Time { h, mi, s, us } => Some(time_to_days(h, mi, s, us)),
        ExcelDate::DateTime {
            y,
            m,
            d,
            h,
            mi,
            s,
            us,
        } => {
            // (dt - epoch).days with floor semantics
            let dt_us = days_from_civil(y as i64, m as i64, d as i64) as i128 * 86_400_000_000
                + (((h as i64 * 60 + mi as i64) * 60 + s as i64) * 1_000_000 + us as i64) as i128;
            let ep_us = epoch.days as i128 * 86_400_000_000 + epoch.us as i128;
            let mut days = (dt_us - ep_us).div_euclid(86_400_000_000) as i64;
            if 0 < days && days <= 60 && epoch.windows {
                days -= 1;
            }
            Some(days as f64 + time_to_days(h, mi, s, us))
        }
        ExcelDate::Error => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn civil_roundtrip() {
        assert_eq!(days_from_civil(1, 1, 1), 0);
        assert_eq!(days_from_civil(9999, 12, 31), MAX_DAY);
        for z in (0..=MAX_DAY).step_by(997) {
            let (y, m, d) = civil_from_days(z);
            assert_eq!(days_from_civil(y, m, d), z);
        }
    }
}
