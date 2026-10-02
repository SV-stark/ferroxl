//! Excel date weirdness (`openpyxl/date_time.py`).
//!
//! Excel has two primary date tracking schemes:
//!
//! * Windows — day 1 is 1900-01-01
//! * Mac — day 1 is 1904-01-01
//!
//! Conversions go through Julian Day Numbers, matching the `jdcal` dependency used by
//! the Python original. `jdcal.gcal2jd` returns a triple whose sum is the fractional
//! Julian day; `jdcal.jd2gcal` inverts it.

use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use std::sync::atomic::{AtomicU64, Ordering};

use crate::exceptions::{Error, Result};

/// The Mac epoch, 1904-01-01.
pub const MAC_EPOCH: [u32; 3] = [1904, 1, 1];

/// The Windows epoch, 1899-12-30.
///
/// This is the day openpyxl treats as serial zero; Excel's own serial 1 is 1900-01-01, and
/// the two differ because Excel believes 1900 was a leap year.
pub const WINDOWS_EPOCH: [u32; 3] = [1899, 12, 30];

/// Julian day of the Mac epoch, 1904-01-01 (`MAC_EPOCH_DAY`).
pub const MAC_EPOCH_DAY: f64 = 2416480.5;

/// Julian day of the Windows epoch, 1899-12-30 (`WINDOWS_EPOCH_DAY`).
pub const WINDOWS_EPOCH_DAY: f64 = 2415018.5;

/// Serial offset for the 1900 date system (`CALENDAR_WINDOWS_1900`).
pub const CALENDAR_WINDOWS_1900: f64 = WINDOWS_EPOCH_DAY;

/// Serial offset for the 1904 date system (`CALENDAR_MAC_1904`).
pub const CALENDAR_MAC_1904: f64 = MAC_EPOCH_DAY;

/// Seconds in a day.
pub const SECS_PER_DAY: f64 = 86400.0;

/// `jdcal.MJD_0`.
pub const MJD_0: f64 = 2_400_000.5;

/// The date system a workbook uses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum BaseDate {
    /// Windows 1900 system (the default).
    #[default]
    Windows1900,
    /// Classic Mac 1904 system.
    Mac1904,
}

impl BaseDate {
    /// The serial offset this calendar uses.
    pub fn offset(self) -> f64 {
        match self {
            BaseDate::Windows1900 => CALENDAR_WINDOWS_1900,
            BaseDate::Mac1904 => CALENDAR_MAC_1904,
        }
    }

    /// Parse the `date1904` workbook property.
    pub fn from_date1904_flag(value: Option<&str>) -> BaseDate {
        match value {
            Some("1") | Some("true") => BaseDate::Mac1904,
            _ => BaseDate::Windows1900,
        }
    }

    /// `True` when this is the Windows 1900 calendar, which drives the leap-year bug fix.
    pub fn is_windows(self) -> bool {
        matches!(self, BaseDate::Windows1900)
    }
}

/// Convert a Gregorian date to the Julian day at midnight.
///
/// This is `jdcal.gcal2jd` as openpyxl uses it: `jdcal` returns `(MJD_0, mjd)` and openpyxl
/// sums the pair, so this returns the full Julian day. A date with no time component
/// therefore lands on a half-integer.
pub fn gcal2jd(year: i32, month: u32, day: u32) -> f64 {
    let year = i64::from(year);
    let month = i64::from(month);
    let day = i64::from(day);
    // January and February are moved to the end of the year so that the leap day comes
    // last, which keeps every month-length calculation uniform.
    let a = (month - 14) / 12;
    let mut mjd = (1461 * (year + 4800 + a)) / 4;
    mjd += (367 * (month - 2 - 12 * a)) / 12;
    let x = (year + 4900 + a) / 100;
    mjd -= (3 * x) / 4;
    // `jdcal` subtracts `2432075.5` to convert to a Julian day and then another half to
    // move from midday to midnight; adding `MJD_0` here reproduces the pair's sum.
    (mjd + day - 2432076 + MJD_0.floor() as i64) as f64 + 0.5
}

/// Convert a Julian day to a Gregorian `(year, month, day)`.
///
/// Mirrors `jdcal.jd2gcal` for the calendar argument openpyxl passes (`MJD_0`), which
/// selects the Gregorian calendar.
pub fn jd2gcal(jd: f64) -> (i32, u32, u32) {
    jd2gcal_int((jd + 0.5).floor() as i64)
}

/// Convert an integer Julian day number to a Gregorian date.
pub fn jd2gcal_int(jdn: i64) -> (i32, u32, u32) {
    if jdn > 2299160 {
        // Gregorian calendar.
        let a = jdn + 32044;
        let b = (4 * a + 3).div_euclid(146097);
        let c = a - (146097 * b).div_euclid(4);
        let d = (4 * c + 3).div_euclid(1461);
        let e = c - (1461 * d).div_euclid(4);
        let m = (5 * e + 2).div_euclid(153);
        let day = e - (153 * m + 2).div_euclid(5) + 1;
        let month = m + 3 - 12 * (m.div_euclid(10));
        let year = 100 * b + d - 4800 + m.div_euclid(10);
        (year as i32, month as u32, day as u32)
    } else {
        // Julian calendar.
        let c = jdn + 32082;
        let d = (4 * c + 3).div_euclid(1461);
        let e = c - (1461 * d).div_euclid(4);
        let m = (5 * e + 2).div_euclid(153);
        let day = e - (153 * m + 2).div_euclid(5) + 1;
        let month = m + 3 - 12 * (m.div_euclid(10));
        let year = d - 4800 + m.div_euclid(10);
        (year as i32, month as u32, day as u32)
    }
}

/// Serialise a datetime to an Excel serial number.
///
/// Mirrors `openpyxl.date_time.to_excel`: the serial is `julian_day - offset`, minus one
/// more for serials at or below 60 under the 1900 system (Excel's phantom 1900-02-29).
pub fn to_excel(dt: NaiveDateTime, base: BaseDate) -> f64 {
    let mut jul = gcal2jd(dt.year(), dt.month(), dt.day()) - base.offset();
    if jul <= 60.0 && base.is_windows() {
        jul -= 1.0;
    }
    jul + time_to_days_datetime(dt)
}

/// Serialise a plain date to an Excel serial number.
///
/// openpyxl treats a `date` as a `datetime` with no time part, so the result is the same
/// value [`to_excel`] produces for midnight.
pub fn date_to_excel(date: NaiveDate, base: BaseDate) -> f64 {
    to_excel(date.and_time(NaiveTime::MIN), base)
}

/// Deserialise an Excel serial number into a datetime.
///
/// Mirrors `openpyxl.date_time.from_excel`. Fractional serials with absolute value below
/// one are returned as a time-of-day instead.
pub fn from_excel(value: f64, base: BaseDate) -> ExcelDateTime {
    let (year, month, day) = jd2gcal(value + base.offset());
    let fractions = value - value.trunc();
    let micro = (fractions * SECS_PER_DAY * 1_000_000.0).round() as i64;
    // A serial strictly between -1 and 1 carries no date part. Zero is a real date
    // (the 1899-12-30 epoch), so the bounds are exclusive at both ends.
    if value > 0.0 && value < 1.0 || value < 0.0 && value > -1.0 {
        return ExcelDateTime::Time(days_to_time(micro));
    }
    // Guard against month/day normalisation for out-of-range results.
    let Some(date) = NaiveDate::from_ymd_opt(year, month, day.max(1)) else {
        return ExcelDateTime::DateTime(NaiveDateTime::default());
    };
    let midnight = date.and_time(NaiveTime::MIN);
    match midnight.checked_add_signed(chrono::Duration::microseconds(micro)) {
        Some(dt) => ExcelDateTime::DateTime(dt),
        None => ExcelDateTime::Date(date),
    }
}

/// The result of deserialising an Excel serial.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelDateTime {
    /// A datetime with a time component.
    DateTime(NaiveDateTime),
    /// A date with no time component.
    Date(NaiveDate),
    /// A time of day with no date component.
    Time(NaiveTime),
}

impl From<ExcelDateTime> for NaiveDateTime {
    fn from(value: ExcelDateTime) -> Self {
        match value {
            ExcelDateTime::DateTime(dt) => dt,
            ExcelDateTime::Date(d) => d.and_time(NaiveTime::MIN),
            ExcelDateTime::Time(t) => NaiveDate::MIN.and_time(t),
        }
    }
}

/// Convert a time value to a fraction of a day.
pub fn time_to_days(time: NaiveTime) -> f64 {
    let secs = (time.hour() * 3600 + time.minute() * 60 + time.second()) as f64
        + f64::from(time.nanosecond()) / 1_000_000_000.0;
    secs / SECS_PER_DAY
}

/// Fraction of a day contributed by the time part of a datetime.
pub fn time_to_days_datetime(dt: NaiveDateTime) -> f64 {
    time_to_days(dt.time())
}

/// Convert a duration to a fraction of a day.
pub fn timedelta_to_days(delta: Duration) -> f64 {
    delta.num_microseconds().unwrap_or(0) as f64 / 1_000_000.0 / SECS_PER_DAY
}

/// Convert a duration in microseconds to a fraction of a day.
pub fn micros_to_days(micros: i64) -> f64 {
    micros as f64 / 1_000_000.0 / SECS_PER_DAY
}

/// Convert a fractional day into a time of day.
pub fn days_to_time(micros: i64) -> NaiveTime {
    let total_secs = micros.div_euclid(1_000_000);
    let rem_micros = micros.rem_euclid(1_000_000);
    let minutes = total_secs.div_euclid(60);
    let seconds = total_secs.rem_euclid(60);
    let hours = minutes.div_euclid(60);
    let mins = minutes.rem_euclid(60);
    NaiveTime::from_hms_micro_opt(hours as u32, mins as u32, seconds as u32, rem_micros as u32)
        .unwrap_or(NaiveTime::MIN)
}

/// Convert a datetime to a W3CDTF timestamp string.
pub fn datetime_to_w3cdtf(dt: NaiveDateTime) -> String {
    dt.format("%Y-%m-%dT%H:%M:%SZ").to_string()
}

/// Parse a W3CDTF timestamp string.
///
/// Matches `openpyxl.date_time.W3CDTF_to_datetime`, which only looks at the first six
/// numeric groups and ignores fractional seconds entirely.
pub fn w3cdtf_to_datetime(formatted: &str) -> Result<NaiveDateTime> {
    let digits: Vec<&str> = formatted
        .split(|c: char| !c.is_ascii_digit())
        .filter(|s| !s.is_empty())
        .collect();
    if digits.len() < 6 {
        return Err(Error::Value(format!(
            "Invalid W3CDTF timestamp: {formatted}"
        )));
    }
    let parse = |s: &str| -> Result<i64> {
        s.parse::<i64>()
            .map_err(|_| Error::Value(format!("Invalid W3CDTF timestamp: {formatted}")))
    };
    let year = parse(digits[0])? as i32;
    let month = parse(digits[1])? as u32;
    let day = parse(digits[2])? as u32;
    let hour = parse(digits[3])? as u32;
    let minute = parse(digits[4])? as u32;
    let second = parse(digits[5])? as u32;
    NaiveDate::from_ymd_opt(year, month, day)
        .and_then(|d| d.and_hms_opt(hour, minute, second))
        .ok_or_else(|| Error::Value(format!("Invalid W3CDTF timestamp: {formatted}")))
}

/// Shared state for date conversion, mirroring `openpyxl.date_time.SharedDate`.
#[derive(Debug, Clone, Copy)]
pub struct SharedDate {
    /// Which calendar is in use.
    pub excel_base_date: BaseDate,
}

impl Default for SharedDate {
    fn default() -> Self {
        SharedDate {
            excel_base_date: BaseDate::Windows1900,
        }
    }
}

impl SharedDate {
    /// Build a `SharedDate`, rejecting unknown calendars.
    pub fn new(base: BaseDate) -> Result<Self> {
        Ok(SharedDate {
            excel_base_date: base,
        })
    }

    /// Convert a datetime to the Excel serial representation.
    pub fn datetime_to_julian(&self, dt: NaiveDateTime) -> f64 {
        to_excel(dt, self.excel_base_date)
    }

    /// Convert an Excel serial to a datetime.
    pub fn from_julian(&self, value: f64) -> ExcelDateTime {
        from_excel(value, self.excel_base_date)
    }

    /// Convert an hours/minutes/seconds triple to a fraction of a day.
    pub fn time_to_julian(&self, hours: i64, minutes: i64, seconds: i64) -> f64 {
        ((hours * 3600) + (minutes * 60) + seconds) as f64 / SECS_PER_DAY
    }
}

/// Monotonic counter backing the `lru_cache` behaviour of the Python module functions.
static CACHE_HITS: AtomicU64 = AtomicU64::new(0);

/// Number of cached date conversions served, exposed for parity with the cached Python
/// implementation (and useful for benchmarks).
pub fn cache_hit_count() -> u64 {
    CACHE_HITS.load(Ordering::Relaxed)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn w3cdtf_round_trip() {
        let dt = NaiveDate::from_ymd_opt(2013, 7, 15)
            .unwrap()
            .and_hms_opt(6, 52, 33)
            .unwrap();
        assert_eq!(datetime_to_w3cdtf(dt), "2013-07-15T06:52:33Z");
        assert_eq!(
            w3cdtf_to_datetime("2011-06-30T13:35:26Z").unwrap(),
            NaiveDate::from_ymd_opt(2011, 6, 30)
                .unwrap()
                .and_hms_opt(13, 35, 26)
                .unwrap()
        );
        // Fractional seconds are dropped, matching the Python regex-based parser.
        assert_eq!(
            w3cdtf_to_datetime("2013-03-04T12:19:01.00Z").unwrap(),
            NaiveDate::from_ymd_opt(2013, 3, 4)
                .unwrap()
                .and_hms_opt(12, 19, 1)
                .unwrap()
        );
    }

    #[test]
    fn to_excel_windows() {
        let cases: Vec<(NaiveDate, f64)> = vec![
            (NaiveDate::from_ymd_opt(1899, 12, 31).unwrap(), 0.0),
            (NaiveDate::from_ymd_opt(1900, 1, 15).unwrap(), 15.0),
            (NaiveDate::from_ymd_opt(1900, 2, 28).unwrap(), 59.0),
            (NaiveDate::from_ymd_opt(1900, 3, 1).unwrap(), 61.0),
            (NaiveDate::from_ymd_opt(2009, 12, 20).unwrap(), 40167.0),
        ];
        for (date, expected) in cases {
            assert_eq!(date_to_excel(date, BaseDate::Windows1900), expected);
        }
        let dt = NaiveDate::from_ymd_opt(2010, 1, 18)
            .unwrap()
            .and_hms_micro_opt(14, 15, 20, 1600)
            .unwrap();
        assert!((to_excel(dt, BaseDate::Windows1900) - 40196.5939815).abs() < 1e-9);
        let old = NaiveDate::from_ymd_opt(1506, 10, 15)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap();
        assert!((to_excel(old, BaseDate::Windows1900) - -143618.0).abs() < 1e-9);
    }

    #[test]
    fn to_excel_mac() {
        let cases: Vec<(NaiveDate, f64)> = vec![
            (NaiveDate::from_ymd_opt(1904, 1, 1).unwrap(), 0.0),
            (NaiveDate::from_ymd_opt(2011, 10, 31).unwrap(), 39385.0),
            (NaiveDate::from_ymd_opt(2009, 12, 20).unwrap(), 38705.0),
        ];
        for (date, expected) in cases {
            assert_eq!(date_to_excel(date, BaseDate::Mac1904), expected);
        }
    }

    #[test]
    fn from_excel_windows() {
        let cases: Vec<(f64, ExcelDateTime)> = vec![
            (
                40167.0,
                ExcelDateTime::DateTime(
                    NaiveDate::from_ymd_opt(2009, 12, 20)
                        .unwrap()
                        .and_hms_opt(0, 0, 0)
                        .unwrap(),
                ),
            ),
            (
                60.0,
                ExcelDateTime::DateTime(
                    NaiveDate::from_ymd_opt(1900, 2, 28)
                        .unwrap()
                        .and_hms_opt(0, 0, 0)
                        .unwrap(),
                ),
            ),
            (
                -25063.0,
                ExcelDateTime::DateTime(
                    NaiveDate::from_ymd_opt(1831, 5, 18)
                        .unwrap()
                        .and_hms_opt(0, 0, 0)
                        .unwrap(),
                ),
            ),
            (
                0.125,
                ExcelDateTime::Time(NaiveTime::from_hms_opt(3, 0, 0).unwrap()),
            ),
        ];
        for (value, expected) in cases {
            assert_eq!(from_excel(value, BaseDate::Windows1900), expected);
        }
        assert_eq!(
            from_excel(21980.0, BaseDate::Windows1900),
            ExcelDateTime::DateTime(
                NaiveDate::from_ymd_opt(1960, 3, 5)
                    .unwrap()
                    .and_hms_opt(0, 0, 0)
                    .unwrap()
            )
        );
    }

    #[test]
    fn from_excel_mac() {
        assert_eq!(
            from_excel(0.0, BaseDate::Mac1904),
            ExcelDateTime::DateTime(
                NaiveDate::from_ymd_opt(1904, 1, 1)
                    .unwrap()
                    .and_hms_opt(0, 0, 0)
                    .unwrap()
            )
        );
        assert_eq!(
            from_excel(21980.0, BaseDate::Mac1904),
            ExcelDateTime::DateTime(
                NaiveDate::from_ymd_opt(1964, 3, 6)
                    .unwrap()
                    .and_hms_opt(0, 0, 0)
                    .unwrap()
            )
        );
    }

    #[test]
    fn time_helpers() {
        let t1 = NaiveTime::from_hms_micro_opt(13, 55, 12, 36).unwrap();
        assert!((time_to_days(t1) - 0.5800000004166667).abs() < 1e-12);
        let t2 = NaiveTime::from_hms_opt(3, 0, 0).unwrap();
        assert!((time_to_days(t2) - 0.125).abs() < 1e-12);
        let td = Duration::days(1) + Duration::hours(3);
        assert!((timedelta_to_days(td) - 1.125).abs() < 1e-12);
        assert_eq!(
            days_to_time(51_320_001_600),
            NaiveTime::from_hms_micro_opt(14, 15, 20, 1600).unwrap()
        );
    }

    #[test]
    fn round_trip_dates() {
        // Serials above Excel's phantom 1900-02-29 round-trip exactly.
        for serial in [61.0, 40167.0, 40372.5] {
            let dt: NaiveDateTime = from_excel(serial, BaseDate::Windows1900).into();
            assert!((to_excel(dt, BaseDate::Windows1900) - serial).abs() < 1e-6);
        }
        // Below it the two functions disagree by one, because openpyxl skips 1900-02-29 on
        // the way out but not on the way back. Serial 1 reads as 1899-12-31, which writes
        // back as 0.
        let first: NaiveDateTime = from_excel(1.0, BaseDate::Windows1900).into();
        assert_eq!(
            NaiveDate::from_ymd_opt(1899, 12, 31).unwrap().and_time(dt_time()),
            first
        );
        assert_eq!(to_excel(first, BaseDate::Windows1900), 0.0);
        // The Mac calendar has no phantom day, so it round-trips from zero.
        for serial in [0.0, 39385.0] {
            let dt: NaiveDateTime = from_excel(serial, BaseDate::Mac1904).into();
            assert!((to_excel(dt, BaseDate::Mac1904) - serial).abs() < 1e-6);
        }
    }

    /// Midnight, as a `NaiveDateTime`'s time component.
    fn dt_time() -> chrono::NaiveTime {
        chrono::NaiveTime::MIN
    }
}
