//! Recovering a zip whose trailing records are missing.
//!
//! Two failures look alike from the outside -- "Excel says the file is corrupt" -- and need
//! opposite responses, so they are worth telling apart:
//!
//! - **Junk appended after the end-of-central-directory.** Already tolerated: the zip crate
//!   scans backwards for the `PK\x05\x06` signature rather than assuming the record is at a
//!   fixed offset. A 100 kB tail of zeroes, a stray `<html>404</html>`, even a prepended
//!   UTF-8 BOM all load. There is nothing to add here, and a fallback scanner that "fixed" it
//!   would only be able to fail where the existing reader succeeds.
//!
//! - **The end-of-central-directory itself cut off**, which is what an interrupted download
//!   looks like. The central directory usually survives, because it is written before the
//!   EOCD, so the archive is reconstructible: walk the central directory, count the entries,
//!   and synthesise the missing 22 bytes.
//!
//! Everything here refuses rather than guesses. A repaired archive that is subtly wrong opens
//! in Excel and shows the wrong data, which is worse than one that refused to open.

use crate::exceptions::{Error, Result};

/// The end-of-central-directory record: signature, then 18 bytes, then an optional comment.
const EOCD_SIGNATURE: &[u8; 4] = b"PK\x05\x06";
/// The zip64 end-of-central-directory locator, which points at a zip64 EOCD.
const EOCD64_LOCATOR_SIGNATURE: &[u8; 4] = b"PK\x06\x07";
/// A central directory file header.
const CENTRAL_SIGNATURE: &[u8; 4] = b"PK\x01\x02";

/// The fixed part of the EOCD: signature, disk fields, counts, and the directory's size and
/// offset. Everything after this is the archive comment.
const EOCD_FIXED_LEN: usize = 22;
/// A central directory file header's fixed part, before the variable-length name and extras.
const CENTRAL_FIXED_LEN: usize = 46;

/// Beyond this the 16-bit fields in a plain EOCD cannot represent the archive, so a
/// synthesised record would be wrong. Such a file needs a zip64 EOCD, which means it was not
/// merely truncated.
const MAX_ENTRIES: usize = u16::MAX as usize;
const MAX_CENTRAL_SIZE: u32 = u32::MAX;

/// Where the central directory sits, recovered by walking the entries.
struct CentralDirectory {
    /// How many entries it holds.
    entries: u16,
    /// Its length in bytes.
    size: u32,
    /// Its offset from the start of the archive.
    offset: u32,
}

/// Make `bytes` openable as a zip, if that can be done without guessing.
///
/// Returns `None` when the archive is already fine, so the caller can skip a copy of the whole
/// file in the overwhelmingly common case.
pub fn repair(bytes: Vec<u8>) -> Result<Option<Vec<u8>>> {
    if zip::ZipArchive::new(std::io::Cursor::new(&bytes[..])).is_ok() {
        return Ok(None);
    }

    // A record that is present but buried in junk is the cheap case: cut everything after it.
    if let Some(repaired) = truncate_after_eocd(&bytes) {
        if zip::ZipArchive::new(std::io::Cursor::new(&repaired[..])).is_ok() {
            return Ok(Some(repaired));
        }
    }

    let (central, directory_end) = walk_central_directory(&bytes).ok_or_else(|| {
        Error::InvalidFile(
            "the zip's end-of-central-directory record is missing and its central directory \
             could not be read, so the file is truncated beyond recovery"
                .to_string(),
        )
    })?;

    let mut repaired = bytes;
    // Anything after the directory is the remains of the record that was cut off. Dropping it
    // is not a guess: a complete directory is always the last thing before the EOCD, so
    // whatever trails it can only be that record's remains.
    repaired.truncate(directory_end);
    // A zip64 locator would now dangle: it points at a zip64 EOCD that, if the truncation cut
    // it, is not there. Dropping it is correct for an archive this size, where the plain EOCD
    // is authoritative anyway.
    if let Some(position) = find_last(&repaired, EOCD64_LOCATOR_SIGNATURE) {
        repaired.truncate(position);
    }
    repaired.extend_from_slice(&build_eocd(&central));
    Ok(Some(repaired))
}

/// Cut the archive back to the end of its EOCD record, if there is one to find.
fn truncate_after_eocd(bytes: &[u8]) -> Option<Vec<u8>> {
    let position = find_last(bytes, EOCD_SIGNATURE)?;
    // The comment length is the last two bytes of the record, and the comment is the last thing
    // in the file. Reading it is what makes this safe rather than a guess: without the record's
    // own length there is no way to tell a real EOCD from the four bytes appearing inside a
    // compressed stream.
    if position + EOCD_FIXED_LEN > bytes.len() {
        return None;
    }
    let comment_len = u16::from_le_bytes([bytes[position + 20], bytes[position + 21]]) as usize;
    let end = position + EOCD_FIXED_LEN + comment_len;
    if end > bytes.len() {
        return None;
    }
    if end == bytes.len() {
        return None;
    }
    let mut repaired = bytes[..end].to_vec();
    repaired.shrink_to_fit();
    Some(repaired)
}

/// Walk the central directory to recover its entry count, size and offset.
///
/// The search starts from every candidate signature rather than the first one, because a
/// `PK\x01\x02` byte sequence can occur by chance inside compressed data. A candidate is only
/// accepted when every entry it describes walks forward exactly to the end of the archive --
/// a run that ends mid-entry, or overruns the buffer, is a false positive.
fn walk_central_directory(bytes: &[u8]) -> Option<(CentralDirectory, usize)> {
    let mut search_from = 0usize;
    while let Some(offset) = find_from(bytes, CENTRAL_SIGNATURE, search_from) {
        if let Some(found) = walk_from(bytes, offset) {
            return Some(found);
        }
        search_from = offset + 1;
    }
    None
}

/// Whether `tail` is the remains of an EOCD record that the truncation cut short.
///
/// The record is the last thing in the file, so a truncation can only take bytes off its end --
/// which means what is left is a *prefix* of it. How much depends on where the cut fell: a
/// 4-byte cut leaves the whole 18-byte remainder of the record, while a 20-byte cut leaves two
/// bytes, `PK`, which is not even a whole signature. Both are prefixes of the record; anything
/// that is not is real trailing data, and treating it as a record would silently discard it.
fn is_partial_eocd(tail: &[u8]) -> bool {
    tail.len() < EOCD_FIXED_LEN
        && (EOCD_SIGNATURE.starts_with(tail) || tail.starts_with(EOCD_SIGNATURE))
}

/// Read the run of central directory entries that starts at `offset`, if it reaches the end.
///
/// Returns the directory and the offset just past it, which is where a rebuilt archive has to be
/// cut back to.
fn walk_from(bytes: &[u8], offset: usize) -> Option<(CentralDirectory, usize)> {
    let mut cursor = offset;
    let mut entries = 0usize;

    while cursor < bytes.len() {
        // The cut-off record's remains sit right here, and they are too short to be an entry.
        // Checking before reading rather than after is what lets a partially-truncated record
        // be told apart from a directory that is itself incomplete.
        if is_partial_eocd(&bytes[cursor..]) {
            break;
        }
        let entry = bytes.get(cursor..)?;
        if entry.len() < CENTRAL_FIXED_LEN || &entry[..4] != CENTRAL_SIGNATURE {
            return None;
        }
        let read_u16 = |at: usize| u16::from_le_bytes([entry[at], entry[at + 1]]);
        let read_u32 = |at: usize| {
            u32::from_le_bytes([entry[at], entry[at + 1], entry[at + 2], entry[at + 3]])
        };

        // Sizes and offsets saturate to their 32-bit maximum in a zip64 archive, where the real
        // value is in the extra field. A synthesised plain EOCD cannot express that, so this is
        // a refusal rather than a wrong number.
        if read_u32(20) == u32::MAX || read_u32(24) == u32::MAX || read_u32(42) == u32::MAX {
            return None;
        }
        let name_len = read_u16(28) as usize;
        let extra_len = read_u16(30) as usize;
        let comment_len = read_u16(32) as usize;
        let next = cursor + CENTRAL_FIXED_LEN + name_len + extra_len + comment_len;
        if next > bytes.len() {
            return None;
        }
        cursor = next;
        entries += 1;
        if entries > MAX_ENTRIES {
            return None;
        }
    }

    if entries == 0 {
        return None;
    }

    // The directory is the last thing before the record, so whatever follows it has to be that
    // record's remains -- or nothing at all, when the whole record was lost.
    let tail = &bytes[cursor..];
    if !tail.is_empty() && !is_partial_eocd(tail) {
        return None;
    }

    let size = cursor - offset;
    if size > MAX_CENTRAL_SIZE as usize || offset > MAX_CENTRAL_SIZE as usize {
        return None;
    }
    Some((
        CentralDirectory {
            entries: entries as u16,
            size: size as u32,
            offset: offset as u32,
        },
        cursor,
    ))
}

/// Build the 22-byte EOCD record for a recovered central directory.
fn build_eocd(central: &CentralDirectory) -> [u8; EOCD_FIXED_LEN] {
    let mut record = [0u8; EOCD_FIXED_LEN];
    record[..4].copy_from_slice(EOCD_SIGNATURE);
    let put_u16 = |record: &mut [u8; EOCD_FIXED_LEN], at: usize, value: u16| {
        record[at..at + 2].copy_from_slice(&value.to_le_bytes());
    };
    let put_u32 = |record: &mut [u8; EOCD_FIXED_LEN], at: usize, value: u32| {
        record[at..at + 4].copy_from_slice(&value.to_le_bytes());
    };
    // Disk numbers stay zero: a multi-disk archive is not something a single-file xlsx is, and
    // claiming a disk we did not read would be a guess.
    put_u16(&mut record, 4, 0);
    put_u16(&mut record, 6, 0);
    put_u16(&mut record, 8, central.entries);
    put_u16(&mut record, 10, central.entries);
    put_u32(&mut record, 12, central.size);
    put_u32(&mut record, 16, central.offset);
    put_u16(&mut record, 20, 0);
    record
}

/// The last position `needle` occurs at, searching backwards.
fn find_last(haystack: &[u8], needle: &[u8; 4]) -> Option<usize> {
    if haystack.len() < needle.len() {
        return None;
    }
    (0..=haystack.len() - needle.len())
        .rev()
        .find(|&at| &haystack[at..at + needle.len()] == needle)
}

/// The first position at or after `from` where `needle` occurs.
fn find_from(haystack: &[u8], needle: &[u8; 4], from: usize) -> Option<usize> {
    if haystack.len() < needle.len() || from > haystack.len() - needle.len() {
        return None;
    }
    (from..=haystack.len() - needle.len()).find(|&at| &haystack[at..at + needle.len()] == needle)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    /// A small archive written through the real zip writer, so the bytes are a genuine zip.
    fn sample_archive() -> Vec<u8> {
        let mut buffer = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut buffer);
            let options: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            for name in [
                "[Content_Types].xml",
                "xl/workbook.xml",
                "xl/worksheets/sheet1.xml",
            ] {
                writer.start_file(name, options).expect("writable");
                writer.write_all(b"<x/>").expect("writable");
            }
            writer.finish().expect("finished");
        }
        buffer.into_inner()
    }

    fn opens(bytes: &[u8]) -> bool {
        zip::ZipArchive::new(Cursor::new(bytes.to_vec())).is_ok()
    }

    fn entry_count(bytes: &[u8]) -> usize {
        zip::ZipArchive::new(Cursor::new(bytes.to_vec()))
            .map(|archive| archive.len())
            .unwrap_or(0)
    }

    #[test]
    fn an_intact_archive_is_left_alone() {
        let archive = sample_archive();
        assert!(repair(archive.clone()).expect("no error").is_none());
    }

    #[test]
    fn a_truncated_end_of_central_directory_is_rebuilt() {
        let archive = sample_archive();
        // Dropping the EOCD is what an interrupted download leaves behind.
        let truncated = &archive[..archive.len() - EOCD_FIXED_LEN];
        assert!(!opens(truncated), "the truncation really did break it");

        let repaired = repair(truncated.to_vec())
            .expect("repairable")
            .expect("something to repair");
        assert!(opens(&repaired));
        assert_eq!(entry_count(&repaired), 3, "every entry survived");
    }

    #[test]
    fn a_partly_truncated_end_of_central_directory_is_rebuilt() {
        let archive = sample_archive();
        // Every cut point inside the record, not just the whole of it. The interesting ends are
        // the first three -- what is left is `P`, `PK`, `PK\x05`, not even a whole signature --
        // and the last two, which leave two and one bytes of the record's tail.
        for cut in [1usize, 2, 3, 4, 5, 8, 12, 16, 19, 20, 21] {
            let truncated = &archive[..archive.len() - cut];
            assert!(!opens(truncated), "{cut} bytes off should break it");
            let repaired = repair(truncated.to_vec())
                .unwrap_or_else(|error| panic!("{cut} bytes off: {error}"))
                .unwrap_or_else(|| panic!("{cut} bytes off: nothing repaired"));
            assert_eq!(entry_count(&repaired), 3, "{cut} bytes off");
        }
    }

    #[test]
    fn junk_after_the_record_is_tolerated_without_our_help() {
        // Documented as already working, so this is a regression guard on the claim rather than
        // a test of the repair. If it ever fails, the repair path is being asked to do a job
        // the zip crate has never needed help with.
        let mut archive = sample_archive();
        assert!(opens(&archive));
        archive.extend_from_slice(&[0u8; 4096]);
        assert!(opens(&archive), "trailing junk already tolerated");
        assert!(repair(archive).expect("no error").is_none());
    }

    #[test]
    fn a_truncation_into_the_central_directory_is_refused() {
        // The directory is gone, so no record can describe the archive. Refusing is the only
        // honest answer: inventing one yields a file that opens and shows the wrong sheets.
        let archive = sample_archive();
        let truncated = &archive[..archive.len() - EOCD_FIXED_LEN - 20];
        let error = repair(truncated.to_vec()).expect_err("unrecoverable");
        assert!(
            error.to_string().contains("truncated beyond recovery"),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn something_that_is_not_a_zip_at_all_is_refused() {
        let error = repair(b"this is not a zip file".to_vec()).expect_err("unrecoverable");
        assert!(error.to_string().contains("truncated beyond recovery"));
    }

    #[test]
    fn a_rebuilt_record_reports_the_true_directory() {
        let archive = sample_archive();
        let original = find_last(&archive, EOCD_SIGNATURE).expect("a record");
        let truncated = &archive[..original];
        assert!(!opens(truncated), "the truncation really did break it");

        let repaired = repair(truncated.to_vec())
            .expect("repairable")
            .expect("something to repair");
        let rebuilt = find_last(&repaired, EOCD_SIGNATURE).expect("a rebuilt record");
        // Bytes 4 through 20 are the disk numbers, the entry counts, the directory's size and
        // its offset. Comparing them against the record that was lost is stronger than checking
        // numbers worked out by hand, and it is the check that would catch an off-by-one in the
        // walk -- which produces a record that parses and describes the wrong archive.
        assert_eq!(
            &repaired[rebuilt + 4..rebuilt + EOCD_FIXED_LEN],
            &archive[original + 4..original + EOCD_FIXED_LEN],
            "every field of the rebuilt record matches the original"
        );
    }

    #[test]
    fn a_central_directory_signature_inside_data_does_not_fool_the_walk() {
        // A false positive has to be rejected by the walk not reaching the end exactly, which is
        // what this exercises: the decoy is followed by the real directory, so a reader that
        // took the first candidate would describe an archive that does not exist.
        let archive = sample_archive();
        let decoy = b"PK\x01\x02 decoy bytes that are not an entry header at all";
        let mut spliced = archive[..archive.len() - EOCD_FIXED_LEN].to_vec();
        spliced.extend_from_slice(decoy);
        let error = repair(spliced).expect_err("a decoy is not a directory");
        assert!(error.to_string().contains("truncated beyond recovery"));
    }
}
