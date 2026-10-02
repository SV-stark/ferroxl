//! Basic password hashing (`openpyxl/worksheet/password_hasher.py`).
//!
//! The algorithm is the one published by Daniel Rentz of OpenOffice and reused by the PEAR
//! package Spreadsheet_Excel_Writer by Xavier Noguer. It is a 15-bit rotating XOR, not a
//! cryptographic hash, and is reproduced here for file-format compatibility only.
//!
//! The accumulator is unbounded in Python, and for long passwords the shifted code points
//! genuinely exceed 64 bits (`hash_password("a" * 50)` is `20FFFFF4E38`), so the arithmetic
//! uses a small arbitrary-precision bit vector rather than a fixed-width integer.

/// An arbitrarily large unsigned integer stored as little-endian 64-bit limbs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct BigBits {
    limbs: Vec<u64>,
}

impl BigBits {
    /// Zero.
    fn zero() -> Self {
        BigBits { limbs: Vec::new() }
    }

    /// OR in `value << shift`, growing the accumulator as needed.
    fn or_shifted(mut self, value: u64, shift: u32) -> Self {
        if value == 0 {
            return self;
        }
        let limb_shift = (shift / 64) as usize;
        let bit_shift = shift % 64;
        let needed = limb_shift + 2;
        if self.limbs.len() < needed {
            self.limbs.resize(needed, 0);
        }
        let low = value << bit_shift;
        let high = if bit_shift == 0 {
            0
        } else {
            value >> (64 - bit_shift)
        };
        self.limbs[limb_shift] |= low;
        if high != 0 {
            self.limbs[limb_shift + 1] |= high;
        }
        self
    }

    /// OR another accumulator into this one, limb by limb.
    fn or_assign(&mut self, other: &BigBits) {
        if self.limbs.len() < other.limbs.len() {
            self.limbs.resize(other.limbs.len(), 0);
        }
        for (index, limb) in other.limbs.iter().enumerate() {
            self.limbs[index] |= limb;
        }
    }

    /// XOR another accumulator into this one, limb by limb.
    fn xor_assign(&mut self, other: &BigBits) {
        if self.limbs.len() < other.limbs.len() {
            self.limbs.resize(other.limbs.len(), 0);
        }
        for (index, limb) in other.limbs.iter().enumerate() {
            self.limbs[index] ^= limb;
        }
    }

    /// Clear every bit above `mask` (`value &= mask`).
    fn and_small(mut self, mask: u64) -> Self {
        if self.limbs.is_empty() {
            self.limbs.push(0);
        }
        self.limbs[0] &= mask;
        for limb in self.limbs.iter_mut().skip(1) {
            *limb = 0;
        }
        self
    }

    /// XOR a small value into the accumulator.
    fn xor_small(&mut self, value: u64) {
        if value == 0 {
            return;
        }
        if self.limbs.is_empty() {
            self.limbs.push(0);
        }
        self.limbs[0] ^= value;
    }

    /// Whether every limb is zero.
    fn is_zero(&self) -> bool {
        self.limbs.iter().all(|limb| *limb == 0)
    }

    /// Render as upper-case hex with no leading zeros, matching Python's `hex()`.
    fn to_hex(&self) -> String {
        if self.is_zero() {
            return "0".to_string();
        }
        let significant = self
            .limbs
            .iter()
            .rposition(|limb| *limb != 0)
            .expect("non-zero accumulator");
        let mut digits = Vec::new();
        for (position, limb) in self.limbs[..=significant].iter().enumerate().rev() {
            let is_top = position == significant;
            if is_top {
                digits.push(format!("{limb:X}"));
            } else {
                digits.push(format!("{limb:016X}"));
            }
        }
        digits.concat()
    }
}

/// Create a password hash from a given string.
///
/// Each character contributes its code point rotated into a 15-bit window; the XOR of all
/// windows is then combined with the password length and a fixed constant. The result is the
/// hex representation, upper-cased and without a leading `0x`.
pub fn hash_password(plaintext_password: &str) -> String {
    let mut password = BigBits::zero();
    for (shift, ch) in (1_u32..).zip(plaintext_password.chars()) {
        // `ord()` in Python is the code point.
        let code = ch as u64;
        // `value = code << shift`, `rotated_bits = value >> 15`, `value &= 0x7fff`.
        // The rotation is taken from the *unmasked* value, so it is computed first: when
        // `shift < 15` the bits are shifted right rather than dropped.
        let mut masked = BigBits::zero().or_shifted(code, shift).and_small(0x7fff);
        let rotated = if shift >= 15 {
            BigBits::zero().or_shifted(code, shift - 15)
        } else {
            BigBits::zero().or_shifted(code >> (15 - shift), 0)
        };
        masked.or_assign(&rotated);
        password.xor_assign(&masked);
    }
    password.xor_small(plaintext_password.chars().count() as u64);
    password.xor_small(0xCE4B);
    password.to_hex()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_python_reference_values() {
        for (password, expected) in [
            ("", "CE4B"),
            ("a", "CE88"),
            ("b", "CE8E"),
            ("password", "83AF"),
            ("test", "CBEB"),
            ("secret", "DAA7"),
            ("aa", "CF0F"),
            ("aaaaa", "C630"),
            ("a".repeat(10).as_str(), "C9FD"),
            ("a".repeat(20).as_str(), "B9DE"),
            ("a".repeat(50).as_str(), "20FFFFF4E38"),
        ] {
            assert_eq!(hash_password(password), expected, "password {password:?}");
        }
    }

    #[test]
    fn unbounded_precision_matches_python() {
        // 100 repeats exceeds 64 bits; a fixed-width accumulator would truncate here.
        assert_eq!(hash_password(&"a".repeat(100)), "83FFFFFFFFFFFFFFFFF4E6E");
    }

    #[test]
    fn hash_is_stable_and_upper_hex() {
        let h = hash_password("test");
        assert!(h.chars().all(|c| c.is_ascii_hexdigit()));
        assert!(h.chars().all(|c| !c.is_ascii_lowercase()));
        assert!(!h.starts_with("0x"));
    }

    #[test]
    fn different_passwords_differ() {
        assert_ne!(hash_password("a"), hash_password("b"));
        assert_ne!(hash_password("password"), hash_password("password "));
    }
}
