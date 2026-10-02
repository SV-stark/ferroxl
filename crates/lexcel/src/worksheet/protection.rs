//! Sheet protection (`openpyxl/worksheet/protection.rs`).

use crate::worksheet::password_hasher::hash_password;

/// Information about protection of various aspects of a sheet.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SheetProtection {
    /// Whether sheet content is protected.
    pub sheet: bool,
    /// Whether objects are protected.
    pub objects: bool,
    /// Whether scenarios are protected.
    pub scenarios: bool,
    /// Whether cell formatting is locked.
    pub format_cells: bool,
    /// Whether column formatting is locked.
    pub format_columns: bool,
    /// Whether row formatting is locked.
    pub format_rows: bool,
    /// Whether inserting columns is locked.
    pub insert_columns: bool,
    /// Whether inserting rows is locked.
    pub insert_rows: bool,
    /// Whether inserting hyperlinks is locked.
    pub insert_hyperlinks: bool,
    /// Whether deleting columns is locked.
    pub delete_columns: bool,
    /// Whether deleting rows is locked.
    pub delete_rows: bool,
    /// Whether selecting locked cells is locked.
    pub select_locked_cells: bool,
    /// Whether sorting is locked.
    pub sort: bool,
    /// Whether using autofilter is locked.
    pub auto_filter: bool,
    /// Whether editing pivot tables is locked.
    pub pivot_tables: bool,
    /// Whether selecting unlocked cells is locked.
    pub select_unlocked_cells: bool,
    /// The hashed password.
    password: String,
    /// Whether protection is active.
    pub enabled: bool,
}

impl SheetProtection {
    /// Build an unprotected sheet.
    pub fn new() -> Self {
        SheetProtection::default()
    }

    /// Set a password, hashing it unless `already_hashed`.
    pub fn set_password(&mut self, value: &str, already_hashed: bool) {
        self.password = if already_hashed {
            value.to_string()
        } else {
            hash_password(value)
        };
        self.enabled = true;
    }

    /// The stored password, hashed or not.
    pub fn password(&self) -> &str {
        &self.password
    }

    /// Turn protection on.
    pub fn enable(&mut self) {
        self.enabled = true;
    }

    /// Turn protection off.
    pub fn disable(&mut self) {
        self.enabled = false;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_are_unprotected() {
        let p = SheetProtection::new();
        assert!(!p.enabled);
        assert!(!p.sheet);
        assert_eq!(p.password(), "");
    }

    #[test]
    fn set_password_enables_and_hashes() {
        let mut p = SheetProtection::new();
        p.set_password("secret", false);
        assert!(p.enabled);
        assert_eq!(p.password(), hash_password("secret"));
        assert_ne!(p.password(), "secret");
    }

    #[test]
    fn already_hashed_is_kept_verbatim() {
        let mut p = SheetProtection::new();
        p.set_password("ABCD", true);
        assert_eq!(p.password(), "ABCD");
    }

    #[test]
    fn enable_disable_toggle() {
        let mut p = SheetProtection::new();
        p.enable();
        assert!(p.enabled);
        p.disable();
        assert!(!p.enabled);
    }
}
