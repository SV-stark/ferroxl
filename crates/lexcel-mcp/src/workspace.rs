//! Filesystem access for the tool handlers.
//!
//! Every path an agent supplies is resolved against a single root — the `--root` flag, or
//! `LEXCEL_ROOT`, or the process working directory. A path that escapes the root is
//! refused. This is not a security sandbox, but it does stop a mistyped `../` from
//! reaching somewhere the user never meant, and it gives the tools a documented boundary
//! an agent can reason about.

use std::path::{Component, Path, PathBuf};

use lexcel::{LoadOptions, Workbook};

/// A rooted view of the filesystem, with the helpers the tools need.
#[derive(Debug, Clone)]
pub struct Workspace {
    root: PathBuf,
}

impl Workspace {
    /// Create a workspace rooted at `root`, making the directory if it does not exist.
    pub fn new(root: impl Into<PathBuf>) -> std::io::Result<Self> {
        let root = root.into();
        std::fs::create_dir_all(&root)?;
        Ok(Workspace {
            root: normalise(&root),
        })
    }

    /// The directory every path is resolved against.
    pub fn root(&self) -> &Path {
        &self.root
    }

    /// Resolve a user-supplied path, refusing anything that leaves the root.
    pub fn resolve(&self, path: &str) -> Result<PathBuf, String> {
        if path.trim().is_empty() {
            return Err("path must not be empty".to_string());
        }
        let candidate = Path::new(path);
        let joined = if candidate.is_absolute() {
            candidate.to_path_buf()
        } else {
            self.root.join(candidate)
        };
        let normalised = normalise(&joined);
        if normalised == self.root || normalised.starts_with(&self.root) {
            Ok(normalised)
        } else {
            Err(format!(
                "path {path:?} is outside the workspace root {}",
                self.root.display()
            ))
        }
    }

    /// Resolve a path and check the suffix is one Excel understands.
    pub fn resolve_spreadsheet(&self, path: &str) -> Result<PathBuf, String> {
        let resolved = self.resolve(path)?;
        let extension = resolved
            .extension()
            .map(|e| e.to_string_lossy().to_ascii_lowercase())
            .unwrap_or_default();
        if !matches!(extension.as_str(), "xlsx" | "xlsm" | "xltx" | "xltm") {
            return Err(format!(
                "{path:?} must end in .xlsx or .xlsm, not {extension:?}"
            ));
        }
        Ok(resolved)
    }

    /// Read a workbook from disk.
    pub fn load(&self, path: &str, options: LoadOptions) -> Result<Workbook, String> {
        let resolved = self.resolve_spreadsheet(path)?;
        if !resolved.exists() {
            return Err(format!("no such workbook: {path}"));
        }
        lexcel::load_workbook(&resolved, options).map_err(|e| format!("could not read {path}: {e}"))
    }

    /// Write a workbook back to disk.
    pub fn save(&self, workbook: Workbook, path: &str) -> Result<(), String> {
        let resolved = self.resolve_spreadsheet(path)?;
        if let Some(parent) = resolved.parent() {
            std::fs::create_dir_all(parent)
                .map_err(|e| format!("could not create {}: {e}", parent.display()))?;
        }
        workbook
            .save(&resolved)
            .map_err(|e| format!("could not write {path}: {e}"))
    }

    /// Read a non-spreadsheet file, such as an image to embed.
    pub fn read_asset(&self, path: &str, limit: usize) -> Result<Vec<u8>, String> {
        let resolved = self.resolve(path)?;
        let data = std::fs::read(&resolved).map_err(|e| format!("could not read {path}: {e}"))?;
        if data.len() > limit {
            return Err(format!(
                "{path} is {} bytes, over the {limit}-byte limit",
                data.len()
            ));
        }
        Ok(data)
    }
}

/// Resolve `.` and `..` lexically, without touching the filesystem.
///
/// `canonicalize` would be simpler but it fails for paths that do not exist yet, and most
/// of the paths an agent hands over are about to be created.
fn normalise(path: &Path) -> PathBuf {
    let mut out = PathBuf::new();
    for component in path.components() {
        match component {
            Component::CurDir => {}
            Component::ParentDir => {
                // A leading `..` on a relative path has nowhere to go, so it is kept.
                if !out.pop() {
                    out.push("..");
                }
            }
            other => out.push(other.as_os_str()),
        }
    }
    if out.as_os_str().is_empty() {
        out.push(".");
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn workspace() -> Workspace {
        let dir = std::env::temp_dir().join("lexcel-mcp-workspace-tests");
        Workspace::new(dir).expect("workspace")
    }

    #[test]
    fn relative_paths_land_inside_the_root() {
        let workspace = workspace();
        let resolved = workspace.resolve("book.xlsx").unwrap();
        assert!(resolved.starts_with(workspace.root()));
        assert!(resolved.ends_with("book.xlsx"));
    }

    #[test]
    fn dot_segments_are_resolved_away() {
        let workspace = workspace();
        let resolved = workspace.resolve("nested/../book.xlsx").unwrap();
        assert_eq!(resolved.file_name().unwrap(), "book.xlsx");
    }

    #[test]
    fn escaping_the_root_is_refused() {
        let workspace = workspace();
        assert!(workspace.resolve("../escape.xlsx").is_err());
        assert!(workspace.resolve("a/../../escape.xlsx").is_err());
    }

    #[test]
    fn an_empty_path_is_refused() {
        assert!(workspace().resolve("   ").is_err());
    }

    #[test]
    fn only_spreadsheet_extensions_are_accepted() {
        let workspace = workspace();
        assert!(workspace.resolve_spreadsheet("book.xlsx").is_ok());
        assert!(workspace.resolve_spreadsheet("macro.xlsm").is_ok());
        assert!(workspace.resolve_spreadsheet("notes.txt").is_err());
        assert!(workspace.resolve_spreadsheet("noextension").is_err());
    }

    #[test]
    fn missing_workbooks_are_reported_by_name() {
        let workspace = workspace();
        let error = workspace.load("absent.xlsx", LoadOptions::default()).unwrap_err();
        assert!(error.contains("absent.xlsx"), "{error}");
    }

    #[test]
    fn a_workbook_round_trips_through_the_workspace() {
        let workspace = workspace();
        let mut workbook = Workbook::new();
        workbook.worksheets[0]
            .set("A1", lexcel::CellValue::text("hi"))
            .unwrap();
        workspace.save(workbook, "round-trip.xlsx").unwrap();
        let reloaded = workspace
            .load("round-trip.xlsx", LoadOptions::default())
            .unwrap();
        assert_eq!(
            reloaded.worksheets[0].cell_value("A1"),
            Some(lexcel::CellValue::text("hi"))
        );
    }

    #[test]
    fn oversized_assets_are_refused() {
        let workspace = workspace();
        workspace
            .save(Workbook::new(), "asset-holder.xlsx")
            .unwrap();
        let error = workspace.read_asset("asset-holder.xlsx", 4).unwrap_err();
        assert!(error.contains("limit"), "{error}");
    }
}
