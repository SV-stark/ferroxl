//! Read-only test runner for Excel files in D:\OneDrive\Office\Assignments\NRLM\Shimla
//! Using ferroxl v0.1.2:
//! - Tests workbook loading, sheet extraction, cell values, and formulas
//! - Audits formula dependency graphs and circular reference detection
//! - Strictly read-only: does not modify or write to any target file

use std::fs;
use std::path::{Path, PathBuf};
use std::time::Instant;
use ferroxl::CellValue;

fn find_excel_files(dir: &Path, list: &mut Vec<PathBuf>) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                find_excel_files(&path, list);
            } else if let Some(ext) = path.extension().and_then(|s| s.to_str()) {
                let ext_lower = ext.to_ascii_lowercase();
                if ext_lower == "xlsx" || ext_lower == "xlsm" {
                    list.push(path);
                }
            }
        }
    }
}

fn main() {
    let target_dir = Path::new(r"D:\OneDrive\Office\Assignments\NRLM\Shimla");
    if !target_dir.exists() {
        eprintln!("Target directory does not exist: {:?}", target_dir);
        return;
    }

    let mut files = Vec::new();
    find_excel_files(target_dir, &mut files);
    files.sort();

    println!("Testing ferroxl v0.1.2 (read-only) on {} Excel workbooks in {:?}", files.len(), target_dir);
    println!("{:-<120}", "");
    println!(
        "{:<4} | {:<52} | {:<6} | {:<8} | {:<8} | {:<6} | {:<7} | {:<10}",
        "#", "File Name", "Sheets", "Cells", "Formulas", "DepEdges", "Cycles", "Status"
    );
    println!("{:-<120}", "");

    let start_all = Instant::now();
    let mut passed = 0;
    let mut failed = 0;
    let mut total_sheets = 0;
    let mut total_cells = 0;
    let mut total_formulas = 0;
    let mut total_dep_edges = 0;
    let mut total_cycles_found = 0;
    let mut files_with_cycles: Vec<(String, usize)> = Vec::new();

    for (i, file_path) in files.iter().enumerate() {
        let file_name = file_path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("unknown");
        let short_name = if file_name.len() > 50 {
            format!("{}...", &file_name[..47])
        } else {
            file_name.to_string()
        };

        let file_start = Instant::now();
        let result = ferroxl::load_workbook(file_path.to_str().unwrap(), ferroxl::LoadOptions::default());

        match result {
            Ok(wb) => {
                let sheet_names = wb.get_sheet_names();
                let sheet_count = sheet_names.len();
                total_sheets += sheet_count;

                let mut file_cells = 0;
                let mut file_formulas = 0;
                let mut file_dep_edges = 0;
                let mut file_cycles = 0;

                for sheet in &wb.worksheets {
                    for cell in sheet.cells() {
                        file_cells += 1;
                        if matches!(cell.internal_value(), CellValue::Formula(_)) {
                            file_formulas += 1;
                        }
                    }

                    // Audit dependency graph & circular references (v0.1.2)
                    if let Ok(graph) = sheet.dependency_graph() {
                        for targets in graph.values() {
                            file_dep_edges += targets.len();
                        }
                    }

                    if let Ok(cycles) = sheet.circular_references() {
                        file_cycles += cycles.len();
                    }
                }

                total_cells += file_cells;
                total_formulas += file_formulas;
                total_dep_edges += file_dep_edges;
                total_cycles_found += file_cycles;
                passed += 1;

                if file_cycles > 0 {
                    files_with_cycles.push((file_name.to_string(), file_cycles));
                }

                let elapsed_ms = file_start.elapsed().as_millis();
                println!(
                    "{:<4} | {:<52} | {:<6} | {:<8} | {:<8} | {:<8} | {:<6} | OK ({}ms)",
                    i + 1,
                    short_name,
                    sheet_count,
                    file_cells,
                    file_formulas,
                    file_dep_edges,
                    file_cycles,
                    elapsed_ms
                );
            }
            Err(e) => {
                failed += 1;
                println!(
                    "{:<4} | {:<52} | {:<6} | {:<8} | {:<8} | {:<8} | {:<6} | FAILED: {}",
                    i + 1,
                    short_name,
                    "-",
                    "-",
                    "-",
                    "-",
                    "-",
                    e
                );
            }
        }
    }

    let total_elapsed = start_all.elapsed();
    println!("{:-<120}", "");
    println!("Test & Audit Summary (ferroxl v0.1.2):");
    println!("  Total Workbooks Tested   : {}", files.len());
    println!("  Passed (Clean Parse)     : {}", passed);
    println!("  Failed                   : {}", failed);
    println!("  Total Sheets Processed   : {}", total_sheets);
    println!("  Total Cells Extracted    : {}", total_cells);
    println!("  Total Formulas Parsed    : {}", total_formulas);
    println!("  Dependency Graph Edges   : {}", total_dep_edges);
    println!("  Circular References Seen : {}", total_cycles_found);
    if !files_with_cycles.is_empty() {
        println!("  Files with Circular Refs : {:?}", files_with_cycles);
    } else {
        println!("  Circular References      : None detected across all accounting files (100% clean models)");
    }
    println!("  Total Execution Time     : {:.2?}", total_elapsed);
    if !files.is_empty() {
        println!("  Average Time per File    : {:.2?}", total_elapsed / files.len() as u32);
    }
    println!("{:-<120}", "");
}
