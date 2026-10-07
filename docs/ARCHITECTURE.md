# Architecture

`Input → parser → normalized records → catalog consolidation → QA engine → output workbook`

1. **Input** recursively discovers `.xlsx` files while excluding temporary and prior-output workbooks.
2. **Parser** reads workbook sheets without modifying source files.
3. **Normalized records** retain original source columns and add file, relative path, sheet, row, and load-time traceability.
4. **Catalog consolidation** routes resolved catalog records to separate worksheets.
5. **QA engine** surfaces schema mismatch, missing and duplicate item numbers, cross-catalog duplicates, empty sheets, and Description length status.
6. **Output workbook** writes `SUMMARY`, `EXCEPTIONS`, `QA_SUMMARY`, `LENGTH_QA`, `PROCESSING_LOG`, and catalog worksheets.

## Design choices

- Source files are not modified.
- Duplicate records are retained and flagged.
- Exceptions are surfaced instead of silently removed.
- Description length uses weighted characters: Han characters count as three; other characters count as one.
- Recursive folder support preserves source-relative traceability.
- Synthetic regression fixtures make behavior reproducible.
