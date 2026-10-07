# Excel Consolidator for ERP / EBS Master-Data Preparation

**v1.0.0 — Public Portfolio Edition**

A portfolio project demonstrating Excel consolidation and QA automation for ERP / EBS master-data preparation. It is automation tooling for data preparation and quality controls, not a commercial ERP or MES product.

## Workflow

Source Excel files → recursive ingestion → schema validation → consolidation → duplicate detection → Description length QA → exception reporting → Master workbook

## Capabilities

- Recursive Excel discovery across nested folders
- Catalog-aware consolidation with source traceability
- Duplicate and missing `ITEM_NUMBER` checks
- Schema and empty-sheet QA
- Weighted Description length validation
- QA summary, exception detail, and processing-log generation
- Synthetic regression test suite and frozen reference baseline

## Engineering evidence

- Regression suite: **PASS — 25/25**
- Synthetic regression fixtures: **9**
- Frozen reference baseline: derived from **v1.0.0**

## Run

```powershell
node .\cli\Run_ExcelConsolidatorV1.mjs --root-folder "C:\Data\EBS\Input" --recursive --output "C:\Data\EBS\Output\Master_EBS_YYYYMMDD.xlsx"
node .\cli\Run_Migration_Length_QA.mjs "C:\Data\EBS\Input" "C:\Data\EBS\Output\Migration_Cell_Length_QA.xlsx"
node .\tests\recursive_consolidator_tests.mjs
```

See [architecture](docs/ARCHITECTURE.md) and [synthetic fixture notes](tests/README.md).

## Security and privacy

This public repository contains **synthetic test data only**. Production company data, customer data, credentials, ERP access logic, and private filesystem locations are not included. Examples use generic paths and synthetic identifiers.

## Portfolio positioning

This project demonstrates Industrial Engineering process automation, manufacturing data preparation, ERP/EBS migration support, data-quality controls, regression testing, and version-controlled automation tooling.

## License

License: Not yet selected.
