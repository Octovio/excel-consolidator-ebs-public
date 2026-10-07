import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import { evaluateDescriptionLength } from "./lengthRules.mjs";

const TRACKING = ["SOURCE_FILE", "SOURCE_FOLDER", "SOURCE_RELATIVE_PATH", "SOURCE_SHEET", "SOURCE_ROW", "LOAD_TIMESTAMP"];
const GENERIC_SHEETS = new Set(["sheet", "sheet1", "data", "data1", "input", "inputs", "raw", "rawdata"]);
const EXCEPTION_HEADERS = ["CATALOG", "ITEM_NUMBER", "SOURCE_FILE", "SOURCE_SHEET", "EXCEPTION_TYPE", "EXCEPTION_DETAIL"];
const SUMMARY_HEADERS = ["CATALOG", "TOTAL_RECORDS", "UNIQUE_ITEM_NUMBER", "DUPLICATE_ITEM_NUMBER", "MISSING_ITEM_NUMBER", "SOURCE_FILE_COUNT", "EXCEPTION_COUNT"];
const OUTPUT_NAME_PATTERN = /(?:consolidated|consolidator|merged_output|master_ebs|migration_cell_length_qa)/i;
export const FORMAL_MASTER_OUTPUT_FOLDER = process.env.EBS_OUTPUT_FOLDER || "C:\\Data\\EBS\\Output";
const norm = (v) => String(v ?? "").trim();
const normHeader = (v) => norm(v).toUpperCase();
const safeCatalog = (v) => /^[A-Za-z0-9][A-Za-z0-9 _/\-]{0,80}$/.test(norm(v)) ? norm(v) : null;
const colName = (n) => { let s = ""; while (n >= 0) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } return s; };

function dateStamp(date = new Date()) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}`;
}

export function formalMasterOutputPath(date = new Date(), outputFolder = FORMAL_MASTER_OUTPUT_FOLDER) {
  return path.join(outputFolder, `Master_EBS_${dateStamp(date)}.xlsx`);
}

function normalizeRelativePath(relativePath) {
  return relativePath.split(path.sep).join("\\");
}

function isExcludedExcelName(fileName, outputBase) {
  const lower = fileName.toLowerCase();
  return fileName.startsWith("~$")
    || lower === outputBase
    || OUTPUT_NAME_PATTERN.test(fileName);
}

export async function scanExcelFilesRecursive(rootFolder, outputFile = "") {
  const root = path.resolve(rootFolder);
  const outputBase = path.basename(outputFile).toLowerCase();
  const files = [], excluded = [];
  async function visit(folder) {
    const entries = await fs.readdir(folder, { withFileTypes: true });
    for (const entry of entries) {
      const absolutePath = path.join(folder, entry.name);
      if (entry.isDirectory()) {
        await visit(absolutePath);
        continue;
      }
      if (!entry.isFile() || !entry.name.toLowerCase().endsWith(".xlsx")) continue;
      const relativePath = normalizeRelativePath(path.relative(root, absolutePath));
      if (isExcludedExcelName(entry.name, outputBase)) {
        excluded.push({ file: entry.name, relativePath, reason: "Excluded temporary or Consolidator output workbook." });
        continue;
      }
      const parts = relativePath.split("\\");
      files.push({
        name: entry.name,
        absolutePath,
        relativePath,
        sourceFolder: parts.length > 1 ? parts[0] : path.basename(root),
      });
    }
  }
  await visit(root);
  files.sort((a, b) => a.relativePath.localeCompare(b.relativePath));
  excluded.sort((a, b) => a.relativePath.localeCompare(b.relativePath));
  return { root, files, excluded };
}

function duplicateFileCandidates(files) {
  const groups = new Map();
  for (const file of files) {
    const key = file.name.toLocaleLowerCase();
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(file);
  }
  return [...groups.values()].filter((group) => group.length > 1);
}

function makeWorksheetAliases(catalogs) {
  const used = new Set(["summary", "exceptions"]);
  const aliases = new Map();
  for (const catalog of [...catalogs].sort()) {
    // Excel tab constraints are an output concern only; CATALOG itself is never changed.
    let base = norm(catalog).replace(/[\\/:?*\[\]\x00-\x1F]/g, "_").trim() || "CATALOG";
    base = base.slice(0, 31);
    let alias = base, number = 2;
    while (used.has(alias.toLowerCase())) {
      const suffix = `_${number++}`;
      alias = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    }
    used.add(alias.toLowerCase());
    aliases.set(catalog, alias);
  }
  return aliases;
}

function usage() {
  console.log("Usage: node ExcelConsolidatorV1.mjs <source-folder> [output-file]");
}
function addException(exceptions, catalog, item, file, sheet, type, detail) {
  exceptions.push([catalog || "", item || "", file || "", sheet || "", type, detail]);
}
function catalogFrom(sheetName, fileStem, headers, row, unlabeledCatalog = null) {
  const catIndex = headers.findIndex((h) => norm(h) === "CATALOG");
  if (catIndex >= 0) {
    const v = norm(row[catIndex]);
    return v ? { catalog: v, method: "CATALOG column" } : { catalog: null, method: "CATALOG column is blank" };
  }
  if (unlabeledCatalog) return { catalog: unlabeledCatalog, method: "constant unlabeled column" };
  const sheet = safeCatalog(sheetName);
  if (sheet && !GENERIC_SHEETS.has(sheet.toLowerCase().replace(/\s+/g, ""))) return { catalog: sheet, method: "Sheet Name" };
  const file = safeCatalog(fileStem);
  if (file) return { catalog: file, method: "File Name" };
  return { catalog: null, method: "no safe CATALOG, Sheet Name, or File Name" };
}
function compareSchema(baseline, candidate) {
  const a = new Set(baseline), b = new Set(candidate);
  const missing = baseline.filter((h) => !b.has(h));
  const extra = candidate.filter((h) => !a.has(h));
  const caseOnly = baseline.some((h) => !candidate.includes(h) && candidate.some((x) => normHeader(x) === normHeader(h)));
  const order = !missing.length && !extra.length && baseline.some((h, i) => candidate[i] !== h);
  if (!missing.length && !extra.length && !caseOnly && !order) return "";
  const bits = [];
  if (missing.length) bits.push(`Missing columns: ${missing.join(", ")}`);
  if (extra.length) bits.push(`Extra columns: ${extra.join(", ")}`);
  if (caseOnly) bits.push("Header name/case inconsistency");
  if (order) bits.push("Header order inconsistency");
  return bits.join("; ");
}
function styleSheet(sheet, headers, rows, isSummary = false) {
  sheet.showGridLines = false;
  if (!headers.length) return;
  const end = colName(headers.length - 1);
  sheet.getRange(`A1:${end}1`).format = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", wrapText: true, borders: { preset: "outside", style: "thin", color: "#17365D" } };
  sheet.getRange(`A1:${end}1`).format.rowHeight = 28;
  sheet.freezePanes.freezeRows(1);
  if (rows.length) {
    sheet.getRange(`A2:${end}${rows.length + 1}`).format.borders = { preset: "insideHorizontal", style: "thin", color: "#E7E6E6" };
    sheet.tables.add(`A1:${end}${rows.length + 1}`, true, `${sheet.name.replace(/[^A-Za-z0-9]/g, "_").slice(0, 120)}_Table`);
  }
  // Apply after table creation, which can reset column widths in some Excel clients.
  for (let i = 0; i < headers.length; i++) {
    const width = Math.min(42, Math.max(13, String(headers[i]).length + 3));
    sheet.getRange(`${colName(i)}1:${colName(i)}${Math.max(rows.length + 1, 1)}`).format.columnWidth = width;
  }
  if (isSummary && rows.length) sheet.getRange(`B2:${end}${rows.length + 1}`).format.numberFormat = "#,##0";
}

export async function consolidate(sourceDir, outputFile) {
  const resolvedOutput = path.resolve(outputFile || formalMasterOutputPath());
  const scan = await scanExcelFilesRecursive(sourceDir, resolvedOutput);
  const { root: resolvedSource, files, excluded } = scan;
  const exceptions = [], records = [], schemaBase = new Map();
  const processingLog = excluded.map((entry) => ({
    EVENT_TYPE: "EXCLUDED_FILE", SKIPPED_FILE: entry.relativePath, SKIPPED_SHEET: "", SKIP_REASON: entry.reason, DETAIL: "",
  }));
  const duplicateCandidates = duplicateFileCandidates(files);
  for (const group of duplicateCandidates) {
    const paths = group.map((entry) => entry.relativePath).join(" | ");
    for (const entry of group) {
      processingLog.push({
        EVENT_TYPE: "DUPLICATE_FILE_CANDIDATE", SKIPPED_FILE: "", SKIPPED_SHEET: "", SKIP_REASON: "Same filename appears in more than one relative path; all files retained.", DETAIL: paths,
      });
    }
  }
  const log = {
    root_folder: resolvedSource, recursive_scan: true, files_found: files.length, files_excluded: excluded.length,
    scanned_files: files.length, successfully_read: 0, skipped_files: 0, sheets_processed: 0,
    total_records: 0, catalogs: {}, exception_count: 0, duplicate_file_candidates: duplicateCandidates.reduce((count, group) => count + group.length, 0), skipped: [],
  };
  const now = new Date().toISOString();

  for (const entry of files) {
    try {
      const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(entry.absolutePath));
      log.successfully_read++;
      for (const ws of wb.worksheets.items) {
        const used = ws.getUsedRange(true);
        const matrix = used ? used.values : [];
        if (!matrix?.length || !matrix[0]?.some((v) => norm(v))) {
          addException(exceptions, "", "", entry.name, ws.name, "EMPTY_SHEET", "Worksheet contains no header/data values.");
          processingLog.push({ EVENT_TYPE: "SKIPPED_SHEET", SKIPPED_FILE: entry.relativePath, SKIPPED_SHEET: ws.name, SKIP_REASON: "Worksheet contains no header/data values.", DETAIL: "" });
          continue;
        }
        const sourceHeaders = matrix[0].map(norm);
        const rows = matrix.slice(1).map((row, index) => ({ row, sourceRow: index + 2 })).filter(({ row }) => row.some((v) => norm(v)));
        if (!rows.length) {
          addException(exceptions, "", "", entry.name, ws.name, "EMPTY_SHEET", "Worksheet has a header but no data rows.");
          processingLog.push({ EVENT_TYPE: "SKIPPED_SHEET", SKIPPED_FILE: entry.relativePath, SKIPPED_SHEET: ws.name, SKIP_REASON: "Worksheet has a header but no data rows.", DETAIL: "" });
          continue;
        }
        let headers = [...sourceHeaders];
        let unlabeledCatalog = null;
        // Some EBS exports have values (rather than headers) in their first two cells.
        // This rule is deliberately narrow: column 2 must be non-empty, safe, and identical
        // in the title row and every retained data row.
        if (!headers.includes("CATALOG") && headers.length >= 2) {
          const candidate = safeCatalog(sourceHeaders[1]);
          if (candidate && rows.length && rows.every(({ row }) => norm(row[1]) === candidate)) {
            unlabeledCatalog = candidate;
            headers[0] = "UNNAMED_COLUMN_1";
            headers[1] = "UNNAMED_COLUMN_2";
          }
        }
        const itemIndex = headers.findIndex((h) => h === "ITEM_NUMBER");
        if (itemIndex < 0) {
          const detail = "Missing required validation header: ITEM_NUMBER. Worksheet skipped.";
          addException(exceptions, unlabeledCatalog || "", "", entry.name, ws.name, "SCHEMA_MISMATCH", detail);
          processingLog.push({ EVENT_TYPE: "SKIPPED_SHEET", SKIPPED_FILE: entry.relativePath, SKIPPED_SHEET: ws.name, SKIP_REASON: detail, DETAIL: "" });
          continue;
        }
        // A row repeating ITEM_NUMBER is an empty-template field guide, not an item record.
        const dataRows = rows.filter(({ row }) => !row.some((v) => norm(v) === "ITEM_NUMBER"));
        // A template with only an owner/catalog row and no ITEM_NUMBER is empty. This
        // preserves ordinary two-column data sheets whose ITEM_NUMBER is in column 1.
        const contentRows = dataRows.length && dataRows.every(({ row }) => !norm(row[itemIndex])) ? [] : dataRows;
        if (!contentRows.length) {
          addException(exceptions, unlabeledCatalog || "", "", entry.name, ws.name, "EMPTY_SHEET", "Worksheet contains only a template/header row; no item records loaded.");
          processingLog.push({ EVENT_TYPE: "SKIPPED_SHEET", SKIPPED_FILE: entry.relativePath, SKIPPED_SHEET: ws.name, SKIP_REASON: "Worksheet contains only a template/header row; no item records loaded.", DETAIL: "" });
          continue;
        }
        log.sheets_processed++;
        const fileStem = path.parse(entry.name).name;
        const sheetCatalogs = new Set();
        for (const { row, sourceRow } of contentRows) {
          const resolved = catalogFrom(ws.name, fileStem, headers, row, unlabeledCatalog);
          const item = itemIndex >= 0 ? norm(row[itemIndex]) : "";
          if (!resolved.catalog) {
            addException(exceptions, "", item, entry.name, ws.name, "UNKNOWN_CATALOG", `Unable to determine catalog: ${resolved.method}.`);
            continue;
          }
          sheetCatalogs.add(resolved.catalog);
          records.push({ catalog: resolved.catalog, headers, row, item, file: entry.name, sourceFolder: entry.sourceFolder, relativePath: entry.relativePath, sheet: ws.name, sourceRow, timestamp: now, lengthQA: evaluateDescriptionLength(headers, row) });
        }
        for (const catalog of sheetCatalogs) {
          const detail = schemaBase.has(catalog) ? compareSchema(schemaBase.get(catalog), headers) : "";
          if (!schemaBase.has(catalog)) schemaBase.set(catalog, headers);
          if (detail) addException(exceptions, catalog, "", entry.name, ws.name, "SCHEMA_MISMATCH", detail);
          if (unlabeledCatalog) addException(exceptions, catalog, "", entry.name, ws.name, "SCHEMA_MISMATCH", "The first two source columns have no header labels; retained as UNNAMED_COLUMN_1 and UNNAMED_COLUMN_2. CATALOG was read only because column 2 is constant across all data rows.");
        }
      }
    } catch (error) {
      log.skipped_files++;
      const reason = String(error.message || error);
      log.skipped.push({ file: entry.relativePath, reason });
      processingLog.push({ EVENT_TYPE: "SKIPPED_FILE", SKIPPED_FILE: entry.relativePath, SKIPPED_SHEET: "", SKIP_REASON: reason, DETAIL: "" });
    }
  }
  const byCatalog = new Map();
  for (const r of records) { if (!byCatalog.has(r.catalog)) byCatalog.set(r.catalog, []); byCatalog.get(r.catalog).push(r); }
  const itemCatalogs = new Map();
  for (const r of records) if (r.item) { if (!itemCatalogs.has(r.item)) itemCatalogs.set(r.item, new Set()); itemCatalogs.get(r.item).add(r.catalog); }
  for (const [catalog, list] of byCatalog) {
    const counts = new Map();
    for (const r of list) if (r.item) counts.set(r.item, (counts.get(r.item) || 0) + 1);
    for (const r of list) {
      if (!r.item) addException(exceptions, catalog, "", r.file, r.sheet, "MISSING_ITEM_NUMBER", "ITEM_NUMBER is blank or missing; record retained.");
      else if (counts.get(r.item) > 1) addException(exceptions, catalog, r.item, r.file, r.sheet, "DUPLICATE_ITEM_NUMBER", `ITEM_NUMBER occurs ${counts.get(r.item)} times within CATALOG; record retained.`);
      if (r.item && itemCatalogs.get(r.item).size > 1) addException(exceptions, catalog, r.item, r.file, r.sheet, "CROSS_CATALOG_DUPLICATE", `ITEM_NUMBER also occurs in: ${[...itemCatalogs.get(r.item)].filter((c) => c !== catalog).join(", ")}.`);
    }
  }
  log.total_records = records.length;
  for (const [catalog, list] of byCatalog) log.catalogs[catalog] = list.length;
  log.exception_count = exceptions.length;
  const summaryRows = [...byCatalog.keys()].sort().map((catalog) => {
    const list = byCatalog.get(catalog), items = list.map((r) => r.item).filter(Boolean), freq = new Map(); items.forEach((x) => freq.set(x, (freq.get(x) || 0) + 1));
    return [catalog, list.length, new Set(items).size, [...freq.values()].filter((n) => n > 1).reduce((s, n) => s + n, 0), list.filter((r) => !r.item).length, new Set(list.map((r) => r.file)).size, exceptions.filter((e) => e[0] === catalog).length];
  });
  const lengthQARows = records.map((record) => [
    record.catalog, record.item, record.file, record.relativePath, record.sheet, record.sourceRow,
    record.lengthQA.descriptionColumns.join(", "), record.lengthQA.status, record.lengthQA.maxWeightedLength,
    record.lengthQA.overLimit.map(({ column }) => column).join(", "), record.lengthQA.blankColumns.join(", "),
  ]);
  const lengthStatusCounts = new Map();
  for (const record of records) lengthStatusCounts.set(record.lengthQA.status, (lengthStatusCounts.get(record.lengthQA.status) || 0) + 1);
  const lengthQAStatus = (lengthStatusCounts.get("OVER_30") || 0) ? "OVER_30" : ((lengthStatusCounts.get("BLANK") || 0) ? "BLANK" : "OK");
  const qaSummaryRows = [
    ["ROOT_FOLDER", resolvedSource], ["RECURSIVE_SCAN", "ON"], ["FILES_FOUND", files.length], ["FILES_EXCLUDED", excluded.length],
    ["FILES_PROCESSED", log.successfully_read], ["FILES_SKIPPED", log.skipped_files], ["SHEETS_PROCESSED", log.sheets_processed],
    ["TOTAL_ROWS", records.length], ["OK_ROWS", lengthStatusCounts.get("OK") || 0], ["OVER_30_ROWS", lengthStatusCounts.get("OVER_30") || 0],
    ["BLANK_ROWS", lengthStatusCounts.get("BLANK") || 0], ["DUPLICATE_FILE_CANDIDATES", log.duplicate_file_candidates],
    ["LENGTH_QA_STATUS", lengthQAStatus], ["OUTPUT_EXCEL", resolvedOutput],
  ];
  const out = Workbook.create();
  const summary = out.worksheets.add("SUMMARY"); summary.getRange(`A1:${colName(SUMMARY_HEADERS.length - 1)}${summaryRows.length + 1}`).values = [SUMMARY_HEADERS, ...summaryRows]; styleSheet(summary, SUMMARY_HEADERS, summaryRows, true);
  const ex = out.worksheets.add("EXCEPTIONS"); ex.getRange(`A1:${colName(EXCEPTION_HEADERS.length - 1)}${exceptions.length + 1}`).values = [EXCEPTION_HEADERS, ...exceptions]; styleSheet(ex, EXCEPTION_HEADERS, exceptions);
  const qaSummary = out.worksheets.add("QA_SUMMARY"); qaSummary.getRange(`A1:B${qaSummaryRows.length + 1}`).values = [["METRIC", "VALUE"], ...qaSummaryRows]; styleSheet(qaSummary, ["METRIC", "VALUE"], qaSummaryRows);
  qaSummary.getRange("A:A").format.columnWidth = 32; qaSummary.getRange("B:B").format.columnWidth = 70;
  const qaHeaders = ["CATALOG", "ITEM_NUMBER", "SOURCE_FILE", "SOURCE_RELATIVE_PATH", "SOURCE_SHEET", "SOURCE_ROW", "DESCRIPTION_COLUMNS", "LENGTH_QA_STATUS", "MAX_WEIGHTED_LENGTH", "OVER_30_COLUMNS", "BLANK_DESCRIPTION_COLUMNS"];
  const qa = out.worksheets.add("LENGTH_QA"); qa.getRange(`A1:${colName(qaHeaders.length - 1)}${lengthQARows.length + 1}`).values = [qaHeaders, ...lengthQARows]; styleSheet(qa, qaHeaders, lengthQARows);
  if (lengthQARows.length) qa.getRange(`H2:H${lengthQARows.length + 1}`).conditionalFormats.add("containsText", { text: "OVER_30", format: { fill: "#FCE4D6", font: { bold: true, color: "#C00000" } } });
  const processingHeaders = ["EVENT_TYPE", "SKIPPED_FILE", "SKIPPED_SHEET", "SKIP_REASON", "DETAIL"];
  const processing = out.worksheets.add("PROCESSING_LOG"); processing.getRange(`A1:${colName(processingHeaders.length - 1)}${processingLog.length + 1}`).values = [processingHeaders, ...processingLog.map((entry) => processingHeaders.map((header) => entry[header] ?? ""))]; styleSheet(processing, processingHeaders, processingLog);
  processing.getRange("A:A").format.columnWidth = 28; processing.getRange("B:B").format.columnWidth = 55; processing.getRange("C:C").format.columnWidth = 28; processing.getRange("D:D").format.columnWidth = 55; processing.getRange("E:E").format.columnWidth = 75;
  const worksheetAliases = makeWorksheetAliases(byCatalog.keys());
  for (const catalog of [...byCatalog.keys()].sort()) {
    const list = byCatalog.get(catalog), rawHeaders = [...new Set(list.flatMap((r) => r.headers))];
    const tracking = TRACKING.map((h) => rawHeaders.includes(h) ? `${h}__TRACKING` : h);
    if (tracking.some((h) => h.includes("__TRACKING"))) addException(exceptions, catalog, "", "", "", "SCHEMA_MISMATCH", "A source column collides with a required tracking header; original retained and tracking header suffixed __TRACKING.");
    const headers = [...rawHeaders, ...tracking];
    const rows = list.map((r) => { const values = Object.fromEntries(r.headers.map((h, i) => [h, r.row[i] ?? null])); return [...rawHeaders.map((h) => values[h] ?? null), r.file, r.sourceFolder, r.relativePath, r.sheet, r.sourceRow, r.timestamp]; });
    const sheetName = worksheetAliases.get(catalog);
    const ws = out.worksheets.add(sheetName); ws.getRange(`A1:${colName(headers.length - 1)}${rows.length + 1}`).values = [headers, ...rows]; styleSheet(ws, headers, rows);
  }
  // Re-write exception rows after late tracking-collision checks.
  if (exceptions.length) ex.getRange(`A2:${colName(EXCEPTION_HEADERS.length - 1)}${exceptions.length + 1}`).values = exceptions;
  log.length_qa_status = lengthQAStatus;
  log.length_qa = { ok_rows: lengthStatusCounts.get("OK") || 0, over_30_rows: lengthStatusCounts.get("OVER_30") || 0, blank_rows: lengthStatusCounts.get("BLANK") || 0 };
  // Artifact Tool writes its inspection sidecar next to the exported workbook. Export to the
  // project diagnostics folder first, then publish only the finished Master workbook.
  const diagnosticsFolder = path.resolve("outputs");
  const outputName = path.basename(resolvedOutput);
  const stagingOutput = path.join(diagnosticsFolder, `.${outputName}.${process.pid}.staging.xlsx`);
  const stagingInspect = `${stagingOutput}.inspect.ndjson`;
  const inspectFile = path.join(diagnosticsFolder, `${outputName}.inspect.ndjson`);
  const logFile = path.join(diagnosticsFolder, outputName.replace(/\.xlsx$/i, "_processing_summary.json"));
  await fs.mkdir(diagnosticsFolder, { recursive: true });
  await fs.mkdir(path.dirname(resolvedOutput), { recursive: true });
  const blob = await SpreadsheetFile.exportXlsx(out);
  await blob.save(stagingOutput);
  await fs.copyFile(stagingOutput, resolvedOutput);
  await fs.unlink(stagingOutput);
  try {
    await fs.copyFile(stagingInspect, inspectFile);
    await fs.unlink(stagingInspect);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
  await fs.writeFile(logFile, JSON.stringify(log, null, 2), "utf8");
  return { outputFile: resolvedOutput, inspectFile, logFile, log, summaryRows, exceptions };
}

if (import.meta.url === `file:///${process.argv[1].replace(/\\/g, "/")}`) {
  const [sourceDir, outputFile] = process.argv.slice(2); if (!sourceDir) { usage(); process.exitCode = 1; } else { const result = await consolidate(sourceDir, outputFile); console.log(JSON.stringify(result.log, null, 2)); console.log(`Master workbook: ${result.outputFile}`); console.log(`Inspect result: ${result.inspectFile}`); console.log(`Processing summary: ${result.logFile}`); }
}
