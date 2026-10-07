import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import { consolidate, FORMAL_MASTER_OUTPUT_FOLDER, formalMasterOutputPath, scanExcelFilesRecursive } from "../src/ExcelConsolidatorV1.mjs";

async function createWorkbook(filePath, sheets) {
  const workbook = Workbook.create();
  for (const { name, rows } of sheets) {
    const sheet = workbook.worksheets.add(name);
    sheet.getRange(`A1:${columnName(rows[0].length - 1)}${rows.length}`).values = rows;
  }
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(filePath);
}

function columnName(index) {
  let result = "";
  for (let current = index; current >= 0; current = Math.floor(current / 26) - 1) result = String.fromCharCode(65 + (current % 26)) + result;
  return result;
}

function matrixToObjects(values) {
  const [headers, ...rows] = values;
  return rows.filter((row) => row.some((value) => String(value ?? "").trim())).map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
}

const fixtureRoot = await fs.mkdtemp(path.join(os.tmpdir(), "public-consolidator-"));
const formalTestFolder = path.join(fixtureRoot, "formal-master-output");
const output = formalMasterOutputPath(new Date(2099, 0, 2), formalTestFolder);
const direct = path.join(fixtureRoot, "direct.xlsx");

assert.equal(
  formalMasterOutputPath(new Date(2026, 9, 1)),
  path.join(FORMAL_MASTER_OUTPUT_FOLDER, "Master_EBS_20261001.xlsx"),
  "default formal path uses the requested folder and date-only filename",
);

await createWorkbook(direct, [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER", "DESC1"], ["DIRECT", "D-001", "short"]] }]);
await createWorkbook(path.join(fixtureRoot, "_AUTOMATION_TEST_正式.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER", "DESC1"], ["AUTOMATION", "A-001", "kept"]] }]);
await createWorkbook(path.join(fixtureRoot, "~$temp.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER"], ["TEMP", "T-001"]] }]);
await createWorkbook(path.join(fixtureRoot, "CONSOLIDATED_OLD.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER"], ["OUTPUT", "O-001"]] }]);
await createWorkbook(path.join(fixtureRoot, "WHEEL ASSEMBLY", "one.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER", "DESC1"], ["WHEEL", "W-001", "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDE"]] }]);
await createWorkbook(path.join(fixtureRoot, "中文 資料夾", "深層 folder", "第二.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER", "DESC1"], ["中文", "C-001", ""]] }]);
await createWorkbook(path.join(fixtureRoot, "duplicate a", "duplicate.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER", "DESC1"], ["DUP", "DUP-001", "same description"]] }]);
await createWorkbook(path.join(fixtureRoot, "duplicate b", "duplicate.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "ITEM_NUMBER", "DESC1"], ["DUP", "DUP-002", "same description"]] }]);
await createWorkbook(path.join(fixtureRoot, "bad schema", "BAD_SCHEMA.xlsx"), [{ name: "工作表1", rows: [["CATALOG", "DESC1"], ["BAD", "not processed"]] }]);

const directBefore = await fs.readFile(direct);
const scan = await scanExcelFilesRecursive(fixtureRoot, output);
assert.equal(scan.files.some((file) => file.name === "_AUTOMATION_TEST_正式.xlsx"), true, "automation output remains a source");
assert.equal(scan.files.some((file) => file.name === "~$temp.xlsx"), false, "temporary file excluded");
assert.equal(scan.files.some((file) => file.name === "CONSOLIDATED_OLD.xlsx"), false, "Consolidator output excluded");
assert.equal(scan.files.some((file) => file.relativePath === "中文 資料夾\\深層 folder\\第二.xlsx"), true, "nested Chinese path found");

const result = await consolidate(fixtureRoot, output);
assert.equal(result.outputFile, output, "Master output path is the formal output folder");
assert.equal((await fs.stat(output)).isFile(), true, "Master workbook was published to the formal output folder");
assert.equal(path.dirname(result.inspectFile), path.resolve("outputs"), "inspect file stays in project outputs");
assert.equal((await fs.stat(result.inspectFile)).isFile(), true, "inspect file was retained in outputs");
assert.equal((await fs.stat(result.logFile)).isFile(), true, "processing log was retained in outputs");
await assert.rejects(fs.stat(`${output}.inspect.ndjson`), { code: "ENOENT" }, "formal output folder contains no inspect sidecar");
assert.equal(result.log.successfully_read, 7, "all non-excluded fixture workbooks read");
assert.equal(result.log.skipped_files, 0, "schema mismatch skips only its sheet");
assert.equal(result.log.total_records, 6, "schema mismatch sheet was not ingested");
assert.equal(result.log.duplicate_file_candidates, 2, "same filename produces a warning only");
assert.deepEqual(result.log.length_qa, { ok_rows: 4, over_30_rows: 1, blank_rows: 1 }, "output routing does not change Length QA statistics");
assert.equal((await fs.readFile(direct)).equals(directBefore), true, "source workbook bytes unchanged");

const master = await SpreadsheetFile.importXlsx(await FileBlob.load(output));
const directSheet = master.worksheets.getItem("DIRECT");
const directRows = matrixToObjects(directSheet.getUsedRange(true).values);
assert.equal(directRows[0].SOURCE_FILE, "direct.xlsx");
assert.equal(directRows[0].SOURCE_RELATIVE_PATH, "direct.xlsx");
assert.equal(directRows[0].SOURCE_SHEET, "工作表1");
assert.equal(directRows[0].SOURCE_ROW, 2);
assert.equal(directRows[0].SOURCE_FOLDER, path.basename(fixtureRoot));
assert.equal(matrixToObjects(master.worksheets.getItem("AUTOMATION").getUsedRange(true).values)[0].SOURCE_FILE, "_AUTOMATION_TEST_正式.xlsx", "automation output is retained as a source");

const wheelRows = matrixToObjects(master.worksheets.getItem("WHEEL").getUsedRange(true).values);
assert.equal(wheelRows[0].SOURCE_FOLDER, "WHEEL ASSEMBLY");
assert.equal(wheelRows[0].SOURCE_RELATIVE_PATH, "WHEEL ASSEMBLY\\one.xlsx");
assert.equal(wheelRows[0].DESC1, "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDE", "description was not truncated");

const qaRows = matrixToObjects(master.worksheets.getItem("LENGTH_QA").getUsedRange(true).values);
assert.equal(qaRows.find((row) => row.ITEM_NUMBER === "W-001").LENGTH_QA_STATUS, "OVER_30");
assert.equal(qaRows.find((row) => row.ITEM_NUMBER === "D-001").LENGTH_QA_STATUS, "OK");
assert.equal(qaRows.find((row) => row.ITEM_NUMBER === "C-001").LENGTH_QA_STATUS, "BLANK");

const processRows = matrixToObjects(master.worksheets.getItem("PROCESSING_LOG").getUsedRange(true).values);
assert.equal(processRows.some((row) => row.EVENT_TYPE === "SKIPPED_SHEET" && row.SKIP_REASON.includes("ITEM_NUMBER")), true, "schema mismatch is recorded and does not stop batch");
assert.equal(processRows.filter((row) => row.EVENT_TYPE === "DUPLICATE_FILE_CANDIDATE").length, 2, "duplicate candidates are logged but retained");

await fs.unlink(result.inspectFile);
await fs.unlink(result.logFile);

console.log(JSON.stringify({ status: "PASS", fixtureRoot, output, records: result.log.total_records, tests: 25 }, null, 2));
