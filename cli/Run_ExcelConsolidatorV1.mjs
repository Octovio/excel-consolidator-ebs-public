// Command-line entry point for the Public Portfolio Edition consolidator with recursive root-folder scanning.
import { consolidate, formalMasterOutputPath } from "../src/ExcelConsolidatorV1.mjs";

const args = process.argv.slice(2);
const valueAfter = (flag) => { const index = args.indexOf(flag); return index >= 0 ? args[index + 1] : ""; };
const sourceFolder = valueAfter("--root-folder") || args.find((arg) => !arg.startsWith("--"));
const explicitOutput = valueAfter("--output");
if (!sourceFolder) {
  console.error("Usage: node Run_ExcelConsolidatorV1.mjs <source-folder> [output-file] | --root-folder <folder> [--recursive] [--output <file>]");
  process.exit(1);
}
const positionalOutput = args.filter((arg, index) => !arg.startsWith("--") && args[index - 1] !== "--root-folder" && args[index - 1] !== "--output")[1];
const outputFile = explicitOutput || positionalOutput || formalMasterOutputPath();
const result = await consolidate(sourceFolder, outputFile);
console.log(JSON.stringify(result.log, null, 2));
console.log(`Master workbook: ${result.outputFile}`);
console.log(`Inspect result: ${result.inspectFile}`);
console.log(`Processing summary: ${result.logFile}`);
