import { runLengthQA } from "../src/MigrationLengthQA.mjs";
const [inputDir, outputFile] = process.argv.slice(2);
if (!inputDir) { console.error("Usage: node Run_Migration_Length_QA.mjs <input-folder> [output.xlsx]"); process.exit(1); }
console.log(JSON.stringify(await runLengthQA(inputDir, outputFile || "outputs/Migration_Cell_Length_QA.xlsx"), null, 2));
