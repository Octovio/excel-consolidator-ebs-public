import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import { WEIGHTED_LENGTH_LIMIT, weightedLength } from "./lengthRules.mjs";

const LIMIT = WEIGHTED_LENGTH_LIMIT;
const INFORMATIONAL_EXCLUDED_COLUMN = "未拆資訊";
const col = n => { let s=""; while(n>=0){s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)-1;} return s; };
const text = v => String(v ?? "").trim();
const isExcluded = name => name.startsWith("~$") || /(?:^|_)(?:master|qa report)|migration_cell_length_qa/i.test(name);
function style(ws, headers, rows, tableName) {
  ws.showGridLines=false; const end=col(headers.length-1);
  ws.getRange(`A1:${end}1`).format={fill:"#1F4E78",font:{bold:true,color:"#FFFFFF"},horizontalAlignment:"center",wrapText:true};
  ws.getRange(`A1:${end}1`).format.rowHeight=30; ws.freezePanes.freezeRows(1);
  if(rows.length) ws.tables.add(`A1:${end}${rows.length+1}`,true,tableName);
  for(let i=0;i<headers.length;i++) ws.getRange(`${col(i)}1:${col(i)}${Math.max(1,rows.length+1)}`).format.columnWidth=Math.min(45,Math.max(12,String(headers[i]).length+3));
}
export async function runLengthQA(inputDir, outputFile) {
  const entries=await fs.readdir(inputDir,{withFileTypes:true});
  const files=entries.filter(e=>e.isFile()&&e.name.toLowerCase().endsWith(".xlsx")&&!isExcluded(e.name));
  const exceptions=[]; let sheets=0, rowsScanned=0,cellsScanned=0,maxLen=0;
  let excludedNonEmptyCells=0, excludedOverLimitCells=0, excludedMaxLen=0;
  const fileStats=new Map(),colStats=new Map();
  for(const entry of files){
    const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(inputDir,entry.name)));
    for(const ws of wb.worksheets.items){
      sheets++; const matrix=ws.getUsedRange(true)?.values??[];
      if(!matrix.length||!matrix[0]?.some(v=>text(v))) continue;
      const headers=matrix[0].map(v=>text(v)); const itemCol=headers.findIndex(h=>h==="ITEM_NUMBER"), catalogCol=headers.findIndex(h=>h==="CATALOG");
      for(let r=1;r<matrix.length;r++){
        const row=matrix[r]; if(!row?.some(v=>text(v))) continue; rowsScanned++;
        for(let c=0;c<row.length;c++){
          const columnName=headers[c]||`COLUMN_${col(c)}`; const item=itemCol>=0?text(row[itemCol]):""; const catalog=catalogCol>=0?text(row[catalogCol]):"";
          const value=row[c]; if(!text(value)) continue; cellsScanned++; const length=weightedLength(String(value)); maxLen=Math.max(maxLen,length);
          if(columnName===INFORMATIONAL_EXCLUDED_COLUMN){
            excludedNonEmptyCells++; excludedMaxLen=Math.max(excludedMaxLen,length);
            if(length>LIMIT) excludedOverLimitCells++;
            continue;
          }
          if(length<=LIMIT) continue;
          exceptions.push([entry.name,ws.name,`${col(c)}${r+1}`,r+1,columnName,item,catalog,String(value),length,LIMIT,length-LIMIT,"MIGRATION_LENGTH_ERROR"]);
          const fsStat=fileStats.get(entry.name)||{count:0,sheets:new Set(),max:0}; fsStat.count++;fsStat.sheets.add(ws.name);fsStat.max=Math.max(fsStat.max,length);fileStats.set(entry.name,fsStat);
          const cs=colStats.get(columnName)||{count:0,max:0};cs.count++;cs.max=Math.max(cs.max,length);colStats.set(columnName,cs);
        }
      }
    }
  }
  exceptions.sort((a,b)=>String(a[0]).localeCompare(String(b[0]))||String(a[1]).localeCompare(String(b[1]))||a[3]-b[3]||String(a[2]).localeCompare(String(b[2])));
  const exceptionReportRows=exceptions.map(([sourceFile,sourceSheet,cellAddress,rowNumber,columnName,itemNumber,catalog,cellValue,weightedLength,_limit,exceededBy])=>[
    sourceFile,sourceSheet,cellAddress,rowNumber,itemNumber,catalog,columnName,cellValue,weightedLength,exceededBy
  ]);
  const summaryRows=[["FILES_SCANNED",files.length],["WORKSHEETS_SCANNED",sheets],["DATA_ROWS_SCANNED",rowsScanned],["NON_EMPTY_DATA_CELLS_SCANNED",cellsScanned],["MIGRATION_LENGTH_ERROR_COUNT",exceptions.length],["AFFECTED_SOURCE_FILE_COUNT",fileStats.size],["AFFECTED_WORKSHEET_COUNT",new Set(exceptions.map(x=>`${x[0]}|${x[1]}`)).size],["MAX_WEIGHTED_LENGTH",maxLen],["UNSPLIT_INFO_NON_EMPTY_CELL_COUNT",excludedNonEmptyCells],["UNSPLIT_INFO_OVER_LIMIT_CELL_COUNT",excludedOverLimitCells],["UNSPLIT_INFO_MAX_WEIGHTED_LENGTH",excludedMaxLen],["QA_STATUS",exceptions.length?"LENGTH QA NEEDS REVIEW":"LENGTH QA PASSED"]];
  const fileRows=[...fileStats].map(([file,s])=>[file,s.count,s.sheets.size,s.max]).sort((a,b)=>b[1]-a[1]||String(a[0]).localeCompare(String(b[0])));
  const columnRows=[...colStats].map(([name,s])=>[name,s.count,s.max]).sort((a,b)=>b[1]-a[1]||String(a[0]).localeCompare(String(b[0])));
  const out=Workbook.create();
  const sum=out.worksheets.add("SUMMARY"); sum.getRange(`A1:B${summaryRows.length+1}`).values=[["METRIC","VALUE"],...summaryRows]; style(sum,["METRIC","VALUE"],summaryRows,"LengthSummaryTable"); sum.getRange("A:A").format.columnWidth=40; sum.getRange("B:B").format.columnWidth=28;
  const eh=["SOURCE_FILE","SOURCE_SHEET","CELL_ADDRESS","ROW_NUMBER","ITEM_NUMBER","CATALOG","COLUMN_NAME","CELL_VALUE","WEIGHTED_LENGTH","EXCEEDED_BY"];
  const ex=out.worksheets.add("EXCEPTIONS"); ex.getRange(`A1:${col(eh.length-1)}${exceptionReportRows.length+1}`).values=[eh,...exceptionReportRows]; style(ex,eh,exceptionReportRows,"LengthExceptionsTable"); if(exceptionReportRows.length) ex.getRange(`H2:H${exceptionReportRows.length+1}`).format.wrapText=true;
  const fh=["SOURCE_FILE","MIGRATION_LENGTH_ERROR_COUNT","AFFECTED_WORKSHEET_COUNT","MAX_WEIGHTED_LENGTH"]; const fws=out.worksheets.add("SOURCE_FILE_SUMMARY"); fws.getRange(`A1:D${fileRows.length+1}`).values=[fh,...fileRows]; style(fws,fh,fileRows,"LengthFileSummaryTable"); fws.getRange("A:A").format.columnWidth=70; fws.getRange("B:B").format.columnWidth=38; fws.getRange("C:C").format.columnWidth=32; fws.getRange("D:D").format.columnWidth=28;
  const ch=["COLUMN_NAME","MIGRATION_LENGTH_ERROR_COUNT","MAX_WEIGHTED_LENGTH"]; const cws=out.worksheets.add("COLUMN_SUMMARY"); cws.getRange(`A1:C${columnRows.length+1}`).values=[ch,...columnRows]; style(cws,ch,columnRows,"LengthColumnSummaryTable"); cws.getRange("A:A").format.columnWidth=28; cws.getRange("B:B").format.columnWidth=38; cws.getRange("C:C").format.columnWidth=28;
  await fs.mkdir(path.dirname(outputFile),{recursive:true}); const blob=await SpreadsheetFile.exportXlsx(out);await blob.save(outputFile);
  return {files:files.length,sheets,rowsScanned,cellsScanned,errors:exceptions.length,affectedFiles:fileStats.size,affectedSheets:new Set(exceptions.map(x=>`${x[0]}|${x[1]}`)).size,maxLen,excludedColumn:INFORMATIONAL_EXCLUDED_COLUMN,excludedNonEmptyCells,excludedOverLimitCells,excludedMaxLen,exceptionDetails:exceptionReportRows,fileRows,columnRows,status:exceptions.length?"LENGTH QA NEEDS REVIEW":"LENGTH QA PASSED"};
}
if(import.meta.url===`file:///${process.argv[1].replace(/\\/g,"/")}`){const [dir,out]=process.argv.slice(2);if(!dir){console.error("Usage: node MigrationLengthQA.mjs <input-dir> <output.xlsx>");process.exit(1);}console.log(JSON.stringify(await runLengthQA(dir,out||path.join(dir,"Migration_Cell_Length_QA.xlsx")),null,2));}
