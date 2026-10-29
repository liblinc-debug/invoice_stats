import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const outputDir = process.argv[3];
if (!inputPath || !outputDir) throw new Error("usage: inspect_workbook.mjs input.xlsx outputDir");

await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const summary = await workbook.inspect({
  kind: "workbook,sheet,table,definedName,drawing",
  maxChars: 12000,
  tableMaxRows: 12,
  tableMaxCols: 12,
  tableMaxCellChars: 120,
});
console.log(summary.ndjson);

for (const sheet of workbook.worksheets.items) {
  const detail = await workbook.inspect({
    kind: "region",
    sheetId: sheet.name,
    range: "A1:L40",
    maxChars: 12000,
  });
  console.log(`\n--- ${sheet.name} ---`);
  console.log(detail.ndjson);
  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1.5,
    format: "png",
  });
  await fs.writeFile(`${outputDir}/${sheet.name}.png`, new Uint8Array(await preview.arrayBuffer()));
}
