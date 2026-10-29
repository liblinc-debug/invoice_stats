import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = process.cwd();
const inputPath = path.join(root, "IDEA项目材料台账.xlsx");
const outputDir = path.join(root, "outputs", "01a185e3-f23b-7633-8125-e52dcad7b006");
const outputPath = path.join(outputDir, "IDEA项目材料台账_发票信息统计.xlsx");
const previewDir = path.join(root, "tmp", "final-previews");

const materialRecords = [
  ["2026-01-15", "26952000000168054676", "*计算机外部设备*摄像头模组", 45.70],
  ["2026-01-26", "26327000000238137466", "*运输服务*客运服务费", 31.63],
  ["2026-03-08", "26447000000558405770", "*配电控制设备*公牛4位总控3米插座GN-S1440", 72.90],
  ["2026-03-10", "26952000000966887836", "*供电*排插", 76.00],
  ["2026-03-30", "26312000001919704966", "*通信传输设备*移远RG200U物联网5G全网通模组", 900.00],
  ["2026-03-30", "26312000001919713546", "*通信传输设备*移远RG200U物联网5G全网通模组", 98.80],
  ["2026-04-10", "26447000000784079908", "*计算机外部设备*金士顿64GB MicroSD存储卡；*照相器材*绿联USB/Type-C 3.0高速读卡器", 417.90],
  ["2026-04-20", "26952000001601800036", "*电子元件*扩展板 USB 3.2 Gen1 5G DONGLE", 364.32],
  ["2026-05-17", "26447000001022400180", "*快递服务*收派服务费", 13.00],
  ["2026-07-20", "26952000003077263546", "*塑料制品*3D打印服务", 190.00],
  ["2026-08-04", "26322000006407666686", "*密封用填料*胶水（75g 200R）", 21.00],
  ["2026-08-04", "26952000003312060376", "*电子元件*转接线", 32.34],
  ["2026-08-06", "26442000009030337366", "*通信传输设备*转接线", 33.00],
  ["2026-08-14", "26412000003076073791", "*金属制品*铝箔卷", 19.00],
  ["2026-08-21", "26322000006868194721", "*金属制品*螺母（M3*3*5；M3*3*4）", 5.70],
  ["2026-08-23", "26117000001277146953", "*其他机械设备*彩格3D打印机平台热床固体胶", 9.00],
  ["2026-08-26", "26957000000228899086", "*生产生活服务*收派服务费", 12.00],
  ["2026-08-31", "26322000007195046926", "*日用杂品*电烙铁套装", 35.34],
  ["2026-08-31", "26322000007199457841", "*金属制品*烙铁头", 18.35],
  ["2026-08-31", "26952000003701032201", "*塑料制品*3D打印机线材 PETG Basic 黑色（带盘）", 46.83],
  ["2026-09-05", "26957000000238385844", "*生产生活服务*收派服务费", 12.00],
  ["2026-09-14", "26952000003905827591", "*塑料制品*3D打印机线材 PLA Lite 1kg黑色（无料盘）", 123.73],
  ["2026-10-05", "26334000000390703771", "*电子元件*连接器", 89.40],
  ["2026-10-15", "26957000000267165771", "*生产生活服务*收派服务费", 15.00],
];

const travelRecords = [
  ["2026-04-30", "26442000004764055396", "*住宿服务*住宿", 674.22],
  ["2026-05-11", "26442000005137930711", "*住宿服务*住宿", 738.00],
  ["2026-05-18", "26442000005438322376", "*住宿服务*住宿", 1014.00],
  ["2026-05-24", "26444000000134833081", "*住宿服务*住宿", 1152.00],
  ["2026-05-31", "26442000006079206016", "*生产生活服务*住宿", 738.00],
  ["2026-06-12", "26442000006581709526", "*生产生活服务*住宿", 600.00],
  ["2026-06-18", "26444000000165681121", "*生产生活服务*住宿", 138.00],
  ["2026-06-29", "26442000007294512331", "*生产生活服务*住宿", 552.00],
  ["2026-09-18", "26444000000280750246", "*生产生活服务*住宿", 1338.00],
];

function assertSortedAndUnique(records, label) {
  const invoiceNumbers = new Set();
  let previousDate = "";
  for (const [date, invoiceNumber] of records) {
    if (date < previousDate) throw new Error(`${label} records are not sorted: ${date} < ${previousDate}`);
    if (invoiceNumbers.has(invoiceNumber)) throw new Error(`${label} duplicate invoice number: ${invoiceNumber}`);
    previousDate = date;
    invoiceNumbers.add(invoiceNumber);
  }
}

assertSortedAndUnique(materialRecords, "材料费");
assertSortedAndUnique(travelRecords, "差旅费");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const materialSheet = workbook.worksheets.getItem("材料费");
const travelSheet = workbook.worksheets.getItem("差旅费");

materialSheet.getRange("A4:J100").clear({ applyTo: "contents" });
travelSheet.getRange("A4:G100").clear({ applyTo: "contents" });

const materialValues = materialRecords.map(([date, invoiceNumber, itemName, amount], index) => [
  index + 1,
  new Date(`${date}T00:00:00+08:00`),
  `\u2060${invoiceNumber}`,
  itemName,
  null,
  null,
  amount,
  null,
  null,
  null,
]);
const materialEndRow = 3 + materialValues.length;
materialSheet.getRange(`A4:J${materialEndRow}`).values = materialValues;
materialSheet.getRange(`B4:B${materialEndRow}`).format.numberFormat = "yyyy-mm-dd";
materialSheet.getRange(`C4:C${materialEndRow}`).setNumberFormat("@");
materialSheet.getRange(`G4:G${materialEndRow}`).format.numberFormat = "#,##0.00";
materialSheet.getRange(`D4:D${materialEndRow}`).format.wrapText = true;
materialSheet.getRange(`A4:J${materialEndRow}`).format.verticalAlignment = "center";
materialSheet.getRange(`A4:A${materialEndRow}`).format.horizontalAlignment = "center";
materialSheet.getRange(`B4:C${materialEndRow}`).format.horizontalAlignment = "center";
materialSheet.getRange(`G4:G${materialEndRow}`).format.horizontalAlignment = "right";
materialSheet.getRange(`D4:D${materialEndRow}`).format.autofitRows();

const travelValues = travelRecords.map(([date, invoiceNumber, itemName, amount], index) => [
  index + 1,
  new Date(`${date}T00:00:00+08:00`),
  null,
  `\u2060${invoiceNumber}`,
  itemName,
  amount,
  null,
]);
const travelEndRow = 3 + travelValues.length;
travelSheet.getRange(`A4:G${travelEndRow}`).values = travelValues;
travelSheet.getRange(`B4:B${travelEndRow}`).format.numberFormat = "yyyy-mm-dd";
travelSheet.getRange(`D4:D${travelEndRow}`).setNumberFormat("@");
travelSheet.getRange(`F4:F${travelEndRow}`).format.numberFormat = "#,##0.00";
travelSheet.getRange(`E4:E${travelEndRow}`).format.wrapText = true;
travelSheet.getRange(`A4:G${travelEndRow}`).format.verticalAlignment = "center";
travelSheet.getRange(`D1:D${travelEndRow}`).format.columnWidth = 25;
travelSheet.getRange(`A4:A${travelEndRow}`).format.horizontalAlignment = "center";
travelSheet.getRange(`B4:B${travelEndRow}`).format.horizontalAlignment = "center";
travelSheet.getRange(`D4:D${travelEndRow}`).format.horizontalAlignment = "center";
travelSheet.getRange(`F4:F${travelEndRow}`).format.horizontalAlignment = "right";
travelSheet.getRange(`E4:E${travelEndRow}`).format.autofitRows();

workbook.recalculate();

const materialCheck = await workbook.inspect({
  kind: "table",
  range: `材料费!A1:J${materialEndRow}`,
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 10,
  maxChars: 30000,
});
const travelCheck = await workbook.inspect({
  kind: "table",
  range: `差旅费!A1:G${travelEndRow}`,
  include: "values,formulas",
  tableMaxRows: 15,
  tableMaxCols: 7,
  maxChars: 20000,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log("--- MATERIAL CHECK ---");
console.log(materialCheck.ndjson);
console.log("--- TRAVEL CHECK ---");
console.log(travelCheck.ndjson);
console.log("--- FORMULA ERRORS ---");
console.log(errors.ndjson);

await fs.mkdir(previewDir, { recursive: true });
const materialPreview = await workbook.render({
  sheetName: "材料费",
  range: `A1:J${materialEndRow}`,
  scale: 1.5,
  format: "png",
});
await fs.writeFile(path.join(previewDir, "材料费.png"), new Uint8Array(await materialPreview.arrayBuffer()));
const travelPreview = await workbook.render({
  sheetName: "差旅费",
  range: `A1:G${travelEndRow}`,
  scale: 1.5,
  format: "png",
});
await fs.writeFile(path.join(previewDir, "差旅费.png"), new Uint8Array(await travelPreview.arrayBuffer()));

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

const savedWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const savedMaterial = await savedWorkbook.inspect({
  kind: "table",
  range: `材料费!A3:G${materialEndRow}`,
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 7,
  maxChars: 30000,
});
const savedTravel = await savedWorkbook.inspect({
  kind: "table",
  range: `差旅费!A3:F${travelEndRow}`,
  include: "values,formulas",
  tableMaxRows: 15,
  tableMaxCols: 6,
  maxChars: 20000,
});
const savedMaterialNumbers = savedWorkbook.worksheets
  .getItem("材料费")
  .getRange(`C4:C${materialEndRow}`)
  .values
  .flat()
  .map((value) => String(value).replaceAll("\u2060", ""));
const savedTravelNumbers = savedWorkbook.worksheets
  .getItem("差旅费")
  .getRange(`D4:D${travelEndRow}`)
  .values
  .flat()
  .map((value) => String(value).replaceAll("\u2060", ""));
if (JSON.stringify(savedMaterialNumbers) !== JSON.stringify(materialRecords.map((row) => row[1]))) {
  throw new Error("Saved material invoice numbers do not match source invoice numbers");
}
if (JSON.stringify(savedTravelNumbers) !== JSON.stringify(travelRecords.map((row) => row[1]))) {
  throw new Error("Saved travel invoice numbers do not match source invoice numbers");
}
console.log("--- SAVED MATERIAL ---");
console.log(savedMaterial.ndjson);
console.log("--- SAVED TRAVEL ---");
console.log(savedTravel.ndjson);
console.log(JSON.stringify({
  outputPath,
  materialCount: materialRecords.length,
  travelCount: travelRecords.length,
  materialTotal: materialRecords.reduce((sum, row) => sum + row[3], 0),
  travelTotal: travelRecords.reduce((sum, row) => sum + row[3], 0),
}));
