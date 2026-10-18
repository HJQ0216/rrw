import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const files = [
  {
    path: "D:/钢结构RRW方向/outputs/01a0ad35-6e80-7170-975c-a65e666fbf16/串联系统_RRW单参数计算需求表.xlsx",
    sheet: "单参数计算需求",
    expectedRange: "A1:R33",
    inputRange: "J8:M8",
    resultRange: "N8:Q8",
  },
  {
    path: "D:/钢结构RRW方向/outputs/01a0ad35-6e80-7170-975c-a65e666fbf16/串联系统_RRW多参数计算需求表.xlsx",
    sheet: "多参数计算需求",
    expectedRange: "A1:U110",
    inputRange: "M8:P8",
    resultRange: "Q8:T8",
  },
];

for (const item of files) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(item.path));
  const sheet = wb.worksheets.getItem(item.sheet);
  const used = sheet.getUsedRange();
  if (used.address !== item.expectedRange) throw new Error(`${item.path}: used range ${used.address}`);

  sheet.getRange(item.inputRange).values = [[0.01092, 0.01092, 0.0093, 0.0093]];
  wb.recalculate();
  const result = await wb.inspect({
    kind: "table",
    range: `${item.sheet}!${item.resultRange}`,
    include: "values,formulas",
    tableMaxRows: 2,
    tableMaxCols: 6,
    maxChars: 2500,
  });
  console.log(`VERIFIED ${item.path}`);
  console.log(result.ndjson);
}
