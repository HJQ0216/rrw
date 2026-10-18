import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const items = [
  {
    path: "D:/钢结构RRW方向/制表/串联系统_RRW单参数计算需求表.xlsx",
    sheetName: "单参数计算需求",
    range: "A1:R40",
    preview: "D:/钢结构RRW方向/.codex_read/current_single.png",
  },
  {
    path: "D:/钢结构RRW方向/制表/串联系统_RRW多参数计算需求表.xlsx",
    sheetName: "多参数计算需求",
    range: "A1:U18",
    preview: "D:/钢结构RRW方向/.codex_read/current_multi.png",
  },
];

for (const item of items) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(item.path));
  const sheet = wb.worksheets.getItem(item.sheetName);
  const used = sheet.getUsedRange();
  const inspect = await wb.inspect({
    kind: "table",
    range: `${item.sheetName}!${used.address}`,
    include: "values,formulas",
    tableMaxRows: 45,
    tableMaxCols: 21,
    maxChars: 25000,
  });
  console.log(JSON.stringify({ path: item.path, used: used.address }));
  console.log(inspect.ndjson);
  const preview = await wb.render({ sheetName: item.sheetName, range: item.range, scale: 1, format: "png" });
  await fs.writeFile(item.preview, new Uint8Array(await preview.arrayBuffer()));
}
