import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = "D:/钢结构RRW方向";
const previewDir = path.join(root, ".codex_read", "rrw_submission_previews");
await fs.mkdir(previewDir, { recursive: true });

const files = [
  path.join(root, "数据文件", "串联系统单参数_RLP_RRW计算表.xlsx"),
  path.join(root, "数据文件", "串联系统多参数_RLP_RRW计算表.xlsx"),
];

for (const file of files) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
  const summary = await wb.inspect({
    kind: "workbook,sheet,table",
    maxChars: 16000,
    tableMaxRows: 15,
    tableMaxCols: 24,
    tableMaxCellChars: 100,
  });
  console.log(`FILE: ${file}`);
  console.log(summary.ndjson);

  for (const sheet of wb.worksheets.items) {
    const used = sheet.getUsedRange();
    console.log(`SHEET: ${sheet.name}; USED: ${used?.address ?? "none"}`);
    const preview = await wb.render({
      sheetName: sheet.name,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    const safeName = `${path.basename(file, ".xlsx")}_${sheet.name}`.replace(/[\\/:*?"<>|]/g, "_");
    await fs.writeFile(path.join(previewDir, `${safeName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
}
