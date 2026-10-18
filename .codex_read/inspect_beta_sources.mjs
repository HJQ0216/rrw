import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sources = [
  [2.7, "D:/钢结构RRW方向/数据文件/β = 2.7串联系统 .xlsx"],
  [3.0, "D:/钢结构RRW方向/数据文件/β = 3.0串联系统 .xlsx"],
  [3.2, "D:/钢结构RRW方向/数据文件/β = 3.2串联系统 .xlsx"],
  [3.7, "D:/钢结构RRW方向/数据文件/β = 3.7串联系统 .xlsx"],
];

for (const [beta, path] of sources) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const sheets = wb.worksheets.items.map((sheet) => ({
    name: sheet.name,
    trimmed: sheet.name.trim(),
    used: sheet.getUsedRange()?.address ?? null,
  }));
  const sheet10 = wb.worksheets.items.find((sheet) => sheet.name.trim() === "10根构件");
  const probes = sheet10 ? {
    C13: sheet10.getRange("C13").values[0][0],
    D13: sheet10.getRange("D13").values[0][0],
    G13: sheet10.getRange("G13").values[0][0],
    H13: sheet10.getRange("H13").values[0][0],
    K13: sheet10.getRange("K13").values[0][0],
    L13: sheet10.getRange("L13").values[0][0],
    O13: sheet10.getRange("O13").values[0][0],
    P13: sheet10.getRange("P13").values[0][0],
  } : null;
  const sheet8 = wb.worksheets.items.find((sheet) => sheet.name.trim() === "8根构件");
  const sheet8Probes = sheet8 ? {
    B9: sheet8.getRange("B9").values[0][0],
    C13: sheet8.getRange("C13").values[0][0],
    D13: sheet8.getRange("D13").values[0][0],
    G13: sheet8.getRange("G13").values[0][0],
    H13: sheet8.getRange("H13").values[0][0],
  } : null;
  console.log(JSON.stringify({ beta, path, sheets, sheet10Probes: probes, sheet8Probes }));
}
