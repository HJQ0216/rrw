import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sources = [
  [2.7, "D:/钢结构RRW方向/数据文件/β = 2.7串联系统 .xlsx", "11根构件"],
  [3.0, "D:/钢结构RRW方向/数据文件/β = 3.0串联系统 .xlsx", "10根构件"],
  [3.2, "D:/钢结构RRW方向/数据文件/β = 3.2串联系统 .xlsx", "10根构件"],
  [3.7, "D:/钢结构RRW方向/数据文件/β = 3.7串联系统 .xlsx", "10根构件"],
];

for (const [beta, path, wanted] of sources) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const sheet = wb.worksheets.items.find((item) => item.name.trim() === wanted);
  if (!sheet) throw new Error(`β=${beta}缺少${wanted}`);
  console.log(JSON.stringify({ beta, sheet: sheet.name, range: "A1:P15", values: sheet.getRange("A1:P15").values }));
}
