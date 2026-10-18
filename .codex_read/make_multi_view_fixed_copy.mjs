import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "D:/钢结构RRW方向/制表/串联系统_RRW多参数计算需求表.xlsx";
const output = "D:/钢结构RRW方向/制表/串联系统_RRW多参数计算需求表_修正版.xlsx";

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const sheet = wb.worksheets.getItem("多参数计算需求");
sheet.freezePanes.unfreeze();
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(output);
