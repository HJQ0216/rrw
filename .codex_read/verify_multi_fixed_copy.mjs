import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = "D:/钢结构RRW方向/制表/串联系统_RRW多参数计算需求表_修正版.xlsx";
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sheet = wb.worksheets.getItem("多参数计算需求");
if (sheet.getUsedRange().address !== "A1:U110") throw new Error("工作表范围异常");
const values = sheet.getRange("M8:N110").values;
const filled = values.filter(([lower, upper]) => Number.isFinite(lower) && Number.isFinite(upper));
if (filled.length !== 103) throw new Error(`Pf(S)回填行数异常：${filled.length}`);
if (!filled.every(([lower, upper]) => lower >= 0 && lower <= upper && upper <= 1)) throw new Error("Pf(S)上下界异常");
console.log(JSON.stringify({ usedRange: "A1:U110", pfFilled: filled.length, verified: true }));
