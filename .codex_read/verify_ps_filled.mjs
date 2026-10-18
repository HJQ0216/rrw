import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const files = {
  single: "D:/钢结构RRW方向/制表/串联系统_RRW单参数计算需求表.xlsx",
  multi: "D:/钢结构RRW方向/制表/串联系统_RRW多参数计算需求表.xlsx",
};

const close = (actual, expected, tolerance = 1e-12) => Math.abs(actual - expected) <= tolerance;
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
};

const singleWb = await SpreadsheetFile.importXlsx(await FileBlob.load(files.single));
const single = singleWb.worksheets.getItem("单参数计算需求");
assert(single.getUsedRange().address === "A1:R29", `单参数范围异常：${single.getUsedRange().address}`);
const singlePs = single.getRange("J8:K29").values;
const singleFilled = singlePs.filter(([lower, upper]) => Number.isFinite(lower) && Number.isFinite(upper));
assert(singleFilled.length === 22, `单参数Pf(S)应回填22行，实际${singleFilled.length}行`);
assert(singleFilled.every(([lower, upper]) => lower >= 0 && lower <= upper && upper <= 1), "单参数Pf(S)上下界顺序或范围异常");
assert(close(single.getRange("J8").values[0][0], 0.01092), "单参数J8校核失败");
assert(close(single.getRange("J18").values[0][0], 0.03614), "n=28下界校核失败");
assert(close(single.getRange("K18").values[0][0], 0.03632), "n=28上界校核失败");
assert(close(single.getRange("J24").values[0][0], 0.01354), "β=3.0校核失败");
assert(close(single.getRange("J25").values[0][0], 0.00694), "β=3.2校核失败");
assert(close(single.getRange("K25").values[0][0], 0.00694), "β=3.2上界校核失败");
assert(close(single.getRange("J26").values[0][0], 0.00112), "β=3.7校核失败");
assert(close(single.getRange("K26").values[0][0], 0.00112), "β=3.7上界校核失败");
assert(close(single.getRange("J28").values[0][0], 0.011144), "n1=2校核失败");
assert(close(single.getRange("J29").values[0][0], 0.0113712), "n1=3校核失败");

const multiWb = await SpreadsheetFile.importXlsx(await FileBlob.load(files.multi));
const multi = multiWb.worksheets.getItem("多参数计算需求");
assert(multi.getUsedRange().address === "A1:U110", `多参数范围异常：${multi.getUsedRange().address}`);
const multiPs = multi.getRange("M8:N110").values;
const multiFilled = multiPs.filter(([lower, upper]) => Number.isFinite(lower) && Number.isFinite(upper));
assert(multiFilled.length === 103, `多参数Pf(S)应回填103行，实际${multiFilled.length}行`);
assert(multiFilled.every(([lower, upper]) => lower >= 0 && lower <= upper && upper <= 1), "多参数Pf(S)上下界顺序或范围异常");
assert(close(multi.getRange("M8").values[0][0], 0.01092), "多参数M8校核失败");
assert(close(multi.getRange("M13").values[0][0], 0.011144), "多参数n1=2校核失败");
assert(close(multi.getRange("M18").values[0][0], 0.0113712), "多参数n1=3校核失败");
assert(close(multi.getRange("M105").values[0][0], 0.00382), "多参数尾部下界校核失败");
assert(close(multi.getRange("N110").values[0][0], 0.02268), "多参数尾部上界校核失败");

console.log(JSON.stringify({
  single: { rows: 22, pfFilled: singleFilled.length, pfBlank: 22 - singleFilled.length },
  multi: { rows: 103, pfFilled: multiFilled.length, pfBlank: 103 - multiFilled.length },
  spotChecks: "passed",
}));
