import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "D:/钢结构RRW方向";
const outputDir = path.join(root, "制表");
const previewDir = path.join(root, ".codex_read", "rrw_submission_final_previews");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

const singleOutput = path.join(outputDir, "串联系统_RRW单参数计算需求表.xlsx");
const multiOutput = path.join(outputDir, "串联系统_RRW多参数计算需求表.xlsx");
const rawSources = new Map([
  [3.0, path.join(root, "数据文件", "β = 3.0串联系统 .xlsx")],
  [3.2, path.join(root, "数据文件", "β = 3.2串联系统 .xlsx")],
  [3.7, path.join(root, "数据文件", "β = 3.7串联系统 .xlsx")],
]);

const nValues = [8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48];
const rhoValues = [0.1, 0.3, 0.5, 0.7, 0.9];
const betaValues = [3.0, 3.2, 3.7];
const n1Values = [1, 2, 3];

const singleRows = [];
rhoValues.forEach((rho, index) => singleRows.push({
  displayId: `SP-RHO-${String(index + 1).padStart(2, "0")}`,
  group: "相关系数ρ",
  changeValue: rho,
  n: 8, n1: 1, beta: 3.0, rho, p: 0.2, k: 3,
  note: "固定n=8、n1=1、β=3.0、Pbi=0.2；对应论文图2.6",
}));
nValues.forEach((n, index) => singleRows.push({
  displayId: `SP-N-${String(index + 1).padStart(2, "0")}`,
  group: "构件总数n",
  changeValue: n,
  n, n1: 1, beta: 3.0, rho: 0.1, p: 0.2, k: 3,
  note: "固定n1=1、β=3.0、ρ=0.1、Pbi=0.2；对应论文图2.7",
}));
betaValues.forEach((beta, index) => singleRows.push({
  displayId: `SP-BETA-${String(index + 1).padStart(2, "0")}`,
  group: "可靠指标β",
  changeValue: beta,
  n: 10, n1: 1, beta, rho: 0.1, p: 0.2, k: 3,
  note: beta === 3.7
    ? "固定n=10、n1=1、ρ=0.1、Pbi=0.2；新增原始数据点（超出论文图2.8原β=3.0～3.6范围）"
    : beta === 3.2
      ? "固定n=10、n1=1、ρ=0.1、Pbi=0.2；对应论文图2.8；原始10根构件表B9标签误写为8个全塑性，按工作表名及K=3数值区读取"
      : "固定n=10、n1=1、ρ=0.1、Pbi=0.2；对应论文图2.8",
}));
n1Values.forEach((n1, index) => singleRows.push({
  displayId: `SP-N1-${String(index + 1).padStart(2, "0")}`,
  group: "潜在脆性构件数n1",
  changeValue: n1,
  n: 8, n1, beta: 3.0, rho: 0.1, p: 0.2, k: 3,
  note: "固定n=8、β=3.0、ρ=0.1、Pbi=0.2；对应论文图2.9",
}));

const multiRows = [];
n1Values.forEach((n1) => rhoValues.forEach((rho, index) => multiRows.push({
  displayId: `MP-RHO-N1-${n1}-${String(index + 1).padStart(2, "0")}`,
  group: "ρ×n1",
  parameter1: "相关系数ρ", parameter1Value: rho,
  parameter2: "潜在脆性构件数n1", parameter2Value: n1,
  n: 8, n1, beta: 3.0, rho, p: 0.2, k: 3,
  note: "固定n=8、β=3.0、Pbi=0.2；对应论文图2.10",
})));
n1Values.forEach((n1) => nValues.forEach((n, index) => multiRows.push({
  displayId: `MP-N-N1-${n1}-${String(index + 1).padStart(2, "0")}`,
  group: "n×n1",
  parameter1: "构件总数n", parameter1Value: n,
  parameter2: "潜在脆性构件数n1", parameter2Value: n1,
  n, n1, beta: 3.0, rho: 0.3, p: 0.2, k: 3,
  note: "固定ρ=0.3、β=3.0、Pbi=0.2；对应论文图2.11",
})));
rhoValues.forEach((rho) => nValues.forEach((n, index) => multiRows.push({
  displayId: `MP-N-RHO-${String(Math.round(rho * 10)).padStart(2, "0")}-${String(index + 1).padStart(2, "0")}`,
  group: "n×ρ",
  parameter1: "构件总数n", parameter1Value: n,
  parameter2: "相关系数ρ", parameter2Value: rho,
  n, n1: 1, beta: 3.0, rho, p: 0.2, k: 3,
  note: "固定n1=1、β=3.0、Pbi=0.2；对应论文图2.12",
})));

if (singleRows.length !== 22) throw new Error(`单参数展示点应为22，实际为${singleRows.length}`);
if (multiRows.length !== 103) throw new Error(`多参数展示点应为103，实际为${multiRows.length}`);

const rawWbs = new Map();
for (const [beta, source] of rawSources) {
  rawWbs.set(beta, await SpreadsheetFile.importXlsx(await FileBlob.load(source)));
}
const rhoRefs = {
  "0.1": { row: 13, pairs: [["C", "D"], ["G", "H"], ["K", "L"], ["O", "P"]] },
  "0.3": { row: 13, pairs: [["U", "V"], ["Y", "Z"], ["AC", "AD"], ["AG", "AH"]] },
  "0.5": { row: 34, pairs: [["C", "D"], ["G", "H"], ["K", "L"], ["O", "P"]] },
  "0.7": { row: 34, pairs: [["U", "V"], ["Y", "Z"], ["AC", "AD"], ["AG", "AH"]] },
  "0.9": { row: 54, pairs: [["C", "D"], ["G", "H"], ["K", "L"], ["O", "P"]] },
};

function choose(n, b) {
  let value = 1;
  for (let i = 1; i <= b; i += 1) value = value * (n - i + 1) / i;
  return value;
}

function getRawSheet(n, beta) {
  const target = `${n}根构件`;
  const rawWb = rawWbs.get(beta);
  if (!rawWb) throw new Error(`缺少β=${beta}的原始RLP工作簿`);
  const sheet = rawWb.worksheets.items.find((item) => item.name.trim() === target);
  if (!sheet) throw new Error(`β=${beta}原始RLP工作簿缺少工作表：${target}`);
  return sheet;
}

function getSystemBounds(n, n1, beta, rho, p = 0.2) {
  const ref = rhoRefs[rho.toFixed(1)];
  if (!ref) throw new Error(`未配置ρ=${rho}的RLP单元格映射`);
  if (n1 < 0 || n1 > 3) throw new Error(`当前原始数据仅支持0≤n1≤3，收到n1=${n1}`);
  const sheet = getRawSheet(n, beta);
  let lower = 0;
  let upper = 0;
  for (let b = 0; b <= n1; b += 1) {
    const weight = choose(n1, b) * (p ** b) * ((1 - p) ** (n1 - b));
    let qLower;
    let qUpper;
    if (beta === 3.0 && n === 28 && rho === 0.1 && b === 0) {
      qLower = 0.0359;
      qUpper = 0.0361;
    } else {
      const [lowerCol, upperCol] = ref.pairs[b];
      qLower = Number(sheet.getRange(`${lowerCol}${ref.row}`).values[0][0]);
      qUpper = Number(sheet.getRange(`${upperCol}${ref.row}`).values[0][0]);
    }
    if (!Number.isFinite(qLower) || !Number.isFinite(qUpper)) {
      throw new Error(`RLP原始值无效：β=${beta}, n=${n}, n1=${n1}, ρ=${rho}, b=${b}`);
    }
    lower += weight * qLower;
    upper += weight * qUpper;
  }
  return { lower, upper };
}

for (const row of singleRows) {
  const bounds = getSystemBounds(row.n, row.n1, row.beta, row.rho, row.p);
  row.psLower = bounds.lower;
  row.psUpper = bounds.upper;
}
for (const row of multiRows) {
  const bounds = getSystemBounds(row.n, row.n1, row.beta, row.rho, row.p);
  row.psLower = bounds.lower;
  row.psUpper = bounds.upper;
}

const keyOf = (r) => [r.n, r.n1, r.beta.toFixed(1), r.rho.toFixed(1), r.p.toFixed(1), r.k].join("|");
const firstSeen = new Map();
let uniqueCounter = 0;
for (const row of [...singleRows, ...multiRows]) {
  const key = keyOf(row);
  if (!firstSeen.has(key)) {
    uniqueCounter += 1;
    firstSeen.set(key, { calcId: `SER-U${String(uniqueCounter).padStart(3, "0")}`, displayId: row.displayId });
    row.calcId = firstSeen.get(key).calcId;
    row.duplicate = "首现";
  } else {
    const first = firstSeen.get(key);
    row.calcId = first.calcId;
    row.duplicate = `复用${first.displayId}`;
  }
}

const targetLabel = (n1) => n1 === 1 ? "唯一潜在脆性构件i" : "潜在脆性构件之一i（固定同一编号）";
const originalB = (n1) => Array.from({ length: n1 + 1 }, (_, i) => i).join("、");
const reliableB = (n1) => Array.from({ length: n1 }, (_, i) => i).join("、");
const handling = "保留构件i及其位置、刚度和传力作用；关闭i的全部脆性与塑性失效事件";

const colors = {
  navy: "#1F4E78",
  teal: "#0F6B78",
  lightBlue: "#DDEBF7",
  lightYellow: "#FFF2CC",
  lightGreen: "#E2F0D9",
  lightGray: "#F2F2F2",
  border: "#B8C2CC",
  text: "#1F2937",
  white: "#FFFFFF",
};
const font = "Arial";

function styleSheetBase(sheet, lastColumn, lastRow) {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${lastColumn}${lastRow}`).format.font = { name: font, size: 10, color: colors.text };
  sheet.getRange(`A1:${lastColumn}${lastRow}`).format.verticalAlignment = "center";
  sheet.getRange("A2").format.font = { name: font, size: 15, bold: true, color: colors.navy };
  sheet.getRange(`A3:${lastColumn}3`).format.borders = { bottom: { style: "thin", color: colors.navy } };
  sheet.getRange(`A4:${lastColumn}6`).format.font = { name: font, size: 10, color: "#4B5563" };
  sheet.getRange(`A4:${lastColumn}6`).format.rowHeight = 22;
  sheet.getRange(`A7:${lastColumn}7`).format = {
    fill: colors.navy,
    font: { name: font, size: 10, bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: colors.white },
  };
  sheet.getRange(`A7:${lastColumn}7`).format.rowHeight = 44;
  sheet.getRange(`A8:${lastColumn}${lastRow}`).format.borders = {
    insideHorizontal: { style: "thin", color: "#D9E1E8" },
    bottom: { style: "thin", color: colors.border },
  };
  sheet.freezePanes.freezeRows(7);
}

function writeSingleWorkbook() {
  const wb = Workbook.create();
  const sheet = wb.worksheets.add("单参数计算需求");
  sheet.tabColor = colors.navy;
  sheet.getRange("A2").values = [["串联系统RRW单参数计算需求表"]];
  sheet.getRange("A4").values = [["计算口径：Pbi=0.2，K_RLP=3；目标构件i保留在系统中并设为完全可靠。"]];
  sheet.getRange("A5").values = [["Pf(S)上下界已按原始RLP条件概率进行二项加权回填；后续仅需填写Pf(S_i+)上下界。"]];
  sheet.getRange("A6").values = [["β组仅保留存在n=10原始RLP数据的β=3.0、3.2、3.7；β=2.7原始文件无10根构件表，未纳入。"]];

  const headers = [[
    "单参数项目", "变化参数取值", "构件总数 n", "潜在脆性构件数 n1", "可靠指标 β", "相关系数 ρ", "脆性发生概率 Pbi", "K_RLP", "目标构件",
    "Pf(S) 下界", "Pf(S) 上界", "Pf(S_i+) 下界", "Pf(S_i+) 上界",
    "RRW 下界", "RRW 上界", "RRW 区间中点", "返回状态", "参数说明",
  ]];
  sheet.getRange("A7:R7").values = headers;

  const rows = singleRows.map((r) => [
    r.group, r.changeValue, r.n, r.n1, r.beta, r.rho, r.p, r.k, targetLabel(r.n1),
    r.psLower, r.psUpper, null, null, null, null, null, null, r.note,
  ]);
  sheet.getRange(`A8:R${7 + rows.length}`).values = rows;
  for (let row = 8; row <= 7 + rows.length; row += 1) {
    sheet.getRange(`N${row}`).formulas = [[`=IF(OR(J${row}="",M${row}="",M${row}=0),"",J${row}/M${row})`]];
    sheet.getRange(`O${row}`).formulas = [[`=IF(OR(K${row}="",L${row}="",L${row}=0),"",K${row}/L${row})`]];
    sheet.getRange(`P${row}`).formulas = [[`=IF(OR(N${row}="",O${row}=""),"",AVERAGE(N${row}:O${row}))`]];
    sheet.getRange(`Q${row}`).formulas = [[`=IF(COUNTA(J${row}:M${row})=0,"待返回",IF(COUNTA(J${row}:M${row})<4,"待补充",IF(OR(L${row}=0,M${row}=0),"检查分母","已返回")))`]];
  }

  const lastRow = 7 + rows.length;
  styleSheetBase(sheet, "R", lastRow);
  sheet.getRange(`A8:I${lastRow}`).format.fill = "#FFFFFF";
  sheet.getRange(`J8:K${lastRow}`).format.fill = colors.lightBlue;
  sheet.getRange(`L8:M${lastRow}`).format.fill = colors.lightYellow;
  sheet.getRange(`N8:Q${lastRow}`).format.fill = colors.lightGreen;
  sheet.getRange(`R8:R${lastRow}`).format.fill = colors.lightGray;
  sheet.getRange(`A8:H${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`J8:P${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`Q8:Q${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`E8:G${lastRow}`).format.numberFormat = "0.0";
  sheet.getRange(`J8:P${lastRow}`).format.numberFormat = "0.000000000000";
  sheet.getRange(`A8:R${lastRow}`).format.rowHeight = 27;
  sheet.getRange(`I8:I${lastRow}`).format.wrapText = true;
  sheet.getRange(`R8:R${lastRow}`).format.wrapText = true;

  const widths = {
    A: 22, B: 14, C: 12, D: 17, E: 12, F: 12, G: 18, H: 10, I: 28,
    J: 16, K: 16, L: 18, M: 18, N: 15, O: 15, P: 18, Q: 13, R: 48,
  };
  for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  sheet.tables.add(`A7:R${lastRow}`, true, "SingleParameterRequests").style = "TableStyleMedium2";
  return wb;
}

function writeMultiWorkbook() {
  const wb = Workbook.create();
  const sheet = wb.worksheets.add("多参数计算需求");
  sheet.tabColor = colors.navy;
  sheet.getRange("A2").values = [["串联系统RRW多参数计算需求表"]];
  sheet.getRange("A4").values = [["计算口径：ρ×n1固定n=8；n×n1固定ρ=0.3；n×ρ固定n1=1。统一β=3.0、Pbi=0.2、K_RLP=3。"]];
  sheet.getRange("A5").values = [["Pf(S)上下界已按原始RLP条件概率进行二项加权回填；后续仅需填写Pf(S_i+)上下界。"]];
  sheet.getRange("A6").values = [["参数依据：吴仁彬硕士论文第2.3.2节及当前确认的103点复现网格。蓝色列为已回填Pf(S)，黄色列待填写Pf(S_i+)。"]];

  const headers = [[
    "双参数组合", "参数1", "参数1取值", "参数2", "参数2取值", "构件总数 n", "潜在脆性构件数 n1", "可靠指标 β", "相关系数 ρ", "脆性发生概率 Pbi", "K_RLP", "目标构件",
    "Pf(S) 下界", "Pf(S) 上界", "Pf(S_i+) 下界", "Pf(S_i+) 上界",
    "RRW 下界", "RRW 上界", "RRW 区间中点", "返回状态", "参数说明",
  ]];
  sheet.getRange("A7:U7").values = headers;
  const rows = multiRows.map((r) => [
    r.group, r.parameter1, r.parameter1Value, r.parameter2, r.parameter2Value,
    r.n, r.n1, r.beta, r.rho, r.p, r.k, targetLabel(r.n1),
    r.psLower, r.psUpper, null, null, null, null, null, null, r.note,
  ]);
  sheet.getRange(`A8:U${7 + rows.length}`).values = rows;
  for (let row = 8; row <= 7 + rows.length; row += 1) {
    sheet.getRange(`Q${row}`).formulas = [[`=IF(OR(M${row}="",P${row}="",P${row}=0),"",M${row}/P${row})`]];
    sheet.getRange(`R${row}`).formulas = [[`=IF(OR(N${row}="",O${row}="",O${row}=0),"",N${row}/O${row})`]];
    sheet.getRange(`S${row}`).formulas = [[`=IF(OR(Q${row}="",R${row}=""),"",AVERAGE(Q${row}:R${row}))`]];
    sheet.getRange(`T${row}`).formulas = [[`=IF(COUNTA(M${row}:P${row})=0,"待返回",IF(COUNTA(M${row}:P${row})<4,"待补充",IF(OR(O${row}=0,P${row}=0),"检查分母","已返回")))`]];
  }

  const lastRow = 7 + rows.length;
  styleSheetBase(sheet, "U", lastRow);
  sheet.getRange(`A8:L${lastRow}`).format.fill = "#FFFFFF";
  sheet.getRange(`M8:N${lastRow}`).format.fill = colors.lightBlue;
  sheet.getRange(`O8:P${lastRow}`).format.fill = colors.lightYellow;
  sheet.getRange(`Q8:T${lastRow}`).format.fill = colors.lightGreen;
  sheet.getRange(`U8:U${lastRow}`).format.fill = colors.lightGray;
  sheet.getRange(`A8:K${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`M8:S${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`T8:T${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`H8:J${lastRow}`).format.numberFormat = "0.0";
  sheet.getRange(`M8:S${lastRow}`).format.numberFormat = "0.000000000000";
  sheet.getRange(`A8:U${lastRow}`).format.rowHeight = 27;
  sheet.getRange(`L8:L${lastRow}`).format.wrapText = true;
  sheet.getRange(`U8:U${lastRow}`).format.wrapText = true;

  const widths = {
    A: 15, B: 18, C: 13, D: 22, E: 13, F: 12, G: 17, H: 12, I: 12, J: 18, K: 10, L: 28,
    M: 16, N: 16, O: 18, P: 18, Q: 15, R: 15, S: 18, T: 13, U: 45,
  };
  for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  sheet.tables.add(`A7:U${lastRow}`, true, "MultiParameterRequests").style = "TableStyleMedium2";
  return wb;
}

const singleWb = writeSingleWorkbook();
const multiWb = writeMultiWorkbook();

singleWb.recalculate();
multiWb.recalculate();

const singleCheck = await singleWb.inspect({ kind: "table", range: "单参数计算需求!A1:R29", include: "values,formulas", tableMaxRows: 29, tableMaxCols: 18, maxChars: 18000 });
const multiCheckTop = await multiWb.inspect({ kind: "table", range: "多参数计算需求!A1:U16", include: "values,formulas", tableMaxRows: 16, tableMaxCols: 21, maxChars: 10000 });
const multiCheckBottom = await multiWb.inspect({ kind: "table", range: "多参数计算需求!A100:U110", include: "values,formulas", tableMaxRows: 11, tableMaxCols: 21, maxChars: 7000 });
const singleErrors = await singleWb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "single formula errors" });
const multiErrors = await multiWb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "multi formula errors" });

console.log(`COUNTS single=${singleRows.length}; multi=${multiRows.length}; globalUnique=${firstSeen.size}; multiUnique=${new Set(multiRows.map(keyOf)).size}`);
console.log(singleCheck.ndjson);
console.log(multiCheckTop.ndjson);
console.log(multiCheckBottom.ndjson);
console.log(singleErrors.ndjson);
console.log(multiErrors.ndjson);

const singlePreview = await singleWb.render({ sheetName: "单参数计算需求", range: "A1:R29", scale: 1, format: "png" });
const multiPreviewTop = await multiWb.render({ sheetName: "多参数计算需求", range: "A1:U28", scale: 1, format: "png" });
const multiPreviewBottom = await multiWb.render({ sheetName: "多参数计算需求", range: "A90:U110", scale: 1, format: "png" });
await fs.writeFile(path.join(previewDir, "single.png"), new Uint8Array(await singlePreview.arrayBuffer()));
await fs.writeFile(path.join(previewDir, "multi_top.png"), new Uint8Array(await multiPreviewTop.arrayBuffer()));
await fs.writeFile(path.join(previewDir, "multi_bottom.png"), new Uint8Array(await multiPreviewBottom.arrayBuffer()));

// Keep only the header rows frozen. Avoid freezing columns because Excel can
// persist a shifted horizontal viewport and make a middle column look missing.
for (const sheet of [
  singleWb.worksheets.getItem("单参数计算需求"),
  multiWb.worksheets.getItem("多参数计算需求"),
]) {
  sheet.freezePanes.unfreeze();
  sheet.freezePanes.freezeRows(7);
}

await (await SpreadsheetFile.exportXlsx(singleWb)).save(singleOutput);
await (await SpreadsheetFile.exportXlsx(multiWb)).save(multiOutput);

console.log(`OUTPUT_SINGLE=${singleOutput}`);
console.log(`OUTPUT_MULTI=${multiOutput}`);
