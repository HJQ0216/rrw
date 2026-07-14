param(
  [string]$OutputPath = '.\串联系统_RLP_RRW导师计算任务单.xlsx',
  [switch]$SingleParameterOnly,
  [switch]$MultiParameterOnly
)

if ($SingleParameterOnly -and $MultiParameterOnly) { throw 'SingleParameterOnly and MultiParameterOnly cannot be used together.' }

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function XmlEscape([object]$Value) {
  if ($null -eq $Value) { return '' }
  return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ColName([int]$Index) {
  $name = ''
  while ($Index -gt 0) {
    $Index--
    $name = [char](65 + ($Index % 26)) + $name
    $Index = [Math]::Floor($Index / 26)
  }
  return $name
}

function CellXml([string]$Ref, [object]$Value, [int]$Style) {
  if ($Value -is [string] -and $Value.StartsWith('=')) {
    $f = XmlEscape $Value.Substring(1)
    return "<c r=`"$Ref`" s=`"$Style`"><f>$f</f></c>"
  }
  if ($null -eq $Value -or [string]$Value -eq '') {
    return "<c r=`"$Ref`" s=`"$Style`"/>"
  }
  if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or
      $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
    $num = [Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
    return "<c r=`"$Ref`" s=`"$Style`"><v>$num</v></c>"
  }
  $text = XmlEscape $Value
  return "<c r=`"$Ref`" s=`"$Style`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$text</t></is></c>"
}

function StyleFor([string]$Group, [object]$Value) {
  $numeric = $Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or
             $Value -is [single] -or $Value -is [double] -or $Value -is [decimal] -or
             ($Value -is [string] -and $Value.StartsWith('='))
  switch ($Group) {
    'output' { if ($numeric) { return 7 } else { return 8 } }
    'calc'   { if ($numeric) { return 9 } else { return 10 } }
    'note'   { return 11 }
    'warn'   { return 12 }
    default  { if ($numeric) { return 6 } else { return 5 } }
  }
}

function New-TableSheetXml($Sheet) {
  $cols = $Sheet.Columns
  $rows = $Sheet.Rows
  $lastCol = ColName $cols.Count
  $lastRow = $rows.Count + 2
  $sb = [System.Text.StringBuilder]::new()
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
  [void]$sb.Append("<dimension ref=`"A1:$lastCol$lastRow`"/>")
  [void]$sb.Append('<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
  [void]$sb.Append('<sheetFormatPr defaultRowHeight="18"/>')
  [void]$sb.Append('<cols>')
  for ($c=0; $c -lt $cols.Count; $c++) {
    $idx = $c + 1
    $width = $cols[$c].Width
    [void]$sb.Append("<col min=`"$idx`" max=`"$idx`" width=`"$width`" customWidth=`"1`"/>")
  }
  [void]$sb.Append('</cols><sheetData>')

  [void]$sb.Append('<row r="1" ht="30" customHeight="1">')
  [void]$sb.Append((CellXml 'A1' $Sheet.Title 1))
  [void]$sb.Append('</row>')

  [void]$sb.Append('<row r="2" ht="38" customHeight="1">')
  for ($c=0; $c -lt $cols.Count; $c++) {
    $ref = "$(ColName ($c+1))2"
    $group = $cols[$c].Group
    $style = if ($group -eq 'output') { 3 } elseif ($group -eq 'calc') { 4 } else { 2 }
    [void]$sb.Append((CellXml $ref $cols[$c].Header $style))
  }
  [void]$sb.Append('</row>')

  for ($r=0; $r -lt $rows.Count; $r++) {
    $rowNum = $r + 3
    [void]$sb.Append("<row r=`"$rowNum`" ht=`"31`" customHeight=`"1`">")
    for ($c=0; $c -lt $cols.Count; $c++) {
      $col = $cols[$c]
      $value = $rows[$r][$col.Key]
      $style = StyleFor $col.Group $value
      $ref = "$(ColName ($c+1))$rowNum"
      [void]$sb.Append((CellXml $ref $value $style))
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData>')
  [void]$sb.Append("<autoFilter ref=`"A2:$lastCol$lastRow`"/>")
  [void]$sb.Append("<mergeCells count=`"1`"><mergeCell ref=`"A1:$lastCol`1`"/></mergeCells>")
  [void]$sb.Append('<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>')
  [void]$sb.Append('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
  [void]$sb.Append('</worksheet>')
  return $sb.ToString()
}

function MainColumns {
  return @(
    @{Key='case_id';Header='工况ID';Width=13;Group='input'},
    @{Key='group';Header='工况组';Width=17;Group='input'},
    @{Key='var';Header='变化参数';Width=14;Group='input'},
    @{Key='n';Header='构件总数 n';Width=11;Group='input'},
    @{Key='n1';Header='潜在脆性数 n1';Width=13;Group='input'},
    @{Key='beta_other';Header='β 其他构件';Width=12;Group='input'},
    @{Key='beta_d';Header='βd 目标塑性';Width=13;Group='input'},
    @{Key='beta_b';Header='βb 目标脆性';Width=15;Group='input'},
    @{Key='pb';Header='脆性概率 Pb';Width=12;Group='input'},
    @{Key='rho';Header='相关系数 ρ';Width=11;Group='input'},
    @{Key='rho_def';Header='ρ定义/相关矩阵';Width=24;Group='input'},
    @{Key='target';Header='评价对象';Width=25;Group='input'},
    @{Key='model_note';Header='模型与荷载说明';Width=35;Group='input'},
    @{Key='pf_s_l';Header='Pf(S) 下界';Width=14;Group='output'},
    @{Key='pf_s_u';Header='Pf(S) 上界';Width=14;Group='output'},
    @{Key='pf_rm_l';Header='Pf(S_-i) 下界';Width=16;Group='output'},
    @{Key='pf_rm_u';Header='Pf(S_-i) 上界';Width=16;Group='output'},
    @{Key='pf_plus_l';Header='Pf(S_i+) 下界';Width=16;Group='output'},
    @{Key='pf_plus_u';Header='Pf(S_i+) 上界';Width=16;Group='output'},
    @{Key='lp_status';Header='LP状态';Width=13;Group='output'},
    @{Key='constraint_error';Header='最大约束残差';Width=15;Group='output'},
    @{Key='teacher_note';Header='导师备注';Width=25;Group='output'},
    @{Key='pf_s_mean';Header='Pf(S) 均值';Width=14;Group='calc'},
    @{Key='pf_rm_mean';Header='Pf(S_-i) 均值';Width=16;Group='calc'},
    @{Key='pf_plus_mean';Header='Pf(S_i+) 均值';Width=16;Group='calc'},
    @{Key='raw_l';Header='论文RAW 下界';Width=15;Group='calc'},
    @{Key='raw_u';Header='论文RAW 上界';Width=15;Group='calc'},
    @{Key='raw_mean';Header='论文RAW 均值';Width=15;Group='calc'},
    @{Key='rrw_inverse';Header='RAW倒数';Width=14;Group='calc'},
    @{Key='rrw_l';Header='标准RRW 下界';Width=15;Group='calc'},
    @{Key='rrw_u';Header='标准RRW 上界';Width=15;Group='calc'},
    @{Key='rrw_mean';Header='标准RRW 均值';Width=15;Group='calc'},
    @{Key='rrw_diff';Header='标准RRW-RAW倒数';Width=19;Group='calc'},
    @{Key='check';Header='自动校核';Width=18;Group='calc'}
  )
}

function AddFormulas([hashtable]$Row, [int]$ExcelRow) {
  $Row.pf_s_mean = "=IF(COUNT(N${ExcelRow}:O${ExcelRow})=2,AVERAGE(N${ExcelRow}:O${ExcelRow}),`"`" )"
  $Row.pf_rm_mean = "=IF(COUNT(P${ExcelRow}:Q${ExcelRow})=2,AVERAGE(P${ExcelRow}:Q${ExcelRow}),`"`" )"
  $Row.pf_plus_mean = "=IF(COUNT(R${ExcelRow}:S${ExcelRow})=2,AVERAGE(R${ExcelRow}:S${ExcelRow}),`"`" )"
  $Row.raw_l = "=IFERROR(P$ExcelRow/O$ExcelRow,`"`")"
  $Row.raw_u = "=IFERROR(Q$ExcelRow/N$ExcelRow,`"`")"
  # 论文式(2.6)：RAW平均值为RAW上下界的算术平均，不是概率均值之比。
  $Row.raw_mean = "=IF(COUNT(Z${ExcelRow}:AA${ExcelRow})=2,AVERAGE(Z${ExcelRow}:AA${ExcelRow}),`"`")"
  $Row.rrw_inverse = "=IFERROR(1/AB$ExcelRow,`"`")"
  $Row.rrw_l = "=IFERROR(N$ExcelRow/S$ExcelRow,`"`")"
  $Row.rrw_u = "=IFERROR(O$ExcelRow/R$ExcelRow,`"`")"
  # 标准RRW代表值采用RRW上下界的算术平均，并在研究方法中明确为本文口径。
  $Row.rrw_mean = "=IF(COUNT(AD${ExcelRow}:AE${ExcelRow})=2,AVERAGE(AD${ExcelRow}:AE${ExcelRow}),`"`")"
  $Row.rrw_diff = "=IFERROR(AF$ExcelRow-AC$ExcelRow,`"`")"
  $Row.check = "=IF(COUNTA(N${ExcelRow}:S${ExcelRow})<6,`"待导师填写`",IF(AND(N$ExcelRow<=O$ExcelRow,P$ExcelRow<=Q$ExcelRow,R$ExcelRow<=S$ExcelRow,AF$ExcelRow>=1),`"通过`",`"检查输入/模型`"))"
}

function NewMainRow([string]$Id,[string]$Group,[string]$Var,[object]$n,[object]$n1,[object]$BetaOther,[object]$BetaD,[object]$BetaB,[object]$Pb,[object]$Rho,[string]$Target,[string]$Note) {
  $row = @{
    case_id=$Id; group=$Group; var=$Var; n=$n; n1=$n1; beta_other=$BetaOther; beta_d=$BetaD; beta_b=$BetaB;
    pb=$Pb; rho=$Rho; rho_def='等相关；具体含义按导师原RLP程序确认'; target=$Target; model_note=$Note;
    pf_s_l=$null;pf_s_u=$null;pf_rm_l=$null;pf_rm_u=$null;pf_plus_l=$null;pf_plus_u=$null;
    lp_status='';constraint_error=$null;teacher_note='';
  }
  return $row
}

function FinalizeRows($Rows) {
  for ($i=0; $i -lt $Rows.Count; $i++) { AddFormulas $Rows[$i] ($i+3) }
  return $Rows
}

$sheets = [System.Collections.Generic.List[object]]::new()

$infoCols = @(
  @{Key='section';Header='类别';Width=18;Group='input'},
  @{Key='item';Header='项目';Width=30;Group='input'},
  @{Key='content';Header='具体要求';Width=95;Group='note'},
  @{Key='status';Header='确认状态';Width=18;Group='output'}
)
$infoRows = @(
  @{section='研究范围';item='第一阶段系统';content='仅计算串联系统／静定钢结构，与论文串联系统数据逐项对比；理想并联系统和蒙特卡洛模拟暂不纳入本批任务。';status='已确定'},
  @{section='指标定义';item='论文RAW';content='RAW_文章 = Pf(S_-i) / Pf(S)，其中S_-i为论文中的移除系统。';status='已确定'},
  @{section='指标定义';item='标准RRW';content='RRW_标准 = Pf(S) / Pf(S_i+)，S_i+保留构件、刚度和传力路径，只令评价对象完全可靠。该定义已由用户确认。';status='定义已确定'},
  @{section='三种模型';item='原系统 S';content='原构件数量、可靠指标、相关关系、荷载与结构模型保持不变。';status='必算'},
  @{section='三种模型';item='移除系统 S_-i';content='按论文方法移除评价对象，用于复现论文RAW。';status='必算'},
  @{section='三种模型';item='完全可靠系统 S_i+';content='不得删除评价对象；保持原结构受力和刚度，只关闭目标构件的塑性、脆性全部失效事件。程序实现优先删除目标失效随机事件，而不是输入有限的大β。';status='定义已确定／必算'},
  @{section='三种模型';item='S_i+随机维数';content='物理构件数量仍为n；参与随机失效计算的构件事件数为n-1。相关矩阵中删除目标失效事件对应的行列，其余相关关系保持不变。';status='定义已确定'},
  @{section='颜色说明';item='黄色列';content='你提交给导师的输入参数及模型说明。';status='填写后提交'},
  @{section='颜色说明';item='绿色列';content='导师运行RLP后填写的失效概率上下界、LP状态和约束残差。';status='导师返回'},
  @{section='颜色说明';item='蓝色列';content='Excel自动计算RAW、RAW倒数、标准RRW及校核结果，请勿手工覆盖公式。';status='自动计算'},
  @{section='未决问题';item='βb取值';content='论文通用参数分析只明确给出β，未明确βb转换规则；βb列留空，由导师按原程序补充或确认。';status='必须确认'},
  @{section='未决问题';item='ρ的含义';content='必须确认ρ是极限状态变量、抗力、荷载还是失效事件相关系数，以及程序是否直接接收相关矩阵。';status='必须确认'},
  @{section='未决问题';item='评价对象';content='理论主线建议评价单根潜在脆性构件；四杆基本单元另算，用于严格复现论文工程算例。';status='必须区分'}
)
$sheets.Add(@{Name='00_说明';Title='串联系统RLP／RRW导师计算任务单——使用说明';Columns=$infoCols;Rows=$infoRows})

$firstRows = [System.Collections.Generic.List[hashtable]]::new()
$firstRows.Add((NewMainRow 'B01' '通用基准' '基准' 8 1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 0.1 'PB1潜在脆性构件i' '同论文串联系统通用基准；同时计算S、S_-i与S_i+'))
$firstRows.Add((NewMainRow 'B02-S' '11杆桁架' '单构件口径' 11 1 3.56 3.56 3.36 0.05 0.1 '11号潜在脆性杆' 'F=100 kN；保留11杆结构；S_-i仅移除11号杆；S_i+令11号杆完全可靠'))
$firstRows.Add((NewMainRow 'B02-G' '11杆桁架' '论文四杆单元口径' 11 1 3.56 3.56 3.36 0.05 0.1 '8～11号四杆基本单元' 'F=100 kN；S_-i为论文7杆系统；S_i+保留8～11号杆并令该单元完全可靠'))
$firstRows.Add((NewMainRow 'B03-S' '清水塘大桥' '单构件口径' 33 1 3.24 3.24 3.04 0.05 $null '6号潜在脆性杆' '局部33杆模型；ρ按导师原程序／论文原值填写；S_i+仅令6号杆完全可靠'))
$firstRows.Add((NewMainRow 'B03-G' '清水塘大桥' '论文四杆单元口径' 33 1 3.24 3.24 3.04 0.05 $null '5～8号四杆基本单元' '严格复现论文RAW；S_i+保留5～8号杆并令该单元完全可靠'))
$sheets.Add(@{Name='01_首批任务';Title='第一批提交导师的RLP计算任务（先复现，再批量）';Columns=(MainColumns);Rows=(FinalizeRows $firstRows)})

$singleRows = [System.Collections.Generic.List[hashtable]]::new()
$idx=1
foreach($rho in @(0.1,0.3,0.5,0.7,0.9)) {
  $singleRows.Add((NewMainRow ("S-RHO-{0:D2}" -f $idx) '单参数' '相关系数ρ' 8 1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 $rho 'PB1潜在脆性构件i' '固定n=8、n1=1、β=3.0、Pb=0.2；按论文图5/图2.6'))
  $idx++
}
$idx=1
foreach($n in @(8,12,16,20,24,28,32,36,40,44,48)) {
  $singleRows.Add((NewMainRow ("S-N-{0:D2}" -f $idx) '单参数' '构件总数n' $n 1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 0.1 'PB1潜在脆性构件i' '固定n1=1、β=3.0、ρ=0.1、Pb=0.2；按论文图6/图2.7'))
  $idx++
}
$idx=1
foreach($beta in @(3.0,3.1,3.2,3.3,3.4,3.5,3.6)) {
  $singleRows.Add((NewMainRow ("S-BETA-{0:D2}" -f $idx) '单参数' '可靠指标β' 10 1 $beta $beta '论文未单列；按原RLP程序确认' 0.2 0.1 'PB1潜在脆性构件i' '固定n=10、n1=1、ρ=0.1、Pb=0.2；按论文图7/图2.8'))
  $idx++
}
$idx=1
foreach($n1 in @(1,2,3)) {
  $singleRows.Add((NewMainRow ("S-N1-{0:D2}" -f $idx) '单参数' '潜在脆性数n1' 8 $n1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 0.1 'PB1为目标，其余潜在脆性构件保留' '固定n=8、β=3.0、ρ=0.1、Pb=0.2；按论文图8/图2.9'))
  $idx++
}
$sheets.Add(@{Name='02_单参数工况';Title='论文串联系统单参数对比工况（共26组）';Columns=(MainColumns);Rows=(FinalizeRows $singleRows)})

$doubleRows = [System.Collections.Generic.List[hashtable]]::new()
$idx=1
foreach($n1 in @(1,2,3)) { foreach($rho in @(0.1,0.3,0.5,0.7,0.9)) {
  $doubleRows.Add((NewMainRow ("D-RHO-N1-{0:D3}" -f $idx) '双参数' 'ρ×n1' 8 $n1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 $rho 'PB1为目标，其余潜在脆性构件保留' '固定n=8、β=3.0、Pb=0.2；按论文图9/图2.10'))
  $idx++
}}
$idx=1
foreach($n1 in @(1,2,3)) { foreach($n in @(8,12,16,20,24,28,32,36,40,44,48)) {
  $doubleRows.Add((NewMainRow ("D-N-N1-{0:D3}" -f $idx) '双参数' 'n×n1' $n $n1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 0.3 'PB1为目标，其余潜在脆性构件保留' '固定ρ=0.3、β=3.0、Pb=0.2；按论文图10/图2.11'))
  $idx++
}}
$idx=1
foreach($rho in @(0.1,0.3,0.5,0.7,0.9)) { foreach($n in @(8,12,16,20,24,28,32,36,40,44,48)) {
  $doubleRows.Add((NewMainRow ("D-N-RHO-{0:D3}" -f $idx) '双参数' 'n×ρ' $n 1 3.0 3.0 '论文未单列；按原RLP程序确认' 0.2 $rho 'PB1潜在脆性构件i' '固定n1=1、β=3.0、Pb=0.2；按论文图11/图2.12'))
  $idx++
}}
$sheets.Add(@{Name='03_双参数工况';Title='论文串联系统双参数对比工况（共103组）';Columns=(MainColumns);Rows=(FinalizeRows $doubleRows)})

$engRows = [System.Collections.Generic.List[hashtable]]::new()
$engRows.Add((NewMainRow 'E11-S' '11杆桁架' '单构件RRW主线' 11 1 3.56 3.56 3.36 0.05 0.1 '11号杆' 'F=100 kN；比较保留且完全可靠的11号杆与移除11号杆'))
$engRows.Add((NewMainRow 'E11-G' '11杆桁架' '四杆单元论文复现' 11 1 3.56 3.56 3.36 0.05 0.1 '8～11号单元' '论文给出Pf(S)约0.00204、Pf(S_-i)约0.00130、RAW约0.6387'))
$engRows.Add((NewMainRow 'EQS-S' '清水塘大桥' '单构件RRW主线' 33 1 3.24 3.24 3.04 0.05 $null '6号杆' 'ρ待导师按原程序补充；只评价6号潜在脆性杆'))
$engRows.Add((NewMainRow 'EQS-G' '清水塘大桥' '四杆单元论文复现' 33 1 3.24 3.24 3.04 0.05 $null '5～8号单元' '论文给出Pf(S)=0.01917～0.01919、Pf(S_-i)=0.01689～0.01690、RAW=0.88056～0.88084'))
$sheets.Add(@{Name='04_工程算例';Title='工程算例：单构件主线与四杆单元论文复现必须分开';Columns=(MainColumns);Rows=(FinalizeRows $engRows)})

$stateCols = @(
  @{Key='template_id';Header='模板ID';Width=15;Group='input'},
  @{Key='n1';Header='n1';Width=8;Group='input'},
  @{Key='variant';Header='模型变体';Width=15;Group='input'},
  @{Key='state_id';Header='状态ID';Width=15;Group='input'},
  @{Key='target_state';Header='目标i状态';Width=15;Group='input'},
  @{Key='other_states';Header='其余潜在脆性构件状态';Width=30;Group='input'},
  @{Key='weight';Header='状态权重';Width=35;Group='input'},
  @{Key='pf_l';Header='条件Pf下界';Width=16;Group='output'},
  @{Key='pf_u';Header='条件Pf上界';Width=16;Group='output'},
  @{Key='lp';Header='LP状态';Width=14;Group='output'},
  @{Key='err';Header='约束残差';Width=14;Group='output'},
  @{Key='note';Header='备注';Width=35;Group='note'}
)
$stateRows = [System.Collections.Generic.List[hashtable]]::new()
foreach($n1 in 1..3) {
  $count=[Math]::Pow(2,$n1)
  for($mask=0;$mask -lt $count;$mask++) {
    $bits = [Convert]::ToString($mask,2).PadLeft($n1,'0')
    $targetState = if($bits[0] -eq '1'){'脆性b'}else{'塑性d'}
    $others = if($n1 -eq 1){'无'}else{($bits.Substring(1).ToCharArray() | ForEach-Object {if($_ -eq '1'){'b'}else{'d'}}) -join ','}
    $weightParts=@()
    for($j=0;$j -lt $n1;$j++){if($bits[$j] -eq '1'){$weightParts += "Pb$($j+1)"}else{$weightParts += "(1-Pb$($j+1))"}}
    $stateRows.Add(@{template_id="N$n1-O-$bits";n1=$n1;variant='original';state_id="Z$bits";target_state=$targetState;other_states=$others;weight=($weightParts -join '×');pf_l=$null;pf_u=$null;lp='';err=$null;note='原系统条件失效概率'})
  }
  $remaining=[Math]::Pow(2,$n1-1)
  foreach($variant in @('removed','perfect_i')) {
    for($mask=0;$mask -lt $remaining;$mask++) {
      $bits = if($n1 -eq 1){'-'}else{[Convert]::ToString($mask,2).PadLeft($n1-1,'0')}
      $others = if($n1 -eq 1){'无'}else{($bits.ToCharArray() | ForEach-Object {if($_ -eq '1'){'b'}else{'d'}}) -join ','}
      $weightParts=@()
      if($n1 -gt 1){for($j=0;$j -lt $n1-1;$j++){if($bits[$j] -eq '1'){$weightParts += "Pb$($j+2)"}else{$weightParts += "(1-Pb$($j+2))"}}}
      $weight = if($weightParts.Count){$weightParts -join '×'}else{'1'}
      $targetState = if($variant -eq 'removed'){'已移除'}else{'完全可靠'}
      $note = if($variant -eq 'removed'){'用于论文RAW'}else{'用于标准RRW'}
      $stateRows.Add(@{template_id="N$n1-$variant-$bits";n1=$n1;variant=$variant;state_id="Z$bits";target_state=$targetState;other_states=$others;weight=$weight;pf_l=$null;pf_u=$null;lp='';err=$null;note=$note})
    }
  }
}
$sheets.Add(@{Name='05_状态组合模板';Title='RLP逐状态返回模板（n1=1～3）';Columns=$stateCols;Rows=$stateRows})

$checkCols = @(
  @{Key='category';Header='类别';Width=18;Group='input'},
  @{Key='item';Header='提交/返回项目';Width=38;Group='input'},
  @{Key='required';Header='是否必需';Width=12;Group='input'},
  @{Key='source';Header='来源或填写人';Width=25;Group='input'},
  @{Key='detail';Header='要求';Width=75;Group='note'},
  @{Key='status';Header='状态';Width=18;Group='output'}
)
$checkRows = @(
  @{category='输入';item='系统类型、结构模型及构件编号';required='是';source='你';detail='本批只提交串联系统／静定结构；附模型图并标明目标构件。';status='待填写'},
  @{category='输入';item='n、n1和目标构件i';required='是';source='你';detail='单构件与四杆基本单元使用不同工况ID，严禁混为同一评价对象。';status='待填写'},
  @{category='输入';item='β其他、βd目标、βb目标';required='是';source='你＋导师确认';detail='βb通用取值规则未在论文参数段明确，批量计算前必须由原程序口径确认。';status='待确认'},
  @{category='输入';item='Pb';required='是';source='你';detail='论文参数分析为0.2；工程算例为0.05。多个潜在脆性构件时需说明是否相同。';status='待填写'},
  @{category='输入';item='ρ或完整相关矩阵';required='是';source='你＋导师确认';detail='必须写清ρ对应的随机变量，并检查相关矩阵半正定。';status='待确认'},
  @{category='输入';item='原系统S定义';required='是';source='你';detail='结构、荷载、构件数量及状态组合。';status='待填写'},
  @{category='输入';item='移除系统S_-i定义';required='是';source='你';detail='用于复现论文RAW，明确移除单杆还是四杆单元。';status='待填写'},
  @{category='输入';item='完全可靠系统S_i+定义';required='是';source='你；导师落实程序';detail='定义已确定：保留构件刚度、位置和传力路径，同时关闭目标构件全部失效事件。仍需导师确认原RLP程序的具体实现方法。';status='定义已确定'},
  @{category='返回';item='三种模型的Pf下界与上界';required='是';source='导师RLP程序';detail='最好逐脆性／塑性状态组合返回，不只给最终加权概率。';status='待返回'},
  @{category='返回';item='LP状态和最大约束残差';required='是';source='导师RLP程序';detail='用于识别不可行、未收敛和数值精度不足的工况。';status='待返回'},
  @{category='校核';item='论文RAW复现';required='是';source='Excel自动＋你';detail='先通过通用基准、11杆桁架和清水塘大桥，再做129组参数扫描。';status='待校核'},
  @{category='校核';item='标准RRW合理性';required='是';source='Excel自动＋你';detail='串联系统标准RRW一般不小于1；异常值需检查S_i+定义。';status='待校核'}
)
$sheets.Add(@{Name='06_提交清单';Title='提交导师前与导师返回后的核对清单';Columns=$checkCols;Rows=$checkRows})

if ($SingleParameterOnly) {
  $paperCols = @(
    @{Key='case_id';Header='工况ID';Width=14;Group='input'},
    @{Key='factor';Header='单参数项目';Width=20;Group='input'},
    @{Key='vary_value';Header='变化参数取值';Width=16;Group='input'},
    @{Key='n';Header='构件总数 n';Width=12;Group='input'},
    @{Key='n1';Header='潜在脆性构件数 n1';Width=17;Group='input'},
    @{Key='beta';Header='可靠指标 β';Width=13;Group='input'},
    @{Key='rho';Header='相关系数 ρ';Width=13;Group='input'},
    @{Key='pb';Header='脆性发生概率 Pbi';Width=17;Group='input'},
    @{Key='pf_s_l';Header='Pf(S) 下界';Width=15;Group='output'},
    @{Key='pf_s_u';Header='Pf(S) 上界';Width=15;Group='output'},
    @{Key='pf_plus_l';Header='Pf(S_i+) 下界';Width=17;Group='output'},
    @{Key='pf_plus_u';Header='Pf(S_i+) 上界';Width=17;Group='output'},
    @{Key='lp_status';Header='RLP状态';Width=14;Group='output'},
    @{Key='constraint_error';Header='最大约束残差';Width=16;Group='output'},
    @{Key='rrw_l';Header='RRW 下界';Width=14;Group='calc'},
    @{Key='rrw_u';Header='RRW 上界';Width=14;Group='calc'},
    @{Key='rrw_mean';Header='RRW 均值';Width=14;Group='calc'},
    @{Key='pf_rm_l';Header='Pf(S_-i) 下界';Width=17;Group='output'},
    @{Key='pf_rm_u';Header='Pf(S_-i) 上界';Width=17;Group='output'},
    @{Key='raw_l';Header='论文RAW 下界';Width=16;Group='calc'},
    @{Key='raw_u';Header='论文RAW 上界';Width=16;Group='calc'},
    @{Key='raw_mean';Header='论文RAW 均值';Width=16;Group='calc'},
    @{Key='check';Header='自动校核';Width=17;Group='calc'},
    @{Key='note';Header='论文参数设定';Width=48;Group='note'}
  )

  function NewPaperRow([string]$Id,[string]$Factor,[object]$VaryValue,[object]$n,[object]$n1,[object]$Beta,[object]$Rho,[object]$Pb,[string]$Note) {
    return @{
      case_id=$Id;factor=$Factor;vary_value=$VaryValue;n=$n;n1=$n1;beta=$Beta;rho=$Rho;pb=$Pb;
      pf_s_l=$null;pf_s_u=$null;pf_plus_l=$null;pf_plus_u=$null;lp_status='';constraint_error=$null;
      pf_rm_l=$null;pf_rm_u=$null;note=$Note
    }
  }

  $paperRows = [System.Collections.Generic.List[hashtable]]::new()
  $idx=1
  foreach($rho in @(0.1,0.3,0.5,0.7,0.9)) {
    $paperRows.Add((NewPaperRow ("SP-RHO-{0:D2}" -f $idx) '相关系数ρ' $rho 8 1 3.0 $rho 0.2 '固定n=8、n1=1、β=3.0、Pbi=0.2；对应毕业论文图2.6'))
    $idx++
  }
  $idx=1
  foreach($n in @(8,12,16,20,24,28,32,36,40,44,48)) {
    $paperRows.Add((NewPaperRow ("SP-N-{0:D2}" -f $idx) '构件总数n' $n $n 1 3.0 0.1 0.2 '固定n1=1、β=3.0、ρ=0.1、Pbi=0.2；对应毕业论文图2.7'))
    $idx++
  }
  $idx=1
  foreach($beta in @(3.0,3.1,3.2,3.3,3.4,3.5,3.6)) {
    $paperRows.Add((NewPaperRow ("SP-BETA-{0:D2}" -f $idx) '可靠指标β' $beta 10 1 $beta 0.1 0.2 '固定n=10、n1=1、ρ=0.1、Pbi=0.2；对应毕业论文图2.8'))
    $idx++
  }
  $idx=1
  foreach($n1 in @(1,2,3)) {
    $paperRows.Add((NewPaperRow ("SP-N1-{0:D2}" -f $idx) '潜在脆性构件数n1' $n1 8 $n1 3.0 0.1 0.2 '固定n=8、β=3.0、ρ=0.1、Pbi=0.2；对应毕业论文图2.9'))
    $idx++
  }

  for($i=0;$i -lt $paperRows.Count;$i++) {
    $r=$i+3
    $paperRows[$i].rrw_l="=IFERROR(I$r/L$r,`"`")"
    $paperRows[$i].rrw_u="=IFERROR(J$r/K$r,`"`")"
    $paperRows[$i].rrw_mean="=IF(COUNT(O${r}:P${r})=2,AVERAGE(O${r}:P${r}),`"`")"
    $paperRows[$i].raw_l="=IFERROR(R$r/J$r,`"`")"
    $paperRows[$i].raw_u="=IFERROR(S$r/I$r,`"`")"
    $paperRows[$i].raw_mean="=IF(COUNT(T${r}:U${r})=2,AVERAGE(T${r}:U${r}),`"`")"
    $paperRows[$i].check="=IF(COUNT(I${r}:L${r})+COUNT(R${r}:S${r})<6,`"待导师填写`",IF(AND(I$r>=0,J$r<=1,K$r>=0,L$r<=1,R$r>=0,S$r<=1,I$r<=J$r,K$r<=L$r,R$r<=S$r,O$r<=P$r,T$r<=U$r,Q$r>=1),`"通过`",`"检查`"))"
  }

  $sheets.Clear()
  $sheets.Add(@{Name='串联单参数';Title='毕业论文第2.3.1节：串联系统中单参数影响（严格按原文参数）';Columns=$paperCols;Rows=$paperRows})
}

if ($MultiParameterOnly) {
  $multiCols = @(
    @{Key='case_id';Header='工况ID';Width=16;Group='input'},
    @{Key='group';Header='双参数组合';Width=27;Group='input'},
    @{Key='factor1';Header='参数1';Width=16;Group='input'},
    @{Key='value1';Header='参数1取值';Width=14;Group='input'},
    @{Key='factor2';Header='参数2';Width=16;Group='input'},
    @{Key='value2';Header='参数2取值';Width=14;Group='input'},
    @{Key='n';Header='构件总数 n';Width=12;Group='input'},
    @{Key='n1';Header='潜在脆性构件数 n1';Width=17;Group='input'},
    @{Key='beta';Header='可靠指标 β';Width=13;Group='input'},
    @{Key='rho';Header='相关系数 ρ';Width=13;Group='input'},
    @{Key='pb';Header='脆性发生概率 Pbi';Width=17;Group='input'},
    @{Key='pf_s_l';Header='Pf(S) 下界';Width=15;Group='output'},
    @{Key='pf_s_u';Header='Pf(S) 上界';Width=15;Group='output'},
    @{Key='pf_plus_l';Header='Pf(S_i+) 下界';Width=17;Group='output'},
    @{Key='pf_plus_u';Header='Pf(S_i+) 上界';Width=17;Group='output'},
    @{Key='lp_status';Header='RLP状态';Width=14;Group='output'},
    @{Key='constraint_error';Header='最大约束残差';Width=16;Group='output'},
    @{Key='rrw_l';Header='RRW 下界';Width=14;Group='calc'},
    @{Key='rrw_u';Header='RRW 上界';Width=14;Group='calc'},
    @{Key='rrw_mean';Header='RRW 均值';Width=14;Group='calc'},
    @{Key='pf_rm_l';Header='Pf(S_-i) 下界';Width=17;Group='output'},
    @{Key='pf_rm_u';Header='Pf(S_-i) 上界';Width=17;Group='output'},
    @{Key='raw_l';Header='论文RAW 下界';Width=16;Group='calc'},
    @{Key='raw_u';Header='论文RAW 上界';Width=16;Group='calc'},
    @{Key='raw_mean';Header='论文RAW 均值';Width=16;Group='calc'},
    @{Key='check';Header='自动校核';Width=17;Group='calc'},
    @{Key='note';Header='论文参数设定';Width=56;Group='note'}
  )

  function NewMultiRow([string]$Id,[string]$Group,[string]$Factor1,[object]$Value1,[string]$Factor2,[object]$Value2,[object]$n,[object]$n1,[object]$Beta,[object]$Rho,[object]$Pb,[string]$Note) {
    return @{
      case_id=$Id;group=$Group;factor1=$Factor1;value1=$Value1;factor2=$Factor2;value2=$Value2;
      n=$n;n1=$n1;beta=$Beta;rho=$Rho;pb=$Pb;
      pf_s_l=$null;pf_s_u=$null;pf_plus_l=$null;pf_plus_u=$null;lp_status='';constraint_error=$null;
      pf_rm_l=$null;pf_rm_u=$null;note=$Note
    }
  }

  $multiRows = [System.Collections.Generic.List[hashtable]]::new()

  # 图2.10：ρ=0.1～0.9（步长0.1），n1=1、2、3；固定n=8、β=3.0、Pbi=0.2。
  $idx=1
  foreach($n1 in @(1,2,3)) {
    foreach($rho in @(0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9)) {
      $multiRows.Add((NewMultiRow ("MP-RHO-N1-{0:D3}" -f $idx) '相关系数ρ × 潜在脆性构件数n1' '相关系数ρ' $rho '潜在脆性构件数n1' $n1 8 $n1 3.0 $rho 0.2 '固定n=8、β=3.0、Pbi=0.2；ρ按图轴0.1步长；目标为PB1，S_i+保留PB1并令其完全可靠，其余n1-1个潜在脆性构件不变；图2.10'))
      $idx++
    }
  }

  # 图2.11：n=8～52（步长4），n1=1、2、3；固定ρ=0.3、β=3.0、Pbi=0.2。
  $idx=1
  foreach($n1 in @(1,2,3)) {
    foreach($n in @(8,12,16,20,24,28,32,36,40,44,48,52)) {
      $multiRows.Add((NewMultiRow ("MP-N-N1-{0:D3}" -f $idx) '构件总数n × 潜在脆性构件数n1' '构件总数n' $n '潜在脆性构件数n1' $n1 $n $n1 3.0 0.3 0.2 '固定ρ=0.3、β=3.0、Pbi=0.2；目标为PB1，S_i+保留PB1并令其完全可靠；第2.3.2(2)段末误写“相关系数”，以小标题、图题和图轴n×n1为准；图2.11'))
      $idx++
    }
  }

  # 图2.12：n=8～52（步长4），ρ=0.1～0.9（步长0.1）；固定n1=1、β=3.0、Pbi=0.2。
  $idx=1
  foreach($rho in @(0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9)) {
    foreach($n in @(8,12,16,20,24,28,32,36,40,44,48,52)) {
      $multiRows.Add((NewMultiRow ("MP-N-RHO-{0:D3}" -f $idx) '构件总数n × 相关系数ρ' '构件总数n' $n '相关系数ρ' $rho $n 1 3.0 $rho 0.2 '固定n1=1、β=3.0、Pbi=0.2；ρ按图轴0.1步长；目标为PB1，S_i+保留PB1并令其完全可靠；图2.12'))
      $idx++
    }
  }

  for($i=0;$i -lt $multiRows.Count;$i++) {
    $r=$i+3
    $multiRows[$i].rrw_l="=IFERROR(L$r/O$r,`"`")"
    $multiRows[$i].rrw_u="=IFERROR(M$r/N$r,`"`")"
    $multiRows[$i].rrw_mean="=IF(COUNT(R${r}:S${r})=2,AVERAGE(R${r}:S${r}),`"`")"
    $multiRows[$i].raw_l="=IFERROR(U$r/M$r,`"`")"
    $multiRows[$i].raw_u="=IFERROR(V$r/L$r,`"`")"
    $multiRows[$i].raw_mean="=IF(COUNT(W${r}:X${r})=2,AVERAGE(W${r}:X${r}),`"`")"
    $multiRows[$i].check="=IF(COUNT(L${r}:O${r})+COUNT(U${r}:V${r})<6,`"待导师填写`",IF(AND(L$r>=0,M$r<=1,N$r>=0,O$r<=1,U$r>=0,V$r<=1,L$r<=M$r,N$r<=O$r,U$r<=V$r,R$r<=S$r,W$r<=X$r,T$r>=1),`"通过`",`"检查`"))"
  }

  $sheets.Clear()
  $sheets.Add(@{Name='串联多参数';Title='毕业论文第2.3.2节：串联系统中多参数影响（严格按图2.10～图2.12参数网格）';Columns=$multiCols;Rows=$multiRows})
}

$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2"><numFmt numFmtId="164" formatCode="0.000000E+00"/><numFmt numFmtId="165" formatCode="0.000000"/></numFmts>
  <fonts count="5">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="16"/><name val="Calibri"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FF000000"/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FF9C0006"/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="10">
    <fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF548235"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF008C95"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="165" fontId="0" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="8" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="9" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@

$out = [System.IO.Path]::GetFullPath($OutputPath)
$workspace = [System.IO.Path]::GetFullPath((Get-Location).Path)
if (-not $out.StartsWith($workspace, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'Output must stay inside workspace.' }
if (Test-Path -LiteralPath $out) { throw "Output already exists: $out" }

$fs = [System.IO.File]::Open($out,[System.IO.FileMode]::CreateNew)
$zip = [System.IO.Compression.ZipArchive]::new($fs,[System.IO.Compression.ZipArchiveMode]::Create,$false)
function AddEntry([string]$Name,[string]$Content) {
  $entry=$zip.CreateEntry($Name,[System.IO.Compression.CompressionLevel]::Optimal)
  $stream=$entry.Open(); $writer=[System.IO.StreamWriter]::new($stream,[System.Text.UTF8Encoding]::new($false));
  $writer.Write($Content); $writer.Dispose(); $stream.Dispose()
}

$contentTypes = [System.Text.StringBuilder]::new()
[void]$contentTypes.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
for($i=1;$i -le $sheets.Count;$i++){[void]$contentTypes.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")}
[void]$contentTypes.Append('</Types>')
AddEntry '[Content_Types].xml' $contentTypes.ToString()
AddEntry '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'

$now=[DateTime]::UtcNow.ToString('s')+'Z'
AddEntry 'docProps/core.xml' "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><cp:coreProperties xmlns:cp=`"http://schemas.openxmlformats.org/package/2006/metadata/core-properties`" xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:dcterms=`"http://purl.org/dc/terms/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"><dc:title>串联系统RLP/RRW导师计算任务单</dc:title><dc:creator>Codex</dc:creator><dcterms:created xsi:type=`"dcterms:W3CDTF`">$now</dcterms:created><dcterms:modified xsi:type=`"dcterms:W3CDTF`">$now</dcterms:modified></cp:coreProperties>"
AddEntry 'docProps/app.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel Compatible</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop></Properties>'

$wb=[System.Text.StringBuilder]::new(); [void]$wb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>')
$wbRels=[System.Text.StringBuilder]::new(); [void]$wbRels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
for($i=0;$i -lt $sheets.Count;$i++){
  $id=$i+1; $name=XmlEscape $sheets[$i].Name
  [void]$wb.Append("<sheet name=`"$name`" sheetId=`"$id`" r:id=`"rId$id`"/>")
  [void]$wbRels.Append("<Relationship Id=`"rId$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$id.xml`"/>")
  AddEntry "xl/worksheets/sheet$id.xml" (New-TableSheetXml $sheets[$i])
}
$styleRid=$sheets.Count+1
[void]$wb.Append('</sheets><calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1" calcMode="auto"/></workbook>')
[void]$wbRels.Append("<Relationship Id=`"rId$styleRid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>")
AddEntry 'xl/workbook.xml' $wb.ToString()
AddEntry 'xl/_rels/workbook.xml.rels' $wbRels.ToString()
AddEntry 'xl/styles.xml' $styles

$zip.Dispose(); $fs.Dispose()
Write-Output $out
