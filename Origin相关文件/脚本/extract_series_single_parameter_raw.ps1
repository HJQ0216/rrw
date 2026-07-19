param(
    [string]$Workbook = 'D:\钢结构RRW方向\数据文件\串联系统单参数_RLP_RRW计算表.xlsx',
    [string]$Output = 'D:\钢结构RRW方向\Origin相关文件\源数据\串联系统单参数_RAW绘图数据.tsv'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function Read-Entry($zip, [string]$name) {
    $entry = $zip.GetEntry($name)
    if (-not $entry) { throw "Missing XLSX entry: $name" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-CellText($sheetXml, [string]$reference) {
    $cell = $sheetXml.worksheet.sheetData.row.c | Where-Object { $_.r -eq $reference }
    if (-not $cell -or [string]::IsNullOrWhiteSpace([string]$cell.v)) { return '' }
    if ([string]$cell.t -eq 's') {
        return ([string]$script:sharedStrings[[int]$cell.v]).Trim()
    }
    if ([string]$cell.t -eq 'inlineStr') { return ([string]$cell.is.InnerText).Trim() }
    ([string]$cell.v).Trim()
}

$stream = [IO.File]::Open($Workbook, 'Open', 'Read', 'ReadWrite')
$zip = [IO.Compression.ZipArchive]::new($stream, 'Read', $false)
$records = @()
try {
    $script:sharedStrings = @()
    $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
    if ($sharedEntry) {
        [xml]$sharedXml = Read-Entry $zip 'xl/sharedStrings.xml'
        foreach ($si in $sharedXml.sst.si) { $script:sharedStrings += [string]$si.InnerText }
    }
    [xml]$workbookXml = Read-Entry $zip 'xl/workbook.xml'
    [xml]$relsXml = Read-Entry $zip 'xl/_rels/workbook.xml.rels'
    $sheet = $workbookXml.workbook.sheets.sheet | Where-Object { $_.name -eq '串联单参数' }
    if (-not $sheet) { throw 'Worksheet not found: 串联单参数' }
    $rid = $sheet.GetAttribute('id', $relNs)
    $target = [string](($relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target)
    if ($target.StartsWith('/')) { $sheetEntry = $target.TrimStart('/') }
    elseif ($target.StartsWith('xl/')) { $sheetEntry = $target }
    else { $sheetEntry = 'xl/' + $target }
    [xml]$sheetXml = Read-Entry $zip $sheetEntry

    $groups = @(
        [pscustomobject]@{ Name='相关系数'; Symbol='ρ'; Start=3; End=7 },
        [pscustomobject]@{ Name='构件总数'; Symbol='n'; Start=8; End=18 },
        [pscustomobject]@{ Name='可靠指标'; Symbol='β'; Start=19; End=25 },
        [pscustomobject]@{ Name='潜在脆性构件数量'; Symbol='n1'; Start=26; End=28 }
    )
    foreach ($group in $groups) {
        for ($row = $group.Start; $row -le $group.End; $row++) {
            $x = Get-CellText $sheetXml "C$row"
            $lower = Get-CellText $sheetXml "T$row"
            $upper = Get-CellText $sheetXml "U$row"
            $mean = Get-CellText $sheetXml "V$row"
            $status = if ($lower -ne '' -and $upper -ne '' -and $mean -ne '') { '完整' } else { '缺少RAW' }
            $records += [pscustomobject]@{
                参数组 = $group.Name
                参数符号 = $group.Symbol
                原表行 = $row
                参数值 = $x
                RAW下界 = $lower
                RAW上界 = $upper
                RAW平均值 = $mean
                数据状态 = $status
            }
        }
    }
}
finally {
    $zip.Dispose()
    $stream.Dispose()
}

$outputDir = Split-Path -Parent $Output
if (-not (Test-Path -LiteralPath $outputDir)) { [void](New-Item -ItemType Directory -Path $outputDir) }
$records | Export-Csv -LiteralPath $Output -Delimiter "`t" -NoTypeInformation -Encoding UTF8
$records | Format-Table -AutoSize
