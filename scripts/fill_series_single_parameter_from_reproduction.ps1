param(
    [string]$SourceWorkbook = "D:\钢结构RRW方向\数据文件\RAW论文单参数复现与对比.xlsx",
    [string]$TargetWorkbook = "D:\钢结构RRW方向\数据文件\串联系统单参数_RLP_RRW计算表.xlsx",
    [string]$OutputWorkbook = "D:\钢结构RRW方向\数据文件\串联系统单参数_RLP_RRW计算表_更新中.xlsx"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

$SpreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$RelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function Read-ZipEntryText {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { throw "ZIP entry not found: $EntryName" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

function Get-WorksheetEntryName {
    param($WorkbookXml, $RelationshipsXml, [string]$SheetName)
    $sheet = $WorkbookXml.workbook.sheets.sheet |
        Where-Object { ([string]$_.name).Trim() -eq $SheetName.Trim() }
    if (-not $sheet) { throw "Worksheet not found: $SheetName" }
    $relationshipId = $sheet.GetAttribute('id', $RelationshipNamespace)
    $relationship = $RelationshipsXml.Relationships.Relationship |
        Where-Object { $_.Id -eq $relationshipId }
    if (-not $relationship) { throw "Worksheet relationship not found: $SheetName" }
    $target = [string]$relationship.Target
    if ($target.StartsWith('/')) {
        return $target.TrimStart('/')
    }
    if ($target.StartsWith('xl/')) {
        return $target
    }
    return 'xl/' + $target
}

function Get-CellNumber {
    param($WorksheetXml, [string]$CellReference)
    $cell = $WorksheetXml.worksheet.sheetData.row.c |
        Where-Object { $_.r -eq $CellReference }
    if (-not $cell -or [string]::IsNullOrWhiteSpace([string]$cell.v)) {
        return $null
    }
    return [double]::Parse(
        [string]$cell.v,
        [Globalization.CultureInfo]::InvariantCulture
    )
}

function Set-NumericCell {
    param($WorksheetXml, [string]$CellReference, [double]$Value)
    $cell = $WorksheetXml.worksheet.sheetData.row.c |
        Where-Object { $_.r -eq $CellReference }
    if (-not $cell) { throw "Target cell not found: $CellReference" }
    if ($cell.f) { throw "Refusing to replace formula cell: $CellReference" }
    if ($cell.HasAttribute('t')) { $cell.RemoveAttribute('t') }
    @($cell.v) | ForEach-Object {
        if ($_ -and $_.ParentNode) { [void]$cell.RemoveChild($_) }
    }
    $valueNode = $WorksheetXml.CreateElement('v', $SpreadsheetNamespace)
    $valueNode.InnerText = $Value.ToString(
        'G17',
        [Globalization.CultureInfo]::InvariantCulture
    )
    [void]$cell.AppendChild($valueNode)
}

function Replace-ZipXmlEntry {
    param($Zip, [string]$EntryName, $XmlDocument)
    $oldEntry = $Zip.GetEntry($EntryName)
    if (-not $oldEntry) { throw "ZIP entry not found for replacement: $EntryName" }
    $oldEntry.Delete()
    $newEntry = $Zip.CreateEntry($EntryName, [IO.Compression.CompressionLevel]::Optimal)
    $writer = [IO.StreamWriter]::new(
        $newEntry.Open(),
        [Text.UTF8Encoding]::new($false)
    )
    try {
        $XmlDocument.Save($writer)
    }
    finally {
        $writer.Dispose()
    }
}

if (-not (Test-Path -LiteralPath $SourceWorkbook)) {
    throw "Source workbook not found: $SourceWorkbook"
}
if (-not (Test-Path -LiteralPath $TargetWorkbook)) {
    throw "Target workbook not found: $TargetWorkbook"
}

$sourceStream = [IO.File]::Open(
    $SourceWorkbook,
    [IO.FileMode]::Open,
    [IO.FileAccess]::Read,
    [IO.FileShare]::ReadWrite
)
$sourceZip = [IO.Compression.ZipArchive]::new(
    $sourceStream,
    [IO.Compression.ZipArchiveMode]::Read,
    $false
)

$records = @()
try {
    [xml]$sourceWorkbookXml = Read-ZipEntryText $sourceZip 'xl/workbook.xml'
    [xml]$sourceRelationshipsXml = Read-ZipEntryText $sourceZip 'xl/_rels/workbook.xml.rels'

    $groups = @(
        [pscustomobject]@{ Sheet = '相关系数对比'; SourceStart = 4; TargetStart = 3; Count = 5 },
        [pscustomobject]@{ Sheet = '构件总数对比'; SourceStart = 4; TargetStart = 8; Count = 11 },
        [pscustomobject]@{ Sheet = '可靠度指标对比'; SourceStart = 4; TargetStart = 19; Count = 7 },
        [pscustomobject]@{ Sheet = '潜在脆性数量对比'; SourceStart = 4; TargetStart = 26; Count = 3 }
    )

    foreach ($group in $groups) {
        $entryName = Get-WorksheetEntryName `
            $sourceWorkbookXml $sourceRelationshipsXml $group.Sheet
        [xml]$worksheetXml = Read-ZipEntryText $sourceZip $entryName
        for ($offset = 0; $offset -lt $group.Count; $offset++) {
            $sourceRow = $group.SourceStart + $offset
            $psLower = Get-CellNumber $worksheetXml "G$sourceRow"
            $psUpper = Get-CellNumber $worksheetXml "H$sourceRow"
            $psNewLower = Get-CellNumber $worksheetXml "I$sourceRow"
            $psNewUpper = Get-CellNumber $worksheetXml "J$sourceRow"
            if ($null -eq $psLower -or $null -eq $psUpper -or
                $null -eq $psNewLower -or $null -eq $psNewUpper) {
                continue
            }
            $records += [pscustomobject]@{
                Group = $group.Sheet
                TargetRow = $group.TargetStart + $offset
                PsLower = $psLower
                PsUpper = $psUpper
                PsNewLower = $psNewLower
                PsNewUpper = $psNewUpper
            }
        }
    }
}
finally {
    $sourceZip.Dispose()
    $sourceStream.Dispose()
}

if ($records.Count -ne 20) {
    throw "Expected 20 complete reproduction records, found $($records.Count)."
}

[IO.File]::Copy($TargetWorkbook, $OutputWorkbook, $true)
$targetStream = [IO.File]::Open(
    $OutputWorkbook,
    [IO.FileMode]::Open,
    [IO.FileAccess]::ReadWrite,
    [IO.FileShare]::None
)
$targetZip = [IO.Compression.ZipArchive]::new(
    $targetStream,
    [IO.Compression.ZipArchiveMode]::Update,
    $false
)

try {
    [xml]$targetWorksheetXml = Read-ZipEntryText $targetZip 'xl/worksheets/sheet1.xml'
    foreach ($record in $records) {
        $row = $record.TargetRow
        Set-NumericCell $targetWorksheetXml "I$row" $record.PsLower
        Set-NumericCell $targetWorksheetXml "J$row" $record.PsUpper
        Set-NumericCell $targetWorksheetXml "R$row" $record.PsNewLower
        Set-NumericCell $targetWorksheetXml "S$row" $record.PsNewUpper
    }
    Replace-ZipXmlEntry $targetZip 'xl/worksheets/sheet1.xml' $targetWorksheetXml

    [xml]$targetWorkbookXml = Read-ZipEntryText $targetZip 'xl/workbook.xml'
    $calcPr = $targetWorkbookXml.workbook.calcPr
    if (-not $calcPr) {
        $calcPr = $targetWorkbookXml.CreateElement('calcPr', $SpreadsheetNamespace)
        [void]$targetWorkbookXml.workbook.AppendChild($calcPr)
    }
    $calcPr.SetAttribute('calcMode', 'auto')
    $calcPr.SetAttribute('fullCalcOnLoad', '1')
    $calcPr.SetAttribute('forceFullCalc', '1')
    $calcPr.SetAttribute('calcId', '0')
    Replace-ZipXmlEntry $targetZip 'xl/workbook.xml' $targetWorkbookXml
}
finally {
    $targetZip.Dispose()
    $targetStream.Dispose()
}

$n28 = $records | Where-Object { $_.Group -eq '构件总数对比' -and $_.TargetRow -eq 13 }
Write-Output "OUTPUT=$OutputWorkbook"
Write-Output "FILLED_RECORDS=$($records.Count)"
Write-Output "FILLED_CELLS=$($records.Count * 4)"
Write-Output "N28_PS=$($n28.PsLower.ToString('G17',[Globalization.CultureInfo]::InvariantCulture))/$($n28.PsUpper.ToString('G17',[Globalization.CultureInfo]::InvariantCulture))"
Write-Output "N28_PSNEW=$($n28.PsNewLower.ToString('G17',[Globalization.CultureInfo]::InvariantCulture))/$($n28.PsNewUpper.ToString('G17',[Globalization.CultureInfo]::InvariantCulture))"
