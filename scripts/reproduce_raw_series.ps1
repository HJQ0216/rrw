param(
    [string]$RawWorkbook = "D:\钢结构RRW方向\数据文件\β = 3.0串联系统 .xlsx",
    [string]$PaperWorkbook = "D:\钢结构RRW方向\吴仁彬 毕业论文\数据串联系统.xlsx",
    [string]$OutputWorkbook = "D:\钢结构RRW方向\数据文件\RAW论文单参数复现与对比.xlsx",
    [double]$N28AllPlasticLower = 0.0359,
    [double]$N28AllPlasticUpper = 0.0361
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

function Read-ZipXml {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { throw "ZIP entry not found: $EntryName" }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

function Open-XlsxReadOnly {
    param([string]$Path)
    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )
    $zip = [System.IO.Compression.ZipArchive]::new(
        $stream,
        [System.IO.Compression.ZipArchiveMode]::Read,
        $false
    )
    $workbookXml = Read-ZipXml $zip 'xl/workbook.xml'
    $relsXml = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'

    $sharedStrings = @()
    $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
    if ($sharedEntry) {
        $sharedXml = Read-ZipXml $zip 'xl/sharedStrings.xml'
        foreach ($si in $sharedXml.sst.si) {
            if ($si.t) {
                $sharedStrings += [string]$si.t
            } else {
                $sharedStrings += (($si.r | ForEach-Object { [string]$_.t }) -join '')
            }
        }
    }

    $sheetMap = @{}
    foreach ($sheet in $workbookXml.workbook.sheets.sheet) {
        $relId = $sheet.GetAttribute(
            'id',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        )
        $rel = $relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $relId }
        $target = [string]$rel.Target
        if ($target.StartsWith('/')) {
            $entryName = $target.TrimStart('/')
        } elseif ($target.StartsWith('xl/')) {
            $entryName = $target
        } else {
            $entryName = 'xl/' + $target
        }
        $sheetMap[([string]$sheet.name).Trim()] = $entryName
    }

    return [pscustomobject]@{
        Path = $Path
        Stream = $stream
        Zip = $zip
        SharedStrings = $sharedStrings
        SheetMap = $sheetMap
        SheetCache = @{}
    }
}

function Close-XlsxReadOnly {
    param($Book)
    if ($Book.Zip) { $Book.Zip.Dispose() }
    if ($Book.Stream) { $Book.Stream.Dispose() }
}

function Get-SheetXml {
    param($Book, [string]$SheetName)
    $key = $SheetName.Trim()
    if (-not $Book.SheetMap.ContainsKey($key)) {
        throw "Worksheet not found in $($Book.Path): [$SheetName]"
    }
    if (-not $Book.SheetCache.ContainsKey($key)) {
        $Book.SheetCache[$key] = Read-ZipXml $Book.Zip $Book.SheetMap[$key]
    }
    return $Book.SheetCache[$key]
}

function Get-CellObject {
    param($Book, [string]$SheetName, [string]$CellReference)
    $sheetXml = Get-SheetXml $Book $SheetName
    $cell = $sheetXml.worksheet.sheetData.row.c | Where-Object { $_.r -eq $CellReference }
    if (-not $cell) {
        throw "Cell not found: $SheetName!$CellReference"
    }
    return $cell
}

function Get-CellNumber {
    param($Book, [string]$SheetName, [string]$CellReference)
    $cell = Get-CellObject $Book $SheetName $CellReference
    $text = [string]$cell.v
    $value = 0.0
    $ok = [double]::TryParse(
        $text,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$value
    )
    if (-not $ok) { throw "Cell is not numeric: $SheetName!$CellReference = [$text]" }
    return $value
}

function Get-CellText {
    param($Book, [string]$SheetName, [string]$CellReference)
    $cell = Get-CellObject $Book $SheetName $CellReference
    $value = [string]$cell.v
    if ($cell.t -eq 's') { return $Book.SharedStrings[[int]$value] }
    if ($cell.t -eq 'inlineStr') {
        if ($cell.is.t) { return [string]$cell.is.t }
        return (($cell.is.r | ForEach-Object { [string]$_.t }) -join '')
    }
    return $value
}

function Get-Combination {
    param([int]$N, [int]$K)
    if ($K -lt 0 -or $K -gt $N) { return 0.0 }
    if ($K -eq 0 -or $K -eq $N) { return 1.0 }
    $k2 = [Math]::Min($K, $N - $K)
    $result = 1.0
    for ($i = 1; $i -le $k2; $i++) {
        $result = $result * ($N - $k2 + $i) / $i
    }
    return $result
}

function Get-ConditionalRefs {
    param([string]$Rho, [int]$BrittleCount)
    $row = switch ($Rho) {
        '0.1' { 13 }
        '0.3' { 13 }
        '0.5' { 34 }
        '0.7' { 34 }
        '0.9' { 54 }
        default { throw "Unsupported rho: $Rho" }
    }
    $pairs = if ($Rho -in @('0.1', '0.5', '0.9')) {
        @(@('C','D'), @('G','H'), @('K','L'), @('O','P'))
    } else {
        @(@('U','V'), @('Y','Z'), @('AC','AD'), @('AG','AH'))
    }
    if ($BrittleCount -lt 0 -or $BrittleCount -gt 3) {
        throw "Only 0-3 brittle-state groups exist in the source workbook."
    }
    $lowerRef = '{0}{1}' -f $pairs[$BrittleCount][0], $row
    $upperRef = '{0}{1}' -f $pairs[$BrittleCount][1], $row
    return @($lowerRef, $upperRef)
}

function Get-SystemBounds {
    param(
        $RawBook,
        [int]$ComponentCount,
        [int]$PotentialBrittleCount,
        [string]$Rho,
        [double]$BrittleProbability = 0.2
    )
    $sheet = "$ComponentCount`根构件"
    $lower = 0.0
    $upper = 0.0
    $sources = @()
    for ($b = 0; $b -le $PotentialBrittleCount; $b++) {
        $refs = Get-ConditionalRefs $Rho $b
        $conditionalLower = Get-CellNumber $RawBook $sheet $refs[0]
        $conditionalUpper = Get-CellNumber $RawBook $sheet $refs[1]
        $sourceLabel = "$($refs[0])/$($refs[1])"
        if ($ComponentCount -eq 28 -and $Rho -eq '0.1' -and $b -eq 0) {
            $conditionalLower = $N28AllPlasticLower
            $conditionalUpper = $N28AllPlasticUpper
            $sourceLabel = "$($refs[0])/$($refs[1])按确认值替换为$($conditionalLower.ToString('0.0000',[Globalization.CultureInfo]::InvariantCulture))/$($conditionalUpper.ToString('0.0000',[Globalization.CultureInfo]::InvariantCulture))"
        }
        $weight = (Get-Combination $PotentialBrittleCount $b) *
            [Math]::Pow($BrittleProbability, $b) *
            [Math]::Pow(1.0 - $BrittleProbability, $PotentialBrittleCount - $b)
        $lower += $weight * $conditionalLower
        $upper += $weight * $conditionalUpper
        $sources += "$sourceLabel×$($weight.ToString('0.###',[Globalization.CultureInfo]::InvariantCulture))"
    }
    return [pscustomobject]@{
        Lower = $lower
        Upper = $upper
        Source = "${sheet}:" + ($sources -join '; ')
    }
}

function New-ReproductionRecord {
    param(
        $RawBook,
        [string]$Group,
        [double]$Parameter,
        [int]$N,
        [int]$N1,
        [double]$Beta,
        [string]$Rho,
        [double]$Pbi,
        $PaperLower,
        $PaperUpper,
        $PaperAverage,
        [string]$PaperSource,
        [bool]$CanReproduce = $true,
        [string]$MissingReason = ''
    )
    $ps = $null
    $psNew = $null
    $rawLower = $null
    $rawUpper = $null
    $rawAverage = $null
    $diffLower = $null
    $diffUpper = $null
    $diffAverage = $null
    $source = ''
    $verdict = ''

    if ($CanReproduce) {
        $ps = Get-SystemBounds $RawBook $N $N1 $Rho $Pbi
        $psNew = Get-SystemBounds $RawBook ($N - 1) ($N1 - 1) $Rho $Pbi
        $rawLower = $psNew.Lower / $ps.Lower
        $rawUpper = $psNew.Upper / $ps.Upper
        $rawAverage = ($rawLower + $rawUpper) / 2.0
        $source = "Ps={$($ps.Source)}; Psnew={$($psNew.Source)}"
        if ($null -ne $PaperLower -and $null -ne $PaperUpper -and $null -ne $PaperAverage) {
            $diffLower = $rawLower - [double]$PaperLower
            $diffUpper = $rawUpper - [double]$PaperUpper
            $diffAverage = $rawAverage - [double]$PaperAverage
            if ([Math]::Abs($diffLower) -le 1e-10 -and
                [Math]::Abs($diffUpper) -le 1e-10 -and
                [Math]::Abs($diffAverage) -le 1e-10) {
                $verdict = '一致'
            } else {
                $verdict = '不一致'
            }
        } else {
            $verdict = '缺论文数值，无法比对'
        }
    } else {
        $verdict = $MissingReason
    }

    return [pscustomobject]@{
        Group = $Group
        Parameter = $Parameter
        N = $N
        N1 = $N1
        Beta = $Beta
        Rho = [double]$Rho
        Pbi = $Pbi
        PsLower = if ($ps) { $ps.Lower } else { $null }
        PsUpper = if ($ps) { $ps.Upper } else { $null }
        PsNewLower = if ($psNew) { $psNew.Lower } else { $null }
        PsNewUpper = if ($psNew) { $psNew.Upper } else { $null }
        RawLower = $rawLower
        RawUpper = $rawUpper
        RawAverage = $rawAverage
        PaperLower = $PaperLower
        PaperUpper = $PaperUpper
        PaperAverage = $PaperAverage
        DiffLower = $diffLower
        DiffUpper = $diffUpper
        DiffAverage = $diffAverage
        Verdict = $verdict
        Source = $source
        PaperSource = $PaperSource
    }
}

function Escape-XmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Get-ExcelColumnName {
    param([int]$Number)
    $name = ''
    while ($Number -gt 0) {
        $Number--
        $name = [char](65 + ($Number % 26)) + $name
        $Number = [Math]::Floor($Number / 26)
    }
    return $name
}

function TextCell {
    param($Value, [int]$Style = 3)
    return [pscustomobject]@{ Kind = 'text'; Value = [string]$Value; Style = $Style; Formula = $null }
}

function NumberCell {
    param($Value, [int]$Style = 4, [string]$Formula = $null)
    if ($null -eq $Value) { return $null }
    return [pscustomobject]@{ Kind = 'number'; Value = [double]$Value; Style = $Style; Formula = $Formula }
}

function Build-CellXml {
    param([string]$Reference, $Cell)
    if ($null -eq $Cell) { return '' }
    $style = $Cell.Style
    if ($Cell.Kind -eq 'text') {
        $value = Escape-XmlText $Cell.Value
        return "<c r=`"$Reference`" s=`"$style`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$value</t></is></c>"
    }
    $number = ([double]$Cell.Value).ToString('G17', [Globalization.CultureInfo]::InvariantCulture)
    if ($Cell.Formula) {
        $formula = Escape-XmlText $Cell.Formula
        return "<c r=`"$Reference`" s=`"$style`"><f>$formula</f><v>$number</v></c>"
    }
    return "<c r=`"$Reference`" s=`"$style`"><v>$number</v></c>"
}

function Build-SheetXml {
    param(
        [array]$Rows,
        [double[]]$Widths,
        [string[]]$Merges = @(),
        [int]$FreezeRows = 0,
        [string]$AutoFilter = ''
    )
    $maxCols = 1
    foreach ($row in $Rows) { if ($row.Count -gt $maxCols) { $maxCols = $row.Count } }
    $maxRows = $Rows.Count
    $dimension = "A1:$(Get-ExcelColumnName $maxCols)$maxRows"
    $sheetViews = '<sheetViews><sheetView workbookViewId="0">'
    if ($FreezeRows -gt 0) {
        $active = "A$($FreezeRows + 1)"
        $sheetViews += "<pane ySplit=`"$FreezeRows`" topLeftCell=`"$active`" activePane=`"bottomLeft`" state=`"frozen`"/><selection pane=`"bottomLeft`" activeCell=`"$active`" sqref=`"$active`"/>"
    }
    $sheetViews += '</sheetView></sheetViews>'

    $colsXml = '<cols>'
    for ($i = 0; $i -lt $Widths.Count; $i++) {
        $index = $i + 1
        $width = $Widths[$i].ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)
        $colsXml += "<col min=`"$index`" max=`"$index`" width=`"$width`" customWidth=`"1`"/>"
    }
    $colsXml += '</cols>'

    $rowsXml = '<sheetData>'
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $rowNumber = $r + 1
        $height = if ($rowNumber -eq 1) { ' ht="24" customHeight="1"' } elseif ($rowNumber -eq 2) { ' ht="30" customHeight="1"' } else { '' }
        $rowsXml += "<row r=`"$rowNumber`"$height>"
        $row = $Rows[$r]
        for ($c = 0; $c -lt $row.Count; $c++) {
            $reference = "$(Get-ExcelColumnName ($c + 1))$rowNumber"
            $rowsXml += Build-CellXml $reference $row[$c]
        }
        $rowsXml += '</row>'
    }
    $rowsXml += '</sheetData>'

    $mergeXml = ''
    if ($Merges.Count -gt 0) {
        $mergeXml = "<mergeCells count=`"$($Merges.Count)`">"
        foreach ($merge in $Merges) { $mergeXml += "<mergeCell ref=`"$merge`"/>" }
        $mergeXml += '</mergeCells>'
    }
    $filterXml = if ($AutoFilter) { "<autoFilter ref=`"$AutoFilter`"/>" } else { '' }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="$dimension"/>
  $sheetViews
  <sheetFormatPr defaultRowHeight="15"/>
  $colsXml
  $rowsXml
  $filterXml
  $mergeXml
  <pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
  <pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>
</worksheet>
"@
}

function Add-ZipTextEntry {
    param($Zip, [string]$EntryName, [string]$Content)
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
    try { $writer.Write($Content) }
    finally { $writer.Dispose() }
}

function Write-Xlsx {
    param([string]$Path, [array]$Sheets)
    if (Test-Path -LiteralPath $Path) {
        $directory = Split-Path -Parent $Path
        $base = [IO.Path]::GetFileNameWithoutExtension($Path)
        $extension = [IO.Path]::GetExtension($Path)
        $index = 2
        do {
            $candidate = Join-Path $directory "$base`_v$index$extension"
            $index++
        } while (Test-Path -LiteralPath $candidate)
        $Path = $candidate
    }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
"@
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            $contentTypes += "  <Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>`n"
        }
        $contentTypes += '</Types>'

        $rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

        $workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="120" yWindow="45" windowWidth="24000" windowHeight="12000"/></bookViews>
  <sheets>
"@
        $workbookRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"@
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            $sheetName = Escape-XmlText $Sheets[$i - 1].Name
            $workbook += "    <sheet name=`"$sheetName`" sheetId=`"$i`" r:id=`"rId$i`"/>`n"
            $workbookRels += "  <Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>`n"
        }
        $styleRel = $Sheets.Count + 1
        $workbook += "  </sheets><calcPr calcId=`"191029`" fullCalcOnLoad=`"1`" forceFullCalc=`"1`"/></workbook>"
        $workbookRels += "  <Relationship Id=`"rId$styleRel`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>`n</Relationships>"

        $styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2"><numFmt numFmtId="164" formatCode="0.000000000000"/><numFmt numFmtId="165" formatCode="0.0"/></numFmts>
  <fonts count="4">
    <font><sz val="10.5"/><name val="等线"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="10.5"/><name val="等线"/></font>
    <font><b/><sz val="15"/><color rgb="FF1F4E78"/><name val="等线"/></font>
    <font><b/><sz val="10.5"/><name val="等线"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E1F2"/></left><right style="thin"><color rgb="FFD9E1F2"/></right><top style="thin"><color rgb="FFD9E1F2"/></top><bottom style="thin"><color rgb="FFD9E1F2"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@

        $now = [DateTime]::UtcNow.ToString('s') + 'Z'
        $core = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>RAW论文单参数复现与对比</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@
        $app = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application></Properties>
"@

        Add-ZipTextEntry $zip '[Content_Types].xml' $contentTypes
        Add-ZipTextEntry $zip '_rels/.rels' $rootRels
        Add-ZipTextEntry $zip 'xl/workbook.xml' $workbook
        Add-ZipTextEntry $zip 'xl/_rels/workbook.xml.rels' $workbookRels
        Add-ZipTextEntry $zip 'xl/styles.xml' $styles
        Add-ZipTextEntry $zip 'docProps/core.xml' $core
        Add-ZipTextEntry $zip 'docProps/app.xml' $app
        for ($i = 1; $i -le $Sheets.Count; $i++) {
            Add-ZipTextEntry $zip "xl/worksheets/sheet$i.xml" $Sheets[$i - 1].Xml
        }
    }
    finally {
        $zip.Dispose()
        $stream.Dispose()
    }
    return $Path
}

function Build-ComparisonSheet {
    param([string]$Title, [string]$FixedParameters, [array]$Records, [string]$ParameterHeader)
    $headers = @(
        $ParameterHeader, '构件总数 n', '潜在脆性构件数 n₁', '可靠度指标 β', '相关系数 ρ', '脆性发生概率 Pbi',
        '复现 Ps 下界', '复现 Ps 上界', '复现 Psnew 下界', '复现 Psnew 上界',
        '复现 RAW 下界', '复现 RAW 上界', '复现 RAW 平均值',
        '论文 RAW 下界', '论文 RAW 上界', '论文 RAW 平均值',
        '下界差值', '上界差值', '平均值差值', '比对结论', '原始数据来源', '论文数据来源'
    )
    $rows = @()
    $rows += ,@((TextCell $Title 1))
    $rows += ,@((TextCell $FixedParameters 9))
    $rows += ,@($headers | ForEach-Object { TextCell $_ 2 })
    $excelRow = 4
    foreach ($record in $Records) {
        $verdictStyle = if ($record.Verdict -eq '一致') { 6 } elseif ($record.Verdict -eq '不一致') { 7 } else { 8 }
        $rawLowerFormula = if ($null -ne $record.RawLower) { "I$excelRow/G$excelRow" } else { $null }
        $rawUpperFormula = if ($null -ne $record.RawUpper) { "J$excelRow/H$excelRow" } else { $null }
        $rawAverageFormula = if ($null -ne $record.RawAverage) { "AVERAGE(K${excelRow}:L${excelRow})" } else { $null }
        $diffLowerFormula = if ($null -ne $record.DiffLower) { "K$excelRow-N$excelRow" } else { $null }
        $diffUpperFormula = if ($null -ne $record.DiffUpper) { "L$excelRow-O$excelRow" } else { $null }
        $diffAverageFormula = if ($null -ne $record.DiffAverage) { "M$excelRow-P$excelRow" } else { $null }
        $rows += ,@(
            (NumberCell $record.Parameter 5), (NumberCell $record.N 5), (NumberCell $record.N1 5),
            (NumberCell $record.Beta 5), (NumberCell $record.Rho 5), (NumberCell $record.Pbi 5),
            (NumberCell $record.PsLower 4), (NumberCell $record.PsUpper 4),
            (NumberCell $record.PsNewLower 4), (NumberCell $record.PsNewUpper 4),
            (NumberCell $record.RawLower 4 $rawLowerFormula), (NumberCell $record.RawUpper 4 $rawUpperFormula),
            (NumberCell $record.RawAverage 4 $rawAverageFormula),
            (NumberCell $record.PaperLower 4), (NumberCell $record.PaperUpper 4), (NumberCell $record.PaperAverage 4),
            (NumberCell $record.DiffLower 4 $diffLowerFormula), (NumberCell $record.DiffUpper 4 $diffUpperFormula),
            (NumberCell $record.DiffAverage 4 $diffAverageFormula),
            (TextCell $record.Verdict $verdictStyle), (TextCell $record.Source 3), (TextCell $record.PaperSource 3)
        )
        $excelRow++
    }
    $widths = @(12,12,16,12,12,16,15,15,17,17,17,17,17,17,17,17,15,15,15,20,52,30)
    $lastRow = $Rows.Count
    return [pscustomobject]@{
        Xml = Build-SheetXml $rows $widths @('A1:V1','A2:V2') 3 "A3:V$lastRow"
        RowCount = $lastRow
    }
}

$rawBook = $null
$paperBook = $null
try {
    $rawBook = Open-XlsxReadOnly $RawWorkbook
    $paperBook = Open-XlsxReadOnly $PaperWorkbook

    $rhoRecords = @()
    $rhoRows = @{'0.1'=3; '0.3'=6; '0.5'=9; '0.7'=12; '0.9'=15}
    foreach ($rho in @('0.1','0.3','0.5','0.7','0.9')) {
        $row = $rhoRows[$rho]
        $rhoRecords += New-ReproductionRecord $rawBook '相关系数' ([double]$rho) 8 1 3.0 $rho 0.2 `
            (Get-CellNumber $paperBook 'Sheet1' "C$row") `
            (Get-CellNumber $paperBook 'Sheet1' "D$row") `
            (Get-CellNumber $paperBook 'Sheet1' "E$row") `
            "数据串联系统.xlsx!A$row:E$row"
    }

    $nRecords = @()
    $nValues = @(8,12,16,20,24,28,32,36,40,44,48)
    for ($i = 0; $i -lt $nValues.Count; $i++) {
        $n = $nValues[$i]
        $row = 59 + $i
        $nRecords += New-ReproductionRecord $rawBook '构件总数' $n $n 1 3.0 '0.1' 0.2 `
            (Get-CellNumber $paperBook 'Sheet1' "C$row") `
            (Get-CellNumber $paperBook 'Sheet1' "D$row") `
            (Get-CellNumber $paperBook 'Sheet1' "E$row") `
            "数据串联系统.xlsx!B$row:E$row"
    }

    $n1Records = @()
    foreach ($n1 in 1,2,3) {
        $row = 146 + $n1
        $n1Records += New-ReproductionRecord $rawBook '潜在脆性构件数量' $n1 8 $n1 3.0 '0.1' 0.2 `
            (Get-CellNumber $paperBook 'Sheet1' "C$row") `
            (Get-CellNumber $paperBook 'Sheet1' "D$row") `
            (Get-CellNumber $paperBook 'Sheet1' "E$row") `
            "数据串联系统.xlsx!B$row:E$row"
    }

    $betaRecords = @()
    foreach ($beta in @(3.0,3.1,3.2,3.3,3.4,3.5,3.6)) {
        $canReproduce = ($beta -eq 3.0)
        $paperLower = $null
        $paperUpper = $null
        $paperAverage = $null
        $paperSource = '数据串联系统.xlsx中未找到该β对应的最终精确RAW数据'
        if ($beta -eq 3.0) {
            $paperLower = Get-CellNumber $paperBook 'Sheet1' 'AA130'
            $paperUpper = Get-CellNumber $paperBook 'Sheet1' 'AB130'
            $paperAverage = ($paperLower + $paperUpper) / 2.0
            $paperSource = '数据串联系统.xlsx!Z130:AB130'
        } elseif ($beta -eq 3.2) {
            $paperLower = Get-CellNumber $paperBook 'Sheet1' 'AA131'
            $paperUpper = Get-CellNumber $paperBook 'Sheet1' 'AB131'
            $paperAverage = ($paperLower + $paperUpper) / 2.0
            $paperSource = '数据串联系统.xlsx!Z131:AB131（该值与最终论文图2.8曲线不一致）'
        }
        $reason = if ($canReproduce) { '' } else { '缺少该β的原始RLP Ps/Psnew，无法复现' }
        $betaRecords += New-ReproductionRecord $rawBook '可靠度指标' $beta 10 1 $beta '0.1' 0.2 `
            $paperLower $paperUpper $paperAverage $paperSource $canReproduce $reason
    }

    $paper37Lower = Get-CellNumber $paperBook 'Sheet1' 'AA132'
    $paper37Upper = Get-CellNumber $paperBook 'Sheet1' 'AB132'
    $betaRecords += New-ReproductionRecord $rawBook '可靠度指标（数据文件额外点）' 3.7 10 1 3.7 '0.1' 0.2 `
        $paper37Lower $paper37Upper (($paper37Lower + $paper37Upper) / 2.0) `
        '数据串联系统.xlsx!Z132:AB132（β=3.7不属于论文图2.8的3.0~3.6范围）' $false `
        '数据文件额外点；缺少β=3.7原始RLP数据'

    $allComparable = @($rhoRecords + $nRecords + $n1Records + ($betaRecords | Where-Object { $_.Parameter -eq 3.0 }))
    $matchCount = @($allComparable | Where-Object { $_.Verdict -eq '一致' }).Count
    $mismatchCount = @($allComparable | Where-Object { $_.Verdict -eq '不一致' }).Count

    $explanationRows = @(
        ,@((TextCell 'RAW论文单参数复现与对比' 1)),
        ,@((TextCell '复现范围、计算定义、数据来源与结论' 9)),
        ,@((TextCell '项目' 2), (TextCell '内容' 2)),
        ,@((TextCell 'RAW定义' 3), (TextCell '按论文：从原串联系统S中去掉目标构件i，计算S_-i与S的失效概率之比。' 3)),
        ,@((TextCell 'Ps' 3), (TextCell '原系统S的失效概率边界；由k=3条件RLP数据按Pbi=0.2进行全概率加权。' 3)),
        ,@((TextCell 'Psnew' 3), (TextCell '去掉构件i后的系统S_-i失效概率边界；构件总数n→n-1，潜在脆性构件数n₁→n₁-1。' 3)),
        ,@((TextCell 'RAW计算' 3), (TextCell '论文实际数据采用RAW下界=Psnew下界/Ps下界，RAW上界=Psnew上界/Ps上界，平均值=(上下界)/2。' 3)),
        ,@((TextCell 'RLP阶次' 3), (TextCell '统一采用三阶RLP（k=3）。' 3)),
        ,@((TextCell '原始RLP文件' 3), (TextCell $RawWorkbook 3)),
        ,@((TextCell '论文数据文件' 3), (TextCell $PaperWorkbook 3)),
        ,@((TextCell '可直接比对结果' 3), (TextCell "共$($allComparable.Count)个有完整原始输入且有论文值的点：一致$matchCount个，不一致$mismatchCount个。" 3)),
        ,@((TextCell '重要原则' 3), (TextCell '未修改两个原始文件；按已确认要求，仅在本复现文件中将28根构件全塑性、k=3的下/上界替换为0.0359/0.0361；缺少β数据的点保持空白并注明原因。' 8))
    )
    $explanationXml = Build-SheetXml $explanationRows @(22,115) @('A1:B1','A2:B2') 3 ''

    $rhoSheet = Build-ComparisonSheet '相关系数单参数：RAW复现与论文数据对比' '固定参数：n=8，n₁=1，β=3.0，Pbi=0.2；ρ=0.1、0.3、0.5、0.7、0.9' $rhoRecords '相关系数 ρ'
    $nSheet = Build-ComparisonSheet '构件总数单参数：RAW复现与论文数据对比' '固定参数：n₁=1，β=3.0，ρ=0.1，Pbi=0.2；n=8、12、16、20、24、28、32、36、40、44、48' $nRecords '构件总数 n'
    $betaSheet = Build-ComparisonSheet '可靠度指标单参数：RAW复现与论文数据对比' '固定参数：n=10，n₁=1，ρ=0.1，Pbi=0.2；论文图2.8参数范围β=3.0~3.6。当前原始RLP文件仅有β=3.0。' $betaRecords '可靠度指标 β'
    $n1Sheet = Build-ComparisonSheet '潜在脆性构件数量单参数：RAW复现与论文数据对比' '固定参数：n=8，β=3.0，ρ=0.1，Pbi=0.2；n₁=1、2、3' $n1Records '潜在脆性构件数 n₁'

    $n28 = $nRecords | Where-Object { $_.N -eq 28 }
    $anomalyRows = @(
        ,@((TextCell '复现过程中发现的问题' 1)),
        ,@((TextCell '以下记录同时包含已确认替换和仍待处理的数据问题' 9)),
        ,@((TextCell '序号' 2), (TextCell '级别' 2), (TextCell '位置' 2), (TextCell '发现' 2), (TextCell '影响/判断' 2)),
        ,@((NumberCell 1 5), (TextCell '已确认替换' 6), (TextCell '28根构件：全塑性、k=3，C13/D13' 3),
            (TextCell "本复现文件采用下界0.0359、上界0.0361；加权后Ps下界=$($n28.PsLower.ToString('0.000000000000'))、Ps上界=$($n28.PsUpper.ToString('0.000000000000'))，复现RAW下界=$($n28.RawLower.ToString('0.000000000000'))、RAW上界=$($n28.RawUpper.ToString('0.000000000000'))、平均值=$($n28.RawAverage.ToString('0.000000000000'))。" 3),
            (TextCell "原始RLP文件保持不变；复现RAW下界与论文值$($n28.PaperLower.ToString('0.000000000000'))仍有差异，未再作其他修正。" 6)),
        ,@((NumberCell 2 5), (TextCell '数据缺口' 8), (TextCell '可靠度指标β=3.1~3.6' 3),
            (TextCell 'β=3.0原始表只包含β=3.0，缺少β=3.1~3.6的条件RLP数据及Ps/Psnew。' 3),
            (TextCell '除β=3.0外无法重新计算RAW；需要补充n=10、n₁=1、ρ=0.1、Pbi=0.2对应的原始RLP数据。' 8)),
        ,@((NumberCell 3 5), (TextCell '论文数据文件内部不一致' 8), (TextCell '数据串联系统.xlsx：D4 与 D148' 3),
            (TextCell "同一参数n=8、n₁=2、β=3.0、ρ=0.1、Pbi=0.2，D4=$((Get-CellNumber $paperBook 'Sheet1' 'D4').ToString('0.000000000000'))，D148=$((Get-CellNumber $paperBook 'Sheet1' 'D148').ToString('0.000000000000'))。" 3),
            (TextCell '三阶RLP复现结果与D148一致，D4疑似复制错误。' 8)),
        ,@((NumberCell 4 5), (TextCell '论文数据文件内部不一致' 8), (TextCell '数据串联系统.xlsx：可靠度指标块Z128:AB132' 3),
            (TextCell '该块只有β=3.0、3.2、3.7，而最终论文图2.8采用β=3.0、3.1、3.2、3.3、3.4、3.5、3.6；其中β=3.7不在最终范围内。' 3),
            (TextCell '该块应属于早期试算或不完整数据，不能作为最终图2.8的完整原始数据。' 8)),
        ,@((NumberCell 5 5), (TextCell '论文数据文件重复值差异' 8), (TextCell '数据串联系统.xlsx：C67 与 P185' 3),
            (TextCell "相同参数n=40、n₁=1、β=3.0、ρ=0.1、Pbi=0.2，C67=$((Get-CellNumber $paperBook 'Sheet1' 'C67').ToString('0.000000000000'))，P185=$((Get-CellNumber $paperBook 'Sheet1' 'P185').ToString('0.000000000000'))。" 3),
            (TextCell '本次采用专门用于图2.7的B57:E69数据块；原始RLP复现结果与C67一致。' 8))
    )
    $anomalyXml = Build-SheetXml $anomalyRows @(8,22,42,78,78) @('A1:E1','A2:E2') 3 'A3:E8'

    $sheets = @(
        [pscustomobject]@{ Name='复现说明'; Xml=$explanationXml },
        [pscustomobject]@{ Name='相关系数对比'; Xml=$rhoSheet.Xml },
        [pscustomobject]@{ Name='构件总数对比'; Xml=$nSheet.Xml },
        [pscustomobject]@{ Name='可靠度指标对比'; Xml=$betaSheet.Xml },
        [pscustomobject]@{ Name='潜在脆性数量对比'; Xml=$n1Sheet.Xml },
        [pscustomobject]@{ Name='异常清单'; Xml=$anomalyXml }
    )

    $actualOutput = Write-Xlsx $OutputWorkbook $sheets
    Write-Output "OUTPUT=$actualOutput"
    Write-Output "COMPARABLE=$($allComparable.Count)"
    Write-Output "MATCH=$matchCount"
    Write-Output "MISMATCH=$mismatchCount"
    Write-Output "N28_REPRO_RAW_L=$($n28.RawLower.ToString('G17',[Globalization.CultureInfo]::InvariantCulture))"
    Write-Output "N28_PAPER_RAW_L=$($n28.PaperLower.ToString('G17',[Globalization.CultureInfo]::InvariantCulture))"
}
finally {
    if ($paperBook) { Close-XlsxReadOnly $paperBook }
    if ($rawBook) { Close-XlsxReadOnly $rawBook }
}
