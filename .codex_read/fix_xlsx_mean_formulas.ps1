param(
  [Parameter(Mandatory=$true)][string]$Path,
  [Parameter(Mandatory=$true)][ValidateSet('single','multi')][string]$Mode
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$fullPath = [System.IO.Path]::GetFullPath($Path)
if (-not (Test-Path -LiteralPath $fullPath)) {
  throw "Workbook not found: $fullPath"
}

$workspace = [System.IO.Path]::GetFullPath((Get-Location).Path)
if (-not $fullPath.StartsWith($workspace, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw 'Workbook must stay inside workspace.'
}

$backupDir = Join-Path $workspace '.codex_read\formula_fix_backups'
if (-not (Test-Path -LiteralPath $backupDir)) {
  [void](New-Item -ItemType Directory -Path $backupDir -Force)
}
$backupPath = Join-Path $backupDir ([System.IO.Path]::GetFileName($fullPath))
if (-not (Test-Path -LiteralPath $backupPath)) {
  Copy-Item -LiteralPath $fullPath -Destination $backupPath
}

$zip = [System.IO.Compression.ZipFile]::Open($fullPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
  if ($null -eq $entry) { throw 'sheet1.xml is missing.' }

  $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
  try {
    [xml]$xml = $reader.ReadToEnd()
  } finally {
    $reader.Dispose()
  }

  $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
  $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
  # Keep formula status text encoding-safe under Windows PowerShell 5.1.
  $pendingText = -join @([char]0x5F85,[char]0x5BFC,[char]0x5E08,[char]0x586B,[char]0x5199)
  $passText = -join @([char]0x901A,[char]0x8FC7)
  $checkText = -join @([char]0x68C0,[char]0x67E5)

  function Set-Formula([string]$CellRef, [string]$Formula) {
    $cell = $xml.SelectSingleNode("//m:c[@r='$CellRef']", $ns)
    if ($null -eq $cell) { throw "Cell $CellRef is missing." }
    $formulaNode = $cell.SelectSingleNode('m:f', $ns)
    if ($null -eq $formulaNode) {
      $formulaNode = $xml.CreateElement('f', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
      [void]$cell.PrependChild($formulaNode)
    }
    foreach ($attributeName in @('t','ref','si')) {
      if ($formulaNode.HasAttribute($attributeName)) {
        $formulaNode.RemoveAttribute($attributeName)
      }
    }
    $formulaNode.InnerText = $Formula
    $cachedValue = $cell.SelectSingleNode('m:v', $ns)
    if ($null -ne $cachedValue) {
      [void]$cell.RemoveChild($cachedValue)
    }
  }

  if ($Mode -eq 'single') {
    foreach ($row in 3..28) {
      Set-Formula "Q$row" ('IF(COUNT(O{0}:P{0})=2,AVERAGE(O{0}:P{0}),"")' -f $row)
      Set-Formula "V$row" ('IF(COUNT(T{0}:U{0})=2,AVERAGE(T{0}:U{0}),"")' -f $row)
      Set-Formula "W$row" ('IF(COUNT(I{0}:L{0})+COUNT(R{0}:S{0})<6,"{1}",IF(AND(I{0}>=0,J{0}<=1,K{0}>=0,L{0}<=1,R{0}>=0,S{0}<=1,I{0}<=J{0},K{0}<=L{0},R{0}<=S{0},O{0}<=P{0},T{0}<=U{0},Q{0}>=1),"{2}","{3}"))' -f $row,$pendingText,$passText,$checkText)
    }
  } else {
    foreach ($row in 3..173) {
      Set-Formula "T$row" ('IF(COUNT(R{0}:S{0})=2,AVERAGE(R{0}:S{0}),"")' -f $row)
      Set-Formula "Y$row" ('IF(COUNT(W{0}:X{0})=2,AVERAGE(W{0}:X{0}),"")' -f $row)
      Set-Formula "Z$row" ('IF(COUNT(L{0}:O{0})+COUNT(U{0}:V{0})<6,"{1}",IF(AND(L{0}>=0,M{0}<=1,N{0}>=0,O{0}<=1,U{0}>=0,V{0}<=1,L{0}<=M{0},N{0}<=O{0},U{0}<=V{0},R{0}<=S{0},W{0}<=X{0},T{0}>=1),"{2}","{3}"))' -f $row,$pendingText,$passText,$checkText)
    }
  }

  $settings = [System.Xml.XmlWriterSettings]::new()
  $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
  $settings.Indent = $false
  $settings.OmitXmlDeclaration = $false

  $stream = $entry.Open()
  try {
    $stream.SetLength(0)
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    try {
      $xml.Save($writer)
    } finally {
      $writer.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
} finally {
  $zip.Dispose()
}

Write-Output $fullPath
