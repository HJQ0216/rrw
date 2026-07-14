param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$out = [System.IO.Path]::GetFullPath($OutputDir)
$workspace = [System.IO.Path]::GetFullPath((Get-Location).Path)
if (-not $out.StartsWith($workspace, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "OutputDir must stay inside the current workspace."
}
if (-not (Test-Path -LiteralPath $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
$zip = Join-Path $out 'source.zip'
$unpacked = Join-Path $out 'unpacked'
Copy-Item -LiteralPath $InputDocx -Destination $zip -Force
if (Test-Path -LiteralPath $unpacked) { Remove-Item -LiteralPath $unpacked -Recurse -Force }
[System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $unpacked)

$docPath = Join-Path $unpacked 'word\document.xml'
$relsPath = Join-Path $unpacked 'word\_rels\document.xml.rels'
[xml]$doc = Get-Content -LiteralPath $docPath -Raw -Encoding UTF8
[xml]$rels = Get-Content -LiteralPath $relsPath -Raw -Encoding UTF8

$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
$ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
$ns.AddNamespace('v','urn:schemas-microsoft-com:vml')

$relMap = @{}
foreach ($rel in $rels.Relationships.Relationship) { $relMap[$rel.Id] = $rel.Target }

$paragraphs = New-Object System.Collections.Generic.List[object]
$imageRefs = New-Object System.Collections.Generic.List[object]
$i = 0
foreach ($p in $doc.SelectNodes('//w:body//w:p', $ns)) {
  $i++
  $texts = @($p.SelectNodes('.//w:t|.//w:tab|.//w:br', $ns) | ForEach-Object {
    if ($_.LocalName -eq 'tab') { "`t" } elseif ($_.LocalName -eq 'br') { "`n" } else { $_.'#text' }
  })
  $text = ($texts -join '').Trim()
  $styleNode = $p.SelectSingleNode('./w:pPr/w:pStyle', $ns)
  $style = if ($styleNode) { $styleNode.GetAttribute('val','http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { '' }
  $paragraphs.Add([pscustomobject]@{ Index=$i; Style=$style; Text=$text })
  foreach ($blip in $p.SelectNodes('.//a:blip', $ns)) {
    $rid = $blip.GetAttribute('embed','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $imageRefs.Add([pscustomobject]@{ Paragraph=$i; CaptionContext=$text; RelationshipId=$rid; Target=$relMap[$rid] })
  }
  foreach ($img in $p.SelectNodes('.//v:imagedata', $ns)) {
    $rid = $img.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $imageRefs.Add([pscustomobject]@{ Paragraph=$i; CaptionContext=$text; RelationshipId=$rid; Target=$relMap[$rid] })
  }
}

$paragraphs | Export-Csv -LiteralPath (Join-Path $out 'paragraphs.csv') -NoTypeInformation -Encoding UTF8
$paragraphs | ForEach-Object { if ($_.Text) { $_.Text } } | Set-Content -LiteralPath (Join-Path $out 'fulltext.txt') -Encoding UTF8
$imageRefs | Export-Csv -LiteralPath (Join-Path $out 'image_refs.csv') -NoTypeInformation -Encoding UTF8

$mediaDir = Join-Path $unpacked 'word\media'
$media = if (Test-Path -LiteralPath $mediaDir) { Get-ChildItem -LiteralPath $mediaDir -File } else { @() }
$media | Select-Object Name,Length,Extension,FullName | Export-Csv -LiteralPath (Join-Path $out 'media.csv') -NoTypeInformation -Encoding UTF8

Write-Output ("paragraphs={0}; imageRefs={1}; media={2}" -f $paragraphs.Count,$imageRefs.Count,$media.Count)
