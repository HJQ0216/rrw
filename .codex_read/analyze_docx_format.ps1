param([string]$DocxPath = '.codex_read\source_readcopy.docx')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntry([IO.Compression.ZipArchive]$Zip, [string]$Name) {
  $entry = $Zip.GetEntry($Name)
  if ($null -eq $entry) { return $null }
  $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
  try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Attr($Node, [string]$LocalName) {
  if ($null -eq $Node) { return '' }
  foreach ($a in $Node.Attributes) { if ($a.LocalName -eq $LocalName) { return $a.Value } }
  return ''
}

function ChildAttr($Node, [string]$XPath, $Ns, [string]$AttrName='val') {
  if ($null -eq $Node) { return '' }
  return Attr ($Node.SelectSingleNode($XPath, $Ns)) $AttrName
}

function Select-One($Node, [string]$XPath, $Ns) {
  if ($null -eq $Node) { return $null }
  return $Node.SelectSingleNode($XPath, $Ns)
}

$path = (Resolve-Path -LiteralPath $DocxPath).Path
$zip = [IO.Compression.ZipFile]::OpenRead($path)
try {
  [xml]$doc = Read-ZipEntry $zip 'word/document.xml'
  [xml]$styles = Read-ZipEntry $zip 'word/styles.xml'
  $ns = [Xml.XmlNamespaceManager]::new($doc.NameTable)
  $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $sns = [Xml.XmlNamespaceManager]::new($styles.NameTable)
  $sns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')

  '=== SECTION PROPERTIES ==='
  foreach ($sect in $doc.SelectNodes('//w:sectPr',$ns)) {
    $size = $sect.SelectSingleNode('./w:pgSz',$ns)
    $mar = $sect.SelectSingleNode('./w:pgMar',$ns)
    "page w=$(Attr $size 'w') h=$(Attr $size 'h') orient=$(Attr $size 'orient'); margin top=$(Attr $mar 'top') right=$(Attr $mar 'right') bottom=$(Attr $mar 'bottom') left=$(Attr $mar 'left') header=$(Attr $mar 'header') footer=$(Attr $mar 'footer') gutter=$(Attr $mar 'gutter')"
  }

  '=== PARAGRAPH STYLE USAGE ==='
  $paras = $doc.SelectNodes('//w:body//w:p',$ns)
  $items = foreach ($p in $paras) {
    $style = ChildAttr $p './w:pPr/w:pStyle' $ns
    if (-not $style) { $style='(none)' }
    $text = (($p.SelectNodes('.//w:t',$ns) | ForEach-Object {$_.InnerText}) -join '')
    [pscustomobject]@{Style=$style;Text=$text}
  }
  $items | Group-Object Style | Sort-Object Count -Descending | ForEach-Object {
    $samples = ($_.Group | Where-Object {$_.Text.Trim()} | Select-Object -First 3 -ExpandProperty Text) -join ' || '
    "style=$($_.Name) count=$($_.Count) samples=$samples"
  }

  '=== STYLE DEFINITIONS ==='
  foreach ($s in $styles.SelectNodes('//w:style',$sns)) {
    $type=Attr $s 'type'; if($type -notin @('paragraph','character')){continue}
    $id=Attr $s 'styleId'; $name=ChildAttr $s './w:name' $sns
    $based=ChildAttr $s './w:basedOn' $sns; $next=ChildAttr $s './w:next' $sns
    $ppr=$s.SelectSingleNode('./w:pPr',$sns);$rpr=$s.SelectSingleNode('./w:rPr',$sns)
    $fonts=Select-One $rpr './w:rFonts' $sns;$spacing=Select-One $ppr './w:spacing' $sns;$ind=Select-One $ppr './w:ind' $sns
    $sz=ChildAttr $rpr './w:sz' $sns;$szCs=ChildAttr $rpr './w:szCs' $sns
    $bold=($null-ne(Select-One $rpr './w:b' $sns));$italic=($null-ne(Select-One $rpr './w:i' $sns))
    $jc=ChildAttr $ppr './w:jc' $sns;$outline=ChildAttr $ppr './w:outlineLvl' $sns
    "id=$id type=$type name=$name basedOn=$based next=$next fonts ascii=$(Attr $fonts 'ascii') eastAsia=$(Attr $fonts 'eastAsia') hAnsi=$(Attr $fonts 'hAnsi') size=$sz/$szCs bold=$bold italic=$italic jc=$jc outline=$outline spacing before=$(Attr $spacing 'before') after=$(Attr $spacing 'after') line=$(Attr $spacing 'line') rule=$(Attr $spacing 'lineRule') ind firstLine=$(Attr $ind 'firstLine') firstLineChars=$(Attr $ind 'firstLineChars') left=$(Attr $ind 'left')"
  }

  '=== KEY PARAGRAPH DIRECT FORMATTING ==='
  $patterns = @('摘  要','摘要','ABSTRACT','目  录','第1章','第2章','1.1 ','1.1.1','图 2.6','Figure 2.6','表2.1','Table 2.1','参考文献')
  foreach ($p in $paras) {
    $text=(($p.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText})-join '')
    if(-not ($patterns|Where-Object{$text -like "*$_*"})){continue}
    $ppr=$p.SelectSingleNode('./w:pPr',$ns);$style=ChildAttr $ppr './w:pStyle' $ns
    $spacing=$ppr.SelectSingleNode('./w:spacing',$ns);$ind=$ppr.SelectSingleNode('./w:ind',$ns);$jc=ChildAttr $ppr './w:jc' $ns
    $runs=@()
    foreach($r in $p.SelectNodes('./w:r',$ns)){
      $t=(($r.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText})-join '');if(-not$t){continue}
      $rp=$r.SelectSingleNode('./w:rPr',$ns);$f=Select-One $rp './w:rFonts' $ns
      $runs += "[$t|ascii=$(Attr $f 'ascii');ea=$(Attr $f 'eastAsia');size=$(ChildAttr $rp './w:sz' $ns);b=$($null-ne(Select-One $rp './w:b' $ns));i=$($null-ne(Select-One $rp './w:i' $ns))]"
    }
    "text=$text || style=$style jc=$jc line=$(Attr $spacing 'line') rule=$(Attr $spacing 'lineRule') before=$(Attr $spacing 'before') after=$(Attr $spacing 'after') firstLine=$(Attr $ind 'firstLine') chars=$(Attr $ind 'firstLineChars') || runs=$($runs -join '')"
  }

  '=== TABLE SUMMARY ==='
  $ti=0
  foreach($tbl in $doc.SelectNodes('//w:body//w:tbl',$ns)){
    $ti++;$rows=$tbl.SelectNodes('./w:tr',$ns);$grid=$tbl.SelectNodes('./w:tblGrid/w:gridCol',$ns)|ForEach-Object{Attr $_ 'w'}
    $text=(($tbl.SelectNodes('.//w:t',$ns)|Select-Object -First 12|ForEach-Object{$_.InnerText})-join '|')
    "table=$ti rows=$($rows.Count) widths=$($grid -join ',') sample=$text"
  }
} finally { $zip.Dispose() }
