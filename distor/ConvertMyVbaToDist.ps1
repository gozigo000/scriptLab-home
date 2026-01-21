param(
  [Parameter(Mandatory = $false)]
  [string]$SourceDir = "",

  [Parameter(Mandatory = $false)]
  [string]$OutRoot = "",

  [Parameter(Mandatory = $false)]
  [switch]$InPlace
)

$ErrorActionPreference = "Stop"

function Get-ScriptDir() {
  $p = $MyInvocation.MyCommand.Path
  if ([string]::IsNullOrEmpty($p)) { return (Get-Location).Path }
  return (Split-Path -Parent $p)
}

if ([string]::IsNullOrWhiteSpace($SourceDir)) {
  $candidates = @(
    (Join-Path (Get-Location).Path "myVba"),
    (Join-Path (Get-ScriptDir) "..\myVba"),
    (Join-Path (Get-ScriptDir) "myVba")
  )
  foreach ($c in $candidates) {
    if (Test-Path -LiteralPath $c) { $SourceDir = $c; break }
  }
  if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    $SourceDir = (Join-Path (Get-Location).Path "myVba")
  }
}

if ([string]::IsNullOrWhiteSpace($OutRoot)) {
  $OutRoot = (Join-Path (Get-Location).Path "dist")
}

function Ensure-Dir([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function New-StringFromCodePoints([int[]]$CodePoints) {
  return -join ($CodePoints | ForEach-Object { [char]$_ })
}

function Get-ModuleNameFromFileName([string]$FilePath) {
  return [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
}

function Has-VBNameHeader([string]$Content) {
  return ($Content -match '(?m)^\s*Attribute\s+VB_Name\s*=')
}

function Add-VBNameHeaderIfMissing([string]$Content, [string]$ModuleName) {
  if (Has-VBNameHeader $Content) { return $Content }
  return "Attribute VB_Name = `"$ModuleName`"`r`n" + $Content
}

function Strip-ExportAttributeLines([string]$Content) {
  # Remove export-only Attribute lines (safe for document-module paste/replacement).
  return ($Content -split "`r?`n" | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' }) -join "`r`n"
}

function Has-ClassExportHeader([string]$Content) {
  return ($Content -match '(?m)^\s*VERSION\s+\d+(\.\d+)?\s+CLASS\s*$')
}

function Wrap-ClassModuleIfNeeded([string]$Content, [string]$ClassName) {
  if (Has-ClassExportHeader $Content -and Has-VBNameHeader $Content) { return $Content }

  $header =
    "VERSION 1.0 CLASS`r`n" +
    "BEGIN`r`n" +
    "  MultiUse = -1  'True`r`n" +
    "END`r`n" +
    "Attribute VB_Name = `"$ClassName`"`r`n" +
    "Attribute VB_GlobalNameSpace = False`r`n" +
    "Attribute VB_Creatable = False`r`n" +
    "Attribute VB_PredeclaredId = False`r`n" +
    "Attribute VB_Exposed = False`r`n"

  return $header + $Content
}

if (-not (Test-Path -LiteralPath $SourceDir)) {
  throw "SourceDir not found: $SourceDir"
}

$sourceFull = (Resolve-Path -LiteralPath $SourceDir).Path.TrimEnd("\")

$vbaFiles = Get-ChildItem -LiteralPath $SourceDir -Recurse -File -Filter "*.vba"
$clsFiles = Get-ChildItem -LiteralPath $SourceDir -Recurse -File -Filter "*.cls"
$frmFiles = Get-ChildItem -LiteralPath $SourceDir -Recurse -File -Filter "*.frm"

if (($vbaFiles.Count + $clsFiles.Count + $frmFiles.Count) -eq 0) {
  Write-Host "No .vba/.cls files found under $SourceDir"
  exit 0
}

if ($InPlace) {
  foreach ($f in $vbaFiles) {
    $newPath = [System.IO.Path]::ChangeExtension($f.FullName, ".bas")
    if (Test-Path -LiteralPath $newPath) {
      Remove-Item -LiteralPath $newPath -Force
    }
    Rename-Item -LiteralPath $f.FullName -NewName ([System.IO.Path]::GetFileName($newPath))
  }
  Write-Host "Renamed $($vbaFiles.Count) files: .vba -> .bas (in place)."
  exit 0
}

# Windows PowerShell 5 can misread non-ASCII literals depending on script encoding.
# Build the Korean folder names from Unicode code points.
$nameMsWordObjects = New-StringFromCodePoints @(0x004D,0x0073,0x0057,0x006F,0x0072,0x0064,0xAC1D,0xCCB4)
$nameModules = New-StringFromCodePoints @(0xBAA8,0xB4C8)
$nameClassModules = New-StringFromCodePoints @(0xD074,0xB798,0xC2A4,0xBAA8,0xB4C8)
$nameForms = New-StringFromCodePoints @(0xD3FC)

$outMsWordObjects = (Join-Path $OutRoot $nameMsWordObjects)
$outModules = (Join-Path $OutRoot $nameModules)
$outClassModules = (Join-Path $OutRoot $nameClassModules)
$outForms = (Join-Path $OutRoot $nameForms)

Ensure-Dir $OutRoot
Ensure-Dir $outMsWordObjects
Ensure-Dir $outModules
Ensure-Dir $outClassModules
Ensure-Dir $outForms

foreach ($f in $vbaFiles) {
  $moduleName = Get-ModuleNameFromFileName $f.FullName

  $targetDir = if ($moduleName -ieq "ThisDocument") { $outMsWordObjects } else { $outModules }
  $targetExt = if ($moduleName -ieq "ThisDocument") { ".cls" } else { ".bas" }
  $targetPath = Join-Path $targetDir ([System.IO.Path]::ChangeExtension([System.IO.Path]::GetFileName($f.Name), $targetExt))

  # Read source as UTF-8 (common for repo-managed files).
  $content = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8

  if ($moduleName -ieq "ThisDocument") {
    # Document module code must NOT include Attribute headers (causes syntax error if pasted/replaced).
    $content = Strip-ExportAttributeLines $content
  } else {
    $content = Add-VBNameHeaderIfMissing $content $moduleName
  }

  # Normalize line endings to CRLF (VBA export convention).
  $content = $content -replace "`r?`n", "`r`n"
  # Save using system ANSI for VBA IDE import compatibility.
  [System.IO.File]::WriteAllText($targetPath, $content, [System.Text.Encoding]::Default)
}

foreach ($f in $clsFiles) {
  $className = Get-ModuleNameFromFileName $f.FullName
  $targetDir = if ($className -ieq "ThisDocument") { $outMsWordObjects } else { $outClassModules }
  $targetPath = Join-Path $targetDir ([System.IO.Path]::GetFileName($f.Name))

  $content = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
  if ($className -ieq "ThisDocument") {
    # Treat ThisDocument.cls as a document-module code file (no class export header).
    $content = Strip-ExportAttributeLines $content
  } else {
    $content = Wrap-ClassModuleIfNeeded $content $className
  }

  $content = $content -replace "`r?`n", "`r`n"
  [System.IO.File]::WriteAllText($targetPath, $content, [System.Text.Encoding]::Default)
}

foreach ($f in $frmFiles) {
  $targetPath = Join-Path $outForms ([System.IO.Path]::GetFileName($f.Name))

  # .frm is typically already in VBA export format; copy as text.
  $content = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
  $content = $content -replace "`r?`n", "`r`n"
  [System.IO.File]::WriteAllText($targetPath, $content, [System.Text.Encoding]::Default)

  # If a companion .frx exists, copy it too (binary).
  $frxPath = [System.IO.Path]::ChangeExtension($f.FullName, ".frx")
  if (Test-Path -LiteralPath $frxPath) {
    $frxTarget = Join-Path $outForms ([System.IO.Path]::GetFileName($frxPath))
    Copy-Item -LiteralPath $frxPath -Destination $frxTarget -Force
  }
}

$msWordCount = ($clsFiles | Where-Object { (Get-ModuleNameFromFileName $_.FullName) -ieq "ThisDocument" }).Count
$classModuleCount = ($clsFiles | Where-Object { (Get-ModuleNameFromFileName $_.FullName) -ine "ThisDocument" }).Count

Write-Host ("Generated {0} MsWord object file into: {1}" -f $msWordCount, $outMsWordObjects)
Write-Host ("Generated {0} .frm into: {1}" -f $frmFiles.Count, $outForms)
Write-Host ("Generated {0} .bas into: {1}" -f $vbaFiles.Count, $outModules)
Write-Host ("Generated {0} .cls into: {1}" -f $classModuleCount, $outClassModules)
