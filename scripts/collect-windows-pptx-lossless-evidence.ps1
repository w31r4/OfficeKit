[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$SourceRoot,

  [Parameter(Mandatory = $true)]
  [string]$OutputRoot,

  [Parameter(Mandatory = $true)]
  [string]$EvidencePath,

  [string]$Commit
)

$ErrorActionPreference = "Stop"

function Fail([string]$Message) {
  throw "Windows PPTX evidence collection failed: $Message"
}

function Ask-Confirmed([string]$Prompt) {
  $answer = Read-Host "$Prompt [y/N]"
  return $answer -match '^(?i:y|yes)$'
}

function Require-AbsoluteWindowsPath([string]$Value, [string]$Label) {
  if ($Value -notmatch '^[A-Za-z]:[\\/].+') {
    Fail "$Label must be an absolute Windows path"
  }
  return [IO.Path]::GetFullPath($Value)
}

function Get-Sha256([string]$Path) {
  return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Export-SlideHashes($Presentation, [string]$Directory, [string]$Prefix) {
  New-Item -ItemType Directory -Force -Path $Directory | Out-Null
  $hashes = @{}
  for ($page = 1; $page -le $Presentation.Slides.Count; $page++) {
    $imagePath = Join-Path $Directory ("{0}-{1}.png" -f $Prefix, $page)
    $Presentation.Slides.Item($page).Export($imagePath, "PNG", 1600, 900)
    if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) {
      Fail "PowerPoint did not export slide $page for $Prefix"
    }
    $hashes[$page] = Get-Sha256 $imagePath
  }
  return $hashes
}

if (-not [Environment]::Is64BitOperatingSystem) { Fail "a 64-bit Windows host is required" }
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { Fail "a Windows host is required" }

$SourceRoot = Require-AbsoluteWindowsPath $SourceRoot "SourceRoot"
$OutputRoot = Require-AbsoluteWindowsPath $OutputRoot "OutputRoot"
$EvidencePath = Require-AbsoluteWindowsPath $EvidencePath "EvidencePath"
if (-not $Commit) {
  $Commit = (& git rev-parse HEAD).Trim()
}
if ($Commit -notmatch '^[0-9a-f]{40}$|^[0-9a-f]{64}$') { Fail "Commit must be a full Git commit hash" }

$sources = @(
  @{ id = "suanzhi-future-2026"; fileName = "b34ddad8cf8b_012_算秩未来2026_0127_极致技术&长期主义.pptx"; sha256 = "b34ddad8cf8bbd083b60e07f8488267b1a0e4199db422468faa0eeb5d83e1762"; slides = 21; targetPage = 1; targetNodeId = "presentation/slide/1/element/1" },
  @{ id = "blue-gray-acid-template"; fileName = "template.pptx"; sha256 = "558ce85c0d64cd2a06faf88d6a4aa331e8cd4c685c59101c835ded2fbc87696d"; slides = 19; targetPage = 1; targetNodeId = "presentation/slide/1/element/6" },
  @{ id = "mckinsey-customer-loyalty"; fileName = "ppt169_麦肯锡风_kimsoong_customer_loyalty.pptx"; sha256 = "e0bfb89454f51c400ac03797c255aa93919328ff8dba36fe414e5bcfed0536c5"; slides = 8; targetPage = 1; targetNodeId = "presentation/slide/1/element/1" }
)

$evidenceDirectory = Split-Path -Parent $EvidencePath
New-Item -ItemType Directory -Force -Path $evidenceDirectory | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$powerPoint = $null
$powerPointVersion = $null
$results = @()
$pageComparisons = @()
$openedPresentations = @()

try {
  try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = $false
    $powerPointVersion = [string]$powerPoint.Version
  } catch {
    Fail "Microsoft PowerPoint COM automation is unavailable: $($_.Exception.Message)"
  }

  foreach ($source in $sources) {
    $sourcePath = Require-AbsoluteWindowsPath (Join-Path $SourceRoot $source.fileName) "$($source.id) sourcePath"
    $candidatePath = Require-AbsoluteWindowsPath (Join-Path $OutputRoot "$($source.id).pptx") "$($source.id) candidate outputPath"
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { Fail "$($source.id) source does not exist: $sourcePath" }
    if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) { Fail "$($source.id) output does not exist: $candidatePath" }
    if ($sourcePath -eq $candidatePath) { Fail "$($source.id) output must not overwrite its source" }
    if ((Get-Sha256 $sourcePath) -ne $source.sha256) { Fail "$($source.id) source SHA-256 does not match the frozen manifest" }

    $sourcePresentation = $null
    $candidatePresentation = $null
    $savedCopyPresentation = $null
    try {
      $sourcePresentation = $powerPoint.Presentations.Open($sourcePath, $true, $false, $false)
      $candidatePresentation = $powerPoint.Presentations.Open($candidatePath, $true, $false, $false)
      if ($sourcePresentation.Slides.Count -ne $source.slides) { Fail "$($source.id) source slide count changed" }
      if ($candidatePresentation.Slides.Count -ne $source.slides) { Fail "$($source.id) output slide count differs from the source" }

      $copyPath = Join-Path $evidenceDirectory "$($source.id)-powerpoint-copy-$stamp.pptx"
      $candidatePresentation.SaveCopyAs($copyPath)
      if (-not (Test-Path -LiteralPath $copyPath -PathType Leaf)) { Fail "$($source.id) PowerPoint did not save a copy" }
      $savedCopyPresentation = $powerPoint.Presentations.Open($copyPath, $true, $false, $false)
      if ($savedCopyPresentation.Slides.Count -ne $source.slides) { Fail "$($source.id) saved copy cannot be reopened" }

      $sourceImages = Join-Path $evidenceDirectory "$($source.id)-source-pages"
      $outputImages = Join-Path $evidenceDirectory "$($source.id)-output-pages"
      $sourceHashes = Export-SlideHashes $sourcePresentation $sourceImages "source"
      $outputHashes = Export-SlideHashes $savedCopyPresentation $outputImages "output"
      for ($page = 1; $page -le $source.slides; $page++) {
        $target = $page -eq $source.targetPage
        $identical = $sourceHashes[$page] -eq $outputHashes[$page]
        $pageComparisons += [ordered]@{
          sourceId = $source.id
          page = $page
          target = $target
          pixelIdentical = $identical
          sourcePixelSha256 = $sourceHashes[$page]
          outputPixelSha256 = $outputHashes[$page]
        }
      }
      if (($pageComparisons | Where-Object { $_.sourceId -eq $source.id -and $_.target -and $_.pixelIdentical }).Count -gt 0) {
        Fail "$($source.id) target page did not show a PowerPoint pixel delta"
      }

      $checks = [ordered]@{
        opened = (Ask-Confirmed "$($source.id): did both source and output open in PowerPoint without an error?")
        noRepairPrompt = (Ask-Confirmed "$($source.id): was there no repair/recovery prompt for either file?")
        browsedAllSlides = (Ask-Confirmed "$($source.id): did you browse every slide in both presentations?")
        targetEditVisible = (Ask-Confirmed "$($source.id): is the declared target edit visibly present in the saved copy?")
        nonTargetPagesPixelIdentical = (($pageComparisons | Where-Object { $_.sourceId -eq $source.id -and -not $_.target } | Where-Object { -not $_.pixelIdentical }).Count -eq 0)
        advancedObjectsPreserved = (Ask-Confirmed "$($source.id): are the source's advanced objects still present and usable?")
        savedCopy = $true
        reopenedCopy = $true
        sourceProtected = ((Get-Sha256 $sourcePath) -eq $source.sha256)
        unsupportedCapabilityFailClosed = (Ask-Confirmed "$($source.id): did an unsupported capability refuse safely without changing the file?")
      }
      if (-not ($checks.Values -contains $false)) {
        $results += [ordered]@{
          id = $source.id
          sourceSha256 = $source.sha256
          sourcePath = $sourcePath
          outputPath = [IO.Path]::GetFullPath($copyPath)
          target = [ordered]@{ nodeId = $source.targetNodeId; operation = "native-leaf-edit" }
          checks = $checks
          evidencePath = [IO.Path]::GetFullPath($EvidencePath)
        }
      } else {
        Fail "$($source.id) has a failed human confirmation; do not publish partial evidence"
      }
    } finally {
      if ($savedCopyPresentation) { $savedCopyPresentation.Close() }
      if ($candidatePresentation) { $candidatePresentation.Close() }
      if ($sourcePresentation) { $sourcePresentation.Close() }
    }
  }
} finally {
  if ($powerPoint) { $powerPoint.Quit() }
}

$observedAt = (Get-Date).ToString("o")
$evidence = [ordered]@{
  schema = "office-kit.windows-pptx-lossless-evidence.v1"
  method = "human-observed-windows-powerpoint"
  checkedAt = $observedAt
  commit = $Commit
  host = [ordered]@{
    platform = "win32-x64"
    observedAt = $observedAt
    powerpoint = [ordered]@{ installed = $true; version = $powerPointVersion }
  }
  visualReview = [ordered]@{
    observedAt = $observedAt
    renderer = "Microsoft PowerPoint"
    pagesCompared = $pageComparisons.Count
    evidencePath = [IO.Path]::GetFullPath($EvidencePath)
    pageComparisons = $pageComparisons
  }
  sources = $results
}

$json = $evidence | ConvertTo-Json -Depth 20
$utf8NoBom = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
[IO.File]::WriteAllText($EvidencePath, $json, $utf8NoBom)
Write-Output ("Wrote Windows PPTX lossless evidence: {0}" -f $EvidencePath)
