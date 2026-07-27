$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$OfficeKitVersion = "0.5.0"
$OfficeKitRepository = "w31r4/OfficeKit"

function Fail([string] $Message) {
  throw "OfficeKit installer: $Message"
}

function Get-FullPath([string] $Value) {
  if ([string]::IsNullOrWhiteSpace($Value)) {
    Fail "an installation path is empty."
  }
  return [System.IO.Path]::GetFullPath($Value)
}

function Assert-RealDirectory([string] $Path, [string] $Label) {
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
  $item = Get-Item -LiteralPath $Path -Force
  if (
    -not $item.PSIsContainer -or
    (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
  ) {
    Fail "$Label must be a real directory: $Path"
  }
}

function Assert-RegularFile([string] $Path, [string] $Label) {
  $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
  if (
    $item.PSIsContainer -or
    (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
  ) {
    Fail "$Label must be a regular non-reparse file: $Path"
  }
}

function Assert-PathInside([string] $Path, [string] $Root, [string] $Label) {
  $fullPath = [System.IO.Path]::GetFullPath($Path)
  $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd(
    [char[]]@(
      [System.IO.Path]::DirectorySeparatorChar,
      [System.IO.Path]::AltDirectorySeparatorChar
    )
  )
  if (
    -not $fullPath.StartsWith(
      "$fullRoot$([System.IO.Path]::DirectorySeparatorChar)",
      [System.StringComparison]::OrdinalIgnoreCase
    )
  ) {
    Fail "$Label escapes its managed root."
  }
}

function Get-ExpectedTarget() {
  if (-not [System.Environment]::Is64BitOperatingSystem) {
    Fail "no self-contained 32-bit Windows build is available."
  }
  return "win32-x64"
}

function Get-Sha256([string] $Path) {
  $stream = [System.IO.File]::OpenRead($Path)
  $algorithm = [System.Security.Cryptography.SHA256]::Create()
  try {
    $hash = $algorithm.ComputeHash($stream)
    return ([System.BitConverter]::ToString($hash).Replace("-", "")).ToLowerInvariant()
  } finally {
    $algorithm.Dispose()
    $stream.Dispose()
  }
}

function Test-ZipEntryPath([string] $Value) {
  if (
    [string]::IsNullOrWhiteSpace($Value) -or
    $Value.Contains("\") -or
    $Value.IndexOf([char] 0) -ge 0 -or
    $Value.StartsWith("/") -or
    $Value -match "^[A-Za-z]:"
  ) {
    return $false
  }
  $segments = $Value.Split("/")
  foreach ($segment in $segments) {
    if ([string]::IsNullOrWhiteSpace($segment) -or $segment -eq "." -or $segment -eq "..") {
      return $false
    }
  }
  return $true
}

function Expand-VerifiedZip([string] $Archive, [string] $Destination, [string] $ExpectedRoot) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::OpenRead($Archive)
  try {
    if ($zip.Entries.Count -eq 0) {
      Fail "archive contains no entries."
    }
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($entry in $zip.Entries) {
      $raw = $entry.FullName
      $relative = $raw.TrimEnd("/")
      if (-not (Test-ZipEntryPath $relative)) {
        Fail "archive contains an unsafe path: $raw"
      }
      if ($relative -ne $ExpectedRoot -and -not $relative.StartsWith("$ExpectedRoot/", [System.StringComparison]::Ordinal)) {
        Fail "archive entry is outside ${ExpectedRoot}: $raw"
      }
      if (-not $seen.Add($relative)) {
        Fail "archive contains a duplicate path: $relative"
      }
      $mode = (($entry.ExternalAttributes -shr 16) -band 0xffff)
      if (($mode -band 0xf000) -eq 0xa000) {
        Fail "archive contains a symlink: $raw"
      }
    }

    Assert-RealDirectory $Destination "archive extraction root"
    foreach ($entry in $zip.Entries) {
      if ($entry.FullName.EndsWith("/")) {
        continue
      }
      $relative = $entry.FullName.TrimEnd("/")
      $parts = $relative.Split("/")
      $output = [System.IO.Path]::GetFullPath(
        [System.IO.Path]::Combine($Destination, [string]::Join([System.IO.Path]::DirectorySeparatorChar, $parts))
      )
      Assert-PathInside $output $Destination "archive entry"
      $parent = [System.IO.Path]::GetDirectoryName($output)
      Assert-RealDirectory $parent "archive output directory"
      $inputStream = $entry.Open()
      try {
        $outputStream = [System.IO.File]::Open(
          $output,
          [System.IO.FileMode]::CreateNew,
          [System.IO.FileAccess]::Write,
          [System.IO.FileShare]::None
        )
        try {
          $inputStream.CopyTo($outputStream)
        } finally {
          $outputStream.Dispose()
        }
      } finally {
        $inputStream.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function Invoke-InstallationVerification(
  [string] $Root,
  [string] $Target,
  [switch] $AllowReceipt
) {
  $node = Join-Path $Root "runtime\node\node.exe"
  $verifier = Join-Path $Root "lib\verify-install.mjs"
  $entry = Join-Path $Root "app\node_modules\office-kit\bin\officekit.mjs"
  Assert-RegularFile $node "bundled Node executable"
  Assert-RegularFile $verifier "installation verifier"
  Assert-RegularFile $entry "OfficeKit command entrypoint"
  $arguments = @($verifier, $Root, $OfficeKitVersion, $Target)
  if ($AllowReceipt) {
    $arguments += "--allow-receipt"
  }
  & $node @arguments | Out-Null
  if ($LASTEXITCODE -ne 0) {
    Fail "installed file integrity verification failed."
  }
  $actualVersion = (& $node $entry --version).Trim()
  if ($LASTEXITCODE -ne 0 -or $actualVersion -ne $OfficeKitVersion) {
    Fail "bundled OfficeKit failed its version probe."
  }
}

function Ensure-Launcher([string] $BinRoot, [string] $InstallRoot) {
  $launcherPath = Join-Path $BinRoot "officekit.cmd"
  $launcher = @'
@echo off
setlocal EnableExtensions
set "ROOT=%~dp0.."
set /p VERSION=<"%ROOT%\current.version"
if "%VERSION%"=="" (
  echo OfficeKit installation is incomplete: active version is missing. 1>&2
  exit /b 1
)
set "NODE=%ROOT%\versions\%VERSION%\runtime\node\node.exe"
set "ENTRY=%ROOT%\versions\%VERSION%\app\node_modules\office-kit\bin\officekit.mjs"
if not exist "%NODE%" (
  echo OfficeKit installation is incomplete: bundled Node is missing. 1>&2
  exit /b 1
)
if not exist "%ENTRY%" (
  echo OfficeKit installation is incomplete: command entrypoint is missing. 1>&2
  exit /b 1
)
"%NODE%" "%ENTRY%" %*
exit /b %ERRORLEVEL%
'@
  if (Test-Path -LiteralPath $launcherPath) {
    Assert-RegularFile $launcherPath "OfficeKit command"
    $existing = [System.IO.File]::ReadAllText($launcherPath)
    if ($existing -ne $launcher) {
      Fail "$launcherPath belongs to another installation."
    }
    return
  }
  [System.IO.File]::WriteAllText(
    $launcherPath,
    $launcher,
    [System.Text.UTF8Encoding]::new($false)
  )
}

function Add-UserPath([string] $BinRoot, [bool] $Persist) {
  $existing = [System.Environment]::GetEnvironmentVariable("Path", "User")
  $parts = @()
  if (-not [string]::IsNullOrWhiteSpace($existing)) {
    $parts = $existing.Split(";") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  }
  $alreadyPresent = $parts | Where-Object {
    [string]::Equals($_, $BinRoot, [System.StringComparison]::OrdinalIgnoreCase)
  }
  if ($Persist -and $null -eq $alreadyPresent) {
    $next = (@($parts) + @($BinRoot)) -join ";"
    [System.Environment]::SetEnvironmentVariable("Path", $next, "User")
  }
  if (-not (($env:Path -split ";") | Where-Object {
    [string]::Equals($_, $BinRoot, [System.StringComparison]::OrdinalIgnoreCase)
  })) {
    $env:Path = "$BinRoot;$env:Path"
  }
}

$target = if ($env:OFFICE_KIT_INSTALL_TEST -eq "1") {
  if ([string]::IsNullOrWhiteSpace($env:OFFICE_KIT_TEST_TARGET)) {
    Fail "OFFICE_KIT_TEST_TARGET is required in test mode."
  }
  $env:OFFICE_KIT_TEST_TARGET
} else {
  Get-ExpectedTarget
}
if ($target -ne "win32-x64") {
  Fail "unsupported target $target."
}

$expectedSha256 = "0f3ae269edf6b52d5e57d6cc69e9d9f95e995ce607fd1b736b751f9eef6f0d61"
$expectedSize = 80459043
if ($env:OFFICE_KIT_INSTALL_TEST -eq "1") {
  if ([string]::IsNullOrWhiteSpace($env:OFFICE_KIT_TEST_ARCHIVE)) {
    Fail "OFFICE_KIT_TEST_ARCHIVE is required in test mode."
  }
  if ([string]::IsNullOrWhiteSpace($env:OFFICE_KIT_TEST_SHA256)) {
    Fail "OFFICE_KIT_TEST_SHA256 is required in test mode."
  }
  if ([string]::IsNullOrWhiteSpace($env:OFFICE_KIT_TEST_SIZE)) {
    Fail "OFFICE_KIT_TEST_SIZE is required in test mode."
  }
  $expectedSha256 = $env:OFFICE_KIT_TEST_SHA256
  $expectedSize = [Int64] $env:OFFICE_KIT_TEST_SIZE
}
if ($expectedSha256 -notmatch "^[a-f0-9]{64}$" -or $expectedSize -le 0) {
  Fail "release identity is not finalized."
}

if ($env:OFFICE_KIT_HOME) {
  $installRoot = Get-FullPath $env:OFFICE_KIT_HOME
} else {
  $installRoot = Get-FullPath (Join-Path $env:LOCALAPPDATA "OfficeKit")
}
if ($env:OFFICE_KIT_BIN_DIR) {
  $binRoot = Get-FullPath $env:OFFICE_KIT_BIN_DIR
} else {
  $binRoot = Get-FullPath (Join-Path $installRoot "bin")
}
$versionsRoot = Join-Path $installRoot "versions"
$versionRoot = Join-Path $versionsRoot $OfficeKitVersion
$currentPath = Join-Path $installRoot "current.version"

Assert-RealDirectory $installRoot "OfficeKit home"
Assert-RealDirectory $versionsRoot "OfficeKit versions directory"
Assert-RealDirectory $binRoot "OfficeKit command directory"
Assert-PathInside $versionsRoot $installRoot "versions directory"
Assert-PathInside $versionRoot $versionsRoot "version directory"
Assert-PathInside $binRoot $installRoot "command directory"

$temporary = Join-Path $installRoot (".install." + [Guid]::NewGuid().ToString("N"))
Assert-RealDirectory $temporary "installation transaction"
try {
  $asset = "office-kit-$OfficeKitVersion-$target.zip"
  $archive = Join-Path $temporary $asset
  if ($env:OFFICE_KIT_INSTALL_TEST -eq "1") {
    Copy-Item -LiteralPath $env:OFFICE_KIT_TEST_ARCHIVE -Destination $archive -ErrorAction Stop
  } else {
    $url = "https://github.com/$OfficeKitRepository/releases/download/v$OfficeKitVersion/$asset"
    Invoke-WebRequest -Uri $url -OutFile $archive -UseBasicParsing
  }

  Assert-RegularFile $archive "downloaded archive"
  $actualSize = (Get-Item -LiteralPath $archive).Length
  if ($actualSize -ne $expectedSize) {
    Fail "archive size is $actualSize; expected $expectedSize."
  }
  $actualSha256 = Get-Sha256 $archive
  if ($actualSha256 -ne $expectedSha256) {
    Fail "archive SHA-256 is $actualSha256; expected $expectedSha256."
  }

  $archiveRoot = "office-kit-$OfficeKitVersion-$target"
  if (Test-Path -LiteralPath $versionRoot) {
    Assert-RealDirectory $versionRoot "existing version path"
    $receipt = Join-Path $versionRoot ".office-kit-install-receipt"
    Assert-RegularFile $receipt "existing version receipt"
    $expectedReceipt = "office-kit.standalone-install.v1 $OfficeKitVersion $target $expectedSha256 $expectedSize"
    if ([System.IO.File]::ReadAllText($receipt).Trim() -ne $expectedReceipt) {
      Fail "existing version receipt does not match this release."
    }
    Invoke-InstallationVerification $versionRoot $target -AllowReceipt
  } else {
    $extraction = Join-Path $temporary "extracted"
    Expand-VerifiedZip $archive $extraction $archiveRoot
    $candidate = Join-Path $extraction $archiveRoot
    Assert-RealDirectory $candidate "archive root"
    Invoke-InstallationVerification $candidate $target
    $receipt = Join-Path $candidate ".office-kit-install-receipt"
    [System.IO.File]::WriteAllText(
      $receipt,
      "office-kit.standalone-install.v1 $OfficeKitVersion $target $expectedSha256 $expectedSize" + [Environment]::NewLine,
      [System.Text.UTF8Encoding]::new($false)
    )
    Move-Item -LiteralPath $candidate -Destination $versionRoot -ErrorAction Stop
  }

  $nextCurrent = Join-Path $installRoot (".current.next." + [Guid]::NewGuid().ToString("N"))
  [System.IO.File]::WriteAllText(
    $nextCurrent,
    "$OfficeKitVersion" + [Environment]::NewLine,
    [System.Text.UTF8Encoding]::new($false)
  )
  try {
    if (Test-Path -LiteralPath $currentPath) {
      Assert-RegularFile $currentPath "active version record"
      $previousCurrent = Join-Path $installRoot (".current.previous." + [Guid]::NewGuid().ToString("N"))
      try {
        [System.IO.File]::Replace($nextCurrent, $currentPath, $previousCurrent)
      } finally {
        if (Test-Path -LiteralPath $previousCurrent) {
          Assert-RegularFile $previousCurrent "previous active version record"
          Remove-Item -LiteralPath $previousCurrent -Force
        }
      }
    } else {
      [System.IO.File]::Move($nextCurrent, $currentPath)
    }
  } finally {
    if (Test-Path -LiteralPath $nextCurrent) {
      Assert-RegularFile $nextCurrent "pending active version record"
      Remove-Item -LiteralPath $nextCurrent -Force
    }
  }
  Ensure-Launcher $binRoot $installRoot
  Add-UserPath $binRoot ($env:OFFICE_KIT_INSTALL_TEST -ne "1")
} finally {
  if (Test-Path -LiteralPath $temporary) {
    Remove-Item -LiteralPath $temporary -Force -Recurse
  }
}

Write-Output "OfficeKit $OfficeKitVersion installed at $versionRoot"
Write-Output "Command: officekit"
