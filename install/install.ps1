param(
  [string]$Version = $env:OFFICEGEN_VERSION,
  [string]$Repo = $(if ($env:OFFICEGEN_REPO) { $env:OFFICEGEN_REPO } else { "Aero123421/officegen-CLI" }),
  [string]$InstallDir = $(if ($env:OFFICEGEN_INSTALL_DIR) { $env:OFFICEGEN_INSTALL_DIR } else { Join-Path $env:LOCALAPPDATA "Officegen\bin" })
)

$ErrorActionPreference = "Stop"

function Resolve-OfficegenVersion {
  param([string]$RequestedVersion, [string]$Repository)
  if ($RequestedVersion) {
    return $RequestedVersion.TrimStart("v")
  }

  $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repository/releases/latest" -Headers @{ "User-Agent" = "officegen-installer" }
  if (-not $release.tag_name) {
    throw "Could not resolve latest release for $Repository."
  }
  return ([string]$release.tag_name).TrimStart("v")
}

function Resolve-OfficegenTarget {
  if ($PSVersionTable.PSEdition -eq "Core" -and -not $IsWindows) {
    throw "install.ps1 supports Windows. Use install/install.sh on macOS or Linux."
  }

  switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { return "x86_64-pc-windows-msvc" }
    "ARM64" { return "aarch64-pc-windows-msvc" }
    default { throw "Unsupported Windows architecture: $env:PROCESSOR_ARCHITECTURE" }
  }
}

function Set-OfficegenPathPrecedence {
  param([string]$Directory)

  $currentSession = ($env:PATH -split ";") | Where-Object { $_ -and ($_ -ne $Directory) }
  $env:PATH = (@($Directory) + $currentSession) -join ";"

  $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
  $userParts = @()
  if ($userPath) {
    $userParts = ($userPath -split ";") | Where-Object { $_ -and ($_ -ne $Directory) }
  }
  [Environment]::SetEnvironmentVariable("Path", (@($Directory) + $userParts) -join ";", "User")
}

function Test-OfficegenCommandResolution {
  param([string]$ExpectedExe, [string]$ExpectedVersion)

  $commands = @(Get-Command officegen -All -ErrorAction SilentlyContinue)
  $conflicts = @($commands | Where-Object { $_.Source -and ([System.IO.Path]::GetFullPath($_.Source) -ne [System.IO.Path]::GetFullPath($ExpectedExe)) })
  if ($conflicts.Count -gt 0) {
    Write-Warning "Detected older or competing officegen commands earlier/on PATH:"
    $conflicts | ForEach-Object { Write-Warning ("  " + $_.Source) }
  }

  $resolved = Get-Command officegen -ErrorAction SilentlyContinue
  if (-not $resolved -or ([System.IO.Path]::GetFullPath($resolved.Source) -ne [System.IO.Path]::GetFullPath($ExpectedExe))) {
    throw "officegen command still resolves to '$($resolved.Source)' instead of '$ExpectedExe'. Restart the shell or remove the old npm shim (for example AppData\\Roaming\\npm\\officegen.ps1)."
  }
  $actualVersion = (& officegen --version).Trim()
  if ($actualVersion -ne $ExpectedVersion) {
    throw "officegen --version returned $actualVersion, expected $ExpectedVersion. A stale shim is still winning PATH resolution."
  }
}

$resolvedVersion = Resolve-OfficegenVersion -RequestedVersion $Version -Repository $Repo
$target = Resolve-OfficegenTarget
$asset = "officegen-v$resolvedVersion-$target.zip"
$baseUrl = "https://github.com/$Repo/releases/download/v$resolvedVersion"
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("officegen-install-" + [System.Guid]::NewGuid().ToString("N"))
$archive = Join-Path $tempDir $asset
$checksum = "$archive.sha256"
$extractDir = Join-Path $tempDir "extract"

try {
  New-Item -ItemType Directory -Force -Path $tempDir, $extractDir, $InstallDir | Out-Null
  Write-Host "Installing officegen v$resolvedVersion for $target"
  Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/$asset" -OutFile $archive
  Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/$asset.sha256" -OutFile $checksum

  $expected = ((Get-Content -Raw $checksum).Trim() -split "\s+")[0].ToLowerInvariant()
  $actual = (Get-FileHash -Algorithm SHA256 $archive).Hash.ToLowerInvariant()
  if ($expected -ne $actual) {
    throw "Checksum mismatch for $asset."
  }

  Expand-Archive -Path $archive -DestinationPath $extractDir -Force
  $binary = Get-ChildItem -Path $extractDir -Recurse -File -Filter "officegen.exe" | Select-Object -First 1
  if (-not $binary) {
    throw "Archive did not contain officegen.exe."
  }

  $destination = Join-Path $InstallDir "officegen.exe"
  Copy-Item -Force -Path $binary.FullName -Destination $destination
  Write-Host "Installed $destination"

  Set-OfficegenPathPrecedence -Directory $InstallDir
  Test-OfficegenCommandResolution -ExpectedExe $destination -ExpectedVersion $resolvedVersion
  Write-Host "officegen v$resolvedVersion is first on PATH for this session and future user sessions."
}
finally {
  Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $tempDir
}
