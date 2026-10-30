param(
    [ValidateSet("install", "download")][string]$Action = "install",
    [string]$Target = "",
    [string]$Output = "",
    [string]$Wheel = "",
    [string]$Cache = "",
    [ValidateSet("auto", "codex", "claude", "none")][string]$Client = "auto",
    [switch]$NoIntegrate
)
$ErrorActionPreference = "Stop"
$OutputEncoding = [Text.UTF8Encoding]::new()
[Console]::OutputEncoding = $OutputEncoding
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not [Environment]::Is64BitOperatingSystem -or $env:PROCESSOR_ARCHITECTURE -eq "ARM64") {
    throw "This installer supports Windows x64 only."
}
$packageVersion = "0.9.13"
$repositoryManifest = Join-Path (Split-Path -Parent $PSScriptRoot) "pyproject.toml"
$repositorySource = Join-Path (Split-Path -Parent $PSScriptRoot) "src\easy_uiauto\deployment.py"
if ((Test-Path -LiteralPath $repositoryManifest) -and (Test-Path -LiteralPath $repositorySource)) {
    $manifestText = Get-Content -LiteralPath $repositoryManifest -Raw -Encoding UTF8
    if ($manifestText -match '(?m)^version = "([^"]+)"') { $packageVersion = $Matches[1] }
}
$pythonVersion = "3.13.17"
$pythonSha256 = "b26ca4b4ee51e157836049934d2b8eabd55939be296f3130254251b6df96ec2f"
$deploymentSha256 = "__DEPLOYMENT_SHA256__"
$cacheRoot = $Cache
if (-not $cacheRoot) { $cacheRoot = Join-Path $env:LOCALAPPDATA "easy-uiauto\setup-cache" }
New-Item -ItemType Directory -Path $cacheRoot -Force | Out-Null
$pythonZip = Join-Path $cacheRoot "python.zip"
if (-not (Test-Path -LiteralPath $pythonZip) -or (Get-FileHash -LiteralPath $pythonZip -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pythonSha256) {
    Write-Host "Downloading Python $pythonVersion ..."
    $partialFile = $pythonZip + ".partial"
    Invoke-WebRequest -UseBasicParsing -Uri "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-amd64.zip" -OutFile $partialFile
    if ((Get-FileHash -LiteralPath $partialFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pythonSha256) { throw "Python checksum mismatch" }
    Move-Item -LiteralPath $partialFile -Destination $pythonZip -Force
}
$bootstrapRoot = Join-Path $cacheRoot "bootstrap-python-$pythonVersion"
$bootstrapMarker = Join-Path $bootstrapRoot ".easy-uiauto-runtime"
if (-not (Test-Path -LiteralPath $bootstrapMarker) -or (Get-Content -LiteralPath $bootstrapMarker -Raw).Trim() -ne $pythonSha256) {
    Expand-Archive -LiteralPath $pythonZip -DestinationPath $bootstrapRoot -Force
    Set-Content -LiteralPath $bootstrapMarker -Value $pythonSha256 -Encoding ASCII -NoNewline
}
$repositorySource = Join-Path (Split-Path -Parent $PSScriptRoot) "src\easy_uiauto\deployment.py"
if (Test-Path -LiteralPath $repositorySource) {
    $deploymentFile = $repositorySource
} else {
    if ($deploymentSha256 -eq "__DEPLOYMENT_SHA256__") { throw "Download the released easy-uiauto-setup.ps1 or use a repository checkout." }
    $deploymentFile = Join-Path $cacheRoot "deployment-$packageVersion.py"
    if (-not (Test-Path -LiteralPath $deploymentFile) -or (Get-FileHash -LiteralPath $deploymentFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne $deploymentSha256) {
        Invoke-WebRequest -UseBasicParsing -Uri "https://raw.githubusercontent.com/Poggi-Tang/easyautomation/v$packageVersion/src/easy_uiauto/deployment.py" -OutFile $deploymentFile
        if ((Get-FileHash -LiteralPath $deploymentFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne $deploymentSha256) { throw "Setup script checksum mismatch" }
    }
}
$setupArguments = @("-I", "-X", "utf8", $deploymentFile, $Action, "--version", $packageVersion, "--cache", $cacheRoot)
if ($Wheel) { $setupArguments += @("--wheel", $Wheel) }
if ($Action -eq "download") {
    if (-not $Output) { $Output = Join-Path $env:USERPROFILE "Downloads\easy-uiauto-offline" }
    $setupArguments += $Output
} else {
    $setupArguments += @("--client", $Client)
    if ($Target) { $setupArguments += @("--target", $Target) }
    if ($NoIntegrate) { $setupArguments += "--no-integrate" }
}
& (Join-Path $bootstrapRoot "python.exe") @setupArguments
if ($LASTEXITCODE -ne 0) { throw "Setup failed ($LASTEXITCODE). See $cacheRoot\download.log" }
