param(
    [string]$Target = "",
    [switch]$Gpu,
    [ValidatePattern('^cu[0-9]{3}$')][string]$CudaIndex = "cu128",
    [switch]$NoLaunch,
    [switch]$NoShortcut
)
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSHOME "Modules\Microsoft.PowerShell.Utility\Microsoft.PowerShell.Utility.psd1")
[Console]::OutputEncoding = [Text.UTF8Encoding]::new()
$guiRoot = Split-Path -Parent $PSScriptRoot
if (-not [Environment]::Is64BitOperatingSystem -or $env:PROCESSOR_ARCHITECTURE -eq "ARM64") {
    throw "GUI supports Windows x64 only."
}

function Assert-GuiFile([string]$Path, [string]$Digest) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing file: $Path. Extract the complete ZIP again." }
    if ($Digest -and (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $Digest) {
        throw "File checksum mismatch: $Path. Extract the complete ZIP again."
    }
}

if (Test-Path -LiteralPath (Join-Path $guiRoot "gui-distribution.json")) {
    $distribution = Get-Content -LiteralPath (Join-Path $guiRoot "gui-distribution.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($distribution.format -ne 1) { throw "Unsupported GUI distribution." }
    $guiExe = Join-Path $guiRoot "easy-uiauto-library.exe"
    Assert-GuiFile $guiExe $distribution.exe_sha256
    foreach ($required in @("_internal\python313.dll", "_internal\PySide6\Qt6Widgets.dll", "_internal\image-runtime\python.exe")) {
        Assert-GuiFile (Join-Path $guiRoot $required) ""
    }
    $vcRuntime = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" -ErrorAction SilentlyContinue
    if (-not $vcRuntime -or $vcRuntime.Installed -ne 1) {
        $installer = Join-Path $guiRoot "prerequisites\VC_redist.x64.exe"
        Assert-GuiFile $installer $distribution.vc_redist_sha256
        Write-Host "Preparing Microsoft Visual C++ Runtime..."
        $installed = Start-Process -FilePath $installer -ArgumentList "/install", "/quiet", "/norestart" -Verb RunAs -WindowStyle Hidden -Wait -PassThru
        if ($installed.ExitCode -notin @(0, 1638, 3010)) { throw "Visual C++ Runtime setup failed: $($installed.ExitCode)" }
    }
    $launchFile = $guiExe
    $launchArguments = ""
    $workingDirectory = $guiRoot
    $shortcutIcon = $guiExe
    $inferencePython = Join-Path $guiRoot "_internal\image-runtime\python.exe"
} else {
    if (-not (Test-Path -LiteralPath (Join-Path $guiRoot "pyproject.toml"))) {
        throw "Neither a complete GUI ZIP nor a source checkout was found."
    }
    if (-not $Target) { $Target = Join-Path $env:LOCALAPPDATA "easy-uiauto\gui" }
    $Target = [IO.Path]::GetFullPath($Target)
    # Compare actual source contents, including uncommitted edits and ZIP downloads.
    $sourceFiles = @(
        Get-Item -LiteralPath (Join-Path $guiRoot "pyproject.toml"), (Join-Path $guiRoot "MANIFEST.in")
        Get-ChildItem -LiteralPath (Join-Path $guiRoot "src\easy_uiauto"), (Join-Path $guiRoot "scripts") -Recurse -File |
            Where-Object { $_.Extension -in @(".py", ".ps1", ".svg", ".json", ".ico", ".png", ".md") -and $_.FullName -notmatch '[\\/]__pycache__[\\/]' }
    )
    $fingerprintText = ($sourceFiles | Sort-Object FullName | ForEach-Object {
        $_.FullName.Substring($guiRoot.Length) + ":" + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    }) -join "`n"
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $fingerprint = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($fingerprintText))).Replace("-", "").ToLowerInvariant() }
    finally { $sha.Dispose() }
    $markerPath = Join-Path $Target "gui-source.json"
    $marker = $null
    if (Test-Path -LiteralPath $markerPath) {
        try { $marker = Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json } catch { $marker = $null }
    }
    # Reuse a managed runtime or source venv before considering a download.
    $runtimePython = Join-Path $Target "python\python.exe"
    $runtimeHome = Join-Path $Target "python"
    if (Test-Path -LiteralPath (Join-Path $runtimeHome "Scripts\python.exe")) {
        $runtimePython = Join-Path $runtimeHome "Scripts\python.exe"
    }
    $launchFile = Join-Path (Split-Path -Parent $runtimePython) "pythonw.exe"
    $installedGui = Join-Path $Target "python\Lib\site-packages\easy_uiauto\qt_ui\window.py"
    if (-not $marker -or $marker.fingerprint -ne $fingerprint -or -not (Test-Path -LiteralPath $runtimePython) -or -not (Test-Path -LiteralPath $launchFile) -or -not (Test-Path -LiteralPath $installedGui)) {
        if (-not (Test-Path -LiteralPath $runtimePython)) {
            $localPython = $null
            $probe = "import sys,struct,venv; sys.exit(0 if sys.version_info >= (3,11) and struct.calcsize('P') == 8 else 1)"
            foreach ($command in @("py", "python")) {
                $candidate = Get-Command $command -CommandType Application -ErrorAction SilentlyContinue
                if (-not $candidate -or $candidate.Source -match '[\\/]WindowsApps[\\/]') { continue }
                $candidateArgs = @()
                if ($command -eq "py") { $candidateArgs = @("-3") }
                & $candidate.Source @candidateArgs -I -c $probe
                if ($LASTEXITCODE -eq 0) {
                    $localPython = $candidate.Source
                    break
                }
            }
            if ($localPython) {
                if ((Test-Path -LiteralPath $runtimeHome) -and
                    (Get-ChildItem -LiteralPath $runtimeHome -Force | Select-Object -First 1)) {
                    throw "Incomplete Python environment at $runtimeHome. Choose a new -Target or repair that environment."
                }
                Write-Host "Creating a GUI environment with local Python: $localPython (no Python download)."
                & $localPython @candidateArgs -I -m venv $runtimeHome
                if ($LASTEXITCODE -ne 0) { throw "Could not create the GUI Python environment." }
                $runtimePython = Join-Path $runtimeHome "Scripts\python.exe"
                $launchFile = Join-Path $runtimeHome "Scripts\pythonw.exe"
            }
        }
        $dependencyFingerprint = (Get-FileHash -LiteralPath (Join-Path $guiRoot "pyproject.toml") -Algorithm SHA256).Hash
        if (Test-Path -LiteralPath $runtimePython) {
            Write-Host "Updating local GUI source using $runtimePython (no Python download)."
            $installArguments = @("-I", "-m", "pip", "--isolated", "install")
            if ($marker -and $marker.dependencyFingerprint -eq $dependencyFingerprint -and (Test-Path -LiteralPath $installedGui)) {
                # A local source requirement is rebuilt/reinstalled even at the same version.
                $installArguments += @("--no-deps", $guiRoot)
            } else {
                Write-Host "Checking GUI and image dependencies; missing packages may be downloaded."
                $installArguments += ($guiRoot + "[full]")
            }
            & $runtimePython @installArguments
            if ($LASTEXITCODE -ne 0) { throw "GUI source installation failed; the source marker was not updated." }
            & $runtimePython -I -c "from easy_uiauto import qt_gui, local_image_worker"
            if ($LASTEXITCODE -ne 0) { throw "GUI import verification failed; the source marker was not updated." }
        } else {
            Write-Host "No local 64-bit Python 3.11+ found. Preparing the complete runtime and models."
            & (Join-Path $PSScriptRoot "full_setup.ps1") -Target $Target -NoIntegrate
            if ($LASTEXITCODE -ne 0) { throw "GUI installation failed ($LASTEXITCODE)." }
        }
        @{ fingerprint = $fingerprint; dependencyFingerprint = $dependencyFingerprint; source = $guiRoot } |
            ConvertTo-Json | Set-Content -LiteralPath $markerPath -Encoding UTF8
    }
    $launchArguments = "-I -X utf8 -m easy_uiauto.library_launcher"
    $workingDirectory = $Target
    $shortcutIcon = Join-Path $Target "python\Lib\site-packages\easy_uiauto\resources\easy-uiauto.ico"
    $inferencePython = $runtimePython
}

if ($Gpu) {
    Write-Host "Preparing optional NVIDIA GPU inference. Internet access and a compatible NVIDIA driver are required."
    & $inferencePython -I -X utf8 (Join-Path $PSScriptRoot "install_gpu.py") --cuda-index $CudaIndex
    if ($LASTEXITCODE -ne 0) { throw "GPU setup failed. The previous runtime is unchanged. Check the NVIDIA driver or use CPU mode." }
}

if (-not $NoShortcut) {
    $desktop = [Environment]::GetFolderPath("DesktopDirectory")
    if ($desktop -and (Test-Path -LiteralPath $desktop)) {
        try {
            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $desktop "easy-uiauto GUI.lnk"))
            $shortcut.TargetPath = $launchFile
            $shortcut.Arguments = $launchArguments
            $shortcut.WorkingDirectory = $workingDirectory
            $shortcut.IconLocation = $shortcutIcon
            $shortcut.Save()
            # Match the GUI EXE's elevation behavior for the Python shortcut.
            if ($launchArguments) {
                $shortcutPath = Join-Path $desktop "easy-uiauto GUI.lnk"
                $shortcutBytes = [IO.File]::ReadAllBytes($shortcutPath)
                $shortcutBytes[21] = $shortcutBytes[21] -bor 0x20
                [IO.File]::WriteAllBytes($shortcutPath, $shortcutBytes)
            }
            Write-Host "Desktop shortcut ready: easy-uiauto GUI"
        } catch { Write-Warning "Could not create a shortcut. Run setup.bat again to open the GUI. $_" }
    }
}
Write-Host "GUI is ready: $workingDirectory"
if (-not $NoLaunch) {
    $launch = @{ FilePath = $launchFile; WorkingDirectory = $workingDirectory; Verb = "RunAs" }
    if ($launchArguments) { $launch.ArgumentList = $launchArguments }
    Start-Process @launch
}
