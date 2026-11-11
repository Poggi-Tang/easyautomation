param(
    [string]$Target = "",
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
    $launchFile = Join-Path $Target "python\pythonw.exe"
    $installedGui = Join-Path $Target "python\Lib\site-packages\easy_uiauto\qt_ui\window.py"
    if (-not $marker -or $marker.fingerprint -ne $fingerprint -or -not (Test-Path -LiteralPath $launchFile) -or -not (Test-Path -LiteralPath $installedGui)) {
        Write-Host "Installing this GUI source and its local models. The first run needs Internet access."
        & (Join-Path $PSScriptRoot "full_setup.ps1") -Target $Target -NoIntegrate
        if ($LASTEXITCODE -ne 0) { throw "GUI installation failed ($LASTEXITCODE)." }
        @{ fingerprint = $fingerprint; source = $guiRoot } | ConvertTo-Json | Set-Content -LiteralPath $markerPath -Encoding UTF8
    }
    $launchArguments = "-I -X utf8 -m easy_uiauto.library_launcher"
    $workingDirectory = $Target
    $shortcutIcon = Join-Path $Target "python\Lib\site-packages\easy_uiauto\resources\easy-uiauto.ico"
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
