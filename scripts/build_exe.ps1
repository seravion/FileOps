param(
    [string]$PythonExe = "python",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

function Resolve-Python {
    param([string]$Requested)

    if ($Requested -and $Requested -ne "python") {
        $cmd = Get-Command $Requested -ErrorAction SilentlyContinue
        if ($cmd) {
            return $cmd.Source
        }
        if (Test-Path $Requested) {
            return (Resolve-Path $Requested).Path
        }
    }

    $whereMatches = @()
    try {
        $whereMatches = where.exe python 2>$null
    }
    catch {
        $whereMatches = @()
    }

    foreach ($item in $whereMatches) {
        if ($item -and $item -notmatch "WindowsApps") {
            return $item
        }
    }

    $localPythonRoot = Join-Path $env:LocalAppData "Programs\\Python"
    if (Test-Path $localPythonRoot) {
        $candidate = Get-ChildItem $localPythonRoot -Recurse -Filter python.exe |
            Where-Object { $_.FullName -notmatch "Lib\\venv" } |
            Select-Object -First 1
        if ($candidate) {
            return $candidate.FullName
        }
    }

    throw "No usable Python executable found. Install Python 3.11+ and retry."
}

function Invoke-Checked {
    param(
        [string]$Command,
        [string[]]$Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $Command $($Arguments -join ' ')"
    }
}

function Test-Module {
    param(
        [string]$PythonPath,
        [string]$Module
    )

    & $PythonPath -c "import importlib.util,sys;sys.exit(0 if importlib.util.find_spec('$Module') else 1)"
    return $LASTEXITCODE -eq 0
}

if ($Clean) {
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
    if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
}

if (Test-Path "build\\fileops_new") {
    Remove-Item -Recurse -Force "build\\fileops_new"
}

if (Test-Path "build\\fileops_new.spec") {
    Remove-Item -Force "build\\fileops_new.spec"
}

if (Test-Path "dist\\fileops_new.exe") {
    Remove-Item -Force "dist\\fileops_new.exe"
}

if (Test-Path "fileops_new.spec") {
    Remove-Item -Force "fileops_new.spec"
}

if (Test-Path "fileops.spec") {
    Remove-Item -Force "fileops.spec"
}

if (Test-Path "dist\\fileops.exe") {
    try {
        Remove-Item -Force "dist\\fileops.exe"
    }
    catch {
        throw "Cannot overwrite dist\\fileops.exe. Please close FileOps and retry."
    }
}

$resolvedPython = Resolve-Python -Requested $PythonExe

if (-not (Test-Path ".venv\\Scripts\\python.exe")) {
    Invoke-Checked -Command $resolvedPython -Arguments @("-m", "venv", ".venv")
}

$venvPython = ".\.venv\Scripts\python.exe"
Invoke-Checked -Command $venvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip")

$requiredModules = @("pytest", "PyInstaller", "hatchling", "send2trash", "docx", "PIL", "pytesseract", "PySide6", "pypdf", "cryptography")
$missing = @()
foreach ($module in $requiredModules) {
    if (-not (Test-Module -PythonPath $venvPython -Module $module)) {
        $missing += $module
    }
}

if ($missing.Count -gt 0) {
    Write-Host "Installing missing modules: $($missing -join ', ')"
    Invoke-Checked -Command $venvPython -Arguments @("-m", "pip", "install", "-r", "requirements-dev.txt")
}
else {
    Write-Host "All required modules already available in .venv"
}

Invoke-Checked -Command $venvPython -Arguments @("-m", "pip", "install", "--no-build-isolation", "--no-deps", ".")

$iconPath = (Resolve-Path "assets\fileops.ico").Path
Invoke-Checked -Command ".\.venv\Scripts\pyinstaller.exe" -Arguments @("--clean", "--noconfirm", "--windowed", "--onefile", "--name", "fileops", "--icon", $iconPath, "--add-data", "$iconPath;assets", "--specpath", "build", "--paths", "src", "scripts/entrypoint.py")

Write-Host "Build completed: dist\\fileops.exe"
