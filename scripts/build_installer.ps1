param(
    [string]$IsccPath = "iscc"
)

$ErrorActionPreference = "Stop"

function Resolve-Iscc {
    param([string]$Requested)

    if ($Requested -and $Requested -ne "iscc") {
        $cmd = Get-Command $Requested -ErrorAction SilentlyContinue
        if ($cmd) {
            return $cmd.Source
        }
        if (Test-Path $Requested) {
            return (Resolve-Path $Requested).Path
        }
    }

    $defaultCandidates = @(
        "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
        "C:\\Program Files\\Inno Setup 6\\ISCC.exe",
        "$env:LocalAppData\\Programs\\Inno Setup 6\\ISCC.exe"
    )

    foreach ($candidate in $defaultCandidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $cmd2 = Get-Command iscc -ErrorAction SilentlyContinue
    if ($cmd2) {
        return $cmd2.Source
    }

    throw "ISCC.exe was not found. Install Inno Setup 6 and retry."
}

if (-not (Test-Path "dist\\fileops.exe")) {
    throw "dist\\fileops.exe not found. Run scripts/build_exe.ps1 first."
}

$resolvedIscc = Resolve-Iscc -Requested $IsccPath
& $resolvedIscc "installer\\FileOps.iss"
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed with exit code $LASTEXITCODE"
}

Write-Host "Installer build completed."
