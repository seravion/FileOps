param(
    [string]$FolioProject = "Folio-master",
    [string]$OutputDir = "vendor\folio\bin",
    [switch]$Release
)

$ErrorActionPreference = "Stop"

function Resolve-Cargo {
    $cmd = Get-Command cargo -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }
    throw "cargo not found. Install Rust toolchain first."
}

$cargo = Resolve-Cargo
$projectPath = (Resolve-Path $FolioProject).Path
$targetProfile = if ($Release) { "release" } else { "debug" }

Push-Location $projectPath
try {
    $buildArgs = @("build", "-p", "scribe-cli")
    if ($Release) {
        $buildArgs += "--release"
    }
    & $cargo @buildArgs
    if ($LASTEXITCODE -ne 0) {
        throw "cargo build failed for scribe-cli."
    }
}
finally {
    Pop-Location
}

$binaryPath = Join-Path $projectPath "target\$targetProfile\scribe-cli.exe"
if (-not (Test-Path $binaryPath)) {
    throw "Built binary not found: $binaryPath"
}

$outputPath = Join-Path (Get-Location).Path $OutputDir
$null = New-Item -Path $outputPath -ItemType Directory -Force
Copy-Item -Path $binaryPath -Destination (Join-Path $outputPath "scribe-cli.exe") -Force

Write-Host "Folio CLI prepared at: $outputPath\scribe-cli.exe"
