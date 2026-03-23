# FileOps
Lightweight and safe file operation toolkit with `copy`, `move`, `rename`, and `delete`.

## Features
- Supports `copy`, `move`, `rename`, `delete` commands
- Built-in safety defaults: confirmation, workspace boundary checks, conflict policy
- Supports `--dry-run` preview mode
- Supports JSON execution report output (`--report`)
- Supports text or JSON logs (`--log-json`)
- Supports packaging to Windows `.exe` and installer

## Repository Layout
- `src/fileops/`: core package and CLI
- `tests/`: pytest tests
- `scripts/build_exe.ps1`: build one-file `fileops.exe` via PyInstaller
- `scripts/build_installer.ps1`: build installer via Inno Setup
- `installer/FileOps.iss`: Inno Setup configuration
- `docs/PRD.md`: product requirements
- `docs/ARCHITECTURE.md`: technical architecture
- `docs/INSTALLATION.md`: build and release steps

## Local Development
Prerequisites:
- Python 3.11+
- Git (for version control and pushing)

Install:
```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -U pip
pip install . -r requirements-dev.txt
```

Run tests:
```powershell
pytest
```

CLI help:
```powershell
fileops --help
fileops copy --help
```

## Command Examples
Copy files to a directory:
```powershell
fileops copy "input\\*.txt" --recursive --dest "output" --workspace "." --overwrite rename --yes
```

Move files:
```powershell
fileops move "staging\\*.csv" --dest "archive" --workspace "." --overwrite never --yes
```

Rename by template:
```powershell
fileops rename "photos\\*.jpg" --pattern "{date}_{index}{ext}" --start-index 1 --workspace "." --yes
```

Delete to trash (default):
```powershell
fileops delete "temp\\*" --recursive --workspace "." --yes
```

Hard delete:
```powershell
fileops delete "temp\\*" --recursive --workspace "." --hard --yes
```

Generate report:
```powershell
fileops copy "input\\*.txt" --dest "output" --workspace "." --yes --report "reports\\copy-report.json"
```

## Build Windows EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
```
Output binary:
- `dist/fileops.exe`

## Build Installer
Prerequisite:
- Inno Setup (`ISCC.exe`)

Build:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```
Installer output:
- `dist/FileOps-Setup.exe`

## Notes
- All source and destination paths must stay inside `--workspace`.
- For destructive operations, use `--yes` for non-interactive CI runs.
- `--overwrite` options for copy/move/rename: `never`, `always`, `rename`.
