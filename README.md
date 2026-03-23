# FileOps
Desktop file operation tool with a click-based UI for `copy`, `move`, `rename`, and `delete`.

## Features
- Click-based desktop UI (no command-line required for end users)
- Supports `copy`, `move`, `rename`, `delete`
- Built-in safety defaults: confirmation, workspace boundary checks, conflict policy
- Supports `Dry Run` preview mode in UI
- Supports JSON execution report export
- Windows `.exe` and installer packaging support

## Repository Layout
- `src/fileops/gui.py`: desktop UI
- `src/fileops/operations.py`: core operation logic
- `scripts/build_exe.ps1`: build one-file GUI `fileops.exe`
- `scripts/build_installer.ps1`: build installer via Inno Setup
- `installer/FileOps.iss`: Inno Setup configuration
- `tests/`: pytest tests

## Run From Source
Prerequisites:
- Python 3.11+

Setup:
```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -U pip
pip install . -r requirements-dev.txt
```

Launch UI:
```powershell
python scripts/entrypoint.py
```

Run tests:
```powershell
pytest
```

## Build GUI EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
```
Output:
- `dist/fileops.exe`

## Build Installer
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```
Output:
- `dist/FileOps-Setup.exe`
