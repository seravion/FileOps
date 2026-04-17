# Build and Release (Windows)

## 1. Build GUI EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
```
Expected output:
- `dist/fileops.exe`

## 2. Build Installer
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```
Expected output:
- `dist/FileOps-Setup.exe`

## 3. Verify GUI Features
After launch, verify these operations in the GUI:
- document split by heading levels
- word format by template
- template compare and adjusted output
- optional OCR extraction for image text
