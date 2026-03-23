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

## 3. Install and Verify
After installation, launch `FileOps` from Start Menu and verify:
- Window opens successfully
- You can add source files/folders
- You can run copy/move/rename/delete through the UI
