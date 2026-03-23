# Build and Release (Windows)

## 1. Create EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
```

Expected output:
- `dist/fileops.exe`

## 2. Create Installer
Install Inno Setup first, then run:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```

Expected output:
- `dist/FileOps-Setup.exe`

## 3. Install and Verify
After installation:
```powershell
fileops --help
fileops copy --help
```
