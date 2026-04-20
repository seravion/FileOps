# Folio Runtime Bundle

Place the Folio CLI binary used by FileOps at:

- `vendor/folio/bin/scribe-cli.exe` (Windows)

You can generate this binary from the local `Folio-master` source tree by running:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_folio_cli.ps1 -Release
```

`scripts/build_exe.ps1` will bundle this binary into `fileops.exe` under the `folio/` runtime directory.
