# Release Signing

Sign the `.xll` with the Sectigo code-signing certificate before publishing a GitHub Release.

## Prerequisites

- **signtool.exe** — ships with Visual Studio (ClickOnce SDK):
  `C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe`
- **Sectigo PFX file** — not committed to the repo (`.gitignore` excludes `*.pfx`)

## Steps

1. Build in Release configuration:

   ```
   dotnet build formula-boss/formula-boss.slnx -c Release
   ```

2. Sign the 64-bit XLL:

   ```
   signtool sign /f "path\to\certificate.pfx" /p "pfx-password" /fd sha256 /tr http://timestamp.sectigo.com /td sha256 formula-boss\bin\Release\net6.0-windows\formula-boss64.xll
   ```

3. Verify the signature:

   ```
   signtool verify /pa formula-boss\bin\Release\net6.0-windows\formula-boss64.xll
   ```

## Notes

- The `/tr` and `/td` flags add an RFC 3161 timestamp so the signature remains valid after the certificate expires.
- Only `formula-boss64.xll` needs signing — it's the native shim that Windows evaluates for SmartScreen. The managed DLLs it loads are not checked independently.
- The PFX file and password must never be committed to source control.
