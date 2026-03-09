# Building the Installer

Formula Boss uses [InnoSetup](https://jrsoftware.org/isinfo.php) to create a Windows installer.

## Prerequisites

- **InnoSetup 6.x** — download from https://jrsoftware.org/isdl.php
- **Python 3.10+** with **Pillow** — for regenerating icon/banner assets
- **.NET 6 SDK** — for building the project
- **.NET 6 Desktop Runtime installer** — bundled into the installer for end users

## Directory Structure

```
installer/
  formula-boss.iss       # InnoSetup script
  generate-icons.py      # Generates logo.ico and wizard-banner.bmp from pixel art source
  logo.ico               # Multi-resolution icon (16, 32, 48, 256)
  wizard-banner.bmp      # 164x314 sidebar image for the install wizard
  bundled-runtime/       # .NET runtime installer (not committed, .gitignored)
  output/                # Compiled installer output (.gitignored)
```

## Steps

### 1. Build the project in Release mode

```
dotnet build formula-boss/formula-boss.slnx -c Release
```

### 2. Download the .NET 6 Desktop Runtime installer

Download `windowsdesktop-runtime-6.0.36-win-x64.exe` (or the latest 6.0.x patch) from:
https://dotnet.microsoft.com/en-us/download/dotnet/6.0

Place it in `installer/bundled-runtime/`.

> **Note:** Update the filename in the `[Run]` section of `formula-boss.iss` if using a different patch version.

### 3. Sign the XLL (optional, recommended for distribution)

See [release-signing.md](release-signing.md) for signing instructions. Sign the XLL before building the installer so the signed binary is bundled.

### 4. Regenerate icons (if the source art changed)

```
python installer/generate-icons.py
```

The pixel art source is in `resources/formula-boss icon.html`.

### 5. Compile the installer

Open `installer/formula-boss.iss` in InnoSetup Compiler and press **Build > Compile** (Ctrl+F9), or use the command line:

```
iscc installer/formula-boss.iss
```

The output installer will be at `installer/output/FormulaBoss-{version}-Setup.exe`.

### 6. Sign the installer

```
signtool sign /f "path\to\certificate.pfx" /p "pfx-password" /fd sha256 /tr http://timestamp.sectigo.com /td sha256 installer/output/FormulaBoss-0.1.0-Setup.exe
```

## How It Works

- Installs to `%LOCALAPPDATA%\FormulaBoss` (no admin rights required)
- Registers the XLL via `HKCU\Software\Microsoft\Office\16.0\Excel\Options\OPEN` registry key, so it appears in Excel's **Developer > Excel Add-ins** dialog with a working checkbox
- Removes the registry entry on uninstall (`uninsdeletevalue`)
- Detects running Excel (`CloseApplications=yes`) and prompts the user to close it
- Silently installs the .NET 6 Desktop Runtime if not already present

## Updating the Version

Edit the `#define MyAppVersion` line at the top of `formula-boss.iss`.

## Testing Checklist

- [ ] Fresh install on clean machine (Office only, no .NET 6): runtime installs, add-in loads
- [ ] Install on dev machine: files present, registry entry set, Excel loads add-in
- [ ] Upgrade install (e.g. v0.0.9 → v0.1.0): files updated, no duplicate registry entries
- [ ] Uninstall: files removed, registry entry removed
- [ ] Add-in appears in **Developer > Excel Add-ins** with working checkbox
