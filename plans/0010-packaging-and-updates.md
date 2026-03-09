# 0010 — Packaging and Updates — Implementation Plan

## Overview

Implement the three components from spec 0010: version infrastructure with About dialog, in-app update checking, and an InnoSetup installer. Work is split into three independent PRs that can be developed in parallel, plus a final release PR.

## PR 1: Version Infrastructure and About Dialog

### Files to Touch

- `Directory.Build.props` — Add `<Version>0.1.0</Version>` property
- `formula-boss/UI/AboutDialog.xaml` (new) — About dialog UI
- `formula-boss/UI/AboutDialog.xaml.cs` (new) — About dialog code-behind
- `formula-boss/UI/RibbonController.cs` — Add About button to ribbon XML and callback
- `formula-boss/UI/FloatingEditorWindow.xaml` — Add version text to the status bar

### Design Details

**Version property:** Add `<Version>0.1.0</Version>` to `Directory.Build.props`. MSBuild flows this into `AssemblyVersion`, `FileVersion`, and `InformationalVersion` automatically. All projects in the solution inherit it. Read at runtime via `Assembly.GetExecutingAssembly().GetName().Version` or `AssemblyInformationalVersionAttribute`.

**About dialog:** Simple modal WPF window matching the editor's dark theme (`#1a1a2e` / `#2a2a40`). Contents:
- Formula Boss logo (reuse embedded `logo32.png`)
- "Formula Boss" title + version number
- "by Taglo" subtitle
- "View on GitHub" link → opens browser to repo URL
- "Release Notes" link → opens browser to latest release page
- "Close" button

Use `Process.Start(new ProcessStartInfo(url) { UseShellExecute = true })` for link opening.

**Ribbon button:** Add an "About" group to the Formula Boss ribbon tab in `RibbonController.GetCustomUI()`:
```xml
<group id='aboutGroup' label='About'>
  <button id='aboutButton'
          label='About'
          imageMso='Info'
          size='normal'
          onAction='OnAbout' />
</group>
```

Callback `OnAbout` creates and shows the dialog. Since the ribbon callback runs on the Excel thread but WPF dialogs need a dispatcher, invoke on the editor's WPF thread via `ShowFloatingEditorCommand`.

**Editor version display:** Add a small version label to the right side of the existing status bar in `FloatingEditorWindow.xaml`, between the branding text and the action buttons. Style it as dim grey text (`#888888`) so it's visible but unobtrusive: `v0.1.0`.

### Order of Operations

1. Add `<Version>` to `Directory.Build.props`
2. Create `AboutDialog.xaml` and `.xaml.cs`
3. Add About button to `RibbonController.cs`
4. Add version label to `FloatingEditorWindow.xaml`
5. Build and manually verify in Excel

### Testing

- Unit test: read version from assembly metadata, assert it matches expected format
- Manual: open Excel, verify About dialog shows correct version, links open browser, version appears in editor status bar

---

## PR 2: Update Check

### Files to Touch

- `formula-boss/Updates/UpdateChecker.cs` (new) — Async version check against GitHub Releases API
- `formula-boss/AddIn.cs` — Fire update check from `AutoOpen`
- `formula-boss/UI/RibbonController.cs` — Add update notification label/button to ribbon, wire up `getVisible`/`getLabel` callbacks to show when update is available

### Design Details

**UpdateChecker class:**
```
static class UpdateChecker
{
    static string? NewVersionAvailable { get; }
    static string? ReleaseUrl { get; }
    static event Action? UpdateAvailable;

    static async void CheckForUpdateAsync() { ... }
}
```

- Uses `HttpClient` with a 10-second timeout to `GET https://api.github.com/repos/TagloGit/taglo-formula-boss/releases/latest`
- Sets `User-Agent: FormulaBoss/{version}` header (GitHub API requires a User-Agent)
- Parses `tag_name` (e.g., `v0.2.0`) and `html_url` from the JSON response using `System.Text.Json`
- Compares against current assembly version — if remote is newer, sets properties and raises `UpdateAvailable`
- Entire method is wrapped in try/catch that swallows all exceptions and logs via `Logger.Info` — network failures are silent to the user

**Integration in AddIn.AutoOpen:** After existing initialization, call `UpdateChecker.CheckForUpdateAsync()` (fire-and-forget, no await). This runs entirely on a background thread.

**Ribbon notification:** Add a dynamic button to the ribbon XML:
```xml
<button id='updateNotification'
        getLabel='GetUpdateLabel'
        getVisible='GetUpdateVisible'
        imageMso='Refresh'
        size='normal'
        onAction='OnUpdateClick' />
```

When `UpdateChecker.UpdateAvailable` fires, call `IRibbonUI.InvalidateControl("updateNotification")` to make the button appear. The `getVisible` callback returns `UpdateChecker.NewVersionAvailable != null`. The `getLabel` callback returns `"Update: v{version}"`. `OnUpdateClick` opens `UpdateChecker.ReleaseUrl` in the browser.

To get the `IRibbonUI` reference, override `OnConnection` in `RibbonController` and capture the `ribbonUI` parameter from `GetCustomUI`'s `onLoad` callback.

### Order of Operations

1. Create `UpdateChecker.cs` with the async check logic
2. Add ribbon notification elements to `RibbonController.cs` (with `onLoad` callback to capture `IRibbonUI`)
3. Wire up `UpdateChecker.CheckForUpdateAsync()` in `AddIn.AutoOpen`
4. Build and test

### Testing

- Unit test: `UpdateChecker` version comparison logic (parse tag, compare versions) — mock the HTTP response or test the parsing separately
- Manual: build, create a GitHub release with a higher version tag, launch Excel, verify the ribbon notification appears and opens the correct URL
- Manual: verify no visible effect when GitHub is unreachable (disconnect network)

---

## PR 3: InnoSetup Installer

### Files to Touch

- `installer/formula-boss.iss` (new) — InnoSetup script
- `installer/dotnet-desktop-runtime.exe` (downloaded, gitignored) — Bundled .NET 6 runtime
- `.gitignore` — Add `installer/dotnet-desktop-runtime.exe` and installer output directory
- `docs/release-signing.md` — Update to document installer signing steps
- `docs/building-installer.md` (new) — Instructions for building the installer

### Design Details

**InnoSetup script (`formula-boss.iss`):**

```
[Setup]
AppName=Formula Boss
AppVersion={version}
AppPublisher=Taglo
AppPublisherURL=https://github.com/TagloGit/taglo-formula-boss
DefaultDirName={localappdata}\FormulaBoss
PrivilegesRequired=lowest
OutputBaseFilename=formula-boss-{version}-x64-setup
SetupIconFile=..\formula-boss\Resources\logo.ico
WizardImageFile=wizard-banner.bmp
DisableProgramGroupPage=yes
Compression=lzma2
SolidCompression=yes
```

**Prerequisite check:** InnoSetup's `[Code]` section with a Pascal Script function that checks for the .NET 6 Desktop Runtime:
- Check registry key `HKLM\SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App` for a `6.x` entry
- If missing, run the bundled `dotnet-desktop-runtime.exe /install /quiet /norestart` before proceeding
- Show a progress message: "Installing .NET 6 Desktop Runtime..."

**Files section:** Reference the Release build output:
```
[Files]
Source: "..\formula-boss\bin\Release\net6.0-windows\formula-boss64.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\formula-boss.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\formula-boss64.dna"; DestDir: "{app}"; DestName: "formula-boss64.dna"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\FormulaBoss.Runtime.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\ExcelDna.Integration.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\ICSharpCode.AvalonEdit.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\Microsoft.CodeAnalysis*.dll"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
Source: "..\formula-boss\bin\Release\net6.0-windows\System.*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\Humanizer.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\Microsoft.DiaSymReader.dll"; DestDir: "{app}"; Flags: ignoreversion
; Roslyn satellite assemblies for localized error messages
Source: "..\formula-boss\bin\Release\net6.0-windows\cs\*"; DestDir: "{app}\cs"; Flags: ignoreversion
Source: "..\formula-boss\bin\Release\net6.0-windows\de\*"; DestDir: "{app}\de"; Flags: ignoreversion
; ... (remaining locales)
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
```

**Registry — Add-in Manager entry:**
```
[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Add-in Manager"; ValueType: string; ValueName: "{app}\formula-boss64.xll"; ValueData: ""; Flags: uninsdeletevalue
```

This registers the XLL so it appears in Developer > Excel Add-ins with a checkbox. The user can toggle it on/off. The `uninsdeletevalue` flag removes the entry on uninstall.

**CloseApplications:** InnoSetup's `CloseApplications=yes` will detect running Excel and prompt the user.

**Logo/branding:** Need to create:
- `installer/logo.ico` — convert the existing `logo32.png` to `.ico` format for the installer and uninstaller icons
- `installer/wizard-banner.bmp` — 164x314 wizard sidebar image with Formula Boss branding (or skip for v0.1.0 and use the default)

**Signing:** After InnoSetup compiles the installer, sign it with the Sectigo certificate using `signtool.exe`, same process as the XLL signing documented in `docs/release-signing.md`.

### Order of Operations

1. Create `installer/` directory
2. Download .NET 6 Desktop Runtime installer, add to `.gitignore`
3. Create the `.iss` script
4. Create/convert logo to `.ico` format
5. Build in Release mode, run InnoSetup compiler (`ISCC.exe formula-boss.iss`)
6. Test installer on a clean machine or VM

### Testing

- Install on a clean Windows machine (or VM) with no .NET SDK, no dev tools
- Verify .NET runtime gets installed silently
- Verify Excel opens with Formula Boss loaded
- Verify the add-in appears in Developer > Excel Add-ins with a checkbox
- Verify unchecking the box disables the add-in, rechecking re-enables
- Verify running the installer again upgrades cleanly
- Verify uninstall removes files and registry entry

---

## PR 4: Release v0.1.0

This is the final step after PRs 1-3 are merged. It encompasses the existing issue #133 scope.

### Steps

1. Update version in `Directory.Build.props` to `0.1.0` (if not already set)
2. Build in Release configuration
3. Sign `formula-boss64.xll` with Sectigo certificate
4. Run InnoSetup to produce the installer
5. Sign the installer `.exe` with Sectigo certificate
6. Tag the commit as `v0.1.0`
7. Create GitHub Release with:
   - Signed installer as the release asset
   - Release notes covering key capabilities
   - Note about 64-bit Excel requirement
8. Verify the update check works by launching an older build and confirming the notification appears

### Testing

- Fresh machine install from the GitHub Release download
- End-to-end: install → open Excel → Formula Boss loads → write a formula → it works
- Update notification: install v0.1.0, create a test v0.1.1 release, verify notification appears

---

## Testing Strategy

Testing is split into four stages:

### Stage 1: Unit and manual testing (during PRs 1-3)

- **Version infrastructure:** Unit test that reads version from assembly metadata. Manual check of About dialog and editor version label.
- **Update checker:** Unit test the version comparison/parsing logic with mock responses. Manual test with network disconnected to verify silent failure.
- **Installer:** Build installer locally, run on dev machine, verify files, registry, and Excel loading.

### Stage 2: Upgrade and clean-machine testing (after PRs 1-3 merged)

- **Upgrade path:** Build installer at v0.0.9 and v0.1.0. Install v0.0.9, then run v0.1.0 installer over it. Verify clean upgrade.
- **Clean-machine test:** Run the installer on Tim's VM (Office only, no .NET SDK). Validates runtime bundling and installer end-to-end.

### Stage 3: Update check end-to-end (pre-release)

- Publish a GitHub pre-release tagged `v0.0.1-test`. Install that build.
- Publish `v0.0.2-test` (or briefly publish a non-prerelease, since `/releases/latest` ignores pre-releases). Verify ribbon notification appears and links to the release page.
- Clean up test releases afterward.

### Stage 4: Beta test (before public announcement)

- Publish the real v0.1.0 release on GitHub but don't promote it yet.
- Share the installer link with 2-3 trusted testers on different machines.
- Validate: SmartScreen behaviour with signed installer, antivirus compatibility, .NET runtime install on varied configurations.
- Once confirmed working, announce broadly.

## Resolved Questions

1. **Icon file:** No existing `.ico`. Generate from the source HTML in `resources/` at the repo root (better quality than converting the 32px PNG). Include multiple sizes (16, 32, 48, 256) in the `.ico`.
2. **Wizard banner image:** Create a branded sidebar image with Formula Boss logo and branding.
3. **Office version registry path:** `Office\16.0` is sufficient (covers 2016, 2019, 2021, 365).
4. **InnoSetup installation:** Not currently installed. Add setup instructions to `docs/building-installer.md`. InnoSetup is free and can be installed via `winget install JRSoftware.InnoSetup` or downloaded from jrsoftware.org.
