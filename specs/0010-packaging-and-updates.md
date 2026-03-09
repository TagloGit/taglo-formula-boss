# 0010 — Packaging and Updates

## Problem

Formula Boss currently has no distribution mechanism beyond copying build output files manually. The `.xll` cannot work standalone — it requires companion DLLs (`formula-boss.dll`, `ExcelDna.Integration.dll`, `FormulaBoss.Runtime.dll`) to remain unpacked on disk for Roslyn compilation and ALC loading. This makes "just ship the XLL" impossible.

For a potentially large audience of Excel experts (not necessarily developers), we need:
- A frictionless install experience (single `.exe`, no manual file management)
- Visibility into what version is running (for bug reports and support)
- A smooth update path for frequent early releases

## Proposed Solution

Three components: an InnoSetup installer, version infrastructure with an About dialog, and an in-app update check against GitHub Releases.

### 1. InnoSetup Installer

**Why InnoSetup:** Free, mature, produces a single `.exe`, handles upgrades natively (reinstall over previous version), supports custom branding, and is well-understood. It also produces artifacts compatible with WinGet for future distribution.

**Install behaviour:**
- Single `formula-boss-v{version}-x64-setup.exe` published as a GitHub Release asset
- Installs to `%LocalAppData%\FormulaBoss\` (no admin rights required)
- Registers the XLL via Excel's Add-in Manager registry key (`HKCU\Software\Microsoft\Office\16.0\Excel\Add-in Manager`) so it auto-loads but users can disable/enable from Developer > Excel Add-ins
- Shows the Formula Boss logo on the installer wizard
- No license agreement screen (MIT doesn't require click-through acceptance; `LICENSE` file is included in the install folder)
- Code-signed with the Sectigo certificate to avoid SmartScreen warnings

**Prerequisite handling:**
- Bundles the .NET 6 Desktop Runtime installer (~50MB) as a payload
- On install, checks whether the runtime is already present; if not, runs the bundled installer silently before proceeding
- This ensures non-technical users never hit a cryptic runtime error

**Upgrade behaviour:**
- Running a newer installer over an existing installation replaces files cleanly
- InnoSetup's `AppId` ensures the old version is recognized and upgraded in place
- If Excel is running, the installer prompts the user to close it before proceeding

**Uninstall:**
- Appears in Windows Add/Remove Programs
- Removes installed files and the Add-in Manager registry key
- Does not remove user data (if any exists in future)

**Files installed:**
- `formula-boss64.xll` (signed)
- `formula-boss.dll`
- `formula-boss.dna`
- `ExcelDna.Integration.dll`
- `FormulaBoss.Runtime.dll`
- Roslyn compiler assemblies
- AvalonEdit and other managed dependencies
- `LICENSE`

### 2. Version Infrastructure

**Single source of truth:** Version number defined in `Directory.Build.props` as `<Version>0.1.0</Version>`. This flows into `AssemblyVersion`, `FileVersion`, and `InformationalVersion` for all projects automatically.

**About dialog:**
- Accessible from the Formula Boss ribbon tab (e.g., an "About" or "Info" button)
- Displays: Formula Boss logo, version number, "by Taglo", link to GitHub repo, link to release notes
- Simple modal dialog — no need for a full settings/preferences UI yet

**Version in editor:**
- Display version in the floating editor's status bar (bottom of the editor window, alongside any existing status information)

### 3. Update Check

**Mechanism:**
- On Excel startup (in `AutoOpen`), fire an async background HTTP request to `https://api.github.com/repos/TagloGit/taglo-formula-boss/releases/latest`
- Parse the `tag_name` field, compare against the embedded assembly version
- If a newer version is available, show a notification in the ribbon (e.g., a label or button that says "Update available: v0.2.0")
- Clicking the notification opens the browser to the GitHub Release page where the user can download the new installer

**Performance:** The API call is a single HTTPS GET returning ~2KB of JSON, executed on a background thread. Zero impact on Excel startup time. If the network is unavailable or the request fails, nothing happens — no error, no retry, no user-visible effect.

**Frequency:** Every Excel launch. No throttling or opt-out for v0.1.0.

**No auto-download:** The update check only notifies. The user downloads and runs the new installer themselves. This keeps the implementation simple and avoids SmartScreen issues with programmatically downloaded executables.

## User Stories

- As an Excel user, I want to install Formula Boss by double-clicking a setup file, so that I don't need to understand file management or registry editing.
- As an Excel user, I want Formula Boss to load automatically when I open Excel, so that I don't have to enable it manually each time.
- As an Excel user, I want to disable Formula Boss from the standard Excel Add-ins dialog, so that I have control without needing to uninstall.
- As a user reporting a bug, I want to see what version I'm running, so that I can include it in my report.
- As a user, I want to be notified when an update is available, so that I get bug fixes and new features without having to check manually.
- As a user, I want the installer to handle .NET runtime prerequisites, so that I never see a cryptic framework error.

## Acceptance Criteria

- [ ] InnoSetup script produces a working installer from build output
- [ ] Installer is code-signed with Sectigo certificate
- [ ] Installer bundles and conditionally installs .NET 6 Desktop Runtime
- [ ] Installation creates Add-in Manager registry entry (user can toggle in Excel Add-ins dialog)
- [ ] Upgrade install over previous version works cleanly
- [ ] Uninstaller removes files and registry entry
- [ ] Version number defined in `Directory.Build.props`, visible in assembly metadata
- [ ] About dialog accessible from ribbon, shows version and links
- [ ] Version displayed in floating editor status bar
- [ ] Async update check runs on startup, compares against GitHub Releases API
- [ ] Ribbon notification shown when a newer version is available
- [ ] Clicking update notification opens the release page in the browser
- [ ] Update check failure (network down, API error) is silent — no user impact
- [ ] Fresh machine test: installer works on a machine with no dev tools or .NET runtime

## Relationship to Spec 0007

This spec supersedes the Distribution section of spec 0007 (MVP Launch). Spec 0007's acceptance criterion "Signed 64-bit .xll published as GitHub Release v0.1.0" is replaced by "Signed installer published as GitHub Release v0.1.0." Issue #133 should be updated to reflect this.

## Out of Scope

- Auto-download and silent update (future enhancement)
- WinGet / Chocolatey publishing (straightforward follow-up once we have an installer)
- 32-bit installer variant
- Settings/preferences UI beyond the About dialog
- Telemetry or crash reporting
- Update check opt-out or frequency configuration

## Open Questions

None.
