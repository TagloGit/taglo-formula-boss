<#
.SYNOPSIS
    End-to-end release publishing script for Formula Boss.

.DESCRIPTION
    Walks through every step of building, signing, and publishing a Formula Boss release.
    Must be run from the main branch with a clean working directory.

    Steps:
      1. Prompt for new version and update Directory.Build.props
      2. Commit version bump and push to main
      3. Build in Release configuration
      4. Run tests
      5. Sign the XLL with code-signing certificate
      6. Sync version to InnoSetup script
      7. Compile the installer
      8. Sign the installer
      9. Create git tag
     10. Create GitHub Release (draft) with signed installer

    Prompts for version, certificate path, and password. No secrets are stored.

.PARAMETER Version
    The version to release (e.g. "0.2.0"). Prompted if not provided.

.PARAMETER CertPath
    Path to the .pfx code-signing certificate. Prompted if not provided.

.PARAMETER CertPassword
    Password for the .pfx certificate. Prompted securely if not provided.

.PARAMETER SkipTests
    Skip running tests (use when you've already verified).

.PARAMETER SkipSign
    Skip code signing (for local testing only).

.PARAMETER DryRun
    Show what would happen without executing destructive steps (commit, tag, release).

.EXAMPLE
    .\scripts\publish-release.ps1
    .\scripts\publish-release.ps1 -Version "0.2.0" -CertPath "C:\certs\sectigo.pfx"
    .\scripts\publish-release.ps1 -DryRun -SkipSign
#>

[CmdletBinding()]
param(
    [string]$Version,
    [string]$CertPath,
    [string]$CertPassword,
    [switch]$SkipTests,
    [switch]$SkipSign,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path "$PSScriptRoot\..").Path
$BuildOutput = "$RepoRoot\formula-boss\bin\Release\net6.0-windows"
$InstallerDir = "$RepoRoot\installer"
$SignTool = "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe"
$IsccExe = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
$TimestampServer = "http://timestamp.sectigo.com"

# --- Helpers ---

function Write-Step([int]$StepNumber, [string]$StepDescription) {
    Write-Host ""
    Write-Host "=== Step ${StepNumber}: ${StepDescription} ===" -ForegroundColor Cyan
}

function Write-Success([string]$Message) {
    Write-Host "  OK: $Message" -ForegroundColor Green
}

function Write-Skip([string]$Message) {
    Write-Host "  SKIP: $Message" -ForegroundColor Yellow
}

function Confirm-Continue([string]$Prompt) {
    $response = Read-Host "$Prompt (y/n)"
    if ($response -notin @("y", "Y", "yes")) {
        Write-Host "Aborted." -ForegroundColor Red
        exit 1
    }
}

# --- Step 0: Preflight checks ---

Write-Host ""
Write-Host "Formula Boss Release Publisher" -ForegroundColor Magenta
Write-Host "==============================" -ForegroundColor Magenta

# Check tools
if (-not (Test-Path $IsccExe)) {
    Write-Error "InnoSetup not found at: $IsccExe`nInstall from https://jrsoftware.org/isdl.php"
}
if (-not $SkipSign -and -not (Test-Path $SignTool)) {
    Write-Error "signtool.exe not found at: $SignTool`nInstall the ClickOnce Publishing Tools via Visual Studio Installer."
}
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Error "GitHub CLI (gh) not found. Install from https://cli.github.com/"
}
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Error ".NET SDK not found on PATH."
}

# Must be on main
$currentBranch = git -C $RepoRoot rev-parse --abbrev-ref HEAD
if ($currentBranch -ne "main") {
    Write-Error "Must be on the main branch to publish a release. Current branch: $currentBranch"
}

# Working directory must be clean
$gitStatus = git -C $RepoRoot status --porcelain
if ($gitStatus) {
    Write-Host ""
    Write-Host "Uncommitted changes:" -ForegroundColor Red
    Write-Host $gitStatus
    Write-Error "Working directory must be clean before publishing. Commit or stash your changes."
}

# Pull latest
git -C $RepoRoot pull --ff-only
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to pull latest main. Resolve any divergence first." }

# --- Step 1: Version bump ---

Write-Step 1 "Version bump"

$propsFile = "$RepoRoot\Directory.Build.props"
[xml]$props = Get-Content $propsFile
$currentVersion = $props.Project.PropertyGroup.Version
if (-not $currentVersion) {
    Write-Error "Could not read <Version> from $propsFile"
}

Write-Host "  Current version: $currentVersion"

if (-not $Version) {
    $Version = Read-Host "New version (e.g. 0.2.0)"
}

if (-not ($Version -match '^\d+\.\d+\.\d+$')) {
    Write-Error "Invalid version format: $Version (expected X.Y.Z)"
}

if ($Version -eq $currentVersion) {
    Write-Host "  Version unchanged: $Version" -ForegroundColor Yellow
    Confirm-Continue "Release with the same version?"
}

$version = $Version
$xllPath = "$BuildOutput\formula-boss64.xll"
$installerName = "FormulaBoss-$version-Setup.exe"
$installerPath = "$InstallerDir\output\$installerName"
$tag = "v$version"

# Check if tag already exists
$existingTag = git -C $RepoRoot tag -l $tag
if ($existingTag) {
    Write-Host ""
    Write-Host "WARNING: Tag $tag already exists!" -ForegroundColor Yellow
    Confirm-Continue "This will overwrite the existing tag. Continue?"
}

# Update Directory.Build.props
$propsContent = Get-Content $propsFile -Raw
$propsContent = $propsContent -replace '<Version>.*?</Version>', "<Version>$version</Version>"
Set-Content $propsFile $propsContent -NoNewline

Write-Success "Updated Directory.Build.props: $currentVersion -> $version"
Write-Host "  Installer: $installerName"
Write-Host "  Tag:       $tag"

# Commit and push version bump
if (-not $DryRun) {
    git -C $RepoRoot add $propsFile
    git -C $RepoRoot commit -m "Bump version to $version"
    if ($LASTEXITCODE -ne 0) { Write-Error "Failed to commit version bump." }
    git -C $RepoRoot push
    if ($LASTEXITCODE -ne 0) { Write-Error "Failed to push version bump." }
    Write-Success "Version bump committed and pushed"
} else {
    Write-Skip "Version bump commit (--DryRun)"
}

# --- Step 2: Collect signing credentials ---

if (-not $SkipSign) {
    Write-Step 2 "Collect signing credentials"

    if (-not $CertPath) {
        $CertPath = Read-Host "Path to .pfx certificate"
    }
    if (-not (Test-Path $CertPath)) {
        Write-Error "Certificate not found: $CertPath"
    }
    Write-Success "Certificate: $CertPath"

    if (-not $CertPassword) {
        $securePwd = Read-Host "Certificate password" -AsSecureString
        $CertPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd)
        )
    }
} else {
    Write-Skip "Code signing (--SkipSign)"
}

# --- Step 3: Build ---

Write-Step 3 "Build in Release configuration"

dotnet build "$RepoRoot\formula-boss\formula-boss.slnx" -c Release
if ($LASTEXITCODE -ne 0) { Write-Error "Build failed." }
Write-Success "Build succeeded"

if (-not (Test-Path $xllPath)) {
    Write-Error "Expected XLL not found at: $xllPath"
}

# --- Step 4: Tests ---

if (-not $SkipTests) {
    Write-Step 4 "Run tests"
    dotnet test "$RepoRoot\formula-boss\formula-boss.slnx" -c Release --no-build
    if ($LASTEXITCODE -ne 0) { Write-Error "Tests failed." }
    Write-Success "All tests passed"
} else {
    Write-Step 4 "Run tests"
    Write-Skip "Tests (--SkipTests)"
}

# --- Step 5: Sign the XLL ---

if (-not $SkipSign) {
    Write-Step 5 "Sign the XLL"

    & $SignTool sign /f $CertPath /p $CertPassword /fd sha256 /tr $TimestampServer /td sha256 $xllPath
    if ($LASTEXITCODE -ne 0) { Write-Error "XLL signing failed." }

    & $SignTool verify /pa $xllPath
    if ($LASTEXITCODE -ne 0) { Write-Error "XLL signature verification failed." }
    Write-Success "XLL signed and verified"
} else {
    Write-Step 5 "Sign the XLL"
    Write-Skip "XLL signing (--SkipSign)"
}

# --- Step 6: Sync version to InnoSetup script ---

Write-Step 6 "Sync version to InnoSetup script"

$issFile = "$InstallerDir\formula-boss.iss"
$issContent = Get-Content $issFile -Raw
$issVersionPattern = '#define MyAppVersion ".*?"'
$issVersionNew = "#define MyAppVersion ""$version"""

if ($issContent -match [regex]::Escape($issVersionNew)) {
    Write-Success "InnoSetup version already matches: $version"
} else {
    $issContent = $issContent -replace $issVersionPattern, $issVersionNew
    Set-Content $issFile $issContent -NoNewline
    Write-Success "Updated InnoSetup version to: $version"
}

# --- Step 7: Compile the installer ---

Write-Step 7 "Compile the installer"

# Check bundled runtime exists
$runtimeFiles = Get-ChildItem "$InstallerDir\bundled-runtime\windowsdesktop-runtime-6.0.*-win-x64.exe" -ErrorAction SilentlyContinue
if (-not $runtimeFiles) {
    Write-Host ""
    Write-Host "WARNING: No bundled .NET runtime found in installer/bundled-runtime/" -ForegroundColor Yellow
    Write-Host "Download from: https://dotnet.microsoft.com/en-us/download/dotnet/6.0"
    Confirm-Continue "Continue without bundled runtime?"
}

& $IsccExe $issFile
if ($LASTEXITCODE -ne 0) { Write-Error "InnoSetup compilation failed." }

if (-not (Test-Path $installerPath)) {
    Write-Error "Expected installer not found at: $installerPath"
}
Write-Success "Installer built: $installerPath"

# --- Step 8: Sign the installer ---

if (-not $SkipSign) {
    Write-Step 8 "Sign the installer"

    & $SignTool sign /f $CertPath /p $CertPassword /fd sha256 /tr $TimestampServer /td sha256 $installerPath
    if ($LASTEXITCODE -ne 0) { Write-Error "Installer signing failed." }

    & $SignTool verify /pa $installerPath
    if ($LASTEXITCODE -ne 0) { Write-Error "Installer signature verification failed." }
    Write-Success "Installer signed and verified"
} else {
    Write-Step 8 "Sign the installer"
    Write-Skip "Installer signing (--SkipSign)"
}

# --- Step 9: Summary before publish ---

Write-Host ""
Write-Host "==============================" -ForegroundColor Magenta
Write-Host "Ready to publish!" -ForegroundColor Magenta
Write-Host "==============================" -ForegroundColor Magenta
Write-Host ""
Write-Host "  Version:   $version"
Write-Host "  Tag:       $tag"
Write-Host "  Installer: $installerPath"
$installerSize = (Get-Item $installerPath).Length / 1MB
Write-Host ("  Size:      {0:N1} MB" -f $installerSize)
Write-Host ""

if ($DryRun) {
    Write-Host "DRY RUN -- skipping tag and GitHub Release creation." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Would run:" -ForegroundColor Yellow
    Write-Host "  git tag $tag"
    Write-Host "  git push origin $tag"
    Write-Host "  gh release create $tag --title 'Formula Boss $version' ..."
    exit 0
}

Confirm-Continue "Create git tag and GitHub Release?"

# --- Step 10: Tag and release ---

Write-Step 9 "Create git tag"

# Delete existing tag if present (user already confirmed above)
$existingTag = git -C $RepoRoot tag -l $tag
if ($existingTag) {
    git -C $RepoRoot tag -d $tag
}

git -C $RepoRoot tag $tag
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to create tag." }
Write-Success "Tag $tag created locally"

git -C $RepoRoot push origin $tag
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to push tag." }
Write-Success "Tag pushed to origin"

Write-Step 10 "Create GitHub Release"

# Write release notes to a temp file to avoid here-string parsing issues
$releaseNotesFile = [System.IO.Path]::GetTempFileName()
try {
    $notes = @(
        "## Formula Boss $version",
        "",
        "### What's New",
        "<!-- TODO: Replace these placeholder bullets with actual release highlights -->",
        "- <TODO: describe key change 1>",
        "- <TODO: describe key change 2>",
        "- <TODO: describe key change 3>",
        "",
        "### Requirements",
        "- 64-bit Excel (Microsoft 365 or Excel 2019+)",
        "- Windows 10/11",
        "- .NET 6 Desktop Runtime (bundled in installer)",
        "",
        "### Installation",
        "Download and run ``$installerName``. The installer will:",
        "- Install to ``%LOCALAPPDATA%\FormulaBoss``",
        "- Install .NET 6 Desktop Runtime if needed",
        "- Register the add-in with Excel automatically"
    )
    $notes -join "`n" | Set-Content $releaseNotesFile -NoNewline

    gh release create $tag `
        --repo TagloGit/taglo-formula-boss `
        --title "Formula Boss $version" `
        --notes-file $releaseNotesFile `
        --draft `
        $installerPath
} finally {
    Remove-Item $releaseNotesFile -ErrorAction SilentlyContinue
}

if ($LASTEXITCODE -ne 0) { Write-Error "Failed to create GitHub Release." }

Write-Host ""
Write-Host "==============================" -ForegroundColor Green
Write-Host "Release created as DRAFT" -ForegroundColor Green
Write-Host "==============================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Edit the release notes on GitHub"
Write-Host "  2. Test the installer from the release download"
Write-Host "  3. Publish the release when ready"
Write-Host ""
