; Formula Boss - InnoSetup Installer Script
; Installs to {localappdata}\FormulaBoss (no admin rights required)

#define MyAppName "Formula Boss"
#define MyAppVersion "0.1.2"
#define MyAppPublisher "Taglo"
#define MyAppURL "https://github.com/TagloGit/taglo-formula-boss"

; Path to Release build output — update if building from a different location
#define BuildOutput "..\formula-boss\bin\Release\net6.0-windows"

[Setup]
AppId={{B3F7A2E1-4D5C-4E6F-8A9B-1C2D3E4F5A6B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={localappdata}\FormulaBoss
DisableProgramGroupPage=yes
OutputBaseFilename=FormulaBoss-{#MyAppVersion}-Setup
SetupIconFile=logo.ico
WizardImageFile=wizard-banner.bmp
WizardSmallImageFile=wizard-small.bmp
Compression=lzma2/ultra
SolidCompression=yes
PrivilegesRequired=lowest
CloseApplications=yes
CloseApplicationsFilter=EXCEL.EXE
RestartApplications=no
UninstallDisplayIcon={app}\logo.ico
OutputDir=output
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; XLL shim (64-bit)
Source: "{#BuildOutput}\formula-boss64.xll"; DestDir: "{app}"; Flags: ignoreversion

; .dna configuration
Source: "{#BuildOutput}\formula-boss64.dna"; DestDir: "{app}"; Flags: ignoreversion

; Core managed assemblies (unpacked per .dna config)
Source: "{#BuildOutput}\formula-boss.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\FormulaBoss.Runtime.dll"; DestDir: "{app}"; Flags: ignoreversion

; Supporting config files
Source: "{#BuildOutput}\formula-boss.deps.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\formula-boss.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\formula-boss64.deps.json"; DestDir: "{app}"; Flags: ignoreversion

; Roslyn compiler assemblies
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.CSharp.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.CSharp.Features.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.CSharp.Workspaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.Features.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.Scripting.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.Workspaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.Elfie.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.CodeAnalysis.AnalyzerUtilities.dll"; DestDir: "{app}"; Flags: ignoreversion

; AvalonEdit (floating editor UI)
Source: "{#BuildOutput}\ICSharpCode.AvalonEdit.dll"; DestDir: "{app}"; Flags: ignoreversion

; Other dependencies
Source: "{#BuildOutput}\Humanizer.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.Bcl.AsyncInterfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.DiaSymReader.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Collections.Immutable.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Composition.AttributedModel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Composition.Convention.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Composition.Hosting.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Composition.Runtime.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Composition.TypedParts.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Configuration.ConfigurationManager.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.IO.Pipelines.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Reflection.Metadata.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Security.Cryptography.ProtectedData.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Text.Encoding.CodePages.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Text.Encodings.Web.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Text.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Threading.Channels.dll"; DestDir: "{app}"; Flags: ignoreversion

; Resources
Source: "{#BuildOutput}\Resources\*"; DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs

; Satellite assemblies (Roslyn localisation)
Source: "{#BuildOutput}\cs\*"; DestDir: "{app}\cs"; Flags: ignoreversion
Source: "{#BuildOutput}\de\*"; DestDir: "{app}\de"; Flags: ignoreversion
Source: "{#BuildOutput}\es\*"; DestDir: "{app}\es"; Flags: ignoreversion
Source: "{#BuildOutput}\fr\*"; DestDir: "{app}\fr"; Flags: ignoreversion
Source: "{#BuildOutput}\it\*"; DestDir: "{app}\it"; Flags: ignoreversion
Source: "{#BuildOutput}\ja\*"; DestDir: "{app}\ja"; Flags: ignoreversion
Source: "{#BuildOutput}\ko\*"; DestDir: "{app}\ko"; Flags: ignoreversion
Source: "{#BuildOutput}\pl\*"; DestDir: "{app}\pl"; Flags: ignoreversion
Source: "{#BuildOutput}\pt-BR\*"; DestDir: "{app}\pt-BR"; Flags: ignoreversion
Source: "{#BuildOutput}\ru\*"; DestDir: "{app}\ru"; Flags: ignoreversion
Source: "{#BuildOutput}\tr\*"; DestDir: "{app}\tr"; Flags: ignoreversion
Source: "{#BuildOutput}\zh-Hans\*"; DestDir: "{app}\zh-Hans"; Flags: ignoreversion
Source: "{#BuildOutput}\zh-Hant\*"; DestDir: "{app}\zh-Hant"; Flags: ignoreversion

; Icon for uninstall entry
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion

; LICENSE
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

; .NET 6 Desktop Runtime installer (bundled, not checked into source control)
; Download from: https://dotnet.microsoft.com/en-us/download/dotnet/6.0
Source: "bundled-runtime\windowsdesktop-runtime-6.0.*-win-x64.exe"; DestDir: "{tmp}"; Flags: ignoreversion deleteafterinstall; Check: not IsDotNet6Installed

[Registry]
; Register XLL with Excel Add-in Manager (HKCU, per-user, no admin needed)
; Uses OPEN key under Excel Options — the [Code] section finds the next free slot
; uninsdeletevalue removes it on uninstall so the add-in disappears from Excel's list
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Options"; ValueType: string; ValueName: "{code:GetOpenKeyName}"; ValueData: """{app}\formula-boss64.xll"""; Flags: uninsdeletevalue

[Run]
; Install .NET 6 Desktop Runtime silently if not present
; NOTE: Update this filename to match the runtime version in bundled-runtime/
Filename: "{tmp}\windowsdesktop-runtime-6.0.36-win-x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "Installing .NET 6 Desktop Runtime..."; Check: not IsDotNet6Installed; Flags: waituntilterminated

[Code]
var
  OpenKeyName: string;

function IsDotNet6Installed: Boolean;
var
  FindRec: TFindRec;
begin
  // Check for any .NET 6.0.x Windows Desktop Runtime by looking for version directories
  Result := FindFirst(
    ExpandConstant('{commonpf}\dotnet\shared\Microsoft.WindowsDesktop.App\6.0.*'),
    FindRec);
  if Result then
    FindClose(FindRec);
end;

function FindNextOpenKey: String;
var
  I: Integer;
  KeyName: String;
  Value: String;
  XllPath: String;
begin
  // Check if our XLL is already registered under an existing OPEN key
  XllPath := LowerCase(ExpandConstant('{app}\formula-boss64.xll'));

  // Check OPEN (no number)
  if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN', Value) then
  begin
    if Pos(LowerCase(XllPath), LowerCase(Value)) > 0 then
    begin
      Result := 'OPEN';
      Exit;
    end;
  end
  else
  begin
    Result := 'OPEN';
    Exit;
  end;

  // Check OPEN1 through OPEN99
  for I := 1 to 99 do
  begin
    KeyName := 'OPEN' + IntToStr(I);
    if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName, Value) then
    begin
      if Pos(LowerCase(XllPath), LowerCase(Value)) > 0 then
      begin
        Result := KeyName;
        Exit;
      end;
    end
    else
    begin
      Result := KeyName;
      Exit;
    end;
  end;

  // Fallback (should never reach here)
  Result := 'OPEN99';
end;

function GetOpenKeyName(Param: String): String;
begin
  if OpenKeyName = '' then
    OpenKeyName := FindNextOpenKey;
  Result := OpenKeyName;
end;

function IsExcelRunning: Boolean;
var
  WMIService: Variant;
  Processes: Variant;
begin
  try
    WMIService := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := WMIService.ConnectServer('.', 'root\cimv2');
    Processes := WMIService.ExecQuery('SELECT Name FROM Win32_Process WHERE Name = "EXCEL.EXE"');
    Result := (Processes.Count > 0);
  except
    // If WMI fails, fall back to assuming Excel is not running
    Result := False;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  while IsExcelRunning do
  begin
    if MsgBox('Excel is currently running and must be closed before Formula Boss can be installed.' + #13#10 + #13#10 +
              'Please close Excel and click OK to continue, or click Cancel to abort the installation.',
              mbError, MB_OKCANCEL) = IDCANCEL then
    begin
      Result := 'Installation cancelled. Please close Excel and try again.';
      Exit;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  I: Integer;
  KeyName: String;
  Value: String;
  XllPath: String;
begin
  // On uninstall, clean up any OPEN key that references our XLL
  // (the [Registry] uninsdeletevalue handles the one we set, but this catches
  //  cases where the user or Excel moved it to a different OPEN slot)
  if CurUninstallStep = usUninstall then
  begin
    XllPath := LowerCase(ExpandConstant('{app}\formula-boss64.xll'));

    if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN', Value) then
    begin
      if Pos(LowerCase(XllPath), LowerCase(Value)) > 0 then
        RegDeleteValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN');
    end;

    for I := 1 to 99 do
    begin
      KeyName := 'OPEN' + IntToStr(I);
      if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName, Value) then
      begin
        if Pos(LowerCase(XllPath), LowerCase(Value)) > 0 then
          RegDeleteValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName);
      end;
    end;
  end;
end;
