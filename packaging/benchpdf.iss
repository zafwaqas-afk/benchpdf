; Inno Setup script for BenchPDF — a per-user installer (no admin rights).
; Installs the PyInstaller onedir build, adds Start Menu + optional Desktop
; shortcuts, and removes everything (including %LOCALAPPDATA%\BenchPDF) on
; uninstall so nothing is left behind.

#define AppName     "BenchPDF"
#define AppVersion  "1.0.1"
#define AppPublisher "BenchPDF"
#define AppExe      "BenchPDF.exe"

[Setup]
AppId={{B3F1C2A6-9D4E-4E37-9E2A-BENCHPDF0001}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; ---- per-user install: no admin prompt ----
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist_installer
OutputBaseFilename=BenchPDF-Setup-{#AppVersion}
SetupIconFile=benchpdf.ico
UninstallDisplayIcon={app}\{#AppExe}
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
LicenseFile=..\LICENSE
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &Desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
; the entire PyInstaller onedir output
Source: "..\dist\BenchPDF\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\{#AppName} diagnostics"; Filename: "{app}\{#AppExe}"; Parameters: "--diagnostics"; Comment: "Run the built-in fidelity regression suite"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; remove the per-user working area + logs so uninstall leaves nothing behind
Type: filesandordirs; Name: "{localappdata}\{#AppName}"
