; windows.iss — Inno Setup script for CBMI Loop (Windows x64 beta installer).
; Wraps the PyInstaller onedir output (dist\CBMI-Loop\) into a single Setup.exe.
; Unsigned for beta: users click through the SmartScreen "More info -> Run anyway"
; prompt (see TESTER_GUIDE.md). Build with: iscc installer\windows.iss
; (expects TesterId + AppVersion passed via /D or defaulted below).

#ifndef TesterId
  #define TesterId "beta"
#endif
#ifndef AppVersion
  #define AppVersion "0.1.0-beta"
#endif

[Setup]
AppId={{A9E4C1B2-CB11-4E7A-9C3F-CBMILOOP0001}
AppName=CBMI Loop
AppVersion={#AppVersion}
AppPublisher=Crystalball AI
DefaultDirName={autopf}\CBMI Loop
DefaultGroupName=CBMI Loop
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=CBMI-Loop-{#TesterId}-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern
UninstallDisplayIcon={app}\CBMI-Loop.exe

[Files]
; the entire PyInstaller onedir tree (built by CI on a Windows runner)
Source: "dist\CBMI-Loop\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\CBMI Loop"; Filename: "{app}\CBMI-Loop.exe"
Name: "{autodesktop}\CBMI Loop"; Filename: "{app}\CBMI-Loop.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\CBMI-Loop.exe"; Description: "Launch CBMI Loop"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; leave the user's app-data (results/history) intact on uninstall
Type: filesandordirs; Name: "{app}\_internal\__pycache__"
