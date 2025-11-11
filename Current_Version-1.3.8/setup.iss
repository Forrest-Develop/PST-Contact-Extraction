; -------------------------------------------
; setup.iss — PST Contacts Extractor Installer
; -------------------------------------------
#define MyAppName        "PST Contacts Extractor"
#define MyAppVersion     "1.3.8"
#define MyAppPublisher   "ForrestDev"
#define MyAppExeName     "PSTContactExtractor_1.3.8.exe"

[Setup]
; Generate a GUID once and keep it stable for future versions
AppId={{F6E0F67F-4A5A-4C9C-9C6F-2C9B1D5E7B1A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
LicenseFile=License.txt
SetupIconFile=app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
OutputDir=.
OutputBaseFilename=PST_Contacts_Extractor_{#MyAppVersion}_Setup
DisableProgramGroupPage=yes
CloseApplications=yes
; If you have any signing tool (self-signed or CA), set this to enable auto-signing at build
;SignTool=mysigntool

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "PSTContactExtractor_1.3.8.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "License.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\app.ico"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"
Name: "openreadme"; Description: "Open README after install"; Flags: unchecked

[Run]
Filename: "{app}\README.txt"; Description: "View README"; Flags: postinstall shellexec skipifsilent; Tasks: openreadme
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

; ---- Code signing integration (optional but recommended) ----
; Inno substitutes $f with the file it is signing (the built installer).
; If you don't have a PFX yet, you can still compile; signing will just be skipped/failed.
[SignTool]
mysigntool=sign /fd sha256 /tr http://timestamp.digicert.com /td sha256 /f "codesign.pfx" /p "PFX_PASSWORD" $f
