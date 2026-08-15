#define MyAppName "PrintDesk"
#ifndef AppVersion
  #define AppVersion "1.2.1"
#endif
#define MyAppPublisher "PrintDesk"
#define MyAppExeName "PrintCalc.exe"

[Setup]
AppId={{8E6C3C11-1A11-4A9A-B40C-PRINTDESK1200}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\PrintDesk
DefaultGroupName=PrintDesk
DisableProgramGroupPage=yes
OutputDir=..\installer-output
OutputBaseFilename=PrintDesk-Setup-v{#AppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}

[Files]
Source: "..\bin\Release\net8.0-windows\win-x64\publish\PrintCalc.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\PrintDesk"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\PrintDesk"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch PrintDesk"; Flags: nowait postinstall skipifsilent
