; Inno Setup Script for Donut Data Cleaner
; Download Inno Setup free from: https://jrsoftware.org/isinfo.php
; Then open this file with Inno Setup Compiler and click "Compile".
; Before compiling, run build_windows.bat first so dist\DonutDataCleaner exists.

#define MyAppName "Donut Data Cleaner"
#define MyAppVersion "1.0"
#define MyAppPublisher "Adrian Mohammed"
#define MyAppExeName "DonutDataCleaner.exe"

[Setup]
AppId={{B3C1A2E4-9F21-4C7A-8E3D-DONUTCLEANER1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=DonutDataCleaner_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\DonutDataCleaner\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
