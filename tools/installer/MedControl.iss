; Inno Setup script for MedControl
; Build requirements: Inno Setup 6 (ISCC)
; To build: Right-click this file and choose "Compile" (or run ISCC from command line)

#define MyAppName "MedControl"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "DayanFA"
#define MyAppURL "https://github.com/DayanFA/MedControl"
#define MyAppExeName "MedControl.exe"
#define PublishDir "C:\Users\dolby\OneDrive\Área de Trabalho\medalgo\tools\installer\dist"
#define AppIcon "C:\Users\dolby\OneDrive\Área de Trabalho\medalgo\MedControl\Assets\app.ico"

[Setup]
AppId={{A2E5F2A0-1F5B-4C73-9F5E-3B5E0E8E8A77}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir={#SourcePath}\\dist
OutputBaseFilename={#MyAppName}-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
SetupIconFile="{#AppIcon}"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "portuguesebr"; MessagesFile: "compiler:Languages\\BrazilianPortuguese.isl"

[Files]
; include everything from publish folder (single-file exe + any sidecar files if present)
Source: "{#PublishDir}\\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{autoprograms}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"
Name: "{autodesktop}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Run]
Filename: "{app}\\{#MyAppExeName}"; Description: "Executar {#MyAppName}"; Flags: nowait postinstall skipifsilent
