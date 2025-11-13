; Inno Setup Script for MedControl
; Requisitos: Instalar Inno Setup (https://jrsoftware.org/isinfo.php)
; Para compilar: Abra este arquivo no Inno Setup Compiler (ISCC) e clique em Build.

[Setup]
AppId={{A4E6E7C0-53C6-4C4F-A3A5-1C8B2C9E8A11}
AppName=MedControl
AppVersion=1.2.2
AppPublisher=MedControl
; Ícone do instalador (usa o mesmo ícone do aplicativo)
SetupIconFile=..\..\MedControl\Assets\app.ico
DefaultDirName={autopf}\\MedControl
DefaultGroupName=MedControl
UninstallDisplayIcon={app}\\MedControl.exe
Compression=lzma2
SolidCompression=yes
; Salvar saída dentro da pasta do projeto para facilitar localização
OutputDir=.\\dist
OutputBaseFilename=MedControl-Setup
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
DisableDirPage=no
DisableReadyMemo=yes

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\\BrazilianPortuguese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
; Ajuste o caminho se publicar em outra configuração. Este caminho assume Publish padrão Release/net8.0-windows/publish
Source: "..\\..\\MedControl\\bin\\Release\\net8.0-windows\\publish\\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\\MedControl"; Filename: "{app}\\MedControl.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\\MedControl"; Filename: "{app}\\MedControl.exe"; Tasks: desktopicon; WorkingDir: "{app}"

[Run]
Filename: "{app}\\MedControl.exe"; Description: "Iniciar MedControl"; Flags: nowait postinstall skipifsilent

[Registry]
; Cria entrada de inicialização automática (usuário atual)
Root: HKCU; Subkey: "Software\\Microsoft\\Windows\\CurrentVersion\\Run"; ValueType: string; ValueName: "MedControl"; ValueData: "{app}\\MedControl.exe"; Flags: uninsdeletevalue
