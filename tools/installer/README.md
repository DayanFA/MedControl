# Instalador do MedControl (Windows)

Este diretório contém o script do Inno Setup para empacotar o MedControl como instalador `.exe` para Windows.

## Pré-requisitos
- .NET 8 SDK (para publicar o app)
- Inno Setup instalado (ISCC) – https://jrsoftware.org/isinfo.php

## Passo 1 – Publicar o aplicativo
Use o perfil de publicação existente (Release, self-contained) ou o task do VS Code.

No VS Code:
1. Abra a paleta de tarefas e execute a task `publish`.
   - Isso gera os arquivos em `MedControl/bin/Release/net8.0-windows/publish/`.

## Passo 2 – Compilar o instalador
Opção A – GUI do Inno Setup:
1. Abra `tools/installer/MedControl.iss` no Inno Setup Compiler.
2. Clique em `Build`.
3. O instalador será gerado em `tools/installer/dist/MedControl-Setup.exe`.

Opção B – Linha de comando (ISCC no PATH):
```cmd
ISCC tools\installer\MedControl.iss
```

## Ajustes comuns
- Versão do app: edite `AppVersion` em `MedControl.iss`.
- Caminho da publicação: se você publicar em outra pasta (por exemplo, win-x64\publish), ajuste a linha `[Files]` do script para apontar para a pasta correta.
- Ícone: se você tiver um `.ico`, adicione `SetupIconFile=` em `[Setup]` e aponte para `MedControl\Assets\app.ico` ou similar.

## Notas
- O instalador cria atalhos no menu iniciar e (opcional) na área de trabalho.
- É recomendado publicar como `self-contained` para evitar dependências no .NET Runtime.
