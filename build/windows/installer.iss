; Inno Setup script for Speedy Scandocs
; Requires Inno Setup 6: https://jrsoftware.org/isinfo.php
;
; Before running this script:
;  1. Build the PyInstaller output first (run build_windows.bat, or manually:
;     pyinstaller build\windows\scandocs.spec --clean)
;  2. Place the Tesseract installer in build\windows\Installers\
;     Download from: https://github.com/UB-Mannheim/tesseract/wiki
;     File should be named: tesseract-ocr-w64-setup.exe

#define AppName      "Speedy Scandocs"
#define AppVersion   "1.0"
#define AppPublisher "GDJ"
#define AppExeName   "SpeedyScandocs.exe"
#define DistDir      "..\..\dist\SpeedyScandocs"

[Setup]
AppId={{A3F1C2D4-8B5E-4F6A-9C3D-1E2F7A8B9C0D}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=SpeedyScandocsSetup
SetupIconFile=..\..\assets\GDJ Logo.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Require 64-bit Windows
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Main application (PyInstaller one-folder output)
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; (Tesseract is bundled inside the PyInstaller app folder — no separate install needed)

[Icons]
Name: "{group}\{#AppName}";   Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\GDJ Logo.ico"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\assets\GDJ Logo.ico"; Tasks: desktopicon

[Run]
; Offer to launch app when installer finishes
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; \
  Flags: nowait postinstall skipifsilent

