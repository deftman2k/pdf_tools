; installer.iss — PDF_Tools 설치 파일 제작용 스크립트

[Setup]
AppName=PDF_Tools
AppVersion=1.0
DefaultDirName={pf}\PDF_Tools
DefaultGroupName=PDF_Tools
OutputDir=installer
OutputBaseFilename=PDFManagerSetup
Compression=lzma
SolidCompression=yes
DisableDirPage=no
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
SetupIconFile=

[Files]
; PyInstaller로 빌드된 모든 파일을 포함
Source: "dist\pdf_tools\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PDF_Tools"; Filename: "{app}\pdf_tools.exe"
Name: "{group}\프로그램 제거"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\pdf_tools.exe"; Description: "{cm:LaunchProgram,PDF_Tools}"; Flags: nowait postinstall skipifsilent
