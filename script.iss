[Setup]
AppName=Conversor OCR
AppVersion=1.0
DefaultDirName={pf}\Conversor OCR
DefaultGroupName=Conversor OCR
OutputBaseFilename=ConversorOCRSetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\index.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "tesseract-setup.exe"; DestDir: "{tmp}"; Flags: ignoreversion

[Icons]
Name: "{group}\Conversor OCR"; Filename: "{app}\index.exe"
Name: "{commondesktop}\Conversor OCR"; Filename: "{app}\index.exe"

[Run]
Filename: "{tmp}\tesseract-setup.exe"; Parameters: "/SILENT"; Flags: waituntilterminated
Filename: "{app}\index.exe"; Description: "Execute o Conversor OCR"; Flags: nowait postinstall skipifsilent