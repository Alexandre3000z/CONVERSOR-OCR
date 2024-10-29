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
Source: "poppler\poppler-24.08.0\Library\bin\*"; DestDir: "{app}\poppler\bin"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Conversor OCR"; Filename: "{app}\index.exe"
Name: "{commondesktop}\Conversor OCR"; Filename: "{app}\index.exe"

[Run]
Filename: "{tmp}\tesseract-setup.exe"; Parameters: "/SILENT"; Flags: waituntilterminated
Filename: "{app}\index.exe"; Description: "Execute o Conversor OCR"; Flags: nowait postinstall skipifsilent

[Code]
const
  SMTO_ABORTIFHUNG = 2;
  WM_SETTINGCHANGE = $1A;

function SendMessageTimeout(hWnd: Integer; Msg: Cardinal; wParam: Integer; lParam: Integer; fuFlags: Cardinal; uTimeout: Cardinal; var lpdwResult: Integer): Integer;
  external 'SendMessageTimeoutA@user32.dll stdcall';

procedure AddToPath(PathToAdd: string);
var
  OldPath: string;
  SendResult: Integer;
begin
  if not RegQueryStringValue(HKLM, 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path', OldPath) then
    OldPath := '';
  if Pos(';' + PathToAdd + ';', ';' + OldPath + ';') = 0 then begin
    if OldPath <> '' then
      OldPath := OldPath + ';';
    RegWriteStringValue(HKLM, 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path', OldPath + PathToAdd);
    SendMessageTimeout($FFFF, WM_SETTINGCHANGE, 0, 0, SMTO_ABORTIFHUNG, 5000, SendResult);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    AddToPath(ExpandConstant('{app}\poppler\bin'));
end;