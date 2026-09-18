#define MyAppName "Flank"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Flank Technologies, Inc."

#define SsmsRoot "C:\Program Files\Microsoft SQL Server Management Studio 22\Release"
#define SsmsExe SsmsRoot + "\Common7\IDE\SSMS.exe"
#define FlankDir SsmsRoot + "\Common7\IDE\Extensions\Flank"

[Setup]
AppId={{D5B81B12-1A97-4C4A-9B77-6A996793EBC1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={#FlankDir}
DisableDirPage=yes
DisableProgramGroupPage=yes

PrivilegesRequired=admin

OutputDir=output
OutputBaseFilename=Flank-SSMS-Setup
Compression=lzma2
SolidCompression=yes

UninstallDisplayName=Flank for SSMS

[Files]
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.SsmsExtension.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.SsmsExtension.pkgdef"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{#SsmsExe}"; Parameters: "/setup"; StatusMsg: "Registering Flank with SSMS..."; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "{#SsmsExe}"; Parameters: "/setup"; Flags: runhidden waituntilterminated; RunOnceId: "SsmsSetup"

[Code]
function InitializeSetup(): Boolean;
begin
  if not FileExists('{#SsmsExe}') then
  begin
    MsgBox(
      'Flank requires SQL Server Management Studio 22.' + #13#10 + #13#10 +
      'SSMS 22 was not found at:' + #13#10 +
      '{#SsmsExe}',
      mbError,
      MB_OK
    );

    Result := False;
    exit;
  end;

  Result := True;
end;