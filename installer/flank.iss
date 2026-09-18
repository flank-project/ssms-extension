#define MyAppName "Flank"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Flank Technologies, Inc."

#define SsmsRoot "{autopf}\Microsoft SQL Server Management Studio 22\Release"
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
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

OutputDir=output
OutputBaseFilename=Flank-SSMS-Setup
Compression=lzma2
SolidCompression=yes

UninstallDisplayName=Flank for SSMS
UninstallFilesDir={app}\uninstall

CloseApplications=yes
RestartApplications=no

[Files]
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.SsmsExtension.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.SsmsExtension.pkgdef"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion

[Run]
; Register Flank with SSMS
Filename: "{#SsmsExe}"; \
    Parameters: "/setup"; \
    StatusMsg: "Registering Flank with SSMS..."; \
    Flags: runhidden waituntilterminated

; Optional launch from Finish page
Filename: "{#SsmsExe}"; \
    Description: "Launch SQL Server Management Studio"; \
    Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{#SsmsExe}"; \
    Parameters: "/setup"; \
    Flags: runhidden waituntilterminated; \
    RunOnceId: "SsmsSetup"

[Code]

function IsSsmsRunning(): Boolean;
var
  WbemLocator: Variant;
  WbemServices: Variant;
  Processes: Variant;
begin
  Result := False;

  try
    WbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    WbemServices := WbemLocator.ConnectServer('.', 'root\CIMV2');

    Processes := WbemServices.ExecQuery(
      'SELECT * FROM Win32_Process WHERE Name="SSMS.exe"');

    Result := Processes.Count > 0;
  except
    { If process detection fails, don't block installation. }
    Result := False;
  end;
end;


function InitializeSetup(): Boolean;
begin
  Result := False;

  if not FileExists(ExpandConstant('{#SsmsExe}')) then
  begin
    MsgBox(
      'Flank requires SQL Server Management Studio 22.' +
      Chr(13) + Chr(10) + Chr(13) + Chr(10) +
      'SSMS 22 could not be found.',
      mbError,
      MB_OK
    );

    exit;
  end;

  if IsSsmsRunning() then
  begin
    MsgBox(
      'SQL Server Management Studio is currently running.' +
      Chr(13) + Chr(10) + Chr(13) + Chr(10) +
      'Please close SSMS, then run the Flank installer again.',
      mbInformation,
      MB_OK
    );

    exit;
  end;

  Result := True;
end;


function InitializeUninstall(): Boolean;
begin
  Result := False;

  if IsSsmsRunning() then
  begin
    MsgBox(
      'SQL Server Management Studio is currently running.' +
      Chr(13) + Chr(10) + Chr(13) + Chr(10) +
      'Please close SSMS before uninstalling Flank.',
      mbInformation,
      MB_OK
    );

    exit;
  end;

  Result := True;
end;