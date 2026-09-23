#define MyAppName "Flank"
#define MyAppVersion "0.1.2"
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
OutputBaseFilename=Flank-SSMS-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes

UninstallDisplayName=Flank for SSMS

CloseApplications=yes
RestartApplications=no


[Files]
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.SsmsExtension.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.SsmsExtension.pkgdef"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion


[Run]
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
  ResultCode: Integer;
  PowerShell: String;
begin
  PowerShell := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');

  if not FileExists(PowerShell) then
  begin
    Result := False;
    exit;
  end;

  if not Exec(
    PowerShell,
    '-NoProfile -NonInteractive -Command "if (Get-Process SSMS -ErrorAction SilentlyContinue) { exit 10 } else { exit 20 }"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  ) then
  begin
    Result := False;
    exit;
  end;

  Result := ResultCode = 10;
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
      'Close SSMS, then run the Flank installer again.',
      mbInformation,
      MB_OK
    );

    exit;
  end;

  Result := True;
end;


procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    WizardForm.StatusLabel.Caption := 'Registering Flank with SSMS...';

    { SSMS /setup can take several seconds and doesn't expose progress,
      so show an indeterminate progress bar while we wait. }
    WizardForm.ProgressGauge.Style := npbstMarquee;

    try
      Exec(
        ExpandConstant('{#SsmsExe}'),
        '/setup',
        '',
        SW_HIDE,
        ewWaitUntilTerminated,
        ResultCode
      );
    finally
      WizardForm.ProgressGauge.Style := npbstNormal;
    end;
  end;
end;


function InitializeUninstall(): Boolean;
begin
  Result := False;

  if IsSsmsRunning() then
  begin
    MsgBox(
      'SQL Server Management Studio is currently running.' +
      Chr(13) + Chr(10) + Chr(13) + Chr(10) +
      'Close SSMS, then uninstall Flank again.',
      mbInformation,
      MB_OK
    );

    exit;
  end;

  Result := True;
end;