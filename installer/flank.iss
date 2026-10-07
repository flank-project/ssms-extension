#define MyAppName "Flank"
#define MyAppVersion "0.2.4"
#define MyAppPublisher "Flank Technologies, Inc."

[Setup]
AppId={{D5B81B12-1A97-4C4A-9B77-6A996793EBC1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={code:GetFlankDir}
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
Source: "..\Flank.SsmsExtension\bin\Release\net472\Flank.Ssrs.dll"; DestDir: "{app}"; Flags: ignoreversion


[Run]
Filename: "{code:GetSsmsExe}"; \
    Description: "Launch SQL Server Management Studio"; \
    Flags: nowait postinstall skipifsilent


[Code]

var
  SsmsExePath: String;
  SsmsInstallRoot: String;


function FindSsms22(): Boolean;
var
  VsWhere: String;
  TempFile: String;
  Command: String;
  ResultCode: Integer;
  Output: AnsiString;
  InstallPath: String;
begin
  Result := False;

  { ---------------------------------------------------------
    First try the standard SSMS 22 installation location.
    --------------------------------------------------------- }

  SsmsInstallRoot :=
    ExpandConstant(
      '{autopf}\Microsoft SQL Server Management Studio 22\Release'
    );

  SsmsExePath :=
    SsmsInstallRoot + '\Common7\IDE\SSMS.exe';

  if FileExists(SsmsExePath) then
  begin
    Result := True;
    exit;
  end;


  { ---------------------------------------------------------
    SSMS 22 uses the Visual Studio Installer infrastructure.
    Use vswhere.exe to discover installations in non-default
    locations.

    vswhere itself normally lives under Program Files (x86),
    regardless of where SSMS was installed.
    --------------------------------------------------------- }

  VsWhere :=
    ExpandConstant(
      '{pf32}\Microsoft Visual Studio\Installer\vswhere.exe'
    );

  if not FileExists(VsWhere) then
    exit;


  { Have vswhere write the installation path to a temp file. }

  TempFile :=
    ExpandConstant('{tmp}\flank-ssms-install-path.txt');

  DeleteFile(TempFile);

  Command :=
    '-latest ' +
    '-products Microsoft.SQLServer.ManagementStudio ' +
    '-property installationPath ' +
    '>"' + TempFile + '"';


  { Redirection requires cmd.exe. }

  if not Exec(
    ExpandConstant('{cmd}'),
    '/C ""' + VsWhere + '" ' + Command + '"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  ) then
    exit;

  if ResultCode <> 0 then
    exit;

  if not LoadStringFromFile(TempFile, Output) then
    exit;

  InstallPath := Trim(String(Output));

  if InstallPath = '' then
    exit;


  { ---------------------------------------------------------
    installationPath should be the SSMS root, e.g.

      D:\Apps\Microsoft SQL Server Management Studio 22\Release

    SSMS.exe is beneath Common7\IDE.
    --------------------------------------------------------- }

  SsmsInstallRoot := InstallPath;
  SsmsExePath :=
    SsmsInstallRoot + '\Common7\IDE\SSMS.exe';

  Result := FileExists(SsmsExePath);
end;


function GetSsmsExe(Param: String): String;
begin
  Result := SsmsExePath;
end;


function GetFlankDir(Param: String): String;
begin
  Result :=
    SsmsInstallRoot + '\Common7\IDE\Extensions\Flank';
end;


function IsSsmsRunning(): Boolean;
var
  ResultCode: Integer;
  PowerShell: String;
begin
  PowerShell :=
    ExpandConstant(
      '{sys}\WindowsPowerShell\v1.0\powershell.exe'
    );

  if not FileExists(PowerShell) then
  begin
    Result := False;
    exit;
  end;

  if not Exec(
    PowerShell,
    '-NoProfile -NonInteractive -Command ' +
    '"if (Get-Process SSMS -ErrorAction SilentlyContinue) ' +
    '{ exit 10 } else { exit 20 }"',
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

  { Discover SSMS before Setup evaluates DefaultDirName. }

  if not FindSsms22() then
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
    WizardForm.StatusLabel.Caption :=
      'Registering Flank with SSMS...';

    { SSMS /setup can take several seconds and doesn't expose
      progress, so show an indeterminate progress bar. }

    WizardForm.ProgressGauge.Style := npbstMarquee;

    try
      if (not Exec(
        SsmsExePath,
        '/setup',
        '',
        SW_HIDE,
        ewWaitUntilTerminated,
        ResultCode
      )) or (ResultCode <> 0) then
      begin
        MsgBox(
          'Flank could not be registered with SSMS.',
          mbError,
          MB_OK
        );
      end;
    finally
      WizardForm.ProgressGauge.Style := npbstNormal;
    end;
  end;
end;


function InitializeUninstall(): Boolean;
begin
  Result := False;

  { Variables are not preserved between installation and
    uninstallation, so rediscover SSMS here. }

  if not FindSsms22() then
  begin
    MsgBox(
      'SQL Server Management Studio 22 could not be found.',
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
      'Close SSMS, then uninstall Flank again.',
      mbInformation,
      MB_OK
    );

    exit;
  end;

  Result := True;
end;


procedure CurUninstallStepChanged(
  CurUninstallStep: TUninstallStep
);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    if (not Exec(
      SsmsExePath,
      '/setup',
      '',
      SW_HIDE,
      ewWaitUntilTerminated,
      ResultCode
    )) or (ResultCode <> 0) then
    begin
      MsgBox(
        'Flank could not be unregistered from SSMS.',
        mbError,
        MB_OK
      );
    end;
  end;
end;