; Copyright © Leonid Maliutin <mals@live.ru> 2019

#define Configuration "Debug"
#define AppName "CSV 4180"
#define AppVersion "1.0.21"

#define LicensePath "..\..\LICENSE"
#define NetSetupPath "..\Libraries\dotNetFx40_Full_setup.exe"
#define UpgradeCode "{8FCA08C9-3406-4AAB-8481-638CBDBCE918}"
#define Msi32Path "..\Csv4180.Installer.MSI\bin\x86\Release\en-US\CSV_4180.msi"
#define Msi64Path "..\Csv4180.Installer.MSI\bin\x64\Release\en-US\CSV_4180.msi"

#define OutputDir "..\Csv4180.Installer.EXE\bin\Release"
#define OutputBaseFilename "CSV_4180_v" + AppVersion

#define appPublisher "Leonid Maliutin"
#define copyright "Copyright © Leonid Maliutin <mals@live.ru> 2019"
#define netSetupFile "dotnetfx.exe"
#define msiSetupFile "setup.msi"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
VersionInfoVersion={#AppVersion}.0

MinVersion=0,5.01sp3

AppVerName={#AppName} v{#AppVersion}

AppCopyright={#copyright}
AppPublisher={#appPublisher}

Uninstallable=no
CreateAppDir=no

DisableReadyMemo=yes
DisableProgramGroupPage=yes

LicenseFile={#LicensePath}

OutputDir={#OutputDir}
OutputBaseFilename={#OutputBaseFilename}
Compression=none
SolidCompression=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#NetSetupPath}"; DestDir: "{tmp}"; DestName: "{#netSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and not IsFrameworkDetected(); AfterInstall: InstallFramework();
Source: "{#Msi32Path}";    DestDir: "{tmp}"; DestName: "{#msiSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and not IsWin64();             AfterInstall: InstallAddIn();
Source: "{#Msi64Path}";    DestDir: "{tmp}"; DestName: "{#msiSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and     IsWin64();             AfterInstall: InstallAddIn();

[CustomMessages]
en.FrameworkInstallation=Installation of Microsoft .NET Framework 4.
en.RestartFrameworkInstallationMessage=ATTENTION!!!%nTo fully install Microsoft .NET Framework 4, you need to restart your computer. After the computer restarts, run the installation again.
en.FrameworkInstallationError=The installation of Microsoft .NET Framework 4 failed.
en.PreviousProductUninstallation=Uninstall {#AppName} version %1.
en.RebootPreviousProductUninstallationMessage=To completely remove %1 version {#AppName} you need to restart your computer. After the computer restarts, run the installation again.
en.AddInInstallationError=The installation of {#AppName} failed.
en.AddInInstallation=The installation {#AppName} version {#AppVersion}.
en.CurrentVersionMessage=This version {#AppVersion} of {#AppName} is already installed.
en.NextVersionMessage=The next version of {#AppName} is already installed.
en.UnsuccessfulFinish=The installation of {#AppName} failed.%n%nClick "Finish" to exit the setup program.
en.RestartFrameworkInstallationMessage=ATTENTION!!!%nTo fully install Microsoft .NET Framework 4, you need to restart your computer. After the computer restarts, run the installation again.
en.RebootPreviousProductUninstallationMessage=To completely remove %1 version of {#AppName} you need to restart your computer. After the computer restarts, run the installation again.
en.InstallInfo=Information of installation
en.InternetRequiredCaption=Attention! An Internet connection is required.
en.InternetRequiredMessage=To ensure that the program {#AppName} to work correctly, you must install Microsoft .NET Framework 4.%nThis requires an Internet connection.
en.UpdateCaption=The installed program {#AppName} version %1 will be updated.
en.UpdateMessage=The program {#AppName} version %1 is installed on the computer.%nIt will be updated to {#AppVersion}.

[Code]
#include "setup.pas"

type
   TProcessStatus = (psSuccess, psFinish, psError, psFrameworkRebootRequired, psPreviousProgramUninstallationRebootRequired, psProgramInstallationRebootRequired);

var
  ProcessStatus: TProcessStatus;
  HasFramework: Boolean;
  AnotherVersion: TAnotherVersion;
  PreviousProgramId: String;
  PreviousProgramVersion: String;
  PreviousProgramFolderPath: String;
  FrameworkMsgPage: TOutputMsgWizardPage;
  UpdateMsgPage: TOutputMsgWizardPage;

function IsFrameworkDetected(): Boolean;
begin
  Result := HasFramework;
end;

function CheckProcess(): Boolean;
begin
  Result := ProcessStatus = psSuccess;

  if Result then 
    Log('CheckProcess: True')
  else 
    Log('CheckProcess: False');
end;

///////////////////////////////////////////////////////////////////////////////
// Event Functions ////////////////////////////////////////////////////////////

function InitializeSetup(): Boolean;
begin
  Log('InitializeSetup');

  Result := true;
end;

procedure InitializeWizard();
begin
  Log('InitializeWizard');

    Log('Upgrade Code: {#UpgradeCode}');
  if GetPreviousProduct('{#UpgradeCode}', {OUT} PreviousProgramId, {OUT} PreviousProgramVersion, {OUT} PreviousProgramFolderPath) then
  begin
    Log('Previous program id: ' + PreviousProgramId);
    Log('Previous program version: ' + PreviousProgramVersion);
    Log('Previous program folder path: ' + PreviousProgramFolderPath);
    case CompareVersion(PreviousProgramVersion, '{#AppVersion}') of
      0: AnotherVersion := avCurrent;
      1: AnotherVersion := avNext;
      else AnotherVersion := avPrevious;
    end;
  end
  else
  begin
    Log('The version of previous product is not found');
    AnotherVersion := avNone;
  end;
  
  UpdateMsgPage := CreateOutputMsgPage(wpSelectDir, CustomMessage('InstallInfo'), FmtMessage(CustomMessage('UpdateCaption'), [PreviousProgramVersion]), FmtMessage(CustomMessage('UpdateMessage'), [PreviousProgramVersion]));
  FrameworkMsgPage := CreateOutputMsgPage(UpdateMsgPage.Id, CustomMessage('InstallInfo'), CustomMessage('InternetRequiredCaption'), CustomMessage('InternetRequiredMessage'));

  HasFramework := IsDotNetDetected('v4\Full', 0);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Log('PrepareToInstall');

  case AnotherVersion of
    avCurrent: begin
      Result := CustomMessage('CurrentVersionMessage');
      exit
    end;
    avNext: begin
      Result := CustomMessage('NextVersionMessage');
      exit;
    end;
  end;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  case PageId of
  wpWelcome, wpLicense: Result := false;
  UpdateMsgPage.Id:     Result := (ProcessStatus = psError) or (AnotherVersion <> avPrevious);
  FrameworkMsgPage.Id:  Result := (ProcessStatus = psError) or (HasFramework);
  wpReady:              Result := (ProcessStatus = psError);
  end;
end;

function NeedRestart(): Boolean;
begin
  case ProcessStatus of
    psFrameworkRebootRequired,
    psPreviousProgramUninstallationRebootRequired,
	psProgramInstallationRebootRequired:
	  Result := true;
  end;
  
  if Result then Log('NeedRestart(): true')
  else Log('NeedRestart(): false');
end;

function GetCustomSetupExitCode: Integer;
begin
  case ProcessStatus of
    psSuccess: 
	  Result := 0;
	psError: 
	  Result := 1;
    psFrameworkRebootRequired,
	psPreviousProgramUninstallationRebootRequired,
	psProgramInstallationRebootRequired: 
	  Result := ERROR_SUCCESS_REBOOT_REQUIRED;
  end;
  
  Log(Format('Returning exit code %d', [Result]));
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  Log(Format('CurPageChanged(CurPageID: %s)', [GetPageName(CurPageID)]));
  if CurPageID = wpFinished then
  begin
    case ProcessStatus of
	  psError:                                       WizardForm.FinishedLabel.Caption := CustomMessage('UnsuccessfulFinish');
	  psFrameworkRebootRequired:                     WizardForm.FinishedLabel.Caption := CustomMessage('RestartFrameworkInstallationMessage');
    psPreviousProgramUninstallationRebootRequired: WizardForm.FinishedLabel.Caption := FmtMessage(CustomMessage('RebootPreviousProductUninstallationMessage'), [PreviousProgramVersion]);
   // psProgramInstallationRebootRequired:           WizardForm.FinishedLabel.Caption := CustomMessage('');
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
	Log('CurStepChanged')
end;

// Event Functions ////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////

procedure InstallFramework();
var
  statusText: String;
  ResultCode: integer;
begin
  Log('InstallFramework');

  statusText := WizardForm.StatusLabel.Caption;
  WizardForm.StatusLabel.Caption := CustomMessage('FrameworkInstallation');
  WizardForm.ProgressGauge.Style := npbstMarquee;
  try
    InstallFrameworkExe(ExpandConstant('{tmp}\{#netSetupFile}'), ResultCode);
    if ResultCode = ERROR_SUCCESS_REBOOT_REQUIRED then
    begin
      ProcessStatus := psFrameworkRebootRequired;
      MsgBox(CustomMessage('RestartFrameworkInstallationMessage'), mbInformation, MB_OK);
    end
    else if ResultCode <> ERROR_SUCCESS then
      RaiseException(CustomMessage('FrameworkInstallationError'));
  except
    ShowExceptionMessage();
    ProcessStatus := psError;
  finally
    WizardForm.StatusLabel.Caption := statusText;
    WizardForm.ProgressGauge.Style := npbstNormal;
  end;
end;

procedure InstallAddIn();
var
  StatusText: string;
  ResultCode: integer;
  MsiFilePath: string;
begin
  Log('InstallAddIn');

  StatusText := WizardForm.StatusLabel.Caption;
  WizardForm.ProgressGauge.Style := npbstMarquee;

  try
    if AnotherVersion = avPrevious then
    begin
      WizardForm.StatusLabel.Caption := FmtMessage(CustomMessage('PreviousProductUninstallation'), [PreviousProgramVersion]);

      UninstallProduct(PreviousProgramId, {OUT} ResultCode);
      if ResultCode = ERROR_SUCCESS_REBOOT_REQUIRED then
      begin
        ProcessStatus := psPreviousProgramUninstallationRebootRequired;
        MsgBox(FmtMessage(CustomMessage('RebootPreviousProductUninstallationMessage'), [PreviousProgramVersion]), mbInformation, MB_OK);
        exit;
      end
      else if ResultCode <> ERROR_SUCCESS then
      begin
        RaiseException(CustomMessage('AddInInstallationError'));
      end;
    end;

    WizardForm.StatusLabel.Caption := CustomMessage('AddInInstallation');

    MsiFilePath := ExpandConstant('{tmp}') + '\{#msiSetupFile}'

    Log('Insall "' + MsiFilePath + '"');
    InstallMsi(MsiFilePath, WizardForm.DirEdit.Text, {OUT} ResultCode);
    if ResultCode = ERROR_SUCCESS then
      ProcessStatus := psFinish
    else if ResultCode = ERROR_SUCCESS_REBOOT_REQUIRED then
      ProcessStatus := psProgramInstallationRebootRequired
    else
      aiseException(CustomMessage('AddInInstallationError'));
  except
      ShowExceptionMessage();
      ProcessStatus := psError;
  finally
    WizardForm.StatusLabel.Caption := StatusText;
    WizardForm.ProgressGauge.Style := npbstNormal;
  end;
end;