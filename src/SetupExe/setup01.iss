; Copyright © Leonid Maliutin <mals@live.ru> 2019

#define Configuration "Debug"
#define AppName "CSV 4180"
#define AppVersion "1.0.21"

#define LicensePath "..\..\LICENSE"
#define NetSetupPath "..\Libraries\dotNetFx40_Full_setup.exe"
#define OutputDir "..\Csv4180.Installer.EXE\bin\Release"
#define OutputBaseFilename "CSV_4180_v" + AppVersion

#define appPublisher "Leonid Maliutin"
#define copyright "Copyright © Leonid Maliutin <mals@live.ru> 2019"
#define netSetupFile "dotnetfx.exe"

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
Source: "{#NetSetupPath}";   DestDir: "{tmp}"; DestName: "{#netSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and not IsFrameworkDetected();        AfterInstall: InstallFramework();

[CustomMessages]
en.FrameworkInstallation=Installation of Microsoft .NET Framework 4.
en.RestartFrameworkInstallationMessage=ATTENTION!!!%nTo fully install Microsoft .NET Framework 4, you need to restart your computer. After the computer restarts, run the installation again.
en.FrameworkInstallationError=The installation of Microsoft .NET Framework 4 failed.

[Code]
#include "setup.pas"

type
   TProcessStatus = (psSuccess, psFinish, psError, psFrameworkRebootRequired, psPreviousProgramUninstallationRebootRequired, psProgramInstallationRebootRequired);

var
  ProcessStatus: TProcessStatus;
  HasFramework: Boolean;

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
  HasFramework := IsDotNetDetected('v4\Full', 0);
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