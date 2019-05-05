; Copyright © Leonid Maliutin <mals@live.ru> 2019

#ifndef AppId
  #define testCase "csv4180"
#else
  #define testCase "none"
#endif

#if testCase == "csv4180"
  #define AppId "csv4180"
  #define Configuration "Debug"
  #define AppName "CSV 4180"
  #define AppVersion "1.0.20"
  #define OfficeProgram "Excel"
  #define OfficeVersion "14.0"
  #define OfficeYear "2010"
  #define LicensePath "..\Licences\licence (CSV 4180).rtf"
  #define NetSetupPath "..\Libraries\dotNetFx40_Full_setup.exe"
  #define UpgradeCode "{8FCA08C9-3406-4AAB-8481-638CBDBCE918}"
  #define Msi3232Path "..\Csv4180.Installer.EXE\bin\Release\CSV_4180_ru_(Excel_2010_x86)_Windows_x86.msi"
  #define Msi6432Path "..\Csv4180.Installer.EXE\bin\Release\CSV_4180_ru_(Excel_2010_x86)_Windows_x64.msi"
  #define Msi6464Path "..\Csv4180.Installer.EXE\bin\Release\CSV_4180_ru_(Excel_2010_x86)_Windows_x64.msi"
  #define OutputDir "..\Csv4180.Installer.EXE\bin\Release"
  #define OutputBaseFilename "CSV_4180_v" + AppVersion +"_ru_(Excel_" + OfficeYear + ")"
#endif

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
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
Source: "{#NetSetupPath}";   DestDir: "{tmp}"; DestName: "{#netSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and not IsFrameworkDetected();        AfterInstall: InstallFramework();
Source: "{#Msi3232Path}";    DestDir: "{tmp}"; DestName: "{#msiSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and not IsWin64();                    AfterInstall: InstallAddIn();
Source: "{#Msi6432Path}";    DestDir: "{tmp}"; DestName: "{#msiSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and     IsWin64() and not IsHost64(); AfterInstall: InstallAddIn();
Source: "{#Msi6464Path}";    DestDir: "{tmp}"; DestName: "{#msiSetupFile}"; Flags: deleteafterinstall; Check: CheckProcess() and     IsWin64() and     IsHost64(); AfterInstall: InstallAddIn();

[CustomMessages]
russian.NextVersionMessage=Следующая версия программы {#AppName} уже установлена.
russian.CurrentVersionMessage=Дання версия {#AppVersion} программы {#AppName} уже установлена.
russian.HostConditionMessage=Требуется установленная программа Microsoft {#OfficeProgram} {#OfficeYear}.
russian.Host12ConditionMessage=Требуется установленная программа Microsoft {#OfficeProgram} 2007 со 2-ым (или более поздним) пакетом обновлений.
russian.FrameworkInstallation=Установка клиентского профиля Microsoft .NET Framework 4.
russian.RestartFrameworkInstallationMessage=ВНИМАНИЕ!!!%nДля полной установки Microsoft .NET Framework 4 вам нужно перезагрузить компьютер. После перезагрузки компьютера запустите установку снова.
russian.FrameworkInstallationError=Установка Microsoft .NET Framework 4 закончилась неудачей.
russian.AddInInstallation=Установка {#AppName} версии {#AppVersion}.
russian.AddInInstallationError=Установка {#AppName} закончилась неудачей.
russian.InstallInfo=Информация по установке
russian.InternetRequiredCaption=Внимание! Требуется подключение к интернету.
russian.InternetRequiredMessage=Для того чтобы программа {#AppName} работала корректно, требуется установить Microsoft .NET Framework 4.%nДля этого требуется подключение к интернету.
russian.UpdateCaption=Установленная программа {#AppName} версии %1 будет обновлена.
russian.UpdateMessage=На компьютере установлена программа {#AppName} версии %1.%nОна будет обновлена до версии {#AppVersion}.
russian.PreviousProductUninstallation=Удаление {#AppName} версии %1.
russian.RebootPreviousProductUninstallationMessage=Для полного удаления %1 версии {#AppName} вам нужно перезагрузить компьютер. После перезагрузки компьютера запустите установку снова.
russian.UnsuccessfulFinish=Установка программы {#AppName} закончилась неудачей.%n%nНажмите "завершить", чтобы выйти из программы установки.

[Code]
#include "setup.pas"

type
   TProcessStatus = (psSuccess, psFinish, psError, psFrameworkRebootRequired, psPreviousProgramUninstallationRebootRequired, psProgramInstallationRebootRequired);

var
  ProcessStatus: TProcessStatus;
  OfficeBit: TOfficeBit;
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

function IsHost64(): Boolean;
begin
  Result := OfficeBit = ob64;
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
var
  OfficeProgram: TOfficeProgram;
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
  
  if Length(PreviousProgramFolderPath) > 0 then
    WizardForm.DirEdit.Text := PreviousProgramFolderPath
  else
    case OfficeBit of
      ob64: WizardForm.DirEdit.Text := ExpandConstant('{pf64}') + '\DocInOffice\{#AppName} ({#OfficeProgram} {#OfficeYear})';
      ob32: WizardForm.DirEdit.Text := ExpandConstant('{pf32}') + '\DocInOffice\{#AppName} ({#OfficeProgram} {#OfficeYear})';
    end;
  
  UpdateMsgPage := CreateOutputMsgPage(wpSelectDir, CustomMessage('InstallInfo'), FmtMessage(CustomMessage('UpdateCaption'), [PreviousProgramVersion]), FmtMessage(CustomMessage('UpdateMessage'), [PreviousProgramVersion]));
  FrameworkMsgPage := CreateOutputMsgPage(UpdateMsgPage.Id, CustomMessage('InstallInfo'), CustomMessage('InternetRequiredCaption'), CustomMessage('InternetRequiredMessage'));
  
  OfficeProgram := opExcel;

  CheckOfficeInstance(OfficeProgram, '{#OfficeVersion}', OfficeBit);
  if OfficeBit = obNone then
  begin
	ProcessStatus := psError;
	exit;
  end;

  HasFramework := IsDotNetDetected('v4\Full', 0);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Log('PrepareToInstall');

  if OfficeBit = obNone then
  begin
    if '{#OfficeVersion}' = '12.0' then
      Result := CustomMessage('Host12ConditionMessage')
    else
      Result := CustomMessage('HostConditionMessage');
    Exit
  end;

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

procedure DeinitializeSetup();
var 
  HostValue: string;
  ActionValue: string;
begin
  if ProcessStatus = psFinish then
  begin
    case OfficeBit of
      ob32: HostValue := '{#OfficeProgram} {#OfficeYear} x86';
      ob64: HostValue := '{#OfficeProgram} {#OfficeYear} x64';
    end;

    if AnotherVersion = avPrevious then
      ActionValue := 'update'
    else
      ActionValue := 'installation';

    ShowThanksPage('{#AppName}', '{#AppVersion}', HostValue, ActionValue);
  end;
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
	  RaiseException(CustomMessage('AddInInstallationError'));
  except
      ShowExceptionMessage();
      ProcessStatus := psError;
  finally
    WizardForm.StatusLabel.Caption := StatusText;
    WizardForm.ProgressGauge.Style := npbstNormal;
  end;
end;