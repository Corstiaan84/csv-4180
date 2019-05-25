function CompareVersion(V1, V2: string): Integer;
var
  P, N1, N2: Integer;
begin
  Result := 0;
  while (Result = 0) and ((V1 <> '') or (V2 <> '')) do
  begin
    P := Pos('.', V1);
    if P > 0 then
    begin
      N1 := StrToInt(Copy(V1, 1, P - 1));
      Delete(V1, 1, P);
    end
      else
    if V1 <> '' then
    begin
      N1 := StrToInt(V1);
      V1 := '';
    end
      else
    begin
      N1 := 0;
    end;

    P := Pos('.', V2);
    if P > 0 then
    begin
      N2 := StrToInt(Copy(V2, 1, P - 1));
      Delete(V2, 1, P);
    end
      else
    if V2 <> '' then
    begin
      N2 := StrToInt(V2);
      V2 := '';
    end
      else
    begin
      N2 := 0;
    end;

    if N1 < N2 then Result := -1
      else
    if N1 > N2 then Result := 1;
  end;
end;

function IsDotNetDetected(version: string; service: cardinal): boolean;
// Source: http://kynosarges.org/DotNetVersion.html
// Indicates whether the specified version and service pack of the .NET Framework is installed.
//
// version -- Specify one of these strings for the required .NET Framework version:
//    'v1.1'          .NET Framework 1.1
//    'v2.0'          .NET Framework 2.0
//    'v3.0'          .NET Framework 3.0
//    'v3.5'          .NET Framework 3.5
//    'v4\Client'     .NET Framework 4.0 Client Profile
//    'v4\Full'       .NET Framework 4.0 Full Installation
//    'v4.5'          .NET Framework 4.5
//    'v4.5.1'        .NET Framework 4.5.1
//    'v4.5.2'        .NET Framework 4.5.2
//    'v4.6'          .NET Framework 4.6
//    'v4.6.1'        .NET Framework 4.6.1
//    'v4.6.2'        .NET Framework 4.6.2
//
// service -- Specify any non-negative integer for the required service pack level:
//    0               No service packs required
//    1, 2, etc.      Service pack 1, 2, etc. required
var
    key, versionKey: string;
    install, release, serviceCount, versionRelease: cardinal;
    success: boolean;
begin
    versionKey := version;
    versionRelease := 0;

    // .NET 1.1 and 2.0 embed release number in version key
    if version = 'v1.1' then begin
        versionKey := 'v1.1.4322';
    end else if version = 'v2.0' then begin
        versionKey := 'v2.0.50727';
    end

    // .NET 4.5 and newer install as update to .NET 4.0 Full
    else if Pos('v4.', version) = 1 then begin
        versionKey := 'v4\Full';
        case version of
          'v4.5':   versionRelease := 378389;
          'v4.5.1': versionRelease := 378675; // 378758 on Windows 8 and older
          'v4.5.2': versionRelease := 379893;
          'v4.6':   versionRelease := 393295; // 393297 on Windows 8.1 and older
          'v4.6.1': versionRelease := 394254; // 394271 on Windows 8.1 and older
          'v4.6.2': versionRelease := 394802; // 394806 on Windows 8.1 and older
        end;
    end;

    // installation key group for all .NET versions
    key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\' + versionKey;

    // .NET 3.0 uses value InstallSuccess in subkey Setup
    if Pos('v3.0', version) = 1 then begin
        success := RegQueryDWordValue(HKLM, key + '\Setup', 'InstallSuccess', install);
    end else begin
        success := RegQueryDWordValue(HKLM, key, 'Install', install);
    end;

    // .NET 4.0 and newer use value Servicing instead of SP
    if Pos('v4', version) = 1 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Servicing', serviceCount);
    end else begin
        success := success and RegQueryDWordValue(HKLM, key, 'SP', serviceCount);
    end;

    // .NET 4.5 and newer use additional value Release
    if versionRelease > 0 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Release', release);
        success := success and (release >= versionRelease);
    end;

    result := success and (install = 1) and (serviceCount >= service);
end;

type
	TAnotherVersion = (avNone, avCurrent, avNext, avPrevious);

function MsiEnumRelatedProducts(lpUpgradeCode: string; dwReserved: DWORD; iProductIndex: DWORD; lpProductBuf: string): UINT;
  external 'MsiEnumRelatedProductsW@msi.dll stdcall';

function MsiGetProductInfo(szProduct: string; szProperty: String; lpValueBuf: String; var pcchValueBuf: DWORD): UINT;
  external 'MsiGetProductInfoW@msi.dll stdcall';
  
const // Full list: https://msdn.microsoft.com/en-us/library/windows/desktop/aa376931(v=vs.85).aspx
  ERROR_SUCCESS = 0;
  ERROR_SUCCESS_REBOOT_INITIATED = 1641;
  ERROR_SUCCESS_REBOOT_REQUIRED = 3010;

  INSTALLPROPERTY_VERSIONSTRING = 'VersionString';
  INSTALLPROPERTY_INSTALLLOCATION = 'InstallLocation';

function GetPreviousProduct(UpgradeCode: string; var PreviousProductId: string; var PreviousProductVersion: string; var Location: string): Boolean;
var
  ProductBuf: string;
  VersionBuf: string;
  LocationBuf: string;
  PcchValueBuf: DWORD;
begin
  SetLength(ProductBuf, 39);
  if MsiEnumRelatedProducts(upgradeCode, 0, 0, ProductBuf) <> ERROR_SUCCESS then
  begin
    PreviousProductId := '';
    PreviousProductVersion := '';
    Result := false;
    exit;
  end;

  SetLength(VersionBuf, 255);
  PcchValueBuf := Length(VersionBuf);
  if MsiGetProductInfo(ProductBuf, INSTALLPROPERTY_VERSIONSTRING, VersionBuf, PcchValueBuf) <> ERROR_SUCCESS then
  begin
	Log('The VersionString is not received.');
    RaiseException('The VersionString is not received.');
    exit;
  end;
  
  SetLength(VersionBuf, PcchValueBuf);
  
  SetLength(LocationBuf, 255)
  PcchValueBuf := Length(LocationBuf);
  if MsiGetProductInfo(ProductBuf, INSTALLPROPERTY_INSTALLLOCATION, LocationBuf, PcchValueBuf) <> ERROR_SUCCESS then
  begin
	Log('The VersionString is not received.');
    RaiseException('The VersionString is not received.');
    exit;
  end;
  
  SetLength(LocationBuf, PcchValueBuf);

  PreviousProductId := ProductBuf;
  PreviousProductVersion := VersionBuf;
  Location := LocationBuf;
  Result := true;
end;

procedure UninstallProduct(ProductId: string; var ResultCode: integer);
begin
  Exec('msiexec.exe', '/quiet /x "' + ProductId + '"', '', SW_SHOW, ewWaitUntilTerminated, ResultCode);
  Log('The previous program is uninstalled with code: ' + IntToStr(ResultCode) + '.');
end;

procedure InstallFrameworkExe(FilePath: string; var ResultCode: integer);
begin
  Exec(FilePath, '/norestart', '', SW_SHOW, ewWaitUntilTerminated, resultCode);
  Log('Framework is installed with code: ' + IntToStr(ResultCode) + '.');
end;

procedure InstallMsi(FilePath, InstallFolderPath: string; var ResultCode: integer);
begin
  Exec('msiexec.exe', '/quiet /i "' + FilePath + '" AGREETOLICENSE=yes INSTALLFOLDER="' + InstallFolderPath + '"' , '', SW_SHOW, ewWaitUntilTerminated, ResultCode);
  Log('The program is installed with code: ' + IntToStr(ResultCode) + '.');
end;

function GetPageName(PageID: Integer): string;
begin
  case PageId of
	wpWelcome:            Result := 'wpWelcome';
	wpLicense:            Result := 'wpLicense';
	wpPassword:           Result := 'wpPassword';
	wpInfoBefore:         Result := 'wpInfoBefore';
	wpUserInfo:           Result := 'wpUserInfo';
	wpSelectDir:          Result := 'wpSelectDir';
	wpSelectComponents:   Result := 'wpSelectComponents';
	wpSelectProgramGroup: Result := 'wpSelectProgramGroup';
	wpSelectTasks:        Result := 'wpSelectTasks';
	wpReady:              Result := 'wpReady';
	wpPreparing:          Result := 'wpPreparing';
	wpInstalling:         Result := 'wpInstalling';
	wpInfoAfter:          Result := 'wpInfoAfter';
	wpFinished:           Result := 'wpFinished';
    else                  Result := IntToStr(PageID)
  end;
end;