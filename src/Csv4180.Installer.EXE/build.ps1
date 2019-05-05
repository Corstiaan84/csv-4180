PARAM (
    [string]$SignPassword,
	[string]$MSBuildPath="C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe"
)

$Configuration = "Release"
$binDir = ".\bin\$Configuration\"
$versionParametrs = (& "..\AssemblyInfo.Shared.ps1")
$AppVersion = $versionParametrs.Csv4180Version + "." + $versionParametrs.BuildNumber
$FilesFolderPath = "..\Csv4180.FileExtension\bin\$Configuration\"
$LicensePath = (Get-Item "..\Licences\licence (CSV 4180).rtf").FullName
$NetSetupPath = (Get-Item "..\Libraries\dotNetFx40_Full_setup.exe").FullName

Function ComputeMsiPath($binDir, $Platform, $OfficeYear, $OfficePlatform) {
    $msiPath = (Join-Path $binDir ("CSV_4180_ru_(Excel_${OfficeYear}_${OfficePlatform})_Windows_${Platform}.msi"))
    Return $msiPath
}

If (Test-Path -Path $binDir) {
    Remove-Item $binDir -Recurse
}

[void](New-Item $binDir -ItemType directory)

$binDir=(Get-Item $binDir).FullName

("x86", "x64")| ForEach-Object {
	& "$MSBuildPath" /target:Rebuild /p:Configuration="$Configuration" /p:Platform="$_" "..\CSV 4180.sln" 
	If ($LASTEXITCODE -ne 0) { throw [System.Exception] }
}

(   @{"OfficeVersion"="12.0"; "OfficeYear"="2007"; "UpgradeCode"="{A6EDC28C-7AF3-4BFB-A0A4-1E798FF43715}" },
    @{"OfficeVersion"="14.0"; "OfficeYear"="2010"; "UpgradeCode"="{8FCA08C9-3406-4AAB-8481-638CBDBCE918}" },
    @{"OfficeVersion"="15.0"; "OfficeYear"="2013"; "UpgradeCode"="{36CCCA1D-5C8C-4514-9D52-E69568C0F8FC}" },
    @{"OfficeVersion"="16.0"; "OfficeYear"="2016"; "UpgradeCode"="{4FE29D2C-0BA3-4B8A-881A-3A033FD29E9A}" }

)| ForEach-Object {
    $OfficeVersion = $_.OfficeVersion
    $OfficeYear = $_.OfficeYear
    $UpgradeCode = $_.UpgradeCode

    ( @{"Platform"="x86"; "IsOffice64"="no"}, @{"Platform"="x64"; "IsOffice64"="no"}, @{"Platform"="x64"; "IsOffice64"="yes"} ) | ForEach-Object {
        $Platform = $_.Platform
        $IsOffice64 = $_.IsOffice64
        
		& "$MSBuildPath" /target:Rebuild /p:Configuration="$Configuration" /p:Platform="$Platform" /p:AppVersion="$AppVersion" /p:OfficeVersion="$OfficeVersion" /p:OfficeYear="$OfficeYear" /p:IsOffice64="$IsOffice64" /p:ProductUpgradeCode="$UpgradeCode" /p:FilesFolderPath="$FilesFolderPath" "..\Csv4180.Installer.MSI\Csv4180.Installer.MSI.wixproj"
		If ($LASTEXITCODE -ne 0) { throw [System.Exception] }

        if ($IsOffice64 -eq "yes") { $OfficePlatform = "x64" } else { $OfficePlatform = "x86" }

        $msiPath = ComputeMsiPath $binDir $Platform $OfficeYear $OfficePlatform

        Copy-Item -Path "..\Csv4180.Installer.MSI\bin\$Platform\$Configuration\ru-Ru\CSV_4180_ru.msi" -Destination $msiPath

        If (-Not [string]::IsNullOrEmpty($SignPassword)) {
            & "..\Signing\sign.ps1" -Password $SignPassword -FilePath $msiPath
        }
    }
    
    $setupExeName = "CSV_4180_v${AppVersion}_ru_(Excel_${OfficeYear})"

    &'..\SetupExe\make exe.ps1' -AppId csv4180 -AppName "CSV 4180" -AppVersion $AppVersion -OfficeProgram Excel -OfficeVersion $OfficeVersion -OfficeYear $OfficeYear -LicensePath "$LicensePath" -NetSetupPath "$NetSetupPath" -UpgradeCode $UpgradeCode -Msi3232Path (ComputeMsiPath $binDir x86 ${OfficeYear} x86) -Msi6432Path (ComputeMsiPath $binDir x64 ${OfficeYear} x86) -Msi6464Path (ComputeMsiPath $binDir x64 ${OfficeYear} x64) -OutputPath "$binDir" -OutputFilename $setupExeName

    $setupExePath = Join-Path $binDir "$setupExeName.exe"

    If (-Not [string]::IsNullOrEmpty($SignPassword)) {
        & "..\Signing\sign.ps1" -Password $SignPassword -FilePath $setupExePath -sha1
    }
}