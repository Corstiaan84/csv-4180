Param (
    [Parameter(Mandatory=$True)][string]$AppId,
    [Parameter(Mandatory=$True)][string]$AppName,
    [Parameter(Mandatory=$True)][string]$AppVersion,
    [Parameter(Mandatory=$True)][string]$OfficeVersion,
    [Parameter(Mandatory=$True)][string]$OfficeYear,
    [Parameter(Mandatory=$True)][string]$OfficeProgram,
    [Parameter(Mandatory=$True)][string]$LicensePath,
    [Parameter(Mandatory=$True)][string]$NetSetupPath,
    [Parameter(Mandatory=$True)][string]$UpgradeCode,
    [Parameter(Mandatory=$True)][string]$Msi3232Path,
    [Parameter(Mandatory=$True)][string]$Msi6432Path,
    [Parameter(Mandatory=$True)][string]$Msi6464Path,
    [Parameter(Mandatory=$True)][string]$OutputPath,
    [Parameter(Mandatory=$True)][string]$OutputFilename
)

$currentPath = Get-Location
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition

Set-Location -Path $scriptPath

try {

& "..\Libraries\Inno Setup 5\ISCC.exe" /DAppId="$AppId" /DConfiguration="$Configuration" /DAppName="$AppName" /DAppVersion="$AppVersion" /DOfficeVersion="$OfficeVersion" /DOfficeYear="$OfficeYear" /DOfficeProgram="$OfficeProgram" /DLicensePath="$LicensePath" /DNetSetupPath="$NetSetupPath" /DUpgradeCode="$UpgradeCode" /DMsi3232Path="$msi3232Path" /DMsi6432Path="$msi6432Path" /DMsi6464Path="$msi6464Path" /DOutputDir="$OutputPath" /DOutputBaseFilename="$OutputFilename" setup.iss
If ($LASTEXITCODE -ne 0) {
    throw [System.Exception]
}

}
Finally {
    Set-Location -Path $currentPath
}
