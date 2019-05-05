Param (
    [Parameter(Mandatory=$False)][string]$SignFilePath = "CodeSigning_2019-08-16.p12",
    [Parameter(Mandatory=$True )][string]$Password,
    [Parameter(Mandatory=$True )][string]$FilePath,
                                 [switch]$sha1
)

$currentPath = Get-Location
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition

Set-Location -Path $scriptPath

try {
	$sdkPath="..\Libraries\WinSDK\x86\signtool.exe"
	
	# Параметры команды "sign": https://docs.microsoft.com/ru-ru/dotnet/framework/tools/signtool-exe#sign
	
	If ($sha1) {
		& "${sdkPath}" sign /f "$SignFilePath" /p "$Password" /fd sha1 /t http://timestamp.verisign.com/scripts/timstamp.dll /v "$FilePath"
		If ($LASTEXITCODE -ne 0) {
			throw [System.Exception] ("Error: Siginig sha1: exit code " + $LASTEXITCODE)
		}
		
		& "${sdkPath}" sign /as /f "$SignFilePath" /p "$Password" /fd sha256 /tr "http://sha256timestamp.ws.symantec.com/sha256/timestamp" /td sha256 /v "$FilePath"
		If ($LASTEXITCODE -ne 0) {
			throw [System.Exception] ("Error: Siginig sha256: exit code " + $LASTEXITCODE)
		}
	} else {
		& "${sdkPath}" sign /f "$SignFilePath" /p "$Password" /fd sha256 /tr "http://sha256timestamp.ws.symantec.com/sha256/timestamp" /td sha256 /v "$FilePath"
		If ($LASTEXITCODE -ne 0) {
			throw [System.Exception] ("Error: Siginig sha256: exit code " + $LASTEXITCODE)
		}
	}
}
Finally {
    Set-Location -Path $currentPath
}