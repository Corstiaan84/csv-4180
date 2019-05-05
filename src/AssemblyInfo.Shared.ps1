
function GetValue($name) {
	$value = Select-String -Path .\AssemblyInfo.Shared.cs -Pattern ($name + '\s*=\s*"(?<value>(\d|\.)+)"') | %{$_.Matches} | %{$_.Groups["value"].Value}
	Return ($value)
}

$currentPath = Get-Location
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition

Set-Location -Path $scriptPath

try {
	# $buildNumber = Select-String -Path .\AssemblyInfo.Shared.cs -Pattern 'BuildNumber = "(?<bn>\d+)"' | %{$_.Matches} | %{$_.Groups["bn"].Value}

	$buildNumber = GetValue("BuildNumber")
	$MSOfficeExtensionVersion = GetValue("MSOfficeExtensionVersion")
	$ActivationVersion = GetValue("ActivationVersion")
	$Fb2WordViewerVersion = GetValue("Fb2WordViewerVersion")
	$Csv4180Version = GetValue("Csv4180Version")
	$Fb2AddInVersion = GetValue("Fb2AddInVersion")

	(@{ 
		"BuildNumber" = $buildNumber; 
		"MSOfficeExtensionVersion" = $MSOfficeExtensionVersion;
		"ActivationVersion" = $ActivationVersion;
		"Fb2WordViewerVersion" = $Fb2WordViewerVersion;
		"Csv4180Version" = $Csv4180Version;
		"Fb2AddInVersion" = $Fb2AddInVersion;
	})
}
Finally {
    Set-Location -Path $currentPath
}