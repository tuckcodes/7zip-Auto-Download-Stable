<#
	Some code logic/inspiration recycled from https://github.com/glombard/Scripts/blob/master/PowerShell-Installers/Get-7zip.ps1
#>


# Global variables
$Script:instPath = "$env:ProgramFiles\7-Zip\7z.exe"
$Script:wshell = New-Object -ComObject Wscript.Shell

# Determines the download URL for the latest stable version from the 7-zip site
# and then compares the local version with the newest stable release.
function Script:GetDownloadURL
{
	# Local variables
	$Script:installedVersion = (Get-Item $instPath).LastWriteTime.Date

	$Script:web = New-Object System.Net.WebClient
	$page = $web.DownloadString("http://www.7-zip.org/download.html")
	$64bit = ''
	
	if ($env:PROCESSOR_ARCHITECTURE -match '64')
	{
		$64bit = 'x64'
	}
	$downloadRoot = "http://www.7-zip.org/a/"
	$pattern = "(7z.*?${64bit}\.msi)"
	
	$Script:url = $downloadRoot + ($page | Select-String -Pattern $pattern | Select-Object -ExpandProperty Matches -First 1 | ForEach-Object { $_.Value })

	$webRequest = [System.Net.HttpWebRequest]::Create($url);
	$webRequest.Method = "HEAD";
	$webResponse = $webRequest.GetResponse()
	$Script:remoteStableVersion = ($webResponse.LastModified.Date)
	$webResponse.Close()

}


# Fetches and installs the latest stable version
function Global:InstallStable
{
	
	$file = "$env:TEMP\7z.msi"
	if (Test-Path $file)
	{
		Remove-Item $file | Out-Null
	}
	
	$web.DownloadFile($url, $file)	
	$cmd = "$file /passive"
	
	Invoke-Expression $cmd | Out-Null
	
	while (!(Test-Path $instPath))
	{
		Start-Sleep -Seconds 10
	}
	
	$wshell.Popup("The latest version of 7-Zip has been installed.", 0, "Done")
	
}


<#
	//						//
	// START OF APPLICATION //
	//						//
#>

# Dependency function that other functions will require
GetDownloadURL

# Determines if 7-zip is up-to-date or even installed at all
if (!(Test-Path $instPath))
{
	# Determined that it was not installed at all, so install from scratch
	InstallStable
}

# Previous installation detected
else
{
	if ($remoteStableVersion -gt $installedVersion)
	{
		# Current installation is not the newest version, so installs newest
		InstallStable
	}
	else
	{
		# Current installation is already newest version
		$wshell.Popup("Latest version of 7-Zip is already installed.", 0, "Done")
	}	
}
