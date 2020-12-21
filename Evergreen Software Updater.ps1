# ***********************************************************
# D. Mohrmann, S&L Firmengruppe, Twitter: @mohrpheus78
# Download Software packages with Evergreen powershell module
# ***********************************************************

<#
.SYNOPSIS
This script downloads software packages if new versions are available.
		
.Description
The script uses the excellent Powershell Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. 
To update a software package just switch from 0 to 1 in the section "Select software to download".
A new folder for every single package will be created, together with a version file, a download date file and a log file. If a new version is available
the scriot checks the version number and will update the package.

.EXAMPLE
'$NotePadPlusPlus = 1' Downloads Notepad++

.NOTES
Thanks to Trond Eric Haarvarstein, I used some code from his great Automation Framework!
Many thanks to Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein for the module!
https://github.com/aaronparker/Evergreen
You can run this script daily with a scheduled task.
Run as admin!
#>

# Do you run the script as admin?
# ========================================================================================================================================
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

if ($myWindowsPrincipal.IsInRole($adminRole))
   {
    # OK, runs as admin
    Write-Verbose "OK, script is running with Admin rights" -Verbose
    Write-Output ""
   }

else
   {
    # Script doesn't run as admin, stop!
    Write-Verbose "Error! Script is NOT running with Admin rights!" -Verbose
    BREAK
   }
# ========================================================================================================================================

Write-Verbose "Setting Variables" -Verbose
Write-Output ""

# Variables
$Date = $Date = Get-Date -UFormat "%m.%d.%Y"

# Select software to download
$NotePadPlusPlus = 1
$GoogleChrome = 1
$MSEdge = 1
$VLCPlayer =1
$BISF = 1
$WorkspaceApp_Current_Relase = 1 
$WorkspaceApp_LTSR_Relase = 1 
$7ZIP = 1
$AdobeReaderDC_MUI = 1 # Only MSP Updates
$FSLogix = 1
$MSTeams = 1
$OneDrive = 1
$OfficeDT = 1 # Office Deployment Toolkit for installing Office 365
$VMWareTools = 1
$OpenJDK = 1
$OracleJava8 = 1
$KeepPass = 1
$mRemoteNG = 1
$TreeSizeFree = 1

# Disable progress bar while downloading
$ProgressPreference = 'SilentlyContinue'

# Install/Update Evergreen module
Write-Verbose "Installing/updating Evergreen module... please wait" -Verbose
Write-Output ""
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies}
if (!(Get-Module -ListAvailable -Name Evergreen)) {Install-Module Evergreen -Force | Import-Module Evergreen}
Update-Module Evergreen -force

Write-Output "Starting downloads..."
Write-Output ""


# Download Notepad ++
IF ($NotePadPlusPlus -eq 1) {
$Product = "NotePadPlusPlus"
$PackageName = "NotePadPlusPlus_x64"
$Notepad = Get-NotepadPlusPlus | Where-Object {$_.Architecture -eq "x64" -and $_.URI -match ".exe"}
$Version = $Notepad.Version
$URL = $Notepad.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Get-ChildItem "$PSScriptRoot\$Product\" -Exclude lang | Remove-Item -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download Chrome
IF ($GoogleChrome -eq 1) {
$Product = "Google Chrome"
$ChromeURL = Get-GoogleChrome | Where-Object {$_.Architecture -eq "x64"} | Select-Object -ExpandProperty URI
$Version = (Get-GoogleChrome | Where-Object {$_.Architecture -eq "x64"}).Version
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $ChromeURL -OutFile ("$PSScriptRoot\$Product\" + ($ChromeURL | Split-Path -Leaf))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download MS Edge
IF ($MSEdge -eq 1) {
$Product = "MS Edge"
$EdgeURL = Get-MicrosoftEdge | Where-Object {$_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "x64"}
$EdgeURL  = $EdgeURL | Sort-Object -Property Version -Descending | Select-Object -First 1
$Version =  (Get-MicrosoftEdge | Where-Object {$_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "x64"}).Version
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue 
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $EdgeURL.Uri -OutFile ("$PSScriptRoot\$Product\" + ($EdgeURL.URI | Split-Path -Leaf))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download VLC Player
IF ($VLCPlayer -eq 1) {
$Product = "VLC Player"
$PackageName = "VLC-Player"
$VLC = Get-VideoLanVlcPlayer | Where-Object {$_.Platform -eq "Windows"  -and $_.Architecture -eq "x64" -and $_.Type -eq "MSI"}
$Version = $VLC.Version
$URL = $VLC.uri
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download BIS-F
IF ($BISF -eq 1) {
$Product = "BIS-F"
$PackageName = "setup-BIS-F"
$BISF = Get-BISF| Where-Object {$_.URI -like "*msi*"}
$Version = $BISF.Version
$URL = $BISF.uri
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Exclude *.ps1, *.lnk -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download WorkspaceApp Current
IF ($WorkspaceApp_Current_Relase -eq 1) {
$Product = "WorkspaceApp"
$PackageName = "CitrixWorkspaceApp"
$WSA = Get-CitrixWorkspaceApp | Where-Object {$_.Title -like "*Workspace*" -and "*Current*" -and $_.Platform -eq "Windows" -and $_.Title -like "*Current*" }
$Version = $WSA.Version
$URL = $WSA.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\Current\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product\Windows\Current")) {New-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\Current" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\Citrix\$Product\Windows\Current\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\Citrix\$Product\Windows\Current\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\Current" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\Current\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version Current Release" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\Citrix\$Product\Windows\Current\" + ($Source))
Copy-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\Current\CitrixWorkspaceApp.exe" -Destination "$PSScriptRoot\Citrix\$Product\Windows\Current\CitrixWorkspaceAppWeb.exe" | Out-Null
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download WorkspaceApp LTSR
IF ($WorkspaceApp_LTSR_Relase -eq 1) {
$Product = "WorkspaceApp"
$PackageName = "CitrixWorkspaceApp"
$WSA = Get-CitrixWorkspaceApp | Where-Object {$_.Title -like "*Workspace*" -and "*LTSR*" -and $_.Platform -eq "Windows" -and $_.Title -like "*LTSR*" }
$Version = $WSA.Version
$URL = $WSA.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product LTSR" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR")) {New-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\Citrix\$Product\Windows\LTSR\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\Citrix\$Product\Windows\LTSR\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version LTSR Release" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\Citrix\$Product\Windows\LTSR\" + ($Source))
Copy-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR\CitrixWorkspaceApp.exe" -Destination "$PSScriptRoot\Citrix\$Product\Windows\LTSR\CitrixWorkspaceAppWeb.exe" | Out-Null
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download 7-ZIP
IF ($7ZIP -eq 1) {
$Product = "7-Zip"
$PackageName = "7-Zip_x64"
$7Zip = Get-7zip | Where-Object {$_.Architecture -eq "x64" -and $_.URI -like "*exe*"}
$Version = $7Zip.Version
$URL = $7Zip.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product.txt" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download Adobe Reader DC MUI Update
IF ($AdobeReaderDC_MUI -eq 1) {
$Product = "Adobe Reader DC MUI"
$PackageName = "Adobe_DC_MUI_Update"
$Adobe = Get-AdobeAcrobatReaderDC | Where-Object {$_.Platform -eq "Windows" -and $_.Language -eq "Multi"}
$Version = $Adobe.Version
$URL = $Adobe.uri
$InstallerType = "msp"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Include *.msp, *.log, Version.txt, Download* -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source)) 
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download FSLogix
IF ($FSLogix -eq 1) {
$Product = "FSLogix"
$PackageName = "FSLogixAppsSetup"
$FSLogix = Get-MicrosoftFSLogixApps
$Version = $FSLogix.Version
$URL = $FSLogix.uri
$InstallerType = "zip"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Install\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product\Install")) {New-Item -Path "$PSScriptRoot\$Product\Install" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\Install\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\Install\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product\Install" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Install\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\Install\" + ($Source))
expand-archive -path "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.zip" -destinationpath "$PSScriptRoot\$Product\Install"
Remove-Item -Path "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.zip" -Force
Move-Item -Path "$PSScriptRoot\$Product\Install\x64\Release\*" -Destination "$PSScriptRoot\$Product\Install"
Remove-Item -Path "$PSScriptRoot\$Product\Install\Win32" -Force -Recurse
Remove-Item -Path "$PSScriptRoot\$Product\Install\x64" -Force -Recurse
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download MS Teams
IF ($MSTeams -eq 1) {
$Product = "MS Teams"
$PackageName = "Teams_windows_x64"
$Teams = Get-MicrosoftTeams | Where-Object {$_.Architecture -eq "x64"}
$Version = $Teams.Version
$URL = $Teams.uri
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Include *.msi, *.log, Version.txt, Download* -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download OneDrive
IF ($OneDrive -eq 1) {
$Product = "MS OneDrive"
$PackageName = "OneDriveSetup"
$OneDrive = Get-MicrosoftOneDrive | Where-Object {$_.Ring -eq "Production"} | Sort-Object -Property Version -Descending | Select-Object -Last 1
$Version = $OneDrive.Version
$URL = $OneDrive.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
IF (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download Office Deployment Toolkit (ODT)
IF ($OfficeDT -eq 1) {
$Product = "MS Office 365"
$PackageName = "officedeploymenttool"
$URL = $(Get-ODTUri)
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$Version = $URL.Split("_") | Select-Object -Last 1
$Version = $Version -replace ".{4}$"
$CurrentVersion = Get-Content -Path "$UpdateFolder\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$UpdateFolder\$Product")) {New-Item -Path "$UpdateFolder\$Product" -ItemType Directory | Out-Null}
$LogPS = "$UpdateFolder\$Product\" + "$Product $Version.log"
Remove-Item "$UpdateFolder\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$UpdateFolder\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$UpdateFolder\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$UpdateFolder\$Product\" + ($Source))
Start-Sleep 3
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
& "$UpdateFolder\$Product\$PackageName.$InstallerType" /quiet /extract:"$UpdateFolder\$Product"
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download VMWareTools
IF ($VMWareTools -eq 1) {
$Product = "VMWare Tools"
$PackageName = "VMWareTools"
$VMWareTools = Get-VMwareTools | Where-Object {$_.Architecture -eq "x64"}
$Version = $VMWareTools.Version
$URL = $VMWareTools.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download openJDK
IF ($OpenJDK -eq 1) {
$Product = "open JDK"
$PackageName = "OpenJDK"
$OpenJDK = Get-OpenJDK | Where-Object {$_.Architecture -eq "x64" -and $_.URI -like "*msi*"} | Sort-Object -Property Version -Descending | Select-Object -First 1
$Version = $OpenJDK.Version
$URL = $OpenJDK.uri
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download OracleJava8
IF ($OracleJava8 -eq 1) {
$Product = "Oracle Java 8"
$PackageName = "Oracle Java 8"
$OracleJava8 = Get-OracleJava8 | Where-Object {$_.Architecture -eq "x64"}
$Version = $OracleJava8.Version
$URL = $OracleJava8.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download KeePass
IF ($KeepPass -eq 1) {
$Product = "KeePass"
$PackageName = "KeePass"
$KeepPass = Get-KeePass | Where-Object {$_.URI -like "*msi*"}
$Version = $KeepPass.Version
$URL = $KeepPass.uri
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product $Version" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download mRemoteNG
IF ($mRemoteNG -eq 1) {
$Product = "mRemoteNG"
$PackageName = "mRemoteNG"
$mRemoteNG = Get-mRemoteNG | Where-Object {$_.URI -like "*msi*"}
$Version = $mRemoteNG.Version
$URL = $mRemoteNG.uri
$InstallerType = "msi"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product.txt" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}


# Download Tree Size Free
IF ($TreeSizeFree -eq 1) {
$Product = "TreeSizeFree"
$PackageName = "TreeSizeFree"
$TreeSizeFree = Get-JamTreeSizeFree
$Version = $TreeSizeFree.Version
$URL = $TreeSizeFree.uri
$InstallerType = "exe"
$Source = "$PackageName" + "." + "$InstallerType"
$CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
Write-Verbose "Download $Product" -Verbose
Write-Host "Download Version: $Version"
Write-Host "Current Version: $CurrentVersion"
IF (!($CurrentVersion -eq $Version)) {
if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
$LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
Remove-Item "$PSScriptRoot\$Product\*" -Recurse
Start-Transcript $LogPS
New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
Write-Verbose "Starting Download of $Product.txt" -Verbose
Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
Write-Verbose "Stop logging" -Verbose
Stop-Transcript
Write-Output ""
}
Write-Verbose "No new version available" -Verbose
Write-Output ""
}
