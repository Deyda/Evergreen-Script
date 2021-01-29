#requires -version 3
<#
.SYNOPSIS
Download and Install several Software with the Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. 
.DESCRIPTION
To update or download a software package just switch from 0 to 1 in the section "Select software".
A new folder for every single package will be created, together with a version file, a download date file and a log file. If a new version is available
the script checks the version number and will update the package.
.NOTES
  Version:        1.0
  Author:         Manuel Winkel <www.deyda.net>
  Creation Date:  2021-01-29
  Purpose/Change:
  2021-01-29    Initial Version
<#


.PARAMETER download

Only download the software packages.

.PARAMETER install

Only install the software packages.

.PARAMETER gui

Start a GUI to select the Software Packages.

.EXAMPLE

& '.\FSLogix-DiffDiskToUniqueDisk.ps1 -path D:\CTXFslogix -tmp D:\TMP -target D:\FSLogixCTX

Copy and rename the disks in the specified locations and in all child items from Path D:\CTXFSLogix to D:\FSLogixCTX, with temporary storage in D:\TMP and create 1 session disk.

.EXAMPLE

& '.\FSLogix-DiffDiskToUniqueDisk.ps1 -path D:\CTXFslogix -count 9 -delete

Copy and rename the disks in the specified locations and in all child items from Path D:\CTXFSLogix to D:\CTXFSLogix, and create 9 session disk. After that the original Difference Container are deleted
#>

[CmdletBinding()]


Param (
    
        [Parameter(
            HelpMessage='Only Download Software?',
            ValuefromPipelineByPropertyName = $true
        )]
        [switch]$download,

        [Parameter(
            HelpMessage='Only Install Software?',
            ValuefromPipelineByPropertyName = $true
        )]
        [switch]$install,
    
        [Parameter(
            HelpMessage='Start the Gui to select the Software',
            ValuefromPipelineByPropertyName = $true
        )]
        [switch]$gui
    
    )

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

# Select software
$7ZIP = 1
$AdobeProDC = 1 #Only Download @ the moment
$AdobeReaderDC = 1
$BISF = 1
$FSLogix = 1
$GoogleChrome = 1
$KeepPass = 1
$mRemoteNG = 1
$MS365Apps = 1 # Office Deployment Toolkit for installing Office 365 / Only Download @ the moment
$MSEdge = 1
$MSOffice2019 = 1 # Deployment Toolkit for installing Office 2019 / Only Download @ the moment
$MSTeams = 1
$NotePadPlusPlus = 1
$OneDrive = 1
$OpenJDK = 1 #Only Download @ the moment
$OracleJava8 = 1 #Only Download @ the moment
$TreeSizeFree = 1
$VLCPlayer = 1
$VMWareTools = 1 #Only Download @ the moment
$WinSCP = 1
$WorkspaceApp_Current_Relase = 1
$WorkspaceApp_LTSR_Relase = 1

# Disable progress bar while downloading
$ProgressPreference = 'SilentlyContinue'

# Install/Update Evergreen module
Write-Verbose "Installing/updating Evergreen module... please wait" -Verbose
Write-Output ""
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies}
if (!(Get-Module -ListAvailable -Name Evergreen)) {Install-Module Evergreen -Force | Import-Module Evergreen}
Update-Module Evergreen -force

IF ($install -eq $False) {

    Write-Output "Starting downloads..."
    Write-Output ""

    # Download 7-ZIP
    IF ($7ZIP -eq 1) {
        $Product = "7-Zip"
        $PackageName = "7-Zip_x64"
        $7Zip = Get-7zip | Where-Object { $_.Architecture -eq "x64" -and $_.URI -like "*exe*" }
        $Version = $7Zip.Version
        $URL = $7Zip.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download Adobe Pro DC Update
    IF ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        $PackageName = "Adobe_Pro_DC_Update"
        $Adobe = Get-AdobeAcrobatProDC | Where-Object { $_.Type -eq "Updater" }
        $Version = $Adobe.Version
        $URL = $Adobe.uri
        $InstallerType = "msp"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Include *.msp, *.log, Version.txt, Download* -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source)) 
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download Adobe Reader DC
    IF ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        $PackageName = "Adobe_DC_Update"
        $Adobe = Get-AdobeAcrobatReaderDC | Where-Object {$_.Type -eq "Updater" -and $_.Language -eq "Multi"}
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
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download BIS-F
    IF ($BISF -eq 1) {
        $Product = "BIS-F"
        $PackageName = "setup-BIS-F"
        $BISF = Get-BISF | Where-Object { $_.URI -like "*msi*" }
        $Version = $BISF.Version
        $URL = $BISF.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Exclude *.ps1, *.lnk -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
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
            if (!(Test-Path -Path "$PSScriptRoot\$Product\Install")) { New-Item -Path "$PSScriptRoot\$Product\Install" -ItemType Directory | Out-Null }
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
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download Google Chrome
    IF ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        $ChromeURL = Get-GoogleChrome | Where-Object { $_.Architecture -eq "x64" } | Select-Object -ExpandProperty URI
        $Version = (Get-GoogleChrome | Where-Object { $_.Architecture -eq "x64" }).Version
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $ChromeURL -OutFile ("$PSScriptRoot\$Product\" + ($ChromeURL | Split-Path -Leaf))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download KeePass
    IF ($KeepPass -eq 1) {
        $Product = "KeePass"
        $PackageName = "KeePass"
        $KeepPass = Get-KeePass | Where-Object { $_.URI -like "*msi*" }
        $Version = $KeepPass.Version
        $URL = $KeepPass.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download mRemoteNG
    IF ($mRemoteNG -eq 1) {
        $Product = "mRemoteNG"
        $PackageName = "mRemoteNG"
        $mRemoteNG = Get-mRemoteNG | Where-Object { $_.URI -like "*msi*" }
        $Version = $mRemoteNG.Version
        $URL = $mRemoteNG.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download MS Office365Apps
    IF ($MS365Apps -eq 1) {
        $Product = "MS 365 Apps (Semi Annual Channel)"
        $PackageName = "setup"
        $MS365Apps = Get-Microsoft365Apps | Where-Object {$_.Channel -eq "Semi-Annual Channel"}
        $Version = $MS365Apps.Version
        $URL = $MS365Apps.uri
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
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download MS Edge
    IF ($MSEdge -eq 1) {
        $Product = "MS Edge"
        $EdgeURL = Get-MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "x64" }
        $EdgeURL = $EdgeURL | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = (Get-MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "x64" }).Version
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue 
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $EdgeURL.Uri -OutFile ("$PSScriptRoot\$Product\" + ($EdgeURL.URI | Split-Path -Leaf))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download MS Office 2019
    IF ($MSOffice2019 -eq 1) {
        $Product = "MS Office 2019"
        $PackageName = "setup"
        $MSOffice2019 = Get-Microsoft365Apps | Where-Object {$_.Channel -eq "Office 2019 Enterprise"}
        $Version = $MSOffice2019.Version
        $URL = $MSOffice2019.uri
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
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download MS Teams
    IF ($MSTeams -eq 1) {
        $Product = "MS Teams"
        $PackageName = "Teams_windows_x64"
        $Teams = Get-MicrosoftTeams | Where-Object { $_.Architecture -eq "x64" }
        $Version = $Teams.Version
        $URL = $Teams.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Include *.msi, *.log, Version.txt, Download* -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download Notepad ++
    IF ($NotePadPlusPlus -eq 1) {
        $Product = "NotePadPlusPlus"
        $PackageName = "NotePadPlusPlus_x64"
        $Notepad = Get-NotepadPlusPlus | Where-Object { $_.Architecture -eq "x64" -and $_.URI -match ".exe" }
        $Version = $Notepad.Version
        $URL = $Notepad.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Get-ChildItem "$PSScriptRoot\$Product\" -Exclude lang | Remove-Item -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download OneDrive
    IF ($OneDrive -eq 1) {
        $Product = "MS OneDrive"
        $PackageName = "OneDriveSetup"
        $OneDrive = Get-MicrosoftOneDrive | Where-Object { $_.Ring -eq "Production" -and $_.Type -eq "Exe" } | Sort-Object -Property Version -Descending | Select-Object -Last 1
        $Version = $OneDrive.Version
        $URL = $OneDrive.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            IF (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download openJDK
    IF ($OpenJDK -eq 1) {
        $Product = "open JDK"
        $PackageName = "OpenJDK"
        $OpenJDK = Get-OpenJDK | Where-Object { $_.Architecture -eq "x64" -and $_.URI -like "*msi*" } | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $OpenJDK.Version
        $URL = $OpenJDK.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download OracleJava8
    IF ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        $PackageName = "Oracle Java 8"
        $OracleJava8 = Get-OracleJava8 | Where-Object { $_.Architecture -eq "x64" }
        $Version = $OracleJava8.Version
        $URL = $OracleJava8.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
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
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download VLC Player
    IF ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        $PackageName = "VLC-Player"
        $VLC = Get-VideoLanVlcPlayer | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" -and $_.Type -eq "MSI" }
        $Version = $VLC.Version
        $URL = $VLC.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download VMWareTools
    IF ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"
        $PackageName = "VMWareTools"
        $VMWareTools = Get-VMwareTools | Where-Object { $_.Architecture -eq "x64" }
        $Version = $VMWareTools.Version
        $URL = $VMWareTools.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            New-Item -Path "$PSScriptRoot\$Product" -Name "Download date $Date" | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download WinSCP
IF ($WinSCP -eq 1) {
    $Product = "WinSCP"
    $PackageName = "WinSCP"
    $WinSCP = Get-WinSCP | Where-Object {$_.URI -like "*Setup*"}
    $Version = $WinSCP.Version
    $URL = $WinSCP.uri
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
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download WorkspaceApp Current
    IF ($WorkspaceApp_Current_Relase -eq 1) {
        $Product = "WorkspaceApp"
        $PackageName = "CitrixWorkspaceApp"
        $WSA = Get-CitrixWorkspaceApp | Where-Object { $_.Title -like "*Workspace*" -and "*Current*" -and $_.Platform -eq "Windows" -and $_.Title -like "*Current*" }
        $Version = $WSA.Version
        $URL = $WSA.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\Current\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product\Windows\Current")) { New-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\Current" -ItemType Directory | Out-Null }
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
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }

    # Download WorkspaceApp LTSR
    IF ($WorkspaceApp_LTSR_Relase -eq 1) {
        $Product = "WorkspaceApp"
        $PackageName = "CitrixWorkspaceApp"
        $WSA = Get-CitrixWorkspaceApp | Where-Object { $_.Title -like "*Workspace*" -and "*LTSR*" -and $_.Platform -eq "Windows" -and $_.Title -like "*LTSR*" }
        $Version = $WSA.Version
        $URL = $WSA.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product LTSR" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        IF (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR")) { New-Item -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR" -ItemType Directory | Out-Null }
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
            Write-Verbose "Download of the new version $Version finished" -Verbose
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
    }
}

IF ($download -eq $False) {

    # define Error handling
    # note: do not change these values
    $global:ErrorActionPreference = "Stop"
    if($verbose){ $global:VerbosePreference = "Continue" }

    # FUNCTION Logging
    #========================================================================================================================================
    Function DS_WriteLog {
    
        [CmdletBinding()]
        Param( 
            [Parameter(Mandatory=$true, Position = 0)][ValidateSet("I","S","W","E","-",IgnoreCase = $True)][String]$InformationType,
            [Parameter(Mandatory=$true, Position = 1)][AllowEmptyString()][String]$Text,
            [Parameter(Mandatory=$true, Position = 2)][AllowEmptyString()][String]$LogFile
        )
 
        begin {
        }
 
        process {
        $DateTime = (Get-Date -format dd-MM-yyyy) + " " + (Get-Date -format HH:mm:ss)
 
            if ( $Text -eq "" ) {
                Add-Content $LogFile -value ("") # Write an empty line
            } Else {
                Add-Content $LogFile -value ($DateTime + " " + $InformationType.ToUpper() + " - " + $Text)
            }
        }
 
        end {
        }
    }

    # Logging

    # Custom variables [edit]
    $BaseLogDir = "$PSScriptRoot\_Install Logs"      # [edit] add the location of your log directory here
    $PackageName = "$Product" 		            	# [edit] enter the display name of the software (e.g. 'Arcobat Reader' or 'Microsoft Office')

    # Global variables
    $StartDir = $PSScriptRoot # the directory path of the script currently being executed
    $LogDir = (Join-Path $BaseLogDir $PackageName)
    $LogFileName = ("$ENV:COMPUTERNAME - $PackageName.log")
    $LogFile = Join-path $LogDir $LogFileName

    # Create the log directory if it does not exist
    if (!(Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType directory | Out-Null }

    # Create new log file (overwrite existing one)
    New-Item $LogFile -ItemType "file" -force | Out-Null

    DS_WriteLog "I" "START SCRIPT - $PackageName" $LogFile
    DS_WriteLog "-" "" $LogFile

    #========================================================================================================================================

    # FUNCTION MSI Installation
    #========================================================================================================================================
    function Install-MSIFile {

        [CmdletBinding()]
        Param(
            [parameter(mandatory=$true,ValueFromPipeline=$true,ValueFromPipelinebyPropertyName=$true)]
            [ValidateNotNullorEmpty()]
            [string]$msiFile,
    
            [parameter()]
            [ValidateNotNullorEmpty()]
            [string]$targetDir
        )
        if (!(Test-Path $msiFile)){
            throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
        }
        $arguments = @(
            "/i"
            "`"$msiFile`""
            "/qn"
        )
        if ($targetDir){
            if (!(Test-Path $targetDir)){
                throw "Path to the installation directory $($targetDir) is invalid. Please check path and file name!"
            }
            $arguments += "INSTALLDIR=`"$targetDir`""
        }
        $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0){
        }
        else {
            Write-Verbose "Installer Exit Code  $($process.ExitCode) for File  $($msifile)"
        }
    }
    
    #========================================================================================================================================
    

    # Install 7-ZIP
    IF ($7ZIP -eq 1) {
        $Product = "7-Zip"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $SevenZip = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion
        IF ($SevenZip -ne $Version) {
            # 7-Zip
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\7-Zip_x64.exe" –ArgumentList /S –NoNewWindow -Wait
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install Adobe Pro DC
    IF ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion
        IF ($Adobe -ne $Version) {
            # Adobe Pro DC
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                $mspArgs = "/P `"$PSScriptRoot\$Product\Adobe_Pro_DC_Update.msp`" /quiet /qn"
                Start-Process -FilePath msiexec.exe -ArgumentList $mspArgs -Wait
                # Update Dienst und Task deaktivieren
                Stop-Service AdobeARMservice
                Set-Service AdobeARMservice -StartupType Disabled
                Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
            } catch {
                DS_WriteLog "E" "Error installinng $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install Adobe Reader DC
    IF ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion
        IF ($Adobe -ne $Version) {
            # Adobe Reader DC Update
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                $mspArgs = "/P `"$PSScriptRoot\$Product\Adobe_DC_MUI_Update.msp`" /quiet /qn"
                Start-Process -FilePath msiexec.exe -ArgumentList $mspArgs -Wait
                # Update Dienst und Task deaktivieren
                Stop-Service AdobeARMservice
                Set-Service AdobeARMservice -StartupType Disabled
                Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
            } catch {
                DS_WriteLog "E" "Error installinng $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install BIS-F
    IF ($BISF -eq 1) {
        $Product = "BIS-F"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $BISF = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Base Image*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        IF ($BISF) {$BISF = $BISF -replace ".{6}$"}
        IF ($BISF -ne $Version) {
            # Base Image Script Framework
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\setup-BIS-F.msi" | Install-MSIFile
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            write-Output ""
        }
        # Customize scripts, it's best practise to enable Task Offload and RSS and to disable DEP
        write-Verbose "Customize scripts" -Verbose
        DS_WriteLog "I" "Customize scripts" $LogFile
        try {
            ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace "DisableTaskOffload' -Value '1'","DisableTaskOffload' -Value '0'") | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
            ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'nx AlwaysOff','nx OptOut') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
            ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'rss=disable','rss=enable') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
        } catch {
            DS_WriteLog "E" "Error beim Anpassen der Skripte (error: $($Error[0]))" $LogFile
        }
        DS_WriteLog "-" "" $LogFile
        write-Output ""
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install FSLogix
    IF ($FSLogix -eq 1) {
        $Product = "FSLogix"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Install\Version.txt"
        $FSLogix = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).DisplayVersion
        IF (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}) {
            $UninstallFSL = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).UninstallString.replace("/uninstall","")
        }
        IF (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps RuleEditor"}) {
            $UninstallFSLRE = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps RuleEditor"}).UninstallString.replace("/uninstall","")
        }
        IF ($FSLogix -ne $Version) {
            # FSLogix Uninstall
            IF (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}) {
                Write-Verbose "Uninstalling $Product" -Verbose
                DS_WriteLog "I" "Uninstalling $Product" $LogFile
                try	{
                    Start-process $UninstallFSL -ArgumentList '/uninstall /quiet /norestart' –NoNewWindow -Wait
                    Start-process $UninstallFSLRE -ArgumentList '/uninstall /quiet /norestart' –NoNewWindow -Wait
                } catch {
                    DS_WriteLog "E" "Error Uninstalling $Product (error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
                Write-Verbose "Server needs to reboot, start script again after reboot" -Verbose
                Write-Output ""
                Write-Output "Hit any key to reboot server"
                Read-Host
                Restart-Computer
            }
            # FSLogix Install
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.exe" -ArgumentList '/install /norestart /quiet'  –NoNewWindow -Wait
                Start-Process "$PSScriptRoot\$Product\Install\FSLogixAppsRuleEditorSetup.exe" -ArgumentList '/install /norestart /quiet'  –NoNewWindow -Wait
                reg add "HKLM\SOFTWARE\FSLogix\Profiles" /v GroupPolicyState /t REG_DWORD /d 0 /f | Out-Null
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install Chrome
    IF ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Chrome = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion
        IF ($Chrome -ne $Version) {
            # Google Chrome
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\googlechromestandaloneenterprise64.msi" | Install-MSIFile
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install KeePass
    IF ($KeepPass -eq 1) {
        $Product = "KeePass"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $KeePass = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*KeePass*"}).DisplayVersion
        IF ($KeePass) {$KeePass = $KeePass -replace ".{2}$"}
        IF ($KeePass -ne $Version) {
            # KeePass
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\KeePass.msi" | Install-MSIFile
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install mRemoteNG
    IF ($mRemoteNG -eq 1) {
        $Product = "mRemoteNG"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $mRemoteNG = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "mRemoteNG"}).DisplayVersion
        IF ($mRemoteNG) {$mRemoteNG = $mRemoteNG -replace ".{6}$"}
        IF ($mRemoteNG -ne $Version) {
            # mRemoteNG
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\mRemoteNG.msi" | Install-MSIFile
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install MS Edge
    IF ($MSEdge -eq 1) {
        $Product = "MS Edge"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Edge = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion
        IF ($Edge -ne $Version) {
            # MS Edge
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\MicrosoftEdgeEnterpriseX64.msi" | Install-MSIFile
                # Update Task deaktivieren
                # Disable-ScheduledTask -TaskName MicrosoftEdgeUpdateTaskMachineCore | Out-Null
                # Disable-ScheduledTask -TaskName MicrosoftEdgeUpdateTaskMachineUA | Out-Null
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            # Disable Citrix API Hooks (MS Edge) on Citrix VDA
            $(
                $RegPath = "HKLM:SYSTEM\CurrentControlSet\services\CtxUvi"
                IF (Test-Path $RegPath) {
                    $RegName = "UviProcessExcludes"
                    $EdgeRegvalue = "msedge.exe"
                    # Get current values in UviProcessExcludes
                    $CurrentValues = Get-ItemProperty -Path $RegPath | Select-Object -ExpandProperty $RegName
                    # Add the msedge.exe value to existing values in UviProcessExcludes
                    Set-ItemProperty -Path $RegPath -Name $RegName -Value "$CurrentValues$EdgeRegvalue;"
                }
            ) | Out-Null
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install MS OneDrive
    IF ($OneDrive -eq 1) {
        $Product = "MS OneDrive"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $OneDrive = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion
        IF ($OneDrive -ne $Version) {
            # Installation OneDrive
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $null = Start-Process "$PSScriptRoot\$Product\OneDriveSetup.exe" –ArgumentList '/allusers' –NoNewWindow -PassThru
                while (Get-Process -Name "OneDriveSetup" -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                # onedrive starts automatically after setup. kill!
                Stop-Process -Name "OneDrive" -Force
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install MS Teams
    IF ($MSTeams -eq 1) {
        $Product = "MS Teams"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Teams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).DisplayVersion
        IF ($Teams) {$Teams = $Teams.Insert(5,'0')}
        IF ($Teams -ne $Version) {
            #Uninstalling MS Teams
            Write-Verbose "Uninstalling $Product" -Verbose
            DS_WriteLog "I" "Uninstalling $Product" $LogFile
            try {
                $UninstallTeams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).UninstallString
                $UninstallTeams = $UninstallTeams -Replace("MsiExec.exe /I","")
                Start-Process -FilePath msiexec.exe -ArgumentList "/X $UninstallTeams /qn"
                Start-Sleep 20
            } catch {
                DS_WriteLog "E" "Ein Fehler ist aufgetreten beim Deinstallieren von $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Verbose " ...ready!" -Verbose
            #MS Teams Installation
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\Teams_windows_x64.msi" | Install-MSIFile
                Start-Sleep 5
                # Prevents MS Teams from starting at logon, better do this with WEM or similar
                Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -Force
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install Notepad ++
    IF ($NotePadPlusPlus -eq 1) {
        $Product = "NotepadPlusPlus"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Notepad = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion
        IF ($Notepad -ne $Version) {
            # Installation Notepad++
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\NotePadPlusPlus_x64.exe" –ArgumentList /S –NoNewWindow -Wait
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install TreeSizeFree
    IF ($TreeSizeFree -eq 1) {
        $Product = "TreeSizeFree"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Version = $Version.Insert(3,'.')
        $TreeSize = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion
        IF ($TreeSize -ne $Version) {
            # Installation Tree Size Free
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\TreeSizeFree.exe" –ArgumentList /VerySilent –NoNewWindow -Wait
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install VLC Player
    IF ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $VLC = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion
        IF ($VLC) {$VLC = $VLC -replace ".{2}$"}
        IF ($VLC -ne $Version) {
            # VLC Player
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\VLC-Player.msi" | Install-MSIFile
            } catch {
                DS_WriteLog "E" "Ein Fehler ist aufgetreten beim Installieren von $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install WorkspaceApp Current
    IF ($WorkspaceApp_Current_Release -eq 1) {
        $Product = "WorkspaceApp Current Release"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\Citrix\WorkspaceApp\Windows\Current\Version.txt"
        $WSA = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Workspace*" -and $_.UninstallString -like "*Trolley*"}).DisplayVersion
        IF ($WSA -ne $Version) {
            # Citrix WSA Installation
            $Options = @(
                "/silent"
                "/EnableCEIP=false"
                "/FORCE_LAA=1"
                "/AutoUpdateCheck=disabled"
                "/EnableCEIP=false"
                "/ALLOWADDSTORE=S"
                "/ALLOWSAVEPWD=S"
                "/includeSSON"
                "/ENABLE_SSON=Yes"
            )
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\Citrix\WorkspaceApp\Windows\Current\CitrixWorkspaceAppWeb.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                } 
                reg add "HKLM\SOFTWARE\Wow6432Node\Policies\Citrix" /v EnableX1FTU /t REG_DWORD /d 0 /f | Out-Null
                reg add "HKCU\Software\Citrix\Splashscreen" /v SplashscrrenShown /d 1 /f | Out-Null
                reg add "HKLM\SOFTWARE\Policies\Citrix" /f /v EnableFTU /t REG_DWORD /d 0 | Out-Null
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Verbose " ... ready!" -Verbose
            Write-Verbose "Server needs to reboot after installation!" -Verbose
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    # Install WorkspaceApp LTSR
    IF ($WorkspaceApp_LTSR_Release -eq 1) {
        $Product = "WorkspaceApp LTSR"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\Citrix\WorkspaceApp\Windows\LTSR\Version.txt"
        $WSA = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Workspace*" -and $_.UninstallString -like "*Trolley*"}).DisplayVersion
        IF ($WSA -ne $Version) {
            # Citrix WSA Installation
            $Options = @(
                "/silent"
                "/EnableCEIP=false"
                "/FORCE_LAA=1"
                "/AutoUpdateCheck=disabled"
                "/EnableCEIP=false"
                "/ALLOWADDSTORE=S"
                "/ALLOWSAVEPWD=S"
                "/includeSSON"
                "/ENABLE_SSON=Yes"
            )
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\Citrix\WorkspaceApp\Windows\LTSR\CitrixWorkspaceAppWeb.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                } 
                reg add "HKLM\SOFTWARE\Wow6432Node\Policies\Citrix" /v EnableX1FTU /t REG_DWORD /d 0 /f | Out-Null
                reg add "HKCU\Software\Citrix\Splashscreen" /v SplashscrrenShown /d 1 /f | Out-Null
                reg add "HKLM\SOFTWARE\Policies\Citrix" /f /v EnableFTU /t REG_DWORD /d 0 | Out-Null
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Verbose " ... ready!" -Verbose
            Write-Verbose "Server needs to reboot after installation!" -Verbose
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }
}