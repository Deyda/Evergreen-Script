# *****************************************************
# D. Mohrmann, S&L Firmengruppe, Twitter: @mohrpheus78
# Install Software package on your master server/client
# *****************************************************

<#
.SYNOPSIS
This script installs the Citrix WorkspaceApp LTSR on a MCS/PVS master server/client or wherever you want.
		
.Description
Use the Software Updater script first, to check if a new version is available! After that use the Software Installer script. If you select this software
package it gets installed. 
The script compares the software version and will install or update the software. A log file will be created in the 'Install Logs' folder. 

.EXAMPLE

.NOTES
Always call this script with the Software Installer script!
#>


# define Error handling
# note: do not change these values
$global:ErrorActionPreference = "Stop"
if($verbose){ $global:VerbosePreference = "Continue" }

# Variables
$Product = "WorkspaceApp LTSR"

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
$BaseLogDir = "$PSScriptRoot\_Install Logs"       # [edit] add the location of your log directory here
$PackageName = "Citrix WorkspaceApp LTSR" 		    # [edit] enter the display name of the software (e.g. 'Arcobat Reader' or 'Microsoft Office')

# Global variables
#$StartDir = $PSScriptRoot # the directory path of the script currently being executed
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
	if($inst -ne $null)
	{
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