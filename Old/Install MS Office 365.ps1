# *****************************************************
# Manuel Winkel (www.deyda.net) @deyda84
# Original Version by
# D. Mohrmann, S&L Firmengruppe, Twitter: @mohrpheus78
# Install Software package on your master server/client
# *****************************************************

<#
.SYNOPSIS
This script installs MS Office 356 on a MCS/PVS master server/client or wherever you want.
		
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
$Product = "MS Office 365"

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
$PackageName = "$Product" 		    # [edit] enter the display name of the software (e.g. 'Arcobat Reader' or 'Microsoft Office')

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
$Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
$Office365 = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\AppVMachineRegistryStore\9AC08E99-230B-47E8-9721-4577B7F124EA\Versions\1A8308C7-90D1-4200-B16E-646F163A08E8\Catalog\ -ErrorAction Ignore)
$Office365V = $Office365.InternalVersion
IF ($Office365V -ne $Version) {

# Uninstallation Office365
Write-Verbose "Uninstalling $Product" -Verbose
DS_WriteLog "I" "Uninstalling $Product" $LogFile
try	{
    if (!(Test-Path "$PSScriptRoot\$Product\remove.xml" -PathType leaf))
    {
        [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
        [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
        $XML.appendChild($Root)
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Remove"))
        $Node1.SetAttribute("All","True")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Display"))
        $Node1.SetAttribute("Level","None")
        $Node1.SetAttribute("AcceptEULA","TRUE")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
        $Node1.SetAttribute("Name","AUTOACTIVATE")
        $Node1.SetAttribute("Value","0")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
        $Node1.SetAttribute("Name","FORCEAPPSHUTDOWN")
        $Node1.SetAttribute("Value","TRUE")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
        $Node1.SetAttribute("Name","SharedComputerLicensing")
        $Node1.SetAttribute("Value","0")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
        $Node1.SetAttribute("Name","PinIconsToTaskbar")
        $Node1.SetAttribute("Value","FALSE")
        $XML.Save("$PSScriptRoot\$Product\remove.xml")
    }
    Start-Process "$PSScriptRoot\$Product\Setup.exe" -ArgumentList '/configure remove.xml' -NoNewWindow -Wait
    Stop-Process -Name Setup
	} catch {
DS_WriteLog "E" "Error uninstalling $Product (error: $($Error[0]))" $LogFile       
}
DS_WriteLog "-" "" $LogFile
Write-Output ""

# Download Office 365
Write-Verbose "Download $Product" -Verbose
DS_WriteLog "I" "Downloading $Product" $LogFile
try	{
    if (!(Test-Path "$PSScriptRoot\$Product\install.xml" -PathType leaf))
    {
        [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
        [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
        $XML.appendChild($Root)
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
        $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product\$Version")
        $Node1.SetAttribute("OfficeClientEdition","64")
        $Node1.SetAttribute("Channel","SemiAnnual")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Product"))
        $Node1.SetAttribute("ID","O365ProPlusRetail")
        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Language"))
        $Node2.SetAttribute("ID","MatchOS")
        $Node2.SetAttribute("Fallback","en-us")
        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("ExcludeApp"))
        $Node2.SetAttribute("ID","Teams")
        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("ExcludeApp"))
        $Node2.SetAttribute("ID","Lync")
        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("ExcludeApp"))
        $Node2.SetAttribute("ID","Groove")
        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("ExcludeApp"))
        $Node2.SetAttribute("ID","OneDrive")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Display"))
        $Node1.SetAttribute("Level","None")
        $Node1.SetAttribute("AcceptEULA","TRUE")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Logging"))
        $Node1.SetAttribute("Level","Standard")
        $Node1.SetAttribute("Path","%temp%")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
        $Node1.SetAttribute("Name","SharedComputerLicensing")
        $Node1.SetAttribute("Value","1")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
        $Node1.SetAttribute("Name","FORCEAPPSHUTDOWN")
        $Node1.SetAttribute("Value","TRUE")
        [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Updates"))
        $Node1.SetAttribute("Enabled","FALSE")
        $XML.Save("$PSScriptRoot\$Product\install.xml")
    }
    Start-Process "$PSScriptRoot\$Product\setup.exe" -ArgumentList '/download install.xml' -NoNewWindow -Wait
    Stop-Process -Name Setup
	} catch {
DS_WriteLog "E" "Error downloading $Product (error: $($Error[0]))" $LogFile       
}
DS_WriteLog "-" "" $LogFile
Write-Output ""


# Installation Office 365
Write-Verbose "Installing $Product" -Verbose
DS_WriteLog "I" "Installing $Product" $LogFile
try	{
    Start-Process "$PSScriptRoot\$Product\setup.exe" -ArgumentList '/configure install.xml' -NoNewWindow -Wait
    Stop-Process -Name Setup
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