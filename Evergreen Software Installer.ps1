# ******************************************************
# D. Mohrmann, S&L Firmengruppe, Twitter: @mohrpheus78
# Install Software packages on your master server/client
# ******************************************************
# Update Version: Manuel Winkel (www.deyda.net)
# Addition of Office 365 Installation
<#
.SYNOPSIS
This script calls other scripts to install software on a MCS/PVS master server/client or wherever you want. Install scripts have to be in the root folder. 
		
.Description
To install a software package just switch from 0 to 1 in the section "Select software to install"

.EXAMPLE
$NotePadPlusPlus = 1 installs Notepad++

.NOTES
There are no install scripts for VMWare Tools, openJDK and Oracle Java 8 yet!
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

# Select software to install
$NotePadPlusPlus = 0
$GoogleChrome = 0
$MSEdge = 1
$VLCPlayer = 1
$BISF = 0
$FSLogix = 1
$WorkspaceApp_Current_Release = 0
$WorkspaceApp_LTSR_Release = 0
$7ZIP = 0
$AdobeReaderDCUpdate = 0
$MSTeams = 1
$OneDrive = 1
$KeepPass = 0
$mRemoteNG = 0
$TreeSizeFree = 0
$Office365 = 1


# Install Notepad ++
IF ($NotePadPlusPlus -eq 1)
	{
		& "$psscriptroot\Install NotepadPlusPlus.ps1"
	}

# Install Chrome
IF ($GoogleChrome -eq 1)
	{
		& "$psscriptroot\Install Google Chrome.ps1"
	}

# Install MS Edge
IF ($MSEdge -eq 1)
	{
		& "$psscriptroot\Install MS Edge.ps1"
	}


# Install VLC Player
IF ($VLCPlayer -eq 1)
	{
		& "$psscriptroot\Install VLC Player.ps1"
	}


# Install BIS-F
IF ($BISF -eq 1)
	{
		& "$psscriptroot\Install BIS-F.ps1"
	}

# Install FSLogix
IF ($FSLogix -eq 1)
{
	& "$psscriptroot\Install FSLogix.ps1"
}

# Install WorkspaceApp Current
IF ($WorkspaceApp_Current_Release -eq 1)
	{
		& "$psscriptroot\Install WorkspaceApp Current.ps1"
	}

# Install WorkspaceApp LTSR
IF ($WorkspaceApp_LTSR_Release -eq 1)
	{
		& "$psscriptroot\Install WorkspaceApp LTSR.ps1"
	}

# Install 7-ZIP
IF ($7ZIP -eq 1)
	{
		& "$psscriptroot\Install 7-Zip.ps1"
	}
	
# Install Adobe Reader DC MUI Update
IF ($AdobeReaderDCUpdate -eq 1)
	{
		& "$psscriptroot\Install Adobe Reader DC Update.ps1"
	}

# Install MS Teams
IF ($MSTeams -eq 1)
	{
		& "$psscriptroot\Install MS Teams.ps1"
	}
	
# Install MS OneDrive
IF ($OneDrive -eq 1)
	{
		& "$psscriptroot\Install MS OneDrive.ps1"
	}

# Install KeePass
IF ($KeepPass -eq 1)
	{
		& "$psscriptroot\Install KeePass.ps1"
	}

# Install mRemoteNG
IF ($mRemoteNG -eq 1)
	{
		& "$psscriptroot\Install mRemoteNG.ps1"
	}
	
# Install TreeSizeFree
IF ($TreeSizeFree -eq 1)
	{
		& "$psscriptroot\Install TreeSizeFree.ps1"
	}

# Install Office 365
IF ($Office365 -eq 1)
{
	& "$psscriptroot\Install MS Office 365.ps1"
}