#requires -version 3
<#
.SYNOPSIS
Download and Install several Software with the Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. 
.DESCRIPTION
To update or download a software package just switch from 0 to 1 in the section "Select software" (With parameter -list) or select your Software out of the GUI.
A new folder for every single package will be created, together with a version file and a log file. If a new version is available
the script checks the version number and will update the package.
.NOTES
  Version:          1.42
  Author:           Manuel Winkel <www.deyda.net>
  Creation Date:    2021-01-29
  // NOTE: Purpose/Change
  2021-01-29        Initial Version
  2021-01-30        Error solved: No installation without parameters / Add WinSCP Install
  2021-01-31        Error solved: Installation Workspace App -> Wrong Variable / Error solved: Detection acute version 7-Zip -> Limitation of the results
  2021-02-01        Add Gui Mode as Standard
  2021-02-02        Add Install OpenJDK / Add Install VMWare Tools / Add Install Oracle Java 8 / Add Install Adobe Reader DC
  2021-02-03        Addition of verbose comments. Chrome and Edge customization regarding disabling services and scheduled tasks.
  2021-02-04        Correction OracleJava8 detection / Add Environment Variable $env:evergreen for script path
  2021-02-12        Add Download Citrix Hypervisor Tools, Greenshot, Firefox, Foxit Reader & Filezilla / Correction Citrix Workspace Download & Install Folder / Adding Citrix Receiver Cleanup Utility
  2021-02-14        Change Adobe Acrobat DC Downloader
  2021-02-15        Change MS Teams Downloader / Correction GUI Select All / Add Download MS Apps 365 & Office 2019 Install Files / Add Uninstall and Install MS Apps 365 & Office 2019
  2021-02-18        Correction Code regarding location of scripts at MS365Apps and MSOffice2019. Removing Download Time Files.
  2021-02-19        Implementation of new GUI / Add choice of architecture option in 7-Zip / Add choice of language option in Adobe Reader DC / Add choice of architecture option in Citrix Hypervisor Tools / Add choice of release option in Citrix Workspace App (Merge LTSR and CR script part)
  2021-02-22        Add choice of architecture, language and channel (Latest and ESR) options in Mozilla Firefox / Add choice of language option in Foxit Reader / Add choice of architecture option in Google Chrome / Add choice of channel, architecture and language options in Microsoft 365 Apps / Add choice of architecture option in Microsoft Edge / Add choice of architecture and language options in Microsoft Office 2019 / Add choice of update ring option in Microsoft OneDrive
  2021-02-23        Correction Microsoft Edge Download / Google Chrome Version File
  2021-02-25        Set Mark Jump markers for better editing / Add choice of architecture and update ring options in Microsoft Teams / Add choice of architecture option in Notepad++ / Add choice of architecture option in openJDK / Add choice of architecture option in Oracle Java 8
  2021-02-26        Add choice of version type option in TreeSize / Add choice of version type option in VLC-Player / Add choice of version type option in VMWare Tools / Fix installed version detection for x86 / x64 for Microsoft Edge, Google Chrome, 7-Zip, Citrix Hypervisor Tools, Mozilla Firefox, Microsoft365, Microsoft Teams, Microsoft Edge, Notepad++, openJDK, Oracle Java 8, VLC Player and VMWare Tols/ Correction Foxit Reader gui variable / Correction version.txt for Microsoft Teams, Notepad++, openJDK, Oracle Java 8, VLC Player and VMWare Tools
  2021-02-28        Implementation of LastSetting memory
  2021-03-02        Add Microsoft Teams Citrix Api Hook / Correction En dash Error
  2021-03-05        Adjustment regarding merge #122 (Get-AdobeAcrobatReader)
  2021-03-10        Fix Citrix Workspace App File / Adding advanced logging for Microsoft Teams installation
  2021-03-13        Adding advanced logging for BIS-F, Citrix Hypervisor Tools, Google Chrome, KeePass, Microsoft Edge, Mozilla Firefox, mRemoteNG, Open JDK and VLC Player installation / Adobe Reader Registry Filter Customization / New install parameter Foxit Reader
  2021-03-14        New Install Parameter Adobe Reader DC, Mozilla Firefox and Oracle Java 8 / GUI new Logo Location
  2021-03-15        New Install Parameter Microsoft Edge and Microsoft Teams / Post Setup Customization FSLogix, Microsoft Teams and Microsoft FSLogix
  2021-03-16        Fix Silent Installation of Foxit Reader / Delete Public Desktop Icon of Microsoft Teams, VLC Player and Foxit Reader / Add IrfanView in GUI / Add IrfanView Install and Download / Add Microsoft Teams Developer Ring
  2021-03-22        Add Comments / Add (AddScript) to find the places faster when new application is added / Change Install Logging function / Change Adobe Pro DC Download request
  2021-03-23        Added the possibility to delete Microsoft Teams AutoStart in the GUI / Change Microsoft Edge service to manual
  2021-03-24        Add Download Microsoft PowerShell, Microsoft .Net, RemoteDesktopManager, deviceTRUST and Zoom
  2021-03-25        Add Download Slack and ShareX / Add new Software to GUI
  2021-03-26        Add Pending Reboot Check / Add Install RemoteDesktopManager / Icon Delete Public Desktop for KeePass, mRemoteNG, WinSCP and VLC Player
  2021-03-29        Correction Microsoft FSLogix registry entries / Correction Microsoft OneDrive Installer / Add Install Microsoft .Net Framework, ShareX, Slack and Microsoft PowerShell / Correction Zoom and deviceTRUST Download
  2021-03-30        Add Install Zoom + Zoom Plugin for Citrix Receiver and deviceTRUST (Client, Host and Console)
  2021-04-06        Change to new Evergreen Commands
  2021-04-07        Change to faster download method
  2021-04-08        Change color scheme of the messages in Download section / New central MSI Install Function
  2021-04-09        Change color scheme of the messages in Install section
  2021-04-11        Implement new MSI Install Function
  2021-04-12        Correction eng dash
  2021-04-13        Change encoding to UTF-8withBOM / Correction displayed Current Version Install Adobe Reader DC
  2021-04-15        Add Microsoft Edge Dev and Beta Channel / Add Microsoft OneDrive ADM64
  2021-04-16        Script cleanup using the PSScriptAnalyzer suggestions / Add new version check with auto download
  2021-04-21        Customize Auto Update (TLS12 Error) / Teams AutoStart Kill registry query / Correction Teams Outlook Addin registration

  .PARAMETER list

Don't start the GUI to select the Software Packages and use the hardcoded list in the script.

.PARAMETER download

Only download the software packages in list Mode (-list).

.PARAMETER install

Only install the software packages in list Mode (-list).

.EXAMPLE

.\Evergreen.ps1 -list -download

Download the selected Software out of the list.

.EXAMPLE

.\Evergreen.ps1 -list -install

Install the selected Software out of the list.

.EXAMPLE

.\Evergreen.ps1 -list

Download and install the selected Software out of the list.

.EXAMPLE

.\Evergreen.ps1

Start the GUI to select the mode (Install and/or Download) and the Software.
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
        [switch]$list
    
)

#Add Functions here

# Function MSI Installation
#========================================================================================================================================
Function Install-MSI {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        $msiFile, 
        $Arguments
    )
    If (!(Test-Path $msiFile)) {
        Write-Host -ForegroundColor Red "Path to MSI file ($msiFile) is invalid. Please check name and path"
    }
    $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $Arguments -NoNewWindow -PassThru
    If ($inst) {
            Wait-Process -InputObject $inst
    }
    If ($process.ExitCode -eq 0) {
        Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
        DS_WriteLog "I" "Installation $Product finished!" $LogFile
    }
    Else {
        Write-Host -ForegroundColor Red "Error installing $Product (Exit Code $($process.ExitCode) for file $($msifile))"
    }
}

# Function File Download with Progress Bar
#========================================================================================================================================
Function Get-Download {
    Param (
        [Parameter(Mandatory=$true)]
        $url, 
        $destinationFolder="$PSScriptRoot\$Product\",
        $file="$Source",
        [switch]$includeStats
    )
    $wc = New-Object Net.WebClient
    $wc.UseDefaultCredentials = $true
    $destination = Join-Path $destinationFolder $file
    $start = Get-Date
    $wc.DownloadFile($url, $destination)
    $elapsed = ((Get-Date) - $start).ToString('hh\:mm\:ss')
    $totalSize = (Get-Item $destination).Length | Get-FileSize
    If ($includeStats.IsPresent){
        $DownloadStat = [PSCustomObject]@{TotalSize=$totalSize;Time=$elapsed}
        Write-Information $DownloadStat
    }
    Get-Item $destination | Unblock-File
}
Filter Get-FileSize {
	"{0:N2} {1}" -f $(
	If ($_ -lt 1kb) { $_, 'Bytes' }
	ElseIf ($_ -lt 1mb) { ($_/1kb), 'KB' }
	ElseIf ($_ -lt 1gb) { ($_/1mb), 'MB' }
	ElseIf ($_ -lt 1tb) { ($_/1gb), 'GB' }
	ElseIf ($_ -lt 1pb) { ($_/1tb), 'TB' }
	Else { ($_/1pb), 'PB' }
	)
}

# Function Microsoft Teams Download Developer Version
#========================================================================================================================================
Function Get-MicrosoftTeamsDev() {
    <#
    .NOTES
    Author: Jonathan Pitre
    Twitter: @PitreJonathan
    #>
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://whatpulse.org/app/microsoft-teams#versions"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "\<td\>\d.\d.\d{2}.\d+<\/td\>\n.+windows"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersion = $webVersion.Split()[0].Trim("</td>")
        $appx64URL = "https://statics.teams.cdn.office.net/production-windows-x64/$appVersion/Teams_windows_x64.msi"
        $appx86URL = "https://statics.teams.cdn.office.net/production-windows-x86/$appVersion/Teams_windows_x86.msi"

        $PSObjectx86 = [PSCustomObject] @{
            Version      = $appVersion
            Ring         = "Developer"
            Architecture = "x86"
            URI          = $appx86URL
        }

        $PSObjectx64 = [PSCustomObject] @{
            Version      = $appVersion
            Ring         = "Developer"
            Architecture = "x64"
            URI          = $appx64URL
        }
        Write-Output -InputObject $PSObjectx86
        Write-Output -InputObject $PSObjectx64
    }
}

# Function Test RegistryValue Pending Reboot
#========================================================================================================================================
Function Test-RegistryValue {
    Param (
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]$Value
    )
    Try {
        Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        Return $true
    }
    Catch {
        Return $false
    }
}

# Function Test RegistryValue
#========================================================================================================================================
Function Test-RegistryValue2 {
    Param (
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]$Value
    )
    Try {
        Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction SilentlyContinue | Out-Null
        Return $true
    }
    Catch {
        Return $false
    }
}

# Function Logging
#========================================================================================================================================
Function DS_WriteLog {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, Position = 0)][ValidateSet("I","S","W","E","-",IgnoreCase = $True)][String]$InformationType,
        [Parameter(Mandatory=$true, Position = 1)][AllowEmptyString()][String]$Text,
        [Parameter(Mandatory=$true, Position = 2)][AllowEmptyString()][String]$LogFile
    )
    Begin {
    }
    Process {
        $DateTime = (Get-Date -format dd-MM-yyyy) + " " + (Get-Date -format HH:mm:ss)
        If ( $Text -eq "" ) {
            Add-Content $LogFile -value ("") # Write an empty line
        } Else {
            Add-Content $LogFile -value ($DateTime + " " + $InformationType.ToUpper() + " - " + $Text)
        }
    }
    End {
    }
}

# Disable progress bar while downloading
$ProgressPreference = 'SilentlyContinue'

# Is there a newer Evergreen Script version?
# ========================================================================================================================================
$eVersion = "1.42"
[bool]$NewerVersion = $false
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$WebResponseVersion = Invoke-WebRequest "https://raw.githubusercontent.com/Deyda/Evergreen-Script/main/Evergreen.ps1"
$WebVersion = (($WebResponseVersion.tostring() -split "[`r`n]" | select-string "Version:" | Select-Object -First 1) -split ":")[1].Trim()
If ($WebVersion -gt $eVersion) {
    $NewerVersion = $true
}

# Do you run the script as admin?
# ========================================================================================================================================
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

# Is there a pending reboot?
# ========================================================================================================================================
[bool]$PendingReboot = $false
#Check for Keys
If ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -eq $true) {
    $PendingReboot = $true
}
If ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting") -eq $true) {
    $PendingReboot = $true
}
If ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -eq $true) {
    $PendingReboot = $true
}
If ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") -eq $true) {
    $PendingReboot = $true
}
If ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts") -eq $true) {
    $PendingReboot = $true
}
#Check for Values
If ((Test-RegistryValue -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing" -Value "RebootInProgress") -eq $true) {
    $PendingReboot = $true
}
If ((Test-RegistryValue -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing" -Value "PackagesPending") -eq $true) {
    $PendingReboot = $true
}
If ((Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Value "PendingFileRenameOperations") -eq $true) {
    $PendingReboot = $true
}
If ((Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Value "PendingFileRenameOperations2") -eq $true) {
    $PendingReboot = $true
}
<#If ((Test-RegistryValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Value "DVDRebootSignal") -eq $true) {
    $PendingReboot = $true
}#>
If ((Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon" -Value "JoinDomain") -eq $true) {
    $PendingReboot = $true
}
If ((Test-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon" -Value "AvoidSpnSet") -eq $true) {
    $PendingReboot = $true
}

# Script Version
# ========================================================================================================================================
Write-Output ""
Write-Host -BackgroundColor DarkGreen -ForegroundColor Yellow "   Evergreen Script - Update your Software, the lazy way   "
Write-Host -BackgroundColor DarkGreen -ForegroundColor Yellow "               Manuel Winkel (www.deyda.net)               "
Write-Host -BackgroundColor DarkGreen -ForegroundColor Yellow "                     Version $eVersion                           "
$host.ui.RawUI.WindowTitle ="Evergreen Script - Update your Software, the lazy way - Manuel Winkel (www.deyda.net) - Version $eVersion"
If (Test-Path "$PSScriptRoot\update.ps1" -PathType leaf) {
    Remove-Item -Path "$PSScriptRoot\Update.ps1" -Force
}
Write-Output ""
Write-Host -Foregroundcolor DarkGray "Is there a newer Evergreen Script version?"
If ($NewerVersion -eq $false) {
    # No new version available
    Write-Host -Foregroundcolor Green "OK, script is newest version!"
    Write-Output ""
}
Else {
    # There is a new Evergreen Script Version
    Write-Host -Foregroundcolor Red "Attention! There is a new version of the Evergreen Script."
    Write-Output ""
    $wshell = New-Object -ComObject Wscript.Shell
    $AnswerPending = $wshell.Popup("Do you want to download the new version?",0,"New Version Alert!",32+4)
    If ($AnswerPending -eq "6") {
        Start-Process "https://www.deyda.net/index.php/en/evergreen-script/"
        $update = @'
            Remove-Item -Path "$PSScriptRoot\Evergreen.ps1" -Force 
            Invoke-WebRequest -Uri https://raw.githubusercontent.com/Deyda/Evergreen-Script/main/Evergreen.ps1 -OutFile ("$PSScriptRoot\" + "Evergreen.ps1")
            & "$PSScriptRoot\evergreen.ps1"
'@
        $update > $PSScriptRoot\update.ps1
        & "$PSScriptRoot\update.ps1"
        Break
    }
}

Write-Host -Foregroundcolor DarkGray "Does the script run under admin rights?"
If ($myWindowsPrincipal.IsInRole($adminRole)) {
    # OK, runs as admin
    Write-Host -Foregroundcolor Green "OK, script is running with admin rights."
    Write-Output ""
}
Else {
    # Script doesn't run as admin, stop!
    Write-Host -Foregroundcolor Red "Error! Script is NOT running with admin rights!"
    Break
}

Write-Host -Foregroundcolor DarkGray "Are there still pending reboots?"
If ($PendingReboot -eq $false) {
    # OK, no pending reboot
    Write-Host -Foregroundcolor Green "OK, no pending reboot"
    Write-Output ""
}
Else {
    # Oh Oh pending reboot, stop the script and reboot!
    Write-Host -Foregroundcolor Red "Error! Pending reboot! Reboot System!"
    Write-Output ""
    $wshell = New-Object -ComObject Wscript.Shell
    $AnswerPending = $wshell.Popup("Do you want to restart?",0,"Pending reboot alert!",32+4)
    If ($AnswerPending -eq "6") {
        Restart-Computer -Force
    }
    #Break
}

# Function GUI
# ========================================================================================================================================

Function gui_mode {
#// MARK: XAML Code (AddScript)
$inputXML = @"
<Window x:Class="GUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:GUI"
        mc:Ignorable="d"
        Title="Evergreen Script - Update your Software, the lazy way - Version $eVersion" Height="518" Width="855">
    <Grid x:Name="Evergreen_GUI" Margin="0,0,0,0" VerticalAlignment="Stretch">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="13*"/>
            <ColumnDefinition Width="234*"/>
            <ColumnDefinition Width="586*"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="Image_Logo" Height="100" Margin="472,0,24,0" VerticalAlignment="Top" Width="100" Source="$PSScriptRoot\img\Logo_DEYDA_no_cta.png" Grid.Column="2" ToolTip="www.deyda.net"/>
        <Button x:Name="Button_Start" Content="Start" HorizontalAlignment="Left" Margin="258,421,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
        <Button x:Name="Button_Cancel" Content="Cancel" HorizontalAlignment="Left" Margin="353,421,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
        <Label x:Name="Label_SelectMode" Content="Select Mode" HorizontalAlignment="Left" Margin="15,3,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Download" Content="Download" HorizontalAlignment="Left" Margin="15,34,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Install" Content="Install" HorizontalAlignment="Left" Margin="103,34,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <Label x:Name="Label_SelectLanguage" Content="Select Language" HorizontalAlignment="Left" Margin="131,3,0,0" VerticalAlignment="Top" Grid.Column="2"/>
        <ComboBox x:Name="Box_Language" HorizontalAlignment="Left" Margin="147,30,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="2" ToolTip="If this is selectable at download!">
            <ListBoxItem Content="Danish"/>
            <ListBoxItem Content="Dutch"/>
            <ListBoxItem Content="English"/>
            <ListBoxItem Content="Finnish"/>
            <ListBoxItem Content="French"/>
            <ListBoxItem Content="German"/>
            <ListBoxItem Content="Italian"/>
            <ListBoxItem Content="Japanese"/>
            <ListBoxItem Content="Korean"/>
            <ListBoxItem Content="Norwegian"/>
            <ListBoxItem Content="Polish"/>
            <ListBoxItem Content="Portuguese"/>
            <ListBoxItem Content="Russian"/>
            <ListBoxItem Content="Spanish"/>
            <ListBoxItem Content="Swedish"/>
        </ComboBox>
        <Label x:Name="Label_SelectArchitecture" Content="Select Architecture" HorizontalAlignment="Left" Margin="286,3,0,0" VerticalAlignment="Top" Grid.Column="2"/>
        <ComboBox x:Name="Box_Architecture" HorizontalAlignment="Left" Margin="320,30,0,0" VerticalAlignment="Top" SelectedIndex="0" RenderTransformOrigin="0.864,0.591" Grid.Column="2" ToolTip="If this is selectable at download!">
            <ListBoxItem Content="x64"/>
            <ListBoxItem Content="x86"/>
        </ComboBox>
        <Label x:Name="Label_Explanation" Content="When software download can be filtered on language or architecture." HorizontalAlignment="Left" Margin="100,52,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2"/>
        <Label x:Name="Label_Software" Content="Select Software" HorizontalAlignment="Left" Margin="15,51,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_7Zip" Content="7 Zip" HorizontalAlignment="Left" Margin="15,82,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_AdobeProDC" Content="Adobe Pro DC" HorizontalAlignment="Left" Margin="15,102,0,0" VerticalAlignment="Top" Grid.Column="1" ToolTip="Update Only!"/>
        <CheckBox x:Name="Checkbox_AdobeReaderDC" Content="Adobe Reader DC" HorizontalAlignment="Left" Margin="15,122,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_BISF" Content="BIS-F" HorizontalAlignment="Left" Margin="15,142,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_CitrixHypervisorTools" Content="Citrix Hypervisor Tools" HorizontalAlignment="Left" Margin="15,162,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_CitrixWorkspaceApp" Content="Citrix Workspace App" HorizontalAlignment="Left" Margin="15,182,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <ComboBox x:Name="Box_CitrixWorkspaceApp" HorizontalAlignment="Left" Margin="179,177,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.ColumnSpan="2" Grid.Column="1">
            <ListBoxItem Content="Current Release"/>
            <ListBoxItem Content="Long Term Service Release"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_Filezilla" Content="Filezilla" HorizontalAlignment="Left" Margin="15,222,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_FoxitReader" Content="Foxit Reader" HorizontalAlignment="Left" Margin="15,242,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_GoogleChrome" Content="Google Chrome" HorizontalAlignment="Left" Margin="15,262,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Greenshot" Content="Greenshot" HorizontalAlignment="Left" Margin="15,282,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_KeePass" Content="KeePass" HorizontalAlignment="Left" Margin="15,322,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_mRemoteNG" Content="mRemoteNG" HorizontalAlignment="Left" Margin="15,342,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSEdge" Content="Microsoft Edge" HorizontalAlignment="Left" Margin="15,402,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSFSlogix" Content="Microsoft FSLogix" HorizontalAlignment="Left" Margin="15,422,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSOffice2019" Content="Microsoft Office 2019" HorizontalAlignment="Left" Margin="153,82,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_MSOneDrive" Content="Microsoft OneDrive" HorizontalAlignment="Left" Margin="153,102,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2" ToolTip="Machine Based Install"/>
        <ComboBox x:Name="Box_MSOneDrive" HorizontalAlignment="Left" Margin="320,96,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="2" ToolTip="Machine Based Install">
            <ListBoxItem Content="Insider Ring"/>
            <ListBoxItem Content="Production Ring"/>
            <ListBoxItem Content="Enterprise Ring"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_MSTeams" Content="Microsoft Teams" HorizontalAlignment="Left" Margin="153,142,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2" ToolTip="Machine Based Install"/>
        <ComboBox x:Name="Box_MSTeams" HorizontalAlignment="Left" Margin="320,138,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="2" ToolTip="Machine Based Install">
            <ListBoxItem Content="Developer Ring"/>
            <ListBoxItem Content="Preview Ring"/>
            <ListBoxItem Content="General Ring"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_Firefox" Content="Mozilla Firefox" HorizontalAlignment="Left" Margin="153,162,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <ComboBox x:Name="Box_Firefox" HorizontalAlignment="Left" Margin="320,159,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="Current"/>
            <ListBoxItem Content="ESR"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_NotepadPlusPlus" Content="Notepad ++" HorizontalAlignment="Left" Margin="153,182,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_OpenJDK" Content="Open JDK" HorizontalAlignment="Left" Margin="153,202,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_OracleJava8" Content="Oracle Java 8" HorizontalAlignment="Left" Margin="153,222,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_TreeSize" Content="TreeSize" HorizontalAlignment="Left" Margin="153,302,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <ComboBox x:Name="Box_TreeSize" HorizontalAlignment="Left" Margin="320,299,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="Free"/>
            <ListBoxItem Content="Professional"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_VLCPlayer" Content="VLC Player" HorizontalAlignment="Left" Margin="153,322,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_VMWareTools" Content="VMWare Tools" HorizontalAlignment="Left" Margin="153,342,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_WinSCP" Content="WinSCP" HorizontalAlignment="Left" Margin="153,362,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_SelectAll" Content="Select All" HorizontalAlignment="Left" Margin="127,425,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <Label x:Name="Label_author" Content="Manuel Winkel / @deyda84 / www.deyda.net / 2021 / Version $eVersion" HorizontalAlignment="Left" Margin="286,453,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_MS365Apps" Content="Microsoft 365 Apps" HorizontalAlignment="Left" Margin="15,382,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <ComboBox x:Name="Box_MS365Apps" HorizontalAlignment="Left" Margin="179,379,0,0" VerticalAlignment="Top" SelectedIndex="4" Grid.Column="1" Grid.ColumnSpan="2">
            <ListBoxItem Content="Current (Preview)"/>
            <ListBoxItem Content="Current"/>
            <ListBoxItem Content="Monthly Enterprise"/>
            <ListBoxItem Content="Semi-Annual Enterprise (Preview)"/>
            <ListBoxItem Content="Semi-Annual Enterprise"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_IrfanView" Content="IrfanView" HorizontalAlignment="Left" Margin="15,302,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSTeams_No_AutoStart" Content="No AutoStart" HorizontalAlignment="Left" Margin="438,142,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2" ToolTip="Delete the HKLM Run entry to AutoStart Microsoft Teams"/>
        <CheckBox x:Name="Checkbox_deviceTRUST" Content="deviceTRUST" HorizontalAlignment="Left" Margin="15,202,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_MSDotNetFramework" Content="Microsoft .Net Framework" HorizontalAlignment="Left" Margin="15,362,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_MSPowerShell" Content="Microsoft PowerShell" HorizontalAlignment="Left" Margin="153,122,0,0" VerticalAlignment="Top" Grid.Column="2" />
        <CheckBox x:Name="Checkbox_RemoteDesktopManager" Content="Remote Desktop Manager" HorizontalAlignment="Left" Margin="153,242,0,0" VerticalAlignment="Top" Grid.Column="2" />
        <CheckBox x:Name="Checkbox_ShareX" Content="ShareX" HorizontalAlignment="Left" Margin="153,262,0,0" VerticalAlignment="Top" Grid.Column="2" />
        <CheckBox x:Name="Checkbox_Slack" Content="Slack" HorizontalAlignment="Left" Margin="153,282,0,0" VerticalAlignment="Top" Grid.Column="2" />
        <CheckBox x:Name="Checkbox_Zoom" Content="Zoom" HorizontalAlignment="Left" Margin="153,382,0,0" VerticalAlignment="Top" Grid.Column="2" />
        <ComboBox x:Name="Box_MSDotNetFramework" HorizontalAlignment="Left" Margin="179,358,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
            <ListBoxItem Content="Current"/>
            <ListBoxItem Content="LTS (Long Term Support)"/>
        </ComboBox>
        <ComboBox x:Name="Box_MSPowerShell" HorizontalAlignment="Left" Margin="320,117,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="2">
            <ListBoxItem Content="Stable"/>
            <ListBoxItem Content="LTS (Long Term Support)"/>
        </ComboBox>
        <ComboBox x:Name="Box_RemoteDesktopManager" HorizontalAlignment="Left" Margin="320,238,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="Free"/>
            <ListBoxItem Content="Enterprise"/>
        </ComboBox>
        <ComboBox x:Name="Box_Zoom" HorizontalAlignment="Left" Margin="320,379,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="VDI Client"/>
            <ListBoxItem Content="VDI Client + Citrix Plugin"/>
        </ComboBox>
        <ComboBox x:Name="Box_Slack" HorizontalAlignment="Left" Margin="320,278,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="Per Machine"/>
            <ListBoxItem Content="Per User"/>
        </ComboBox>
        <ComboBox x:Name="Box_deviceTRUST" HorizontalAlignment="Left" Margin="179,198,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.ColumnSpan="2" Grid.Column="1">
            <ListBoxItem Content="Client"/>
            <ListBoxItem Content="Host"/>
            <ListBoxItem Content="Console"/>
            <ListBoxItem Content="Client + Host"/>
            <ListBoxItem Content="Host + Console"/>
        </ComboBox>
        <ComboBox x:Name="Box_MSEdge" HorizontalAlignment="Left" Margin="179,400,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="1" Grid.ColumnSpan="2">
            <ListBoxItem Content="Developer"/>
            <ListBoxItem Content="Beta"/>
            <ListBoxItem Content="Stable"/>
        </ComboBox>
    </Grid>
</Window>
"@

    #Correction XAML
    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    #Read XAML
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    Try {
        $Form=[Windows.Markup.XamlReader]::Load( $reader )
    }
    Catch {
        Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
        Throw
    }

    # Load XAML Objects In PowerShell  
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object{"trying item $($_.Name)";
        Try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
        Catch {Throw}
    } | out-null

    # Set Variable
    $Script:install = $true
    $Script:download = $true

    # Read LastSettings.txt to get the settings of the last session. (AddScript)
    If (Test-Path "$PSScriptRoot\LastSetting.txt" -PathType leaf) {
        $LastSetting = Get-Content "$PSScriptRoot\LastSetting.txt"
        $WPFBox_Language.SelectedIndex = $LastSetting[0] -as [int]
        $WPFBox_Architecture.SelectedIndex = $LastSetting[1] -as [int]
        $WPFBox_CitrixWorkspaceApp.SelectedIndex = $LastSetting[2] -as [int]
        $WPFBox_MS365Apps.SelectedIndex = $LastSetting[3] -as [int]
        $WPFBox_MSOneDrive.SelectedIndex = $LastSetting[4] -as [int]
        $WPFBox_MSTeams.SelectedIndex = $LastSetting[5] -as [int]
        $WPFBox_Firefox.SelectedIndex = $LastSetting[6] -as [int]
        $WPFBox_TreeSize.SelectedIndex = $LastSetting[7] -as [int]
        $WPFBox_MSDotNetFramework.SelectedIndex = $LastSetting[40] -as [int]
        $WPFBox_MSPowerShell.SelectedIndex = $LastSetting[42] -as [int]
        $WPFBox_RemoteDesktopManager.SelectedIndex = $LastSetting[44] -as [int]
        $WPFBox_Slack.SelectedIndex = $LastSetting[46] -as [int]
        $WPFBox_Zoom.SelectedIndex = $LastSetting[49] -as [int]
        $WPFBox_deviceTRUST.SelectedIndex = $LastSetting[50] -as [int]
        $WPFBox_MSEdge.SelectedIndex = $LastSetting[51] -as [int]
        Switch ($LastSetting[8]) {
            1 { $WPFCheckbox_7ZIP.IsChecked = "True"}
        }
        Switch ($LastSetting[9]) {
            1 { $WPFCheckbox_AdobeProDC.IsChecked = "True"}
        }
        Switch ($LastSetting[10]) {
            1 { $WPFCheckbox_AdobeReaderDC.IsChecked = "True"}
        }
        Switch ($LastSetting[11]) {
            1 { $WPFCheckbox_BISF.IsChecked = "True"}
        }
        Switch ($LastSetting[12]) {
            1 { $WPFCheckbox_CitrixHypervisorTools.IsChecked = "True"}
        }
        Switch ($LastSetting[13]) {
            1 { $WPFCheckbox_CitrixWorkspaceApp.IsChecked = "True"}
        }
        Switch ($LastSetting[14]) {
            1 { $WPFCheckbox_Filezilla.IsChecked = "True"}
        }
        Switch ($LastSetting[15]) {
            1 { $WPFCheckbox_Firefox.IsChecked = "True"}
        }
        Switch ($LastSetting[16]) {
            1 { $WPFCheckbox_FoxitReader.IsChecked = "True"}
        }
        Switch ($LastSetting[17]) {
            1 { $WPFCheckbox_MSFSLogix.IsChecked = "True"}
        }
        Switch ($LastSetting[18]) {
            1 { $WPFCheckbox_GoogleChrome.IsChecked = "True"}
        }
        Switch ($LastSetting[19]) {
            1 { $WPFCheckbox_Greenshot.IsChecked = "True"}
        }
        Switch ($LastSetting[20]) {
            1 { $WPFCheckbox_KeePass.IsChecked = "True"}
        }
        Switch ($LastSetting[21]) {
            1 { $WPFCheckbox_mRemoteNG.IsChecked = "True"}
        }
        Switch ($LastSetting[22]) {
            1 { $WPFCheckbox_MS365Apps.IsChecked = "True"}
        }
        Switch ($LastSetting[23]) {
            1 { $WPFCheckbox_MSEdge.IsChecked = "True"}
        }
        Switch ($LastSetting[24]) {
            1 { $WPFCheckbox_MSOffice2019.IsChecked = "True"}
        }
        Switch ($LastSetting[25]) {
            1 { $WPFCheckbox_MSOneDrive.IsChecked = "True"}
        }
        Switch ($LastSetting[26]) {
            1 { $WPFCheckbox_MSTeams.IsChecked = "True"}
        }
        Switch ($LastSetting[27]) {
            1 { $WPFCheckbox_NotePadPlusPlus.IsChecked = "True"}
        }
        Switch ($LastSetting[28]) {
            1 { $WPFCheckbox_OpenJDK.IsChecked = "True"}
        }
        Switch ($LastSetting[29]) {
            1 { $WPFCheckbox_OracleJava8.IsChecked = "True"}
        }
        Switch ($LastSetting[30]) {
            1 { $WPFCheckbox_TreeSize.IsChecked = "True"}
        }
        Switch ($LastSetting[31]) {
            1 { $WPFCheckbox_VLCPlayer.IsChecked = "True"}
        }
        Switch ($LastSetting[32]) {
            1 { $WPFCheckbox_VMWareTools.IsChecked = "True"}
        }
        Switch ($LastSetting[33]) {
            1 { $WPFCheckbox_WinSCP.IsChecked = "True"}
        }
        Switch ($LastSetting[34]) {
            True { $WPFCheckbox_Download.IsChecked = "True"}
        }
        Switch ($LastSetting[35]) {
            True { $WPFCheckbox_Install.IsChecked = "True"}
        }
        Switch ($LastSetting[36]) {
            1 { $WPFCheckbox_IrfanView.IsChecked = "True"}
        }
        Switch ($LastSetting[37]) {
            1 { $WPFCheckbox_MSTeams_No_AutoStart.IsChecked = "True"}
        }
        Switch ($LastSetting[38]) {
            1 { $WPFCheckbox_deviceTRUST.IsChecked = "True"}
        }
        Switch ($LastSetting[39]) {
            1 { $WPFCheckbox_MSDotNetFramework.IsChecked = "True"}
        }
        Switch ($LastSetting[41]) {
            1 { $WPFCheckbox_MSPowerShell.IsChecked = "True"}
        }
        Switch ($LastSetting[43]) {
            1 { $WPFCheckbox_RemoteDesktopManager.IsChecked = "True"}
        }
        Switch ($LastSetting[45]) {
            1 { $WPFCheckbox_Slack.IsChecked = "True"}
        }
        Switch ($LastSetting[47]) {
            1 { $WPFCheckbox_ShareX.IsChecked = "True"}
        }
        Switch ($LastSetting[48]) {
            1 { $WPFCheckbox_Zoom.IsChecked = "True"}
        }
    }
    
    #// MARK: Event Handler
    # Checkbox SelectAll (AddScript)
    $WPFCheckbox_SelectAll.Add_Checked({
        $WPFCheckbox_7Zip.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_AdobeProDC.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_AdobeReaderDC.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_BISF.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CitrixHypervisorTools.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CitrixWorkspaceApp.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Filezilla.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Firefox.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_FoxitReader.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GoogleChrome.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Greenshot.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_IrfanView.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_KeePass.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_mRemoteNG.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MS365Apps.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSEdge.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSFSLogix.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSOffice2019.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSOneDrive.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSTeams.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_NotePadPlusPlus.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OpenJDK.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OracleJava8.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TreeSize.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_VLCPlayer.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_VMWareTools.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_WinSCP.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_deviceTRUST.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSDotNetFramework.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerShell.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_RemoteDesktopManager.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Slack.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ShareX.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Zoom.IsChecked = $WPFCheckbox_SelectAll.IsChecked
    })
    # Checkbox SelectAll to Uncheck (AddScript)
    $WPFCheckbox_SelectAll.Add_Unchecked({
        $WPFCheckbox_7Zip.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_AdobeProDC.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_AdobeReaderDC.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_BISF.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CitrixHypervisorTools.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CitrixWorkspaceApp.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Filezilla.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Firefox.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_FoxitReader.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GoogleChrome.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Greenshot.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_IrfanView.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_KeePass.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_mRemoteNG.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MS365Apps.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSEdge.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSFSLogix.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSOffice2019.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSOneDrive.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSTeams.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_NotePadPlusPlus.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OpenJDK.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OracleJava8.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TreeSize.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_VLCPlayer.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_VMWareTools.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_WinSCP.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_deviceTRUST.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSDotNetFramework.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerShell.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_RemoteDesktopManager.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Slack.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ShareX.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Zoom.IsChecked = $WPFCheckbox_SelectAll.IsChecked
    })

    # Button Start (AddScript)
    $WPFButton_Start.Add_Click({
        If ($WPFCheckbox_Download.IsChecked -eq $True) {$Script:install = $false}
        Else {$Script:install = $true}
        If ($WPFCheckbox_Install.IsChecked -eq $True) {$Script:download = $false}
        Else {$Script:download = $true}
        If ($WPFCheckbox_7Zip.IsChecked -eq $true) {$Script:7ZIP = 1}
        Else {$Script:7ZIP = 0}
        If ($WPFCheckbox_AdobeProDC.IsChecked -eq $true) {$Script:AdobeProDC = 1}
        Else {$Script:AdobeProDC = 0}
        If ($WPFCheckbox_AdobeReaderDC.IsChecked -eq $true) {$Script:AdobeReaderDC = 1}
        Else {$Script:AdobeReaderDC = 0}
        If ($WPFCheckbox_BISF.IsChecked -eq $true) {$Script:BISF = 1}
        Else {$Script:BISF = 0}
        If ($WPFCheckbox_CitrixHypervisorTools.IsChecked -eq $true) {$Script:Citrix_Hypervisor_Tools = 1}
        Else {$Script:Citrix_Hypervisor_Tools = 0}
        If ($WPFCheckbox_CitrixWorkspaceApp.IsChecked -eq $true) {$Script:Citrix_WorkspaceApp = 1}
        Else {$Script:Citrix_WorkspaceApp = 0}
        If ($WPFCheckbox_Filezilla.IsChecked -eq $true) {$Script:Filezilla = 1}
        Else {$Script:Filezilla = 0}
        If ($WPFCheckbox_Firefox.IsChecked -eq $true) {$Script:Firefox = 1}
        Else {$Script:Firefox = 0}
        If ($WPFCheckbox_MSFSLogix.IsChecked -eq $true) {$Script:FSLogix = 1}
        Else {$Script:FSLogix = 0}
        If ($WPFCheckbox_FoxitReader.Ischecked -eq $true) {$Script:Foxit_Reader = 1}
        Else {$Script:Foxit_Reader = 0}
        If ($WPFCheckbox_GoogleChrome.ischecked -eq $true) {$Script:GoogleChrome = 1}
        Else {$Script:GoogleChrome = 0}
        If ($WPFCheckbox_Greenshot.ischecked -eq $true) {$Script:Greenshot = 1}
        Else {$Script:Greenshot = 0}
        If ($WPFCheckbox_IrfanView.ischecked -eq $true) {$Script:IrfanView = 1}
        Else {$Script:IrfanView = 0}
        If ($WPFCheckbox_KeePass.ischecked -eq $true) {$Script:KeePass = 1}
        Else {$Script:KeePass = 0}
        If ($WPFCheckbox_mRemoteNG.ischecked -eq $true) {$Script:mRemoteNG = 1}
        Else {$Script:mRemoteNG = 0}
        If ($WPFCheckbox_MS365Apps.ischecked -eq $true) {$Script:MS365Apps = 1}
        Else {$Script:MS365Apps = 0}
        If ($WPFCheckbox_MSEdge.ischecked -eq $true) {$Script:MSEdge = 1}
        Else {$Script:MSEdge = 0}
        If ($WPFCheckbox_MSOffice2019.ischecked -eq $true) {$Script:MSOffice2019 = 1}
        Else {$Script:MSOffice2019 = 0}
        If ($WPFCheckbox_MSOneDrive.ischecked -eq $true) {$Script:MSOneDrive = 1}
        Else {$Script:MSOneDrive = 0}
        If ($WPFCheckbox_MSTeams.ischecked -eq $true) {$Script:MSTeams = 1}
        Else {$Script:MSTeams = 0}
        If ($WPFCheckbox_NotePadPlusPlus.ischecked -eq $true) {$Script:NotePadPlusPlus = 1}
        Else {$Script:NotePadPlusPlus = 0}
        If ($WPFCheckbox_OpenJDK.ischecked -eq $true) {$Script:OpenJDK = 1}
        Else {$Script:OpenJDK = 0}
        If ($WPFCheckbox_OracleJava8.ischecked -eq $true) {$Script:OracleJava8 = 1}
        Else {$Script:OracleJava8 = 0}
        If ($WPFCheckbox_TreeSize.ischecked -eq $true) {$Script:TreeSize = 1}
        Else {$Script:TreeSize = 0}
        If ($WPFCheckbox_VLCPlayer.ischecked -eq $true) {$Script:VLCPlayer = 1}
        Else {$Script:VLCPlayer = 0}
        If ($WPFCheckbox_VMWareTools.ischecked -eq $true) {$Script:VMWareTools = 1}
        Else {$Script:VMWareTools = 0}
        If ($WPFCheckbox_WinSCP.ischecked -eq $true) {$Script:WinSCP = 1}
        Else {$Script:WinSCP = 0}        
        If ($WPFCheckbox_MSTeams_No_AutoStart.ischecked -eq $true) {$Script:MSTeamsNoAutoStart = 1}
        Else {$Script:MSTeamsNoAutoStart = 0}
        If ($WPFCheckbox_deviceTRUST.ischecked -eq $true) {$Script:deviceTRUST = 1}
        Else {$Script:deviceTRUST = 0}
        If ($WPFCheckbox_MSDotNetFramework.ischecked -eq $true) {$Script:MSDotNetFramework = 1}
        Else {$Script:MSDotNetFramework = 0}
        If ($WPFCheckbox_MSPowerShell.ischecked -eq $true) {$Script:MSPowerShell = 1}
        Else {$Script:MSPowerShell = 0}
        If ($WPFCheckbox_RemoteDesktopManager.ischecked -eq $true) {$Script:RemoteDesktopManager = 1}
        Else {$Script:RemoteDesktopManager = 0}
        If ($WPFCheckbox_Slack.ischecked -eq $true) {$Script:Slack = 1}
        Else {$Script:Slack = 0}
        If ($WPFCheckbox_ShareX.ischecked -eq $true) {$Script:ShareX = 1}
        Else {$Script:ShareX = 0}
        If ($WPFCheckbox_Zoom.ischecked -eq $true) {$Script:Zoom = 1}
        Else {$Script:Zoom = 0}
        $Script:Language = $WPFBox_Language.SelectedIndex
        $Script:Architecture = $WPFBox_Architecture.SelectedIndex
        $Script:FirefoxChannel = $WPFBox_Firefox.SelectedIndex
        $Script:CitrixWorkspaceAppRelease = $WPFBox_CitrixWorkspaceApp.SelectedIndex
        $Script:MS365AppsChannel = $WPFBox_MS365Apps.SelectedIndex
        $Script:MSOneDriveRing = $WPFBox_MSOneDrive.SelectedIndex
        $Script:MSTeamsRing = $WPFBox_MSTeams.SelectedIndex
        $Script:TreeSizeType = $WPFBox_TreeSize.SelectedIndex
        $Script:MSDotNetFrameworkChannel = $WPFBox_MSDotNetFramework.SelectedIndex
        $Script:MSPowerShellRelease = $WPFBox_MSPowerShell.SelectedIndex
        $Script:RemoteDesktopManagerType = $WPFBox_RemoteDesktopManager.SelectedIndex
        $Script:SlackPlatform = $WPFBox_Slack.SelectedIndex
        $Script:ZoomCitrixClient = $WPFBox_Zoom.SelectedIndex
        $Script:deviceTRUSTPackage = $WPFBox_deviceTRUST.SelectedIndex
        $Script:MSEdgeChannel = $WPFBox_MSEdge.SelectedIndex
        # Write LastSettings.txt to get the settings of the last session. (AddScript)
        $Language,$Architecture,$CitrixWorkspaceAppRelease,$MS365AppsChannel,$MSOneDriveRing,$MSTeamsRing,$FirefoxChannel,$TreeSizeType,$7ZIP,$AdobeProDC,$AdobeReaderDC,$BISF,$Citrix_Hypervisor_Tools,$Citrix_WorkspaceApp,$Filezilla,$Firefox,$Foxit_Reader,$FSLogix,$GoogleChrome,$Greenshot,$KeePass,$mRemoteNG,$MS365Apps,$MSEdge,$MSOffice2019,$MSOneDrive,$MSTeams,$NotePadPlusPlus,$OpenJDK,$OracleJava8,$TreeSize,$VLCPlayer,$VMWareTools,$WinSCP,$WPFCheckbox_Download.IsChecked,$WPFCheckbox_Install.IsChecked,$IrfanView,$MSTeamsNoAutoStart,$deviceTRUST,$MSDotNetFramework,$MSDotNetFrameworkChannel,$MSPowerShell,$MSPowerShellRelease,$RemoteDesktopManager,$RemoteDesktopManagerType,$Slack,$SlackPlatform,$ShareX,$Zoom,$ZoomCitrixClient,$deviceTRUSTPackage,$MSEdgeChannel | out-file -filepath "$PSScriptRoot\LastSetting.txt"
        Write-Host "GUI Mode"
        $Form.Close()
    })

    # Button Cancel                                                                    
    $WPFButton_Cancel.Add_Click({
        $Script:install = $true
        $Script:download = $true
        Write-Host -Foregroundcolor Red "GUI Mode Canceled - Nothing happens"
        $Form.Close()
        Break
    })

    # Image Logo
    $WPFImage_Logo.Add_MouseLeftButtonUp({
        [system.Diagnostics.Process]::start('https://www.deyda.net')
    })

    # Shows the form
    $Form.ShowDialog() | out-null
}

#===========================================================================

Write-Host -Foregroundcolor DarkGray "Software selection"

#// MARK: Define and reset variables
$Date = $Date = Get-Date -UFormat "%m.%d.%Y"
$Script:install = $install
$Script:download = $download

# Define the variables for the unattended install or download (Parameter -list) (AddScript)
If ($list -eq $True) {
    # Select Language (If this is selectable at download)
    # 0 = Danish
    # 1 = Dutch
    # 2 = English
    # 3 = Finnish
    # 4 = French
    # 5 = German
    # 6 = Italian
    # 7 = Japanese
    # 8 = Korean
    # 9 = Norwegian
    # 10 = Polish
    # 11 = Portuguese
    # 12 = Russian
    # 13 = Spanish
    # 14 = Swedish
    $Language = 2

    # Select Architecture (If this is selectable at download)
    # 0 = x64
    # 1 = x86
    $Architecture = 0

    # Software Release / Ring / Channel / Type ?!
    # Citrix Workspace App
    # 0 = Current Release
    # 1 = Long Term Service Release
    $CitrixWorkspaceAppRelease = 1

    # deviceTRUST
    # 0 = Client
    # 1 = Host
    # 2 = Console
    # 3 = Client + Host
    # 4 = Host + Console
    $deviceTRUSTPackage = 1

    # Microsoft .Net Framework
    # 0 = Current Channel
    # 1 = LTS (Long Term Support) Channel
    $MSDotNetFrameworkChannel = 1

    # Microsoft 365 Apps
    # 0 = Current (Preview) Channel
    # 1 = Current Channel
    # 2 = Monthly Enterprise Channel
    # 3 = Semi-Annual Enterprise (Preview) Channel
    # 4 = Semi-Annual Enterprise Channel
    $MS365AppsChannel = 4

    # Microsoft Edge
    # 0 = Developer Channel
    # 1 = Beta Channel
    # 2 = Stable Channel
    $MSEdgeChannel = 2

    # Microsoft OneDrive
    # 0 = Insider Ring
    # 1 = Production Ring
    # 2 = Enterprise Ring
    $MSOneDriveRing = 2

    # Microsoft PowerShell
    # 0 = Stable Release
    # 1 = LTS (Long Term Support) Release
    $MSPowerShellRelease = 1

    # Microsoft Teams
    # 0 = Developer Ring
    # 1 = Preview Ring
    # 2 = General Ring
    $MSTeamsRing = 2

    # Microsoft Teams AutoStart
    # 0 = AutoStart Microsoft Teams
    # 1 = No AutoStart (Delete HKLM Registry Entry)
    $MSTeamsNoAutoStart = 0

    # Mozilla Firefox
    # 0 = Current
    # 1 = ESR
    $FirefoxChannel = 0

    # Remote Desktop Manager
    # 0 = Free
    # 1 = Enterprise
    $RemoteDesktopManagerType = 0

    # Slack
    # 0 = Per Machine
    # 1 = Per User
    $SlackPlatform = 0

    # TreeSize
    # 0 = Free
    # 1 = Professional
    $TreeSizeType = 0

    # Zoom
    # 0 = VDI Installer
    # 1 = VDI Installer + Citrix Plugin
    $ZoomCitrixClient = 1

    # Select Software
    # 0 = Not selected
    # 1 = Selected
    $7ZIP = 0
    $AdobeProDC = 0 # Only Update @ the moment
    $AdobeReaderDC = 0
    $BISF = 0
    $Citrix_Hypervisor_Tools = 0
    $Citrix_WorkspaceApp = 0
    $deviceTRUST = 0
    $Filezilla = 0
    $Firefox = 0
    $Foxit_Reader = 0
    $FSLogix = 0
    $GoogleChrome = 0
    $Greenshot = 0
    $IrfanView = 0
    $KeePass = 0
    $mRemoteNG = 0
    $MSDotNetFramework = 0
    $MS365Apps = 0 # Automatically created install.xml is used. Please replace this file if you want to change the installation.
    $MSEdge = 0
    $MSOffice2019 = 0 # Automatically created install.xml is used. Please replace this file if you want to change the installation.
    $MSOneDrive = 0
    $MSPowerShell = 0
    $MSTeams = 0
    $NotePadPlusPlus = 0
    $OpenJDK = 0
    $OracleJava8 = 0
    $RemoteDesktopManager = 0
    $Slack = 0
    $ShareX = 0
    $TreeSize = 0
    $VLCPlayer = 0
    $VMWareTools = 0
    $WinSCP = 0
    $Zoom = 0

    Write-Host "Unattended Mode."
}
Else {
    # Cleanup of the used vaiables (AddScript)
    Clear-Variable -name 7ZIP,AdobeProDC,AdobeReaderDC,BISF,Citrix_Hypervisor_Tools,Filezilla,Firefox,Foxit_Reader,FSLogix,Greenshot,GoogleChrome,KeePass,mRemoteNG,MS365Apps,MSEdge,MSOffice2019,MSTeams,NotePadPlusPlus,MSOneDrive,OpenJDK,OracleJava8,TreeSize,VLCPlayer,VMWareTools,WinSCP,Citrix_WorkspaceApp,Architecture,FirefoxChannel,CitrixWorkspaceAppRelease,Language,MS365AppsChannel,MSOneDriveRing,MSTeamsRing,TreeSizeType,IrfanView,MSTeamsNoAutoStart,deviceTRUST,MSDotNetFramework,MSDotNetFrameworkChannel,MSPowerShell,MSPowerShellRelease,RemoteDesktopManager,RemoteDesktopManagerType,Slack,SlackPlatform,ShareX,Zoom,ZoomCitrixClient,deviceTRUSTPackage,deviceTRUSTClient,deviceTRUSTConsole,deviceTRUSTHost,MSEdgeChannel -ErrorAction SilentlyContinue
    gui_mode
}

#// MARK: Variable definition (Architecture,Language etc) (AddScript)
Switch ($Architecture) {
    0 { $ArchitectureClear = 'x64'}
    1 { $ArchitectureClear = 'x86'}
}

Switch ($Language) {
    0 { $LanguageClear = 'Danish'}
    1 { $LanguageClear = 'Dutch'}
    2 { $LanguageClear = 'English'}
    3 { $LanguageClear = 'Finnish'}
    4 { $LanguageClear = 'French'}
    5 { $LanguageClear = 'German'}
    6 { $LanguageClear = 'Italian'}
    7 { $LanguageClear = 'Japanese'}
    8 { $LanguageClear = 'Korean'}
    9 { $LanguageClear = 'Norwegian'}
    10 { $LanguageClear = 'Polish'}
    11 { $LanguageClear = 'Portuguese'}
    12 { $LanguageClear = 'Russian'}
    13 { $LanguageClear = 'Spanish'}
    14 { $LanguageClear = 'Swedish'}
}

$AdobeLanguageClear = $LanguageClear
Switch ($LanguageClear) {
    Polish { $AdobeLanguageClear = 'English'}
    Portuguese { $AdobeLanguageClear = 'English'}
    Russian { $AdobeLanguageClear = 'English'}
    Swedish { $AdobeLanguageClear = 'English'}
}

$AdobeArchitectureClear = 'x86'
Switch ($LanguageClear) {
    English { $AdobeArchitectureClear = $ArchitectureClear}
}

Switch ($CitrixWorkspaceAppRelease) {
    0 { $CitrixWorkspaceAppReleaseClear = 'Current Release'}
    1 { $CitrixWorkspaceAppReleaseClear = 'LTSR'}
}

Switch ($deviceTRUSTPackage) {
    0 { $deviceTRUSTClient = $True}
    1 { $deviceTRUSTHost = $True}
    2 { $deviceTRUSTConsole = $True}
    3 { $deviceTRUSTClient = $True
        $deviceTRUSTHost = $True}
    4 { $deviceTRUSTConsole = $True
        $deviceTRUSTHost = $True}
}

$FoxitReaderLanguageClear = $LanguageClear
Switch ($LanguageClear) {
    Japanese { $FoxitReaderLanguageClear = 'English'}
}

Switch ($MSDotNetFrameworkChannel) {
    0 { $MSDotNetFrameworkChannelClear = 'Current'}
    1 { $MSDotNetFrameworkChannelClear = 'LTS'}
}

Switch ($MS365AppsChannel) {
    0 { $MS365AppsChannelClear = 'CurrentPreview'}
    1 { $MS365AppsChannelClear = 'Current'}
    2 { $MS365AppsChannelClear = 'MonthlyEnterprise'}
    3 { $MS365AppsChannelClear = 'SemiAnnualPreview'}
    4 { $MS365AppsChannelClear = 'SemiAnnual'}
}

Switch ($MS365AppsChannel) {
    0 { $MS365AppsChannelClearDL = 'Monthly (Targeted)'}
    1 { $MS365AppsChannelClearDL = 'Monthly'}
    2 { $MS365AppsChannelClearDL = 'Monthly Enterprise'}
    3 { $MS365AppsChannelClearDL = 'Semi-Annual Channel (Targeted)'}
    4 { $MS365AppsChannelClearDL = 'Semi-Annual Channel'}
}

Switch ($Architecture) {
    0 { $MS365AppsArchitectureClear = '64'}
    1 { $MS365AppsArchitectureClear = '32'}
}

Switch ($Language) {
    0 { $MS365AppsLanguageClear = 'da-DK'}
    1 { $MS365AppsLanguageClear = 'nl-NL'}
    2 { $MS365AppsLanguageClear = 'en-US'}
    3 { $MS365AppsLanguageClear = 'fi-FI'}
    4 { $MS365AppsLanguageClear = 'fr-FR'}
    5 { $MS365AppsLanguageClear = 'de-DE'}
    6 { $MS365AppsLanguageClear = 'it-IT'}
    7 { $MS365AppsLanguageClear = 'ja-JP'}
    8 { $MS365AppsLanguageClear = 'ko-KR'}
    9 { $MS365AppsLanguageClear = 'nb-NO'}
    10 { $MS365AppsLanguageClear = 'pl-PL'}
    11 { $MS365AppsLanguageClear = 'pt-PT'}
    12 { $MS365AppsLanguageClear = 'ru-RU'}
    13 { $MS365AppsLanguageClear = 'es-ES'}
    14 { $MS365AppsLanguageClear = 'sv-SE'}
}

Switch ($MSEdgeChannel) {
    0 { $MSEdgeChannelClear = 'Dev'}
    1 { $MSEdgeChannelClear = 'Beta'}
    2 { $MSEdgeChannelClear = 'Stable'}
}

Switch ($MSOneDriveRing) {
    0 { $MSOneDriveRingClear = 'Insider'}
    1 { $MSOneDriveRingClear = 'Production'}
    2 { $MSOneDriveRingClear = 'Enterprise'}
}

Switch ($Architecture) {
    0 { $MSOneDriveArchitectureClear = 'AMD64'}
    1 { $MSOneDriveArchitectureClear = 'x86'}
}

Switch ($MSPowerShellRelease) {
    0 { $MSPowerShellReleaseClear = 'Stable'}
    1 { $MSPowerShellReleaseClear = 'LTS'}
}

Switch ($MSTeamsRing) {
    0 { $MSTeamsRingClear = 'Developer'}
    1 { $MSTeamsRingClear = 'Preview'}
    2 { $MSTeamsRingClear = 'General'}
}

Switch ($FirefoxChannel) {
    0 { $FirefoxChannelClear = 'LATEST'}
    1 { $FirefoxChannelClear = 'ESR'}
}

Switch ($LanguageClear) {
    English { $FFLanguageClear = 'en-US'}
    Danish { $FFLanguageClear = 'en-US'}
    Russian { $FFLanguageClear = 'ru'}
    Dutch { $FFLanguageClear = 'nl'}
    Finnish { $FFLanguageClear = 'en-US'}
    French { $FFLanguageClear = 'fr'}
    German { $FFLanguageClear = 'de'}
    Italian { $FFLanguageClear = 'it'}
    Japanese { $FFLanguageClear = 'ja'}
    Korean { $FFLanguageClear = 'en-US'}
    Norwegian { $FFLanguageClear = 'en-US'}
    Polish { $FFLanguageClear = 'en-US'}
    Portuguese { $FFLanguageClear = 'pt-PT'}
    Spanish { $FFLanguageClear = 'es-ES'}
    Swedish { $FFLanguageClear = 'sv-SE'}
}

Switch ($SlackPlatform) {
    0 { $SlackPlatformClear = 'PerMachine'}
    1 { $SlackPlatformClear = 'PerUser'}
}

Switch ($SlackPlatform) {
    0 { $SlackArchitectureClear = $ArchitectureClear}
    1 { $SlackArchitectureClear = 'x64'}
}

Write-Host -ForegroundColor Green "Software selection done."
Write-Output ""

If ($install -eq $False) {
    #// Mark: Install / Update Evergreen module
    Write-Host -ForegroundColor DarkGray "Install / Update Evergreen module!"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    If (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {
        Write-Host "Install Nuget include dependencies."
        Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies | Install-PackageProvider -Force | Out-Null
        Write-Host -ForegroundColor Green "Install Nuget include dependencies done."
    }
    If (!(Get-Module -ListAvailable -Name Evergreen)) {
        Write-Host "Install Evergreen module."
        Install-Module Evergreen -Force | Import-Module Evergreen
        Write-Host -ForegroundColor Green "Install Evergreen module done."
        Write-Output ""
    }
    Else {
        Write-Host "Update Evergreen module."
        Update-Module Evergreen -force
        Write-Host -ForegroundColor Green "Update Evergreen module done."
        Write-Output ""
    }

    Write-Host -ForegroundColor DarkGray "Starting downloads..."
    Write-Output ""

    # Download script part (AddScript)
    #// Mark: Download 7-ZIP
    If ($7ZIP -eq 1) {
        $Product = "7-Zip"
        $PackageName = "7-Zip_" + "$ArchitectureClear"
        $7ZipD = Get-EvergreenApp -Name 7zip | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.Type -eq "exe" }
        $Version = $7ZipD.Version
        $URL = $7ZipD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Adobe Pro DC Update
    If ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        $PackageName = "Adobe_Pro_DC_Update"
        $AdobeProD = Get-EvergreenApp -Name AdobeAcrobat | Where-Object { $_.Track -eq "DC" -and $_.Language -eq "Multi" }
        $Version = $AdobeProD.Version
        $URL = $AdobeProD.uri
        $InstallerType = "msp"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Include *.msp, *.log, Version.txt, Download* -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source)) 
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Adobe Reader DC
    If ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        $PackageName = "Adobe_Reader_DC_"
        $AdobeReaderD = Get-EvergreenApp -Name AdobeAcrobatReaderDC | Where-Object {$_.Architecture -eq "$AdobeArchitectureClear" -and $_.Language -eq "$AdobeLanguageClear"}
        $Version = $AdobeReaderD.Version
        $URL = $AdobeReaderD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "$AdobeArchitectureClear" + "$AdobeLanguageClear" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$AdobeArchitectureClear" + "_$AdobeLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $AdobeArchitectureClear $AdobeLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $AdobeArchitectureClear $AdobeLanguageClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download BIS-F
    If ($BISF -eq 1) {
        $Product = "BIS-F"
        $PackageName = "setup-BIS-F"
        $BISFD = Get-EvergreenApp -Name BISF | Where-Object { $_.Type -eq "msi" }
        $Version = $BISFD.Version
        $URL = $BISFD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Exclude *.ps1, *.lnk -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Citrix Hypervisor Tools
    If ($Citrix_Hypervisor_Tools -eq 1) {
        $Product = "Citrix Hypervisor Tools"
        $PackageName = "managementagent" + "$ArchitectureClear"
        $CitrixHypervisor = Get-EvergreenApp -Name CitrixVMTools | Where-Object {$_.Architecture -eq "$ArchitectureClear"} | Select-Object -Last 1
        $Version = $CitrixHypervisor.Version
        $URL = $CitrixHypervisor.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting Download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\Citrix\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\Citrix\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Citrix WorkspaceApp
    If ($Citrix_WorkspaceApp -eq 1) {
        $Product = "Citrix WorkspaceApp $CitrixWorkspaceAppReleaseClear"
        $PackageName = "CitrixWorkspaceApp"
        $WSACD = Get-EvergreenApp -Name CitrixWorkspaceApp -WarningAction:SilentlyContinue | Where-Object { $_.Title -like "*Workspace*" -and "*$CitrixWorkspaceAppReleaseClear*" -and $_.Platform -eq "Windows" -and $_.Title -like "*$CitrixWorkspaceAppReleaseClear*" }
        $Version = $WSACD.Version
        $URL = $WSACD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -EA SilentlyContinue
        If (!(Test-Path -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility")) { New-Item -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility" -ItemType Directory | Out-Null }
        If (!(Test-Path -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.exe")) {
            Write-Host -ForegroundColor Magenta "Download Citrix Receiver Cleanup Utility"
            Get-Download https://fileservice.citrix.com/downloadspecial/support/article/CTX137494/downloads/ReceiverCleanupUtility.zip "$PSScriptRoot\Citrix\ReceiverCleanupUtility\" ReceiverCleanupUtility.zip -includeStats
            #Invoke-WebRequest -Uri https://fileservice.citrix.com/downloadspecial/support/article/CTX137494/downloads/ReceiverCleanupUtility.zip -OutFile ("$PSScriptRoot\Citrix\ReceiverCleanupUtility\" + "ReceiverCleanupUtility.zip")
            Expand-Archive -path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.zip" -destinationpath "$PSScriptRoot\Citrix\ReceiverCleanupUtility\"
            Remove-Item -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.zip" -Force
            Write-Host -ForegroundColor Green "Download Citrix Receiver Cleanup Utility finished!"
            Write-Output ""
        }
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\Citrix\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\Citrix\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download deviceTRUST
    If ($deviceTRUST -eq 1) {
        $Product = "deviceTRUST"
        $PackageName = "deviceTRUST"
        $URLVersion = "https://docs.devicetrust.com/docs/download/"
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($URLVersion) -SessionVariable websession
        $regexAppVersion = "<td>\d\d.\d.\d\d\d+</td>"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $Version = $webVersion.Trim("</td>").Trim("</td>")
        $URL = "https://storage.devicetrust.com/download/deviceTRUST-$Version.zip"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -UseBasicParsing -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            expand-archive -path "$PSScriptRoot\$Product\deviceTRUST.zip" -destinationpath "$PSScriptRoot\$Product"
            Remove-Item -Path "$PSScriptRoot\$Product\deviceTRUST.zip" -Force
            expand-archive -path "$PSScriptRoot\$Product\dtpolicydefinitions-$Version.0.zip" -destinationpath "$PSScriptRoot\$Product\ADMX"
            copy-item -Path "$PSScriptRoot\$Product\ADMX\*" -Destination "$PSScriptRoot\ADMX\deviceTRUST" -Force
            Remove-Item -Path "$PSScriptRoot\$Product\ADMX" -Force -Recurse
            Remove-Item -Path "$PSScriptRoot\$Product\dtpolicydefinitions-$Version.0.zip" -Force
            Remove-Item -Path "$PSScriptRoot\$Product\dtreporting-$Version.0.zip" -Force
            If (Test-Path -Path "$PSScriptRoot\$Product\dtdemotool-release-$Version.0.exe") {Remove-Item -Path "$PSScriptRoot\$Product\dtdemotool-release-$Version.0.exe" -Force}
            Switch ($Architecture) {
                0 {
                    Get-ChildItem -Path "$PSScriptRoot\$Product" | Where-Object Name -like *"x86"* | Remove-Item
                    Rename-Item -Path "$PSScriptRoot\$Product\dtclient-release-$Version.0.exe" -NewName "dtclient-release.exe"
                    Rename-Item -Path "$PSScriptRoot\$Product\dtconsole-x64-release-$Version.0.msi" -NewName "dtconsole-x64-release.msi"
                    Rename-Item -Path "$PSScriptRoot\$Product\dthost-x64-release-$Version.0.msi" -NewName "dthost-x64-release.msi"
                }
                1 {
                    Get-ChildItem -Path "$PSScriptRoot\$Product" | Where-Object Name -like *"x64"* | Remove-Item
                    Rename-Item -Path "$PSScriptRoot\$Product\dtclient-release-$Version.0.exe" -NewName "dtclient-release.exe"
                    Rename-Item -Path "$PSScriptRoot\$Product\dtconsole-x86-release-$Version.0.msi" -NewName "dtconsole-x86-release.msi"
                    Rename-Item -Path "$PSScriptRoot\$Product\dthost-x86-release-$Version.0.msi" -NewName "dthost-x86-release.msi"
                }
            }
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Filezilla
    If ($Filezilla -eq 1) {
        $Product = "Filezilla"
        $PackageName = "Filezilla-win64"
        $FilezillaD = Get-EvergreenApp -Name Filezilla | Where-Object { $_.URI -like "*win64*"}
        $Version = $FilezillaD.Version
        $URL = $FilezillaD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            #Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Foxit Reader
    If ($Foxit_Reader -eq 1) {
        $Product = "Foxit Reader"
        $PackageName = "FoxitReader-Setup-" + "$FoxitReaderLanguageClear"
        $Foxit_ReaderD = Get-EvergreenApp -Name FoxitReader | Where-Object {$_.Language -eq "$FoxitReaderLanguageClear"}
        $Version = $Foxit_ReaderD.Version
        $URL = $Foxit_ReaderD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitReaderLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $FoxitReaderLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $FoxitReaderLanguageClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Google Chrome
    If ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        $PackageName = "googlechromestandaloneenterprise_" + "$ArchitectureClear"
        $ChromeD = Get-EvergreenApp -Name GoogleChrome | Where-Object { $_.Architecture -eq "$ArchitectureClear" }
        $Version = $ChromeD.Version
        $URL = $ChromeD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Greenshot
    If ($Greenshot -eq 1) {
        $Product = "Greenshot"
        $PackageName = "Greenshot-INSTALLER-x86"
        $GreenshotD = Get-EvergreenApp -Name Greenshot | Where-Object { $_.Architecture -eq "x86" -and $_.URI -like "*INSTALLER*" -and $_.Type -like "exe"}
        $Version = $GreenshotD.Version
        $URL = $GreenshotD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download IrfanView
    If ($IrfanView -eq 1) {
        Function Get-IrfanView {
            <#
                .NOTES
                    Author: Trond Eirik Haavarstein
                    Twitter: @xenappblog
            #>
            [OutputType([System.Management.Automation.PSObject])]
            [CmdletBinding()]
            Param ()
                $url = "https://www.irfanview.com/"
            Try {
                $web = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
            }
            Catch {
                Throw "Failed to connect to URL: $url with error $_."
                Break
            }
            Finally {
                $m = $web.ToString() -split "[`r`n]" | Select-String "Version" | Select-Object -First 1
                $m = $m -replace "<((?!@).)*?>"
                $m = $m.Replace(' ','')
                $Version = $m -replace "Version"
                $File = $Version -replace "\.",""
                $x32 = "http://download.betanews.com/download/967963863-1/iview$($File)_setup.exe"
                $x64 = "http://download.betanews.com/download/967963863-1/iview$($File)_x64_setup.exe"
        
        
                $PSObjectx32 = [PSCustomObject] @{
                Version      = $Version
                Architecture = "x86"
                Language     = "english"
                URI          = $x32
                }
        
                $PSObjectx64 = [PSCustomObject] @{
                Version      = $Version
                Architecture = "x64"
                Language     = "english"
                URI          = $x64
                }
                Write-Output -InputObject $PSObjectx32
                Write-Output -InputObject $PSObjectx64
            }
        }
        $Product = "IrfanView"
        $PackageName = "IrfanView" + "$ArchitectureClear"
        $IrfanViewD = Get-IrfanView | Where-Object {$_.Architecture -eq "$ArchitectureClear"}
        $Version = $IrfanViewD.Version
        $URL = $IrfanViewD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path $VersionPath -EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path $VersionPath -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download KeePass
    If ($KeePass -eq 1) {
        $Product = "KeePass"
        $PackageName = "KeePass"
        $KeePassD = Get-EvergreenApp -Name KeePass | Where-Object { $_.Type -eq "msi" }
        $Version = $KeePassD.Version
        $URL = $KeePassD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"-EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft .Net Framework
    If ($MSDotNetFramework -eq 1) {
        $Product = "Microsoft Dot Net Framework"
        $PackageName = "NetFramework-runtime_" + "$ArchitectureClear" + "_$MSDotNetFrameworkChannelClear"
        $MSDotNetFrameworkD = Get-EvergreenApp -Name Microsoft.NET | Where-Object {$_.Architecture -eq "$ArchitectureClear" -and $_.Channel -eq "$MSDotNetFrameworkChannelClear"}
        $Version = $MSDotNetFrameworkD.Version
        $URL = $MSDotNetFrameworkD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSDotNetFrameworkChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear $MSDotNetFrameworkChannelClear Channel"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $MSDotNetFrameworkChannelClear Channel $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft 365 Apps
    If ($MS365Apps -eq 1) {
        $Product = "Microsoft 365 Apps"
        $PackageName = "setup_" + "$MS365AppsChannelClear"
        $MS365AppsD = Get-EvergreenApp -Name Microsoft365Apps | Where-Object {$_.Channel -eq "$MS365AppsChannelClearDL"}
        $Version = $MS365AppsD.Version
        $URL = $MS365AppsD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MS365AppsChannelClear setup file"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!(Test-Path -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear")) {New-Item -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear" -ItemType Directory | Out-Null}
        If (!(Test-Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\remove.xml" -PathType leaf)) {
            Write-Host "Create remove.xml"
            [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
            [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                $XML.appendChild($Root) | out-null
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
            $XML.Save("$PSScriptRoot\$Product\$MS365AppsChannelClear\remove.xml")
            Write-Host -ForegroundColor Green "Create remove.xml finished!"
        }
        If (!(Test-Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\install.xml" -PathType leaf)) {
            Write-Host "Create install.xml"
            [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
            [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                $XML.appendChild($Root) | out-null
            [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
                $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product\$MS365AppsChannelClear")
                $Node1.SetAttribute("OfficeClientEdition","$MS365AppsArchitectureClear")
                $Node1.SetAttribute("Channel","$MS365AppsChannelClear")
            [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                $Node2.SetAttribute("ID","O365ProPlusRetail")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                $Node3.SetAttribute("ID","MatchOS")
                $Node3.SetAttribute("Fallback","en-us")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                $Node3.SetAttribute("ID","$MS365AppsLanguageClear")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","Teams")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","Lync")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","Groove")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","OneDrive")
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
                $XML.Save("$PSScriptRoot\$Product\$MS365AppsChannelClear\install.xml")
            Write-Host -ForegroundColor Green "Create install.xml finished!"
        }
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            $LogPS = "$PSScriptRoot\$Product\$MS365AppsChannelClear\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\$MS365AppsChannelClear\*" -Recurse -Exclude install.xml,remove.xml
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $MS365AppsChannelClear $Version setup file"
            Get-Download $URL "$PSScriptRoot\$Product\$MS365AppsChannelClear" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\$MS365AppsChannelClear\" + ($Source))
            Write-Host -ForegroundColor Green "Download of the new version $Version setup file finished!"
            # Download Apps 365 install files
            If (!(Test-Path -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Office\Data\$Version")) {
                Write-Host "Starting download of $Product install files"
                $DApps365 = @(
                    "/download install.xml"
                )
                set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                Start-Process ".\$Source" -ArgumentList $DApps365 -wait -NoNewWindow
                set-location $PSScriptRoot
                Write-Host -ForegroundColor Green "Download of the new version $Version install files finished!"
            }
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Edge
    If ($MSEdge -eq 1) {
        $Product = "Microsoft Edge"
        $PackageName = "MicrosoftEdgeEnterprise_" + "$ArchitectureClear" + "_$MSEdgeChannelClear"
        $EdgeD = Get-EvergreenApp -Name MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Release -eq "Enterprise" -and $_.Channel -eq "$MSEdgeChannelClear" -and $_.Architecture -eq "$ArchitectureClear" }
        #$EdgeURL = $EdgeURL | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $EdgeD.Version
        $URL = $EdgeD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSEdgeChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product $MSEdgeChannelClear $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $MSEdgeChannelClear $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft FSLogix
    If ($FSLogix -eq 1) {
        $Product = "Microsoft FSLogix"
        $PackageName = "FSLogixAppsSetup"
        $FSLogixD = Get-EvergreenApp -Name MicrosoftFSLogixApps
        $Version = $FSLogixD.Version
        $URL = $FSLogixD.uri
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Install\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product\Install")) { New-Item -Path "$PSScriptRoot\$Product\Install" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\Install\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\Install\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Install\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\Install" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\Install\" + ($Source))
            expand-archive -path "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.zip" -destinationpath "$PSScriptRoot\$Product\Install"
            Remove-Item -Path "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.zip" -Force
            Switch ($Architecture) {
                1 {
                    Move-Item -Path "$PSScriptRoot\$Product\Install\Win32\Release\*" -Destination "$PSScriptRoot\$Product\Install"
                }
                0 {
                    Move-Item -Path "$PSScriptRoot\$Product\Install\x64\Release\*" -Destination "$PSScriptRoot\$Product\Install"
                }
            }
            Remove-Item -Path "$PSScriptRoot\$Product\Install\Win32" -Force -Recurse
            Remove-Item -Path "$PSScriptRoot\$Product\Install\x64" -Force -Recurse
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Office 2019
    If ($MSOffice2019 -eq 1) {
        $Product = "Microsoft Office 2019"
        $PackageName = "setup"
        $MSOffice2019D = Get-EvergreenApp -Name Microsoft365Apps | Where-Object {$_.Channel -eq "Office 2019 Enterprise"}
        $Version = $MSOffice2019D.Version
        $URL = $MSOffice2019D.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product setup file"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
        If (!(Test-Path "$PSScriptRoot\$Product\remove.xml" -PathType leaf)) {
            Write-Host "Create remove.xml"
            [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
            [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                $XML.appendChild($Root) | out-null
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
            Write-Host -ForegroundColor Green  "Create remove.xml finished!"
        }
        If (!(Test-Path "$PSScriptRoot\$Product\install.xml" -PathType leaf)) {
            Write-Host "Create install.xml"
            [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
            [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                $XML.appendChild($Root) | out-null
            [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
                $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product")
                $Node1.SetAttribute("OfficeClientEdition","$MS365AppsArchitectureClear")
                $Node1.SetAttribute("Channel","PerpetualVL2019")
            [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                $Node2.SetAttribute("ID","ProPlus2019Volume")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                $Node3.SetAttribute("ID","MatchOS")
                $Node3.SetAttribute("Fallback","en-us")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                $Node3.SetAttribute("ID","$MS365AppsLanguageClear")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","Teams")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","Lync")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","Groove")
            [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                $Node3.SetAttribute("ID","OneDrive")
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
            Write-Host -ForegroundColor Green  "Create install.xml finished!"
        }
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse -Exclude install.xml,remove.xml
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version setup file"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            # Download MS Office 2019 install files
            If (!(Test-Path -Path "$PSScriptRoot\$Product\Office\Data\$Version")) {
                Write-Host "Starting download of $Product install files"
                $DOffice2019 = @(
                    "/download install.xml"
                )
                set-location $PSScriptRoot\$Product
                Start-Process ".\setup.exe" -ArgumentList $DOffice2019 -wait -NoNewWindow
                set-location $PSScriptRoot
                Write-Host -ForegroundColor Green "Download of the new version $Version install files finished!"
            }
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft OneDrive
    If ($MSOneDrive -eq 1) {
        $Product = "Microsoft OneDrive"
        $PackageName = "OneDriveSetup-" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear"
        $MSOneDriveD = Get-EvergreenApp -Name MicrosoftOneDrive | Where-Object { $_.Ring -eq "$MSOneDriveRingClear" -and $_.Type -eq "Exe" -and $_.Architecture -eq "$MSOneDriveArchitectureClear"} | Sort-Object -Property Version -Descending | Select-Object -Last 1
        $Version = $MSOneDriveD.Version
        $URL = $MSOneDriveD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft PowerShell
    If ($MSPowerShell -eq 1) {
        $Product = "Microsoft PowerShell"
        $PackageName = "PowerShell" + "$ArchitectureClear" + "_$MSPowerShellReleaseClear"
        $MSPowershellD = Get-EvergreenApp -Name MicrosoftPowerShell | Where-Object {$_.Architecture -eq "$ArchitectureClear" -and $_.Release -eq "$MSPowerShellReleaseClear"}
        $Version = $MSPowershellD.Version
        $URL = $MSPowershellD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSPowerShellReleaseClear" + ".txt"
        $CurrentVersion = Get-Content -Path $VersionPath -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear $MSPowerShellReleaseClear Release"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $MSPowerShellReleaseClear Release $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Teams
    If ($MSTeams -eq 1) {
        $Product = "Microsoft Teams"
        $PackageName = "Teams_" + "$ArchitectureClear" + "_$MSTeamsRingClear"
        If ($MSTeamsRingClear -eq 'Developer') {
            $TeamsD = Get-MicrosoftTeamsDev | Where-Object { $_.Architecture -eq "$ArchitectureClear"}
        }
        Else {
            $TeamsD = Get-EvergreenApp -Name MicrosoftTeams | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.Ring -eq "$MSTeamsRingClear"}
        }
        $Version = $TeamsD.Version
        $URL = $TeamsD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear $MSTeamsRingClear Ring"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $MSTeamsRingClear Ring $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Mozilla Firefox
    If ($Firefox -eq 1) {
        $Product = "Mozilla Firefox"
        $PackageName = "Firefox_Setup_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "_$FFLanguageClear"
        $FirefoxD = Get-EvergreenApp -Name MozillaFirefox | Where-Object { $_.Type -eq "msi" -and $_.Architecture -eq "$ArchitectureClear" -and $_.Channel -like "*$FirefoxChannelClear*" -and $_.Language -eq "$FFLanguageClear"}
        $Version = $FirefoxD.Version
        $URL = $FirefoxD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "$FFLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download mRemoteNG
    If ($mRemoteNG -eq 1) {
        $Product = "mRemoteNG"
        $PackageName = "mRemoteNG"
        $mRemoteNGD = Get-EvergreenApp -Name mRemoteNG | Where-Object { $_.Type -eq "msi" }
        $Version = $mRemoteNGD.Version
        $URL = $mRemoteNGD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Notepad ++
    If ($NotePadPlusPlus -eq 1) {
        $Product = "NotePadPlusPlus"
        $PackageName = "NotePadPlusPlus_" + "$ArchitectureClear"
        $NotepadD = Get-EvergreenApp -Name NotepadPlusPlus | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.Type -eq "exe" }
        $Version = $NotepadD.Version
        $URL = $NotepadD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Get-ChildItem "$PSScriptRoot\$Product\" -Exclude lang | Remove-Item -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download openJDK
    If ($OpenJDK -eq 1) {
        $Product = "open JDK"
        $PackageName = "OpenJDK" + "$ArchitectureClear"
        $OpenJDKD = Get-EvergreenApp -Name OpenJDK | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.URI -like "*msi*" } | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $OpenJDKD.Version
        $URL = $OpenJDKD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download OracleJava8
    If ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        $PackageName = "OracleJava8_" + "$ArchitectureClear"
        $OracleJava8D = Get-EvergreenApp -Name OracleJava8 | Where-Object { $_.Architecture -eq "$ArchitectureClear" }
        $Version = $OracleJava8D.Version
        $URL = $OracleJava8D.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Remote Desktop Manager
    If ($RemoteDesktopManager -eq 1) {
        Switch ($RemoteDesktopManagerType) {
            0 {
                $Product = "RemoteDesktopManager Free"
                $PackageName = "Setup.RemoteDesktopManagerFree"
                $URLVersion = "https://remotedesktopmanager.com/de/release-notes/free"
                $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($URLVersion) -SessionVariable websession
                $regexAppVersion = "\d\d\d\d.\d.\d\d.\d+"
                $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
                $Version = $webVersion.Trim("</td>").Trim("</td>")
                $URL = "https://cdn.devolutions.net/download/Setup.RemoteDesktopManagerFree.$Version.msi"
                $InstallerType = "msi"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $CurrentVersion"
                If (!($CurrentVersion -eq $Version)) {
                    Write-Host -ForegroundColor Green "Update available"
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    Write-Host "Starting download of $Product $Version"
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    #Invoke-WebRequest -UseBasicParsing -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                    Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                    Write-Output ""
                }
                Else {
                    Write-Host -ForegroundColor Cyan "No new version available"
                    Write-Output ""
                }
            }
            1 {
                $Product = "RemoteDesktopManager Enterprise"
                $PackageName = "Setup.RemoteDesktopManagerEnterprise"
                $URLVersion = "https://remotedesktopmanager.com/de/release-notes"
                $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($URLVersion) -SessionVariable websession
                $regexAppVersion = "\d\d\d\d.\d.\d\d.\d+"
                $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
                $Version = $webVersion.Trim("</td>").Trim("</td>")
                $URL = "https://cdn.devolutions.net/download/Setup.RemoteDesktopManager.$Version.msi"
                $InstallerType = "msi"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $CurrentVersion"
                If (!($CurrentVersion -eq $Version)) {
                    Write-Host -ForegroundColor Green "Update available"
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    Write-Host "Starting download of $Product $Version"
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    #Invoke-WebRequest -UseBasicParsing -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                    Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                    Write-Output ""
                }
                Else {
                    Write-Host -ForegroundColor Cyan "No new version available"
                    Write-Output ""
                }
            }
        }
    }

    #// Mark: Download ShareX
    If ($ShareX -eq 1) {
        $Product = "ShareX"
        $PackageName = "ShareX-setup"
        $ShareXD = Get-EvergreenApp -Name ShareX | Where-Object {$_.Type -eq "exe"}
        $Version = $ShareXD.Version
        $URL = $ShareXD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Slack
    If ($Slack -eq 1) {
        $Product = "Slack"
        $PackageName = "Slack.setup" + "_$ArchitectureClear" + "_$SlackPlatformClear"
        $SlackD = Get-EvergreenApp -Name Slack | Where-Object {$_.Architecture -eq "$SlackArchitectureClear" -and $_.Platform -eq "$SlackPlatformClear" }
        $Version = $SlackD.Version
        $URL = $SlackD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$SlackArchitectureClear" + "_$SlackPlatformClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $SlackArchitectureClear $SlackPlatformClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $SlackArchitectureClear $SlackPlatformClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download TreeSize
    If ($TreeSize -eq 1) {
        Switch ($TreeSizeType) {
            0 {
                $Product = "TreeSize Free"
                $PackageName = "TreeSize_Free"
                $TreeSizeFreeD = Get-EvergreenApp -Name JamTreeSizeFree
                $Version = $TreeSizeFreeD.Version
                $URL = $TreeSizeFreeD.uri
                $InstallerType = "exe"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $CurrentVersion"
                If (!($CurrentVersion -eq $Version)) {
                    Write-Host -ForegroundColor Green "Update available"
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    Write-Host "Starting download of $Product $Version"
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                    Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                    Write-Output ""
                }
                Else {
                    Write-Host -ForegroundColor Cyan "No new version available"
                    Write-Output ""
                }
            }
            1 {
                $Product = "TreeSize Professional"
                $PackageName = "TreeSize_Professional"
                $TreeSizeProfD = Get-EvergreenApp -Name JamTreeSizeProfessional
                $Version = $TreeSizeProfD.Version
                $URL = $TreeSizeProfD.uri
                $InstallerType = "exe"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $CurrentVersion"
                If (!($CurrentVersion -eq $Version)) {
                    Write-Host -ForegroundColor Green "Update available"
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    Write-Host "Starting download of $Product $Version"
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                    Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                    Write-Output ""
                }
                Else {
                    Write-Host -ForegroundColor Cyan "No new version available"
                    Write-Output ""
                }
            }
        }
    }

    #// Mark: Download VLC Player
    If ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        $PackageName = "VLC-Player_" + "$ArchitectureClear"
        $VLCD = Get-EvergreenApp -Name VideoLanVlcPlayer | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "$ArchitectureClear" -and $_.Type -eq "MSI" }
        $Version = $VLCD.Version
        $URL = $VLCD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download VMWareTools
    If ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"
        $PackageName = "VMWareTools_" + "$ArchitectureClear"
        $VMWareToolsD = Get-EvergreenApp -Name VMwareTools | Where-Object { $_.Architecture -eq "$ArchitectureClear" }
        $Version = $VMWareToolsD.Version
        $URL = $VMWareToolsD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Host "Starting download of $Product $ArchitectureClear $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download WinSCP
    If ($WinSCP -eq 1) {
        $Product = "WinSCP"
        $PackageName = "WinSCP"
        $WinSCPD = Get-EvergreenApp -Name WinSCP | Where-Object {$_.URI -like "*Setup*"}
        $Version = $WinSCPD.Version
        $URL = $WinSCPD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Zoom VDI Installer
    If ($Zoom -eq 1) {
        $Product = "Zoom VDI"
        $PackageName = "ZoomInstallerVDI"
        $ZoomVDI = Get-EvergreenApp -Name Zoom | Where-Object {$_.Platform -eq "VDI"}
        $URLVersion = "https://support.zoom.us/hc/en-us/articles/360041602711"
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($URLVersion) -SessionVariable websession
        $regexAppVersion = "(\d\.\d\.\d)"
        $Version = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Sort-Object -Descending | Select-Object -First 1
        $URL = $ZoomVDI.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        If (!($CurrentVersion -eq $Version)) {
            Write-Host -ForegroundColor Green "Update available"
            If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS | Out-Null
            Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            Write-Host "Starting download of $Product $Version"
            Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
            #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
            Write-Verbose "Stop logging"
            Stop-Transcript | Out-Null
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
        If ($ZoomCitrixClient -eq 1) {
            $Product2 = "Zoom Citrix Client"
            $PackageName2 = "ZoomCitrixHDXMediaPlugin"
            $ZoomCitrix = Get-EvergreenApp -Name Zoom | Where-Object {$_.Platform -eq "Citrix"}
            $URL = $ZoomCitrix.uri
            $Source2 = "$PackageName2" + "." + "$InstallerType"
            $CurrentVersion2 = Get-Content -Path "$PSScriptRoot\$Product2\Version.txt" -EA SilentlyContinue
            Write-Host -ForegroundColor Magenta "Download $Product2" -Verbose
            Write-Host "Download Version: $Version"
            Write-Host "Current Version: $CurrentVersion2"
            If (!($CurrentVersion2 -eq $Version)) {
                Write-Host -ForegroundColor Green "Update available"
                If (!(Test-Path -Path "$PSScriptRoot\$Product2")) {New-Item -Path "$PSScriptRoot\$Product2" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product2\" + "$Product2 $Version.log"
                Remove-Item "$PSScriptRoot\$Product2\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product2\Version.txt" -Value "$Version"
                Write-Host "Starting download of $Product2 $Version"
                Get-Download $URL "$PSScriptRoot\$Product2\" $Source2 -includeStats
                #Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product2\" + ($Source2))
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
    }
}

If ($download -eq $False) {

    Write-Host -ForegroundColor DarkGray "Starting installs..."
    Write-Output ""

    # Logging
    # Global variables
    # $StartDir = $PSScriptRoot # the directory path of the script currently being executed
    $LogDir = "$PSScriptRoot\_Install Logs"
    $LogFileName = ("$ENV:COMPUTERNAME - $Date.log")
    $LogFile = Join-path $LogDir $LogFileName
    $LogTemp = "$env:windir\Logs\Evergreen"

    # Create the log directories if they don't exist
    If (!(Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType directory | Out-Null }
    If (!(Test-Path $LogTemp)) { New-Item -Path $LogTemp -ItemType directory | Out-Null }

    # Create new log file (overwrite existing one)
    New-Item $LogFile -ItemType "file" -force | Out-Null
    DS_WriteLog "I" "START SCRIPT - " $LogFile
    DS_WriteLog "-" "" $LogFile

    # Install script part (AddScript)

    #// Mark: Install 7-ZIP
    If ($7ZIP -eq 1) {
        $Product = "7-Zip"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $SevenZip = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Select-Object -First 1
        If (!$SevenZip) {
            $SevenZip = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Select-Object -First 1
        }
        $7ZipInstaller = "7-Zip_" + "$ArchitectureClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $SevenZip"
        If ($SevenZip -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Start-Process "$PSScriptRoot\$Product\$7ZipInstaller" -ArgumentList /S
                $p = Get-Process 7-Zip_$ArchitectureClear
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Adobe Pro DC
    If ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Adobe"
        If ($Adobe -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                $mspArgs = "/P `"$PSScriptRoot\$Product\Adobe_Pro_DC_Update.msp`" /quiet /qn"
                $inst = Start-Process -FilePath msiexec.exe -ArgumentList $mspArgs -Wait
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                # Disable update service and scheduled task
                $Service = Get-Service -Name AdobeARMservice -ErrorAction SilentlyContinue
                If ($Service.Length -gt 0) {
                    Write-Host "Customize Service"
                    Stop-Service AdobeARMservice
                    Set-Service AdobeARMservice -StartupType Disabled
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                $ScheduledTask = Get-ScheduledTask -TaskName "Adobe Acrobat Update Task" -ErrorAction SilentlyContinue
                If ($ScheduledTask.Length -gt 0) {
                    Write-Host "Customize Scheduled Task"
                    Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
                    Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
                }
                Write-Host -ForegroundColor Green "Customize scripts $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error customizing (Error: $($Error[0]))"
                DS_WriteLog "E" "Error customizing (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Adobe Reader DC
    If ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$AdobeArchitectureClear" + "_$AdobeLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*" | Sort-Object -Property DisplayVersion | Select-Object -Last 1 }).DisplayVersion
        $AdobeReaderInstaller = "Adobe_Reader_DC_" + "$AdobeArchitectureClear" + "$AdobeLanguageClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $AdobeArchitectureClear $AdobeLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Adobe"
        If ($Adobe -ne $Version) {
            DS_WriteLog "I" "Installing $Product $AdobeArchitectureClear $AdobeLanguageClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/sAll"
                "/rs"
                "/msi EULA_ACCEPT=YES ENABLE_OPTIMIZATION=YES DISABLEDESKTOPSHORTCUT=1 UPDATE_MODE=0 DISABLE_ARM_SERVICE_INSTALL=1 DISABLE_CACHE=1 DISABLE_PDFMAKER=YES ALLUSERS=1"
            )
            Try {
                Write-Host "Starting install of $Product $AdobeArchitectureClear $AdobeLanguageClear $Version"
                Start-Process "$PSScriptRoot\$Product\$AdobeReaderInstaller" -ArgumentList $Options
                $p = Get-Process Adobe_Reader_DC_$AdobeArchitectureClear$AdobeLanguageClear
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $AdobeArchitectureClear $AdobeLanguageClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                # Disable update service and scheduled task
                $Service = Get-Service -Name AdobeARMservice -ErrorAction SilentlyContinue
                If ($Service.Length -gt 0) {
                    Write-Host "Customize Service"
                    Stop-Service AdobeARMservice
                    Set-Service AdobeARMservice -StartupType Disabled
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                $ScheduledTask = Get-ScheduledTask -TaskName "Adobe Acrobat Update Task" -ErrorAction SilentlyContinue
                If ($ScheduledTask.Length -gt 0) {
                    Write-Host "Customize Scheduled Task"
                    Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
                    Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
                }
                Write-Host -ForegroundColor Green "Customize scripts $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error customizing (Error: $($Error[0]))"
                DS_WriteLog "E" "Error customizing (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install BIS-F
    If ($BISF -eq 1) {
        $Product = "BIS-F"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $BISF = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Base Image*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $BISFLog = "$LogTemp\BISF.log"
        $InstallMSI = "$PSScriptRoot\$Product\setup-BIS-F.msi"
        Write-Host -ForegroundColor Magenta "Install $Product"
        If ($BISF) {$BISF = $BISF -replace ".{6}$"}
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $BISF"
        If ($BISF -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $BISFLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $BISFLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $BISFLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            # Customize scripts, it's best practise to enable Task Offload and RSS and to disable DEP
            $BISFDir = "C:\Program Files (x86)\Base Image Script Framework (BIS-F)\Framework\SubCall"
            If (Test-Path -Path "$BISFDir") {
                Try {
                    Write-Host "Customize scripts $Product"
                    ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace "DisableTaskOffload' -Value '1'","DisableTaskOffload' -Value '0'") | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
                    ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'nx AlwaysOff','nx OptOut') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
                    ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'rss=disable','rss=enable') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
                    Write-Host -ForegroundColor Green "Customize scripts $Product finished!"
                } Catch {
                    Write-Host -ForegroundColor Red "Error when customizing scripts (Error: $($Error[0]))"
                    DS_WriteLog "E" "Error when customizing scripts (Error: $($Error[0]))" $LogFile
                }
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Citrix Hypervisor Tools
    If ($Citrix_Hypervisor_Tools -eq 1) {
        $Product = "Citrix Hypervisor Tools"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $HypTools = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Hypervisor*"}).DisplayVersion
        $CitrixHypLog = "$LogTemp\CitrixHypervisor.log"
        $HypToolsInstaller = "managementagent" + "$ArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\Citrix\$Product\$HypToolsInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        If (!$HypTools) {
            $HypTools = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Hypervisor*"}).DisplayVersion
        }
        If ($HypTools) {$HypTools = $HypTools.Insert(3,'.0')}
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $HypTools"
        If ($HypTools -ne $Version) {
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/quiet"
                "/norestart"
                "/L*V $CitrixHypLog"
                )
            Try {
                Write-Host "Starting install of $Product $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $CitrixHypLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $CitrixHypLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Citrix WorkspaceApp
    If ($Citrix_WorkspaceApp -eq 1) {
        $Product = "Citrix WorkspaceApp $CitrixWorkspaceAppReleaseClear"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt"
        $WSA = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Workspace*" -and $_.UninstallString -like "*Trolley*"}).DisplayVersion
        $UninstallWSACR = "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $WSA"
        If ($WSA -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            # Citrix WSA Uninstallation
            Write-Host "Uninstall Citrix Workspace App / Receiver"
            DS_WriteLog "I" "Uninstall Citrix Workspace App / Receiver" $LogFile
            Try {
                Start-process $UninstallWSACR -ArgumentList '/silent /disableCEIP' -NoNewWindow -Wait
                Write-Host -ForegroundColor Green "Uninstall Citrix Workspace App / Receiver finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error uninstalling Citrix Workspace App / Receiver (Error: $($Error[0]))"
                DS_WriteLog "E" "Error Uninstalling Citrix Workspace App / Receiver (Error: $($Error[0]))" $LogFile
            }
            # Citrix WSA Installation
            $Options = @(
                "/forceinstall"
                "/silent"
                "/EnableCEIP=false"
                "/FORCE_LAA=1"
                "/AutoUpdateCheck=disabled"
                "/ALLOWADDSTORE=S"
                "/ALLOWSAVEPWD=S"
                "/includeSSON"
                "/ENABLE_SSON=Yes"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\Citrix\$Product\CitrixWorkspaceApp.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                Write-Host "Customize $Product"
                reg add "HKLM\SOFTWARE\Wow6432Node\Policies\Citrix" /v EnableX1FTU /t REG_DWORD /d 0 /f | Out-Null
                reg add "HKCU\Software\Citrix\Splashscreen" /v SplashscrrenShown /d 1 /f | Out-Null
                reg add "HKLM\SOFTWARE\Policies\Citrix" /f /v EnableFTU /t REG_DWORD /d 0 | Out-Null
                Write-Host -ForegroundColor Green "Customizing $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Host -ForegroundColor Yellow "System needs to reboot after installation!"
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install deviceTRUST
    If ($deviceTRUST -eq 1) {
        $Product = "deviceTRUST"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + "_$ArchitectureClear"+ ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $deviceTRUSTClientV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Client*"}).DisplayVersion
        If (!$deviceTRUSTClientV) {
            $deviceTRUSTClientV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Client*"}).DisplayVersion
        }
        If ($deviceTRUSTClientV.length -ne "8") {$deviceTRUSTClientV = $deviceTRUSTClientV -replace ".{2}$"}
        $deviceTRUSTHostV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Host*"}).DisplayVersion
        If (!$deviceTRUSTHostV) {
            $deviceTRUSTHostV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Host*"}).DisplayVersion
        }
        If ($deviceTRUSTHostV.length -ne "8") {$deviceTRUSTHostV = $deviceTRUSTHostV -replace ".{2}$"}
        $deviceTRUSTConsoleV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Console*"}).DisplayVersion
        If (!$deviceTRUSTConsoleV) {
            $deviceTRUSTConsoleV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Console*"}).DisplayVersion
        }
        If ($deviceTRUSTConsoleV.length -ne "8") {$deviceTRUSTConsoleV = $deviceTRUSTConsoleV -replace ".{2}$"}
        $deviceTRUSTLog = "$LogTemp\deviceTRUST.log"
        $deviceTRUSTClientInstaller = "dtclient-release" + ".exe"
        $deviceTRUSTHostInstaller = "dthost-" + "$ArchitectureClear" + "-release" + ".msi"
        $deviceTRUSTConsoleInstaller = "dtconsole-" + "$ArchitectureClear" + "-release" + ".msi"
        $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/passive"
                "/quiet"
                "/norestart"
                "/L*V $deviceTRUSTLog"
                )
        If ($deviceTRUSTClient -eq $True) {
            Write-Host -ForegroundColor Magenta "Install $Product Client"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version: $deviceTRUSTClientV"
            If ($deviceTRUSTClientV -ne $Version) {
                # deviceTRUST Client
                DS_WriteLog "I" "Installing $Product Client" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                Try {
                    $Options = @(
                        "/INSTALL"
                        "/QUIET"
                    )
                    Write-Host "Starting install of $Product Client $Version"
                    Start-Process -FilePath "$PSScriptRoot\$Product\$deviceTRUSTClientInstaller" -ArgumentList $Options -PassThru -Wait -ErrorAction Stop | Out-Null
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                } Catch {
                    Write-Host -ForegroundColor Red "Error installing $Product Client (Error: $($Error[0]))"
                    DS_WriteLog "E" "Error installing $Product Client (Error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
            # Stop, if no new version is available
            Else {
                Write-Host -ForegroundColor Cyan "No update available for $Product Client"
                Write-Output ""
            }
        }
        If ($deviceTRUSTHost -eq $True) {
            Write-Host -ForegroundColor Magenta "Install $Product Host"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version: $deviceTRUSTHostV"
            If ($deviceTRUSTHostV -ne $Version) {
                # deviceTRUST Host
                DS_WriteLog "I" "Installing $Product Host" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $InstallMSI = "$PSScriptRoot\$Product\$deviceTRUSTHostInstaller"
                Try {
                    Write-Host "Starting install of $Product Host $Version"
                    Install-MSI $InstallMSI $Arguments
                    Get-Content $deviceTRUSTLog | Add-Content $LogFile -Encoding ASCI
                    Remove-Item $deviceTRUSTLog
                } Catch {
                    DS_WriteLog "E" "Error installing $Product Host (Error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
            # Stop, if no new version is available
            Else {
                Write-Host -ForegroundColor Cyan "No update available for $Product Host"
                Write-Output ""
            }
        }
        If ($deviceTRUSTConsole -eq $True) {
            Write-Host -ForegroundColor Magenta "Install $Product Console"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version: $deviceTRUSTConsoleV"
            If ($deviceTRUSTConsoleV -ne $Version) {
                # deviceTRUST Console
                DS_WriteLog "I" "Installing $Product Console" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $InstallMSI = "$PSScriptRoot\$Product\$deviceTRUSTConsoleInstaller"
                Try {
                    Write-Host "Starting install of $Product Console $Version"
                    Install-MSI $InstallMSI $Arguments
                    Get-Content $deviceTRUSTLog | Add-Content $LogFile -Encoding ASCI
                    Remove-Item $deviceTRUSTLog
                } Catch {
                    DS_WriteLog "E" "Error installing $Product Console (Error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
            # Stop, if no new version is available
            Else {
                Write-Host -ForegroundColor Cyan "No update available for $Product Console"
                Write-Output ""
            }
        }
    }

    #// Mark: Install Filezilla
    If ($Filezilla -eq 1) {
        $Product = "Filezilla"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Filezilla = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Filezilla*"}).DisplayVersion
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Filezilla"
        If ($Filezilla -ne $Version) {
            $Options = @(
                "/S"
                "/user=all"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\Filezilla-win64.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Foxit Reader
    If ($Foxit_Reader -eq 1) {
        $Product = "Foxit Reader"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitReaderLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $FReader = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Foxit Reader*"}).DisplayVersion
        $FoxitReaderInstaller = "FoxitReader-Setup-" + "$FoxitReaderLanguageClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $FoxitReaderLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $FReader"
        If ($FReader -ne $Version) {
            $Options = @(
                "/FORCEINSTALL"
                "/VERYSILENT"
                "/PASSIVE"
                "/ALLUSERS"
                "/NORESTART"
                "/NOCLOSEAPPLICATIONS"
                "AUTO_UPDATE=0"
                "LAUNCHCHECKDEFAULT=0"
                "DESKTOP_SHORTCUT=0"
                "/qn"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $FoxitReaderLanguageClear $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$FoxitReaderInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    If (Test-Path -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk" -Force}
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $FoxitReaderLanguageClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Google Chrome
    If ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Chrome = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion
        $ChromeLog = "$LogTemp\GoogleChrome.log"
        If (!$Chrome) {
            $Chrome = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion
        }
        $ChromeInstaller = "googlechromestandaloneenterprise_" + "$ArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$ChromeInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Chrome"
        If ($Chrome -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $ChromeLog"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $ChromeLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $ChromeLog
                If (Test-Path -Path "$env:PUBLIC\Desktop\Google Chrome.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Google Chrome.lnk" -Force}
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                # Disable update service and scheduled task
                $Service = Get-Service -Name gupdate -ErrorAction SilentlyContinue
                If ($Service.Length -gt 0) {
                    Write-Host "Customize Service"
                    Stop-Service gupdate
                    Set-Service gupdate -StartupType Disabled
                    Stop-Service gupdatem
                    Set-Service gupdatem -StartupType Disabled
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                $ScheduledTask = Get-ScheduledTask -TaskName "GoogleUpdateTaskMachineCore" -ErrorAction SilentlyContinue
                If ($ScheduledTask.Length -gt 0) {
                    Write-Host "Customize Scheduled Task"
                    Disable-ScheduledTask -TaskName "GoogleUpdateTaskMachineCore" | Out-Null
                    Disable-ScheduledTask -TaskName "GoogleUpdateTaskMachineUA" | Out-Null
                    #Disable-ScheduledTask -TaskName "GPUpdate on Startup" | Out-Null
                    Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
                }
                Write-Host -ForegroundColor Green "Customize scripts $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error customizing (Error: $($Error[0]))"
                DS_WriteLog "E" "Error customizing (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Greenshot
    If ($Greenshot -eq 1) {
        $Product = "Greenshot"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Greenshot = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Greenshot*"}).DisplayVersion
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Greenshot"
        If ($Greenshot -ne $Version) {
            $Options = @(
                "/VERYSILENT"
                "/NORESTART"
                "/SUPPRESSMSGBOXES"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\Greenshot-INSTALLER-x86.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install IrfanView
    If ($IrfanView -eq 1) {
        $Product = "IrfanView"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $IrfanViewV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*IrfanView*"}).DisplayVersion
        If (!$IrfanViewV) {
            $IrfanViewV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*IrfanView*"}).DisplayVersion
        }
        $IrfanViewInstaller = "IrfanView" + "$ArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $IrfanViewV"
        If ($IrfanViewV -ne $Version) {
            $Options = @(
                "/assoc=1"
                "/group=1"
                "/ini=%APPDATA%\IrfanView"
                "/silent"
                "/allusers=1"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$IrfanViewInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install KeePass
    If ($KeePass -eq 1) {
        $Product = "KeePass"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $KeePassV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*KeePass*"}).DisplayVersion
        If ($KeePassV) {$KeePassV = $KeePassV -replace ".{2}$"}
        $KeePassLog = "$LogTemp\KeePass.log"
        $InstallMSI = "$PSScriptRoot\$Product\KeePass.msi"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $KeePassV"
        If ($KeePassV -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $KeePassLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $KeePassLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $KeePassLog
                If (Test-Path -Path "$env:PUBLIC\Desktop\KeePass.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\KeePass.lnk" -Force}
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft .Net Framework
    If ($MSDotNetFramework -eq 1) {
        $Product = "Microsoft Dot Net Framework"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSDotNetFrameworkChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $MSDotNetFrameworkV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Windows Desktop Runtime*" -and $_.URLInfoAbout -like "https://dot.net/core"}).DisplayVersion | Select-Object -First 1
        If (!$MSDotNetFrameworkV) {
            $MSDotNetFrameworkV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Windows Desktop Runtime*" -and $_.URLInfoAbout -like "https://dot.net/core"}).DisplayVersion | Select-Object -First 1
        }
        $MSDotNetFrameworkInstaller = "NetFramework-runtime_" + "$ArchitectureClear" + "_$MSDotNetFrameworkChannelClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear $MSDotNetFrameworkChannelClear Channel"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $MSDotNetFrameworkV"
        If ($MSDotNetFrameworkV -ne $Version) {
            $Options = @(
                "/q"
                "/norestart"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $MSDotNetFrameworkChannelClear Channel $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$MSDotNetFrameworkInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear $MSDotNetFrameworkChannelClear Channel (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Apps 365
    If ($MS365Apps -eq 1) {
        $Product = "Microsoft 365 Apps"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt"
        $MS365AppsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft 365 Apps*"}).DisplayVersion
        If (!$MS365AppsV) {
            $MS365AppsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft 365 Apps*"}).DisplayVersion
        }
        $MS365AppsInstaller = "setup_" + "$MS365AppsChannelClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $MS365AppsChannelClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $MS365AppsV"
        If ($MS365AppsV -ne $Version) {
            Write-Host -ForegroundColor Green "Update available"
            # MS365Apps Uninstallation
            $Options = @(
                "/configure remove.xml"
            )
            Write-Host "Uninstall Microsoft Office 2019 or Microsoft 365 Apps"
            DS_WriteLog "I" "Uninstall Microsoft Office 2019 or Microsoft 365 Apps" $LogFile
            Try {
                set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                Start-Process -FilePath ".\$MS365AppsInstaller" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Host -ForegroundColor Green "Uninstall Microsoft Office 2019 or Microsoft 365 Apps finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error uninstalling Microsoft Office 2019 or Microsoft 365 Apps (Error: $($Error[0]))"
                DS_WriteLog "E" "Error uninstalling Microsoft Office 2019 or Microsoft 365 Apps (Error: $($Error[0]))" $LogFile
            }
            # MS365Apps Installation
            $Options = @(
                "/configure install.xml"
            )
            Try {
                DS_WriteLog "I" "Install $Product" $LogFile
                Write-Host "Starting install of $Product $Version"
                set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                Start-Process -FilePath ".\$MS365AppsInstaller" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Edge
    If ($MSEdge -eq 1) {
        $Product = "Microsoft Edge"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSEdgeChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Edge = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion
        $EdgeLog = "$LogTemp\MSEdge.log"
        If (!$Edge) {
            $Edge = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion
        }
        $EdgeInstaller = "MicrosoftEdgeEnterprise_" + "$ArchitectureClear" + "_$MSEdgeChannelClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$EdgeInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $MSEdgeChannelClear $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Edge"
        If ($Edge -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "REBOOT=ReallySuppress"
                "DONOTCREATEDESKTOPSHORTCUT=TRUE"
                "DONOTCREATETASKBARSHORTCUT=true"
                "/L*V $EdgeLog"
            )
            try {
                Write-Host "Starting install of $Product $MSEdgeChannelClear $ArchitectureClear $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $EdgeLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $EdgeLog
            } catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Write-Host "Customize $Product"
            Try {
                # Disable Microsoft Edge auto update
                Write-Host "Customize $Product registry"
                If (!(Test-Path -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate)) {
                    New-Item -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate | Out-Null
                    New-ItemProperty -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -Name UpdateDefault -Value 0 -PropertyType DWORD | Out-Null
                }
                Else {
                    $EdgeUpdateState = Get-ItemProperty -path "HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate" | Select-Object -Expandproperty "UpdateDefault"
                    If ($EdgeUpdateState -ne "0") {New-ItemProperty -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -Name UpdateDefault -Value 0 -PropertyType DWORD | Out-Null}
                }
                # Disable Citrix API Hooks (MS Edge) on Citrix VDA
                $(
                    $RegPath = "HKLM:SYSTEM\CurrentControlSet\services\CtxUvi"
                    If (Test-Path $RegPath) {
                        $RegName = "UviProcessExcludes"
                        $EdgeRegvalue = "msedge.exe"
                        # Get current values in UviProcessExcludes
                        $CurrentValues = Get-ItemProperty -Path $RegPath | Select-Object -ExpandProperty $RegName
                        # Add the msedge.exe value to existing values in UviProcessExcludes
                        Set-ItemProperty -Path $RegPath -Name $RegName -Value "$CurrentValues$EdgeRegvalue;"
                    }
                ) | Out-Null
                Write-Host -ForegroundColor Green "Customize $Product registry finished!"
                $Service = Get-Service -Name edgeupdate -ErrorAction SilentlyContinue
                If ($Service.Length -gt 0) {
                    Write-Host "Customize Service"
                    Stop-Service edgeupdate
                    Set-Service -Name edgeupdate -StartupType Manual
                    Stop-Service edgeupdatem
                    Set-Service -Name edgeupdatem -StartupType Manual
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                Write-Host "Customize Scheduled Task"
                Start-Sleep -s 5
                Get-ScheduledTask -TaskName MicrosoftEdgeUpdate* | Disable-ScheduledTask | Out-Null
                Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
                Write-Host -ForegroundColor Green "Customize $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error customizing (Error: $($Error[0]))"
                DS_WriteLog "E" "Error customizing (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft FSLogix
    If ($FSLogix -eq 1) {
        $Product = "Microsoft FSLogix"
        $OS = (Get-WmiObject Win32_OperatingSystem).Caption
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Install\Version.txt"
        $FSLogixV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).DisplayVersion
        If (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}) {
            $UninstallFSL = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).UninstallString.replace("/uninstall","")
        }
        If (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps RuleEditor"}) {
            $UninstallFSLRE = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps RuleEditor"}).UninstallString.replace("/uninstall","")
        }
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $FSLogixV"
        If ($FSLogixV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            # FSLogix Uninstall
            If (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}) {
                Write-Host "Uninstall $Product"
                DS_WriteLog "I" "Uninstall $Product" $LogFile
                Try {
                    Start-process $UninstallFSL -ArgumentList '/uninstall /quiet /norestart' -NoNewWindow -Wait
                    Start-process $UninstallFSLRE -ArgumentList '/uninstall /quiet /norestart' -NoNewWindow -Wait
                    Write-Host -ForegroundColor Green "Uninstall $Product finished!"
                } Catch {
                    Write-Host -ForegroundColor Red "Error uninstalling $Product (Error: $($Error[0]))"
                    DS_WriteLog "E" "Error uninstalling $Product (Error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
                Write-Host -ForegroundColor Red "Server needs to reboot, start script again after reboot"
                Write-Output ""
                Write-Host -ForegroundColor Red "Hit any key to reboot server"
                Read-Host
                Restart-Computer
            }
            # FSLogix Install
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Start-Process "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.exe" -ArgumentList '/install /norestart /quiet' -NoNewWindow -Wait
                Write-Host -ForegroundColor Green "Install $Product finished!"
                Write-Host "Starting install of $Product Rule Editor $ArchitectureClear $Version"
                Start-Process "$PSScriptRoot\$Product\Install\FSLogixAppsRuleEditorSetup.exe" -ArgumentList '/install /norestart /quiet' -NoNewWindow -Wait
                Write-Host -ForegroundColor Green "Install $Product Rule Editor $ArchitectureClear finished!"
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                Start-Sleep -s 20
                # Application post deployment tasks (Thx to Kasper https://github.com/kaspersmjohansen)
                Write-Host "Applying $Product post setup customizations"
                Write-Host "Post setup customizations for $OS"
                If ($OS -Like "*Windows Server 2019*" -or $OS -eq "Microsoft Windows 10 Enterprise for Virtual Desktops") {
                    If ((Test-RegistryValue2 -Path "HKLM:SOFTWARE\FSLogix\Apps" -Value "RoamSearch") -ne $true) {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "0" -Type DWORD | Out-Null
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                    If ((Get-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" | Select-Object -ExpandProperty "RoamSearch") -ne "0") {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        Set-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "0" -Type DWORD
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                }
                If ($OS -Like "*Windows 10*" -and $OS -ne "Microsoft Windows 10 Enterprise for Virtual Desktops") {
                    If ((Test-RegistryValue2 -Path "HKLM:SOFTWARE\FSLogix\Apps" -Value "RoamSearch") -ne $true) {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "1" -Type DWORD | Out-Null
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                    If ((Get-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" | Select-Object -ExpandProperty "RoamSearch") -ne "1") {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        Set-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "1" -Type DWORD
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                }
                Write-Host -ForegroundColor Green "Post setup customizations for $OS finished!"
                # Implement user based group policy processing fix
                If (!(Test-Path -Path HKLM:SOFTWARE\FSLogix\Profiles)) {
                    New-Item -Path "HKLM:SOFTWARE\FSLogix" -Name Profiles | Out-Null
                }
                If ((Test-RegistryValue -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Value "GroupPolicyState") -ne $true) {
                    Write-Host "Deactivate FSLogix GroupPolicy"
                    New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Name "GroupPolicyState" -Value "0" -Type DWORD | Out-Null
                    Write-Host -ForegroundColor Green "Deactivate FSLogix GroupPolicy finished!"
                }
                If ((Get-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" | Select-Object -ExpandProperty "GroupPolicyState") -ne "0") {
                    Write-Host "Deactivate FSLogix GroupPolicy"
                    Set-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Name "GroupPolicyState" -Value "0" -Type DWORD
                    Write-Host -ForegroundColor Green "Deactivate FSLogix GroupPolicy finished!"
                }
                If (!(Get-ScheduledTask -TaskName "Restart Windows Search Service on Event ID 2")) {
                    Write-Host "Implement scheduled task to restart Windows Search service on Event ID 2"
                    # Implement scheduled task to restart Windows Search service on Event ID 2
                    # Define CIM object variables
                    # This is needed for accessing the non-default trigger settings when creating a schedule task using Powershell
                    $Class = Get-CimClass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler
                    $Trigger = $class | New-CimInstance -ClientOnly
                    $Trigger.Enabled = $true
                    $Trigger.Subscription = "<QueryList><Query Id=`"0`" Path=`"Application`"><Select Path=`"Application`">*[System[Provider[@Name='Microsoft-Windows-Search-ProfileNotify'] and EventID=2]]</Select></Query></QueryList>"
                    # Define additional variables containing scheduled task action and scheduled task principal
                    $A = New-ScheduledTaskAction -Execute powershell.exe -Argument "Restart-Service Wsearch"
                    $P = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount
                    $S = New-ScheduledTaskSettingsSet
                    # Cook it all up and create the scheduled task
                    $RegSchTaskParameters = @{
                        TaskName    = "Restart Windows Search Service on Event ID 2"
                        Description = "Restarts the Windows Search service on event ID 2 - Workaround described here - https://virtualwarlock.net/how-to-install-the-fslogix-apps-agent/#Windows_Search_Roaming_workaround_1"
                        TaskPath    = "\"
                        Action      = $A
                        Principal   = $P
                        Settings    = $S
                        Trigger     = $Trigger
                    }
                    Register-ScheduledTask @RegSchTaskParameters
                    Write-Host -ForegroundColor Green "Implement scheduled task to restart Windows Search service on Event ID 2 finished!"
                }
                Write-Host -ForegroundColor Green "Applying $Product post setup customizations finished!"
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Office 2019
    If ($MSOffice2019 -eq 1) {
        $Product = "Microsoft Office 2019"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $MSOffice2019V = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office*"}).DisplayVersion
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $MSOffice2019V"
        If ($MSOffice2019V -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            # MS Office 2019 Uninstallation
            $Options = @(
                "/configure remove.xml"
            )
            Write-Host "Uninstall Microsoft Office 2019 or Microsoft 365 Apps"
            DS_WriteLog "I" "Uninstall Microsoft Office 2019 or Microsoft 365 Apps" $LogFile
            Try {
                set-location $PSScriptRoot\$Product
                Start-Process -FilePath ".\setup.exe" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Host -ForegroundColor Green "Uninstall Microsoft Office 2019 or Microsoft 365 Apps finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error uninstalling Microsoft Office 2019 or Microsoft 365 Apps (Error: $($Error[0]))"
                DS_WriteLog "E" "Error uninstalling Microsoft Office 2019 or Microsoft 365 Apps (Error: $($Error[0]))" $LogFile
            }
            # MS Office 2019 Installation
            $Options = @(
                "/configure install.xml"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host "Starting install of $Product $Version"
            Try {
                set-location $PSScriptRoot\$Product
                Start-Process -FilePath ".\setup.exe" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft OneDrive
    If ($MSOneDrive -eq 1) {
        $Product = "Microsoft OneDrive"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $MSOneDriveV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion
        If (!$MSOneDriveV) {
            $MSOneDriveV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion
        }
        $OneDriveInstaller = "OneDriveSetup-" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear" + ".exe"
        $OneDriveProcess = "OneDriveSetup-" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $MSOneDriveV"
        If ($MSOneDriveV -ne $Version) {
            Write-Host -ForegroundColor Green "Update available"
            DS_WriteLog "I" "Install $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear" $LogFile
            $Options = @(
                "/ALLUSERS"
                "/SILENT"
            )
            Try {
                Write-Host "Starting install of $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear $Version"
                $null = Start-Process "$PSScriptRoot\$Product\$OneDriveInstaller" -ArgumentList $Options -NoNewWindow -PassThru
                while (Get-Process -Name $OneDriveProcess -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                # OneDrive starts automatically after setup. kill!
                #Stop-Process -Name "OneDrive" -Force
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft PowerShell
    If ($MSPowerShell -eq 1) {
        $Product = "Microsoft PowerShell"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSPowerShellReleaseClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $MSPowerShellV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PowerShell*"}).DisplayVersion
        If (!$MSPowerShellV) {
            $MSPowerShellV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PowerShell*"}).DisplayVersion
        }
        If ($MSPowerShellV) {$MSPowerShellV = $MSPowerShellV -replace ".{2}$"}
        $MSPowerShellLog = "$LogTemp\MSPowerShell.log"
        $MSPowerShellInstaller = "PowerShell" + "$ArchitectureClear" + "_$MSPowerShellReleaseClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$MSPowerShellInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear $MSPowerShellReleaseClear Release"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $MSPowerShellV"
        If ($MSPowerShellV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/norestart"
                "/L*V $MSPowerShellLog"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $MSPowerShellReleaseClear Release $Version"
                Install-MSI $InstallMSI $Arguments
                Start-Sleep 25
                Get-Content $MSPowerShellLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $MSPowerShellLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Teams
    If ($MSTeams -eq 1) {
        $Product = "Microsoft Teams"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Teams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).DisplayVersion
        $TeamsInstaller = "Teams_" + "$ArchitectureClear" + "_$MSTeamsRingClear" + ".msi"
        $TeamsLog = "$LogTemp\MSTeams.log"
        $InstallMSI = "$PSScriptRoot\$Product\$TeamsInstaller"
        If ($Teams) {$Teams = $Teams.Insert(5,'0')}
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear $MSTeamsRingClear Ring"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Teams"
        If ($Teams -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            #Uninstalling MS Teams
            If ($Teams) {
                Write-Host "Uninstall $Product"
                DS_WriteLog "I" "Uninstall $Product" $LogFile
                Try {
                    $UninstallTeams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).UninstallString
                    $UninstallTeams = $UninstallTeams -Replace("MsiExec.exe /I","")
                    Start-Process -FilePath msiexec.exe -ArgumentList "/X $UninstallTeams /qn /L*V $TeamsLog"
                    Start-Sleep 20
                    Get-Content $TeamsLog | Add-Content $LogFile -Encoding ASCI
                    Remove-Item $TeamsLog
                    Write-Host -ForegroundColor Green "Uninstall $Product finished!" -Verbose
                    DS_WriteLog "I" "Uninstall $Product finished!" $LogFile
                } Catch {
                    Write-Host -ForegroundColor Red "Error uninstalling $Product (Error: $($Error[0]))"
                    DS_WriteLog "E" "Error uninstalling $Product (Error: $($Error[0]))" $LogFile       
                }
            }
            DS_WriteLog "-" "" $LogFile
            #MS Teams Installation
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "REBOOT=ReallySuppress"
                "ALLUSER=1"
                "ALLUSERS=1"
                "OPTIONS='noAutoStart=true'"
                "/qn"
                "/L*V $TeamsLog"
            )
            #Registry key for Teams machine-based install with Citrix VDA (Thx to Kasper https://github.com/kaspersmjohansen)
            If (!(Test-Path 'HKLM:\Software\Citrix\PortICA\')) {
                Write-Host "Customize System for $Product Machine-Based Install"
                If (!(Test-Path 'HKLM:\Software\Citrix\')) {New-Item -Path "HKLM:Software\Citrix" | Out-Null}
                New-Item -Path "HKLM:Software\Citrix\PortICA" | Out-Null
                Write-Host -ForegroundColor Green "Customize System for $Product Machine-Based Install finished!"
            }
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $MSTeamsRingClear Ring $Version"
                Install-MSI $InstallMSI $Arguments
                Start-Sleep 5
                Get-Content $TeamsLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $TeamsLog
                #Remove public desktop shortcut (Thx to Kasper https://github.com/kaspersmjohansen)
                If (Test-Path "$env:PUBLIC\Desktop\Microsoft Teams.lnk") {
                    Remove-Item -Path "$env:PUBLIC\Desktop\Microsoft Teams.lnk" -Force
                }
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                Write-Host "Customize $Product"
                reg add "HKLM\SOFTWARE\Citrix\CtxHook\AppInit_Dlls\SfrHook" /v Teams.exe /t REG_DWORD /d 204 /f | Out-Null
                If ($MSTeamsNoAutoStart -eq 1) {
                    #Prevents MS Teams from starting at logon, better do this with WEM or similar
                    Write-Host "Customize $Product Autorun"
                    If (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run") {
                        If (Test-RegistryValue2 -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Value "Teams") {
                            Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -Force
                        }
                    }
                    Write-Host -ForegroundColor Green "Customize $Product Autorun finished!"
                }
                Write-Host "Register $Product Add-In for Outlook"
                # Register Teams add-in for Outlook - https://microsoftteams.uservoice.com/forums/555103-public/suggestions/38846044-fix-the-teams-meeting-addin-for-outlook
                $appDLLs = (Get-ChildItem -Path "${Env:ProgramFiles(x86)}\Microsoft\TeamsMeetingAddin" -Include "Microsoft.Teams.AddinLoader.dll" -Recurse).FullName
                $appX64DLL = $appDLLs[0]
                $appX86DLL = $appDLLs[1]
                Start-Process -FilePath "$env:WinDir\SysWOW64\regsvr32.exe" -ArgumentList "/s /n /i:user `"$appX64DLL`""
                Start-Process -FilePath "$env:WinDir\SysWOW64\regsvr32.exe" -ArgumentList "/s /n /i:user `"$appX86DLL`""
                Write-Host -ForegroundColor Green "Register $Product Add-In for Outlook finished!"
                Write-Host -ForegroundColor Green "Customize $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error when customizing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error when customizing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Mozilla Firefox
    If ($Firefox -eq 1) {
        $Product = "Mozilla Firefox"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "$FFLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $FirefoxV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Firefox*"}).DisplayVersion
        $FirefoxLog = "$LogTemp\Firefox.log"
        If (!$FirefoxV) {
            $FirefoxV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Firefox*"}).DisplayVersion
        }
        $FirefoxInstaller = "Firefox_Setup_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "_$FFLanguageClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$FirefoxInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $FirefoxV"
        If ($FirefoxV -ne $Version) {
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/q"
                "DESKTOP_SHORTCUT=false"
                "TASKBAR_SHORTCUT=false"
                "INSTALL_MAINTENANCE_SERVICE=false"
                "PREVENT_REBOOT_REQUIRED=true"
                "/L*V $FirefoxLog"
            )
            DS_WriteLog "I" "Install $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $FirefoxLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $FirefoxLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install mRemoteNG
    If ($mRemoteNG -eq 1) {
        $Product = "mRemoteNG"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $mRemoteNGV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "mRemoteNG"}).DisplayVersion
        $mRemoteLog = "$LogTemp\mRemote.log"
        If ($mRemoteNGV) {$mRemoteNGV = $mRemoteNGV -replace ".{6}$"}
        $InstallMSI = "$PSScriptRoot\$Product\mRemoteNG.msi"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $mRemoteNGV"
        If ($mRemoteNGV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $mRemoteLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $mRemoteLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $mRemoteLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            If (Test-Path -Path "$env:PUBLIC\Desktop\mRemoteNG.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\mRemoteNG.lnk" -Force}
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Notepad ++
    If ($NotePadPlusPlus -eq 1) {
        $Product = "NotepadPlusPlus"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Notepad = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion
        If (!$Notepad) {
            $Notepad = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion
        }
        $NotepadPlusPlusInstaller = "NotePadPlusPlus_" + "$ArchitectureClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $Notepad"
        If ($Notepad -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Start-Process "$PSScriptRoot\$Product\$NotepadPlusPlusInstaller" -ArgumentList /S -NoNewWindow
                $p = Get-Process NotePadPlusPlus_$ArchitectureClear
		        If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install OpenJDK
    If ($OpenJDK -eq 1) {
        $Product = "open JDK"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $OpenJDKV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion
        $openJDKLog = "$LogTemp\OpenJDK.log"
        If (!$OpenJDKV) {
            $OpenJDKV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion
        }
        $OpenJDKInstaller = "OpenJDK" + "$ArchitectureClear" + ".msi"
        If ($Version) {$Version = $Version -replace ".-"}
        $InstallMSI = "$PSScriptRoot\$Product\$OpenJDKInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $OpenJDKV"
        If ($OpenJDKV -ne $Version) {
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "INSTALLLEVEL=3"
                "UPDATE_NOTIFIER=0"
                "/L*V $openJDKLog"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Install-MSI $InstallMSI $Arguments
                Start-Sleep 25
                Get-Content $openJDKLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $openJDKLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install OracleJava8
    If ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        If ($Version) {$Version = $Version -replace "^.{2}"}
        If ($Version) {$Version = $Version -replace "\."}
        If ($Version) {$Version = $Version -replace "_"}
        If ($Version) {$Version = $Version -replace "-b"}
        $OracleJava = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java 8*"}).DisplayVersion
        If (!$OracleJava) {
            $OracleJava = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java 8*"}).DisplayVersion
        }
        If ($OracleJava) {$OracleJava = $OracleJava -replace "\."}
        $OracleJavaInstaller = "OracleJava8_" + "$ArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $OracleJava"
        If ($OracleJava -ne $Version) {
            DS_WriteLog "I" "Install $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/s INSTALL_SILENT=Enable AUTO_UPDATE=Disable REBOOT=Disable SPONSORS=Disable REMOVEOUTOFDATEJRES=1 WEB_ANALYTICS=Disable"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Start-Process "$PSScriptRoot\$Product\$OracleJavaInstaller" -ArgumentList $Options -NoNewWindow
                $p = Get-Process OracleJava8_$ArchitectureClear
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Remote Desktop Manager
    If ($RemoteDesktopManager -eq 1) {
        Switch ($RemoteDesktopManagerType) {
            0 {
                $Product = "RemoteDesktopManager Free"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
                $RemoteDesktopManagerFree = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Remote Desktop Manager*"}).DisplayVersion
                $RemoteDesktopManagerLog = "$LogTemp\RemoteDesktopManager.log"
                $InstallMSI = "$PSScriptRoot\$Product\Setup.RemoteDesktopManagerFree.msi"
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $RemoteDesktopManagerFree"
                If ($RemoteDesktopManagerFree -ne $Version) {
                    DS_WriteLog "I" "Installing $Product" $LogFile
                    Write-Host -ForegroundColor Green "Update available"
                    $Arguments = @(
                        "/i"
                        "`"$InstallMSI`""
                        "/qn"
                        "/L*V $RemoteDesktopManagerLog"
                    )
                    Try {
                        Write-Host "Starting install of $Product $Version"
                        Install-MSI $InstallMSI $Arguments
                        Get-Content $RemoteDesktopManagerLog | Add-Content $LogFile -Encoding ASCI
                        Remove-Item $RemoteDesktopManagerLog
                    } Catch {
                        DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile       
                    }
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
                # Stop, if no new version is available
                Else {
                    Write-Host -ForegroundColor Cyan "No update available for $Product"
                    Write-Output ""
                }
            }
            1 {
                $Product = "RemoteDesktopManager Enterprise"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
                $RemoteDesktopManagerEnterprise = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Remote Desktop Manager*"}).DisplayVersion
                $RemoteDesktopManagerLog = "$LogTemp\RemoteDesktopManager.log"
                $InstallMSI = "$PSScriptRoot\$Product\Setup.RemoteDesktopManagerEnterprise.msi"
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $RemoteDesktopManagerEnterprise"
                If ($RemoteDesktopManagerEnterprise -ne $Version) {
                    DS_WriteLog "I" "Installing $Product" $LogFile
                    Write-Host -ForegroundColor Green "Update available"
                    $Arguments = @(
                        "/i"
                        "`"$InstallMSI`""
                        "/qn"
                        "/L*V $RemoteDesktopManagerLog"
                    )
                    Try {
                        Write-Host "Starting install of $Product $Version"
                        Install-MSI $InstallMSI $Arguments
                        Get-Content $RemoteDesktopManagerLog | Add-Content $LogFile -Encoding ASCI
                        Remove-Item $RemoteDesktopManagerLog
                    } Catch {
                        DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile       
                    }
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
                # Stop, if no new version is available
                Else {
                    Write-Host -ForegroundColor Cyan "No update available for $Product"
                    Write-Output ""
                }
            }
        }
    }

    #// Mark: Install ShareX
    If ($ShareX -eq 1) {
        $Product = "ShareX"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $ShareXV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*ShareX*"}).DisplayVersion
        If (!$ShareXV) {
            $ShareXV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*ShareX*"}).DisplayVersion
        }
        $ShareXInstaller = "ShareX-setup" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $ShareXV"
        If ($ShareXV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/VERYSILENT"
                "/UPDATE"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                Start-Process "$PSScriptRoot\$Product\$ShareXInstaller" -ArgumentList $Options -NoNewWindow
                $p = Get-Process ShareX-setup
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Slack
    If ($Slack -eq 1) {
        $Product = "Slack"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$SlackPlatformClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $SlackV = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Slack*"}).DisplayVersion
        If (!$SlackV) {
            $SlackV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Slack*"}).DisplayVersion
        }
        If (!$SlackV) {
            $SlackV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Slack*"}).DisplayVersion
        }
        If ($SlackV.length -ne "6") {$SlackV = $SlackV -replace ".{2}$"}
        $SlackLog = "$LogTemp\Slack.log"
        $SlackInstaller = "Slack.setup" + "_$ArchitectureClear" + "_$SlackPlatformClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$SlackInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear $SlackPlatformClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $SlackV"
        If ($SlackV -ne $Version) {
            DS_WriteLog "I" "Installing $Product $ArchitectureClear $SlackPlatformClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/norestart"
                "/L*V $SlackLog"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $SlackPlatformClear $Version"
                Install-MSI $InstallMSI $Arguments
                Start-Sleep 25
                Get-Content $SlackLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $SlackLog
            } Catch {
                DS_WriteLog "E" "Error installing $Product $ArchitectureClear $SlackPlatformClear (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install TreeSize
    If ($TreeSize -eq 1) {
        Switch ($TreeSizeType) {
            0 {
                $Product = "TreeSize Free"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
                $Version = $Version.Insert(3,'.')
                $TreeSizeV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $TreeSizeV"
                If ($TreeSizeV -ne $Version) {
                    DS_WriteLog "I" "Install $Product" $LogFile
                    Write-Host -ForegroundColor Green "Update available"
                    Try {
                        Write-Host "Starting install of $Product $Version"
                        Start-Process "$PSScriptRoot\$Product\TreeSize_Free.exe" -ArgumentList /VerySilent -NoNewWindow
                        $p = Get-Process TreeSize_Free
                        If ($p) {
                            $p.WaitForExit()
                            Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                        }
                    } Catch {
                        Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                        DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
                    }
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
                # Stop, if no new version is available
                Else {
                    Write-Host -ForegroundColor Cyan "No update available for $Product"
                    Write-Output ""
                }
            }
            1 {
                $Product = "TreeSize Professional"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
                $Version = $Version.Insert(3,'.')
                $TreeSizeV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $TreeSizeV"
                If ($TreeSizeV -ne $Version) {
                    DS_WriteLog "I" "Install $Product" $LogFile
                    Write-Host -ForegroundColor Green "Update available"
                    Try {
                        Write-Host "Starting install of $Product $Version"
                        Start-Process "$PSScriptRoot\$Product\TreeSize_Professional.exe" -ArgumentList /VerySilent -NoNewWindow
                        $p = Get-Process TreeSize_Professional
                        If ($p) {
                            $p.WaitForExit()
                            Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                        }
                    } Catch {
                        Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                        DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile       
                    }
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
                # Stop, if no new version is available
                Else {
                    Write-Host -ForegroundColor Cyan "No update available for $Product"
                    Write-Output ""
                }
            }
        }
    }

    #// Mark: Install VLC Player
    If ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $VLC = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion
        $VLCLog = "$LogTemp\VLC.log"
        If (!$VLC) {
            $VLC = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion
        }
        If ($VLC) {$VLC = $VLC -replace ".{2}$"}
        $VLCInstaller = "VLC-Player_" + "$ArchitectureClear" +".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$VLCInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $VLC"
        If ($VLC -ne $Version) {
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $VLCLog"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                Install-MSI $InstallMSI $Arguments
                Get-Content $VLCLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $VLCLog
                If (Test-Path -Path "$env:PUBLIC\Desktop\VLC media player.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\VLC media player.lnk" -Force}
            } Catch {
                DS_WriteLog "E" "An error occurred installing $Product (Error: $($Error[0]))" $LogFile 
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install VMWareTools
    If ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $VMWT = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion
        If (!$VMWT) {
            $VMWT = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion
        }
        If ($VMWT) {$VMWT = $VMWT -replace ".{9}$"}
        $VMWareToolsInstaller = "VMWareTools_" + "$ArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $VMWT"
        If ($VMWT -ne $Version) {
            $Options = @(
                "/s"
                "/v"
                "/qn REBOOT=Y"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$VMWareToolsInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                    Write-Host -ForegroundColor Yellow "System needs to reboot after installation!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $ArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install WinSCP
    If ($WinSCP -eq 1) {
        $Product = "WinSCP"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $WSCP = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*WinSCP*"}).DisplayVersion
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $WSCP"
        If ($WSCP -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/VERYSILENT"
                "/ALLUSERS"
                "/NORESTART"
                "/NOCLOSEAPPLICATIONS"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\WinSCP.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                    If (Test-Path -Path "$env:PUBLIC\Desktop\WinSCP.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\WinSCP.lnk" -Force}
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }

    #// Mark: Install Zoom VDI Installer
    If ($Zoom -eq 1) {
        $Product = "Zoom VDI"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $ZoomV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Zoom Client for VDI*"}).DisplayVersion
        If ($ZoomV.length -ne "5") {$ZoomV = $ZoomV -replace ".{4}$"}
        $ZoomLog = "$LogTemp\Zoom.log"
        $ZoomInstaller = "ZoomInstallerVDI" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$ZoomInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $ZoomV"
        If ($ZoomV -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/norestart"
                "/L*V $ZoomLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                Install-MSI $InstallMSI $Arguments
                Start-Sleep 25
                Get-Content $ZoomLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $ZoomLog
                If (Test-Path -Path "$env:PUBLIC\Desktop\Zoom VDI.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Zoom VDI.lnk" -Force}
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
    }
    If ($Zoom -eq 1) {
        If ($ZoomCitrixClient -eq 1) {
            $Product = "Zoom Citrix Client"
            # Check, if a new version is available
            $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
            $Version = Get-Content -Path "$VersionPath"
            $ZoomV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Zoom Plugin*"}).DisplayVersion
            If ($ZoomV.length -ne "5") {$ZoomV = $ZoomV -replace ".{4}$"}
            $ZoomInstaller = "ZoomCitrixHDXMediaPlugin" + ".msi"
            $ZoomLog = "$LogTemp\Zoom.log"
            $InstallMSI = "$PSScriptRoot\$Product\$ZoomInstaller"
            Write-Host -ForegroundColor Magenta "Install $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version: $ZoomV"
            If ($ZoomV -ne $Version) {
                DS_WriteLog "I" "Installing $Product" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $Arguments = @(
                    "/i"
                    "`"$InstallMSI`""
                    "/qn"
                    "/norestart"
                    "/L*V $ZoomLog"
                )
                Try {
                    Write-Host "Starting install of $Product $Version"
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    Get-Content $ZoomLog | Add-Content $LogFile -Encoding ASCI
                    Remove-Item $ZoomLog
                } Catch {
                    DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
            # Stop, if no new version is available
            Else {
                Write-Host -ForegroundColor Cyan "No update available for $Product"
                Write-Output ""
            }
        }
    }
}