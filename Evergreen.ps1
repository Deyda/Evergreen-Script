#requires -version 3
<#
.SYNOPSIS
Download and Install several Software with the Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. and the Nevergreen module from Dan Gough.
.DESCRIPTION
To update or download a software package just select a LastSetting.txt file (With parameter -file) or select your Software out of the GUI.
A new folder for every single package will be created, together with a version file and a log file. If a new version is available
the script checks the version number and will update the package.

.NOTES
  Version:          2.04
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
  2021-04-22        Little customize to the auto update (Error with IE first launch error)
  2021-04-29        Correction Pending Reboot and AutoUpdate Script with List Parameter
  2021-04-30        Add PuTTY Download Function / Add Paint.Net, GIMP, Microsoft PowerToys, Microsoft Visual Studio 2019, Microsoft Visual Studio Code, PuTTY & TeamViewer
  2021-05-01        Adding the new parameter file to extend the parameter execution with a possibility of software selection. / Add auto update for parameter start (-list) with -file parameter / Add Machine Type Selection
  2021-05-02        Add Microsoft Teams User Based Download and Install / Add Visual Studio Code Per User Installer / Connect the Selection Machine Type Physical to Microsoft Teams User Based, Slack Per User and Visual Studio Code Per User
  2021-05-03        GUI Correction deviceTRUST / Add Zoom Full Client Install and Download / Connect the Selection Machine Type Physical to Zoom Full Client, OneDrive User Based and new install.xml file configuration for Microsoft365 Apps and Office 2019 without SharedComputerLicensing / Change download setting for Microsoft365 Apps and Office 2019 install files to Install section (Automated creation of the install.xml is still in the download area and can therefore be adjusted before downloading the install files) / Add Wireshark Download Function / Add Wireshark
  2021-05-05        Add Microsoft Azure Data Studio / Add Save Button
  2021-05-06        Add new LOG and NORESTART Parameter to deviceTRUST Client Install / Auto Create Shortcut on Desktop with ExecutioPolicy ByPass and Noexit Parameter
  2021-05-07        Version formatting customized / Change Oracle Java Version format
  2021-05-12        Implement new languages in Adobe Acrobat Reader DC / Debug No Putty PreRelease / Debug Oracle Java Version Output
  2021-05-18        Implement new Version request for Teams Developer Version / Add new Teams Exploration Version / Add ImageGlass
  2021-05-25        Correction Install GIMP version comparison / Correction OneDrive Machine Based Install / Correction M365 Install
  2021-06-02        Add FSLogix Channel Selection / Move FSLogix ADMX Files to the ADMX folder in Evergreen
  2021-06-11        Correction Notepad++ Download Version
  2021-06-14        Add uberAgent / Correction Foxit Reader Download and Install
  2021-07-02        Minor Update Correction Google Chrome & Microsoft365 Apps
  2021-07-05        Add Cisco Webex Meetings, ControlUp Agent & Console, MS SQL Server Management, MS AVD Remote Desktop, MS Power BI Desktop, Sumatra PDF Reader and RDAnalyzer Download
  2021-07-06        Wireshark download method changed from own to Evergreen / Add Cisco Webex Meetings, ControlUp Agent, MS SQL Server Management Studio Install
  2021-07-07        Correction Notepad++ Version / Add MS AVD Remote Desktop Install / Add Nevergreen PowerShell module
  2021-07-08        Add MS Power BI Desktop Install / Minor Update Correction Microsoft Teams
  2021-07-17        Error Correction FSLogix Installer search, if no preview version is available / Fix Adobe Reader DC update task disable / Fix Microsoft Edge update registry key
  2021-07-18        Activate Change User /Install in Virtual Machine Type Selection / Change Download Method for SumatraPDF
  2021-07-22        Correction MS Edge Download and Install
  2021-07-29        New Log for FW rules (Ray Davis) / Add MS Edge ADMX Download / Correction Citrix Workspace App Download
  2021-07-30        Add MS Office / MS 365 Apps / OneDrive / BISF / Google Chrome / Mozilla Firefox ADMX Download
  2021-08-03        Add Error Action to clean the output
  2021-08-16        Correction Microsoft FSLogix Install and IrfanView Download / Correction FW Log
  2021-08-17        Correction Sumatra PDF Download
  2021-08-18        Correction ADMX Copy MS Edge, Google Chrome, Mozilla Firefox, MS OneDrive and BIS-F / Add ADMX Download Zoom
  2021-08-19        Add ADMX Download Citrix Workspace App Current and LTSR / Add ADMX Download Adobe Acrobat Reader DC / Activate 64 Bit Download Acrobat Reader DC
  2021-08-20        Add Citrix Files, Microsoft Azure CLI, Microsoft Sysinternals, Nmap, TechSmith Snagit, TechSmith Camtasia, LogMeIn GoToMeeting, Git for Windows and Cisco Webex Teams Download
  2021-08-23        Changing the deviceTRUST download from own to Evergreen method / Delete Cisco Webex Meetings / Add WinMerge, PeaZip, Foxit PDF Editor and Microsoft Power BI Report Builder Download / Change Microsoft365 Apps Channels
  2021-08-24        Add 1Password Download / Add 1Password, Citrix Files, Microsoft Azure CLI, Nmap, TechSmith Camtasia, TechSmith SnagIt and Cisco Webex Teams Install
  2021-08-25        Change LogMeIn GoToMeeting to Xen and Local / Add LogMeIn GoToMeeting Xen and Local and Git for Windows Install
  2021-08-26        Add Foxit PDF Editor, WinMerge, Microsoft Power BI Report Builder and PeaZip Install
  2021-08-31        Correction Auto Update List File Parameter
  2021-09-04        Correction Language Parameter M365 Apps
  2021-09-14        Correction Install Parameter Adobe Acrobat Reader
  2021-09-17        Add KeePass Language Function / Add KeePass Language Download and Install / Add IrfanView Language Download and 
  2021-09-21        Add new GUI with second page / Add new variables from second page
  2021-09-22        Change 7 Zip and Adobe Reader DC to new variables
  2021-09-23        Change disable update task for Adobe Acrobat Reader DC, Pro and Google Chrome / Change Citrix Hypervisor, ControlUp Agent, Foxit PDF Editor, Foxit Reader, Git for Windows, Google Chrome, ImageGlass, IrfanView and deviceTRUST to new variables
  2021-09-24        Change KeePass, Microsoft .Net Framework, Microsoft 365 Apps, Microsoft AVD Remote Desktop Microsoft FSLogix, Microsoft Office 2019 and Microsoft Edge to new variables
  2021-09-25        Change Microsoft Power BI Desktop, Microsoft PowerShell Microsoft SQL Server Management Studio, Microsoft Visual Studio Code, Mozilla Firefox, Notepad ++, openJDK, OracleJava8 and Microsoft Teams to new variables / Add Microsoft Edge WebView2 Runtime
  2021-09-27        Change PeaZip, PuTTY, Slack, VLC Player, VMWare Tools, TechSmith SnagIt, WinMerge, Wireshark and Sumatra PDF to new variables / Add Microsoft Project and Microsoft Visio to install.xml creation / Correction Sumatra PDF Reader download link / Change Microsoft Teams download / Add CleanUp Function
  2021-09-28        Add WhatIf Function to Download section / Add OpenFileDialog Function / Add Own Microsoft 365 Apps XML File
  2021-09-29        Add WhatIf Function to Install section / Kill -List Hardcoded Function
  2021-10-04        Add PDF Split & Merge Function / Add PDF Split & Merge Download
  2021-10-05        Correction Copy custom XML / Add PDF Split & Merge Install / Add Autodesk DWG TrueView Function / Add MindView 7 Function / Add Autodesk DWG TrueView Download and Install / Add MindView 7 Download and Install
  2021-10-06        Add Autodesk DWG TrueView, MindView 7 and PDF Split & Merge to GUI and LastSetting.txt
  2021-10-07        Correction Techsmith Camtasia Version / Correction WhatIf Mode
  2021-10-08        Change Machine Type to Installer Label
  2021-10-14        Correction Visio / Project Typo / Fix Microsoft365 Apps Channels
  2021-10-17        Correction Microsoft Teams Machine Based Download
  2021-10-18        Add GUIfile Parameter / Add Mode and Global Information / Add Single Install Type Definition for LogMeInGoToMeeting, Microsoft 365 Apps, Microsoft Visual Studio Code, Microsoft Azure Data Studio, Microsoft OneDrive, Microsoft Teams and Microsoft Office
  2021-10-19        Add Single Install Type Definition for Slack and Zoom / Add Microsoft Office 2021 LTSC / Change Microsoft Office and Microsoft 365 Apps ADMX Architecture / Add Open-Shell Menu Download and Install / Add pdfforge PDFCreator Download Function / Add pdfforge PDFCreator Download and Install / Add Total Commander Download Function / Add Total Commander Download and Install
  2021-10-21        Add Repository Mode
  2021-10-25        Add Start Menu Clean Up Mode
  2021-10-26        Correction Slack and Total Commander installed version detection


.PARAMETER file

Path to file (LastSetting.txt) for software selection in unattended mode.

.PARAMETER GUIfile

Path to GUI file (LastSetting.txt) for software selection in GUI mode.

.EXAMPLE

.\Evergreen.ps1 -file LastSetting.txt

Download and / or Install the selected Software out of the file.

.EXAMPLE

.\Evergreen.ps1 -GUIfile LastSetting.txt

Start the GUI with the options out of the file.

.EXAMPLE

.\Evergreen.ps1

Start the GUI to select the mode (Install and/or Download) and the Software.
#>

[CmdletBinding()]

Param (
    
        [Parameter(
            HelpMessage='Not used anymore',
            ValuefromPipelineByPropertyName = $true
        )]
        [string]$download,

        [Parameter(
            HelpMessage='Not used anymore',
            ValuefromPipelineByPropertyName = $true
        )]
        [string]$install,

        [Parameter(
            HelpMessage='File with Software Selection',
            ValuefromPipelineByPropertyName = $true
        )]
        [string]$file,
    
        [Parameter(
            HelpMessage='File with Software Selection for GUI',
            ValuefromPipelineByPropertyName = $true
        )]
        [string]$GUIfile,

        [Parameter(
            HelpMessage='Not used anymore',
            ValuefromPipelineByPropertyName = $true
        )]
        [switch]$list
    
)

#Add Functions here

# Function OpenFile Dialog
#========================================================================================================================================
function Show-OpenFileDialog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false, Position=0)]
        [System.String]
        $Title = 'Browse',
         
        [Parameter(Mandatory=$false, Position=1)]
        [Object]
        $InitialDirectory = "$PSScriptRoot",
         
        [Parameter(Mandatory=$false, Position=2)]
        [System.String]
        $Filter = 'Extensible Markup Language | *.xml'
    )
     
    Add-Type -AssemblyName PresentationFramework
     
    $dialog = New-Object -TypeName Microsoft.Win32.OpenFileDialog
    $dialog.Title = $Title
    $dialog.InitialDirectory = $InitialDirectory
    $dialog.Filter = $Filter
    if ($dialog.ShowDialog())
    {
        $dialog.FileName
        $WPFTextBox_Filename.Text = $dialog.FileName
    }
    else
    {
        Write-Host -ForegroundColor Red  'Nothing selected.'   
    }
}

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
        $fileD="$Source",
        [switch]$includeStats
    )
    $wc = New-Object Net.WebClient
    $wc.UseDefaultCredentials = $true
    $destination = Join-Path $destinationFolder $fileD
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

# Function IrfanView Download
#========================================================================================================================================
Function Get-IrfanView {
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
        $FileI = $Version -replace "\.",""
        $x32 = "http://download.betanews.com/download/967963863-1/iview$($FileI)_setup.exe"
        $x64 = "http://download.betanews.com/download/967963863-1/iview$($FileI)_x64_setup.exe"


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

# Function Microsoft Teams Download Developer & Beta Version
#========================================================================================================================================
Function Get-MicrosoftTeamsDevBeta() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://github.com/ItzLevvie/MicrosoftTeams-msinternal/blob/master/defconfig"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersionx64dev = '\<td id="LC2".+<\/td\>'
        $webVersionx64dev = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx64dev -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx64dev = $webVersionx64dev.Split("/")
        $appVersionx64dev = $webSplitx64dev[4]
        $regexAppVersionx86dev = '\<td id="LC5".+<\/td\>'
        $webVersionx86dev = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx86dev -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx86dev = $webVersionx86dev.Split("/")
        $appVersionx86dev = $webSplitx86dev[4]
        $regexAppVersionx64beta = '\<td id="LC12".+<\/td\>'
        $webVersionx64beta = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx64beta -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx64beta = $webVersionx64beta.Split("/")
        $appVersionx64beta = $webSplitx64beta[4]
        $regexAppVersionx86beta = '\<td id="LC14".+<\/td\>'
        $webVersionx86beta = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx86beta -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx86beta = $webVersionx86beta.Split("/")
        $appVersionx86beta = $webSplitx86beta[4]
        $appx64URLdev = "https://statics.teams.cdn.office.net/production-windows-x64/$appVersionx64dev/Teams_windows_x64.msi"
        $appx86URLdev = "https://statics.teams.cdn.office.net/production-windows/$appVersionx86dev/Teams_windows.msi"
        $appx64URLbeta = "https://statics.teams.cdn.office.net/production-windows-x64/$appVersionx64beta/Teams_windows_x64.msi"
        $appx86URLbeta = "https://statics.teams.cdn.office.net/production-windows/$appVersionx86beta/Teams_windows.msi"

        $PSObjectx86dev = [PSCustomObject] @{
            Version      = $appVersionx86dev
            Ring         = "Developer"
            Architecture = "x86"
            URI          = $appx86URLdev
        }

        $PSObjectx64dev = [PSCustomObject] @{
            Version      = $appVersionx64dev
            Ring         = "Developer"
            Architecture = "x64"
            URI          = $appx64URLdev
        }

        $PSObjectx86beta = [PSCustomObject] @{
            Version      = $appVersionx86beta
            Ring         = "Exploration"
            Architecture = "x86"
            URI          = $appx86URLbeta
        }

        $PSObjectx64beta = [PSCustomObject] @{
            Version      = $appVersionx64beta
            Ring         = "Exploration"
            Architecture = "x64"
            URI          = $appx64URLbeta
        }
        Write-Output -InputObject $PSObjectx86dev
        Write-Output -InputObject $PSObjectx64dev
        Write-Output -InputObject $PSObjectx86beta
        Write-Output -InputObject $PSObjectx64beta
    }
}

# Function Microsoft Teams Download User Version
#========================================================================================================================================
Function Get-MicrosoftTeamsUser() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://github.com/ItzLevvie/MicrosoftTeams-msinternal/blob/master/defconfig"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
        $TeamsUserVersionGeneral = Get-EvergreenApp -Name MicrosoftTeams | Where-Object { $_.Architecture -eq "x64" -and $_.Ring -eq "General"}
        $VersionGeneral = $TeamsUserVersionGeneral.Version
        $TeamsUserVersionPreview = Get-EvergreenApp -Name MicrosoftTeams | Where-Object { $_.Architecture -eq "x64" -and $_.Ring -eq "Preview"}
        $VersionPreview = $TeamsUserVersionPreview.Version
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersionx64dev = '\<td id="LC3".+<\/td\>'
        $webVersionx64dev = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx64dev -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx64dev = $webVersionx64dev.Split("/")
        $appVersionx64dev = $webSplitx64dev[4]
        $regexAppVersionx86dev = '\<td id="LC4".+<\/td\>'
        $webVersionx86dev = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx86dev -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx86dev = $webVersionx86dev.Split("/")
        $appVersionx86dev = $webSplitx86dev[4]
        $regexAppVersionx64beta = '\<td id="LC11".+<\/td\>'
        $webVersionx64beta = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx64beta -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx64beta = $webVersionx64beta.Split("/")
        $appVersionx64beta = $webSplitx64beta[4]
        $regexAppVersionx86beta = '\<td id="LC13".+<\/td\>'
        $webVersionx86beta = $webRequest.RawContent | Select-String -Pattern $regexAppVersionx86beta -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webSplitx86beta = $webVersionx86beta.Split("/")
        $appVersionx86beta = $webSplitx86beta[4]
        $appx64URLdev = "https://statics.teams.cdn.office.net/production-windows-x64/$appVersionx64dev/Teams_windows_x64.exe"
        $appx86URLdev = "https://statics.teams.cdn.office.net/production-windows/$appVersionx86dev/Teams_windows.exe"
        $appx64URLbeta = "https://statics.teams.cdn.office.net/production-windows-x64/$appVersionx64beta/Teams_windows_x64.exe"
        $appx86URLbeta = "https://statics.teams.cdn.office.net/production-windows/$appVersionx86beta/Teams_windows.exe"
        $appx64URLG = "https://statics.teams.cdn.office.net/production-windows-x64/$VersionGeneral/Teams_windows_x64.exe"
        $appx86URLG = "https://statics.teams.cdn.office.net/production-windows/$VersionGeneral/Teams_windows.exe"
        $appx64URLP = "https://statics.teams.cdn.office.net/production-windows-x64/$VersionPreview/Teams_windows_x64.exe"
        $appx86URLP = "https://statics.teams.cdn.office.net/production-windows/$VersionPreview/Teams_windows.exe"

        $PSObjectx86G = [PSCustomObject] @{
            Version      = $VersionGeneral
            Ring         = "General"
            Architecture = "x86"
            URI          = $appx86URLG
        }

        $PSObjectx64G = [PSCustomObject] @{
            Version      = $VersionGeneral
            Ring         = "General"
            Architecture = "x64"
            URI          = $appx64URLG
        }

        $PSObjectx86P = [PSCustomObject] @{
            Version      = $VersionPreview
            Ring         = "Preview"
            Architecture = "x86"
            URI          = $appx86URLP
        }

        $PSObjectx64P = [PSCustomObject] @{
            Version      = $VersionPreview
            Ring         = "Preview"
            Architecture = "x64"
            URI          = $appx64URLP
        }

        $PSObjectx86dev = [PSCustomObject] @{
            Version      = $appVersionx86dev
            Ring         = "Continuous Deployment"
            Architecture = "x86"
            URI          = $appx86URLdev
        }

        $PSObjectx64dev = [PSCustomObject] @{
            Version      = $appVersionx64dev
            Ring         = "Continuous Deployment"
            Architecture = "x64"
            URI          = $appx64URLdev
        }

        $PSObjectx86beta = [PSCustomObject] @{
            Version      = $appVersionx86beta
            Ring         = "Exploration"
            Architecture = "x86"
            URI          = $appx86URLbeta
        }

        $PSObjectx64beta = [PSCustomObject] @{
            Version      = $appVersionx64beta
            Ring         = "Exploration"
            Architecture = "x64"
            URI          = $appx64URLbeta
        }

        Write-Output -InputObject $PSObjectx86G
        Write-Output -InputObject $PSObjectx64G
        Write-Output -InputObject $PSObjectx86P
        Write-Output -InputObject $PSObjectx64P
        Write-Output -InputObject $PSObjectx86beta
        Write-Output -InputObject $PSObjectx64beta
        Write-Output -InputObject $PSObjectx86dev
        Write-Output -InputObject $PSObjectx64dev
    }
}

# Function PDF Split & Merge Download
#========================================================================================================================================
Function Get-PDFsam() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://pdfsam.org/en/download-pdfsam-basic/"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "Download PDFsam Basic v.{5}"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersion = $webVersion.Split()[3].Trim("v")
        $appx64URL = "https://github.com/torakiki/pdfsam/releases/download/v$appVersion/pdfsam-$appVersion.msi"

        $PSObjectx64 = [PSCustomObject] @{
            Version      = $appVersion
            Channel      = "Stable"
            Architecture = "x64"
            URI          = $appx64URL
        }

        Write-Output -InputObject $PSObjectx64
        
    }
}

# Function Autodesk DWG TrueView Download
#========================================================================================================================================
Function Get-DWGTrueView() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://www.autodesk.de/viewers"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "DWGTrueView_.{4}"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersion = $webVersion.Split("_")[1]
        
        $regexAppURL = "https://efulfillment.*DWGTrueView_" + "$appVersion" + ".*_English_64bit_dlm.sfx.exe"
        $webURL = $webRequest.RawContent | Select-String -Pattern $regexAppURL -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appx64URL = $webURL.Split('"')[0]

        $PSObjectx64 = [PSCustomObject] @{
            Version      = $appVersion
            Language      = "English"
            Architecture = "x64"
            URI          = $appx64URL
        }

        Write-Output -InputObject $PSObjectx64
        
    }
}

# Function Total Commander Download
#========================================================================================================================================
Function Get-TotalCommander() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://www.ghisler.com/ddownload.htm"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "Total Commander Version .{5}"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersion = $webVersion.Split(" ")[3]
        $URLappVersionSplit = $appVersion.Split(".")
        $URLappVersion = $URLappVersionSplit[0] + $URLappVersionSplit[1]
        
        $appx64URL = "https://totalcommander.ch/win/tcmd" + "$URLappVersion" + "x64.exe"
        $appx32URL = "https://totalcommander.ch/win/tcmd" + "$URLappVersion" + "x32.exe"
        
        $PSObjectx64 = [PSCustomObject] @{
            Version      = $appVersion
            Architecture = "x64"
            URI          = $appx64URL
        }

        $PSObjectx32 = [PSCustomObject] @{
            Version      = $appVersion
            Architecture = "x86"
            URI          = $appx32URL
        }

        Write-Output -InputObject $PSObjectx64
        Write-Output -InputObject $PSObjectx32
    }
}

# Function pdfforge PDFCreator
#========================================================================================================================================
Function Get-pdfforgePDFCreator() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersionFree = "https://download.pdfforge.org/download/pdfcreator"
    $appURLVersionProfessional = "https://download.pdfforge.org/download/pdfcreator-professional"
    $appURLVersionTerminal = "https://download.pdfforge.org/download/pdfcreator-terminal-server"
    Try {
        $webRequestFree = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersionFree) -SessionVariable websession
        $webRequestProfessional = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersionProfessional) -SessionVariable websession
        $webRequestTerminal = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersionTerminal) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "Stable Release.{6}"

        $webVersionFree = $webRequestFree.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersionFree = $webVersionFree.Split(" ")[2]
        $webURLFree = "https://download.pdfforge.org/download/pdfcreator/PDFCreator-stable?download"

        $webVersionProfessional = $webRequestProfessional.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersionProfessional = $webVersionProfessional.Split(" ")[2]
        $webURLProfessional = "https://download.pdfforge.org/download/pdfcreator-professional/PDFCreatorProfessional-stable?download"

        $webVersionTerminal = $webRequestTerminal.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $appVersionTerminal = $webVersionTerminal.Split(" ")[2]
        $webURLTerminal = "https://download.pdfforge.org/download/pdfcreator-terminal-server/PDFCreatorTerminalServer-stable?download"

        $PSObjectFree = [PSCustomObject] @{
            Version      = $appVersionFree
            Channel      = "Free"
            URI          = $webURLFree
        }

        $PSObjectProfessional = [PSCustomObject] @{
            Version      = $appVersionProfessional
            Channel      = "Professional"
            URI          = $webURLProfessional
        }

        $PSObjectTerminal = [PSCustomObject] @{
            Version      = $appVersionTerminal
            Channel      = "Terminal Server"
            URI          = $webURLTerminal
        }

        Write-Output -InputObject $PSObjectFree
        Write-Output -InputObject $PSObjectProfessional
        Write-Output -InputObject $PSObjectTerminal
    }
}

# Function MindView 7 Download
#========================================================================================================================================
Function Get-MindView7() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://www.matchware.com/de/mindview-servicepacks"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "MindView 7.*"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $webVersion = $webVersion.Split("(")[1]
        $appVersion = $webVersion.Split(")")[0]
        
        $appx64URLDE = "https://link.matchware.com/mindview7_ge"
        $appx64URLEN = "https://link.matchware.com/mindview7_en"
        $appx64URLFR = "https://link.matchware.com/mindview7_fr"
        $appx64URLDK = "https://link.matchware.com/mindview7_dk"

        $PSObjectx64de = [PSCustomObject] @{
            Version      = $appVersion
            Language      = "German"
            URI          = $appx64URLDE
        }

        $PSObjectx64en = [PSCustomObject] @{
            Version      = $appVersion
            Language      = "English"
            URI          = $appx64URLEN
        }

        $PSObjectx64fr = [PSCustomObject] @{
            Version      = $appVersion
            Language      = "French"
            URI          = $appx64URLFR
        }

        $PSObjectx64dk = [PSCustomObject] @{
            Version      = $appVersion
            Language      = "Danish"
            URI          = $appx64URLDK
        }

        Write-Output -InputObject $PSObjectx64de
        Write-Output -InputObject $PSObjectx64en
        Write-Output -InputObject $PSObjectx64fr
        Write-Output -InputObject $PSObjectx64dk
        
    }
}

# Function PuTTY Download Stable and Pre-Release Version
#========================================================================================================================================
Function Get-PuTTY() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html"
    $appURLVersionPre = "https://www.chiark.greenend.org.uk/~sgtatham/putty/prerel.html"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
        $webRequestPre = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersionPre) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion or $appURLVersionPre with error $_."
        Break
    }
    Finally {
        $regexAppVersion = "\(.*\)\<\/TITLE\>"
        $webVersion = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        $CacheVersion = $webVersion.Split()[0].Trim("</TITLE>")
        $CacheVersion = $CacheVersion.Split()[0].Trim("(")
        $CacheVersion = $CacheVersion.Split()[0].Trim(")")
        $appVersion = $CacheVersion
        $appx64URL = "https://the.earth.li/~sgtatham/putty/latest/w64/putty-64bit-$appVersion-installer.msi"
        $appx86URL = "https://the.earth.li/~sgtatham/putty/latest/w32/putty-$appVersion-installer.msi"
        $regexAppVersionPre = "of .*"
        $webVersionPre = $webRequestPre.RawContent | Select-String -Pattern $regexAppVersionPre -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -First 1
        If ($webVersionPre){
            $appVersionPre = $webVersionPre.Split()[1].Trim("of ")
            $appx64URLPre = "https://tartarus.org/~simon/putty-prerel-snapshots/w64/putty-64bit-installer.msi"
            $appx86URLPre = "https://tartarus.org/~simon/putty-prerel-snapshots/w32/putty-installer.msi"
        }
        $PSObjectx86 = [PSCustomObject] @{
            Version      = $appVersion
            Channel      = "Stable"
            Architecture = "x86"
            URI          = $appx86URL
        }

        $PSObjectx64 = [PSCustomObject] @{
            Version      = $appVersion
            Channel      = "Stable"
            Architecture = "x64"
            URI          = $appx64URL
        }
        If ($webVersionPre){
            $PSObjectx86Pre = [PSCustomObject] @{
                Version      = $appVersionPre
                Channel      = "Pre-Release"
                Architecture = "x86"
                URI          = $appx86URLPre
            }

            $PSObjectx64Pre = [PSCustomObject] @{
                Version      = $appVersionPre
                Channel      = "Pre-Release"
                Architecture = "x64"
                URI          = $appx64URLPre
            }
        }
        else {
            $PSObjectx86Pre = [PSCustomObject] @{
            Version      = $appVersion
            Channel      = "Pre-Release"
            Architecture = "x86"
            URI          = $appx86URL
            }

            $PSObjectx64Pre = [PSCustomObject] @{
                Version      = $appVersion
                Channel      = "Pre-Release"
                Architecture = "x64"
                URI          = $appx64URL
            }
        }

        Write-Output -InputObject $PSObjectx86
        Write-Output -InputObject $PSObjectx64
        Write-Output -InputObject $PSObjectx86Pre
        Write-Output -InputObject $PSObjectx64Pre
        
    }
}

# Function KeePass Language
#========================================================================================================================================
Function Get-KeePassLanguage() {
    [OutputType([System.Management.Automation.PSObject])]
    [CmdletBinding()]
    Param ()
    $appURLVersion = "https://keepass.info/translations.html"
    Try {
        $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($appURLVersion) -SessionVariable websession
    }
    Catch {
        Throw "Failed to connect to URL: $appURLVersion with error $_."
        Break
    }
    Finally {
        $regexLanguageGerman = 'https:\/\/.*German.zip'
        $URLGerman = $webRequest.RawContent | Select-String -Pattern $regexLanguageGerman -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionGerman = $URLGerman.Split("-")[1]
        $regexLanguageDutch = 'https:\/\/.*Dutch.zip'
        $URLDutch = $webRequest.RawContent | Select-String -Pattern $regexLanguageDutch -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionDutch = $URLDutch.Split("-")[1]
        $regexLanguageDanish = 'https:\/\/.*Danish.zip'
        $URLDanish = $webRequest.RawContent | Select-String -Pattern $regexLanguageDanish -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionDanish = $URLDanish.Split("-")[1]
        $regexLanguageFinnish = 'https:\/\/.*Finnish.zip'
        $URLFinnish = $webRequest.RawContent | Select-String -Pattern $regexLanguageFinnish -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionFinnish = $URLFinnish.Split("-")[1]
        $regexLanguageFrench = 'https:\/\/.*French.zip'
        $URLFrench = $webRequest.RawContent | Select-String -Pattern $regexLanguageFrench -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionFrench = $URLFrench.Split("-")[1]
        $regexLanguageItalian = 'https:\/\/.*Italian-b.zip'
        $URLItalian = $webRequest.RawContent | Select-String -Pattern $regexLanguageItalian -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionItalian = $URLItalian.Split("-")[1]
        $regexLanguageJapanese = 'https:\/\/.*Japanese.zip'
        $URLJapanese = $webRequest.RawContent | Select-String -Pattern $regexLanguageJapanese -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionJapanese = $URLJapanese.Split("-")[1]
        $regexLanguageKorean = 'https:\/\/.*Korean.zip'
        $URLKorean = $webRequest.RawContent | Select-String -Pattern $regexLanguageKorean -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionKorean = $URLKorean.Split("-")[1]
        $regexLanguageNorwegian = 'https:\/\/.*Norwegian_NB.zip'
        $URLNorwegian = $webRequest.RawContent | Select-String -Pattern $regexLanguageNorwegian -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionNorwegian = $URLNorwegian.Split("-")[1]
        $regexLanguagePolish = 'https:\/\/.*Polish.zip'
        $URLPolish = $webRequest.RawContent | Select-String -Pattern $regexLanguagePolish -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionPolish = $URLPolish.Split("-")[1]
        $regexLanguagePortuguese = 'https:\/\/.*Portuguese_PT.zip'
        $URLPortuguese = $webRequest.RawContent | Select-String -Pattern $regexLanguagePortuguese -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionPortuguese = $URLPortuguese.Split("-")[1]
        $regexLanguageRussian = 'https:\/\/.*Russian.zip'
        $URLRussian = $webRequest.RawContent | Select-String -Pattern $regexLanguageRussian -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionRussian = $URLRussian.Split("-")[1]
        $regexLanguageSpanish = 'https:\/\/.*Spanish.zip'
        $URLSpanish = $webRequest.RawContent | Select-String -Pattern $regexLanguageSpanish -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionSpanish = $URLSpanish.Split("-")[1]
        $regexLanguageSwedish = 'https:\/\/.*Swedish.zip'
        $URLSwedish = $webRequest.RawContent | Select-String -Pattern $regexLanguageSwedish -AllMatches | ForEach-Object { $_.Matches.Value } | Select-Object -Last 1
        $VersionSwedish = $URLSwedish.Split("-")[1]

        $PSObjectGerman = [PSCustomObject] @{
            Language     = "German"
            Version      = $VersionGerman
            URI          = $URLGerman
        }

        $PSObjectDutch = [PSCustomObject] @{
            Language     = "Dutch"
            Version      = $VersionDutch
            URI          = $URLDutch
        }

        $PSObjectDanish = [PSCustomObject] @{
            Language     = "Danish"
            Version      = $VersionDanish
            URI          = $URLDanish
        }

        $PSObjectFinnish = [PSCustomObject] @{
            Language     = "Finnish"
            Version      = $VersionFinnish
            URI          = $URLFinnish
        }

        $PSObjectFrench = [PSCustomObject] @{
            Language     = "French"
            Version      = $VersionFrench
            URI          = $URLFrench
        }

        $PSObjectItalian = [PSCustomObject] @{
            Language     = "Italian"
            Version      = $VersionItalian
            URI          = $URLItalian
        }

        $PSObjectJapanese = [PSCustomObject] @{
            Language     = "Japanese"
            Version      = $VersionJapanese
            URI          = $URLJapanese
        }

        $PSObjectKorean = [PSCustomObject] @{
            Language     = "Korean"
            Version      = $VersionKorean
            URI          = $URLKorean
        }

        $PSObjectNorwegian = [PSCustomObject] @{
            Language     = "Norwegian"
            Version      = $VersionNorwegian
            URI          = $URLNorwegian
        }

        $PSObjectPolish = [PSCustomObject] @{
            Language     = "Polish"
            Version      = $VersionPolish
            URI          = $URLPolish
        }

        $PSObjectPortuguese = [PSCustomObject] @{
            Language     = "Portuguese"
            Version      = $VersionPortuguese
            URI          = $URLPortuguese
        }

        $PSObjectRussian = [PSCustomObject] @{
            Language     = "Russian"
            Version      = $VersionRussian
            URI          = $URLRussian
        }

        $PSObjectSpanish = [PSCustomObject] @{
            Language     = "Spanish"
            Version      = $VersionSpanish
            URI          = $URLSpanish
        }

        $PSObjectSwedish = [PSCustomObject] @{
            Language     = "Swedish"
            Version      = $VersionSwedish
            URI          = $URLSwedish
        }

        Write-Output -InputObject $PSObjectGerman
        Write-Output -InputObject $PSObjectDutch
        Write-Output -InputObject $PSObjectDanish
        Write-Output -InputObject $PSObjectFinnish
        Write-Output -InputObject $PSObjectFrench
        Write-Output -InputObject $PSObjectJapanese
        Write-Output -InputObject $PSObjectKorean
        Write-Output -InputObject $PSObjectItalian
        Write-Output -InputObject $PSObjectNorwegian
        Write-Output -InputObject $PSObjectPolish
        Write-Output -InputObject $PSObjectPortuguese
        Write-Output -InputObject $PSObjectRussian
        Write-Output -InputObject $PSObjectSpanish
        Write-Output -InputObject $PSObjectSwedish
    }
}

# Function Microsoft Office ADMX Download
#========================================================================================================================================
function Get-MicrosoftOfficeAdmx {
    $id = "49030"
    $urlversion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $urldownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($id)"
    try {
        $ProgressPreference = 'SilentlyContinue'
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urlversion -ErrorAction SilentlyContinue
        $str = ($web.ToString() -split "[`r`n]" | Select-String "Version:").ToString()
        $Version = ($str | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urldownload -ErrorAction SilentlyContinue -MaximumRedirection 0
        $hrefx64 = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.exe" -and $_.href -like "*x64*" } | Select-Object -First 1
        $hrefx86 = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.exe" -and $_.href -like "*x86*" } | Select-Object -First 1
        $PSObjectx86 = [PSCustomObject] @{
            Version      = $Version
            Architecture = "x86"
            URI          = $hrefx86.href
        }

        $PSObjectx64 = [PSCustomObject] @{
            Version      = $Version
            Architecture = "x64"
            URI          = $hrefx64.href
        }
    }
    catch {
        Throw $_
    }
    Write-Output -InputObject $PSObjectx86
    Write-Output -InputObject $PSObjectx64
}

# Function Google Chrome ADMX Download
#========================================================================================================================================
function Get-GoogleChromeAdmx {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $URI = "https://dl.google.com/dl/edgedl/chrome/policy/policy_templates.zip"
        Invoke-WebRequest -Uri $URI -OutFile "$($env:TEMP)\policy_templates.zip"
        Expand-Archive -Path "$($env:TEMP)\policy_templates.zip" -DestinationPath "$($env:TEMP)\chromeadmx" -Force
        $versionfile = (Get-Content -Path "$($env:TEMP)\chromeadmx\VERSION").Split('=')
        $Version = "$($versionfile[1]).$($versionfile[3]).$($versionfile[5]).$($versionfile[7])"
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

# Function Mozilla Firefox ADMX Download
#========================================================================================================================================
function Get-MozillaFirefoxAdmx {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $repo = "mozilla/policy-templates"
        $latest = (Invoke-WebRequest -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]
        $Version = ($latest.tag_name | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        $URI = $latest.assets.browser_download_url
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

# Function Adobe Acrobat Reader DC ADMX Download
#========================================================================================================================================
function Get-AdobeAcrobatReaderDCAdmx {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $fileA = "ReaderADMTemplate.zip"
        $url = "ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC/misc/"
        Write-Verbose "FTP $($url)"
        $listRequest = [Net.WebRequest]::Create($url)
        $listRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
        $lines = New-Object System.Collections.ArrayList
        $listResponse = $listRequest.GetResponse()
        $listStream = $listResponse.GetResponseStream()
        $listReader = New-Object System.IO.StreamReader($listStream)
        while (!$listReader.EndOfStream)
        {
            $line = $listReader.ReadLine()
            if ($line.Contains($fileA)) { $lines.Add($line) | Out-Null }
        }
        $listReader.Dispose()
        $listStream.Dispose()
        $listResponse.Dispose()
        Write-Verbose "received $($line.Length) characters response"
        $tokens = $lines[0].Split(" ", 9, [StringSplitOptions]::RemoveEmptyEntries)
        $Version = Get-Date -Date "$($tokens[6])/$($tokens[5])/$($tokens[7])" -Format "yy.M.d"
        return @{ Version = $Version; URI = "$($url)$($fileA)" }
    }
    catch {
        Throw $_
    }
}

# Function Citrix Workspace App Current ADMX Download
#========================================================================================================================================
function Get-CitrixWorkspaceAppCurrentAdmx {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://www.citrix.com/downloads/workspace-app/windows/workspace-app-for-windows-latest.html"
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Ignore
        $str = ($web.Content -split "`r`n" | Select-String -Pattern "_ADMX_")[0].ToString().Trim()
        $URI = "https:$(((Select-String '(\/\/)([^\s,]+)(?=")' -Input $str).Matches.Value))"
        $filename = $URI.Split("/")[4].Split('?')[0].Split('_')[3]
        $Version = $filename.Replace(".zip", "") #($filename | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        $Path = $Version
        if ($Version -notcontains '.') { $Version += ".0" }
        return @{ Version = $Version; URI = $URI; Path = $Path }
    }
    catch {
        Throw $_
    }
}

# Function Citrix Workspace App LTSR ADMX Download
#========================================================================================================================================
function Get-CitrixWorkspaceAppLTSRAdmx {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://www.citrix.com/downloads/workspace-app/workspace-app-for-windows-long-term-service-release/workspace-app-for-windows-1912ltsr.html"
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Ignore
        $str = ($web.Content -split "`r`n" | Select-String -Pattern "_ADMX_")[0].ToString().Trim()
        $URI = "https:$(((Select-String '(\/\/)([^\s,]+)(?=")' -Input $str).Matches.Value))"
        $filename = $URI.Split("/")[4].Split('?')[0].Split('_')[4]
        $Version = $filename.Replace(".zip", "") 
        $Path = $filename.Replace(".zip", "") #($filename | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        return @{ Version = $Version; URI = $URI; Path = $Path }
    }
    catch {
        Throw $_
    }
}

# Function Zoom ADMX Download
#========================================================================================================================================
function Get-ZoomAdmx {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://support.zoom.us/hc/en-us/articles/360039100051"
        # grab content
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Ignore
        # find ADMX download
        $URI = (($web.Links | Where-Object {$_.href -like "*msi-templates*.zip"})[-1]).href
        # grab version
        $Version = ($URI.Split("/")[-1] | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
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
$ErrorActionPreference = 'SilentlyContinue'

# Is there a newer Evergreen Script version?
# ========================================================================================================================================
$eVersion = "2.04"
[bool]$NewerVersion = $false
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$WebResponseVersion = Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/Deyda/Evergreen-Script/main/Evergreen.ps1"
If (!$WebVersion) {
    $WebVersion = (($WebResponseVersion.tostring() -split "[`r`n]" | select-string "Version:" | Select-Object -First 1) -split ":")[1].Trim()
}
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
Write-Host -BackgroundColor DarkGreen -ForegroundColor Yellow "                     Version $eVersion                          "
$host.ui.RawUI.WindowTitle ="Evergreen Script - Update your Software, the lazy way - Manuel Winkel (www.deyda.net) - Version $eVersion"
If (Test-Path "$PSScriptRoot\update.ps1" -PathType leaf) {
    #Remove-Item -Path "$PSScriptRoot\Update.ps1" -Force
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
    If ($file) {
        #Write-Host -Foregroundcolor Red "File: $file."
        $update = @'
        Remove-Item -Path "$PSScriptRoot\Evergreen.ps1" -Force 
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/Deyda/Evergreen-Script/main/Evergreen.ps1 -OutFile ("$PSScriptRoot\" + "Evergreen.ps1")
        & "$PSScriptRoot\evergreen.ps1" -download -file $file
'@
        $update > $PSScriptRoot\update.ps1
        & "$PSScriptRoot\update.ps1"
        Break
        
    }
    ElseIf ($GUIfile) {
        #Write-Host -Foregroundcolor Red "GUI File: $GUIfile."
        $update = @'
        Remove-Item -Path "$PSScriptRoot\Evergreen.ps1" -Force 
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/Deyda/Evergreen-Script/main/Evergreen.ps1 -OutFile ("$PSScriptRoot\" + "Evergreen.ps1")
        & "$PSScriptRoot\evergreen.ps1" -download -GUIfile $GUIfile
'@
        $update > $PSScriptRoot\update.ps1
        & "$PSScriptRoot\update.ps1"
        Break
        
    }
    Else {
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

If ($file -eq $False) {
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
        Title="Evergreen Script - Update your Software, the lazy way - Version $eVersion" Height="980" Width="940"
        Icon="$PSScriptRoot\shortcut\EvergreenLeafDeyda.ico"
        WindowStartupLocation="CenterScreen">
    <TabControl Grid.Column="1">
        <TabItem Header="General">
            <ScrollViewer ScrollViewer.VerticalScrollBarVisibility="Auto">
                <Grid x:Name="General_Grid" Margin="0,0,0,1" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="13*"/>
                        <ColumnDefinition Width="234*"/>
                        <ColumnDefinition Width="586*"/>
                    </Grid.ColumnDefinitions>
                    <Image x:Name="Image_Logo" Height="100" Margin="510,3,30,0" VerticalAlignment="Top" Width="100" Source="$PSScriptRoot\img\Logo_DEYDA_no_cta.png" Grid.Column="2" ToolTip="www.deyda.net"/>
                    <Button x:Name="Button_Start" Content="Start" HorizontalAlignment="Left" Margin="291,844,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
                    <Button x:Name="Button_Cancel" Content="Cancel" HorizontalAlignment="Left" Margin="386,844,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
                    <Button x:Name="Button_Save" Content="Save" HorizontalAlignment="Left" Margin="522,844,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2" ToolTip="Save Selected Software in LastSetting.txt or -GUIFile Parameter file"/>
                    <Label x:Name="Label_SelectMode" Content="Select Mode" HorizontalAlignment="Left" Margin="11,3,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_Download" Content="Download" HorizontalAlignment="Left" Margin="15,34,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_Install" Content="Install" HorizontalAlignment="Left" Margin="103,34,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <Label x:Name="Label_SelectLanguage" Content="Select Language" HorizontalAlignment="Left" Margin="124,3,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Language" HorizontalAlignment="Left" Margin="140,30,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="2" ToolTip="If this is selectable at download!">
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
                    <Label x:Name="Label_SelectArchitecture" Content="Select Architecture" HorizontalAlignment="Left" Margin="234,3,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Architecture" HorizontalAlignment="Left" Margin="264,30,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2" ToolTip="If this is selectable at download!">
                        <ListBoxItem Content="x64"/>
                        <ListBoxItem Content="x86"/>
                    </ComboBox>
                    <Label x:Name="Label_SelectInstaller" Content="Select Installer Type" HorizontalAlignment="Left" Margin="353,3,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Installer" HorizontalAlignment="Left" Margin="359,30,0,0" VerticalAlignment="Top" SelectedIndex="0" RenderTransformOrigin="0.864,0.591" Grid.Column="2" ToolTip="If this is different at install!">
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <Label x:Name="Label_Explanation" Content="Selected globally, can be defined individually via Detail tab." HorizontalAlignment="Left" Margin="153,52,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2" />
                    <Label x:Name="Label_Software" Content="Select Software" HorizontalAlignment="Left" Margin="11,67,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_1Password" Content="1Password" HorizontalAlignment="Left" Margin="15,98,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_7Zip" Content="7 Zip" HorizontalAlignment="Left" Margin="15,118,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_AdobeProDC" Content="Adobe Pro DC" HorizontalAlignment="Left" Margin="15,138,0,0" VerticalAlignment="Top" Grid.Column="1" ToolTip="Update Only!"/>
                    <CheckBox x:Name="Checkbox_AdobeReaderDC" Content="Adobe Reader DC" HorizontalAlignment="Left" Margin="15,158,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_AutodeskDWGTrueView" Content="Autodesk DWG TrueView" HorizontalAlignment="Left" Margin="15,178,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_BISF" Content="BIS-F" HorizontalAlignment="Left" Margin="15,198,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_CiscoWebexTeams" Content="Cisco Webex Teams" HorizontalAlignment="Left" Margin="15,218,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_CitrixFiles" Content="Citrix Files" HorizontalAlignment="Left" Margin="15,238,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_CitrixHypervisorTools" Content="Citrix Hypervisor Tools" HorizontalAlignment="Left" Margin="15,258,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <CheckBox x:Name="Checkbox_CitrixWorkspaceApp" Content="Citrix Workspace App" HorizontalAlignment="Left" Margin="15,278,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_CitrixWorkspaceApp" HorizontalAlignment="Left" Margin="215,274,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.ColumnSpan="2" Grid.Column="1">
                        <ListBoxItem Content="Current Release"/>
                        <ListBoxItem Content="Long Term Service Release"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_ControlUpAgent" Content="ControlUp Agent" HorizontalAlignment="Left" Margin="15,298,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_ControlUpAgent" HorizontalAlignment="Left" Margin="215,295,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.ColumnSpan="2" Grid.Column="1">
                        <ListBoxItem Content=".Net 3.5"/>
                        <ListBoxItem Content=".Net 4.5"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_ControlUpConsole" Content="ControlUp Console" HorizontalAlignment="Left" Margin="15,318,0,0" VerticalAlignment="Top" Grid.Column="1" ToolTip="Only Download"/>
                    <CheckBox x:Name="Checkbox_deviceTRUST" Content="deviceTRUST" HorizontalAlignment="Left" Margin="15,338,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_deviceTRUST" HorizontalAlignment="Left" Margin="215,335,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Client"/>
                        <ListBoxItem Content="Host"/>
                        <ListBoxItem Content="Console"/>
                        <ListBoxItem Content="Client + Host"/>
                        <ListBoxItem Content="Host + Console"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_Filezilla" Content="Filezilla" HorizontalAlignment="Left" Margin="15,358,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <CheckBox x:Name="Checkbox_FoxitPDFEditor" Content="Foxit PDF Editor" HorizontalAlignment="Left" Margin="15,378,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_FoxitReader" Content="Foxit Reader" HorizontalAlignment="Left" Margin="15,398,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_GIMP" Content="GIMP" HorizontalAlignment="Left" Margin="15,418,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_GitForWindows" Content="Git for Windows" HorizontalAlignment="Left" Margin="15,438,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_GoogleChrome" Content="Google Chrome" HorizontalAlignment="Left" Margin="15,458,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_Greenshot" Content="Greenshot" HorizontalAlignment="Left" Margin="15,478,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_ImageGlass" Content="ImageGlass" HorizontalAlignment="Left" Margin="15,498,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_IrfanView" Content="IrfanView" HorizontalAlignment="Left" Margin="15,518,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_KeePass" Content="KeePass" HorizontalAlignment="Left" Margin="15,538,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_LogMeInGoToMeeting" Content="LogMeIn GoToMeeting" HorizontalAlignment="Left" Margin="15,558,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_MSDotNetFramework" Content="Microsoft .Net Framework" HorizontalAlignment="Left" Margin="15,578,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_MSDotNetFramework" HorizontalAlignment="Left" Margin="215,573,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Current"/>
                        <ListBoxItem Content="LTS (Long Term Support)"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MS365Apps" Content="Microsoft 365 Apps" HorizontalAlignment="Left" Margin="15,598,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MS365Apps" HorizontalAlignment="Left" Margin="215,594,0,0" VerticalAlignment="Top" SelectedIndex="4" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Beta Channel"/>
                        <ListBoxItem Content="Current Channel (Preview)"/>
                        <ListBoxItem Content="Current Channel"/>
                        <ListBoxItem Content="Semi-Annual Channel (Preview)"/>
                        <ListBoxItem Content="Monthly Enterprise Channel"/>
                        <ListBoxItem Content="Semi-Annual Channel"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSAVDRemoteDesktop" Content="Microsoft AVD Remote Desktop" HorizontalAlignment="Left" Margin="15,618,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_MSAVDRemoteDesktop" HorizontalAlignment="Left" Margin="215,615,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Insider"/>
                        <ListBoxItem Content="Public"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSAzureCLI" Content="Microsoft Azure CLI" HorizontalAlignment="Left" Margin="15,638,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_MSAzureDataStudio" Content="Microsoft Azure Data Studio" HorizontalAlignment="Left" Margin="15,658,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSAzureDataStudio" HorizontalAlignment="Left" Margin="215,656,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Insider"/>
                        <ListBoxItem Content="Stable"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSEdge" Content="Microsoft Edge" HorizontalAlignment="Left" Margin="15,678,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSEdge" HorizontalAlignment="Left" Margin="215,676,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Developer"/>
                        <ListBoxItem Content="Beta"/>
                        <ListBoxItem Content="Stable"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSEdgeWebView2" Content="Microsoft Edge WebView2" HorizontalAlignment="Left" Margin="15,698,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_MSFSlogix" Content="Microsoft FSLogix" HorizontalAlignment="Left" Margin="15,718,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSFSlogix" HorizontalAlignment="Left" Margin="215,713,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Preview"/>
                        <ListBoxItem Content="Production"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSOffice" Content="Microsoft Office" HorizontalAlignment="Left" Margin="15,738,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSOffice" HorizontalAlignment="Left" Margin="215,734,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="2019"/>
                        <ListBoxItem Content="2021 LTSC"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSOneDrive" Content="Microsoft OneDrive" HorizontalAlignment="Left" Margin="15,758,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSOneDrive" HorizontalAlignment="Left" Margin="215,755,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Insider Ring"/>
                        <ListBoxItem Content="Production Ring"/>
                        <ListBoxItem Content="Enterprise Ring"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSPowerBIDesktop" Content="Microsoft Power BI Desktop" HorizontalAlignment="Left" Margin="15,778,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <CheckBox x:Name="Checkbox_MSPowerBIReportBuilder" Content="Microsoft Power BI Report Builder" HorizontalAlignment="Left" Margin="15,798,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <CheckBox x:Name="Checkbox_MSPowerShell" Content="Microsoft PowerShell" HorizontalAlignment="Left" Margin="15,818,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_MSPowerShell" HorizontalAlignment="Left" Margin="215,816,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="Stable"/>
                        <ListBoxItem Content="LTS (Long Term Support)"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSPowerToys" Content="Microsoft PowerToys" HorizontalAlignment="Left" Margin="15,838,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <CheckBox x:Name="Checkbox_MSSQLServerManagementStudio" Content="Microsoft SQL Server Management Studio" Margin="190,98,0,0" VerticalAlignment="Top" Grid.Column="2" HorizontalAlignment="Left"/>
                    <CheckBox x:Name="Checkbox_MSSysinternals" Content="Microsoft Sysinternals" Margin="190,118,0,0" VerticalAlignment="Top" Grid.Column="2" HorizontalAlignment="Left" ToolTip="Only Download"/>
                    <CheckBox x:Name="Checkbox_MSTeams" Content="Microsoft Teams" HorizontalAlignment="Left" Margin="190,138,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_MSTeams_No_AutoStart" Content="No AutoStart" HorizontalAlignment="Left" Margin="535,138,0,0" VerticalAlignment="Top" Grid.Column="2" ToolTip="Delete the HKLM Run entry to AutoStart Microsoft Teams"/>
                    <ComboBox x:Name="Box_MSTeams" HorizontalAlignment="Left" Margin="394,135,0,0" VerticalAlignment="Top" SelectedIndex="3" Grid.Column="2">
                        <ListBoxItem Content="Continuous Ring"/>
                        <ListBoxItem Content="Exploration Ring"/>
                        <ListBoxItem Content="Preview Ring"/>
                        <ListBoxItem Content="General Ring"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSVisualStudio" Content="Microsoft Visual Studio 2019" HorizontalAlignment="Left" Margin="190,158,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_MSVisualStudio" HorizontalAlignment="Left" Margin="394,156,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="2">
                        <ListBoxItem Content="Enterprise Edition"/>
                        <ListBoxItem Content="Professional Edition"/>
                        <ListBoxItem Content="Community Edition"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSVisualStudioCode" Content="Microsoft Visual Studio Code" HorizontalAlignment="Left" Margin="190,178,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_MSVisualStudioCode" HorizontalAlignment="Left" Margin="394,177,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="2">
                        <ListBoxItem Content="Insider"/>
                        <ListBoxItem Content="Stable"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MindView7" Content="MindView 7" HorizontalAlignment="Left" Margin="190,198,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <CheckBox x:Name="Checkbox_Firefox" Content="Mozilla Firefox" HorizontalAlignment="Left" Margin="190,218,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Firefox" HorizontalAlignment="Left" Margin="394,217,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="Current"/>
                        <ListBoxItem Content="ESR"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_mRemoteNG" Content="mRemoteNG" HorizontalAlignment="Left" Margin="190,238,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_Nmap" Content="Nmap" HorizontalAlignment="Left" Margin="190,258,0,0" VerticalAlignment="Top" Grid.Column="2" ToolTip="No silent installation!"/>
                    <CheckBox x:Name="Checkbox_NotepadPlusPlus" Content="Notepad ++" HorizontalAlignment="Left" Margin="190,278,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_OpenJDK" Content="Open JDK" HorizontalAlignment="Left" Margin="190,298,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_OpenShellMenu" Content="Open-Shell Menu" HorizontalAlignment="Left" Margin="190,318,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_OracleJava8" Content="Oracle Java 8" HorizontalAlignment="Left" Margin="190,338,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_PaintDotNet" Content="Paint.Net" HorizontalAlignment="Left" Margin="190,358,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_PDFForgeCreator" Content="pdfforge PDFCreator" HorizontalAlignment="Left" Margin="190,378,0,0" VerticalAlignment="Top" Grid.Column="2" ToolTip="No Silent Installation!"/>
                    <ComboBox x:Name="Box_pdfforgePDFCreator" HorizontalAlignment="Left" Margin="394,375,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="Free"/>
                        <ListBoxItem Content="Professional"/>
                        <ListBoxItem Content="Terminal Server"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_PDFsam" Content="PDF Split and Merge" HorizontalAlignment="Left" Margin="190,398,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_PeaZip" Content="PeaZip" HorizontalAlignment="Left" Margin="190,418,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_Putty" Content="PuTTY" HorizontalAlignment="Left" Margin="190,438,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Putty" HorizontalAlignment="Left" Margin="394,435,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="Pre-Release"/>
                        <ListBoxItem Content="Stable"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_RemoteDesktopManager" Content="Remote Desktop Manager" HorizontalAlignment="Left" Margin="190,458,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_RemoteDesktopManager" HorizontalAlignment="Left" Margin="394,456,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="Free"/>
                        <ListBoxItem Content="Enterprise"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_RDAnalyzer" Content="Remote Display Analyzer" HorizontalAlignment="Left" Margin="190,478,0,0" VerticalAlignment="Top" Grid.Column="2" ToolTip="Only Download"/>
                    <CheckBox x:Name="Checkbox_ShareX" Content="ShareX" HorizontalAlignment="Left" Margin="190,498,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <CheckBox x:Name="Checkbox_Slack" Content="Slack" HorizontalAlignment="Left" Margin="190,518,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <CheckBox x:Name="Checkbox_SumatraPDF" Content="Sumatra PDF" HorizontalAlignment="Left" Margin="190,538,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <CheckBox x:Name="Checkbox_TeamViewer" Content="TeamViewer" HorizontalAlignment="Left" Margin="190,558,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_TechSmithCamtasia" Content="TechSmith Camtasia" HorizontalAlignment="Left" Margin="190,578,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_TechSmithSnagIt" Content="TechSmith SnagIt" HorizontalAlignment="Left" Margin="190,598,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_TotalCommander" Content="Total Commander" HorizontalAlignment="Left" Margin="190,618,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_TreeSize" Content="TreeSize" HorizontalAlignment="Left" Margin="190,638,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_TreeSize" HorizontalAlignment="Left" Margin="394,636,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="Free"/>
                        <ListBoxItem Content="Professional"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_uberAgent" Content="uberAgent" HorizontalAlignment="Left" Margin="190,658,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_VLCPlayer" Content="VLC Player" HorizontalAlignment="Left" Margin="190,678,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_VMWareTools" Content="VMWare Tools" HorizontalAlignment="Left" Margin="190,698,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_WinMerge" Content="WinMerge" HorizontalAlignment="Left" Margin="190,718,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_WinSCP" Content="WinSCP" HorizontalAlignment="Left" Margin="190,738,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_Wireshark" Content="Wireshark" HorizontalAlignment="Left" Margin="190,758,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <CheckBox x:Name="Checkbox_Zoom" Content="Zoom" HorizontalAlignment="Left" Margin="190,778,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_Zoom" HorizontalAlignment="Left" Margin="394,775,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="Client"/>
                        <ListBoxItem Content="Client + Citrix Plugin"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_SelectAll" Content="Select All" HorizontalAlignment="Left" Margin="160,851,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <Label x:Name="Label_author" Content="Manuel Winkel / @deyda84 / www.deyda.net / 2021 / Version $eVersion" HorizontalAlignment="Left" Margin="322,888,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2"/>
                </Grid>
            </ScrollViewer>
        </TabItem>
        <TabItem Header="Detail">
            <ScrollViewer ScrollViewer.VerticalScrollBarVisibility="Auto">
                <Grid x:Name="Detail_Grid" Margin="0,0,0,1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="13*"/>
                        <ColumnDefinition Width="234*"/>
                        <ColumnDefinition Width="586*"/>
                    </Grid.ColumnDefinitions>
                    <Image x:Name="Image_Logo_Detail" Height="100" Margin="510,3,30,0" VerticalAlignment="Top" Width="100" Source="$PSScriptRoot\img\Logo_DEYDA_no_cta.png" Grid.Column="2" ToolTip="www.deyda.net"/>
                    <Label x:Name="Label_OptionalMode" Content="Optional Mode" HorizontalAlignment="Left" Margin="11,3,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_WhatIf" Content="WhatIf" HorizontalAlignment="Left" Margin="15,34,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <CheckBox x:Name="Checkbox_Repository" Content="Installer Repository" HorizontalAlignment="Left" Margin="80,34,0,0" VerticalAlignment="Top" Grid.Column="1" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="Checkbox_CleanUp" Content="Installer CleanUp" HorizontalAlignment="Left" Margin="210,34,0,0" VerticalAlignment="Top" Grid.Column="1" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="Checkbox_CleanUpStartMenu" Content="Start Menu CleanUp" HorizontalAlignment="Left" Margin="330,34,0,0" VerticalAlignment="Top" Grid.Column="1" Grid.ColumnSpan="2"/>
                    <Button x:Name="Button_Start_Detail" Content="Start" HorizontalAlignment="Left" Margin="291,844,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
                    <Button x:Name="Button_Cancel_Detail" Content="Cancel" HorizontalAlignment="Left" Margin="386,844,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
                    <Button x:Name="Button_Save_Detail" Content="Save" HorizontalAlignment="Left" Margin="522,844,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2" ToolTip="Save Selected Software in LastSetting.txt or -GUIFile Parameter file"/>
                    <Label x:Name="Label_Software_Detail" Content="Select Software" HorizontalAlignment="Left" Margin="11,67,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <Label x:Name="Label_Architecture_Detail" Content="Architecture" HorizontalAlignment="Left" Margin="187,90,0,0" VerticalAlignment="Top" Grid.Column="1" Width="64" FontSize="10"/>
                    <Label x:Name="Label_Language_Detail" Content="Language" HorizontalAlignment="Left" Margin="260,90,0,0" VerticalAlignment="Top" Grid.Column="1" Width="75" FontSize="10" Grid.ColumnSpan="2"/>
                    <Label x:Name="Label_InstallerType_Detail" Content="Installer Type" HorizontalAlignment="Left" Margin="320,90,0,0" VerticalAlignment="Top" Grid.Column="1" Width="75" FontSize="10" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="Checkbox_7Zip_Detail" Content="7 Zip" HorizontalAlignment="Left" Margin="15,118,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_7Zip_Architecture" HorizontalAlignment="Left" Margin="215,117,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_AdobeReaderDC_Detail" Content="Adobe Reader DC" HorizontalAlignment="Left" Margin="15,143,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_AdobeReaderDC_Architecture" HorizontalAlignment="Left" Margin="215,142,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_AdobeReaderDC_Language" HorizontalAlignment="Left" Margin="265,142,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
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
                        <ListBoxItem Content="Russian"/>
                        <ListBoxItem Content="Spanish"/>
                        <ListBoxItem Content="Swedish"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_CiscoWebexTeams_Detail" Content="Cisco Webex Teams" HorizontalAlignment="Left" Margin="15,168,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_CiscoWebexTeams_Architecture" HorizontalAlignment="Left" Margin="215,167,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_CitrixHypervisorTools_Detail" Content="Citrix Hypervisor Tools" HorizontalAlignment="Left" Margin="15,193,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_CitrixHypervisorTools_Architecture" HorizontalAlignment="Left" Margin="215,192,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_ControlUpAgent_Detail" Content="ControlUp Agent" HorizontalAlignment="Left" Margin="15,218,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_ControlUpAgent_Architecture" HorizontalAlignment="Left" Margin="215,217,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_deviceTRUST_Detail" Content="deviceTRUST" HorizontalAlignment="Left" Margin="15,243,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_deviceTRUST_Architecture" HorizontalAlignment="Left" Margin="215,242,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_FoxitPDFEditor_Detail" Content="Foxit PDF Editor" HorizontalAlignment="Left" Margin="15,268,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_FoxitPDFEditor_Language" HorizontalAlignment="Left" Margin="265,267,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Danish"/>
                        <ListBoxItem Content="Dutch"/>
                        <ListBoxItem Content="English"/>
                        <ListBoxItem Content="Finnish"/>
                        <ListBoxItem Content="French"/>
                        <ListBoxItem Content="German"/>
                        <ListBoxItem Content="Italian"/>
                        <ListBoxItem Content="Korean"/>
                        <ListBoxItem Content="Norwegian"/>
                        <ListBoxItem Content="Polish"/>
                        <ListBoxItem Content="Portuguese"/>
                        <ListBoxItem Content="Russian"/>
                        <ListBoxItem Content="Spanish"/>
                        <ListBoxItem Content="Swedish"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_FoxitReader_Detail" Content="Foxit Reader" HorizontalAlignment="Left" Margin="15,293,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_FoxitReader_Language" HorizontalAlignment="Left" Margin="265,292,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Danish"/>
                        <ListBoxItem Content="Dutch"/>
                        <ListBoxItem Content="English"/>
                        <ListBoxItem Content="Finnish"/>
                        <ListBoxItem Content="French"/>
                        <ListBoxItem Content="German"/>
                        <ListBoxItem Content="Italian"/>
                        <ListBoxItem Content="Norwegian"/>
                        <ListBoxItem Content="Polish"/>
                        <ListBoxItem Content="Portuguese"/>
                        <ListBoxItem Content="Russian"/>
                        <ListBoxItem Content="Spanish"/>
                        <ListBoxItem Content="Swedish"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_GitForWindows_Detail" Content="Git for Windows" HorizontalAlignment="Left" Margin="15,318,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_GitForWindows_Architecture" HorizontalAlignment="Left" Margin="215,317,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_GoogleChrome_Detail" Content="Google Chrome" HorizontalAlignment="Left" Margin="15,343,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_GoogleChrome_Architecture" HorizontalAlignment="Left" Margin="215,342,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_ImageGlass_Detail" Content="ImageGlass" HorizontalAlignment="Left" Margin="15,368,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_ImageGlass_Architecture" HorizontalAlignment="Left" Margin="215,367,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_IrfanView_Detail" Content="IrfanView" HorizontalAlignment="Left" Margin="15,393,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_IrfanView_Architecture" HorizontalAlignment="Left" Margin="215,392,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_IrfanView_Language" HorizontalAlignment="Left" Margin="265,392,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Danish"/>
                        <ListBoxItem Content="Dutch"/>
                        <ListBoxItem Content="English"/>
                        <ListBoxItem Content="Finnish"/>
                        <ListBoxItem Content="French"/>
                        <ListBoxItem Content="German"/>
                        <ListBoxItem Content="Italian"/>
                        <ListBoxItem Content="Japanese"/>
                        <ListBoxItem Content="Korean"/>
                        <ListBoxItem Content="Polish"/>
                        <ListBoxItem Content="Portuguese"/>
                        <ListBoxItem Content="Russian"/>
                        <ListBoxItem Content="Spanish"/>
                        <ListBoxItem Content="Swedish"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_KeePass_Detail" Content="KeePass" HorizontalAlignment="Left" Margin="15,418,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_KeePass_Language" HorizontalAlignment="Left" Margin="265,417,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
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
                    <CheckBox x:Name="Checkbox_LogMeInGoToMeeting_Detail" Content="LogMeIn GoToMeeting" HorizontalAlignment="Left" Margin="15,443,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_LogMeInGoToMeeting_Installer" HorizontalAlignment="Left" Margin="68,443,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="XenApp"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSDotNetFramework_Detail" Content="Microsoft .Net Framework" HorizontalAlignment="Left" Margin="15,468,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_MSDotNetFramework_Architecture" HorizontalAlignment="Left" Margin="215,467,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MS365Apps_Detail" Content="Microsoft 365 Apps" HorizontalAlignment="Left" Margin="15,493,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MS365Apps_Architecture" HorizontalAlignment="Left" Margin="215,492,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_MS365Apps_Language" HorizontalAlignment="Left" Margin="265,492,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
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
                    <ComboBox x:Name="Box_MS365Apps_Installer" HorizontalAlignment="Left" Margin="68,492,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MS365Apps_Visio_Detail" Content="     + Microsoft Visio" HorizontalAlignment="Left" Margin="15,518,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MS365Apps_Visio_Language" HorizontalAlignment="Left" Margin="265,517,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
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
                    <CheckBox x:Name="Checkbox_MS365Apps_Project_Detail" Content="     + Microsoft Project" HorizontalAlignment="Left" Margin="15,543,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MS365Apps_Project_Language" HorizontalAlignment="Left" Margin="265,542,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
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
                    <CheckBox x:Name="Checkbox_MSAVDRemoteDesktop_Detail" Content="Microsoft AVD Remote Desktop" HorizontalAlignment="Left" Margin="15,568,0,0" VerticalAlignment="Top" Grid.Column="1" />
                    <ComboBox x:Name="Box_MSAVDRemoteDesktop_Architecture" HorizontalAlignment="Left" Margin="215,567,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSAzureDataStudio_Detail" Content="Microsoft Azure Data Studio" HorizontalAlignment="Left" Margin="15,593,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSAzureDataStudio_Installer" HorizontalAlignment="Left" Margin="68,592,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSEdge_Detail" Content="Microsoft Edge" HorizontalAlignment="Left" Margin="15,618,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSEdge_Architecture" HorizontalAlignment="Left" Margin="215,617,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSEdgeWebView2_Detail" Content="Microsoft Edge WebView2" HorizontalAlignment="Left" Margin="15,643,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSEdgeWebView2_Architecture" HorizontalAlignment="Left" Margin="215,642,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSFSLogix_Detail" Content="Microsoft FSLogix" HorizontalAlignment="Left" Margin="15,668,0,0" VerticalAlignment="Top" Grid.Column="1"/>
                    <ComboBox x:Name="Box_MSFSLogix_Architecture" HorizontalAlignment="Left" Margin="215,667,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1" Grid.ColumnSpan="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <Label x:Name="Label_Architecture2_Detail" Content="Architecture" HorizontalAlignment="Left" Margin="442,90,0,0" VerticalAlignment="Top" Grid.Column="2" Width="64" FontSize="10"/>
                    <Label x:Name="Label_Language2_Detail" Content="Language" HorizontalAlignment="Left" Margin="514,90,0,0" VerticalAlignment="Top" Grid.Column="2" Width="75" FontSize="10"/>
                    <Label x:Name="Label_InstallerType2_Detail" Content="Installer Type" HorizontalAlignment="Left" Margin="575,90,0,0" VerticalAlignment="Top" Grid.Column="2" Width="75" FontSize="10" Grid.ColumnSpan="2"/>
                    <CheckBox x:Name="Checkbox_MSOffice_Detail" Content="Microsoft Office" HorizontalAlignment="Left" Margin="179,118,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_MSOffice_Architecture" HorizontalAlignment="Left" Margin="470,117,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_MSOffice_Language" HorizontalAlignment="Left" Margin="520,117,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
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
                    <ComboBox x:Name="Box_MSOffice_Installer" HorizontalAlignment="Left" Margin="585,117,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSOneDrive_Detail" Content="Microsoft OneDrive" HorizontalAlignment="Left" Margin="179,143,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_MSOneDrive_Architecture" HorizontalAlignment="Left" Margin="470,142,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_MSOneDrive_Installer" HorizontalAlignment="Left" Margin="585,142,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSPowerBIDesktop_Detail" Content="Microsoft Power BI Desktop" HorizontalAlignment="Left" Margin="179,168,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_MSPowerBIDesktop_Architecture" HorizontalAlignment="Left" Margin="470,167,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSPowerShell_Detail" Content="Microsoft PowerShell" HorizontalAlignment="Left" Margin="179,193,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_MSPowerShell_Architecture" HorizontalAlignment="Left" Margin="470,192,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSSQLServerManagementStudio_Detail" Content="Microsoft SQL Server Management Studio" Margin="179,218,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_MSSQLServerManagementStudio_Language" HorizontalAlignment="Left" Margin="520,217,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="English"/>
                        <ListBoxItem Content="French"/>
                        <ListBoxItem Content="German"/>
                        <ListBoxItem Content="Italian"/>
                        <ListBoxItem Content="Japanese"/>
                        <ListBoxItem Content="Korean"/>
                        <ListBoxItem Content="Portuguese"/>
                        <ListBoxItem Content="Russian"/>
                        <ListBoxItem Content="Spanish"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSTeams_Detail" Content="Microsoft Teams" HorizontalAlignment="Left" Margin="179,243,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_MSTeams_Architecture" HorizontalAlignment="Left" Margin="470,242,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_MSTeams_Installer" HorizontalAlignment="Left" Margin="585,242,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MSVisualStudioCode_Detail" Content="Microsoft Visual Studio Code" HorizontalAlignment="Left" Margin="179,268,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_MSVisualStudioCode_Architecture" HorizontalAlignment="Left" Margin="470,267,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_MSVisualStudioCode_Installer" HorizontalAlignment="Left" Margin="585,267,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_MindView7_Detail" Content="MindView 7" HorizontalAlignment="Left" Margin="179,293,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_MindView7_Language" HorizontalAlignment="Left" Margin="520,292,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Danish"/>
                        <ListBoxItem Content="English"/>
                        <ListBoxItem Content="French"/>
                        <ListBoxItem Content="German"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_Firefox_Detail" Content="Mozilla Firefox" HorizontalAlignment="Left" Margin="179,318,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Firefox_Architecture" HorizontalAlignment="Left" Margin="470,317,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_Firefox_Language" HorizontalAlignment="Left" Margin="520,317,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Dutch"/>
                        <ListBoxItem Content="English"/>
                        <ListBoxItem Content="French"/>
                        <ListBoxItem Content="German"/>
                        <ListBoxItem Content="Italian"/>
                        <ListBoxItem Content="Japanese"/>
                        <ListBoxItem Content="Portuguese"/>
                        <ListBoxItem Content="Russian"/>
                        <ListBoxItem Content="Spanish"/>
                        <ListBoxItem Content="Swedish"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_NotepadPlusPlus_Detail" Content="Notepad ++" HorizontalAlignment="Left" Margin="179,343,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_NotepadPlusPlus_Architecture" HorizontalAlignment="Left" Margin="470,342,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_OpenJDK_Detail" Content="Open JDK" HorizontalAlignment="Left" Margin="179,368,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_OpenJDK_Architecture" HorizontalAlignment="Left" Margin="470,367,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_OracleJava8_Detail" Content="Oracle Java 8" HorizontalAlignment="Left" Margin="179,393,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_OracleJava8_Architecture" HorizontalAlignment="Left" Margin="470,392,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_PeaZip_Detail" Content="PeaZip" HorizontalAlignment="Left" Margin="179,418,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_PeaZip_Architecture" HorizontalAlignment="Left" Margin="470,417,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_Putty_Detail" Content="PuTTY" HorizontalAlignment="Left" Margin="179,443,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Putty_Architecture" HorizontalAlignment="Left" Margin="470,442,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_Slack_Detail" Content="Slack" HorizontalAlignment="Left" Margin="179,468,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_Slack_Architecture" HorizontalAlignment="Left" Margin="470,467,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <ComboBox x:Name="Box_Slack_Installer" HorizontalAlignment="Left" Margin="585,467,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_SumatraPDF_Detail" Content="Sumatra PDF" HorizontalAlignment="Left" Margin="179,493,0,0" VerticalAlignment="Top" Grid.Column="2" />
                    <ComboBox x:Name="Box_SumatraPDF_Architecture" HorizontalAlignment="Left" Margin="470,492,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_TechSmithSnagIt_Detail" Content="TechSmith SnagIt" HorizontalAlignment="Left" Margin="179,518,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_TechSmithSnagIt_Architecture" HorizontalAlignment="Left" Margin="470,517,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_TotalCommander_Detail" Content="Total Commander" HorizontalAlignment="Left" Margin="179,543,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_TotalCommander_Architecture" HorizontalAlignment="Left" Margin="470,542,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_VLCPlayer_Detail" Content="VLC Player" HorizontalAlignment="Left" Margin="179,568,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_VLCPlayer_Architecture" HorizontalAlignment="Left" Margin="470,567,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_VMWareTools_Detail" Content="VMWare Tools" HorizontalAlignment="Left" Margin="179,593,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_VMWareTools_Architecture" HorizontalAlignment="Left" Margin="470,592,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_WinMerge_Detail" Content="WinMerge" HorizontalAlignment="Left" Margin="179,618,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_WinMerge_Architecture" HorizontalAlignment="Left" Margin="470,617,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_Wireshark_Detail" Content="Wireshark" HorizontalAlignment="Left" Margin="179,643,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Wireshark_Architecture" HorizontalAlignment="Left" Margin="470,642,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="x86"/>
                        <ListBoxItem Content="x64"/>
                    </ComboBox>
                    <CheckBox x:Name="Checkbox_Zoom_Detail" Content="Zoom" HorizontalAlignment="Left" Margin="179,668,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <ComboBox x:Name="Box_Zoom_Installer" HorizontalAlignment="Left" Margin="585,667,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
                        <ListBoxItem Content="-"/>
                        <ListBoxItem Content="Machine Based"/>
                        <ListBoxItem Content="User Based"/>
                    </ComboBox>
                    <Label x:Name="Label_MS365Apps_XML" Content="Custom Microsoft 365 Apps XML File" HorizontalAlignment="Left" Margin="170,732,0,0" VerticalAlignment="Top" Grid.Column="2"/>
                    <TextBox HorizontalAlignment="Left" Height="40" Margin="172,763,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="405" Name="TextBox_FileName" Grid.Column="2"/>
                    <Label x:Name="Label_MS365Apps_Text" Content="When a custom file is uploaded, the settings for the XML made here will be overwritten." HorizontalAlignment="Left" Margin="172,803,0,0" VerticalAlignment="Top" Grid.Column="2" FontSize="10" Width="405"/>
                    <Button x:Name="Button_Browse" Content="Browse" HorizontalAlignment="Left" Margin="390,733,0,0" VerticalAlignment="Top" Width="90" RenderTransformOrigin="1.047,0.821" Height="26" Grid.Column="2"/>
                    <Label x:Name="Label_author_Detail" Content="Manuel Winkel / @deyda84 / www.deyda.net / 2021 / Version $eVersion" HorizontalAlignment="Left" Margin="322,888,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2"/>
                </Grid>
            </ScrollViewer>
        </TabItem>
    </TabControl>
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
    #$Script:install = $true
    #$Script:download = $true

    # Read LastSetting.txt or -GUIFile Parameter file to get the settings of the last session. (AddScript)
    If (!$GUIfile) {$GUIfile = "LastSetting.txt"}
    If (Test-Path "$PSScriptRoot\$GUIfile" -PathType leaf) {
        $LastSetting = Get-Content "$PSScriptRoot\$GUIfile"
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
        $WPFBox_Zoom.SelectedIndex = $LastSetting[49] -as [int]
        $WPFBox_deviceTRUST.SelectedIndex = $LastSetting[50] -as [int]
        $WPFBox_MSEdge.SelectedIndex = $LastSetting[51] -as [int]
        $WPFBox_MSVisualStudioCode.SelectedIndex = $LastSetting[56] -as [int]
        $WPFBox_Installer.SelectedIndex = $LastSetting[60] -as [int]
        $WPFBox_MSVisualStudio.SelectedIndex = $LastSetting[61] -as [int]
        $WPFBox_Putty.SelectedIndex = $LastSetting[62] -as [int]
        $WPFBox_MSAzureDataStudio.SelectedIndex = $LastSetting[64] -as [int]
        $WPFBox_MSFSLogix.SelectedIndex = $LastSetting[66] -as [int]
        $WPFBox_ControlUpAgent.SelectedIndex = $LastSetting[71] -as [int]
        $WPFBox_MSAVDRemoteDesktop.SelectedIndex = $LastSetting[75] -as [int]
        $WPFBox_7Zip_Architecture.SelectedIndex = $LastSetting[93] -as [int]
        $WPFBox_AdobeReaderDC_Architecture.SelectedIndex = $LastSetting[94] -as [int]
        $WPFBox_AdobeReaderDC_Language.SelectedIndex = $LastSetting[95] -as [int]
        $WPFBox_CiscoWebexTeams_Architecture.SelectedIndex = $LastSetting[96] -as [int]
        $WPFBox_CitrixHypervisorTools_Architecture.SelectedIndex = $LastSetting[97] -as [int]
        $WPFBox_ControlUpAgent_Architecture.SelectedIndex = $LastSetting[98] -as [int]
        $WPFBox_deviceTRUST_Architecture.SelectedIndex = $LastSetting[99] -as [int]
        $WPFBox_FoxitPDFEditor_Language.SelectedIndex = $LastSetting[100] -as [int]
        $WPFBox_FoxitReader_Language.SelectedIndex = $LastSetting[101] -as [int]
        $WPFBox_GitForWindows_Architecture.SelectedIndex = $LastSetting[102] -as [int]
        $WPFBox_GoogleChrome_Architecture.SelectedIndex = $LastSetting[103] -as [int]
        $WPFBox_ImageGlass_Architecture.SelectedIndex = $LastSetting[104] -as [int]
        $WPFBox_IrfanView_Architecture.SelectedIndex = $LastSetting[105] -as [int]
        $WPFBox_KeePass_Language.SelectedIndex = $LastSetting[106] -as [int]
        $WPFBox_MSDotNetFramework_Architecture.SelectedIndex = $LastSetting[107] -as [int]
        $WPFBox_MS365Apps_Architecture.SelectedIndex = $LastSetting[108] -as [int]
        $WPFBox_MS365Apps_Language.SelectedIndex = $LastSetting[109] -as [int]
        $WPFBox_MS365Apps_Visio_Language.SelectedIndex = $LastSetting[111] -as [int]
        $WPFBox_MS365Apps_Project_Language.SelectedIndex = $LastSetting[113] -as [int]
        $WPFBox_MSAVDRemoteDesktop_Architecture.SelectedIndex = $LastSetting[114] -as [int]
        $WPFBox_MSEdge_Architecture.SelectedIndex = $LastSetting[115] -as [int]
        $WPFBox_MSFSLogix_Architecture.SelectedIndex = $LastSetting[116] -as [int]
        $WPFBox_MSOffice_Architecture.SelectedIndex = $LastSetting[117] -as [int]
        $WPFBox_MSOneDrive_Architecture.SelectedIndex = $LastSetting[118] -as [int]
        $WPFBox_MSPowerBIDesktop_Architecture.SelectedIndex = $LastSetting[119] -as [int]
        $WPFBox_MSPowerShell_Architecture.SelectedIndex = $LastSetting[120] -as [int]
        $WPFBox_MSSQLServerManagementStudio_Language.SelectedIndex = $LastSetting[121] -as [int]
        $WPFBox_MSTeams_Architecture.SelectedIndex = $LastSetting[122] -as [int]
        $WPFBox_MSVisualStudioCode_Architecture.SelectedIndex = $LastSetting[123] -as [int]
        $WPFBox_Firefox_Architecture.SelectedIndex = $LastSetting[124] -as [int]
        $WPFBox_Firefox_Language.SelectedIndex = $LastSetting[125] -as [int]
        $WPFBox_NotepadPlusPlus_Architecture.SelectedIndex = $LastSetting[126] -as [int]
        $WPFBox_OpenJDK_Architecture.SelectedIndex = $LastSetting[127] -as [int]
        $WPFBox_OracleJava8_Architecture.SelectedIndex = $LastSetting[128] -as [int]
        $WPFBox_PeaZip_Architecture.SelectedIndex = $LastSetting[129] -as [int]
        $WPFBox_Putty_Architecture.SelectedIndex = $LastSetting[130] -as [int]
        $WPFBox_Slack_Architecture.SelectedIndex = $LastSetting[131] -as [int]
        $WPFBox_SumatraPDF_Architecture.SelectedIndex = $LastSetting[132] -as [int]
        $WPFBox_TechSmithSnagIT_Architecture.SelectedIndex = $LastSetting[133] -as [int]
        $WPFBox_VLCPlayer_Architecture.SelectedIndex = $LastSetting[134] -as [int]
        $WPFBox_VMWareTools_Architecture.SelectedIndex = $LastSetting[135] -as [int]
        $WPFBox_WinMerge_Architecture.SelectedIndex = $LastSetting[136] -as [int]
        $WPFBox_Wireshark_Architecture.SelectedIndex = $LastSetting[137] -as [int]
        $WPFBox_IrfanView_Language.SelectedIndex = $LastSetting[138] -as [int]
        $WPFBox_MSOffice_Language.SelectedIndex = $LastSetting[139] -as [int]
        $WPFBox_MSEdgeWebView2_Architecture.SelectedIndex = $LastSetting[141] -as [int]
        $WPFBox_MindView7_Language.SelectedIndex = $LastSetting[144] -as [int]
        $WPFBox_MSOffice.SelectedIndex = $LastSetting[146] -as [int]
        $WPFBox_LogMeInGoToMeeting_Installer.SelectedIndex = $LastSetting[150] -as [int]
        $WPFBox_MSAzureDataStudio_Installer.SelectedIndex = $LastSetting[151] -as [int]
        $WPFBox_MSVisualStudioCode_Installer.SelectedIndex = $LastSetting[152] -as [int]
        $WPFBox_MS365Apps_Installer.SelectedIndex = $LastSetting[153] -as [int]
        $WPFBox_MSOffice_Installer.SelectedIndex = $LastSetting[154] -as [int]
        $WPFBox_MSTeams_Installer.SelectedIndex = $LastSetting[155] -as [int]
        $WPFBox_Zoom_Installer.SelectedIndex = $LastSetting[156] -as [int]
        $WPFBox_MSOneDrive_Installer.SelectedIndex = $LastSetting[157] -as [int]
        $WPFBox_Slack_Installer.SelectedIndex = $LastSetting[158] -as [int]
        $WPFBox_pdfforgePDFCreator.SelectedIndex = $LastSetting[159] -as [int]
        $WPFBox_TotalCommander_Architecture.SelectedIndex = $LastSetting[160] -as [int]

        Switch ($LastSetting[8]) {
            1 {
                $WPFCheckbox_7ZIP.IsChecked = "True"
                $WPFCheckbox_7ZIP_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[9]) {
            1 {$WPFCheckbox_AdobeProDC.IsChecked = "True"}
        }
        Switch ($LastSetting[10]) {
            1 {
                $WPFCheckbox_AdobeReaderDC.IsChecked = "True"
                $WPFCheckbox_AdobeReaderDC_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[11]) {
            1 { $WPFCheckbox_BISF.IsChecked = "True"}
        }
        Switch ($LastSetting[12]) {
            1 {
                $WPFCheckbox_CitrixHypervisorTools.IsChecked = "True"
                $WPFCheckbox_CitrixHypervisorTools_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[13]) {
            1 { $WPFCheckbox_CitrixWorkspaceApp.IsChecked = "True"}
        }
        Switch ($LastSetting[14]) {
            1 { $WPFCheckbox_Filezilla.IsChecked = "True"}
        }
        Switch ($LastSetting[15]) {
            1 {
                $WPFCheckbox_Firefox.IsChecked = "True"
                $WPFCheckbox_Firefox_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[16]) {
            1 {
                $WPFCheckbox_FoxitReader.IsChecked = "True"
                $WPFCheckbox_FoxitReader_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[17]) {
            1 {
                $WPFCheckbox_MSFSLogix.IsChecked = "True"
                $WPFCheckbox_MSFSLogix_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[18]) {
            1 {
                $WPFCheckbox_GoogleChrome.IsChecked = "True"
                $WPFCheckbox_GoogleChrome_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[19]) {
            1 { $WPFCheckbox_Greenshot.IsChecked = "True"}
        }
        Switch ($LastSetting[20]) {
            1 {
                $WPFCheckbox_KeePass.IsChecked = "True"
                $WPFCheckbox_KeePass_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[21]) {
            1 { $WPFCheckbox_mRemoteNG.IsChecked = "True"}
        }
        Switch ($LastSetting[22]) {
            1 {
                $WPFCheckbox_MS365Apps.IsChecked = "True"
                $WPFCheckbox_MS365Apps_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[23]) {
            1 {
                $WPFCheckbox_MSEdge.IsChecked = "True"
                $WPFCheckbox_MSEdge_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[24]) {
            1 {
                $WPFCheckbox_MSOffice.IsChecked = "True"
                $WPFCheckbox_MSOffice_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[25]) {
            1 {
                $WPFCheckbox_MSOneDrive.IsChecked = "True"
                $WPFCheckbox_MSOneDrive_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[26]) {
            1 {
                $WPFCheckbox_MSTeams.IsChecked = "True"
                $WPFCheckbox_MSTeams_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[27]) {
            1 {
                $WPFCheckbox_NotePadPlusPlus.IsChecked = "True"
                $WPFCheckbox_NotepadPlusPlus_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[28]) {
            1 {
                $WPFCheckbox_OpenJDK.IsChecked = "True"
                $WPFCheckbox_OpenJDK_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[29]) {
            1 {
                $WPFCheckbox_OracleJava8.IsChecked = "True"
                $WPFCheckbox_OracleJava8_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[30]) {
            1 { $WPFCheckbox_TreeSize.IsChecked = "True"}
        }
        Switch ($LastSetting[31]) {
            1 {
                $WPFCheckbox_VLCPlayer.IsChecked = "True"
                $WPFCheckbox_VLCPlayer_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[32]) {
            1 {
                $WPFCheckbox_VMWareTools.IsChecked = "True"
                $WPFCheckbox_VMWareTools_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[33]) {
            1 { $WPFCheckbox_WinSCP.IsChecked = "True"}
        }
        Switch ($LastSetting[34]) {
            True { $WPFCheckbox_Download.IsChecked = "True"}
            1 { $WPFCheckbox_Download.IsChecked = "True"}
        }
        Switch ($LastSetting[35]) {
            True { $WPFCheckbox_Install.IsChecked = "True"}
            1 { $WPFCheckbox_Install.IsChecked = "True"}
        }
        Switch ($LastSetting[36]) {
            1 {
                $WPFCheckbox_IrfanView.IsChecked = "True"
                $WPFCheckbox_IrfanView_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[37]) {
            1 { $WPFCheckbox_MSTeams_No_AutoStart.IsChecked = "True"}
        }
        Switch ($LastSetting[38]) {
            1 {
                $WPFCheckbox_deviceTRUST.IsChecked = "True"
                $WPFCheckbox_deviceTRUST_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[39]) {
            1 {
                $WPFCheckbox_MSDotNetFramework.IsChecked = "True"
                $WPFCheckbox_MSDotNetFramework_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[41]) {
            1 {
                $WPFCheckbox_MSPowerShell.IsChecked = "True"
                $WPFCheckbox_MSPowerShell_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[43]) {
            1 { $WPFCheckbox_RemoteDesktopManager.IsChecked = "True"}
        }
        Switch ($LastSetting[45]) {
            1 {
                $WPFCheckbox_Slack.IsChecked = "True"
                $WPFCheckbox_Slack_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[46]) {
            1 {
                $WPFCheckbox_Wireshark.IsChecked = "True"
                $WPFCheckbox_Wireshark_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[47]) {
            1 { $WPFCheckbox_ShareX.IsChecked = "True"}
        }
        Switch ($LastSetting[48]) {
            1 { $WPFCheckbox_Zoom.IsChecked = "True"}
        }
        Switch ($LastSetting[52]) {
            1 { $WPFCheckbox_GIMP.IsChecked = "True"}
        }
        Switch ($LastSetting[53]) {
            1 { $WPFCheckbox_MSPowerToys.IsChecked = "True"}
        }
        Switch ($LastSetting[54]) {
            1 { $WPFCheckbox_MSVisualStudio.IsChecked = "True"}
        }
        Switch ($LastSetting[55]) {
            1 {
                $WPFCheckbox_MSVisualStudioCode.IsChecked = "True"
                $WPFCheckbox_MSVisualStudioCode_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[57]) {
            1 { $WPFCheckbox_PaintDotNet.IsChecked = "True"}
        }
        Switch ($LastSetting[58]) {
            1 {
                $WPFCheckbox_Putty.IsChecked = "True"
                $WPFCheckbox_Putty_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[59]) {
            1 { $WPFCheckbox_TeamViewer.IsChecked = "True"}
        }
        Switch ($LastSetting[63]) {
            1 { $WPFCheckbox_MSAzureDataStudio.IsChecked = "True"}
        }
        Switch ($LastSetting[65]) {
            1 {
                $WPFCheckbox_ImageGlass.IsChecked = "True"
                $WPFCheckbox_ImageGlass_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[67]) {
            1 { $WPFCheckbox_uberAgent.IsChecked = "True"}
        }
        Switch ($LastSetting[68]) {
            1 { $WPFCheckbox_1Password.IsChecked = "True"}
        }
        Switch ($LastSetting[69]) {
            1 {
                $WPFCheckbox_SumatraPDF.IsChecked = "True"
                $WPFCheckbox_SumatraPDF_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[70]) {
            1 {
                $WPFCheckbox_ControlUpAgent.IsChecked = "True"
                $WPFCheckbox_ControlUpAgent_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[72]) {
            1 { $WPFCheckbox_ControlUpConsole.IsChecked = "True"}
        }
        Switch ($LastSetting[73]) {
            1 {
                $WPFCheckbox_MSSQLServerManagementStudio.IsChecked = "True"
                $WPFCheckbox_MSSQLServerManagementStudio_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[74]) {
            1 {
                $WPFCheckbox_MSAVDRemoteDesktop.IsChecked = "True"
                $WPFCheckbox_MSAVDRemoteDesktop_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[76]) {
            1 {
                $WPFCheckbox_MSPowerBIDesktop.IsChecked = "True"
                $WPFCheckbox_MSPowerBIDesktop_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[77]) {
            1 { $WPFCheckbox_RDAnalyzer.IsChecked = "True"}
        }
        Switch ($LastSetting[78]) {
            1 {
                $WPFCheckbox_CiscoWebexTeams.IsChecked = "True"
                $WPFCheckbox_CiscoWebexTeams_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[79]) {
            1 { $WPFCheckbox_CitrixFiles.IsChecked = "True"}
        }
        Switch ($LastSetting[80]) {
            1 {
                $WPFCheckbox_FoxitPDFEditor.IsChecked = "True"
                $WPFCheckbox_FoxitPDFEditor_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[81]) {
            1 {
                $WPFCheckbox_GitForWindows.IsChecked = "True"
                $WPFCheckbox_GitForWindows_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[82]) {
            1 { $WPFCheckbox_LogMeInGoToMeeting.IsChecked = "True"}
        }
        Switch ($LastSetting[83]) {
            1 { $WPFCheckbox_MSAzureCLI.IsChecked = "True"}
        }
        Switch ($LastSetting[84]) {
            1 { $WPFCheckbox_MSPowerBIReportBuilder.IsChecked = "True"}
        }
        Switch ($LastSetting[85]) {
            1 { $WPFCheckbox_MSSysinternals.IsChecked = "True"}
        }
        Switch ($LastSetting[86]) {
            1 { $WPFCheckbox_Nmap.IsChecked = "True"}
        }
        Switch ($LastSetting[87]) {
            1 {
                $WPFCheckbox_PeaZip.IsChecked = "True"
                $WPFCheckbox_PeaZip_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[88]) {
            1 { $WPFCheckbox_TechSmithCamtasia.IsChecked = "True"}
        }
        Switch ($LastSetting[89]) {
            1 {
                $WPFCheckbox_TechSmithSnagIt.IsChecked = "True"
                $WPFCheckbox_TechSmithSnagIt_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[90]) {
            1 {
                $WPFCheckbox_WinMerge.IsChecked = "True"
                $WPFCheckbox_WinMerge_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[91]) {
            1 { $WPFCheckbox_WhatIf.IsChecked = "True"}
        }
        Switch ($LastSetting[92]) {
            1 { $WPFCheckbox_CleanUp.IsChecked = "True"}
        }
        Switch ($LastSetting[110]) {
            1 { $WPFCheckbox_MS365Apps_Visio_Detail.IsChecked = "True"}
        }
        Switch ($LastSetting[112]) {
            1 { $WPFCheckbox_MS365Apps_Project_Detail.IsChecked = "True"}
        }
        Switch ($LastSetting[140]) {
            1 {
                $WPFCheckbox_MSEdgeWebView2.IsChecked = "True"
                $WPFCheckbox_MSEdgeWebView2_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[142]) {
            1 {
                $WPFCheckbox_AutodeskDWGTrueView.IsChecked = "True"
            }
        }
        Switch ($LastSetting[143]) {
            1 {
                $WPFCheckbox_MindView7.IsChecked = "True"
                $WPFCheckbox_MindView7_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[145]) {
            1 {
                $WPFCheckbox_PDFsam.IsChecked = "True"
            }
        }
        Switch ($LastSetting[147]) {
            1 {
                $WPFCheckbox_OpenShellMenu.IsChecked = "True"
            }
        }
        Switch ($LastSetting[148]) {
            1 {
                $WPFCheckbox_PDFForgeCreator.IsChecked = "True"
            }
        }
        Switch ($LastSetting[149]) {
            1 {
                $WPFCheckbox_TotalCommander.IsChecked = "True"
                $WPFCheckbox_TotalCommander_Detail.IsChecked = "True"
            }
        }
        Switch ($LastSetting[161]) {
            1 { $WPFCheckbox_Repository.IsChecked = "True"}
        }
        Switch ($LastSetting[162]) {
            1 { $WPFCheckbox_CleanUpStartMenu.IsChecked = "True"}
        }
    }
    
    #// MARK: Event Handler
    # Checkbox Detail (AddScript)
    $WPFCheckbox_7ZIP.Add_Checked({
        $WPFCheckbox_7Zip_Detail.IsChecked = $WPFCheckbox_7ZIP.IsChecked
    })
    $WPFCheckbox_7ZIP.Add_Unchecked({
        $WPFCheckbox_7Zip_Detail.IsChecked = $WPFCheckbox_7Zip.IsChecked
    })
    $WPFCheckbox_7ZIP_Detail.Add_Checked({
        $WPFCheckbox_7Zip.IsChecked = $WPFCheckbox_7ZIP_Detail.IsChecked
    })
    $WPFCheckbox_7ZIP_Detail.Add_Unchecked({
        $WPFCheckbox_7Zip.IsChecked = $WPFCheckbox_7Zip_Detail.IsChecked
    })
    $WPFCheckbox_AdobeReaderDC.Add_Checked({
        $WPFCheckbox_AdobeReaderDC_Detail.IsChecked = $WPFCheckbox_AdobeReaderDC.IsChecked
    })
    $WPFCheckbox_AdobeReaderDC.Add_Unchecked({
        $WPFCheckbox_AdobeReaderDC_Detail.IsChecked = $WPFCheckbox_AdobeReaderDC.IsChecked
    })
    $WPFCheckbox_AdobeReaderDC_Detail.Add_Checked({
        $WPFCheckbox_AdobeReaderDC.IsChecked = $WPFCheckbox_AdobeReaderDC_Detail.IsChecked
    })
    $WPFCheckbox_AdobeReaderDC_Detail.Add_Unchecked({
        $WPFCheckbox_AdobeReaderDC.IsChecked = $WPFCheckbox_AdobeReaderDC_Detail.IsChecked
    })
    $WPFCheckbox_CiscoWebexTeams.Add_Checked({
        $WPFCheckbox_CiscoWebexTeams_Detail.IsChecked = $WPFCheckbox_CiscoWebexTeams.IsChecked
    })
    $WPFCheckbox_CiscoWebexTeams.Add_Unchecked({
        $WPFCheckbox_CiscoWebexTeams_Detail.IsChecked = $WPFCheckbox_CiscoWebexTeams.IsChecked
    })
    $WPFCheckbox_CiscoWebexTeams_Detail.Add_Checked({
        $WPFCheckbox_CiscoWebexTeams.IsChecked = $WPFCheckbox_CiscoWebexTeams_Detail.IsChecked
    })
    $WPFCheckbox_CiscoWebexTeams_Detail.Add_Unchecked({
        $WPFCheckbox_CiscoWebexTeams.IsChecked = $WPFCheckbox_CiscoWebexTeams_Detail.IsChecked
    })
    $WPFCheckbox_CitrixHypervisorTools.Add_Checked({
        $WPFCheckbox_CitrixHypervisorTools_Detail.IsChecked = $WPFCheckbox_CitrixHypervisorTools.IsChecked
    })
    $WPFCheckbox_CitrixHypervisorTools.Add_Unchecked({
        $WPFCheckbox_CitrixHypervisorTools_Detail.IsChecked = $WPFCheckbox_CitrixHypervisorTools.IsChecked
    })
    $WPFCheckbox_CitrixHypervisorTools_Detail.Add_Checked({
        $WPFCheckbox_CitrixHypervisorTools.IsChecked = $WPFCheckbox_CitrixHypervisorTools_Detail.IsChecked
    })
    $WPFCheckbox_CitrixHypervisorTools_Detail.Add_Unchecked({
        $WPFCheckbox_CitrixHypervisorTools.IsChecked = $WPFCheckbox_CitrixHypervisorTools_Detail.IsChecked
    })
    $WPFCheckbox_deviceTRUST.Add_Checked({
        $WPFCheckbox_deviceTRUST_Detail.IsChecked = $WPFCheckbox_deviceTRUST.IsChecked
    })
    $WPFCheckbox_deviceTRUST.Add_Unchecked({
        $WPFCheckbox_deviceTRUST_Detail.IsChecked = $WPFCheckbox_deviceTRUST.IsChecked
    })
    $WPFCheckbox_deviceTRUST_Detail.Add_Checked({
        $WPFCheckbox_deviceTRUST.IsChecked = $WPFCheckbox_deviceTRUST_Detail.IsChecked
    })
    $WPFCheckbox_deviceTRUST_Detail.Add_Unchecked({
        $WPFCheckbox_deviceTRUST.IsChecked = $WPFCheckbox_deviceTRUST_Detail.IsChecked
    })
    $WPFCheckbox_ControlUpAgent.Add_Checked({
        $WPFCheckbox_ControlUpAgent_Detail.IsChecked = $WPFCheckbox_ControlUpAgent.IsChecked
    })
    $WPFCheckbox_ControlUpAgent.Add_Unchecked({
        $WPFCheckbox_ControlUpAgent_Detail.IsChecked = $WPFCheckbox_ControlUpAgent.IsChecked
    })
    $WPFCheckbox_ControlUpAgent_Detail.Add_Checked({
        $WPFCheckbox_ControlUpAgent.IsChecked = $WPFCheckbox_ControlUpAgent_Detail.IsChecked
    })
    $WPFCheckbox_ControlUpAgent_Detail.Add_Unchecked({
        $WPFCheckbox_ControlUpAgent.IsChecked = $WPFCheckbox_ControlUpAgent_Detail.IsChecked
    })
    $WPFCheckbox_FoxitPDFEditor.Add_Checked({
        $WPFCheckbox_FoxitPDFEditor_Detail.IsChecked = $WPFCheckbox_FoxitPDFEditor.IsChecked
    })
    $WPFCheckbox_FoxitPDFEditor.Add_Unchecked({
        $WPFCheckbox_FoxitPDFEditor_Detail.IsChecked = $WPFCheckbox_FoxitPDFEditor.IsChecked
    })
    $WPFCheckbox_FoxitPDFEditor_Detail.Add_Checked({
        $WPFCheckbox_FoxitPDFEditor.IsChecked = $WPFCheckbox_FoxitPDFEditor_Detail.IsChecked
    })
    $WPFCheckbox_FoxitPDFEditor_Detail.Add_Unchecked({
        $WPFCheckbox_FoxitPDFEditor.IsChecked = $WPFCheckbox_FoxitPDFEditor_Detail.IsChecked
    })
    $WPFCheckbox_FoxitReader.Add_Checked({
        $WPFCheckbox_FoxitReader_Detail.IsChecked = $WPFCheckbox_FoxitReader.IsChecked
    })
    $WPFCheckbox_FoxitReader.Add_Unchecked({
        $WPFCheckbox_FoxitReader_Detail.IsChecked = $WPFCheckbox_FoxitReader.IsChecked
    })
    $WPFCheckbox_FoxitReader_Detail.Add_Checked({
        $WPFCheckbox_FoxitReader.IsChecked = $WPFCheckbox_FoxitReader_Detail.IsChecked
    })
    $WPFCheckbox_FoxitReader_Detail.Add_Unchecked({
        $WPFCheckbox_FoxitReader.IsChecked = $WPFCheckbox_FoxitReader_Detail.IsChecked
    })
    $WPFCheckbox_GitForWindows.Add_Checked({
        $WPFCheckbox_GitForWindows_Detail.IsChecked = $WPFCheckbox_GitForWindows.IsChecked
    })
    $WPFCheckbox_GitForWindows.Add_Unchecked({
        $WPFCheckbox_GitForWindows_Detail.IsChecked = $WPFCheckbox_GitForWindows.IsChecked
    })
    $WPFCheckbox_GitForWindows_Detail.Add_Checked({
        $WPFCheckbox_GitForWindows.IsChecked = $WPFCheckbox_GitForWindows_Detail.IsChecked
    })
    $WPFCheckbox_GitForWindows_Detail.Add_Unchecked({
        $WPFCheckbox_GitForWindows.IsChecked = $WPFCheckbox_GitForWindows_Detail.IsChecked
    })
    $WPFCheckbox_GoogleChrome.Add_Checked({
        $WPFCheckbox_GoogleChrome_Detail.IsChecked = $WPFCheckbox_GoogleChrome.IsChecked
    })
    $WPFCheckbox_GoogleChrome.Add_Unchecked({
        $WPFCheckbox_GoogleChrome_Detail.IsChecked = $WPFCheckbox_GoogleChrome.IsChecked
    })
    $WPFCheckbox_GoogleChrome_Detail.Add_Checked({
        $WPFCheckbox_GoogleChrome.IsChecked = $WPFCheckbox_GoogleChrome_Detail.IsChecked
    })
    $WPFCheckbox_GoogleChrome_Detail.Add_Unchecked({
        $WPFCheckbox_GoogleChrome.IsChecked = $WPFCheckbox_GoogleChrome_Detail.IsChecked
    })
    $WPFCheckbox_ImageGlass.Add_Checked({
        $WPFCheckbox_ImageGlass_Detail.IsChecked = $WPFCheckbox_ImageGlass.IsChecked
    })
    $WPFCheckbox_ImageGlass.Add_Unchecked({
        $WPFCheckbox_ImageGlass_Detail.IsChecked = $WPFCheckbox_ImageGlass.IsChecked
    })
    $WPFCheckbox_ImageGlass_Detail.Add_Checked({
        $WPFCheckbox_ImageGlass.IsChecked = $WPFCheckbox_ImageGlass_Detail.IsChecked
    })
    $WPFCheckbox_ImageGlass_Detail.Add_Unchecked({
        $WPFCheckbox_ImageGlass.IsChecked = $WPFCheckbox_ImageGlass_Detail.IsChecked
    })
    $WPFCheckbox_IrfanView.Add_Checked({
        $WPFCheckbox_IrfanView_Detail.IsChecked = $WPFCheckbox_IrfanView.IsChecked
    })
    $WPFCheckbox_IrfanView.Add_Unchecked({
        $WPFCheckbox_IrfanView_Detail.IsChecked = $WPFCheckbox_IrfanView.IsChecked
    })
    $WPFCheckbox_IrfanView_Detail.Add_Checked({
        $WPFCheckbox_IrfanView.IsChecked = $WPFCheckbox_IrfanView_Detail.IsChecked
    })
    $WPFCheckbox_IrfanView_Detail.Add_Unchecked({
        $WPFCheckbox_IrfanView.IsChecked = $WPFCheckbox_IrfanView_Detail.IsChecked
    })
    $WPFCheckbox_KeePass.Add_Checked({
        $WPFCheckbox_KeePass_Detail.IsChecked = $WPFCheckbox_KeePass.IsChecked
    })
    $WPFCheckbox_KeePass.Add_Unchecked({
        $WPFCheckbox_KeePass_Detail.IsChecked = $WPFCheckbox_KeePass.IsChecked
    })
    $WPFCheckbox_KeePass_Detail.Add_Checked({
        $WPFCheckbox_KeePass.IsChecked = $WPFCheckbox_KeePass_Detail.IsChecked
    })
    $WPFCheckbox_KeePass_Detail.Add_Unchecked({
        $WPFCheckbox_KeePass.IsChecked = $WPFCheckbox_KeePass_Detail.IsChecked
    })
    $WPFCheckbox_LogMeInGoToMeeting.Add_Checked({
        $WPFCheckbox_LogMeInGoToMeeting_Detail.IsChecked = $WPFCheckbox_LogMeInGoToMeeting.IsChecked
    })
    $WPFCheckbox_LogMeInGoToMeeting.Add_Unchecked({
        $WPFCheckbox_LogMeInGoToMeeting_Detail.IsChecked = $WPFCheckbox_LogMeInGoToMeeting.IsChecked
    })
    $WPFCheckbox_LogMeInGoToMeeting_Detail.Add_Checked({
        $WPFCheckbox_LogMeInGoToMeeting.IsChecked = $WPFCheckbox_LogMeInGoToMeeting_Detail.IsChecked
    })
    $WPFCheckbox_LogMeInGoToMeeting_Detail.Add_Unchecked({
        $WPFCheckbox_LogMeInGoToMeeting.IsChecked = $WPFCheckbox_LogMeInGoToMeeting_Detail.IsChecked
    })
    $WPFCheckbox_MSDotNetFramework.Add_Checked({
        $WPFCheckbox_MSDotNetFramework_Detail.IsChecked = $WPFCheckbox_MSDotNetFramework.IsChecked
    })
    $WPFCheckbox_MSDotNetFramework.Add_Unchecked({
        $WPFCheckbox_MSDotNetFramework_Detail.IsChecked = $WPFCheckbox_MSDotNetFramework.IsChecked
    })
    $WPFCheckbox_MSDotNetFramework_Detail.Add_Checked({
        $WPFCheckbox_MSDotNetFramework.IsChecked = $WPFCheckbox_MSDotNetFramework_Detail.IsChecked
    })
    $WPFCheckbox_MSDotNetFramework_Detail.Add_Unchecked({
        $WPFCheckbox_MSDotNetFramework.IsChecked = $WPFCheckbox_MSDotNetFramework_Detail.IsChecked
    })
    $WPFCheckbox_MS365Apps.Add_Checked({
        $WPFCheckbox_MS365Apps_Detail.IsChecked = $WPFCheckbox_MS365Apps.IsChecked
    })
    $WPFCheckbox_MS365Apps.Add_Unchecked({
        $WPFCheckbox_MS365Apps_Detail.IsChecked = $WPFCheckbox_MS365Apps.IsChecked
    })
    $WPFCheckbox_MS365Apps_Detail.Add_Checked({
        $WPFCheckbox_MS365Apps.IsChecked = $WPFCheckbox_MS365Apps_Detail.IsChecked
    })
    $WPFCheckbox_MS365Apps_Detail.Add_Unchecked({
        $WPFCheckbox_MS365Apps.IsChecked = $WPFCheckbox_MS365Apps_Detail.IsChecked
    })
    $WPFCheckbox_MSAVDRemoteDesktop.Add_Checked({
        $WPFCheckbox_MSAVDRemoteDesktop_Detail.IsChecked = $WPFCheckbox_MSAVDRemoteDesktop.IsChecked
    })
    $WPFCheckbox_MSAVDRemoteDesktop.Add_Unchecked({
        $WPFCheckbox_MSAVDRemoteDesktop_Detail.IsChecked = $WPFCheckbox_MSAVDRemoteDesktop.IsChecked
    })
    $WPFCheckbox_MSAVDRemoteDesktop_Detail.Add_Checked({
        $WPFCheckbox_MSAVDRemoteDesktop.IsChecked = $WPFCheckbox_MSAVDRemoteDesktop_Detail.IsChecked
    })
    $WPFCheckbox_MSAVDRemoteDesktop_Detail.Add_Unchecked({
        $WPFCheckbox_MSAVDRemoteDesktop.IsChecked = $WPFCheckbox_MSAVDRemoteDesktop_Detail.IsChecked
    })
    $WPFCheckbox_MSAzureDataStudio.Add_Checked({
        $WPFCheckbox_MSAzureDataStudio_Detail.IsChecked = $WPFCheckbox_MSAzureDataStudio.IsChecked
    })
    $WPFCheckbox_MSAzureDataStudio.Add_Unchecked({
        $WPFCheckbox_MSAzureDataStudio_Detail.IsChecked = $WPFCheckbox_MSAzureDataStudio.IsChecked
    })
    $WPFCheckbox_MSAzureDataStudio_Detail.Add_Checked({
        $WPFCheckbox_MSAzureDataStudio.IsChecked = $WPFCheckbox_MSAzureDataStudio_Detail.IsChecked
    })
    $WPFCheckbox_MSAzureDataStudio_Detail.Add_Unchecked({
        $WPFCheckbox_MSAzureDataStudio.IsChecked = $WPFCheckbox_MSAzureDataStudio_Detail.IsChecked
    })
    $WPFCheckbox_MSEdge.Add_Checked({
        $WPFCheckbox_MSEdge_Detail.IsChecked = $WPFCheckbox_MSEdge.IsChecked
    })
    $WPFCheckbox_MSEdge.Add_Unchecked({
        $WPFCheckbox_MSEdge_Detail.IsChecked = $WPFCheckbox_MSEdge.IsChecked
    })
    $WPFCheckbox_MSEdge_Detail.Add_Checked({
        $WPFCheckbox_MSEdge.IsChecked = $WPFCheckbox_MSEdge_Detail.IsChecked
    })
    $WPFCheckbox_MSEdge_Detail.Add_Unchecked({
        $WPFCheckbox_MSEdge.IsChecked = $WPFCheckbox_MSEdge_Detail.IsChecked
    })
    $WPFCheckbox_MSEdgeWebView2.Add_Checked({
        $WPFCheckbox_MSEdgeWebView2_Detail.IsChecked = $WPFCheckbox_MSEdgeWebView2.IsChecked
    })
    $WPFCheckbox_MSEdgeWebView2.Add_Unchecked({
        $WPFCheckbox_MSEdgeWebView2_Detail.IsChecked = $WPFCheckbox_MSEdgeWebView2.IsChecked
    })
    $WPFCheckbox_MSEdgeWebView2_Detail.Add_Checked({
        $WPFCheckbox_MSEdgeWebView2.IsChecked = $WPFCheckbox_MSEdgeWebView2_Detail.IsChecked
    })
    $WPFCheckbox_MSEdgeWebView2_Detail.Add_Unchecked({
        $WPFCheckbox_MSEdgeWebView2.IsChecked = $WPFCheckbox_MSEdgeWebView2_Detail.IsChecked
    })
    $WPFCheckbox_MSFSLogix.Add_Checked({
        $WPFCheckbox_MSFSLogix_Detail.IsChecked = $WPFCheckbox_MSFSLogix.IsChecked
    })
    $WPFCheckbox_MSFSLogix.Add_Unchecked({
        $WPFCheckbox_MSFSLogix_Detail.IsChecked = $WPFCheckbox_MSFSLogix.IsChecked
    })
    $WPFCheckbox_MSFSLogix_Detail.Add_Checked({
        $WPFCheckbox_MSFSLogix.IsChecked = $WPFCheckbox_MSFSLogix_Detail.IsChecked
    })
    $WPFCheckbox_MSFSLogix_Detail.Add_Unchecked({
        $WPFCheckbox_MSFSLogix.IsChecked = $WPFCheckbox_MSFSLogix_Detail.IsChecked
    })
    $WPFCheckbox_MSOffice.Add_Checked({
        $WPFCheckbox_MSOffice_Detail.IsChecked = $WPFCheckbox_MSOffice.IsChecked
    })
    $WPFCheckbox_MSOffice.Add_Unchecked({
        $WPFCheckbox_MSOffice_Detail.IsChecked = $WPFCheckbox_MSOffice.IsChecked
    })
    $WPFCheckbox_MSOffice_Detail.Add_Checked({
        $WPFCheckbox_MSOffice.IsChecked = $WPFCheckbox_MSOffice_Detail.IsChecked
    })
    $WPFCheckbox_MSOffice_Detail.Add_Unchecked({
        $WPFCheckbox_MSOffice.IsChecked = $WPFCheckbox_MSOffice_Detail.IsChecked
    })
    $WPFCheckbox_MSOneDrive.Add_Checked({
        $WPFCheckbox_MSOneDrive_Detail.IsChecked = $WPFCheckbox_MSOneDrive.IsChecked
    })
    $WPFCheckbox_MSOneDrive.Add_Unchecked({
        $WPFCheckbox_MSOneDrive_Detail.IsChecked = $WPFCheckbox_MSOneDrive.IsChecked
    })
    $WPFCheckbox_MSOneDrive_Detail.Add_Checked({
        $WPFCheckbox_MSOneDrive.IsChecked = $WPFCheckbox_MSOneDrive_Detail.IsChecked
    })
    $WPFCheckbox_MSOneDrive_Detail.Add_Unchecked({
        $WPFCheckbox_MSOneDrive.IsChecked = $WPFCheckbox_MSOneDrive_Detail.IsChecked
    })
    $WPFCheckbox_MSPowerBIDesktop.Add_Checked({
        $WPFCheckbox_MSPowerBIDesktop_Detail.IsChecked = $WPFCheckbox_MSPowerBIDesktop.IsChecked
    })
    $WPFCheckbox_MSPowerBIDesktop.Add_Unchecked({
        $WPFCheckbox_MSPowerBIDesktop_Detail.IsChecked = $WPFCheckbox_MSPowerBIDesktop.IsChecked
    })
    $WPFCheckbox_MSPowerBIDesktop_Detail.Add_Checked({
        $WPFCheckbox_MSPowerBIDesktop.IsChecked = $WPFCheckbox_MSPowerBIDesktop_Detail.IsChecked
    })
    $WPFCheckbox_MSPowerBIDesktop_Detail.Add_Unchecked({
        $WPFCheckbox_MSPowerBIDesktop.IsChecked = $WPFCheckbox_MSPowerBIDesktop_Detail.IsChecked
    })
    $WPFCheckbox_MSPowerShell.Add_Checked({
        $WPFCheckbox_MSPowerShell_Detail.IsChecked = $WPFCheckbox_MSPowerShell.IsChecked
    })
    $WPFCheckbox_MSPowerShell.Add_Unchecked({
        $WPFCheckbox_MSPowerShell_Detail.IsChecked = $WPFCheckbox_MSPowerShell.IsChecked
    })
    $WPFCheckbox_MSPowerShell_Detail.Add_Checked({
        $WPFCheckbox_MSPowerShell.IsChecked = $WPFCheckbox_MSPowerShell_Detail.IsChecked
    })
    $WPFCheckbox_MSPowerShell_Detail.Add_Unchecked({
        $WPFCheckbox_MSPowerShell.IsChecked = $WPFCheckbox_MSPowerShell_Detail.IsChecked
    })
    $WPFCheckbox_MSSQLServerManagementStudio.Add_Checked({
        $WPFCheckbox_MSSQLServerManagementStudio_Detail.IsChecked = $WPFCheckbox_MSSQLServerManagementStudio.IsChecked
    })
    $WPFCheckbox_MSSQLServerManagementStudio.Add_Unchecked({
        $WPFCheckbox_MSSQLServerManagementStudio_Detail.IsChecked = $WPFCheckbox_MSSQLServerManagementStudio.IsChecked
    })
    $WPFCheckbox_MSSQLServerManagementStudio_Detail.Add_Checked({
        $WPFCheckbox_MSSQLServerManagementStudio.IsChecked = $WPFCheckbox_MSSQLServerManagementStudio_Detail.IsChecked
    })
    $WPFCheckbox_MSSQLServerManagementStudio_Detail.Add_Unchecked({
        $WPFCheckbox_MSSQLServerManagementStudio.IsChecked = $WPFCheckbox_MSSQLServerManagementStudio_Detail.IsChecked
    })
    $WPFCheckbox_MSTeams.Add_Checked({
        $WPFCheckbox_MSTeams_Detail.IsChecked = $WPFCheckbox_MSTeams.IsChecked
    })
    $WPFCheckbox_MSTeams.Add_Unchecked({
        $WPFCheckbox_MSTeams_Detail.IsChecked = $WPFCheckbox_MSTeams.IsChecked
    })
    $WPFCheckbox_MSTeams_Detail.Add_Checked({
        $WPFCheckbox_MSTeams.IsChecked = $WPFCheckbox_MSTeams_Detail.IsChecked
    })
    $WPFCheckbox_MSTeams_Detail.Add_Unchecked({
        $WPFCheckbox_MSTeams.IsChecked = $WPFCheckbox_MSTeams_Detail.IsChecked
    })
    $WPFCheckbox_MSVisualStudioCode.Add_Checked({
        $WPFCheckbox_MSVisualStudioCode_Detail.IsChecked = $WPFCheckbox_MSVisualStudioCode.IsChecked
    })
    $WPFCheckbox_MSVisualStudioCode.Add_Unchecked({
        $WPFCheckbox_MSVisualStudioCode_Detail.IsChecked = $WPFCheckbox_MSVisualStudioCode.IsChecked
    })
    $WPFCheckbox_MSVisualStudioCode_Detail.Add_Checked({
        $WPFCheckbox_MSVisualStudioCode.IsChecked = $WPFCheckbox_MSVisualStudioCode_Detail.IsChecked
    })
    $WPFCheckbox_MSVisualStudioCode_Detail.Add_Unchecked({
        $WPFCheckbox_MSVisualStudioCode.IsChecked = $WPFCheckbox_MSVisualStudioCode_Detail.IsChecked
    })
    $WPFCheckbox_MindView7.Add_Checked({
        $WPFCheckbox_MindView7_Detail.IsChecked = $WPFCheckbox_MindView7.IsChecked
    })
    $WPFCheckbox_MindView7.Add_Unchecked({
        $WPFCheckbox_MindView7_Detail.IsChecked = $WPFCheckbox_MindView7.IsChecked
    })
    $WPFCheckbox_MindView7_Detail.Add_Checked({
        $WPFCheckbox_MindView7.IsChecked = $WPFCheckbox_MindView7_Detail.IsChecked
    })
    $WPFCheckbox_MindView7_Detail.Add_Unchecked({
        $WPFCheckbox_MindView7.IsChecked = $WPFCheckbox_MindView7_Detail.IsChecked
    })
    $WPFCheckbox_Firefox.Add_Checked({
        $WPFCheckbox_Firefox_Detail.IsChecked = $WPFCheckbox_Firefox.IsChecked
    })
    $WPFCheckbox_Firefox.Add_Unchecked({
        $WPFCheckbox_Firefox_Detail.IsChecked = $WPFCheckbox_Firefox.IsChecked
    })
    $WPFCheckbox_Firefox_Detail.Add_Checked({
        $WPFCheckbox_Firefox.IsChecked = $WPFCheckbox_Firefox_Detail.IsChecked
    })
    $WPFCheckbox_Firefox_Detail.Add_Unchecked({
        $WPFCheckbox_Firefox.IsChecked = $WPFCheckbox_Firefox_Detail.IsChecked
    })
    $WPFCheckbox_SumatraPDF.Add_Checked({
        $WPFCheckbox_SumatraPDF_Detail.IsChecked = $WPFCheckbox_SumatraPDF.IsChecked
    })
    $WPFCheckbox_SumatraPDF.Add_Unchecked({
        $WPFCheckbox_SumatraPDF_Detail.IsChecked = $WPFCheckbox_SumatraPDF.IsChecked
    })
    $WPFCheckbox_SumatraPDF_Detail.Add_Checked({
        $WPFCheckbox_SumatraPDF.IsChecked = $WPFCheckbox_SumatraPDF_Detail.IsChecked
    })
    $WPFCheckbox_SumatraPDF_Detail.Add_Unchecked({
        $WPFCheckbox_SumatraPDF.IsChecked = $WPFCheckbox_SumatraPDF_Detail.IsChecked
    })
    $WPFCheckbox_TechSmithSnagIt.Add_Checked({
        $WPFCheckbox_TechSmithSnagIt_Detail.IsChecked = $WPFCheckbox_TechSmithSnagIt.IsChecked
    })
    $WPFCheckbox_TechSmithSnagIt.Add_Unchecked({
        $WPFCheckbox_TechSmithSnagIt_Detail.IsChecked = $WPFCheckbox_TechSmithSnagIt.IsChecked
    })
    $WPFCheckbox_TechSmithSnagIt_Detail.Add_Checked({
        $WPFCheckbox_TechSmithSnagIt.IsChecked = $WPFCheckbox_TechSmithSnagIt_Detail.IsChecked
    })
    $WPFCheckbox_TechSmithSnagIt_Detail.Add_Unchecked({
        $WPFCheckbox_TechSmithSnagIt.IsChecked = $WPFCheckbox_TechSmithSnagIt_Detail.IsChecked
    })
    $WPFCheckbox_NotepadPlusPlus.Add_Checked({
        $WPFCheckbox_NotepadPlusPlus_Detail.IsChecked = $WPFCheckbox_NotepadPlusPlus.IsChecked
    })
    $WPFCheckbox_NotepadPlusPlus.Add_Unchecked({
        $WPFCheckbox_NotepadPlusPlus_Detail.IsChecked = $WPFCheckbox_NotepadPlusPlus.IsChecked
    })
    $WPFCheckbox_NotepadPlusPlus_Detail.Add_Checked({
        $WPFCheckbox_NotepadPlusPlus.IsChecked = $WPFCheckbox_NotepadPlusPlus_Detail.IsChecked
    })
    $WPFCheckbox_NotepadPlusPlus_Detail.Add_Unchecked({
        $WPFCheckbox_NotepadPlusPlus.IsChecked = $WPFCheckbox_NotepadPlusPlus_Detail.IsChecked
    })
    $WPFCheckbox_OpenJDK.Add_Checked({
        $WPFCheckbox_OpenJDK_Detail.IsChecked = $WPFCheckbox_OpenJDK.IsChecked
    })
    $WPFCheckbox_OpenJDK.Add_Unchecked({
        $WPFCheckbox_OpenJDK_Detail.IsChecked = $WPFCheckbox_OpenJDK.IsChecked
    })
    $WPFCheckbox_OpenJDK_Detail.Add_Checked({
        $WPFCheckbox_OpenJDK.IsChecked = $WPFCheckbox_OpenJDK_Detail.IsChecked
    })
    $WPFCheckbox_OpenJDK_Detail.Add_Unchecked({
        $WPFCheckbox_OpenJDK.IsChecked = $WPFCheckbox_OpenJDK_Detail.IsChecked
    })
    $WPFCheckbox_OracleJava8.Add_Checked({
        $WPFCheckbox_OracleJava8_Detail.IsChecked = $WPFCheckbox_OracleJava8.IsChecked
    })
    $WPFCheckbox_OracleJava8.Add_Unchecked({
        $WPFCheckbox_OracleJava8_Detail.IsChecked = $WPFCheckbox_OracleJava8.IsChecked
    })
    $WPFCheckbox_OracleJava8_Detail.Add_Checked({
        $WPFCheckbox_OracleJava8.IsChecked = $WPFCheckbox_OracleJava8_Detail.IsChecked
    })
    $WPFCheckbox_OracleJava8_Detail.Add_Unchecked({
        $WPFCheckbox_OracleJava8.IsChecked = $WPFCheckbox_OracleJava8_Detail.IsChecked
    })
    $WPFCheckbox_PeaZip.Add_Checked({
        $WPFCheckbox_PeaZip_Detail.IsChecked = $WPFCheckbox_PeaZip.IsChecked
    })
    $WPFCheckbox_PeaZip.Add_Unchecked({
        $WPFCheckbox_PeaZip_Detail.IsChecked = $WPFCheckbox_PeaZip.IsChecked
    })
    $WPFCheckbox_PeaZip_Detail.Add_Checked({
        $WPFCheckbox_PeaZip.IsChecked = $WPFCheckbox_PeaZip_Detail.IsChecked
    })
    $WPFCheckbox_PeaZip_Detail.Add_Unchecked({
        $WPFCheckbox_PeaZip.IsChecked = $WPFCheckbox_PeaZip_Detail.IsChecked
    })
    $WPFCheckbox_Putty.Add_Checked({
        $WPFCheckbox_Putty_Detail.IsChecked = $WPFCheckbox_Putty.IsChecked
    })
    $WPFCheckbox_Putty.Add_Unchecked({
        $WPFCheckbox_Putty_Detail.IsChecked = $WPFCheckbox_Putty.IsChecked
    })
    $WPFCheckbox_Putty_Detail.Add_Checked({
        $WPFCheckbox_Putty.IsChecked = $WPFCheckbox_Putty_Detail.IsChecked
    })
    $WPFCheckbox_Putty_Detail.Add_Unchecked({
        $WPFCheckbox_Putty.IsChecked = $WPFCheckbox_Putty_Detail.IsChecked
    })
    $WPFCheckbox_Slack.Add_Checked({
        $WPFCheckbox_Slack_Detail.IsChecked = $WPFCheckbox_Slack.IsChecked
    })
    $WPFCheckbox_Slack.Add_Unchecked({
        $WPFCheckbox_Slack_Detail.IsChecked = $WPFCheckbox_Slack.IsChecked
    })
    $WPFCheckbox_Slack_Detail.Add_Checked({
        $WPFCheckbox_Slack.IsChecked = $WPFCheckbox_Slack_Detail.IsChecked
    })
    $WPFCheckbox_Slack_Detail.Add_Unchecked({
        $WPFCheckbox_Slack.IsChecked = $WPFCheckbox_Slack_Detail.IsChecked
    })
    $WPFCheckbox_TotalCommander.Add_Checked({
        $WPFCheckbox_TotalCommander_Detail.IsChecked = $WPFCheckbox_TotalCommander.IsChecked
    })
    $WPFCheckbox_TotalCommander.Add_Unchecked({
        $WPFCheckbox_TotalCommander_Detail.IsChecked = $WPFCheckbox_TotalCommander.IsChecked
    })
    $WPFCheckbox_TotalCommander_Detail.Add_Checked({
        $WPFCheckbox_TotalCommander.IsChecked = $WPFCheckbox_TotalCommander_Detail.IsChecked
    })
    $WPFCheckbox_TotalCommander_Detail.Add_Unchecked({
        $WPFCheckbox_TotalCommander.IsChecked = $WPFCheckbox_TotalCommander_Detail.IsChecked
    })
    $WPFCheckbox_VLCPlayer.Add_Checked({
        $WPFCheckbox_VLCPlayer_Detail.IsChecked = $WPFCheckbox_VLCPlayer.IsChecked
    })
    $WPFCheckbox_VLCPlayer.Add_Unchecked({
        $WPFCheckbox_VLCPlayer_Detail.IsChecked = $WPFCheckbox_VLCPlayer.IsChecked
    })
    $WPFCheckbox_VLCPlayer_Detail.Add_Checked({
        $WPFCheckbox_VLCPlayer.IsChecked = $WPFCheckbox_VLCPlayer_Detail.IsChecked
    })
    $WPFCheckbox_VLCPlayer_Detail.Add_Unchecked({
        $WPFCheckbox_VLCPlayer.IsChecked = $WPFCheckbox_VLCPlayer_Detail.IsChecked
    })
    $WPFCheckbox_VMWareTools.Add_Checked({
        $WPFCheckbox_VMWareTools_Detail.IsChecked = $WPFCheckbox_VMWareTools.IsChecked
    })
    $WPFCheckbox_VMWareTools.Add_Unchecked({
        $WPFCheckbox_VMWareTools_Detail.IsChecked = $WPFCheckbox_VMWareTools.IsChecked
    })
    $WPFCheckbox_VMWareTools_Detail.Add_Checked({
        $WPFCheckbox_VMWareTools.IsChecked = $WPFCheckbox_VMWareTools_Detail.IsChecked
    })
    $WPFCheckbox_VMWareTools_Detail.Add_Unchecked({
        $WPFCheckbox_VMWareTools.IsChecked = $WPFCheckbox_VMWareTools_Detail.IsChecked
    })
    $WPFCheckbox_WinMerge.Add_Checked({
        $WPFCheckbox_WinMerge_Detail.IsChecked = $WPFCheckbox_WinMerge.IsChecked
    })
    $WPFCheckbox_WinMerge.Add_Unchecked({
        $WPFCheckbox_WinMerge_Detail.IsChecked = $WPFCheckbox_WinMerge.IsChecked
    })
    $WPFCheckbox_WinMerge_Detail.Add_Checked({
        $WPFCheckbox_WinMerge.IsChecked = $WPFCheckbox_WinMerge_Detail.IsChecked
    })
    $WPFCheckbox_WinMerge_Detail.Add_Unchecked({
        $WPFCheckbox_WinMerge.IsChecked = $WPFCheckbox_WinMerge_Detail.IsChecked
    })
    $WPFCheckbox_Wireshark.Add_Checked({
        $WPFCheckbox_Wireshark_Detail.IsChecked = $WPFCheckbox_Wireshark.IsChecked
    })
    $WPFCheckbox_Wireshark.Add_Unchecked({
        $WPFCheckbox_Wireshark_Detail.IsChecked = $WPFCheckbox_Wireshark.IsChecked
    })
    $WPFCheckbox_Wireshark_Detail.Add_Checked({
        $WPFCheckbox_Wireshark.IsChecked = $WPFCheckbox_Wireshark_Detail.IsChecked
    })
    $WPFCheckbox_Wireshark_Detail.Add_Unchecked({
        $WPFCheckbox_Wireshark.IsChecked = $WPFCheckbox_Wireshark_Detail.IsChecked
    })
    $WPFCheckbox_Zoom.Add_Checked({
        $WPFCheckbox_Zoom_Detail.IsChecked = $WPFCheckbox_Zoom.IsChecked
    })
    $WPFCheckbox_Zoom.Add_Unchecked({
        $WPFCheckbox_Zoom_Detail.IsChecked = $WPFCheckbox_Zoom.IsChecked
    })
    $WPFCheckbox_Zoom_Detail.Add_Checked({
        $WPFCheckbox_Zoom.IsChecked = $WPFCheckbox_Zoom_Detail.IsChecked
    })
    $WPFCheckbox_Zoom_Detail.Add_Unchecked({
        $WPFCheckbox_Zoom.IsChecked = $WPFCheckbox_Zoom_Detail.IsChecked
    })

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
        $WPFCheckbox_MSOffice.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSOneDrive.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSTeams.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_NotePadPlusPlus.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OpenJDK.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OpenShellMenu.IsChecked = $WPFCheckbox_SelectAll.IsChecked
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
        $WPFCheckbox_Putty.IsChecked = $WPFCheckbox_SelectAll.Ischecked
        $WPFCheckbox_PaintDotNet.IsChecked = $WPFCheckbox_SelectAll.Ischecked
        $WPFCheckbox_MSVisualStudio.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSVisualStudioCode.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerToys.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GIMP.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TeamViewer.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Wireshark.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSAzureDataStudio.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ImageGlass.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_uberAgent.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_1Password.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ControlUpAgent.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ControlUpConsole.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSSQLServerManagementStudio.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSAVDRemoteDesktop.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerBIDesktop.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_RDAnalyzer.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_SumatraPDF.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CiscoWebexTeams.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CitrixFiles.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_FoxitPDFEditor.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GitForWindows.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_LogMeInGoToMeeting.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSAzureCLI.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerBIReportBuilder.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSSysinternals.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Nmap.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_PeaZip.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TechSmithCamtasia.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TechSmithSnagIt.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_WinMerge.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSEdgeWebView2.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_AutodeskDWGTrueView.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MindView7.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_PDFsam.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_PDFForgeCreator.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TotalCommander.IsChecked = $WPFCheckbox_SelectAll.IsChecked
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
        $WPFCheckbox_MSOffice.IsChecked = $WPFCheckbox_SelectAll.IsChecked
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
        $WPFCheckbox_Putty.IsChecked = $WPFCheckbox_SelectAll.Ischecked
        $WPFCheckbox_PaintDotNet.IsChecked = $WPFCheckbox_SelectAll.Ischecked
        $WPFCheckbox_MSVisualStudio.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSVisualStudioCode.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerToys.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GIMP.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TeamViewer.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Wireshark.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSAzureDataStudio.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ImageGlass.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_uberAgent.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_1Password.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ControlUpAgent.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_ControlUpConsole.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSSQLServerManagementStudio.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSAVDRemoteDesktop.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerBIDesktop.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_RDAnalyzer.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_SumatraPDF.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CiscoWebexTeams.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_CitrixFiles.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_FoxitPDFEditor.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GitForWindows.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_LogMeInGoToMeeting.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSAzureCLI.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSPowerBIReportBuilder.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSSysinternals.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Nmap.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_PeaZip.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TechSmithCamtasia.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TechSmithSnagIt.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_WinMerge.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSEdgeWebView2.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_AutodeskDWGTrueView.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MindView7.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_PDFsam.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_OpenShellMenu.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_PDFForgeCreator.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_TotalCommander.IsChecked = $WPFCheckbox_SelectAll.IsChecked
    })

    # Button Browse
    $WPFButton_Browse.Add_Click({
        Show-OpenFileDialog
    })

    # Button Start (AddScript)
    $WPFButton_Start.Add_Click({
        If ($WPFCheckbox_Download.IsChecked -eq $True) {$Script:Download = 1}
        Else {$Script:Download = 0}
        If ($WPFCheckbox_Install.IsChecked -eq $True) {$Script:Install = 1}
        Else {$Script:Install = 0}
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
        If ($WPFCheckbox_MSFSLogix.IsChecked -eq $true) {$Script:MSFSLogix = 1}
        Else {$Script:MSFSLogix = 0}
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
        If ($WPFCheckbox_MSEdgeWebView2.ischecked -eq $true) {$Script:MSEdgeWebView2 = 1}
        Else {$Script:MSEdgeWebView2 = 0}
        If ($WPFCheckbox_MSOffice.ischecked -eq $true) {$Script:MSOffice = 1}
        Else {$Script:MSOffice = 0}
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
        If ($WPFCheckbox_GIMP.ischecked -eq $true) {$Script:GIMP = 1}
        Else {$Script:GIMP = 0}
        If ($WPFCheckbox_MSPowerToys.ischecked -eq $true) {$Script:MSPowerToys = 1}
        Else {$Script:MSPowerToys = 0}
        If ($WPFCheckbox_MSVisualStudio.ischecked -eq $true) {$Script:MSVisualStudio = 1}
        Else {$Script:MSVisualStudio = 0}
        If ($WPFCheckbox_MSVisualStudioCode.ischecked -eq $true) {$Script:MSVisualStudioCode = 1}
        Else {$Script:MSVisualStudioCode = 0}
        If ($WPFCheckbox_PaintDotNet.ischecked -eq $true) {$Script:PaintDotNet = 1}
        Else {$Script:PaintDotNet = 0}
        If ($WPFCheckbox_Putty.ischecked -eq $true) {$Script:Putty = 1}
        Else {$Script:Putty = 0}
        If ($WPFCheckbox_TeamViewer.ischecked -eq $true) {$Script:TeamViewer = 1}
        Else {$Script:TeamViewer = 0}
        If ($WPFCheckbox_Wireshark.ischecked -eq $true) {$Script:Wireshark = 1}
        Else {$Script:Wireshark = 0}
        If ($WPFCheckbox_MSAzureDataStudio.ischecked -eq $true) {$Script:MSAzureDataStudio = 1}
        Else {$Script:MSAzureDataStudio = 0}
        If ($WPFCheckbox_ImageGlass.ischecked -eq $true) {$Script:ImageGlass = 1}
        Else {$Script:ImageGlass = 0}
        If ($WPFCheckbox_uberAgent.ischecked -eq $true) {$Script:uberAgent = 1}
        Else {$Script:uberAgent = 0}
        If ($WPFCheckbox_1Password.ischecked -eq $true) {$Script:1Password = 1}
        Else {$Script:1Password = 0}
        If ($WPFCheckbox_ControlUpAgent.ischecked -eq $true) {$Script:ControlUpAgent = 1}
        Else {$Script:ControlUpAgent = 0}
        If ($WPFCheckbox_ControlUpConsole.ischecked -eq $true) {$Script:ControlUpConsole = 1}
        Else {$Script:ControlUpConsole = 0}
        If ($WPFCheckbox_MSSQLServerManagementStudio.ischecked -eq $true) {$Script:MSSQLServerManagementStudio = 1}
        Else {$Script:MSSQLServerManagementStudio = 0}
        If ($WPFCheckbox_MSAVDRemoteDesktop.ischecked -eq $true) {$Script:MSAVDRemoteDesktop = 1}
        Else {$Script:MSAVDRemoteDesktop = 0}
        If ($WPFCheckbox_MSPowerBIDesktop.ischecked -eq $true) {$Script:MSPowerBIDesktop = 1}
        Else {$Script:MSPowerBIDesktop = 0}
        If ($WPFCheckbox_RDAnalyzer.ischecked -eq $true) {$Script:RDAnalyzer = 1}
        Else {$Script:RDAnalyzer = 0}
        If ($WPFCheckbox_SumatraPDF.ischecked -eq $true) {$Script:SumatraPDF = 1}
        Else {$Script:SumatraPDF = 0}
        If ($WPFCheckbox_CiscoWebexTeams.ischecked -eq $true) {$Script:CiscoWebexTeams = 1}
        Else {$Script:CiscoWebexTeams = 0}
        If ($WPFCheckbox_CitrixFiles.ischecked -eq $true) {$Script:CitrixFiles = 1}
        Else {$Script:CitrixFiles = 0}
        If ($WPFCheckbox_FoxitPDFEditor.ischecked -eq $true) {$Script:FoxitPDFEditor = 1}
        Else {$Script:FoxitPDFEditor = 0}
        If ($WPFCheckbox_GitForWindows.ischecked -eq $true) {$Script:GitForWindows = 1}
        Else {$Script:GitForWindows = 0}
        If ($WPFCheckbox_LogMeInGoToMeeting.ischecked -eq $true) {$Script:LogMeInGoToMeeting = 1}
        Else {$Script:LogMeInGoToMeeting = 0}
        If ($WPFCheckbox_MSAzureCLI.ischecked -eq $true) {$Script:MSAzureCLI = 1}
        Else {$Script:MSAzureCLI = 0}
        If ($WPFCheckbox_MSPowerBIReportBuilder.ischecked -eq $true) {$Script:MSPowerBIReportBuilder = 1}
        Else {$Script:MSPowerBIReportBuilder = 0}
        If ($WPFCheckbox_MSSysinternals.ischecked -eq $true) {$Script:MSSysinternals = 1}
        Else {$Script:MSSysinternals = 0}
        If ($WPFCheckbox_Nmap.ischecked -eq $true) {$Script:Nmap = 1}
        Else {$Script:Nmap = 0}
        If ($WPFCheckbox_PeaZip.ischecked -eq $true) {$Script:PeaZip = 1}
        Else {$Script:PeaZip = 0}
        If ($WPFCheckbox_TechSmithCamtasia.ischecked -eq $true) {$Script:TechSmithCamtasia = 1}
        Else {$Script:TechSmithCamtasia = 0}
        If ($WPFCheckbox_TechSmithSnagIt.ischecked -eq $true) {$Script:TechSmithSnagIt = 1}
        Else {$Script:TechSmithSnagIt = 0}
        If ($WPFCheckbox_WinMerge.ischecked -eq $true) {$Script:WinMerge = 1}
        Else {$Script:WinMerge = 0}
        If ($WPFCheckbox_WhatIf.ischecked -eq $true) {$Script:WhatIf = 1}
        Else {$Script:WhatIf = 0}
        If ($WPFCheckbox_CleanUp.ischecked -eq $true) {$Script:CleanUp = 1}
        Else {$Script:CleanUp = 0}
        If ($WPFCheckbox_MS365Apps_Visio_Detail.ischecked -eq $true) {$Script:MS365Apps_Visio = 1}
        Else {$Script:MS365Apps_Visio = 0}
        If ($WPFCheckbox_MS365Apps_Project_Detail.ischecked -eq $true) {$Script:MS365Apps_Project = 1}
        Else {$Script:MS365Apps_Project = 0}
        If ($WPFCheckbox_AutodeskDWGTrueView.ischecked -eq $true) {$Script:AutodeskDWGTrueView = 1}
        Else {$Script:AutodeskDWGTrueView = 0}
        If ($WPFCheckbox_MindView7.ischecked -eq $true) {$Script:MindView7 = 1}
        Else {$Script:MindView7 = 0}
        If ($WPFCheckbox_PDFsam.ischecked -eq $true) {$Script:PDFsam = 1}
        Else {$Script:PDFsam = 0}
        If ($WPFCheckbox_OpenShellMenu.ischecked -eq $true) {$Script:OpenShellMenu = 1}
        Else {$Script:OpenShellMenu = 0}
        If ($WPFCheckbox_PDFForgeCreator.ischecked -eq $true) {$Script:PDFForgeCreator = 1}
        Else {$Script:PDFForgeCreator = 0}
        If ($WPFCheckbox_TotalCommander.ischecked -eq $true) {$Script:TotalCommander = 1}
        Else {$Script:TotalCommander = 0}
        If ($WPFCheckbox_Repository.ischecked -eq $true) {$Script:Repository = 1}
        Else {$Script:Repository = 0}
        If ($WPFCheckbox_CleanUpStartMenu.ischecked -eq $true) {$Script:CleanUpStartMenu = 1}
        Else {$Script:CleanUpStartMenu = 0}
        $Script:Language = $WPFBox_Language.SelectedIndex
        $Script:Architecture = $WPFBox_Architecture.SelectedIndex
        $Script:Installer = $WPFBox_Installer.SelectedIndex
        $Script:FirefoxChannel = $WPFBox_Firefox.SelectedIndex
        $Script:CitrixWorkspaceAppRelease = $WPFBox_CitrixWorkspaceApp.SelectedIndex
        $Script:MS365AppsChannel = $WPFBox_MS365Apps.SelectedIndex
        $Script:MSOneDriveRing = $WPFBox_MSOneDrive.SelectedIndex
        $Script:MSTeamsRing = $WPFBox_MSTeams.SelectedIndex
        $Script:TreeSizeType = $WPFBox_TreeSize.SelectedIndex
        $Script:MSDotNetFrameworkChannel = $WPFBox_MSDotNetFramework.SelectedIndex
        $Script:MSPowerShellRelease = $WPFBox_MSPowerShell.SelectedIndex
        $Script:RemoteDesktopManagerType = $WPFBox_RemoteDesktopManager.SelectedIndex
        $Script:ZoomCitrixClient = $WPFBox_Zoom.SelectedIndex
        $Script:deviceTRUSTPackage = $WPFBox_deviceTRUST.SelectedIndex
        $Script:MSEdgeChannel = $WPFBox_MSEdge.SelectedIndex
        $Script:MSVisualStudioCodeChannel = $WPFBox_MSVisualStudioCode.SelectedIndex
        $Script:MSVisualStudioEdition = $WPFBox_MSVisualStudio.SelectedIndex
        $Script:PuttyChannel = $WPFBox_Putty.SelectedIndex
        $Script:MSAzureDataStudioChannel = $WPFBox_MSAzureDataStudio.SelectedIndex
        $Script:MSFSLogixChannel = $WPFBox_MSFSLogix.SelectedIndex
        $Script:ControlUpAgentFramework = $WPFBox_ControlUpAgent.SelectedIndex
        $Script:MSAVDRemoteDesktopChannel = $WPFBox_MSAVDRemoteDesktop.SelectedIndex
        $Script:7Zip_Architecture = $WPFBox_7Zip_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Architecture = $WPFBox_AdobeReaderDC_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Language = $WPFBox_AdobeReaderDC_Language.SelectedIndex
        $Script:CiscoWebexTeams_Architecture = $WPFBox_CiscoWebexTeams_Architecture.SelectedIndex
        $Script:CitrixHypervisorTools_Architecture = $WPFBox_CitrixHypervisorTools_Architecture.SelectedIndex
        $Script:ControlUpAgent_Architecture = $WPFBox_ControlUpAgent_Architecture.SelectedIndex
        $Script:deviceTRUST_Architecture = $WPFBox_deviceTRUST_Architecture.SelectedIndex
        $Script:FoxitPDFEditor_Language = $WPFBox_FoxitPDFEditor_Language.SelectedIndex
        $Script:FoxitReader_Language = $WPFBox_FoxitReader_Language.SelectedIndex
        $Script:GitForWindows_Architecture = $WPFBox_GitForWindows_Architecture.SelectedIndex
        $Script:GoogleChrome_Architecture = $WPFBox_GoogleChrome_Architecture.SelectedIndex
        $Script:ImageGlass_Architecture = $WPFBox_ImageGlass_Architecture.SelectedIndex
        $Script:IrfanView_Architecture = $WPFBox_IrfanView_Architecture.SelectedIndex
        $Script:Keepass_Language = $WPFBox_KeePass_Language.SelectedIndex
        $Script:MSDotNetFramework_Architecture = $WPFBox_MSDotNetFramework_Architecture.SelectedIndex
        $Script:MS365Apps_Architecture = $WPFBox_MS365Apps_Architecture.SelectedIndex
        $Script:MS365Apps_Language = $WPFBox_MS365Apps_Language.SelectedIndex
        $Script:MS365Apps_Visio_Language = $WPFBox_MS365Apps_Visio_Language.SelectedIndex
        $Script:MS365Apps_Project_Language = $WPFBox_MS365Apps_Project_Language.SelectedIndex
        $Script:MSAVDRemoteDesktop_Architecture = $WPFBox_MSAVDRemoteDesktop_Architecture.SelectedIndex
        $Script:MSEdge_Architecture = $WPFBox_MSEdge_Architecture.SelectedIndex
        $Script:MSFSLogix_Architecture = $WPFBox_MSFSLogix_Architecture.SelectedIndex
        $Script:MSOffice_Architecture = $WPFBox_MSOffice_Architecture.SelectedIndex
        $Script:MSOneDrive_Architecture = $WPFBox_MSOneDrive_Architecture.SelectedIndex
        $Script:MSPowerBIDesktop_Architecture = $WPFBox_MSPowerBIDesktop_Architecture.SelectedIndex
        $Script:MSPowerShell_Architecture = $WPFBox_MSPowerShell_Architecture.SelectedIndex
        $Script:MSSQLServerManagementStudio_Language = $WPFBox_MSSQLServerManagementStudio_Language.SelectedIndex
        $Script:MSTeams_Architecture = $WPFBox_MSTeams_Architecture.SelectedIndex
        $Script:MSVisualStudioCode_Architecture = $WPFBox_MSVisualStudioCode_Architecture.SelectedIndex
        $Script:Firefox_Architecture = $WPFBox_Firefox_Architecture.SelectedIndex
        $Script:Firefox_Language = $WPFBox_Firefox_Language.SelectedIndex
        $Script:NotePadPlusPlus_Architecture = $WPFBox_NotepadPlusPlus_Architecture.SelectedIndex
        $Script:OpenJDK_Architecture = $WPFBox_OpenJDK_Architecture.SelectedIndex
        $Script:OracleJava8_Architecture = $WPFBox_OracleJava8_Architecture.SelectedIndex
        $Script:PeaZip_Architecture = $WPFBox_PeaZip_Architecture.SelectedIndex
        $Script:Putty_Architecture = $WPFBox_Putty_Architecture.SelectedIndex
        $Script:Slack_Architecture = $WPFBox_Slack_Architecture.SelectedIndex
        $Script:SumatraPDF_Architecture = $WPFBox_SumatraPDF_Architecture.SelectedIndex
        $Script:TechSmithSnagIt_Architecture = $WPFBox_TechSmithSnagIT_Architecture.SelectedIndex
        $Script:VLCPlayer_Architecture = $WPFBox_VLCPlayer_Architecture.SelectedIndex
        $Script:VMWareTools_Architecture = $WPFBox_VMWareTools_Architecture.SelectedIndex
        $Script:WinMerge_Architecture = $WPFBox_WinMerge_Architecture.SelectedIndex
        $Script:Wireshark_Architecture = $WPFBox_Wireshark_Architecture.SelectedIndex
        $Script:IrfanView_Language = $WPFBox_IrfanView_Language.SelectedIndex
        $Script:MSOffice_Language = $WPFBox_MSOffice_Language.SelectedIndex
        $Script:MS365Apps_Path = $WPFTextBox_Filename.Text
        $Script:MSEdgeWebView2_Architecture = $WPFBox_MSEdgeWebView2_Architecture.SelectedIndex
        $Script:MindView7_Language = $WPFBox_MindView7_Language.SelectedIndex
        $Script:MSOfficeVersion = $WPFBox_MSOffice.SelectedIndex
        $Script:LogMeInGoToMeeting_Installer = $WPFBox_LogMeInGoToMeeting_Installer.SelectedIndex
        $Script:MSAzureDataStudio_Installer = $WPFBox_MSAzureDataStudio_Installer.SelectedIndex
        $Script:MSVisualStudioCode_Installer = $WPFBox_MSVisualStudioCode_Installer.SelectedIndex
        $Script:MS365Apps_Installer = $WPFBox_MS365Apps_Installer.SelectedIndex
        $Script:MSOffice_Installer = $WPFBox_MSOffice_Installer.SelectedIndex
        $Script:MSTeams_Installer = $WPFBox_MSTeams_Installer.SelectedIndex
        $Script:Zoom_Installer = $WPFBox_Zoom_Installer.SelectedIndex
        $Script:MSOneDrive_Installer = $WPFBox_MSOneDrive_Installer.SelectedIndex
        $Script:Slack_Installer = $WPFBox_Slack_Installer.SelectedIndex
        $Script:pdfforgePDFCreatorChannel = $WPFBox_pdfforgePDFCreator.SelectedIndex
        $Script:TotalCommander_Architecture = $WPFBox_TotalCommander_Architecture.SelectedIndex
        
        # Write LastSetting.txt or -GUIFile Parameter file to get the settings of the last session. (AddScript)
        $Language,$Architecture,$CitrixWorkspaceAppRelease,$MS365AppsChannel,$MSOneDriveRing,$MSTeamsRing,$FirefoxChannel,$TreeSizeType,$7ZIP,$AdobeProDC,$AdobeReaderDC,$BISF,$Citrix_Hypervisor_Tools,$Citrix_WorkspaceApp,$Filezilla,$Firefox,$Foxit_Reader,$MSFSLogix,$GoogleChrome,$Greenshot,$KeePass,$mRemoteNG,$MS365Apps,$MSEdge,$MSOffice,$MSOneDrive,$MSTeams,$NotePadPlusPlus,$OpenJDK,$OracleJava8,$TreeSize,$VLCPlayer,$VMWareTools,$WinSCP,$Download,$Install,$IrfanView,$MSTeamsNoAutoStart,$deviceTRUST,$MSDotNetFramework,$MSDotNetFrameworkChannel,$MSPowerShell,$MSPowerShellRelease,$RemoteDesktopManager,$RemoteDesktopManagerType,$Slack,$Wireshark,$ShareX,$Zoom,$ZoomCitrixClient,$deviceTRUSTPackage,$MSEdgeChannel,$GIMP,$MSPowerToys,$MSVisualStudio,$MSVisualStudioCode,$MSVisualStudioCodeChannel,$PaintDotNet,$Putty,$TeamViewer,$Installer,$MSVisualStudioEdition,$PuttyChannel,$MSAzureDataStudio,$MSAzureDataStudioChannel,$ImageGlass,$MSFSLogixChannel,$uberAgent,$1Password,$SumatraPDF,$ControlUpAgent,$ControlUpAgentFramework,$ControlUpConsole,$MSSQLServerManagementStudio,$MSAVDRemoteDesktop,$MSAVDRemoteDesktopChannel,$MSPowerBIDesktop,$RDAnalyzer,$CiscoWebexTeams,$CitrixFiles,$FoxitPDFEditor,$GitForWindows,$LogMeInGoToMeeting,$MSAzureCLI,$MSPowerBIReportBuilder,$MSSysinternals,$NMap,$PeaZip,$TechSmithCamtasia,$TechSmithSnagit,$WinMerge,$WhatIf,$CleanUp,$7Zip_Architecture,$AdobeReaderDC_Architecture,$AdobeReaderDC_Language,$CiscoWebexTeams_Architecture,$CitrixHypervisorTools_Architecture,$ControlUpAgent_Architecture,$deviceTRUST_Architecture,$FoxitPDFEditor_Language,$FoxitReader_Language,$GitForWindows_Architecture,$GoogleChrome_Architecture,$ImageGlass_Architecture,$IrfanView_Architecture,$Keepass_Language,$MSDotNetFramework_Architecture,$MS365Apps_Architecture,$MS365Apps_Language,$MS365Apps_Visio,$MS365Apps_Visio_Language,$MS365Apps_Project,$MS365Apps_Project_Language,$MSAVDRemoteDesktop_Architecture,$MSEdge_Architecture,$MSFSLogix_Architecture,$MSOffice_Architecture,$MSOneDrive_Architecture,$MSPowerBIDesktop_Architecture,$MSPowerShell_Architecture,$MSSQLServerManagementStudio_Language,$MSTeams_Architecture,$MSVisualStudioCode_Architecture,$Firefox_Architecture,$Firefox_Language,$NotePadPlusPlus_Architecture,$OpenJDK_Architecture,$OracleJava8_Architecture,$PeaZip_Architecture,$Putty_Architecture,$Slack_Architecture,$SumatraPDF_Architecture,$TechSmithSnagIt_Architecture,$VLCPlayer_Architecture,$VMWareTools_Architecture,$WinMerge_Architecture,$Wireshark_Architecture,$IrfanView_Language,$MSOffice_Language,$MSEdgeWebView2,$MSEdgeWebView2_Architecture,$AutodeskDWGTrueView,$MindView7,$MindView7_Language,$PDFsam,$MSOfficeVersion,$OpenShellMenu,$PDFForgeCreator,$TotalCommander,$LogMeInGoToMeeting_Installer,$MSAzureDataStudio_Installer,$MSVisualStudioCode_Installer,$MS365Apps_Installer,$MSOffice_Installer,$MSTeams_Installer,$Zoom_Installer,$MSOneDrive_Installer,$Slack_Installer,$pdfforgePDFCreatorChannel,$TotalCommander_Architecture,$Repository,$CleanUpStartMenu | out-file -filepath "$PSScriptRoot\$GUIfile"
        If ($MS365Apps_Path -ne "") {
            If ($WhatIf -eq '0') {
                Switch ($MS365AppsChannel) {
                    0 { $MS365AppsChannelClear = 'Beta Channel'}
                    1 { $MS365AppsChannelClear = 'Current Channel (Preview)'}
                    2 { $MS365AppsChannelClear = 'Current Channel'}
                    3 { $MS365AppsChannelClear = 'Semi-Annual Channel (Preview)'}
                    4 { $MS365AppsChannelClear = 'Monthly Enterprise Channel'}
                    5 { $MS365AppsChannelClear = 'Semi-Annual Channel'}
                }
                If (!(Test-Path -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear")) {New-Item -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear" -ItemType Directory | Out-Null}
                copy-item -Path "$MS365Apps_Path" -Destination "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear\install.xml" -Force
            }
            Write-Host -ForegroundColor Green "Copy custom XML file as install.xml successfully."
            Write-Output ""
        }
        Write-Host "GUI Mode"
        $Form.Close()
    })

    # Button Start Detail (AddScript)
    $WPFButton_Start_Detail.Add_Click({
        If ($WPFCheckbox_Download.IsChecked -eq $True) {$Script:Download = 1}
        Else {$Script:Download = 0}
        If ($WPFCheckbox_Install.IsChecked -eq $True) {$Script:Install = 1}
        Else {$Script:Install = 0}
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
        If ($WPFCheckbox_MSFSLogix.IsChecked -eq $true) {$Script:MSFSLogix = 1}
        Else {$Script:MSFSLogix = 0}
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
        If ($WPFCheckbox_MSEdgeWebView2.ischecked -eq $true) {$Script:MSEdgeWebView2 = 1}
        Else {$Script:MSEdgeWebView2 = 0}
        If ($WPFCheckbox_MSOffice.ischecked -eq $true) {$Script:MSOffice = 1}
        Else {$Script:MSOffice = 0}
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
        If ($WPFCheckbox_GIMP.ischecked -eq $true) {$Script:GIMP = 1}
        Else {$Script:GIMP = 0}
        If ($WPFCheckbox_MSPowerToys.ischecked -eq $true) {$Script:MSPowerToys = 1}
        Else {$Script:MSPowerToys = 0}
        If ($WPFCheckbox_MSVisualStudio.ischecked -eq $true) {$Script:MSVisualStudio = 1}
        Else {$Script:MSVisualStudio = 0}
        If ($WPFCheckbox_MSVisualStudioCode.ischecked -eq $true) {$Script:MSVisualStudioCode = 1}
        Else {$Script:MSVisualStudioCode = 0}
        If ($WPFCheckbox_PaintDotNet.ischecked -eq $true) {$Script:PaintDotNet = 1}
        Else {$Script:PaintDotNet = 0}
        If ($WPFCheckbox_Putty.ischecked -eq $true) {$Script:Putty = 1}
        Else {$Script:Putty = 0}
        If ($WPFCheckbox_TeamViewer.ischecked -eq $true) {$Script:TeamViewer = 1}
        Else {$Script:TeamViewer = 0}
        If ($WPFCheckbox_Wireshark.ischecked -eq $true) {$Script:Wireshark = 1}
        Else {$Script:Wireshark = 0}
        If ($WPFCheckbox_MSAzureDataStudio.ischecked -eq $true) {$Script:MSAzureDataStudio = 1}
        Else {$Script:MSAzureDataStudio = 0}
        If ($WPFCheckbox_ImageGlass.ischecked -eq $true) {$Script:ImageGlass = 1}
        Else {$Script:ImageGlass = 0}
        If ($WPFCheckbox_uberAgent.ischecked -eq $true) {$Script:uberAgent = 1}
        Else {$Script:uberAgent = 0}
        If ($WPFCheckbox_1Password.ischecked -eq $true) {$Script:1Password = 1}
        Else {$Script:1Password = 0}
        If ($WPFCheckbox_ControlUpAgent.ischecked -eq $true) {$Script:ControlUpAgent = 1}
        Else {$Script:ControlUpAgent = 0}
        If ($WPFCheckbox_ControlUpConsole.ischecked -eq $true) {$Script:ControlUpConsole = 1}
        Else {$Script:ControlUpConsole = 0}
        If ($WPFCheckbox_MSSQLServerManagementStudio.ischecked -eq $true) {$Script:MSSQLServerManagementStudio = 1}
        Else {$Script:MSSQLServerManagementStudio = 0}
        If ($WPFCheckbox_MSAVDRemoteDesktop.ischecked -eq $true) {$Script:MSAVDRemoteDesktop = 1}
        Else {$Script:MSAVDRemoteDesktop = 0}
        If ($WPFCheckbox_MSPowerBIDesktop.ischecked -eq $true) {$Script:MSPowerBIDesktop = 1}
        Else {$Script:MSPowerBIDesktop = 0}
        If ($WPFCheckbox_RDAnalyzer.ischecked -eq $true) {$Script:RDAnalyzer = 1}
        Else {$Script:RDAnalyzer = 0}
        If ($WPFCheckbox_SumatraPDF.ischecked -eq $true) {$Script:SumatraPDF = 1}
        Else {$Script:SumatraPDF = 0}
        If ($WPFCheckbox_CiscoWebexTeams.ischecked -eq $true) {$Script:CiscoWebexTeams = 1}
        Else {$Script:CiscoWebexTeams = 0}
        If ($WPFCheckbox_CitrixFiles.ischecked -eq $true) {$Script:CitrixFiles = 1}
        Else {$Script:CitrixFiles = 0}
        If ($WPFCheckbox_FoxitPDFEditor.ischecked -eq $true) {$Script:FoxitPDFEditor = 1}
        Else {$Script:FoxitPDFEditor = 0}
        If ($WPFCheckbox_GitForWindows.ischecked -eq $true) {$Script:GitForWindows = 1}
        Else {$Script:GitForWindows = 0}
        If ($WPFCheckbox_LogMeInGoToMeeting.ischecked -eq $true) {$Script:LogMeInGoToMeeting = 1}
        Else {$Script:LogMeInGoToMeeting = 0}
        If ($WPFCheckbox_MSAzureCLI.ischecked -eq $true) {$Script:MSAzureCLI = 1}
        Else {$Script:MSAzureCLI = 0}
        If ($WPFCheckbox_MSPowerBIReportBuilder.ischecked -eq $true) {$Script:MSPowerBIReportBuilder = 1}
        Else {$Script:MSPowerBIReportBuilder = 0}
        If ($WPFCheckbox_MSSysinternals.ischecked -eq $true) {$Script:MSSysinternals = 1}
        Else {$Script:MSSysinternals = 0}
        If ($WPFCheckbox_Nmap.ischecked -eq $true) {$Script:Nmap = 1}
        Else {$Script:Nmap = 0}
        If ($WPFCheckbox_PeaZip.ischecked -eq $true) {$Script:PeaZip = 1}
        Else {$Script:PeaZip = 0}
        If ($WPFCheckbox_TechSmithCamtasia.ischecked -eq $true) {$Script:TechSmithCamtasia = 1}
        Else {$Script:TechSmithCamtasia = 0}
        If ($WPFCheckbox_TechSmithSnagIt.ischecked -eq $true) {$Script:TechSmithSnagIt = 1}
        Else {$Script:TechSmithSnagIt = 0}
        If ($WPFCheckbox_WinMerge.ischecked -eq $true) {$Script:WinMerge = 1}
        Else {$Script:WinMerge = 0}
        If ($WPFCheckbox_WhatIf.ischecked -eq $true) {$Script:WhatIf = 1}
        Else {$Script:WhatIf = 0}
        If ($WPFCheckbox_CleanUp.ischecked -eq $true) {$Script:CleanUp = 1}
        Else {$Script:CleanUp = 0}
        If ($WPFCheckbox_Repository.ischecked -eq $true) {$Script:Repository = 1}
        Else {$Script:Repository = 0}
        If ($WPFCheckbox_CleanUpStartMenu.ischecked -eq $true) {$Script:CleanUpStartMenu = 1}
        Else {$Script:CleanUpStartMenu = 0}
        If ($WPFCheckbox_MS365Apps_Visio_Detail.ischecked -eq $true) {$Script:MS365Apps_Visio = 1}
        Else {$Script:MS365Apps_Visio = 0}
        If ($WPFCheckbox_MS365Apps_Project_Detail.ischecked -eq $true) {$Script:MS365Apps_Project = 1}
        Else {$Script:MS365Apps_Project = 0}
        If ($WPFCheckbox_AutodeskDWGTrueView.ischecked -eq $true) {$Script:AutodeskDWGTrueView = 1}
        Else {$Script:AutodeskDWGTrueView = 0}
        If ($WPFCheckbox_MindView7.ischecked -eq $true) {$Script:MindView7 = 1}
        Else {$Script:MindView7 = 0}
        If ($WPFCheckbox_PDFsam.ischecked -eq $true) {$Script:PDFsam = 1}
        Else {$Script:PDFsam = 0}
        If ($WPFCheckbox_OpenShellMenu.ischecked -eq $true) {$Script:OpenShellMenu = 1}
        Else {$Script:OpenShellMenu = 0}
        If ($WPFCheckbox_PDFForgeCreator.ischecked -eq $true) {$Script:PDFForgeCreator = 1}
        Else {$Script:PDFForgeCreator = 0}
        If ($WPFCheckbox_TotalCommander.ischecked -eq $true) {$Script:TotalCommander = 1}
        Else {$Script:TotalCommander = 0}
        $Script:Language = $WPFBox_Language.SelectedIndex
        $Script:Architecture = $WPFBox_Architecture.SelectedIndex
        $Script:Installer = $WPFBox_Installer.SelectedIndex
        $Script:FirefoxChannel = $WPFBox_Firefox.SelectedIndex
        $Script:CitrixWorkspaceAppRelease = $WPFBox_CitrixWorkspaceApp.SelectedIndex
        $Script:MS365AppsChannel = $WPFBox_MS365Apps.SelectedIndex
        $Script:MSOneDriveRing = $WPFBox_MSOneDrive.SelectedIndex
        $Script:MSTeamsRing = $WPFBox_MSTeams.SelectedIndex
        $Script:TreeSizeType = $WPFBox_TreeSize.SelectedIndex
        $Script:MSDotNetFrameworkChannel = $WPFBox_MSDotNetFramework.SelectedIndex
        $Script:MSPowerShellRelease = $WPFBox_MSPowerShell.SelectedIndex
        $Script:RemoteDesktopManagerType = $WPFBox_RemoteDesktopManager.SelectedIndex
        $Script:ZoomCitrixClient = $WPFBox_Zoom.SelectedIndex
        $Script:deviceTRUSTPackage = $WPFBox_deviceTRUST.SelectedIndex
        $Script:MSEdgeChannel = $WPFBox_MSEdge.SelectedIndex
        $Script:MSVisualStudioCodeChannel = $WPFBox_MSVisualStudioCode.SelectedIndex
        $Script:MSVisualStudioEdition = $WPFBox_MSVisualStudio.SelectedIndex
        $Script:PuttyChannel = $WPFBox_Putty.SelectedIndex
        $Script:MSAzureDataStudioChannel = $WPFBox_MSAzureDataStudio.SelectedIndex
        $Script:MSFSLogixChannel = $WPFBox_MSFSLogix.SelectedIndex
        $Script:ControlUpAgentFramework = $WPFBox_ControlUpAgent.SelectedIndex
        $Script:MSAVDRemoteDesktopChannel = $WPFBox_MSAVDRemoteDesktop.SelectedIndex
        $Script:7Zip_Architecture = $WPFBox_7Zip_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Architecture = $WPFBox_AdobeReaderDC_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Language = $WPFBox_AdobeReaderDC_Language.SelectedIndex
        $Script:CiscoWebexTeams_Architecture = $WPFBox_CiscoWebexTeams_Architecture.SelectedIndex
        $Script:CitrixHypervisorTools_Architecture = $WPFBox_CitrixHypervisorTools_Architecture.SelectedIndex
        $Script:ControlUpAgent_Architecture = $WPFBox_ControlUpAgent_Architecture.SelectedIndex
        $Script:deviceTRUST_Architecture = $WPFBox_deviceTRUST_Architecture.SelectedIndex
        $Script:FoxitPDFEditor_Language = $WPFBox_FoxitPDFEditor_Language.SelectedIndex
        $Script:FoxitReader_Language = $WPFBox_FoxitReader_Language.SelectedIndex
        $Script:GitForWindows_Architecture = $WPFBox_GitForWindows_Architecture.SelectedIndex
        $Script:GoogleChrome_Architecture = $WPFBox_GoogleChrome_Architecture.SelectedIndex
        $Script:ImageGlass_Architecture = $WPFBox_ImageGlass_Architecture.SelectedIndex
        $Script:IrfanView_Architecture = $WPFBox_IrfanView_Architecture.SelectedIndex
        $Script:Keepass_Language = $WPFBox_KeePass_Language.SelectedIndex
        $Script:MSDotNetFramework_Architecture = $WPFBox_MSDotNetFramework_Architecture.SelectedIndex
        $Script:MS365Apps_Architecture = $WPFBox_MS365Apps_Architecture.SelectedIndex
        $Script:MS365Apps_Language = $WPFBox_MS365Apps_Language.SelectedIndex
        $Script:MS365Apps_Visio_Language = $WPFBox_MS365Apps_Visio_Language.SelectedIndex
        $Script:MS365Apps_Project_Language = $WPFBox_MS365Apps_Project_Language.SelectedIndex
        $Script:MSAVDRemoteDesktop_Architecture = $WPFBox_MSAVDRemoteDesktop_Architecture.SelectedIndex
        $Script:MSEdge_Architecture = $WPFBox_MSEdge_Architecture.SelectedIndex
        $Script:MSFSLogix_Architecture = $WPFBox_MSFSLogix_Architecture.SelectedIndex
        $Script:MSOffice_Architecture = $WPFBox_MSOffice_Architecture.SelectedIndex
        $Script:MSOneDrive_Architecture = $WPFBox_MSOneDrive_Architecture.SelectedIndex
        $Script:MSPowerBIDesktop_Architecture = $WPFBox_MSPowerBIDesktop_Architecture.SelectedIndex
        $Script:MSPowerShell_Architecture = $WPFBox_MSPowerShell_Architecture.SelectedIndex
        $Script:MSSQLServerManagementStudio_Language = $WPFBox_MSSQLServerManagementStudio_Language.SelectedIndex
        $Script:MSTeams_Architecture = $WPFBox_MSTeams_Architecture.SelectedIndex
        $Script:MSVisualStudioCode_Architecture = $WPFBox_MSVisualStudioCode_Architecture.SelectedIndex
        $Script:Firefox_Architecture = $WPFBox_Firefox_Architecture.SelectedIndex
        $Script:Firefox_Language = $WPFBox_Firefox_Language.SelectedIndex
        $Script:NotePadPlusPlus_Architecture = $WPFBox_NotepadPlusPlus_Architecture.SelectedIndex
        $Script:OpenJDK_Architecture = $WPFBox_OpenJDK_Architecture.SelectedIndex
        $Script:OracleJava8_Architecture = $WPFBox_OracleJava8_Architecture.SelectedIndex
        $Script:PeaZip_Architecture = $WPFBox_PeaZip_Architecture.SelectedIndex
        $Script:Putty_Architecture = $WPFBox_Putty_Architecture.SelectedIndex
        $Script:Slack_Architecture = $WPFBox_Slack_Architecture.SelectedIndex
        $Script:SumatraPDF_Architecture = $WPFBox_SumatraPDF_Architecture.SelectedIndex
        $Script:TechSmithSnagIt_Architecture = $WPFBox_TechSmithSnagIT_Architecture.SelectedIndex
        $Script:VLCPlayer_Architecture = $WPFBox_VLCPlayer_Architecture.SelectedIndex
        $Script:VMWareTools_Architecture = $WPFBox_VMWareTools_Architecture.SelectedIndex
        $Script:WinMerge_Architecture = $WPFBox_WinMerge_Architecture.SelectedIndex
        $Script:Wireshark_Architecture = $WPFBox_Wireshark_Architecture.SelectedIndex
        $Script:IrfanView_Language = $WPFBox_IrfanView_Language.SelectedIndex
        $Script:MSOffice_Language = $WPFBox_MSOffice_Language.SelectedIndex
        $Script:MS365Apps_Path = $WPFTextBox_Filename.Text
        $Script:MSEdgeWebView2_Architecture = $WPFBox_MSEdgeWebView2_Architecture.SelectedIndex
        $Script:MindView7_Language = $WPFBox_MindView7_Language.SelectedIndex
        $Script:MSOfficeVersion = $WPFBox_MSOffice.SelectedIndex
        $Script:LogMeInGoToMeeting_Installer = $WPFBox_LogMeInGoToMeeting_Installer.SelectedIndex
        $Script:MSAzureDataStudio_Installer = $WPFBox_MSAzureDataStudio_Installer.SelectedIndex
        $Script:MSVisualStudioCode_Installer = $WPFBox_MSVisualStudioCode_Installer.SelectedIndex
        $Script:MS365Apps_Installer = $WPFBox_MS365Apps_Installer.SelectedIndex
        $Script:MSOffice_Installer = $WPFBox_MSOffice_Installer.SelectedIndex
        $Script:MSTeams_Installer = $WPFBox_MSTeams_Installer.SelectedIndex
        $Script:Zoom_Installer = $WPFBox_Zoom_Installer.SelectedIndex
        $Script:MSOneDrive_Installer = $WPFBox_MSOneDrive_Installer.SelectedIndex
        $Script:Slack_Installer = $WPFBox_Slack_Installer.SelectedIndex
        $Script:pdfforgePDFCreatorChannel = $WPFBox_pdfforgePDFCreator.SelectedIndex
        $Script:TotalCommander_Architecture = $WPFBox_TotalCommander_Architecture.SelectedIndex
        
        # Write LastSetting.txt or -GUIFile Parameter file to get the settings of the last session. (AddScript)
        $Language,$Architecture,$CitrixWorkspaceAppRelease,$MS365AppsChannel,$MSOneDriveRing,$MSTeamsRing,$FirefoxChannel,$TreeSizeType,$7ZIP,$AdobeProDC,$AdobeReaderDC,$BISF,$Citrix_Hypervisor_Tools,$Citrix_WorkspaceApp,$Filezilla,$Firefox,$Foxit_Reader,$MSFSLogix,$GoogleChrome,$Greenshot,$KeePass,$mRemoteNG,$MS365Apps,$MSEdge,$MSOffice,$MSOneDrive,$MSTeams,$NotePadPlusPlus,$OpenJDK,$OracleJava8,$TreeSize,$VLCPlayer,$VMWareTools,$WinSCP,$Download,$Install,$IrfanView,$MSTeamsNoAutoStart,$deviceTRUST,$MSDotNetFramework,$MSDotNetFrameworkChannel,$MSPowerShell,$MSPowerShellRelease,$RemoteDesktopManager,$RemoteDesktopManagerType,$Slack,$Wireshark,$ShareX,$Zoom,$ZoomCitrixClient,$deviceTRUSTPackage,$MSEdgeChannel,$GIMP,$MSPowerToys,$MSVisualStudio,$MSVisualStudioCode,$MSVisualStudioCodeChannel,$PaintDotNet,$Putty,$TeamViewer,$Installer,$MSVisualStudioEdition,$PuttyChannel,$MSAzureDataStudio,$MSAzureDataStudioChannel,$ImageGlass,$MSFSLogixChannel,$uberAgent,$1Password,$SumatraPDF,$ControlUpAgent,$ControlUpAgentFramework,$ControlUpConsole,$MSSQLServerManagementStudio,$MSAVDRemoteDesktop,$MSAVDRemoteDesktopChannel,$MSPowerBIDesktop,$RDAnalyzer,$CiscoWebexTeams,$CitrixFiles,$FoxitPDFEditor,$GitForWindows,$LogMeInGoToMeeting,$MSAzureCLI,$MSPowerBIReportBuilder,$MSSysinternals,$NMap,$PeaZip,$TechSmithCamtasia,$TechSmithSnagit,$WinMerge,$WhatIf,$CleanUp,$7Zip_Architecture,$AdobeReaderDC_Architecture,$AdobeReaderDC_Language,$CiscoWebexTeams_Architecture,$CitrixHypervisorTools_Architecture,$ControlUpAgent_Architecture,$deviceTRUST_Architecture,$FoxitPDFEditor_Language,$FoxitReader_Language,$GitForWindows_Architecture,$GoogleChrome_Architecture,$ImageGlass_Architecture,$IrfanView_Architecture,$Keepass_Language,$MSDotNetFramework_Architecture,$MS365Apps_Architecture,$MS365Apps_Language,$MS365Apps_Visio,$MS365Apps_Visio_Language,$MS365Apps_Project,$MS365Apps_Project_Language,$MSAVDRemoteDesktop_Architecture,$MSEdge_Architecture,$MSFSLogix_Architecture,$MSOffice_Architecture,$MSOneDrive_Architecture,$MSPowerBIDesktop_Architecture,$MSPowerShell_Architecture,$MSSQLServerManagementStudio_Language,$MSTeams_Architecture,$MSVisualStudioCode_Architecture,$Firefox_Architecture,$Firefox_Language,$NotePadPlusPlus_Architecture,$OpenJDK_Architecture,$OracleJava8_Architecture,$PeaZip_Architecture,$Putty_Architecture,$Slack_Architecture,$SumatraPDF_Architecture,$TechSmithSnagIt_Architecture,$VLCPlayer_Architecture,$VMWareTools_Architecture,$WinMerge_Architecture,$Wireshark_Architecture,$IrfanView_Language,$MSOffice_Language,$MSEdgeWebView2,$MSEdgeWebView2_Architecture,$AutodeskDWGTrueView,$MindView7,$MindView7_Language,$PDFsam,$MSOfficeVersion,$OpenShellMenu,$PDFForgeCreator,$TotalCommander,$LogMeInGoToMeeting_Installer,$MSAzureDataStudio_Installer,$MSVisualStudioCode_Installer,$MS365Apps_Installer,$MSOffice_Installer,$MSTeams_Installer,$Zoom_Installer,$MSOneDrive_Installer,$Slack_Installer,$pdfforgePDFCreatorChannel,$TotalCommander_Architecture,$Repository,$CleanUpStartMenu | out-file -filepath "$PSScriptRoot\$GUIfile"
        If ($MS365Apps_Path -ne "") {
            If ($WhatIf -eq '0') {
                Switch ($MS365AppsChannel) {
                    0 { $MS365AppsChannelClear = 'Beta Channel'}
                    1 { $MS365AppsChannelClear = 'Current Channel (Preview)'}
                    2 { $MS365AppsChannelClear = 'Current Channel'}
                    3 { $MS365AppsChannelClear = 'Semi-Annual Channel (Preview)'}
                    4 { $MS365AppsChannelClear = 'Monthly Enterprise Channel'}
                    5 { $MS365AppsChannelClear = 'Semi-Annual Channel'}
                }
                If (!(Test-Path -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear")) {New-Item -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear" -ItemType Directory | Out-Null}
                copy-item -Path "$MS365Apps_Path" -Destination "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear\install.xml" -Force
            }
            Write-Host -ForegroundColor Green "Copy custom XML file as install.xml successfully."
            Write-Output ""
        }
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

    # Button Cancel Detail
    $WPFButton_Cancel_Detail.Add_Click({
        $Script:install = $true
        $Script:download = $true
        Write-Host -Foregroundcolor Red "GUI Mode Canceled - Nothing happens"
        $Form.Close()
        Break
    })

    # Button Save (AddScript)
    $WPFButton_Save.Add_Click({
        If ($WPFCheckbox_Download.IsChecked -eq $True) {$Script:Download = 1}
        Else {$Script:Download = 0}
        If ($WPFCheckbox_Install.IsChecked -eq $True) {$Script:Install = 1}
        Else {$Script:Install = 0}
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
        If ($WPFCheckbox_MSFSLogix.IsChecked -eq $true) {$Script:MSFSLogix = 1}
        Else {$Script:MSFSLogix = 0}
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
        If ($WPFCheckbox_MSEdgeWebView2.ischecked -eq $true) {$Script:MSEdgeWebView2 = 1}
        Else {$Script:MSEdgeWebView2 = 0}
        If ($WPFCheckbox_MSOffice.ischecked -eq $true) {$Script:MSOffice = 1}
        Else {$Script:MSOffice = 0}
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
        If ($WPFCheckbox_GIMP.ischecked -eq $true) {$Script:GIMP = 1}
        Else {$Script:GIMP = 0}
        If ($WPFCheckbox_MSPowerToys.ischecked -eq $true) {$Script:MSPowerToys = 1}
        Else {$Script:MSPowerToys = 0}
        If ($WPFCheckbox_MSVisualStudio.ischecked -eq $true) {$Script:MSVisualStudio = 1}
        Else {$Script:MSVisualStudio = 0}
        If ($WPFCheckbox_MSVisualStudioCode.ischecked -eq $true) {$Script:MSVisualStudioCode = 1}
        Else {$Script:MSVisualStudioCode = 0}
        If ($WPFCheckbox_PaintDotNet.ischecked -eq $true) {$Script:PaintDotNet = 1}
        Else {$Script:PaintDotNet = 0}
        If ($WPFCheckbox_Putty.ischecked -eq $true) {$Script:Putty = 1}
        Else {$Script:Putty = 0}
        If ($WPFCheckbox_TeamViewer.ischecked -eq $true) {$Script:TeamViewer = 1}
        Else {$Script:TeamViewer = 0}
        If ($WPFCheckbox_Wireshark.ischecked -eq $true) {$Script:Wireshark = 1}
        Else {$Script:Wireshark = 0}
        If ($WPFCheckbox_MSAzureDataStudio.ischecked -eq $true) {$Script:MSAzureDataStudio = 1}
        Else {$Script:MSAzureDataStudio = 0}
        If ($WPFCheckbox_ImageGlass.ischecked -eq $true) {$Script:ImageGlass = 1}
        Else {$Script:ImageGlass = 0}
        If ($WPFCheckbox_uberAgent.ischecked -eq $true) {$Script:uberAgent = 1}
        Else {$Script:uberAgent = 0}
        If ($WPFCheckbox_1Password.ischecked -eq $true) {$Script:1Password = 1}
        Else {$Script:1Password = 0}
        If ($WPFCheckbox_ControlUpAgent.ischecked -eq $true) {$Script:ControlUpAgent = 1}
        Else {$Script:ControlUpAgent = 0}
        If ($WPFCheckbox_ControlUpConsole.ischecked -eq $true) {$Script:ControlUpConsole = 1}
        Else {$Script:ControlUpConsole = 0}
        If ($WPFCheckbox_MSSQLServerManagementStudio.ischecked -eq $true) {$Script:MSSQLServerManagementStudio = 1}
        Else {$Script:MSSQLServerManagementStudio = 0}
        If ($WPFCheckbox_MSAVDRemoteDesktop.ischecked -eq $true) {$Script:MSAVDRemoteDesktop = 1}
        Else {$Script:MSAVDRemoteDesktop = 0}
        If ($WPFCheckbox_MSPowerBIDesktop.ischecked -eq $true) {$Script:MSPowerBIDesktop = 1}
        Else {$Script:MSPowerBIDesktop = 0}
        If ($WPFCheckbox_RDAnalyzer.ischecked -eq $true) {$Script:RDAnalyzer = 1}
        Else {$Script:RDAnalyzer = 0}
        If ($WPFCheckbox_SumatraPDF.ischecked -eq $true) {$Script:SumatraPDF = 1}
        Else {$Script:SumatraPDF = 0}
        If ($WPFCheckbox_CiscoWebexTeams.ischecked -eq $true) {$Script:CiscoWebexTeams = 1}
        Else {$Script:CiscoWebexTeams = 0}
        If ($WPFCheckbox_CitrixFiles.ischecked -eq $true) {$Script:CitrixFiles = 1}
        Else {$Script:CitrixFiles = 0}
        If ($WPFCheckbox_FoxitPDFEditor.ischecked -eq $true) {$Script:FoxitPDFEditor = 1}
        Else {$Script:FoxitPDFEditor = 0}
        If ($WPFCheckbox_GitForWindows.ischecked -eq $true) {$Script:GitForWindows = 1}
        Else {$Script:GitForWindows = 0}
        If ($WPFCheckbox_LogMeInGoToMeeting.ischecked -eq $true) {$Script:LogMeInGoToMeeting = 1}
        Else {$Script:LogMeInGoToMeeting = 0}
        If ($WPFCheckbox_MSAzureCLI.ischecked -eq $true) {$Script:MSAzureCLI = 1}
        Else {$Script:MSAzureCLI = 0}
        If ($WPFCheckbox_MSPowerBIReportBuilder.ischecked -eq $true) {$Script:MSPowerBIReportBuilder = 1}
        Else {$Script:MSPowerBIReportBuilder = 0}
        If ($WPFCheckbox_MSSysinternals.ischecked -eq $true) {$Script:MSSysinternals = 1}
        Else {$Script:MSSysinternals = 0}
        If ($WPFCheckbox_Nmap.ischecked -eq $true) {$Script:Nmap = 1}
        Else {$Script:Nmap = 0}
        If ($WPFCheckbox_PeaZip.ischecked -eq $true) {$Script:PeaZip = 1}
        Else {$Script:PeaZip = 0}
        If ($WPFCheckbox_TechSmithCamtasia.ischecked -eq $true) {$Script:TechSmithCamtasia = 1}
        Else {$Script:TechSmithCamtasia = 0}
        If ($WPFCheckbox_TechSmithSnagIt.ischecked -eq $true) {$Script:TechSmithSnagIt = 1}
        Else {$Script:TechSmithSnagIt = 0}
        If ($WPFCheckbox_WinMerge.ischecked -eq $true) {$Script:WinMerge = 1}
        Else {$Script:WinMerge = 0}
        If ($WPFCheckbox_WhatIf.ischecked -eq $true) {$Script:WhatIf = 1}
        Else {$Script:WhatIf = 0}
        If ($WPFCheckbox_CleanUp.ischecked -eq $true) {$Script:CleanUp = 1}
        Else {$Script:CleanUp = 0}
        If ($WPFCheckbox_MS365Apps_Visio_Detail.ischecked -eq $true) {$Script:MS365Apps_Visio = 1}
        Else {$Script:MS365Apps_Visio = 0}
        If ($WPFCheckbox_MS365Apps_Project_Detail.ischecked -eq $true) {$Script:MS365Apps_Project = 1}
        Else {$Script:MS365Apps_Project = 0}
        If ($WPFCheckbox_AutodeskDWGTrueView.ischecked -eq $true) {$Script:AutodeskDWGTrueView = 1}
        Else {$Script:AutodeskDWGTrueView = 0}
        If ($WPFCheckbox_MindView7.ischecked -eq $true) {$Script:MindView7 = 1}
        Else {$Script:MindView7 = 0}
        If ($WPFCheckbox_PDFsam.ischecked -eq $true) {$Script:PDFsam = 1}
        Else {$Script:PDFsam = 0}
        If ($WPFCheckbox_OpenShellMenu.ischecked -eq $true) {$Script:OpenShellMenu = 1}
        Else {$Script:OpenShellMenu = 0}
        If ($WPFCheckbox_PDFForgeCreator.ischecked -eq $true) {$Script:PDFForgeCreator = 1}
        Else {$Script:PDFForgeCreator = 0}
        If ($WPFCheckbox_TotalCommander.ischecked -eq $true) {$Script:TotalCommander = 1}
        Else {$Script:TotalCommander = 0}
        If ($WPFCheckbox_Repository.ischecked -eq $true) {$Script:Repository = 1}
        Else {$Script:Repository = 0}
        If ($WPFCheckbox_CleanUpStartMenu.ischecked -eq $true) {$Script:CleanUpStartMenu = 1}
        Else {$Script:CleanUpStartMenu = 0}
        $Script:Language = $WPFBox_Language.SelectedIndex
        $Script:Architecture = $WPFBox_Architecture.SelectedIndex
        $Script:Installer = $WPFBox_Installer.SelectedIndex
        $Script:FirefoxChannel = $WPFBox_Firefox.SelectedIndex
        $Script:CitrixWorkspaceAppRelease = $WPFBox_CitrixWorkspaceApp.SelectedIndex
        $Script:MS365AppsChannel = $WPFBox_MS365Apps.SelectedIndex
        $Script:MSOneDriveRing = $WPFBox_MSOneDrive.SelectedIndex
        $Script:MSTeamsRing = $WPFBox_MSTeams.SelectedIndex
        $Script:TreeSizeType = $WPFBox_TreeSize.SelectedIndex
        $Script:MSDotNetFrameworkChannel = $WPFBox_MSDotNetFramework.SelectedIndex
        $Script:MSPowerShellRelease = $WPFBox_MSPowerShell.SelectedIndex
        $Script:RemoteDesktopManagerType = $WPFBox_RemoteDesktopManager.SelectedIndex
        $Script:ZoomCitrixClient = $WPFBox_Zoom.SelectedIndex
        $Script:deviceTRUSTPackage = $WPFBox_deviceTRUST.SelectedIndex
        $Script:MSEdgeChannel = $WPFBox_MSEdge.SelectedIndex
        $Script:MSVisualStudioCodeChannel = $WPFBox_MSVisualStudioCode.SelectedIndex
        $Script:MSVisualStudioEdition = $WPFBox_MSVisualStudio.SelectedIndex
        $Script:PuttyChannel = $WPFBox_Putty.SelectedIndex
        $Script:MSAzureDataStudioChannel = $WPFBox_MSAzureDataStudio.SelectedIndex
        $Script:MSFSLogixChannel = $WPFBox_MSFSLogix.SelectedIndex
        $Script:ControlUpAgentFramework = $WPFBox_ControlUpAgent.SelectedIndex
        $Script:MSAVDRemoteDesktopChannel = $WPFBox_MSAVDRemoteDesktop.SelectedIndex
        $Script:7Zip_Architecture = $WPFBox_7Zip_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Architecture = $WPFBox_AdobeReaderDC_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Language = $WPFBox_AdobeReaderDC_Language.SelectedIndex
        $Script:CiscoWebexTeams_Architecture = $WPFBox_CiscoWebexTeams_Architecture.SelectedIndex
        $Script:CitrixHypervisorTools_Architecture = $WPFBox_CitrixHypervisorTools_Architecture.SelectedIndex
        $Script:ControlUpAgent_Architecture = $WPFBox_ControlUpAgent_Architecture.SelectedIndex
        $Script:deviceTRUST_Architecture = $WPFBox_deviceTRUST_Architecture.SelectedIndex
        $Script:FoxitPDFEditor_Language = $WPFBox_FoxitPDFEditor_Language.SelectedIndex
        $Script:FoxitReader_Language = $WPFBox_FoxitReader_Language.SelectedIndex
        $Script:GitForWindows_Architecture = $WPFBox_GitForWindows_Architecture.SelectedIndex
        $Script:GoogleChrome_Architecture = $WPFBox_GoogleChrome_Architecture.SelectedIndex
        $Script:ImageGlass_Architecture = $WPFBox_ImageGlass_Architecture.SelectedIndex
        $Script:IrfanView_Architecture = $WPFBox_IrfanView_Architecture.SelectedIndex
        $Script:Keepass_Language = $WPFBox_KeePass_Language.SelectedIndex
        $Script:MSDotNetFramework_Architecture = $WPFBox_MSDotNetFramework_Architecture.SelectedIndex
        $Script:MS365Apps_Architecture = $WPFBox_MS365Apps_Architecture.SelectedIndex
        $Script:MS365Apps_Language = $WPFBox_MS365Apps_Language.SelectedIndex
        $Script:MS365Apps_Visio_Language = $WPFBox_MS365Apps_Visio_Language.SelectedIndex
        $Script:MS365Apps_Project_Language = $WPFBox_MS365Apps_Project_Language.SelectedIndex
        $Script:MSAVDRemoteDesktop_Architecture = $WPFBox_MSAVDRemoteDesktop_Architecture.SelectedIndex
        $Script:MSEdge_Architecture = $WPFBox_MSEdge_Architecture.SelectedIndex
        $Script:MSFSLogix_Architecture = $WPFBox_MSFSLogix_Architecture.SelectedIndex
        $Script:MSOffice_Architecture = $WPFBox_MSOffice_Architecture.SelectedIndex
        $Script:MSOneDrive_Architecture = $WPFBox_MSOneDrive_Architecture.SelectedIndex
        $Script:MSPowerBIDesktop_Architecture = $WPFBox_MSPowerBIDesktop_Architecture.SelectedIndex
        $Script:MSPowerShell_Architecture = $WPFBox_MSPowerShell_Architecture.SelectedIndex
        $Script:MSSQLServerManagementStudio_Language = $WPFBox_MSSQLServerManagementStudio_Language.SelectedIndex
        $Script:MSTeams_Architecture = $WPFBox_MSTeams_Architecture.SelectedIndex
        $Script:MSVisualStudioCode_Architecture = $WPFBox_MSVisualStudioCode_Architecture.SelectedIndex
        $Script:Firefox_Architecture = $WPFBox_Firefox_Architecture.SelectedIndex
        $Script:Firefox_Language = $WPFBox_Firefox_Language.SelectedIndex
        $Script:NotePadPlusPlus_Architecture = $WPFBox_NotepadPlusPlus_Architecture.SelectedIndex
        $Script:OpenJDK_Architecture = $WPFBox_OpenJDK_Architecture.SelectedIndex
        $Script:OracleJava8_Architecture = $WPFBox_OracleJava8_Architecture.SelectedIndex
        $Script:PeaZip_Architecture = $WPFBox_PeaZip_Architecture.SelectedIndex
        $Script:Putty_Architecture = $WPFBox_Putty_Architecture.SelectedIndex
        $Script:Slack_Architecture = $WPFBox_Slack_Architecture.SelectedIndex
        $Script:SumatraPDF_Architecture = $WPFBox_SumatraPDF_Architecture.SelectedIndex
        $Script:TechSmithSnagIt_Architecture = $WPFBox_TechSmithSnagIT_Architecture.SelectedIndex
        $Script:VLCPlayer_Architecture = $WPFBox_VLCPlayer_Architecture.SelectedIndex
        $Script:VMWareTools_Architecture = $WPFBox_VMWareTools_Architecture.SelectedIndex
        $Script:WinMerge_Architecture = $WPFBox_WinMerge_Architecture.SelectedIndex
        $Script:Wireshark_Architecture = $WPFBox_Wireshark_Architecture.SelectedIndex
        $Script:IrfanView_Language = $WPFBox_IrfanView_Language.SelectedIndex
        $Script:MSOffice_Language = $WPFBox_MSOffice_Language.SelectedIndex
        $Script:MS365Apps_Path = $WPFTextBox_Filename.Text
        $Script:MSEdgeWebView2_Architecture = $WPFBox_MSEdgeWebView2_Architecture.SelectedIndex
        $Script:MindView7_Language = $WPFBox_MindView7_Language.SelectedIndex
        $Script:MSOfficeVersion = $WPFBox_MSOffice.SelectedIndex
        $Script:LogMeInGoToMeeting_Installer = $WPFBox_LogMeInGoToMeeting_Installer.SelectedIndex
        $Script:MSAzureDataStudio_Installer = $WPFBox_MSAzureDataStudio_Installer.SelectedIndex
        $Script:MSVisualStudioCode_Installer = $WPFBox_MSVisualStudioCode_Installer.SelectedIndex
        $Script:MS365Apps_Installer = $WPFBox_MS365Apps_Installer.SelectedIndex
        $Script:MSOffice_Installer = $WPFBox_MSOffice_Installer.SelectedIndex
        $Script:MSTeams_Installer = $WPFBox_MSTeams_Installer.SelectedIndex
        $Script:Zoom_Installer = $WPFBox_Zoom_Installer.SelectedIndex
        $Script:MSOneDrive_Installer = $WPFBox_MSOneDrive_Installer.SelectedIndex
        $Script:Slack_Installer = $WPFBox_Slack_Installer.SelectedIndex
        $Script:pdfforgePDFCreatorChannel = $WPFBox_pdfforgePDFCreator.SelectedIndex
        $Script:TotalCommander_Architecture = $WPFBox_TotalCommander_Architecture.SelectedIndex

        # Write LastSetting.txt or -GUIFile Parameter file to get the settings of the last session. (AddScript)
        $Language,$Architecture,$CitrixWorkspaceAppRelease,$MS365AppsChannel,$MSOneDriveRing,$MSTeamsRing,$FirefoxChannel,$TreeSizeType,$7ZIP,$AdobeProDC,$AdobeReaderDC,$BISF,$Citrix_Hypervisor_Tools,$Citrix_WorkspaceApp,$Filezilla,$Firefox,$Foxit_Reader,$MSFSLogix,$GoogleChrome,$Greenshot,$KeePass,$mRemoteNG,$MS365Apps,$MSEdge,$MSOffice,$MSOneDrive,$MSTeams,$NotePadPlusPlus,$OpenJDK,$OracleJava8,$TreeSize,$VLCPlayer,$VMWareTools,$WinSCP,$Download,$Install,$IrfanView,$MSTeamsNoAutoStart,$deviceTRUST,$MSDotNetFramework,$MSDotNetFrameworkChannel,$MSPowerShell,$MSPowerShellRelease,$RemoteDesktopManager,$RemoteDesktopManagerType,$Slack,$Wireshark,$ShareX,$Zoom,$ZoomCitrixClient,$deviceTRUSTPackage,$MSEdgeChannel,$GIMP,$MSPowerToys,$MSVisualStudio,$MSVisualStudioCode,$MSVisualStudioCodeChannel,$PaintDotNet,$Putty,$TeamViewer,$Installer,$MSVisualStudioEdition,$PuttyChannel,$MSAzureDataStudio,$MSAzureDataStudioChannel,$ImageGlass,$MSFSLogixChannel,$uberAgent,$1Password,$SumatraPDF,$ControlUpAgent,$ControlUpAgentFramework,$ControlUpConsole,$MSSQLServerManagementStudio,$MSAVDRemoteDesktop,$MSAVDRemoteDesktopChannel,$MSPowerBIDesktop,$RDAnalyzer,$CiscoWebexTeams,$CitrixFiles,$FoxitPDFEditor,$GitForWindows,$LogMeInGoToMeeting,$MSAzureCLI,$MSPowerBIReportBuilder,$MSSysinternals,$NMap,$PeaZip,$TechSmithCamtasia,$TechSmithSnagit,$WinMerge,$WhatIf,$CleanUp,$7Zip_Architecture,$AdobeReaderDC_Architecture,$AdobeReaderDC_Language,$CiscoWebexTeams_Architecture,$CitrixHypervisorTools_Architecture,$ControlUpAgent_Architecture,$deviceTRUST_Architecture,$FoxitPDFEditor_Language,$FoxitReader_Language,$GitForWindows_Architecture,$GoogleChrome_Architecture,$ImageGlass_Architecture,$IrfanView_Architecture,$Keepass_Language,$MSDotNetFramework_Architecture,$MS365Apps_Architecture,$MS365Apps_Language,$MS365Apps_Visio,$MS365Apps_Visio_Language,$MS365Apps_Project,$MS365Apps_Project_Language,$MSAVDRemoteDesktop_Architecture,$MSEdge_Architecture,$MSFSLogix_Architecture,$MSOffice_Architecture,$MSOneDrive_Architecture,$MSPowerBIDesktop_Architecture,$MSPowerShell_Architecture,$MSSQLServerManagementStudio_Language,$MSTeams_Architecture,$MSVisualStudioCode_Architecture,$Firefox_Architecture,$Firefox_Language,$NotePadPlusPlus_Architecture,$OpenJDK_Architecture,$OracleJava8_Architecture,$PeaZip_Architecture,$Putty_Architecture,$Slack_Architecture,$SumatraPDF_Architecture,$TechSmithSnagIt_Architecture,$VLCPlayer_Architecture,$VMWareTools_Architecture,$WinMerge_Architecture,$Wireshark_Architecture,$IrfanView_Language,$MSOffice_Language,$MSEdgeWebView2,$MSEdgeWebView2_Architecture,$AutodeskDWGTrueView,$MindView7,$MindView7_Language,$PDFsam,$MSOfficeVersion,$OpenShellMenu,$PDFForgeCreator,$TotalCommander,$LogMeInGoToMeeting_Installer,$MSAzureDataStudio_Installer,$MSVisualStudioCode_Installer,$MS365Apps_Installer,$MSOffice_Installer,$MSTeams_Installer,$Zoom_Installer,$MSOneDrive_Installer,$Slack_Installer,$pdfforgePDFCreatorChannel,$TotalCommander_Architecture,$Repository,$CleanUpStartMenu | out-file -filepath "$PSScriptRoot\$GUIfile"
        If ($MS365Apps_Path -ne "") {
            If ($WhatIf -eq '0') {
                Switch ($MS365AppsChannel) {
                    0 { $MS365AppsChannelClear = 'Beta Channel'}
                    1 { $MS365AppsChannelClear = 'Current Channel (Preview)'}
                    2 { $MS365AppsChannelClear = 'Current Channel'}
                    3 { $MS365AppsChannelClear = 'Semi-Annual Channel (Preview)'}
                    4 { $MS365AppsChannelClear = 'Monthly Enterprise Channel'}
                    5 { $MS365AppsChannelClear = 'Semi-Annual Channel'}
                }
                If (!(Test-Path -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear")) {New-Item -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear" -ItemType Directory | Out-Null}
                copy-item -Path "$MS365Apps_Path" -Destination "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear\install.xml" -Force
            }
            Write-Host -ForegroundColor Green "Copy custom XML file as install.xml successfully."
            Write-Output ""
        }
        Write-Host "Save Settings"
    })

    # Button Save Detail (AddScript)
    $WPFButton_Save_Detail.Add_Click({
        If ($WPFCheckbox_Download.IsChecked -eq $True) {$Script:Download = 1}
        Else {$Script:Download = 0}
        If ($WPFCheckbox_Install.IsChecked -eq $True) {$Script:Install = 1}
        Else {$Script:Install = 0}
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
        If ($WPFCheckbox_MSFSLogix.IsChecked -eq $true) {$Script:MSFSLogix = 1}
        Else {$Script:MSFSLogix = 0}
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
        If ($WPFCheckbox_MSEdgeWebView2.ischecked -eq $true) {$Script:MSEdgeWebView2 = 1}
        Else {$Script:MSEdgeWebView2 = 0}
        If ($WPFCheckbox_MSOffice.ischecked -eq $true) {$Script:MSOffice = 1}
        Else {$Script:MSOffice = 0}
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
        If ($WPFCheckbox_GIMP.ischecked -eq $true) {$Script:GIMP = 1}
        Else {$Script:GIMP = 0}
        If ($WPFCheckbox_MSPowerToys.ischecked -eq $true) {$Script:MSPowerToys = 1}
        Else {$Script:MSPowerToys = 0}
        If ($WPFCheckbox_MSVisualStudio.ischecked -eq $true) {$Script:MSVisualStudio = 1}
        Else {$Script:MSVisualStudio = 0}
        If ($WPFCheckbox_MSVisualStudioCode.ischecked -eq $true) {$Script:MSVisualStudioCode = 1}
        Else {$Script:MSVisualStudioCode = 0}
        If ($WPFCheckbox_PaintDotNet.ischecked -eq $true) {$Script:PaintDotNet = 1}
        Else {$Script:PaintDotNet = 0}
        If ($WPFCheckbox_Putty.ischecked -eq $true) {$Script:Putty = 1}
        Else {$Script:Putty = 0}
        If ($WPFCheckbox_TeamViewer.ischecked -eq $true) {$Script:TeamViewer = 1}
        Else {$Script:TeamViewer = 0}
        If ($WPFCheckbox_Wireshark.ischecked -eq $true) {$Script:Wireshark = 1}
        Else {$Script:Wireshark = 0}
        If ($WPFCheckbox_MSAzureDataStudio.ischecked -eq $true) {$Script:MSAzureDataStudio = 1}
        Else {$Script:MSAzureDataStudio = 0}
        If ($WPFCheckbox_ImageGlass.ischecked -eq $true) {$Script:ImageGlass = 1}
        Else {$Script:ImageGlass = 0}
        If ($WPFCheckbox_uberAgent.ischecked -eq $true) {$Script:uberAgent = 1}
        Else {$Script:uberAgent = 0}
        If ($WPFCheckbox_1Password.ischecked -eq $true) {$Script:1Password = 1}
        Else {$Script:1Password = 0}
        If ($WPFCheckbox_ControlUpAgent.ischecked -eq $true) {$Script:ControlUpAgent = 1}
        Else {$Script:ControlUpAgent = 0}
        If ($WPFCheckbox_ControlUpConsole.ischecked -eq $true) {$Script:ControlUpConsole = 1}
        Else {$Script:ControlUpConsole = 0}
        If ($WPFCheckbox_MSSQLServerManagementStudio.ischecked -eq $true) {$Script:MSSQLServerManagementStudio = 1}
        Else {$Script:MSSQLServerManagementStudio = 0}
        If ($WPFCheckbox_MSAVDRemoteDesktop.ischecked -eq $true) {$Script:MSAVDRemoteDesktop = 1}
        Else {$Script:MSAVDRemoteDesktop = 0}
        If ($WPFCheckbox_MSPowerBIDesktop.ischecked -eq $true) {$Script:MSPowerBIDesktop = 1}
        Else {$Script:MSPowerBIDesktop = 0}
        If ($WPFCheckbox_RDAnalyzer.ischecked -eq $true) {$Script:RDAnalyzer = 1}
        Else {$Script:RDAnalyzer = 0}
        If ($WPFCheckbox_SumatraPDF.ischecked -eq $true) {$Script:SumatraPDF = 1}
        Else {$Script:SumatraPDF = 0}
        If ($WPFCheckbox_CiscoWebexTeams.ischecked -eq $true) {$Script:CiscoWebexTeams = 1}
        Else {$Script:CiscoWebexTeams = 0}
        If ($WPFCheckbox_CitrixFiles.ischecked -eq $true) {$Script:CitrixFiles = 1}
        Else {$Script:CitrixFiles = 0}
        If ($WPFCheckbox_FoxitPDFEditor.ischecked -eq $true) {$Script:FoxitPDFEditor = 1}
        Else {$Script:FoxitPDFEditor = 0}
        If ($WPFCheckbox_GitForWindows.ischecked -eq $true) {$Script:GitForWindows = 1}
        Else {$Script:GitForWindows = 0}
        If ($WPFCheckbox_LogMeInGoToMeeting.ischecked -eq $true) {$Script:LogMeInGoToMeeting = 1}
        Else {$Script:LogMeInGoToMeeting = 0}
        If ($WPFCheckbox_MSAzureCLI.ischecked -eq $true) {$Script:MSAzureCLI = 1}
        Else {$Script:MSAzureCLI = 0}
        If ($WPFCheckbox_MSPowerBIReportBuilder.ischecked -eq $true) {$Script:MSPowerBIReportBuilder = 1}
        Else {$Script:MSPowerBIReportBuilder = 0}
        If ($WPFCheckbox_MSSysinternals.ischecked -eq $true) {$Script:MSSysinternals = 1}
        Else {$Script:MSSysinternals = 0}
        If ($WPFCheckbox_Nmap.ischecked -eq $true) {$Script:Nmap = 1}
        Else {$Script:Nmap = 0}
        If ($WPFCheckbox_PeaZip.ischecked -eq $true) {$Script:PeaZip = 1}
        Else {$Script:PeaZip = 0}
        If ($WPFCheckbox_TechSmithCamtasia.ischecked -eq $true) {$Script:TechSmithCamtasia = 1}
        Else {$Script:TechSmithCamtasia = 0}
        If ($WPFCheckbox_TechSmithSnagIt.ischecked -eq $true) {$Script:TechSmithSnagIt = 1}
        Else {$Script:TechSmithSnagIt = 0}
        If ($WPFCheckbox_WinMerge.ischecked -eq $true) {$Script:WinMerge = 1}
        Else {$Script:WinMerge = 0}
        If ($WPFCheckbox_WhatIf.ischecked -eq $true) {$Script:WhatIf = 1}
        Else {$Script:WhatIf = 0}
        If ($WPFCheckbox_CleanUp.ischecked -eq $true) {$Script:CleanUp = 1}
        Else {$Script:CleanUp = 0}
        If ($WPFCheckbox_MS365Apps_Visio_Detail.ischecked -eq $true) {$Script:MS365Apps_Visio = 1}
        Else {$Script:MS365Apps_Visio = 0}
        If ($WPFCheckbox_MS365Apps_Project_Detail.ischecked -eq $true) {$Script:MS365Apps_Project = 1}
        Else {$Script:MS365Apps_Project = 0}
        If ($WPFCheckbox_AutodeskDWGTrueView.ischecked -eq $true) {$Script:AutodeskDWGTrueView = 1}
        Else {$Script:AutodeskDWGTrueView = 0}
        If ($WPFCheckbox_MindView7.ischecked -eq $true) {$Script:MindView7 = 1}
        Else {$Script:MindView7 = 0}
        If ($WPFCheckbox_PDFsam.ischecked -eq $true) {$Script:PDFsam = 1}
        Else {$Script:PDFsam = 0}
        If ($WPFCheckbox_OpenShellMenu.ischecked -eq $true) {$Script:OpenShellMenu = 1}
        Else {$Script:OpenShellMenu = 0}
        If ($WPFCheckbox_PDFForgeCreator.ischecked -eq $true) {$Script:PDFForgeCreator = 1}
        Else {$Script:PDFForgeCreator = 0}
        If ($WPFCheckbox_TotalCommander.ischecked -eq $true) {$Script:TotalCommander = 1}
        Else {$Script:TotalCommander = 0}
        If ($WPFCheckbox_Repository.ischecked -eq $true) {$Script:Repository = 1}
        Else {$Script:Repository = 0}
        If ($WPFCheckbox_CleanUpStartMenu.ischecked -eq $true) {$Script:CleanUpStartMenu = 1}
        Else {$Script:CleanUpStartMenu = 0}
        $Script:Language = $WPFBox_Language.SelectedIndex
        $Script:Architecture = $WPFBox_Architecture.SelectedIndex
        $Script:Installer = $WPFBox_Installer.SelectedIndex
        $Script:FirefoxChannel = $WPFBox_Firefox.SelectedIndex
        $Script:CitrixWorkspaceAppRelease = $WPFBox_CitrixWorkspaceApp.SelectedIndex
        $Script:MS365AppsChannel = $WPFBox_MS365Apps.SelectedIndex
        $Script:MSOneDriveRing = $WPFBox_MSOneDrive.SelectedIndex
        $Script:MSTeamsRing = $WPFBox_MSTeams.SelectedIndex
        $Script:TreeSizeType = $WPFBox_TreeSize.SelectedIndex
        $Script:MSDotNetFrameworkChannel = $WPFBox_MSDotNetFramework.SelectedIndex
        $Script:MSPowerShellRelease = $WPFBox_MSPowerShell.SelectedIndex
        $Script:RemoteDesktopManagerType = $WPFBox_RemoteDesktopManager.SelectedIndex
        $Script:ZoomCitrixClient = $WPFBox_Zoom.SelectedIndex
        $Script:deviceTRUSTPackage = $WPFBox_deviceTRUST.SelectedIndex
        $Script:MSEdgeChannel = $WPFBox_MSEdge.SelectedIndex
        $Script:MSVisualStudioCodeChannel = $WPFBox_MSVisualStudioCode.SelectedIndex
        $Script:MSVisualStudioEdition = $WPFBox_MSVisualStudio.SelectedIndex
        $Script:PuttyChannel = $WPFBox_Putty.SelectedIndex
        $Script:MSAzureDataStudioChannel = $WPFBox_MSAzureDataStudio.SelectedIndex
        $Script:MSFSLogixChannel = $WPFBox_MSFSLogix.SelectedIndex
        $Script:ControlUpAgentFramework = $WPFBox_ControlUpAgent.SelectedIndex
        $Script:MSAVDRemoteDesktopChannel = $WPFBox_MSAVDRemoteDesktop.SelectedIndex
        $Script:7Zip_Architecture = $WPFBox_7Zip_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Architecture = $WPFBox_AdobeReaderDC_Architecture.SelectedIndex
        $Script:AdobeReaderDC_Language = $WPFBox_AdobeReaderDC_Language.SelectedIndex
        $Script:CiscoWebexTeams_Architecture = $WPFBox_CiscoWebexTeams_Architecture.SelectedIndex
        $Script:CitrixHypervisorTools_Architecture = $WPFBox_CitrixHypervisorTools_Architecture.SelectedIndex
        $Script:ControlUpAgent_Architecture = $WPFBox_ControlUpAgent_Architecture.SelectedIndex
        $Script:deviceTRUST_Architecture = $WPFBox_deviceTRUST_Architecture.SelectedIndex
        $Script:FoxitPDFEditor_Language = $WPFBox_FoxitPDFEditor_Language.SelectedIndex
        $Script:FoxitReader_Language = $WPFBox_FoxitReader_Language.SelectedIndex
        $Script:GitForWindows_Architecture = $WPFBox_GitForWindows_Architecture.SelectedIndex
        $Script:GoogleChrome_Architecture = $WPFBox_GoogleChrome_Architecture.SelectedIndex
        $Script:ImageGlass_Architecture = $WPFBox_ImageGlass_Architecture.SelectedIndex
        $Script:IrfanView_Architecture = $WPFBox_IrfanView_Architecture.SelectedIndex
        $Script:Keepass_Language = $WPFBox_KeePass_Language.SelectedIndex
        $Script:MSDotNetFramework_Architecture = $WPFBox_MSDotNetFramework_Architecture.SelectedIndex
        $Script:MS365Apps_Architecture = $WPFBox_MS365Apps_Architecture.SelectedIndex
        $Script:MS365Apps_Language = $WPFBox_MS365Apps_Language.SelectedIndex
        $Script:MS365Apps_Visio_Language = $WPFBox_MS365Apps_Visio_Language.SelectedIndex
        $Script:MS365Apps_Project_Language = $WPFBox_MS365Apps_Project_Language.SelectedIndex
        $Script:MSAVDRemoteDesktop_Architecture = $WPFBox_MSAVDRemoteDesktop_Architecture.SelectedIndex
        $Script:MSEdge_Architecture = $WPFBox_MSEdge_Architecture.SelectedIndex
        $Script:MSFSLogix_Architecture = $WPFBox_MSFSLogix_Architecture.SelectedIndex
        $Script:MSOffice_Architecture = $WPFBox_MSOffice_Architecture.SelectedIndex
        $Script:MSOneDrive_Architecture = $WPFBox_MSOneDrive_Architecture.SelectedIndex
        $Script:MSPowerBIDesktop_Architecture = $WPFBox_MSPowerBIDesktop_Architecture.SelectedIndex
        $Script:MSPowerShell_Architecture = $WPFBox_MSPowerShell_Architecture.SelectedIndex
        $Script:MSSQLServerManagementStudio_Language = $WPFBox_MSSQLServerManagementStudio_Language.SelectedIndex
        $Script:MSTeams_Architecture = $WPFBox_MSTeams_Architecture.SelectedIndex
        $Script:MSVisualStudioCode_Architecture = $WPFBox_MSVisualStudioCode_Architecture.SelectedIndex
        $Script:Firefox_Architecture = $WPFBox_Firefox_Architecture.SelectedIndex
        $Script:Firefox_Language = $WPFBox_Firefox_Language.SelectedIndex
        $Script:NotePadPlusPlus_Architecture = $WPFBox_NotepadPlusPlus_Architecture.SelectedIndex
        $Script:OpenJDK_Architecture = $WPFBox_OpenJDK_Architecture.SelectedIndex
        $Script:OracleJava8_Architecture = $WPFBox_OracleJava8_Architecture.SelectedIndex
        $Script:PeaZip_Architecture = $WPFBox_PeaZip_Architecture.SelectedIndex
        $Script:Putty_Architecture = $WPFBox_Putty_Architecture.SelectedIndex
        $Script:Slack_Architecture = $WPFBox_Slack_Architecture.SelectedIndex
        $Script:SumatraPDF_Architecture = $WPFBox_SumatraPDF_Architecture.SelectedIndex
        $Script:TechSmithSnagIt_Architecture = $WPFBox_TechSmithSnagIT_Architecture.SelectedIndex
        $Script:VLCPlayer_Architecture = $WPFBox_VLCPlayer_Architecture.SelectedIndex
        $Script:VMWareTools_Architecture = $WPFBox_VMWareTools_Architecture.SelectedIndex
        $Script:WinMerge_Architecture = $WPFBox_WinMerge_Architecture.SelectedIndex
        $Script:Wireshark_Architecture = $WPFBox_Wireshark_Architecture.SelectedIndex
        $Script:IrfanView_Language = $WPFBox_IrfanView_Language.SelectedIndex
        $Script:MSOffice_Language = $WPFBox_MSOffice_Language.SelectedIndex
        $Script:MS365Apps_Path = $WPFTextBox_Filename.Text
        $Script:MSEdgeWebView2_Architecture = $WPFBox_MSEdgeWebView2_Architecture.SelectedIndex
        $Script:MindView7_Language = $WPFBox_MindView7_Language.SelectedIndex
        $Script:MSOfficeVersion = $WPFBox_MSOffice.SelectedIndex
        $Script:LogMeInGoToMeeting_Installer = $WPFBox_LogMeInGoToMeeting_Installer.SelectedIndex
        $Script:MSAzureDataStudio_Installer = $WPFBox_MSAzureDataStudio_Installer.SelectedIndex
        $Script:MSVisualStudioCode_Installer = $WPFBox_MSVisualStudioCode_Installer.SelectedIndex
        $Script:MS365Apps_Installer = $WPFBox_MS365Apps_Installer.SelectedIndex
        $Script:MSOffice_Installer = $WPFBox_MSOffice_Installer.SelectedIndex
        $Script:MSTeams_Installer = $WPFBox_MSTeams_Installer.SelectedIndex
        $Script:Zoom_Installer = $WPFBox_Zoom_Installer.SelectedIndex
        $Script:MSOneDrive_Installer = $WPFBox_MSOneDrive_Installer.SelectedIndex
        $Script:Slack_Installer = $WPFBox_Slack_Installer.SelectedIndex
        $Script:pdfforgePDFCreatorChannel = $WPFBox_pdfforgePDFCreator.SelectedIndex
        $Script:TotalCommander_Architecture = $WPFBox_TotalCommander_Architecture.SelectedIndex

        # Write LastSetting.txt or -GUIFile Parameter file to get the settings of the last session. (AddScript)
        $Language,$Architecture,$CitrixWorkspaceAppRelease,$MS365AppsChannel,$MSOneDriveRing,$MSTeamsRing,$FirefoxChannel,$TreeSizeType,$7ZIP,$AdobeProDC,$AdobeReaderDC,$BISF,$Citrix_Hypervisor_Tools,$Citrix_WorkspaceApp,$Filezilla,$Firefox,$Foxit_Reader,$MSFSLogix,$GoogleChrome,$Greenshot,$KeePass,$mRemoteNG,$MS365Apps,$MSEdge,$MSOffice,$MSOneDrive,$MSTeams,$NotePadPlusPlus,$OpenJDK,$OracleJava8,$TreeSize,$VLCPlayer,$VMWareTools,$WinSCP,$Download,$Install,$IrfanView,$MSTeamsNoAutoStart,$deviceTRUST,$MSDotNetFramework,$MSDotNetFrameworkChannel,$MSPowerShell,$MSPowerShellRelease,$RemoteDesktopManager,$RemoteDesktopManagerType,$Slack,$Wireshark,$ShareX,$Zoom,$ZoomCitrixClient,$deviceTRUSTPackage,$MSEdgeChannel,$GIMP,$MSPowerToys,$MSVisualStudio,$MSVisualStudioCode,$MSVisualStudioCodeChannel,$PaintDotNet,$Putty,$TeamViewer,$Installer,$MSVisualStudioEdition,$PuttyChannel,$MSAzureDataStudio,$MSAzureDataStudioChannel,$ImageGlass,$MSFSLogixChannel,$uberAgent,$1Password,$SumatraPDF,$ControlUpAgent,$ControlUpAgentFramework,$ControlUpConsole,$MSSQLServerManagementStudio,$MSAVDRemoteDesktop,$MSAVDRemoteDesktopChannel,$MSPowerBIDesktop,$RDAnalyzer,$CiscoWebexTeams,$CitrixFiles,$FoxitPDFEditor,$GitForWindows,$LogMeInGoToMeeting,$MSAzureCLI,$MSPowerBIReportBuilder,$MSSysinternals,$NMap,$PeaZip,$TechSmithCamtasia,$TechSmithSnagit,$WinMerge,$WhatIf,$CleanUp,$7Zip_Architecture,$AdobeReaderDC_Architecture,$AdobeReaderDC_Language,$CiscoWebexTeams_Architecture,$CitrixHypervisorTools_Architecture,$ControlUpAgent_Architecture,$deviceTRUST_Architecture,$FoxitPDFEditor_Language,$FoxitReader_Language,$GitForWindows_Architecture,$GoogleChrome_Architecture,$ImageGlass_Architecture,$IrfanView_Architecture,$Keepass_Language,$MSDotNetFramework_Architecture,$MS365Apps_Architecture,$MS365Apps_Language,$MS365Apps_Visio,$MS365Apps_Visio_Language,$MS365Apps_Project,$MS365Apps_Project_Language,$MSAVDRemoteDesktop_Architecture,$MSEdge_Architecture,$MSFSLogix_Architecture,$MSOffice_Architecture,$MSOneDrive_Architecture,$MSPowerBIDesktop_Architecture,$MSPowerShell_Architecture,$MSSQLServerManagementStudio_Language,$MSTeams_Architecture,$MSVisualStudioCode_Architecture,$Firefox_Architecture,$Firefox_Language,$NotePadPlusPlus_Architecture,$OpenJDK_Architecture,$OracleJava8_Architecture,$PeaZip_Architecture,$Putty_Architecture,$Slack_Architecture,$SumatraPDF_Architecture,$TechSmithSnagIt_Architecture,$VLCPlayer_Architecture,$VMWareTools_Architecture,$WinMerge_Architecture,$Wireshark_Architecture,$IrfanView_Language,$MSOffice_Language,$MSEdgeWebView2,$MSEdgeWebView2_Architecture,$AutodeskDWGTrueView,$MindView7,$MindView7_Language,$PDFsam,$MSOfficeVersion,$OpenShellMenu,$PDFForgeCreator,$TotalCommander,$LogMeInGoToMeeting_Installer,$MSAzureDataStudio_Installer,$MSVisualStudioCode_Installer,$MS365Apps_Installer,$MSOffice_Installer,$MSTeams_Installer,$Zoom_Installer,$MSOneDrive_Installer,$Slack_Installer,$pdfforgePDFCreatorChannel,$TotalCommander_Architecture,$Repository,$CleanUpStartMenu | out-file -filepath "$PSScriptRoot\$GUIfile"
        If ($MS365Apps_Path -ne "") {
            If ($WhatIf -eq '0') {
                Switch ($MS365AppsChannel) {
                    0 { $MS365AppsChannelClear = 'Beta Channel'}
                    1 { $MS365AppsChannelClear = 'Current Channel (Preview)'}
                    2 { $MS365AppsChannelClear = 'Current Channel'}
                    3 { $MS365AppsChannelClear = 'Semi-Annual Channel (Preview)'}
                    4 { $MS365AppsChannelClear = 'Monthly Enterprise Channel'}
                    5 { $MS365AppsChannelClear = 'Semi-Annual Channel'}
                }
                If (!(Test-Path -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear")) {New-Item -Path "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear" -ItemType Directory | Out-Null}
                copy-item -Path "$MS365Apps_Path" -Destination "$PSScriptRoot\Microsoft 365 Apps\$MS365AppsChannelClear\install.xml" -Force
            }
            Write-Host -ForegroundColor Green "Copy custom XML file as install.xml successfully."
            Write-Output ""
        }
        Write-Host "Save Settings"
    })

    # Image Logo
    $WPFImage_Logo.Add_MouseLeftButtonUp({
        [system.Diagnostics.Process]::start('https://www.deyda.net')
    })

    $WPFImage_Logo_Detail.Add_MouseLeftButtonUp({
        [system.Diagnostics.Process]::start('https://www.deyda.net')
    })

    # Shows the form
    $Form.ShowDialog() | out-null
}

#===========================================================================

Write-Host -Foregroundcolor DarkGray "Software selection"

#// MARK: Define and reset variables
$Date = $Date = Get-Date -UFormat "%m.%d.%Y"
#$Script:Install = $Install
#$Script:Download = $Download

# Define the variables for the unattended install or download (Parameter -file) (AddScript)
If ($file) {
    # Read File Parameter to get the settings. (AddScript)
    If (Test-Path "$File" -PathType leaf) {
        $FileSetting = Get-Content "$file"
        $Language = $FileSetting[0] -as [int]
        $Architecture = $FileSetting[1] -as [int]
        $CitrixWorkspaceAppRelease = $FileSetting[2] -as [int]
        $MS365AppsChannel = $FileSetting[3] -as [int]
        $MSOneDriveRing = $FileSetting[4] -as [int]
        $MSTeamsRing = $FileSetting[5] -as [int]
        $FirefoxChannel = $FileSetting[6] -as [int]
        $TreeSizeType = $FileSetting[7] -as [int]
        $MSDotNetFrameworkChannel = $FileSetting[40] -as [int]
        $MSPowerShellRelease = $FileSetting[42] -as [int]
        $RemoteDesktopManagerType = $FileSetting[44] -as [int]
        $ZoomCitrixClient = $FileSetting[49] -as [int]
        $deviceTRUSTPackage = $FileSetting[50] -as [int]
        $MSEdgeChannel = $FileSetting[51] -as [int]
        $MSVisualStudioCodeChannel = $FileSetting[56] -as [int]
        $Installer = $FileSetting[60] -as [int]
        $MSVisualStudioEdition= $FileSetting[61] -as [int]
        $PuttyChannel = $FileSetting[62] -as [int]
        $MSAzureDataStudioChannel = $FileSetting[64] -as [int]
        $7ZIP = $FileSetting[8] -as [int]
        $AdobeProDC = $FileSetting[9] -as [int]
        $AdobeReaderDC = $FileSetting[10] -as [int]
        $BISF = $FileSetting[11] -as [int]
        $Citrix_Hypervisor_Tools = $FileSetting[12] -as [int]
        $Citrix_WorkspaceApp = $FileSetting[13] -as [int]
        $Filezilla = $FileSetting[14] -as [int]
        $Firefox = $FileSetting[15] -as [int]
        $Foxit_Reader = $FileSetting[16] -as [int]
        $MSFSLogix = $FileSetting[17] -as [int]
        $GoogleChrome = $FileSetting[18] -as [int]
        $Greenshot = $FileSetting[19] -as [int]
        $KeePass = $FileSetting[20] -as [int]
        $mRemoteNG = $FileSetting[21] -as [int]
        $MS365Apps = $FileSetting[22] -as [int]
        $MSEdge = $FileSetting[23] -as [int]
        $MSOffice = $FileSetting[24] -as [int]
        $MSOneDrive = $FileSetting[25] -as [int]
        $MSTeams = $FileSetting[26] -as [int]
        $NotePadPlusPlus = $FileSetting[27] -as [int]
        $OpenJDK = $FileSetting[28] -as [int]
        $OracleJava8 = $FileSetting[29] -as [int]
        $TreeSize = $FileSetting[30] -as [int]
        $VLCPlayer = $FileSetting[31] -as [int]
        $VMWareTools = $FileSetting[32] -as [int]
        $WinSCP = $FileSetting[33] -as [int]
        $Download = $FileSetting[34] -as [int]
        If ($FileSetting[34] -eq "True") {$Download = "1"}
        $Install = $FileSetting[35] -as [int]
        If ($FileSetting[35] -eq "True") {$Install = "1"}
        $IrfanView = $FileSetting[36] -as [int]
        $MSTeamsNoAutoStart = $FileSetting[37] -as [int]
        $deviceTRUST = $FileSetting[38] -as [int]
        $MSDotNetFramework = $FileSetting[39] -as [int]
        $MSPowerShell = $FileSetting[41] -as [int]
        $RemoteDesktopManager = $FileSetting[43] -as [int]
        $Slack = $FileSetting[45] -as [int]
        $Wireshark = $FileSetting[46] -as [int]
        $ShareX = $FileSetting[47] -as [int]
        $Zoom = $FileSetting[48] -as [int]
        $GIMP = $FileSetting[52] -as [int]
        $MSPowerToys = $FileSetting[53] -as [int]
        $MSVisualStudio = $FileSetting[54] -as [int]
        $MSVisualStudioCode = $FileSetting[55] -as [int]
        $PaintDotNet = $FileSetting[57] -as [int]
        $Putty = $FileSetting[58] -as [int]
        $TeamViewer = $FileSetting[59] -as [int]
        $MSAzureDataStudio = $FileSetting[63] -as [int]
        $ImageGlass = $FileSetting[65] -as [int]
        $MSFSLogixChannel = $FileSetting[66] -as [int]
        $uberAgent = $FileSetting[67] -as [int]
        $1Password = $FileSetting[68] -as [int]
        $SumatraPDF = $FileSetting[69] -as [int]
        $ControlUpAgent = $FileSetting[70] -as [int]
        $ControlUpAgentFramework = $FileSetting[71] -as [int]
        $ControlUpConsole = $FileSetting[72] -as [int]
        $MSSQLServerManagementStudio = $FileSetting[73] -as [int]
        $MSAVDRemoteDesktop = $FileSetting[74] -as [int]
        $MSAVDRemoteDesktopChannel = $FileSetting[75] -as [int]
        $MSPowerBIDesktop = $FileSetting[76] -as [int]
        $RDAnalyzer = $FileSetting[77] -as [int]
        $CiscoWebexTeams = $FileSetting[78] -as [int]
        $CitrixFiles = $FileSetting[79] -as [int]
        $FoxitPDFEditor = $FileSetting[80] -as [int]
        $GitForWindows = $FileSetting[81] -as [int]
        $LogMeInGoToMeeting = $FileSetting[82] -as [int]
        $MSAzureCLI = $FileSetting[83] -as [int]
        $MSPowerBIReportBuilder = $FileSetting[84] -as [int]
        $MSSysinternals = $FileSetting[85] -as [int]
        $NMap = $FileSetting[86] -as [int]
        $PeaZip = $FileSetting[87] -as [int]
        $TechSmithCamtasia = $FileSetting[88] -as [int]
        $TechSmithSnagit = $FileSetting[89] -as [int]
        $WinMerge = $FileSetting[90] -as [int]
        $WhatIf = $FileSetting[91] -as [int]
        $CleanUp = $FileSetting[92] -as [int]
        $7Zip_Architecture = $FileSetting[93] -as [int]
        $AdobeReaderDC_Architecture = $FileSetting[94] -as [int]
        $AdobeReaderDC_Language = $FileSetting[95] -as [int]
        $CiscoWebexTeams_Architecture = $FileSetting[96] -as [int]
        $CitrixHypervisorTools_Architecture = $FileSetting[97] -as [int]
        $ControlUpAgent_Architecture = $FileSetting[98] -as [int]
        $deviceTRUST_Architecture = $FileSetting[99] -as [int]
        $FoxitPDFEditor_Language = $FileSetting[100] -as [int]
        $FoxitReader_Language = $FileSetting[101] -as [int]
        $GitForWindows_Architecture = $FileSetting[102] -as [int]
        $GoogleChrome_Architecture = $FileSetting[103] -as [int]
        $ImageGlass_Architecture = $FileSetting[104] -as [int]
        $IrfanView_Architecture = $FileSetting[105] -as [int]
        $KeePass_Language = $FileSetting[106] -as [int]
        $MSDotNetFramework_Architecture = $FileSetting[107] -as [int]
        $MS365Apps_Architecture = $FileSetting[108] -as [int]
        $MS365Apps_Language = $FileSetting[109] -as [int]
        $MS365Apps_Visio = $FileSetting[110] -as [int]
        $MS365Apps_Visio_Language = $FileSetting[111] -as [int]
        $MS365Apps_Project = $FileSetting[112] -as [int]
        $MS365Apps_Project_Language = $FileSetting[113] -as [int]
        $MSAVDRemoteDesktop_Architecture = $FileSetting[114] -as [int]
        $MSEdge_Architecture = $FileSetting[115] -as [int]
        $MSFSLogix_Architecture = $FileSetting[116] -as [int]
        $MSOffice_Architecture = $FileSetting[117] -as [int]
        $MSOneDrive_Architecture = $FileSetting[118] -as [int]
        $MSPowerBIDesktop_Architecture = $FileSetting[119] -as [int]
        $MSPowerShell_Architecture = $FileSetting[120] -as [int]
        $MSSQLServerManagementStudio_Language = $FileSetting[121] -as [int]
        $MSTeams_Architecture = $FileSetting[122] -as [int]
        $MSVisualStudioCode_Architecture = $FileSetting[123] -as [int]
        $Firefox_Architecture = $FileSetting[124] -as [int]
        $Firefox_Language = $FileSetting[125] -as [int]
        $NotepadPlusPlus_Architecture = $FileSetting[126] -as [int]
        $OpenJDK_Architecture = $FileSetting[127] -as [int]
        $OracleJava8_Architecture = $FileSetting[128] -as [int]
        $PeaZip_Architecture = $FileSetting[129] -as [int]
        $Putty_Architecture = $FileSetting[130] -as [int]
        $Slack_Architecture = $FileSetting[131] -as [int]
        $SumatraPDF_Architecture = $FileSetting[132] -as [int]
        $TechSmithSnagIt_Architecture = $FileSetting[133] -as [int]
        $VLCPlayer_Architecture = $FileSetting[134] -as [int]
        $VMWareTools_Architecture = $FileSetting[135] -as [int]
        $WinMerge_Architecture = $FileSetting[136] -as [int]
        $Wireshark_Architecture = $FileSetting[137] -as [int]
        $IrfanView_Language = $FileSetting[138] -as [int]
        $MSOffice_Language = $FileSetting[139] -as [int]
        $MSEdgeWebView2 = $FileSetting[140] -as [int]
        $MSEdgeWebView2_Architecture = $FileSetting[141] -as [int]
        $AutodeskDWGTrueView = $FileSetting[142] -as [int]
        $MindView7 = $FileSetting[143] -as [int]
        $MindView7_Language = $FileSetting[144] -as [int]
        $PDFsam = $FileSetting[145] -as [int]
        $MSOfficeVersion = $FileSetting[146] -as [int]
        $OpenShellMenu = $FileSetting[147] -as [int]
        $PDFForgeCreator = $FileSetting[148] -as [int]
        $TotalCommander = $FileSetting[149] -as [int]
        $LogMeInGoToMeeting_Installer = $FileSetting[150] -as [int]
        $MSAzureDataStudio_Installer = $FileSetting[151] -as [int]
        $MSVisualStudioCode_Installer = $FileSetting[152] -as [int]
        $MS365Apps_Installer = $FileSetting[153] -as [int]
        $MSOffice_Installer = $FileSetting[154] -as [int]
        $MSTeams_Installer = $FileSetting[155] -as [int]
        $Zoom_Installer = $FileSetting[156] -as [int]
        $MSOneDrive_Installer = $FileSetting[157] -as [int]
        $Slack_Installer = $FileSetting[158] -as [int]
        $pdfforgePDFCreatorChannel = $FileSetting[159] -as [int]
        $TotalCommander_Architecture = $FileSetting[160] -as [int]
        $Repository = $FileSetting[161] -as [int]
        $CleanUpStartMenu = $FileSetting[162] -as [int]
    }
    Write-Host "Unattended Mode."
}
Else {
    # Cleanup of the used variables (AddScript)
    Clear-Variable -name Download,Install,7ZIP,AdobeProDC,AdobeReaderDC,BISF,Citrix_Hypervisor_Tools,Filezilla,Firefox,Foxit_Reader,MSFSLogix,Greenshot,GoogleChrome,KeePass,mRemoteNG,MS365Apps,MSEdge,MSOffice,MSTeams,NotePadPlusPlus,MSOneDrive,OpenJDK,OracleJava8,TreeSize,VLCPlayer,VMWareTools,WinSCP,Citrix_WorkspaceApp,Architecture,FirefoxChannel,CitrixWorkspaceAppRelease,Language,MS365AppsChannel,MSOneDriveRing,MSTeamsRing,TreeSizeType,IrfanView,MSTeamsNoAutoStart,deviceTRUST,MSDotNetFramework,MSDotNetFrameworkChannel,MSPowerShell,MSPowerShellRelease,RemoteDesktopManager,RemoteDesktopManagerType,Slack,ShareX,Zoom,ZoomCitrixClient,deviceTRUSTPackage,deviceTRUSTClient,deviceTRUSTConsole,deviceTRUSTHost,MSEdgeChannel,Installer,MSVisualStudioCodeChannel,MSVisualStudio,MSVisualStudioCode,TeamViewer,Putty,PaintDotNet,MSPowerToys,GIMP,MSVisualStudioEdition,PuttyChannel,Wireshark,MSAzureDataStudio,MSAzureDataStudioChannel,ImageGlass,MSFSLogixChannel,uberAgent,1Password,CiscoWebexClient,ControlUpAgent,ControlUpAgentFramework,ControlUpConsole,MSSQLServerManagementStudio,MSAVDRemoteDesktop,MSAVDRemoteDesktopChannel,MSPowerBIDesktop,RDAnalyzer,SumatraPDF,CiscoWebexTeams,CitrixFiles,FoxitPDFEditor,GitForWindows,LogMeInGoToMeeting,MSAzureCLI,MSPowerBIReportBuilder,MSSysinternals,NMap,PeaZip,TechSmithCamtasia,TechSmithSnagit,WinMerge,WhatIf,CleanUp,7Zip_Architecture,AdobeReaderDC_Architecture,AdobeReaderDC_Language,CiscoWebexTeams_Architecture,CitrixHypervisorTools_Architecture,ControlUpAgent_Architecture,deviceTRUST_Architecture,FoxitPDFEditor_Language,FoxitReader_Language,GitForWindows_Architecture,GoogleChrome_Architecture,ImageGlass_Architecture,IrfanView_Architecture,Keepass_Language,MSDotNetFramework_Architecture,MS365Apps_Architecture,MS365Apps_Language,MS365Apps_Visio,MS365Apps_Visio_Language,MS365Apps_Project,MS365Apps_Project_Language,MSAVDRemoteDesktop_Architecture,MSEdge_Architecture,MSFSLogix_Architecture,MSOffice_Architecture,MSOneDrive_Architecture,MSPowerBIDesktop_Architecture,MSPowerShell_Architecture,MSSQLServerManagementStudio_Language,MSTeams_Architecture,MSVisualStudioCode_Architecture,Firefox_Architecture,Firefox_Language,NotePadPlusPlus_Architecture,OpenJDK_Architecture,OracleJava8_Architecture,PeaZip_Architecture,Putty_Architecture,Slack_Architecture,SumatraPDF_Architecture,TechSmithSnagIt_Architecture,VLCPlayer_Architecture,VMWareTools_Architecture,WinMerge_Architecture,Wireshark_Architecture,IrfanView_Language,MSOffice_Language,MSEdgeWebView2,MSEdgeWebView2_Architecture,AutodeskDWGTrueView,MindView7,MindView7_Language,PDFsam,MSOfficeVersion,OpenShellMenu,PDFForgeCreator,TotalCommander,LogMeInGoToMeeting_Installer,MSAzureDataStudio_Installer,MSVisualStudioCode_Installer,MS365Apps_Installer,MSOffice_Installer,MSTeams_Installer,Zoom_Installer,MSOneDrive_Installer,Slack_Installer,pdfforgePDFCreatorChannel,TotalCommander_Architecture,Repository,CleanUpStartMenu -ErrorAction SilentlyContinue

    # Shortcut Creation
    If (!(Test-Path -Path "$env:USERPROFILE\Desktop\Evergreen Script.lnk")) {
        $WScriptShell = New-Object -ComObject 'WScript.Shell'
        $ShortcutFile = "$env:USERPROFILE\Desktop\Evergreen Script.lnk"
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Shortcut.WorkingDirectory = "C:\Windows\System32\WindowsPowerShell\v1.0"
        If (!(Test-Path -Path "$PSScriptRoot\shortcut")) { New-Item -Path "$PSScriptRoot\shortcut" -ItemType Directory | Out-Null }
        If (!(Test-Path -Path "$PSScriptRoot\shortcut\EvergreenLeafDeyda.ico")) {Invoke-WebRequest -Uri https://raw.githubusercontent.com/Deyda/Evergreen-Script/main/shortcut/EvergreenLeafDeyda.ico -OutFile ("$PSScriptRoot\shortcut\" + "EvergreenLeafDeyda.ico")}
        $shortcut.IconLocation="$PSScriptRoot\shortcut\EvergreenLeafDeyda.ico"
        $Shortcut.Arguments = '-noexit -ExecutionPolicy Bypass -file "' + "$PSScriptRoot" + '\Evergreen.ps1"'
        $Shortcut.Save()
        $Admin = [System.IO.File]::ReadAllBytes("$ShortcutFile")
        $Admin[0x15] = $Admin[0x15] -bor 0x20
        [System.IO.File]::WriteAllBytes("$ShortcutFile", $Admin)
    }
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

Switch ($Installer) {
    0 { $InstallerClear = 'Machine Based'}
    1 { $InstallerClear = 'User Based'}
}

If ($7Zip_Architecture -ne "") {
    Switch ($7Zip_Architecture) {
        1 { $7ZipArchitectureClear = 'x86'}
        2 { $7ZipArchitectureClear = 'x64'}
    }
}
Else {
    $7ZipArchitectureClear = $ArchitectureClear
}

If ($AdobeReaderDC_Architecture -ne "") {
    Switch ($AdobeReaderDC_Architecture) {
        1 { $AdobeArchitectureClear = 'x86'}
        2 { $AdobeArchitectureClear = 'x64'}
    }
}
Else {
    $AdobeArchitectureClear = $ArchitectureClear
}

If ($AdobeReaderDC_Language -ne "") {
    Switch ($AdobeReaderDC_Language) {
        1 { $AdobeLanguageClear = 'Danish'}
        2 { $AdobeLanguageClear = 'Dutch'}
        3 { $AdobeLanguageClear = 'English'}
        4 { $AdobeLanguageClear = 'Finnish'}
        5 { $AdobeLanguageClear = 'French'}
        6 { $AdobeLanguageClear = 'German'}
        7 { $AdobeLanguageClear = 'Italian'}
        8 { $AdobeLanguageClear = 'Japanese'}
        9 { $AdobeLanguageClear = 'Korean'}
        10 { $AdobeLanguageClear = 'Norwegian'}
        11 { $AdobeLanguageClear = 'Polish'}
        12 { $AdobeLanguageClear = 'Russian'}
        13 { $AdobeLanguageClear = 'Spanish'}
        14 { $AdobeLanguageClear = 'Swedish'}
    }
}
Else {
    $AdobeLanguageClear = $LanguageClear
    Switch ($LanguageClear) {
        Portuguese { $AdobeLanguageClear = 'English'}
    }
}

If ($CiscoWebexTeams_Architecture -ne "") {
    Switch ($CiscoWebexTeams_Architecture) {
        1 { $CiscoWebexTeamsArchitectureClear = 'x86'}
        2 { $CiscoWebexTeamsArchitectureClear = 'x64'}
    }
}
Else {
    $CiscoWebexTeamsArchitectureClear = $ArchitectureClear
}

If ($CitrixHypervisorTools_Architecture -ne "") {
    Switch ($CitrixHypervisorTools_Architecture) {
        1 { $CitrixHypervisorToolsArchitectureClear = 'x86'}
        2 { $CitrixHypervisorToolsArchitectureClear = 'x64'}
    }
}
Else {
    $CitrixHypervisorToolsArchitectureClear = $ArchitectureClear
}

Switch ($CitrixWorkspaceAppRelease) {
    0 { $CitrixWorkspaceAppReleaseClear = 'Current'}
    1 { $CitrixWorkspaceAppReleaseClear = 'LTSR'}
}

If ($ControlUpAgent_Architecture -ne "") {
    Switch ($ControlUpAgent_Architecture) {
        1 { $ControlUpAgentArchitectureClear = 'x86'}
        2 { $ControlUpAgentArchitectureClear = 'x64'}
    }
}
Else {
    $ControlUpAgentArchitectureClear = $ArchitectureClear
}

Switch ($ControlUpAgentFramework) {
    0 { $ControlUpAgentFrameworkClear = 'net35'}
    1 { $ControlUpAgentFrameworkClear = 'net45'}
}

If ($deviceTRUST_Architecture -ne "") {
    Switch ($deviceTRUST_Architecture) {
        1 { $deviceTRUSTArchitectureClear = 'x86'}
        2 { $deviceTRUSTArchitectureClear = 'x64'}
    }
}
Else {
    $deviceTRUSTArchitectureClear = $ArchitectureClear
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

If ($FoxitPDFEditor_Language -ne "") {
    Switch ($FoxitPDFEditor_Language) {
        1 { $FoxitPDFEditorLanguageClear = 'Danish'}
        2 { $FoxitPDFEditorLanguageClear = 'Dutch'}
        3 { $FoxitPDFEditorLanguageClear = 'English'}
        4 { $FoxitPDFEditorLanguageClear = 'Finnish'}
        5 { $FoxitPDFEditorLanguageClear = 'French'}
        6 { $FoxitPDFEditorLanguageClear = 'German'}
        7 { $FoxitPDFEditorLanguageClear = 'Italian'}
        8 { $FoxitPDFEditorLanguageClear = 'Korean'}
        9 { $FoxitPDFEditorLanguageClear = 'Norwegian'}
        10 { $FoxitPDFEditorLanguageClear = 'Polish'}
        11 { $FoxitPDFEditorLanguageClear = 'Portuguese'}
        12 { $FoxitPDFEditorLanguageClear = 'Russian'}
        13 { $FoxitPDFEditorLanguageClear = 'Spanish'}
        14 { $FoxitPDFEditorLanguageClear = 'Swedish'}
    }
}
Else {
    $FoxitPDFEditorLanguageClear = $LanguageClear
    Switch ($LanguageClear) {
        Japanese { $FoxitPDFEditorLanguageClear = 'English'}
    }
}

If ($FoxitReader_Language -ne "") {
    Switch ($FoxitReader_Language) {
        1 { $FoxitReaderLanguageClear = 'Danish'}
        2 { $FoxitReaderLanguageClear = 'Dutch'}
        3 { $FoxitReaderLanguageClear = 'English'}
        4 { $FoxitReaderLanguageClear = 'Finnish'}
        5 { $FoxitReaderLanguageClear = 'French'}
        6 { $FoxitReaderLanguageClear = 'German'}
        7 { $FoxitReaderLanguageClear = 'Italian'}
        8 { $FoxitReaderLanguageClear = 'Norwegian'}
        9 { $FoxitReaderLanguageClear = 'Polish'}
        10 { $FoxitReaderLanguageClear = 'Portuguese'}
        11 { $FoxitReaderLanguageClear = 'Russian'}
        12 { $FoxitReaderLanguageClear = 'Spanish'}
        13 { $FoxitReaderLanguageClear = 'Swedish'}
    }
}
Else {
    $FoxitReaderLanguageClear = $LanguageClear
    Switch ($LanguageClear) {
        Japanese { $FoxitReaderLanguageClear = 'English'}
        Korean { $FoxitReaderLanguageClear = 'English'}
    }
}

If ($GitForWindows_Architecture -ne "") {
    Switch ($GitForWindows_Architecture) {
        1 { $GitForWindowsArchitectureClear = 'x86'}
        2 { $GitForWindowsArchitectureClear = 'x64'}
    }
}
Else {
    $GitForWindowsArchitectureClear = $ArchitectureClear
}

If ($GoogleChrome_Architecture -ne "") {
    Switch ($GoogleChrome_Architecture) {
        1 { $GoogleChromeArchitectureClear = 'x86'}
        2 { $GoogleChromeArchitectureClear = 'x64'}
    }
}
Else {
    $GoogleChromeArchitectureClear = $ArchitectureClear
}

If ($ImageGlass_Architecture -ne "") {
    Switch ($ImageGlass_Architecture) {
        1 { $ImageGlassArchitectureClear = 'x86'}
        2 { $ImageGlassArchitectureClear = 'x64'}
    }
}
Else {
    $ImageGlassArchitectureClear = $ArchitectureClear
}

If ($IrfanView_Architecture -ne "") {
    Switch ($IrfanView_Architecture) {
        1 { $IrfanViewArchitectureClear = 'x86'}
        2 { $IrfanViewArchitectureClear = 'x64'}
    }
}
Else {
    $IrfanViewArchitectureClear = $ArchitectureClear
}

If ($IrfanView_Language -ne "") {
    Switch ($IrfanView_Language) {
        1 { $IrfanViewLanguageClear = 'da'}
        2 { $IrfanViewLanguageClear = 'nl-NL'}
        4 { $IrfanViewLanguageClear = 'fi'}
        5 { $IrfanViewLanguageClear = 'fr'}
        6 { $IrfanViewLanguageClear = 'de'}
        7 { $IrfanViewLanguageClear = 'it'}
        8 { $IrfanViewLanguageClear = 'ja'}
        9 { $IrfanViewLanguageClear = 'ko'}
        10 { $IrfanViewLanguageClear = 'pl'}
        11 { $IrfanViewLanguageClear = 'pt-PT'}
        12 { $IrfanViewLanguageClear = 'ru'}
        13 { $IrfanViewLanguageClear = 'es'}
        14 { $IrfanViewLanguageClear = 'sv'}
    }
    Switch ($IrfanView_Language) {
        1 { $IrfanViewLanguageLongClear = 'Danish'}
        2 { $IrfanViewLanguageLongClear = 'Dutch'}
        3 { $IrfanViewLanguageLongClear = 'English'}
        4 { $IrfanViewLanguageLongClear = 'Finnish'}
        5 { $IrfanViewLanguageLongClear = 'French'}
        6 { $IrfanViewLanguageLongClear = 'German'}
        7 { $IrfanViewLanguageLongClear = 'Italian'}
        8 { $IrfanViewLanguageLongClear = 'Japanese'}
        9 { $IrfanViewLanguageLongClear = 'Korean'}
        10 { $IrfanViewLanguageLongClear = 'Polish'}
        11 { $IrfanViewLanguageLongClear = 'Portuguese'}
        12 { $IrfanViewLanguageLongClear = 'Russian'}
        13 { $IrfanViewLanguageLongClear = 'Spanish'}
        14 { $IrfanViewLanguageLongClear = 'Swedish'}
    }
}
Else {
    Switch ($Language) {
        0 { $IrfanViewLanguageClear = 'da'}
        1 { $IrfanViewLanguageClear = 'nl-NL'}
        3 { $IrfanViewLanguageClear = 'fi'}
        4 { $IrfanViewLanguageClear = 'fr'}
        5 { $IrfanViewLanguageClear = 'de'}
        6 { $IrfanViewLanguageClear = 'it'}
        7 { $IrfanViewLanguageClear = 'ja'}
        8 { $IrfanViewLanguageClear = 'ko'}
        10 { $IrfanViewLanguageClear = 'pl'}
        11 { $IrfanViewLanguageClear = 'pt-PT'}
        12 { $IrfanViewLanguageClear = 'ru'}
        13 { $IrfanViewLanguageClear = 'es'}
        14 { $IrfanViewLanguageClear = 'sv'}
    }
    Switch ($Language) {
        0 { $IrfanViewLanguageLongClear = 'Danish'}
        1 { $IrfanViewLanguageLongClear = 'Dutch'}
        2 { $IrfanViewLanguageLongClear = 'English'}
        3 { $IrfanViewLanguageLongClear = 'Finnish'}
        4 { $IrfanViewLanguageLongClear = 'French'}
        5 { $IrfanViewLanguageLongClear = 'German'}
        6 { $IrfanViewLanguageLongClear = 'Italian'}
        7 { $IrfanViewLanguageLongClear = 'Japanese'}
        8 { $IrfanViewLanguageLongClear = 'Korean'}
        9 { $IrfanViewLanguageLongClear = 'English'}
        10 { $IrfanViewLanguageLongClear = 'Polish'}
        11 { $IrfanViewLanguageLongClear = 'Portuguese'}
        12 { $IrfanViewLanguageLongClear = 'Russian'}
        13 { $IrfanViewLanguageLongClear = 'Spanish'}
        14 { $IrfanViewLanguageLongClear = 'Swedish'}
    }
}

If ($KeePass_Language -ne "") {
    Switch ($KeePass_Language) {
        1 { $KeePassLanguageClear = 'Danish'}
        2 { $KeePassLanguageClear = 'Dutch'}
        3 { $KeePassLanguageClear = 'English'}
        4 { $KeePassLanguageClear = 'Finnish'}
        5 { $KeePassLanguageClear = 'French'}
        6 { $KeePassLanguageClear = 'German'}
        7 { $KeePassLanguageClear = 'Italian'}
        8 { $KeePassLanguageClear = 'Japanese'}
        9 { $KeePassLanguageClear = 'Korean'}
        10 { $KeePassLanguageClear = 'Norwegian'}
        11 { $KeePassLanguageClear = 'Polish'}
        12 { $KeePassLanguageClear = 'Portuguese'}
        13 { $KeePassLanguageClear = 'Russian'}
        14 { $KeePassLanguageClear = 'Spanish'}
        15 { $KeePassLanguageClear = 'Swedish'}
    }
}
Else {
    $KeePassLanguageClear = $LanguageClear
}

If ($LogMeInGoToMeeting_Installer -ne "") {
    Switch ($LogMeInGoToMeeting_Installer) {
        1 { $LogMeInGoToMeetingInstallerClear = 'Machine Based'}
        2 { $LogMeInGoToMeetingInstallerClear = 'User Based'}
    }
}
Else {
    $LogMeInGoToMeetingInstallerClear = $InstallerClear
}

If ($MSDotNetFramework_Architecture -ne "") {
    Switch ($MSDotNetFramework_Architecture) {
        1 { $MSDotNetFrameworkArchitectureClear = 'x86'}
        2 { $MSDotNetFrameworkArchitectureClear = 'x64'}
    }
}
Else {
    $MSDotNetFrameworkArchitectureClear = $ArchitectureClear
}

Switch ($MSDotNetFrameworkChannel) {
    0 { $MSDotNetFrameworkChannelClear = 'Current'}
    1 { $MSDotNetFrameworkChannelClear = 'LTS'}
}

If ($MS365Apps_Architecture -ne "") {
    Switch ($MS365Apps_Architecture) {
        1 { 
            $MS365AppsArchitectureClear = '32'
            $MS365AppsArchitecturePolicyClear = 'x86'
        }
        2 {
            $MS365AppsArchitectureClear = '64'
            $MS365AppsArchitecturePolicyClear = 'x64'
        }
    }
}
Else {
    Switch ($Architecture) {
        0 { 
            $MS365AppsArchitectureClear = '64'
            $MS365AppsArchitecturePolicyClear = 'x64'
        }
        1 { 
            $MS365AppsArchitectureClear = '32'
            $MS365AppsArchitecturePolicyClear = 'x86'
        }
    }
}

If ($MS365Apps_Language -ne "") {
    Switch ($MS365Apps_Language) {
        1 { $MS365AppsLanguageClear = 'da-DK'}
        2 { $MS365AppsLanguageClear = 'nl-NL'}
        3 { $MS365AppsLanguageClear = 'en-US'}
        4 { $MS365AppsLanguageClear = 'fi-FI'}
        5 { $MS365AppsLanguageClear = 'fr-FR'}
        6 { $MS365AppsLanguageClear = 'de-DE'}
        7 { $MS365AppsLanguageClear = 'it-IT'}
        8 { $MS365AppsLanguageClear = 'ja-JP'}
        9 { $MS365AppsLanguageClear = 'ko-KR'}
        10 { $MS365AppsLanguageClear = 'nb-NO'}
        11 { $MS365AppsLanguageClear = 'pl-PL'}
        12 { $MS365AppsLanguageClear = 'pt-PT'}
        13 { $MS365AppsLanguageClear = 'ru-RU'}
        14 { $MS365AppsLanguageClear = 'es-ES'}
        15 { $MS365AppsLanguageClear = 'sv-SE'}
    }
}
Else {
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
}
If ($MS365Apps_Installer -ne "") {
    Switch ($MS365Apps_Installer) {
        1 { $MS365AppsInstallerClear = 'Machine Based'}
        2 { $MS365AppsInstallerClear = 'User Based'}
    }
}
Else {
    $MS365AppsInstallerClear = $InstallerClear
}

Switch ($MS365AppsChannel) {
    0 { $MS365AppsChannelClear = 'Beta Channel'}
    1 { $MS365AppsChannelClear = 'Current Channel (Preview)'}
    2 { $MS365AppsChannelClear = 'Current Channel'}
    3 { $MS365AppsChannelClear = 'Semi-Annual Channel (Preview)'}
    4 { $MS365AppsChannelClear = 'Monthly Enterprise Channel'}
    5 { $MS365AppsChannelClear = 'Semi-Annual Channel'}
}

Switch ($MS365AppsChannel) {
    0 { $MS365AppsChannelClearDL = 'BetaChannel'}
    1 { $MS365AppsChannelClearDL = 'CurrentPreview'}
    2 { $MS365AppsChannelClearDL = 'Current'}
    3 { $MS365AppsChannelClearDL = 'SemiAnnualPreview'}
    4 { $MS365AppsChannelClearDL = 'MonthlyEnterprise'}
    5 { $MS365AppsChannelClearDL = 'SemiAnnual'}
}

If ($MS365Apps_Project_Language -ne "") {
    Switch ($MS365Apps_Project_Language) {
        1 { $MS365AppsProjectLanguageClear = 'da-DK'}
        2 { $MS365AppsProjectLanguageClear = 'nl-NL'}
        3 { $MS365AppsProjectLanguageClear = 'en-US'}
        4 { $MS365AppsProjectLanguageClear = 'fi-FI'}
        5 { $MS365AppsProjectLanguageClear = 'fr-FR'}
        6 { $MS365AppsProjectLanguageClear = 'de-DE'}
        7 { $MS365AppsProjectLanguageClear = 'it-IT'}
        8 { $MS365AppsProjectLanguageClear = 'ja-JP'}
        9 { $MS365AppsProjectLanguageClear = 'ko-KR'}
        10 { $MS365AppsProjectLanguageClear = 'nb-NO'}
        11 { $MS365AppsProjectLanguageClear = 'pl-PL'}
        12 { $MS365AppsProjectLanguageClear = 'pt-PT'}
        13 { $MS365AppsProjectLanguageClear = 'ru-RU'}
        14 { $MS365AppsProjectLanguageClear = 'es-ES'}
        15 { $MS365AppsProjectLanguageClear = 'sv-SE'}
    }
}
Else {
    Switch ($Language) {
        0 { $MS365AppsProjectLanguageClear = 'da-DK'}
        1 { $MS365AppsProjectLanguageClear = 'nl-NL'}
        2 { $MS365AppsProjectLanguageClear = 'en-US'}
        3 { $MS365AppsProjectLanguageClear = 'fi-FI'}
        4 { $MS365AppsProjectLanguageClear = 'fr-FR'}
        5 { $MS365AppsProjectLanguageClear = 'de-DE'}
        6 { $MS365AppsProjectLanguageClear = 'it-IT'}
        7 { $MS365AppsProjectLanguageClear = 'ja-JP'}
        8 { $MS365AppsProjectLanguageClear = 'ko-KR'}
        9 { $MS365AppsProjectLanguageClear = 'nb-NO'}
        10 { $MS365AppsProjectLanguageClear = 'pl-PL'}
        11 { $MS365AppsProjectLanguageClear = 'pt-PT'}
        12 { $MS365AppsProjectLanguageClear = 'ru-RU'}
        13 { $MS365AppsProjectLanguageClear = 'es-ES'}
        14 { $MS365AppsProjectLanguageClear = 'sv-SE'}
    }
}

If ($MS365Apps_Visio_Language -ne "") {
    Switch ($MS365Apps_Visio_Language) {
        1 { $MS365AppsVisioLanguageClear = 'da-DK'}
        2 { $MS365AppsVisioLanguageClear = 'nl-NL'}
        3 { $MS365AppsVisioLanguageClear = 'en-US'}
        4 { $MS365AppsVisioLanguageClear = 'fi-FI'}
        5 { $MS365AppsVisioLanguageClear = 'fr-FR'}
        6 { $MS365AppsVisioLanguageClear = 'de-DE'}
        7 { $MS365AppsVisioLanguageClear = 'it-IT'}
        8 { $MS365AppsVisioLanguageClear = 'ja-JP'}
        9 { $MS365AppsVisioLanguageClear = 'ko-KR'}
        10 { $MS365AppsVisioLanguageClear = 'nb-NO'}
        11 { $MS365AppsVisioLanguageClear = 'pl-PL'}
        12 { $MS365AppsVisioLanguageClear = 'pt-PT'}
        13 { $MS365AppsVisioLanguageClear = 'ru-RU'}
        14 { $MS365AppsVisioLanguageClear = 'es-ES'}
        15 { $MS365AppsVisioLanguageClear = 'sv-SE'}
    }
}
Else {
    Switch ($Language) {
        0 { $MS365AppsVisioLanguageClear = 'da-DK'}
        1 { $MS365AppsVisioLanguageClear = 'nl-NL'}
        2 { $MS365AppsVisioLanguageClear = 'en-US'}
        3 { $MS365AppsVisioLanguageClear = 'fi-FI'}
        4 { $MS365AppsVisioLanguageClear = 'fr-FR'}
        5 { $MS365AppsVisioLanguageClear = 'de-DE'}
        6 { $MS365AppsVisioLanguageClear = 'it-IT'}
        7 { $MS365AppsVisioLanguageClear = 'ja-JP'}
        8 { $MS365AppsVisioLanguageClear = 'ko-KR'}
        9 { $MS365AppsVisioLanguageClear = 'nb-NO'}
        10 { $MS365AppsVisioLanguageClear = 'pl-PL'}
        11 { $MS365AppsVisioLanguageClear = 'pt-PT'}
        12 { $MS365AppsVisioLanguageClear = 'ru-RU'}
        13 { $MS365AppsVisioLanguageClear = 'es-ES'}
        14 { $MS365AppsVisioLanguageClear = 'sv-SE'}
    }
}

Switch ($MSAzureDataStudioChannel) {
    0 { $MSAzureDataStudioChannelClear = 'Insider'}
    1 { $MSAzureDataStudioChannelClear = 'Stable'}
}

If ($MSAzureDataStudio_Installer -ne "") {
    Switch ($MSAzureDataStudio_Installer) {
        1 { $MSAzureDataStudioInstallerClear = 'Machine Based'}
        2 { $MSAzureDataStudioInstallerClear = 'User Based'}
    }
}
Else {
    $MSAzureDataStudioInstallerClear = $InstallerClear
}

If ($MSAzureDataStudioInstallerClear -eq "Machine Based") {
    Switch ($Architecture) {
        0 { $MSAzureDataStudioPlatformClear = 'win32-x64'}
        1 { $MSAzureDataStudioPlatformClear = 'win32'}
    }
    $MSAzureDataStudioModeClear = 'Per Machine'
}

If ($MSAzureDataStudioInstallerClear -eq "User Based") {
    Switch ($Architecture) {
        0 { $MSAzureDataStudioPlatformClear = 'win32-x64-user'}
        1 { $MSAzureDataStudioPlatformClear = 'win32-user'}
    }
    $MSAzureDataStudioModeClear = 'Per User'
}

If ($MSAVDRemoteDesktop_Architecture -ne "") {
    Switch ($MSAVDRemoteDesktop_Architecture) {
        1 { $MSAVDRemoteDesktopArchitectureClear = 'x86'}
        2 { $MSAVDRemoteDesktopArchitectureClear = 'x64'}
    }
}
Else {
    $MSAVDRemoteDesktopArchitectureClear = $ArchitectureClear
}

Switch ($MSAVDRemoteDesktopChannel) {
    0 { $MSAVDRemoteDesktopChannelClear = 'Insider'}
    1 { $MSAVDRemoteDesktopChannelClear = 'Public'}
}

If ($MSEdge_Architecture -ne "") {
    Switch ($MSEdge_Architecture) {
        1 { $MSEdgeArchitectureClear = 'x86'}
        2 { $MSEdgeArchitectureClear = 'x64'}
    }
}
Else {
    $MSEdgeArchitectureClear = $ArchitectureClear
}

Switch ($MSEdgeChannel) {
    0 { $MSEdgeChannelClear = 'Dev'}
    1 { $MSEdgeChannelClear = 'Beta'}
    2 { $MSEdgeChannelClear = 'Stable'}
}

If ($MSEdgeWebView2_Architecture -ne "") {
    Switch ($MSEdgeWebView2_Architecture) {
        1 { $MSEdgeWebView2ArchitectureClear = 'x86'}
        2 { $MSEdgeWebView2ArchitectureClear = 'x64'}
    }
}
Else {
    $MSEdgeWebView2ArchitectureClear = $ArchitectureClear
}

If ($MSFSLogix_Architecture -ne "") {
    Switch ($MSFSLogix_Architecture) {
        1 { $MSFSLogixArchitectureClear = 'x86'}
        2 { $MSFSLogixArchitectureClear = 'x64'}
    }
}
Else {
    $MSFSLogixArchitectureClear = $ArchitectureClear
}

Switch ($MSFSLogixChannel) {
    0 { $MSFSLogixChannelClear = 'Preview'}
    1 { $MSFSLogixChannelClear = 'Production'}
}

If ($MSOffice_Installer -ne "") {
    Switch ($MSOffice_Installer) {
        1 { $MSOfficeInstallerClear = 'Machine Based'}
        2 { $MSOfficeInstallerClear = 'User Based'}
    }
}
Else {
    $MSOfficeInstallerClear = $InstallerClear
}

If ($MSOffice_Architecture -ne "") {
    Switch ($MSOffice_Architecture) {
        1 { 
            $MSOfficeArchitectureClear = '32'
            $MSOfficeArchitecturePolicyClear = 'x86'
        }
        2 { 
            $MSOfficeArchitectureClear = '64'
            $MSOfficeArchitecturePolicyClear = 'x64'
    }
    }
}
Else {
    Switch ($Architecture) {
        0 { 
            $MSOfficeArchitectureClear = '64'
            $MSOfficeArchitecturePolicyClear = 'x64'
        }
        1 { 
            $MSOfficeArchitectureClear = '32'
            $MSOfficeArchitecturePolicyClear = 'x86'
        }
    }
}

If ($MSOffice_Language -ne "") {
    Switch ($MSOffice_Language) {
        1 { $MSOfficeLanguageClear = 'da-DK'}
        2 { $MSOfficeLanguageClear = 'nl-NL'}
        3 { $MSOfficeLanguageClear = 'en-US'}
        4 { $MSOfficeLanguageClear = 'fi-FI'}
        5 { $MSOfficeLanguageClear = 'fr-FR'}
        6 { $MSOfficeLanguageClear = 'de-DE'}
        7 { $MSOfficeLanguageClear = 'it-IT'}
        8 { $MSOfficeLanguageClear = 'ja-JP'}
        9 { $MSOfficeLanguageClear = 'ko-KR'}
        10 { $MSOfficeLanguageClear = 'nb-NO'}
        11 { $MSOfficeLanguageClear = 'pl-PL'}
        12 { $MSOfficeLanguageClear = 'pt-PT'}
        13 { $MSOfficeLanguageClear = 'ru-RU'}
        14 { $MSOfficeLanguageClear = 'es-ES'}
        15 { $MSOfficeLanguageClear = 'sv-SE'}
    }
}
Else {
    Switch ($Language) {
        0 { $MSOfficeLanguageClear = 'da-DK'}
        1 { $MSOfficeLanguageClear = 'nl-NL'}
        2 { $MSOfficeLanguageClear = 'en-US'}
        3 { $MSOfficeLanguageClear = 'fi-FI'}
        4 { $MSOfficeLanguageClear = 'fr-FR'}
        5 { $MSOfficeLanguageClear = 'de-DE'}
        6 { $MSOfficeLanguageClear = 'it-IT'}
        7 { $MSOfficeLanguageClear = 'ja-JP'}
        8 { $MSOfficeLanguageClear = 'ko-KR'}
        9 { $MSOfficeLanguageClear = 'nb-NO'}
        10 { $MSOfficeLanguageClear = 'pl-PL'}
        11 { $MSOfficeLanguageClear = 'pt-PT'}
        12 { $MSOfficeLanguageClear = 'ru-RU'}
        13 { $MSOfficeLanguageClear = 'es-ES'}
        14 { $MSOfficeLanguageClear = 'sv-SE'}
    }
}

Switch ($MSOfficeVersion) {
    0 {
        $MSOfficeVersionClear = 'PerpetualVL2019'
        $MSOfficeChannelClear = '2019'
        $MSOfficeProductIDClear = 'ProPlus2019Volume'
    }
    1 {
        $MSOfficeVersionClear = 'PerpetualVL2021'
        $MSOfficeChannelClear = '2021 LTSC'
        $MSOfficeProductIDClear = 'ProPlus2021Volume'
    }
}

If ($MSOneDrive_Installer -ne "") {
    Switch ($MSOneDrive_Installer) {
        1 { $MSOneDriveInstallerClear = 'Machine Based'}
        2 { $MSOneDriveInstallerClear = 'User Based'}
    }
}
Else {
    $MSOneDriveInstallerClear = $InstallerClear
}

If ($MSOneDrive_Architecture -ne "") {
    Switch ($MSOneDrive_Architecture) {
        1 { $MSOneDriveArchitectureClear = 'x86'}
        2 { $MSOneDriveArchitectureClear = 'AMD64'}
    }
}
Else {
    Switch ($Architecture) {
        0 { $MSOneDriveArchitectureClear = 'AMD64'}
        1 { $MSOneDriveArchitectureClear = 'x86'}
    }
}

Switch ($MSOneDriveRing) {
    0 { $MSOneDriveRingClear = 'Insider'}
    1 { $MSOneDriveRingClear = 'Production'}
    2 { $MSOneDriveRingClear = 'Enterprise'}
}

If ($MSPowerBIDesktop_Architecture -ne "") {
    Switch ($MSPowerBIDesktop_Architecture) {
        1 { $MSPowerBIDesktopArchitectureClear = 'x86'}
        2 { $MSPowerBIDesktopArchitectureClear = 'x64'}
    }
}
Else {
    $MSPowerBIDesktopArchitectureClear = $ArchitectureClear
}

If ($MSPowerShell_Architecture -ne "") {
    Switch ($MSPowerShell_Architecture) {
        1 { $MSPowerShellArchitectureClear = 'x86'}
        2 { $MSPowerShellArchitectureClear = 'x64'}
    }
}
Else {
    $MSPowerShellArchitectureClear = $ArchitectureClear
}

Switch ($MSPowerShellRelease) {
    0 { $MSPowerShellReleaseClear = 'Stable'}
    1 { $MSPowerShellReleaseClear = 'LTS'}
}


If ($MSSQLServerManagementStudio_Language -ne "") {
    Switch ($MSSQLServerManagementStudio_Language) {
        1 { $MSSQLServerManagementStudioLanguageClear = 'English'}
        2 { $MSSQLServerManagementStudioLanguageClear = 'French'}
        3 { $MSSQLServerManagementStudioLanguageClear = 'German'}
        4 { $MSSQLServerManagementStudioLanguageClear = 'Italian'}
        5 { $MSSQLServerManagementStudioLanguageClear = 'Japanese'}
        6 { $MSSQLServerManagementStudioLanguageClear = 'Korean'}
        7 { $MSSQLServerManagementStudioLanguageClear = 'Portuguese (Brazil)'}
        8 { $MSSQLServerManagementStudioLanguageClear = 'Russian'}
        9 { $MSSQLServerManagementStudioLanguageClear = 'Spanish'}
    }
}
Else {
    $MSSQLServerManagementStudioLanguageClear = $LanguageClear
    Switch ($LanguageClear) {
        Danish { $MSSQLServerManagementStudioLanguageClear = 'English'}
        Dutch { $MSSQLServerManagementStudioLanguageClear = 'English'}
        Finnish { $MSSQLServerManagementStudioLanguageClear = 'English'}
        Norwegian { $MSSQLServerManagementStudioLanguageClear = 'English'}
        Polish { $MSSQLServerManagementStudioLanguageClear = 'English'}
        Portuguese { $MSSQLServerManagementStudioLanguageClear = 'Portuguese (Brazil)'}
        Swedish { $MSSQLServerManagementStudioLanguageClear = 'English'}
    }
}

If ($MSTeams_Installer -ne "") {
    Switch ($MSTeams_Installer) {
        1 { $MSTeamsInstallerClear = 'Machine Based'}
        2 { $MSTeamsInstallerClear = 'User Based'}
    }
}
Else {
    $MSTeamsInstallerClear = $InstallerClear
}

If ($MSTeams_Architecture -ne "") {
    Switch ($MSTeams_Architecture) {
        1 { $MSTeamsArchitectureClear = 'x86'}
        2 { $MSTeamsArchitectureClear = 'x64'}
    }
}
Else {
    $MSTeamsArchitectureClear = $ArchitectureClear
}

Switch ($MSTeamsRing) {
    0 { $MSTeamsRingClear = 'Continuous Deployment'}
    1 { $MSTeamsRingClear = 'Exploration'}
    2 { $MSTeamsRingClear = 'Preview'}
    3 { $MSTeamsRingClear = 'General'}
}

Switch ($MSVisualStudioEdition) {
    0 { $MSVisualStudioEditionClear = 'Enterprise'}
    1 { $MSVisualStudioEditionClear = 'Professional'}
    2 { $MSVisualStudioEditionClear = 'Community'}
}

If ($MSVisualStudioCode_Architecture -ne "") {
    Switch ($MSVisualStudioCode_Architecture) {
        1 { $MSVisualStudioCodeArchitectureClear = 'x86'}
        2 { $MSVisualStudioCodeArchitectureClear = 'x64'}
    }
}
Else {
    $MSVisualStudioCodeArchitectureClear = $ArchitectureClear
}

If ($MSVisualStudioCode_Installer -ne "") {
    Switch ($MSVisualStudioCode_Installer) {
        1 { $MSVisualStudioCodeInstallerClear = 'Machine Based'}
        2 { $MSVisualStudioCodeInstallerClear = 'User Based'}
    }
}
Else {
    $MSVisualStudioCodeInstallerClear = $InstallerClear
}

If ($MSVisualStudioCodeInstallerClear -eq "Machine Based") {
    If ($MSVisualStudioCode_Architecture -ne "") {
        Switch ($MSVisualStudioCode_Architecture) {
            1 { $MSVisualStudioCodePlatformClear = 'win32'}
            2 { $MSVisualStudioCodePlatformClear = 'win32-x64'}
        }
    }
    Else {
        Switch ($Architecture) {
            0 { $MSVisualStudioCodePlatformClear = 'win32-x64'}
            1 { $MSVisualStudioCodePlatformClear = 'win32'}
        }
    }
    $MSVisualStudioCodeModeClear = 'Per Machine'
}

If ($MSVisualStudioCodeInstallerClear -eq "User Based") {
    If ($MSVisualStudioCode_Architecture -ne "") {
        Switch ($MSVisualStudioCode_Architecture) {
            1 { $MSVisualStudioCodePlatformClear = 'win32-user'}
            2 { $MSVisualStudioCodePlatformClear = 'win32-x64-user'}
        }
    }
    Else {
        Switch ($Architecture) {
            0 { $MSVisualStudioCodePlatformClear = 'win32-x64-user'}
            1 { $MSVisualStudioCodePlatformClear = 'win32-user'}
        }
    }
    $MSVisualStudioCodeModeClear = 'Per User'
}

Switch ($MSVisualStudioCodeChannel) {
    0 { $MSVisualStudioCodeChannelClear = 'Insider'}
    1 { $MSVisualStudioCodeChannelClear = 'Stable'}
}

If ($MindView7_Language -ne "") {
    Switch ($MindView7_Language) {
        1 { $MindView7LanguageClear = 'Danish'}
        2 { $MindView7LanguageClear = 'English'}
        3 { $MindView7LanguageClear = 'French'}
        4 { $MindView7LanguageClear = 'German'}
    }
}
Else {
    $MindView7LanguageClear = $LanguageClear
    Switch ($LanguageClear) {
        Dutch { $MindView7LanguageClear = 'English'}
        Italian { $MindView7LanguageClear = 'English'}
        Finnish { $MindView7LanguageClear = 'English'}
        Norwegian { $MindView7LanguageClear = 'English'}
        Polish { $MindView7LanguageClear = 'English'}
        Portuguese { $MindView7LanguageClear = 'English'}
        Swedish { $MindView7LanguageClear = 'English'}
        Japanese { $MindView7LanguageClear = 'English'}
        Korean { $MindView7LanguageClear = 'English'}
        Polish { $MindView7LanguageClear = 'English'}
        Russian { $MindView7LanguageClear = 'English'}
        Spanish { $MindView7LanguageClear = 'English'}
    }
}

If ($Firefox_Architecture -ne "") {
    Switch ($Firefox_Architecture) {
        1 { $FirefoxArchitectureClear = 'x86'}
        2 { $FirefoxArchitectureClear = 'x64'}
    }
}
Else {
    $FirefoxArchitectureClear = $ArchitectureClear
}

If ($Firefox_Language -ne "") {
    Switch ($Firefox_Language) {
        1 { $FFLanguageClear = 'nl'}
        2 { $FFLanguageClear = 'en-US'}
        3 { $FFLanguageClear = 'fr'}
        4 { $FFLanguageClear = 'de'}
        5 { $FFLanguageClear = 'it'}
        6 { $FFLanguageClear = 'ja'}
        7 { $FFLanguageClear = 'pt-PT'}
        8 { $FFLanguageClear = 'ru'}
        9 { $FFLanguageClear = 'es-ES'}
        10 { $FFLanguageClear = 'sv-SE'}
    }
}
Else {
    Switch ($Language) {
        0 { $FFLanguageClear = 'en-US'}
        1 { $FFLanguageClear = 'nl'}
        2 { $FFLanguageClear = 'en-US'}
        3 { $FFLanguageClear = 'en-US'}
        4 { $FFLanguageClear = 'fr'}
        5 { $FFLanguageClear = 'de'}
        6 { $FFLanguageClear = 'it'}
        7 { $FFLanguageClear = 'ja'}
        8 { $FFLanguageClear = 'en-US'}
        9 { $FFLanguageClear = 'en-US'}
        10 { $FFLanguageClear = 'en-US'}
        11 { $FFLanguageClear = 'pt-PT'}
        12 { $FFLanguageClear = 'ru'}
        13 { $FFLanguageClear = 'es-ES'}
        14 { $FFLanguageClear = 'sv-SE'}
    }
}

Switch ($FirefoxChannel) {
    0 { $FirefoxChannelClear = 'LATEST'}
    1 { $FirefoxChannelClear = 'ESR'}
}

If ($NotepadPlusPlus_Architecture -ne "") {
    Switch ($NotepadPlusPlus_Architecture) {
        1 { $NotepadPlusPlusArchitectureClear = 'x86'}
        2 { $NotepadPlusPlusArchitectureClear = 'x64'}
    }
}
Else {
    $NotepadPlusPlusArchitectureClear = $ArchitectureClear
}

If ($OpenJDK_Architecture -ne "") {
    Switch ($OpenJDK_Architecture) {
        1 { $OpenJDKArchitectureClear = 'x86'}
        2 { $OpenJDKArchitectureClear = 'x64'}
    }
}
Else {
    $OpenJDKArchitectureClear = $ArchitectureClear
}

If ($OracleJava8_Architecture -ne "") {
    Switch ($OracleJava8_Architecture) {
        1 { $OracleJava8ArchitectureClear = 'x86'}
        2 { $OracleJava8ArchitectureClear = 'x64'}
    }
}
Else {
    $OracleJava8ArchitectureClear = $ArchitectureClear
}

Switch ($pdfforgePDFCreatorChannel) {
    0 { $pdfforgePDFCreatorChannelClear = 'Free'}
    1 { $pdfforgePDFCreatorChannelClear = 'Professional'}
    2 { $pdfforgePDFCreatorChannelClear = 'Terminal Server'}
}

If ($PeaZip_Architecture -ne "") {
    Switch ($PeaZip_Architecture) {
        1 { $PeaZipArchitectureClear = 'x86'}
        2 { $PeaZipArchitectureClear = 'x64'}
    }
}
Else {
    $PeaZipArchitectureClear = $ArchitectureClear
}

If ($Putty_Architecture -ne "") {
    Switch ($Putty_Architecture) {
        1 { $PuttyArchitectureClear = 'x86'}
        2 { $PuttyArchitectureClear = 'x64'}
    }
}
Else {
    $PuttyArchitectureClear = $ArchitectureClear
}

Switch ($PuttyChannel) {
    0 { $PuttyChannelClear = 'Pre-Release'}
    1 { $PuttyChannelClear = 'Stable'}
}

If ($Slack_Installer -ne "") {
    Switch ($Slack_Installer) {
        1 { $SlackInstallerClear = 'Machine Based'}
        2 { $SlackInstallerClear = 'User Based'}
    }
}
Else {
    $SlackInstallerClear = $InstallerClear
}

If ($SlackInstallerClear -eq 'Machine Based') {
    $SlackInstallerID = '0'
} 
else {
    $SlackInstallerID = '1'
}

Switch ($SlackInstallerID) {
    0 { $SlackPlatformClear = 'PerMachine'}
    1 { $SlackPlatformClear = 'PerUser'}
}

Switch ($SlackInstallerID) {
    0 { 
        If ($Slack_Architecture -ne "") {
            Switch ($Slack_Architecture) {
                1 { $SlackArchitectureClear = 'x86'}
                2 { $SlackArchitectureClear = 'x64'}
            }
        }
        Else {
            $SlackArchitectureClear = $ArchitectureClear
        }
    }
    1 { $SlackArchitectureClear = 'x64'}
}

If ($SumatraPDF_Architecture -ne "") {
    Switch ($SumatraPDF_Architecture) {
        1 { $SumatraPDFArchitectureClear = 'x86'}
        2 { $SumatraPDFArchitectureClear = 'x64'}
    }
}
Else {
    $SumatraPDFArchitectureClear = $ArchitectureClear
}

If ($TechSmithSnagIt_Architecture -ne "") {
    Switch ($TechSmithSnagIt_Architecture) {
        1 { $TechSmithSnagItArchitectureClear = 'x86'}
        2 { $TechSmithSnagItArchitectureClear = 'x64'}
    }
}
Else {
    $TechSmithSnagItArchitectureClear = $ArchitectureClear
}

If ($TotalCommander_Architecture -ne "") {
    Switch ($TotalCommander_Architecture) {
        1 { $TotalCommanderArchitectureClear = 'x86'}
        2 { $TotalCommanderArchitectureClear = 'x64'}
    }
}
Else {
    $TotalCommanderArchitectureClear = $ArchitectureClear
}

If ($VLCPlayer_Architecture -ne "") {
    Switch ($VLCPlayer_Architecture) {
        1 { $VLCPlayerArchitectureClear = 'x86'}
        2 { $VLCPlayerArchitectureClear = 'x64'}
    }
}
Else {
    $VLCPlayerArchitectureClear = $ArchitectureClear
}

If ($VMWareTools_Architecture -ne "") {
    Switch ($VMWareTools_Architecture) {
        1 { $VMWareToolsArchitectureClear = 'x86'}
        2 { $VMWareToolsArchitectureClear = 'x64'}
    }
}
Else {
    $VMWareToolsArchitectureClear = $ArchitectureClear
}

If ($WinMerge_Architecture -ne "") {
    Switch ($WinMerge_Architecture) {
        1 { $WinMergeArchitectureClear = 'x86'}
        2 { $WinMergeArchitectureClear = 'x64'}
    }
}
Else {
    $WinMergeArchitectureClear = $ArchitectureClear
}

If ($Wireshark_Architecture -ne "") {
    Switch ($Wireshark_Architecture) {
        1 { $WiresharkArchitectureClear = 'x86'}
        2 { $WiresharkArchitectureClear = 'x64'}
    }
}
Else {
    $WiresharkArchitectureClear = $ArchitectureClear
}

If ($Zoom_Installer -ne "") {
    Switch ($Zoom_Installer) {
        1 { $ZoomInstallerClear = 'Machine Based'}
        2 { $ZoomInstallerClear = 'User Based'}
    }
}
Else {
    $ZoomInstallerClear = $InstallerClear
}

Write-Host -ForegroundColor Green "Software selection done."
Write-Output ""

Write-Host -ForegroundColor DarkGray "Selected Global Mode."
Write-Host "Global Language is $LanguageClear."
Write-Host "Global Architecture is $ArchitectureClear."
Write-Host "Global Installer Type is $InstallerClear."
Write-Output ""

Write-Host -ForegroundColor DarkGray "Selected Options."
If ($Download -eq "1") { Write-Host -ForegroundColor Green "Download Option Enabled." }
Else { Write-Host "Download Option Disabled." }
If ($Install -eq "1") { Write-Host -ForegroundColor Green "Install Option Enabled." }
Else { Write-Host "Install Option Disabled." }
If ($WhatIf) { Write-Host -ForegroundColor Green "What If Option Enabled." }
Else { Write-Host "What If Option Disabled." }
If ($Repository) { Write-Host -ForegroundColor Green "Installer Repository Option Enabled." }
Else { Write-Host "Installer Repository Option Disabled." }
If ($CleanUp) { Write-Host -ForegroundColor Green "Installer Clean Up Option Enabled." }
Else { Write-Host "Installer Clean Up Option Disabled." }
If ($CleanUpStartMenu) { Write-Host -ForegroundColor Green "Start Menu Clean Up Option Enabled." }
Else { Write-Host "Start Menu Clean Up Option Disabled." }
Write-Output ""

If ($Download -eq "1") {
    # Logging
    # Global variables
    # $StartDir = $PSScriptRoot # the directory path of the script currently being executed
    $LogDir = "$PSScriptRoot\_Install Logs"
    $LogFileName = ("$ENV:COMPUTERNAME - $Date.log")
    $LogFile = Join-path $LogDir $LogFileName
    $FWFileName = ("Firewall - $Date.log")
    $FWFile = Join-path $LogDir $FWFileName
    $WhatIfFileName = ("WhatIf - $Date.log")
    $WhatIfFile = Join-path $LogDir $WhatIfFileName
    $LogTemp = "$env:windir\Logs\Evergreen"

    # Create the log directories if they don't exist
    If (!(Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType directory | Out-Null }
    If (!(Test-Path $LogTemp)) { New-Item -Path $LogTemp -ItemType directory | Out-Null }

    # Create new log file (overwrite existing one)
    New-Item $FWFile -ItemType "file" -force | Out-Null
    DS_WriteLog "I" "START SCRIPT - " $FWFile
    DS_WriteLog "-" "" $FWFile

    # Create new log file (overwrite existing one)
    New-Item $WhatIfFile -ItemType "file" -force | Out-Null
    DS_WriteLog "I" "START SCRIPT - " $WhatIfFile
    DS_WriteLog "-" "" $WhatIfFile

    #// Mark: Install / Update PowerShell module
    Write-Host -ForegroundColor DarkGray "Install / Update PowerShell module!"

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
        Write-Host "Check Evergreen module version."
        $version = (Get-Module -ListAvailable Evergreen) | Sort-Object Version -Descending  | Select-Object Version -First 1
        $psgalleryversion = Find-Module -Name Evergreen | Sort-Object Version -Descending | Select-Object Version -First 1
        $stringver = $version | Select-Object @{n='ModuleVersion'; e={$_.Version -as [string]}}
        $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
        $onlinever = $psgalleryversion | select-object @{n='OnlineVersion'; e={$_.Version -as [string]}}
        $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
        if ([version]"$a" -ge [version]"$b") {
            Write-Host -ForegroundColor Green "Installed Evergreen module version is up to date."
            Write-Output ""
        }
        else {
            Write-Host "Update Evergreen module."
            Update-Module Evergreen -force
            Write-Host -ForegroundColor Green "Update Evergreen module done."
            Write-Output ""
      }
    }

    If (!(Get-Module -ListAvailable -Name Nevergreen)) {
        Write-Host "Install Nevergreen module."
        Install-Module Nevergreen -Force | Import-Module Evergreen
        Write-Host -ForegroundColor Green "Install Nevergreen module done."
        Write-Output ""
    }
    Else {
        Write-Host "Check Nevergreen module version."
        $version = (Get-Module -ListAvailable Nevergreen) | Sort-Object Version -Descending  | Select-Object Version -First 1
        $psgalleryversion = Find-Module -Name Nevergreen | Sort-Object Version -Descending | Select-Object Version -First 1
        $stringver = $version | Select-Object @{n='ModuleVersion'; e={$_.Version -as [string]}}
        $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion
        $onlinever = $psgalleryversion | select-object @{n='OnlineVersion'; e={$_.Version -as [string]}}
        $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion
        if ([version]"$a" -ge [version]"$b") {
            Write-Host -ForegroundColor Green "Installed Nevergreen module version is up to date."
            Write-Output ""
        }
        else {
            Write-Host "Update Nevergreen module."
            Update-Module Nevergreen -force
            Write-Host -ForegroundColor Green "Update Nevergreen module done."
            Write-Output ""
      }
    }

    Write-Host -ForegroundColor DarkGray "Starting downloads..."
    Write-Output ""

    # Download script part (AddScript)

    If ($WhatIf -eq '1') {
        Write-Host -BackgroundColor Magenta "WhatIf Mode, nothing will be downloaded !!!"
        Write-Host -BackgroundColor Green "The FW log will be created !!!"
        Write-Output ""
    }

    #// Mark: Download 1Password
    If ($1Password -eq 1) {
        $Product = "1Password"
        $PackageName = "1Password-Setup"
        $1PasswordD = Get-EvergreenApp -Name 1Password | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $1PasswordD.Version
        $URL = $1PasswordD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download 7 Zip
    If ($7ZIP -eq 1) {
        $Product = "7-Zip"
        $PackageName = "7-Zip_" + "$7ZipArchitectureClear"
        $7ZipD = Get-EvergreenApp -Name 7zip | Where-Object { $_.Architecture -eq "$7ZipArchitectureClear" -and $_.Type -eq "exe" }
        $Version = $7ZipD.Version
        $URL = $7ZipD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$7ZipArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $7ZipArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $7ZipArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msp"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msp" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Include *.msp, *.log, Version.txt, Download* -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "$AdobeArchitectureClear" + "$AdobeLanguageClear" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$AdobeArchitectureClear" + "_$AdobeLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $AdobeArchitectureClear $AdobeLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $AdobeArchitectureClear $AdobeLanguageClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            $PackageNameP = "ReaderTemplate"
            $ReaderDP = Get-AdobeAcrobatReaderDCAdmx
            $VersionP = $ReaderDP.version
            $URL = $ReaderDP.uri
            $URL = $URL.Split(":")[1]
            $URL = "http:" + "$URL"
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "zip"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            Write-Host "Starting download of $Product ADMX files $VersionP"
            If ($WhatIf -eq '0') {
                Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($SourceP))
                expand-archive -path "$PSScriptRoot\$Product\$SourceP" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$PSScriptRoot\$Product\AcrobatReaderDC.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\AcrobatReaderDC.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\AcrobatReaderDC.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\en-US\AcrobatReaderDC.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $VersionP finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Autodesk DWG TrueView
    If ($AutodeskDWGTrueView -eq 1) {
        $Product = "Autodesk DWG TrueView"
        $PackageName = "AutodeskDWGTrueView"
        $AutodeskDWGTrueViewD = Get-DWGTrueView
        $Version = $AutodeskDWGTrueViewD.Version
        $URL = $AutodeskDWGTrueViewD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Exclude *.ps1, *.lnk -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Cisco Webex Teams
    If ($CiscoWebexTeams -eq 1) {
        $Product = "Cisco Webex Teams"
        $PackageName = "webexteams-" + "$CiscoWebexTeamsArchitectureClear"
        $WebexTeamsD = Get-NevergreenApp -Name CiscoWebex | Where-Object { $_.Architecture -eq "$CiscoWebexTeamsArchitectureClear" -and $_.Type -eq "Msi" }
        $Version = $WebexTeamsD.Version
        $URL = $WebexTeamsD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$CiscoWebexTeamsArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $CiscoWebexTeamsArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $CiscoWebexTeamsArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Citrix Files
    If ($CitrixFiles -eq 1) {
        $Product = "Citrix Files"
        $PackageName = "CitrixFilesForWindows"
        $CitrixFilesD = Get-NevergreenApp -Name CitrixFiles | Where-Object {$_.Type -eq "Msi"}
        $Version = $CitrixFilesD.Version
        $URL = $CitrixFilesD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\Citrix\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting Download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\Citrix\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "managementagent" + "$CitrixHypervisorToolsArchitectureClear"
        $CitrixHypervisor = Get-EvergreenApp -Name CitrixVMTools | Where-Object {$_.Architecture -eq "$CitrixHypervisorToolsArchitectureClear"} | Select-Object -Last 1
        $Version = $CitrixHypervisor.Version
        $URL = $CitrixHypervisor.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version_" + "$CitrixHypervisorToolsArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $CitrixHypervisorToolsArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\Citrix\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting Download of $Product $CitrixHypervisorToolsArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\Citrix\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $WSACD = Get-EvergreenApp -Name CitrixWorkspaceApp -WarningAction:SilentlyContinue | Where-Object { $_.Title -like "*Workspace*" -and $_.Stream -like "*$CitrixWorkspaceAppReleaseClear*" }
        $Version = $WSACD.Version
        $URL = $WSACD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -EA SilentlyContinue
        If ($WhatIf -eq '0') {
            If (!(Test-Path -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility")) { New-Item -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility" -ItemType Directory | Out-Null }
        }
        If (!(Test-Path -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.exe")) {
            Write-Host -ForegroundColor Magenta "Download Citrix Receiver Cleanup Utility"
            If ($WhatIf -eq '0') {
                Get-Download https://fileservice.citrix.com/downloadspecial/support/article/CTX137494/downloads/ReceiverCleanupUtility.zip "$PSScriptRoot\Citrix\ReceiverCleanupUtility\" ReceiverCleanupUtility.zip -includeStats
                Expand-Archive -path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.zip" -destinationpath "$PSScriptRoot\Citrix\ReceiverCleanupUtility\"
                Remove-Item -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.zip" -Force
            }
            Write-Host -ForegroundColor Green "Download Citrix Receiver Cleanup Utility finished!"
            Write-Output ""
        }
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\Citrix\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\Citrix\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            $PackageNameP = "CitrixWorkspace_ADMX"
            If ($CitrixWorkspaceAppReleaseClear -eq "Current") {
                $CWADP = Get-CitrixWorkspaceAppCurrentAdmx
            }
            else {
                $CWADP = Get-CitrixWorkspaceAppLTSRAdmx
            }
            $URL = $CWADP.uri
            $Path = $CWADP.path
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "zip"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            If ($CitrixWorkspaceAppReleaseClear -eq "Current") {
                $CWAPath = "CitrixWorkspace_ADMX_Files_" + "$Path"
                $CWASubFolderPath = "$CWAPath" + "\Configuration"
            }
            else {
                $CWAPath = "CitrixWorkspace_ADMX_ADML_Files_" + "$Path"
                $CWASubFolderPath = "$CWAPath"
            }
            Write-Host "Starting download of $Product ADMX files $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\Citrix\$Product\" $SourceP -includeStats
                expand-archive -path "$PSScriptRoot\Citrix\$Product\$SourceP" -destinationpath "$PSScriptRoot\Citrix\$Product"
                Remove-Item -Path "$PSScriptRoot\Citrix\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\CitrixBase.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\receiver.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\en-US\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\en-US\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\de-DE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\de-DE\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\de-DE\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\de-DE\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\es-ES")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\es-ES\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\es-ES\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\es-ES\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\fr-FR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fr-FR\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\fr-FR\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\fr-FR\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\it-IT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\it-IT\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\it-IT\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\it-IT\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ja-JP")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ja-JP\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\ja-JP\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\ja-JP\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\nl-NL")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\nl-NL\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\nl-NL\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\nl-NL\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ko-KR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ko-KR\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\ko-KR\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\ko-KR\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pt-BR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pt-BR\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\pt-BR\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\pt-BR\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ru-RU")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ru-RU\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\ru-RU\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\ru-RU\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\zh-CN")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\zh-CN\receiver.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\receiver.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\CitrixBase.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\zh-CN\receiver.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\Citrix\$Product\$CWASubFolderPath\zh-CN\CitrixBase.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\Citrix\$Product\$CWAPath" -Force -Recurse -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download ControlUp Agent
    If ($ControlUpAgent -eq 1) {
        $Product = "ControlUp Agent"
        $PackageName = "ControlUpAgent-" + "$ControlUpAgentFrameworkClear" + "-$ControlUpAgentArchitectureClear"
        $ControlUpAgentD = Get-EvergreenApp -Name ControlUpAgent | Where-Object { $_.Framework -like "*$ControlUpAgentFrameworkClear" -and $_.Architecture -eq "$ControlUpAgentArchitectureClear" }
        $Version = $ControlUpAgentD.Version
        $URL = $ControlUpAgentD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ControlUpAgentFrameworkClear" + "_$ControlUpAgentArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ControlUpAgentFrameworkClear $ControlUpAgentArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $ControlUpAgentFrameworkClear $ControlUpAgentArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download ControlUp Console
    If ($ControlUpConsole -eq 1) {
        $Product = "ControlUp Console"
        $PackageName = "ControlUpConsole"
        $ControlUpConsoleD = Get-EvergreenApp -Name ControlUpConsole
        $Version = $ControlUpConsoleD.Version
        $URL = $ControlUpConsoleD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product" $Source -includeStats
                expand-archive -path "$PSScriptRoot\$Product\$Source" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\$Source" -Force
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $deviceTRUSTD = Get-EvergreenApp -Name deviceTRUST | Where-Object { $_.Platform -eq "Windows" -and $_.Type -eq "Bundle" }  | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $deviceTRUSTD.Version
        $URL = $deviceTRUSTD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$deviceTRUSTArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $deviceTRUSTArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $deviceTRUSTArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                expand-archive -path "$PSScriptRoot\$Product\deviceTRUST.zip" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\deviceTRUST.zip" -Force
                Remove-Item -Path "$PSScriptRoot\$Product\dtreporting-$Version.zip" -Force
                If (Test-Path -Path "$PSScriptRoot\$Product\dtdemotool-release-$Version.exe") {Remove-Item -Path "$PSScriptRoot\$Product\dtdemotool-release-$Version.exe" -Force}
                Switch ($deviceTRUSTArchitectureClear) {
                    x64 {
                        Get-ChildItem -Path "$PSScriptRoot\$Product" | Where-Object Name -like *"x86"* | Remove-Item
                        Rename-Item -Path "$PSScriptRoot\$Product\dtclient-release-$Version.exe" -NewName "dtclient-release.exe"
                        Rename-Item -Path "$PSScriptRoot\$Product\dtconsole-x64-release-$Version.msi" -NewName "dtconsole-x64-release.msi"
                        Rename-Item -Path "$PSScriptRoot\$Product\dthost-x64-release-$Version.msi" -NewName "dthost-x64-release.msi"
                    }
                    x86 {
                        Get-ChildItem -Path "$PSScriptRoot\$Product" | Where-Object Name -like *"x64"* | Remove-Item
                        Rename-Item -Path "$PSScriptRoot\$Product\dtclient-release-$Version.exe" -NewName "dtclient-release.exe"
                        Rename-Item -Path "$PSScriptRoot\$Product\dtconsole-x86-release-$Version.msi" -NewName "dtconsole-x86-release.msi"
                        Rename-Item -Path "$PSScriptRoot\$Product\dthost-x86-release-$Version.msi" -NewName "dthost-x86-release.msi"
                    }
                }
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            Write-Host "Starting copy of $Product ADMX files $Version"
            If ($WhatIf -eq '0') {
                expand-archive -path "$PSScriptRoot\$Product\dtpolicydefinitions-$Version.zip" -destinationpath "$PSScriptRoot\$Product\_ADMX"
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$PSScriptRoot\$Product\_ADMX\deviceTRUST.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\deviceTRUST.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\deviceTRUST.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\_ADMX\en-US\deviceTRUST.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\dtpolicydefinitions-$Version.zip" -Force
                Remove-Item -Path "$PSScriptRoot\$Product\_ADMX" -Force -Recurse
            }
            Write-Host -ForegroundColor Green "Copy of the new ADMX files version $Version finished!"
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Foxit PDF Editor
    If ($FoxitPDFEditor -eq 1) {
        $Product = "Foxit PDF Editor"
        $PackageName = "FoxitPDFEditor-Setup-" + "$FoxitPDFEditorLanguageClear"
        $FoxitPDFEditorD = Get-EvergreenApp -Name FoxitPDFEditor -WarningAction:SilentlyContinue | Where-Object {$_.Language -eq "$FoxitPDFEditorLanguageClear"}
        $Version = $FoxitPDFEditorD.Version
        $URL = $FoxitPDFEditorD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitPDFEditorLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $FoxitPDFEditorLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $FoxitPDFEditorLanguageClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitReaderLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $FoxitReaderLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $FoxitReaderLanguageClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download GIMP
    If ($GIMP -eq 1) {
        $Product = "GIMP"
        $PackageName = "gimp-setup"
        $GIMPD = Get-EvergreenApp -Name GIMP
        $Version = $GIMPD.Version
        $URL = $GIMPD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Git for Windows
    If ($GitForWindows -eq 1) {
        $Product = "Git for Windows"
        $PackageName = "GitForWindows_" + "$GitForWindowsArchitectureClear"
        $GitForWindowsD = Get-EvergreenApp -Name GitForWindows | Where-Object {$_.Architecture -eq "$GitForWindowsArchitectureClear" -and $_.Type -eq "exe" -and $_.URI -like "*bit.exe"}
        $Version = $GitForWindowsD.Version
        $URL = $GitForWindowsD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$GitForWindowsArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $GitForWindowsArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $GitForWindowsArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "googlechromestandaloneenterprise_" + "$GoogleChromeArchitectureClear"
        $ChromeD = Get-EvergreenApp -Name GoogleChrome | Where-Object { $_.Architecture -eq "$GoogleChromeArchitectureClear" }
        $Version = $ChromeD.Version
        $ChromeSplit = $Version.split(".")
        $ChromeStrings = ([regex]::Matches($Version, "\." )).count
        $ChromeStringLast = ([regex]::Matches($ChromeSplit[$ChromeStrings], "." )).count
        If ($ChromeStringLast -lt "3") {
            $ChromeSplit[$ChromeStrings] = "0" + $ChromeSplit[$ChromeStrings]
        }
        Switch ($ChromeStrings) {
            1 {
                $NewVersion = $ChromeSplit[0] + "." + $ChromeSplit[1]
            }
            2 {
                $NewVersion = $ChromeSplit[0] + "." + $ChromeSplit[1] + "." + $ChromeSplit[2]
            }
            3 {
                $NewVersion = $ChromeSplit[0] + "." + $ChromeSplit[1] + "." + $ChromeSplit[2] + "." + $ChromeSplit[3]
            }
        }
        $URL = $ChromeD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$GoogleChromeArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        $NewCurrentVersion = ""
        If ($CurrentVersion) {
            $CurrentChromeSplit = $CurrentVersion.split(".")
            $CurrentChromeStrings = ([regex]::Matches($CurrentVersion, "\." )).count
            $CurrentChromeStringLast = ([regex]::Matches($CurrentChromeSplit[$CurrentChromeStrings], "." )).count
            If ($CurrentChromeStringLast -lt "3") {
                $CurrentChromeSplit[$CurrentChromeStrings] = "0" + $CurrentChromeSplit[$CurrentChromeStrings]
            }
            Switch ($CurrentChromeStrings) {
                1 {
                    $NewCurrentVersion = $CurrentChromeSplit[0] + "." + $CurrentChromeSplit[1]
                }
                2 {
                    $NewCurrentVersion = $CurrentChromeSplit[0] + "." + $CurrentChromeSplit[1] + "." + $CurrentChromeSplit[2]
                }
                3 {
                    $NewCurrentVersion = $CurrentChromeSplit[0] + "." + $CurrentChromeSplit[1] + "." + $CurrentChromeSplit[2] + "." + $CurrentChromeSplit[3]
                }
            }
        }
        Write-Host -ForegroundColor Magenta "Download $Product $GoogleChromeArchitectureClear"
        Write-Host "Download Version: $NewVersion"
        Write-Host "Current Version:  $NewCurrentVersion"
        If ($NewCurrentVersion -lt $NewVersion) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($NewCurrentVersion) {
                        Write-Host "Copy $Product Installer Version $NewCurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$NewCurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$NewCurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$NewCurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $NewCurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $GoogleChromeArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            $PackageNameP = "Chrome-Templates"
            $ChromeDP = Get-GoogleChromeAdmx
            $VersionP = $ChromeDP.version
            $URL = $ChromeDP.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "zip"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            Write-Host "Starting download of $Product ADMX files $VersionP"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                expand-archive -path "$PSScriptRoot\$Product\$SourceP" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\google.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\chrome.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\en-US\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\en-US\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\de-DE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\de-DE\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\de-DE\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\de-DE\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\es-ES")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\es-ES\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\es-ES\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\es-ES\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\fr-FR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fr-FR\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fr-FR\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fr-FR\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\it-IT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\it-IT\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\it-IT\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\it-IT\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ja-JP")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ja-JP\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ja-JP\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ja-JP\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\nl-NL")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\nl-NL\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nl-NL\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nl-NL\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ko-KR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ko-KR\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ko-KR\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ko-KR\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pt-BR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pt-BR\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-BR\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-BR\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ru-RU")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ru-RU\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ru-RU\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ru-RU\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\zh-CN")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\zh-CN\chrome.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\google.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\chrome.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\zh-CN\chrome.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\zh-CN\google.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\common" -Force -Recurse -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\chromeos" -Force -Recurse -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\VERSION" -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\windows" -Force -Recurse -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $VersionP finished!"
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download ImageGlass
    If ($ImageGlass -eq 1) {
        $Product = "ImageGlass"
        $PackageName = "ImageGlass_" + "$ImageGlassArchitectureClear"
        $ImageGlassD = Get-EvergreenApp -Name ImageGlass | Where-Object { $_.Architecture -eq "$ImageGlassArchitectureClear" }
        $Version = $ImageGlassD.Version
        $URL = $ImageGlassD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ImageGlassArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $ImageGlassArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $ImageGlassArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $Product = "IrfanView"
        $PackageName = "IrfanView" + "$IrfanViewArchitectureClear"
        $IrfanViewD = Get-IrfanView | Where-Object {$_.Architecture -eq "$IrfanViewArchitectureClear"}
        $Version = $IrfanViewD.Version
        $URL = $IrfanViewD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$IrfanViewArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path $VersionPath -EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product $IrfanViewLanguageLongClear $IrfanViewArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $IrfanViewArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            If ($IrfanViewLanguageLongClear -ne "English") {
                $IrfanViewLD = Get-NevergreenApp -Name IrfanView | Where-Object {$_.Language -eq "$IrfanViewLanguageClear" -and $_.Type -eq "Zip"}
                $VersionL = $IrfanViewLD.Version
                $URLL = $IrfanViewLD.uri
                $PackageNameL = "IrfanView_lang_" + "$IrfanViewLanguageLongClear"
                $SourceL = "$PackageNameL" + ".zip"
                Write-Host "Starting download of $Product $IrfanViewLanguageLongClear language pack version $VersionL"
                If ($WhatIf -eq '0') {
                    Get-Download $URLL "$PSScriptRoot\$Product\" $SourceL -includeStats
                    expand-archive -path "$PSScriptRoot\$Product\$SourceL" -destinationpath "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear"
                    Remove-Item "$PSScriptRoot\$Product\$SourceL"
                }
                Write-Host -ForegroundColor Green "Download of the $IrfanViewLanguageLongClear language pack version $VersionL finished!"
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"-EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product $KeePassLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            $KeePassLD = Get-KeePassLanguage | Where-Object {$_.Language -eq "$KeePassLanguageClear"}
            $VersionL = $KeePassLD.Version
            $URLL = $KeePassLD.uri
            $PackageNameL = "KeePass-" + "$KeePassLanguageClear"
            $SourceL = "$PackageNameL" + ".zip"
            Write-Host "Starting download of $Product $KeePassLanguageClear language pack version $VersionL"
            If ($WhatIf -eq '0') {
                Get-Download $URLL "$PSScriptRoot\$Product\" $SourceL -includeStats
                expand-archive -path "$PSScriptRoot\$Product\$SourceL" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item "$PSScriptRoot\$Product\$SourceL"
            }
            Write-Host -ForegroundColor Green "Download of the $KeePassLanguageClear language pack version $VersionL finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download LogMeIn GoToMeeting
    If ($LogMeInGoToMeeting -eq 1) {
        If ($LogMeInGoToMeetingInstallerClear -eq 'Machine Based') {
            $Product = "LogMeIn GoToMeeting XenApp"
            $PackageName = "GoToMeeting-Setup"
            $LogMeInGoToMeetingD = Get-EvergreenApp -Name LogMeInGoToMeeting | Where-Object { $_.Type -eq "XenAppLatest" }
            $Version = $LogMeInGoToMeetingD.Version
            $URL = $LogMeInGoToMeetingD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerType = "msi"
            $Source = "$PackageName" + "." + "$InstallerType"
            $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"-EA SilentlyContinue 
            Write-Host -ForegroundColor Magenta "Download $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion"
            If ($CurrentVersion -lt $Version) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                }
                Write-Host "Starting download of $Product $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
        If ($LogMeInGoToMeetingInstallerClear -eq 'User Based') {
            $Product = "LogMeIn GoToMeeting"
            $PackageName = "GoToMeeting-Setup"
            $LogMeInGoToMeetingD = Get-EvergreenApp -Name LogMeInGoToMeeting | Where-Object { $_.Type -eq "Latest" }
            $Version = $LogMeInGoToMeetingD.Version
            $URL = $LogMeInGoToMeetingD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerType = "msi"
            $Source = "$PackageName" + "." + "$InstallerType"
            $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"-EA SilentlyContinue 
            Write-Host -ForegroundColor Magenta "Download $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion"
            If ($CurrentVersion -lt $Version) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                }
                Write-Host "Starting download of $Product $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
    }

    #// Mark: Download Microsoft .Net Framework
    If ($MSDotNetFramework -eq 1) {
        $Product = "Microsoft Dot Net Framework"
        $PackageName = "NetFramework-runtime_" + "$MSDotNetFrameworkArchitectureClear" + "_$MSDotNetFrameworkChannelClear"
        $MSDotNetFrameworkD = Get-EvergreenApp -Name Microsoft.NET | Where-Object {$_.Architecture -eq "$MSDotNetFrameworkArchitectureClear" -and $_.Channel -eq "$MSDotNetFrameworkChannelClear"}
        $Version = $MSDotNetFrameworkD.Version
        $URL = $MSDotNetFrameworkD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSDotNetFrameworkArchitectureClear" + "_$MSDotNetFrameworkChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSDotNetFrameworkArchitectureClear $MSDotNetFrameworkChannelClear Channel"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSDotNetFrameworkArchitectureClear $MSDotNetFrameworkChannelClear Channel $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MS365AppsChannelClear setup file"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($WhatIf -eq '0') {
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
                If ($MS365AppsInstallerClear -eq 'Machine Based') {
                    Write-Host "Create install.xml for Machine Based Install"
                    [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
                    [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                        $XML.appendChild($Root) | out-null
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
                        $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product\$MS365AppsChannelClear")
                        $Node1.SetAttribute("OfficeClientEdition","$MS365AppsArchitectureClear")
                        $Node1.SetAttribute("Channel","$MS365AppsChannelClearDL")
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
                    If ($MS365Apps_Visio -eq '1') {
                        Write-Host "Add Microsoft Visio to install.xml for Machine Based Install"
                        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                        $Node2.SetAttribute("ID","VisioProRetail")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","MatchOS")
                            $Node3.SetAttribute("Fallback","en-us")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","$MS365AppsVisioLanguageClear")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Teams")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Lync")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Groove")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","OneDrive")
                    }
                    If ($MS365Apps_Project -eq '1') {
                        Write-Host "Add Microsoft Project to install.xml for Machine Based Install"
                        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                        $Node2.SetAttribute("ID","ProjectProRetail")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","MatchOS")
                            $Node3.SetAttribute("Fallback","en-us")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","$MS365AppsProjectLanguageClear")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Teams")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Lync")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Groove")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","OneDrive")
                    }
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
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","DeviceBasedLicensing")
                        $Node1.SetAttribute("Value","0")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","AUTOACTIVATE")
                        $Node1.SetAttribute("Value","1")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Updates"))
                        $Node1.SetAttribute("Enabled","FALSE")
                        $XML.Save("$PSScriptRoot\$Product\$MS365AppsChannelClear\install.xml")
                    Write-Host -ForegroundColor Green "Create install.xml for Machine Based Install finished!"
                }
                If ($MS365AppsInstallerClear -eq 'User Based') {
                    Write-Host "Create install.xml for User Based Install"
                    [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
                    [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                        $XML.appendChild($Root) | out-null
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
                        $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product\$MS365AppsChannelClear")
                        $Node1.SetAttribute("OfficeClientEdition","$MS365AppsArchitectureClear")
                        $Node1.SetAttribute("Channel","$MS365AppsChannelClearDL")
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
                    If ($MS365Apps_Visio -eq '1') {
                        Write-Host "Add Microsoft Visio to install.xml for User Based Install"
                        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                        $Node2.SetAttribute("ID","VisioProRetail")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","MatchOS")
                            $Node3.SetAttribute("Fallback","en-us")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","$MS365AppsVisioLanguageClear")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Teams")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Lync")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Groove")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","OneDrive")
                    }
                    If ($MS365Apps_Project -eq '1') {
                        Write-Host "Add Microsoft Project to install.xml for User Based Install"
                        [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                        $Node2.SetAttribute("ID","ProjectProRetail")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","MatchOS")
                            $Node3.SetAttribute("Fallback","en-us")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                            $Node3.SetAttribute("ID","$MS365AppsProjectLanguageClear")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Teams")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Lync")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","Groove")
                        [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("ExcludeApp"))
                            $Node3.SetAttribute("ID","OneDrive")
                    }
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Display"))
                        $Node1.SetAttribute("Level","None")
                        $Node1.SetAttribute("AcceptEULA","TRUE")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Logging"))
                        $Node1.SetAttribute("Level","Standard")
                        $Node1.SetAttribute("Path","%temp%")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","SharedComputerLicensing")
                        $Node1.SetAttribute("Value","0")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","FORCEAPPSHUTDOWN")
                        $Node1.SetAttribute("Value","TRUE")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Updates"))
                        $Node1.SetAttribute("Enabled","FALSE")
                        $XML.Save("$PSScriptRoot\$Product\$MS365AppsChannelClear\install.xml")
                    Write-Host -ForegroundColor Green "Create install.xml for User Based Install finished!"
                }
            }
        }
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                $LogPS = "$PSScriptRoot\$Product\$MS365AppsChannelClear\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\$MS365AppsChannelClear\*" -Recurse -Exclude install.xml,remove.xml
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MS365AppsChannelClear $Version setup file"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\$MS365AppsChannelClear" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version setup file finished!"
            Write-Output ""
            $PackageNameP = "admintemplates-office"
            $MS365AppsPD = Get-MicrosoftOfficeAdmx| Where-Object {$_.Architecture -eq "$MS365AppsArchitecturePolicyClear"}
            $Version = $MS365AppsPD.Version
            $URL = $MS365AppsPD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "exe"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            Write-Host "Starting download of $Product ADMX Files $Version"
            $InstallDir = "$PSScriptRoot\$Product\$SourceP"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                Start-Process -FilePath "$InstallDir" -ArgumentList "/extract:$env:TEMP /passive /quiet" -wait
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\office16.admx" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse
                }
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                copy-item -Path "$env:TEMP\admx\*" -Destination "$PSScriptRoot\_ADMX\$Product" -Force -Recurse
                copy-item -Path "$env:TEMP\office2016grouppolicyandoctsettings.xlsx" -Destination "$PSScriptRoot\_ADMX\$Product" -Force
                Remove-Item -Path "$InstallDir" -Force
                Remove-Item -Path "$env:TEMP\admx" -Force -Recurse
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft AVD Remote Desktop
    If ($MSAVDRemoteDesktop -eq 1) {
        $Product = "Microsoft AVD Remote Desktop"
        $PackageName = "RemoteDesktop_" + "$MSAVDRemoteDesktopArchitectureClear" + "_$MSAVDRemoteDesktopChannelClear"
        $MSAVDRemoteDesktopD = Get-EvergreenApp -Name MicrosoftWVDRemoteDesktop | Where-Object { $_.Architecture -eq "$MSAVDRemoteDesktopArchitectureClear" -and $_.Channel -eq "$MSAVDRemoteDesktopChannelClear" }
        $Version = $MSAVDRemoteDesktopD.Version
        $URL = $MSAVDRemoteDesktopD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSAVDRemoteDesktopChannelClear" + "$MSAVDRemoteDesktopArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSAVDRemoteDesktopChannelClear $MSAVDRemoteDesktopArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSAVDRemoteDesktopChannelClear $MSAVDRemoteDesktopArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Azure CLI
    If ($MSAzureCLI -eq 1) {
        $Product = "Microsoft Azure CLI"
        $PackageName = "AzureCLI"
        $MSAzureCLID = Get-NevergreenApp -Name MicrosoftAzureCLI | Where-Object { $_.Type -eq "Msi" }
        $Version = $MSAzureCLID.Version
        $URL = $MSAzureCLID.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Azure Data Studio
    If ($MSAzureDataStudio -eq 1) {
        $Product = "Microsoft Azure Data Studio"
        $PackageName = "AzureDataStudio-Setup-"
        $MSAzureDataStudioD = Get-EvergreenApp -Name microsoftazuredatastudio | Where-Object { $_.Channel -eq "$MSAzureDataStudioChannelClear" -and $_.Platform -eq "$MSAzureDataStudioPlatformClear"}
        $Version = $MSAzureDataStudioD.Version
        $URL = $MSAzureDataStudioD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "$MSAzureDataStudioChannelClear" + "-$MSAzureDataStudioPlatformClear" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSAzureDataStudioChannelClear" + "-$MSAzureDataStudioPlatformClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSAzureDataStudioChannelClear $ArchitectureClear $MSAzureDataStudioModeClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSAzureDataStudioChannelClear $ArchitectureClear $MSAzureDataStudioModeClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
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
        $PackageName = "MicrosoftEdgeEnterprise_" + "$MSEdgeArchitectureClear" + "_$MSEdgeChannelClear"
        $EdgeD = Get-EvergreenApp -Name MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Release -eq "Enterprise" -and $_.Channel -eq "$MSEdgeChannelClear" -and $_.Architecture -eq "$MSEdgeArchitectureClear" }
        $Version = $EdgeD.Version
        $URL = $EdgeD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSEdgeArchitectureClear" + "_$MSEdgeChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product $MSEdgeChannelClear $MSEdgeArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSEdgeChannelClear $MSEdgeArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            $PackageNameP = "MicrosoftEdgePolicy"
            $EdgeDP = Get-EvergreenApp -name microsoftedge | Where-Object { $_.Channel -eq "Policy" }
            $URL = $EdgeDP.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "cab"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            Write-Host "Starting download of $Product $MSEdgeChannelClear ADMX files $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                expand ."$PSScriptRoot\$Product\$SourceP" ."$PSScriptRoot\$Product\MicrosoftEdgePolicyTemplates.zip" | Out-Null
                expand-archive -path "$PSScriptRoot\$Product\MicrosoftEdgePolicyTemplates.zip" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\MicrosoftEdgePolicyTemplates.zip" -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\mac" -Force -Recurse -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\html" -Force -Recurse -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\examples" -Force -Recurse -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\VERSION" -Force -ErrorAction SilentlyContinue
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\msedge.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\msedgeupdate.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\msedgewebview2.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\en-US\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\en-US\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\en-US\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\de-DE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\de-DE\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\de-DE\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\de-DE\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\de-DE\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\da-DK")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\da-DK" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\da-DK\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\da-DK\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\da-DK\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\da-DK\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\da-DK\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\da-DK" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\da-DK\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\da-DK" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\da-DK\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\da-DK" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\es-ES")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\es-ES\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\es-ES\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\es-ES\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\es-ES\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\fi-FI")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\fi-FI" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fi-FI\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fi-FI\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fi-FI\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fi-FI\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fi-FI\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fi-FI" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fi-FI\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fi-FI" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fi-FI\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fi-FI" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\fr-FR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fr-FR\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fr-FR\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fr-FR\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\fr-FR\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\it-IT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\it-IT\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\it-IT\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\it-IT\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\it-IT\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ja-JP")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ja-JP\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ja-JP\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ja-JP\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ja-JP\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ko-KR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ko-KR\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ko-KR\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ko-KR\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ko-KR\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\nb-NO")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\nb-NO" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\nb-NO\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nb-NO\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nb-NO\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nb-NO\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nb-NO\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nb-NO" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nb-NO\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nb-NO" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nb-NO\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nb-NO" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\nl-NL")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\nl-NL\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nl-NL\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nl-NL\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\nl-NL\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pl-PL")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pl-PL" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pl-PL\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pl-PL\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pl-PL\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pl-PL\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pl-PL\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pl-PL" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pl-PL\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pl-PL" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pl-PL\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pl-PL" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pt-BR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pt-BR\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-BR\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-BR\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-BR\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pt-PT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-PT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pt-PT\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-PT\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-PT\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-PT\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-PT\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-PT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-PT\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-PT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\pt-PT\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-PT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ru-RU")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ru-RU\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ru-RU\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ru-RU\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\ru-RU\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\sv-SE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\sv-SE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\sv-SE\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\sv-SE\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\sv-SE\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\sv-SE\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\sv-SE\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\sv-SE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\sv-SE\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\sv-SE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\sv-SE\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\sv-SE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\zh-CN")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\zh-CN\msedge.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\msedgeupdate.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\msedge.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\msedgewebview2.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\zh-CN\msedge.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\zh-CN\msedgeupdate.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\admx\zh-CN\msedgewebview2.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\windows" -Force -Recurse -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Edge WebView2
    If ($MSEdgeWebView2 -eq 1) {
        $Product = "Microsoft Edge WebView2"
        $PackageName = "MicrosoftEdgeWebView2_" + "$MSEdgeWebView2ArchitectureClear"
        $EdgeWebView2D = Get-EvergreenApp -Name MicrosoftEdgeWebView2Runtime | Where-Object { $_.Architecture -eq "$MSEdgeWebView2ArchitectureClear" }
        $Version = $EdgeWebView2D.Version
        $URL = $EdgeWebView2D.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSEdgeWebView2ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue 
        Write-Host -ForegroundColor Magenta "Download $Product $MSEdgeWebView2ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSEdgeWebView2ArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft FSLogix
    If ($MSFSLogix -eq 1) {
        $Product = "Microsoft FSLogix"
        $PackageName = "FSLogixAppsSetup_" + "$MSFSLogixChannelCLear"
        $MSFSLogixD = Get-EvergreenApp -Name MicrosoftFSLogixApps -ea silentlyContinue -WarningAction silentlyContinue | Where-Object { $_.Channel -eq "$MSFSLogixChannelClear"} 
        $Version = $MSFSLogixD.Version
        $URL = $MSFSLogixD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\$MSFSLogixChannelClear\Version_"+ "$MSFSLogixChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear")) { New-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\$MSFSLogixChannelClear\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\$MSFSLogixChannelClear\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\$MSFSLogixChannelClear" $Source -includeStats
                expand-archive -path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\$Source" -destinationpath "$PSScriptRoot\$Product\$MSFSLogixChannelClear"
                Remove-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\$Source" -Force
                Switch ($MSFSLogixArchitectureClear) {
                    x86 {
                        Move-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\Win32\Release\*" -Destination "$PSScriptRoot\$Product\$MSFSLogixChannelClear"
                    }
                    x64 {
                        Move-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\x64\Release\*" -Destination "$PSScriptRoot\$Product\$MSFSLogixChannelClear"
                    }
                }
                Remove-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\Win32" -Force -Recurse
                Remove-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\x64" -Force -Recurse
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            Write-Host "Starting copy of $Product $MSFSLogixChannelClear ADMX files $Version"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fslogix.admx" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fslogix.admx" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\fslogix.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\fslogix.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\fslogix.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\$MSFSLogixChannelClear\fslogix.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Copy of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Office
    If ($MSOffice -eq 1) {
        $Product = "Microsoft Office " + $MSOfficeChannelClear
        $PackageName = "setup"
        $MSOfficeD = Get-EvergreenApp -Name Microsoft365Apps | Where-Object {$_.Channel -eq "$MSOfficeVersionClear"}
        $Version = $MSOfficeD.Version
        $URL = $MSOfficeD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product setup file"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($WhatIf -eq '0') {
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
                If ($MSOfficeInstallerClear -eq 'Machine Based') {
                    Write-Host "Create install.xml for Machine Based Install"
                    [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
                    [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                        $XML.appendChild($Root) | out-null
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
                        $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product")
                        $Node1.SetAttribute("OfficeClientEdition","$MSOfficeArchitectureClear")
                        $Node1.SetAttribute("Channel","$MSOfficeVersionClear")
                    [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                        $Node2.SetAttribute("ID","$MSOfficeProductIDClear")
                    [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                        $Node3.SetAttribute("ID","MatchOS")
                        $Node3.SetAttribute("Fallback","en-us")
                    [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                        $Node3.SetAttribute("ID","$MSOfficeLanguageClear")
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
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","DeviceBasedLicensing")
                        $Node1.SetAttribute("Value","0")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","AUTOACTIVATE")
                        $Node1.SetAttribute("Value","1")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Updates"))
                        $Node1.SetAttribute("Enabled","FALSE")
                        $XML.Save("$PSScriptRoot\$Product\install.xml")
                    Write-Host -ForegroundColor Green  "Create install.xml for Machine Based Install finished!"
                }
                If ($MSOfficeInstallerClear -eq 'User Based') {
                    Write-Host "Create install.xml for User Based Install"
                    [System.XML.XMLDocument]$XML=New-Object System.XML.XMLDocument
                    [System.XML.XMLElement]$Root = $XML.CreateElement("Configuration")
                        $XML.appendChild($Root) | out-null
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Add"))
                        $Node1.SetAttribute("SourcePath","$PSScriptRoot\$Product")
                        $Node1.SetAttribute("OfficeClientEdition","$MSOfficeArchitectureClear")
                        $Node1.SetAttribute("Channel","$MSOfficeVersionClear")
                    [System.XML.XMLElement]$Node2 = $Node1.AppendChild($XML.CreateElement("Product"))
                        $Node2.SetAttribute("ID","$MSOfficeProductIDClear")
                    [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                        $Node3.SetAttribute("ID","MatchOS")
                        $Node3.SetAttribute("Fallback","en-us")
                    [System.XML.XMLElement]$Node3 = $Node2.AppendChild($XML.CreateElement("Language"))
                        $Node3.SetAttribute("ID","$MSOfficeLanguageClear")
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
                        $Node1.SetAttribute("Value","0")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Property"))
                        $Node1.SetAttribute("Name","FORCEAPPSHUTDOWN")
                        $Node1.SetAttribute("Value","TRUE")
                    [System.XML.XMLElement]$Node1 = $Root.AppendChild($XML.CreateElement("Updates"))
                        $Node1.SetAttribute("Enabled","FALSE")
                        $XML.Save("$PSScriptRoot\$Product\install.xml")
                    Write-Host -ForegroundColor Green  "Create install.xml for User Based Install finished!"
                }
            }
        }
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse -Exclude install.xml,remove.xml
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version setup file"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            $PackageNameP = "admintemplates-office"
            $MSOfficePD = Get-MicrosoftOfficeAdmx | Where-Object {$_.Architecture -eq "$MSOfficeArchitecturePolicyClear"}
            $Version = $MSOfficePD.Version
            $URL = $MSOfficePD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "exe"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            Write-Host "Starting download of $Product ADMX Files $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                $InstallDir = "$PSScriptRoot\$Product\$SourceP"
                Start-Process -FilePath "$InstallDir" -ArgumentList "/extract:$env:TEMP /passive /quiet" -wait
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\office16.admx" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse
                }
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                copy-item -Path "$env:TEMP\admx\*" -Destination "$PSScriptRoot\_ADMX\$Product" -Force -Recurse
                copy-item -Path "$env:TEMP\office2016grouppolicyandoctsettings.xlsx" -Destination "$PSScriptRoot\_ADMX\$Product" -Force
                Remove-Item -Path "$InstallDir" -Force
                Remove-Item -Path "$env:TEMP\admx" -Force -Recurse
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $Version finished!"
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Power BI Desktop
    If ($MSPowerBIDesktop -eq 1) {
        $Product = "Microsoft Power BI Desktop"
        $PackageName = "PBIDesktopSetup_" + "$MSPowerBIDesktopArchitectureClear"
        $MSPowerBIDesktopD = Get-NevergreenApp -Name MicrosoftPowerBIDesktop | Where-Object { $_.Architecture -eq "$MSPowerBIDesktopArchitectureClear"}
        $Version = $MSPowerBIDesktopD.Version
        $URL = $MSPowerBIDesktopD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSPowerBIDesktopArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSPowerBIDesktopArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSPowerBIDesktopArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Power BI Report Builder
    If ($MSPowerBIReportBuilder -eq 1) {
        $Product = "Microsoft Power BI Report Builder"
        $PackageName = "PBIReportBuilderSetup"
        $MSPowerBIReportBuilderD = Get-NevergreenApp -Name MicrosoftPowerBIReportBuilder
        $Version = $MSPowerBIReportBuilderD.Version
        $URL = $MSPowerBIReportBuilderD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "PowerShell" + "$MSPowerShellArchitectureClear" + "_$MSPowerShellReleaseClear"
        $MSPowershellD = Get-EvergreenApp -Name MicrosoftPowerShell | Where-Object {$_.Architecture -eq "$MSPowerShellArchitectureClear" -and $_.Release -eq "$MSPowerShellReleaseClear"}
        $Version = $MSPowershellD.Version
        $URL = $MSPowershellD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSPowerShellArchitectureClear" + "_$MSPowerShellReleaseClear" + ".txt"
        $CurrentVersion = Get-Content -Path $VersionPath -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSPowerShellArchitectureClear $MSPowerShellReleaseClear Release"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSPowerShellArchitectureClear $MSPowerShellReleaseClear Release $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft PowerToys
    If ($MSPowerToys -eq 1) {
        $Product = "Microsoft PowerToys"
        $PackageName = "PowerToysSetup-x64"
        $MSPowerToysD = Get-EvergreenApp -Name MicrosoftPowerToys
        $Version = $MSPowerToysD.Version
        $URL = $MSPowerToysD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft SQL Server Management Studio
    If ($MSSQLServerManagementStudio -eq 1) {
        $Product = "Microsoft SQL Server Management Studio"
        $PackageName = "SSMS-Setup_" + "$MSSQLServerManagementStudioLanguageClear"
        $MSSQLServerManagementStudioD = Get-EvergreenApp -Name MicrosoftSsms -ea silentlyContinue -WarningAction silentlyContinue | Where-Object { $_.Language -eq "$MSSQLServerManagementStudioLanguageClear" }
        $Version = $MSSQLServerManagementStudioD.Version
        $URL = $MSSQLServerManagementStudioD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSSQLServerManagementStudioLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSSQLServerManagementStudioLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSSQLServerManagementStudioLanguageClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Sysinternals
    If ($MSSysinternals -eq 1) {
        $Product = "Microsoft Sysinternals"
        $PackageName = "SysinternalsSuite"
        $MSSysinternalsD = Get-NevergreenApp -Name MicrosoftSysinternals | Where-Object { $_.Type -eq "Zip" -and $_.Architecture -eq "Multi" -and $_.Name -eq "Microsoft Sysinternals Suite" }
        $Version = $MSSysinternalsD.Version
        $URL = $MSSysinternalsD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Copy-Item -Path "$PSScriptRoot\$Product\*.chm" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product" $Source -includeStats
                expand-archive -path "$PSScriptRoot\$Product\$Source" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\$Source" -Force
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "Teams_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear"
        If ($MSTeamsInstallerClear -eq 'Machine Based') {
            $Product = "Microsoft Teams Machine Based"
            If ($MSTeamsRingClear -eq 'Continuous Deployment' -or $MSTeamsRingClear -eq 'Exploration') {
                $TeamsD = Get-NevergreenApp -Name MicrosoftTeams | Where-Object { $_.Architecture -eq "$MSTeamsArchitectureClear" -and $_.Ring -eq "$MSTeamsRingClear" -and $_.Type -eq "MSI"}
            }
            Else {
                $TeamsD = Get-EvergreenApp -Name MicrosoftTeams | Where-Object { $_.Architecture -eq "$MSTeamsArchitectureClear" -and $_.Ring -eq "$MSTeamsRingClear"}
            }
            $Version = $TeamsD.Version
            If ($Version) {
                $TeamsSplit = $Version.split(".")
                $TeamsStrings = ([regex]::Matches($Version, "\." )).count
                $TeamsStringLast = ([regex]::Matches($TeamsSplit[$TeamsStrings], "." )).count
                If ($TeamsStringLast -lt "5") {
                    $TeamsSplit[$TeamsStrings] = "0" + $TeamsSplit[$TeamsStrings]
                }
                $NewVersion = $TeamsSplit[0] + "." + $TeamsSplit[1] + "." + $TeamsSplit[2] + "." + $TeamsSplit[3]
            }
            $URL = $TeamsD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerType = "msi"
            $Source = "$PackageName" + "." + "$InstallerType"
            $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
            $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
            If ($CurrentVersion) {
                $CurrentTeamsSplit = $CurrentVersion.split(".")
                $CurrentTeamsStrings = ([regex]::Matches($CurrentVersion, "\." )).count
                $CurrentTeamsStringLast = ([regex]::Matches($CurrentTeamsSplit[$CurrentTeamsStrings], "." )).count
                If ($CurrentTeamsStringLast -lt "5") {
                    $CurrentTeamsSplit[$CurrentTeamsStrings] = "0" + $CurrentTeamsSplit[$CurrentTeamsStrings]
                }
                $NewCurrentVersion = $CurrentTeamsSplit[0] + "." + $CurrentTeamsSplit[1] + "." + $CurrentTeamsSplit[2] + "." + $CurrentTeamsSplit[3]
            }
            Write-Host -ForegroundColor Magenta "Download $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion"
            If ($NewCurrentVersion -ne $NewVersion) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$VersionPath" -Value "$Version"
                }
                Write-Host "Starting download of $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
        If ($MSTeamsInstallerClear -eq 'User Based') {
            $Product = "Microsoft Teams User Based"
            $TeamsD = Get-MicrosoftTeamsUser | Where-Object { $_.Architecture -eq "$MSTeamsArchitectureClear" -and $_.Ring -eq "$MSTeamsRingClear"}
            $Version = $TeamsD.Version
            $URL = $TeamsD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerType = "exe"
            $Source = "$PackageName" + "." + "$InstallerType"
            $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
            $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
            Write-Host -ForegroundColor Magenta "Download $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion"
            If ($CurrentVersion -lt $Version) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$VersionPath" -Value "$Version"
                }
                Write-Host "Starting download of $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
    }

    #// Mark: Download Microsoft Visual Studio 2019
    If ($MSVisualStudio -eq 1) {
        $Product = "Microsoft Visual Studio 2019"
        $PackageName = "VS-Setup"
        $MSVisualStudioD = Get-EvergreenApp -Name MicrosoftVisualStudio
        $Version = $MSVisualStudioD.Version
        $URL = $MSVisualStudioD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Visual Studio Code
    If ($MSVisualStudioCode -eq 1) {
        $Product = "Microsoft Visual Studio Code"
        $PackageName = "VSCode-Setup-"
        $MSVisualStudioCodeD = Get-EvergreenApp -Name MicrosoftVisualStudioCode | Where-Object { $_.Architecture -eq "$MSVisualStudioCodeArchitectureClear" -and $_.Channel -eq "$MSVisualStudioCodeChannelClear" -and $_.Platform -eq "$MSVisualStudioCodePlatformClear"}
        $Version = $MSVisualStudioCodeD.Version
        $URL = $MSVisualStudioCodeD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "$MSVisualStudioCodeChannelClear" + "-$MSVisualStudioCodePlatformClear" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSVisualStudioCodeChannelClear" + "-$MSVisualStudioCodePlatformClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MSVisualStudioCodeChannelClear $MSVisualStudioCodeArchitectureClear $MSVisualStudioCodeModeClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MSVisualStudioCodeChannelClear $MSVisualStudioCodeArchitectureClear $MSVisualStudioCodeModeClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download MindView 7
    If ($MindView7 -eq 1) {
        $Product = "MindView 7"
        $PackageName = "MindView7-"
        $MindView7D = Get-MindView7 | Where-Object { $_.Language -eq "$MindView7LanguageClear"}
        $Version = $MindView7D.Version
        $URL = $MindView7D.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "$MindView7LanguageClear" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MindView7LanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $MindView7LanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $MindView7LanguageClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "Firefox_Setup_" + "$FirefoxChannelClear" + "$FirefoxArchitectureClear" + "_$FFLanguageClear"
        $FirefoxD = Get-EvergreenApp -Name MozillaFirefox | Where-Object { $_.Type -eq "msi" -and $_.Architecture -eq "$FirefoxArchitectureClear" -and $_.Channel -like "*$FirefoxChannelClear*" -and $_.Language -eq "$FFLanguageClear"}
        $Version = $FirefoxD.Version
        $URL = $FirefoxD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FirefoxChannelClear" + "$FirefoxArchitectureClear" + "$FFLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $FirefoxChannelClear $FirefoxArchitectureClear $FFLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $FirefoxChannelClear $FirefoxArchitectureClear $FFLanguageClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            $PackageNameP = "Firefox-Templates"
            $FirefoxDP = Get-MozillaFirefoxAdmx
            $VersionP = $FirefoxDP.version
            $URL = $FirefoxDP.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerTypeP = "zip"
            $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
            Write-Host "Starting download of $Product ADMX files $VersionP"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                expand-archive -path "$PSScriptRoot\$Product\$SourceP" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$PSScriptRoot\$Product\windows\firefox.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\mozilla.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\firefox.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\mozilla.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\en-US\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\en-US\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\de-DE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\de-DE\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\firefox.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\mozilla.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\de-DE\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\de-DE\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\es-ES")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\es-ES\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\mozilla.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\firefox.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\es-ES\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\es-ES\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\fr-FR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fr-FR\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\mozilla.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\firefox.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\fr-FR\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\fr-FR\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\it-IT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\it-IT\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\mozilla.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\firefox.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\it-IT\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\it-IT\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ru-RU")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ru-RU\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\mozilla.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\firefox.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\ru-RU\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\ru-RU\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\zh-CN")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\zh-CN\firefox.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\mozilla.adml" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\firefox.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$PSScriptRoot\$Product\windows\zh-CN\firefox.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Move-Item -Path "$PSScriptRoot\$Product\windows\zh-CN\mozilla.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\mac" -Force -Recurse -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\README.md" -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\LICENSE" -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$PSScriptRoot\$Product\windows" -Force -Recurse -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Download of the new ADMX files version $VersionP finished!"
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Nmap
    If ($Nmap -eq 1) {
        $Product = "Nmap"
        $PackageName = "Nmap-setup"
        $NMapD = Get-NevergreenApp -Name NMap | Where-Object { $_.Architecture -eq "x86" -and $_.Type -eq "exe" }
        $Version = $NMapD.Version
        $URL = $NMapD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Get-ChildItem "$PSScriptRoot\$Product\" -Exclude lang | Remove-Item -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "NotePadPlusPlus_" + "$NotepadPlusPlusArchitectureClear"
        $NotepadD = Get-EvergreenApp -Name NotepadPlusPlus | Where-Object { $_.Architecture -eq "$NotepadPlusPlusArchitectureClear" -and $_.Type -eq "exe" }
        $Version = $NotepadD.Version
        $URL = $NotepadD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$NotepadPlusPlusArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $NotepadPlusPlusArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Get-ChildItem "$PSScriptRoot\$Product\" -Exclude lang | Remove-Item -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $NotepadPlusPlusArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "OpenJDK" + "$openJDKArchitectureClear"
        $OpenJDKD = Get-EvergreenApp -Name OpenJDK | Where-Object { $_.Architecture -eq "$openJDKArchitectureClear" -and $_.URI -like "*msi*" } | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $OpenJDKD.Version
        $URL = $OpenJDKD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$openJDKArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $openJDKArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $openJDKArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Open-Shell Menu
    If ($OpenShellMenu -eq 1) {
        $Product = "Open-Shell Menu"
        $PackageName = "OpenShellSetup"
        $OpenShellMenuD = Get-EvergreenApp -Name OpenShellMenu
        $Version = $OpenShellMenuD.Version
        $URL = $OpenShellMenuD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "OracleJava8_" + "$OracleJava8ArchitectureClear"
        $OracleJava8D = Get-EvergreenApp -Name OracleJava8 | Where-Object { $_.Architecture -eq "$OracleJava8ArchitectureClear" }
        $Version = $OracleJava8D.Version
        $URL = $OracleJava8D.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$OracleJava8ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $OracleJava8ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $OracleJava8ArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Paint.Net
    If ($PaintDotNet -eq 1) {
        $Product = "Paint Dot Net"
        $PackageName = "Paint.net"
        $PaintDotNetD = Get-EvergreenApp -Name PaintDotNet | Where-Object { $_.URI -like "*files*" }
        $Version = $PaintDotNetD.Version
        $URL = $PaintDotNetD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product" $Source -includeStats
                expand-archive -path "$PSScriptRoot\$Product\Paint.Net.zip" -destinationpath "$PSScriptRoot\$Product"
                Move-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\$Product\paint.net.install.exe"
                Remove-Item -Path "$PSScriptRoot\$Product\Paint.Net.zip" -Force
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download pdfforge PDFCreator
    If ($PDFForgeCreator -eq 1) {
        $Product = "pdfforge PDFCreator"
        $PackageName = "PDFForgeCreatorWebSetup_" + "$pdfforgePDFCreatorChannelClear"
        $PDFForgeCreatorD = Get-PDFForgePDFCreator | Where-Object { $_.Channel -eq "$pdfforgePDFCreatorChannelClear" }
        $Version = $PDFForgeCreatorD.Version
        $URL = $PDFForgeCreatorD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$pdfforgePDFCreatorChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $pdfforgePDFCreatorChannelClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $pdfforgePDFCreatorChannelClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download PDF Split & Merge
    If ($PDFsam -eq 1) {
        $Product = "PDF Split & Merge"
        $PackageName = "PDFsam"
        $PDFsamD = Get-PDFsam
        $Version = $PDFsamD.Version
        $URL = $PDFsamD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download PeaZip
    If ($PeaZip -eq 1) {
        $Product = "PeaZip"
        $PackageName = "PeaZip" + "$PeaZipArchitectureClear"
        $PeaZipD = Get-EvergreenApp -Name PeaZipPeaZip | Where-Object {$_.Architecture -eq "$PeaZipArchitectureClear" -and $_.Type -eq "exe"}
        $Version = $PeaZipD.Version
        $URL = $PeaZipD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$PeaZipArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $PeaZipArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $PeaZipArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download PuTTY
    If ($Putty -eq 1) {
        $Product = "PuTTY"
        $PackageName = "PuTTY-" + "$PuTTYArchitectureClear" + "-$PuttyChannelClear"
        $PuTTYD = Get-Putty | Where-Object { $_.Architecture -eq "$PuTTYArchitectureClear" -and $_.Channel -eq "$PuttyChannelClear"}
        $Version = $PuTTYD.Version
        $URL = $PuTTYD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$PuTTYArchitectureClear" + "_$PuttyChannelClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $PuttyChannelClear $PuTTYArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $PuttyChannelClear $PuTTYArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
                $RemoteDesktopManagerFreeD = $Version
                $URL = "https://cdn.devolutions.net/download/Setup.RemoteDesktopManagerFree.$Version.msi"
                Add-Content -Path "$FWFile" -Value "$URL"
                $InstallerType = "msi"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $CurrentVersion"
                If ($CurrentVersion -lt $Version) {
                    Write-Host -ForegroundColor Green "Update available"
                    If ($WhatIf -eq '0') {
                        If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                        $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                        If ($Repository -eq '1') {
                            If ($CurrentVersion) {
                                Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                                Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                                Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                            }
                        }
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                        Start-Transcript $LogPS | Out-Null
                        Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    }
                    Write-Host "Starting download of $Product $Version"
                    If ($WhatIf -eq '0') {
                        Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                        Write-Verbose "Stop logging"
                        Stop-Transcript | Out-Null
                    }
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
                $RemoteDesktopManagerEnterpriseD = $Version
                $URL = "https://cdn.devolutions.net/download/Setup.RemoteDesktopManager.$Version.msi"
                Add-Content -Path "$FWFile" -Value "$URL"
                $InstallerType = "msi"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $CurrentVersion"
                If ($CurrentVersion -lt $Version) {
                    Write-Host -ForegroundColor Green "Update available"
                    If ($WhatIf -eq '0') {
                        If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                        $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                        If ($Repository -eq '1') {
                            If ($CurrentVersion) {
                                Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                                Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                                Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                            }
                        }
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                        Start-Transcript $LogPS | Out-Null
                        Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    }
                    Write-Host "Starting download of $Product $Version"
                    If ($WhatIf -eq '0') {
                        Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                        Write-Verbose "Stop logging"
                        Stop-Transcript | Out-Null
                    }
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

    #// Mark: Download Remote Display Analyzer
    If ($RDAnalyzer -eq 1) {
        $Product = "Remote Display Analyzer"
        $PackageName = "RDAnalyzer-setup"
        $RDAnalyzerD = Get-EvergreenApp -Name RDAnalyzer | Where-Object {$_.Type -eq "exe"}
        $Version = $RDAnalyzerD.Version
        $URL = $RDAnalyzerD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download ShareX
    If ($ShareX -eq 1) {
        $Product = "ShareX"
        $PackageName = "ShareX-setup"
        $ShareXD = Get-EvergreenApp -Name ShareX | Where-Object {$_.Type -eq "exe"}
        $Version = $ShareXD.Version
        $URL = $ShareXD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "Slack.setup" + "_$SlackArchitectureClear" + "_$SlackPlatformClear"
        $SlackD = Get-EvergreenApp -Name Slack | Where-Object {$_.Architecture -eq "$SlackArchitectureClear" -and $_.Platform -eq "$SlackPlatformClear" }
        $Version = $SlackD.Version
        $URL = $SlackD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$SlackArchitectureClear" + "_$SlackPlatformClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $SlackArchitectureClear $SlackPlatformClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $SlackArchitectureClear $SlackPlatformClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Sumatra PDF
    If ($SumatraPDF -eq 1) {
        $Product = "Sumatra PDF"
        $PackageName = "SumatraPDF-Install-" + "$SumatraPDFArchitectureClear"
        $SumatraPDFD = Get-EvergreenApp -Name SumatraPDFReader | Where-Object {$_.Architecture -eq "$SumatraPDFArchitectureClear" }
        $Version = $SumatraPDFD.Version
        #$URL = $SumatraPDFD.uri
        If ($SumatraPDFArchitectureClear -eq 'x86') {
            $URL = "https://kjkpubsf.sfo2.digitaloceanspaces.com/software/sumatrapdf/rel/SumatraPDF-" + "$Version" + "-install.exe"
        }
        If ($SumatraPDFArchitectureClear -eq 'x64') {
            $URL = "https://kjkpubsf.sfo2.digitaloceanspaces.com/software/sumatrapdf/rel/SumatraPDF-" + "$Version" + "-64-install.exe"
        }
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$SumatraPDFArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $SumatraPDFArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $SumatraPDFArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download TeamViewer
    If ($TeamViewer -eq 1) {
        $Product = "TeamViewer"
        $PackageName = "TeamViewer-setup"
        $TeamViewerD = Get-EvergreenApp -Name TeamViewer
        $Version = $TeamViewerD.Version
        $URL = $TeamViewerD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download TechSmith Camtasia
    If ($TechSmithCamtasia -eq 1) {
        $Product = "TechSmith Camtasia"
        $PackageName = "camtasia-setup"
        $TechSmithCamtasiaD = Get-EvergreenApp -Name TechSmithCamtasia | Where-Object { $_.Type -eq "msi" }
        $Version = $TechSmithCamtasiaD.Version
        $URL = $TechSmithCamtasiaD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download TechSmith Snagit
    If ($TechSmithSnagit -eq 1) {
        $Product = "TechSmith Snagit"
        $PackageName = "snagit-setup" + "_$TechSmithSnagItArchitectureClear"
        $TechSmithSnagitD = Get-EvergreenApp -Name TechSmithSnagit | Where-Object { $_.Architecture -eq "$TechSmithSnagItArchitectureClear" -and $_.Type -eq "msi" }
        $Version = $TechSmithSnagitD.Version
        $URL = $TechSmithSnagitD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$TechSmithSnagItArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $TechSmithSnagItArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $TechSmithSnagItArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Total Commander
    If ($TotalCommander -eq 1) {
        $Product = "Total Commander"
        $PackageName = "TotalCommander_" + "$TotalCommanderArchitectureClear"
        $TotalCommanderD = Get-TotalCommander | Where-Object {$_.Architecture -eq "$TotalCommanderArchitectureClear"}
        $Version = $TotalCommanderD.Version
        $URL = $TotalCommanderD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$TotalCommanderArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $TotalCommanderArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $TotalCommanderArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
                Add-Content -Path "$FWFile" -Value "$URL"
                $InstallerType = "exe"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $CurrentVersion"
                If ($CurrentVersion -lt $Version) {
                    Write-Host -ForegroundColor Green "Update available"
                    If ($WhatIf -eq '0') {
                        If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                        $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                        If ($Repository -eq '1') {
                            If ($CurrentVersion) {
                                Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                                Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                                Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                            }
                        }
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                        Start-Transcript $LogPS | Out-Null
                        Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    }
                    Write-Host "Starting download of $Product $Version"
                    If ($WhatIf -eq '0') {
                        Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                        Write-Verbose "Stop logging"
                        Stop-Transcript | Out-Null
                    }
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
                Add-Content -Path "$FWFile" -Value "$URL"
                $InstallerType = "exe"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Host -ForegroundColor Magenta "Download $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $CurrentVersion"
                If ($CurrentVersion -lt $Version) {
                    Write-Host -ForegroundColor Green "Update available"
                    If ($WhatIf -eq '0') {
                        If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                        $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                        If ($Repository -eq '1') {
                            If ($CurrentVersion) {
                                Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                                If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                                Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                                Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                            }
                        }
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                        Start-Transcript $LogPS | Out-Null
                        Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                    }
                    Write-Host "Starting download of $Product $Version"
                    If ($WhatIf -eq '0') {
                        Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                        Write-Verbose "Stop logging"
                        Stop-Transcript | Out-Null
                    }
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

    #// Mark: Download uberAgent
    If ($uberAgent -eq 1) {
        $Product = "uberAgent"
        $PackageName = "setup_uberAgent"
        $uberAgentD = Get-EvergreenApp -Name VastLimitsUberAgent 
        $Version = $uberAgentD.Version
        $URL = $uberAgentD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse -Exclude silent-install.cmd
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\" + ($Source))
                expand-archive -path "$PSScriptRoot\$Product\$Source" -destinationpath "$PSScriptRoot\$Product"
                Remove-Item -Path "$PSScriptRoot\$Product\$Source" -Force
                Move-Item -Path "$PSScriptRoot\$Product\uberAgent components\uberAgent_endpoint\bin\uberAgent-32.msi" -Destination "$PSScriptRoot\$Product"
                Move-Item -Path "$PSScriptRoot\$Product\uberAgent components\uberAgent_endpoint\bin\uberAgent-64.msi" -Destination "$PSScriptRoot\$Product"
                If (!(Test-Path "$PSScriptRoot\$Product\silent-install.cmd" -PathType leaf)) {
                    Move-Item -Path "$PSScriptRoot\$Product\uberAgent components\uberAgent_endpoint\bin\silent-install.cmd" -Destination "$PSScriptRoot\$Product"
                }
                #Remove-Item -Path "$PSScriptRoot\$Product\uberAgent components" -Force -Recurse
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
            Write-Host "Starting copy of $Product ADMX files $Version"
            If ($WhatIf -eq '0') {
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\uberAgent.admx" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse
                }
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                copy-item -Path "$PSScriptRoot\$Product\uberAgent components\Group Policy\Administrative template (ADMX)\*" -Destination "$PSScriptRoot\_ADMX\$Product" -Force
            }
            Write-Host -ForegroundColor Green "Copy of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download VLC Player
    If ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        $PackageName = "VLC-Player_" + "$VLCPlayerArchitectureClear"
        $VLCD = Get-EvergreenApp -Name VideoLanVlcPlayer | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "$VLCPlayerArchitectureClear" -and $_.Type -eq "MSI" }
        $Version = $VLCD.Version
        $URL = $VLCD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$VLCPlayerArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $VLCPlayerArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $VLCPlayerArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        $PackageName = "VMWareTools_" + "$VMWareToolsArchitectureClear"
        $VMWareToolsD = Get-EvergreenApp -Name VMwareTools | Where-Object { $_.Architecture -eq "$VMWareToolsArchitectureClear" }
        $Version = $VMWareToolsD.Version
        $URL = $VMWareToolsD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$VMWareToolsArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $VMWareToolsArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $VMWareToolsArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download WinMerge
    If ($WinMerge -eq 1) {
        $Product = "WinMerge"
        $PackageName = "WinMerge_" + "$WinMergeArchitectureClear"
        $WinMergeD = Get-EvergreenApp -Name WinMerge | Where-Object {$_.Architecture -eq "$WinMergeArchitectureClear" -and $_.URI -notlike "*PerUser*"}
        $Version = $WinMergeD.Version
        $URL = $WinMergeD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$WinMergeArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $WinMergeArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $WinMergeArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
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
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
            }
            Write-Host "Starting download of $Product $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Wireshark
    If ($Wireshark -eq 1) {
        $Product = "Wireshark"
        $PackageName = "Wireshark-" + "$WiresharkArchitectureClear"
        $WiresharkD = Get-EvergreenApp -Name Wireshark | Where-Object { $_.Architecture -eq "$WiresharkArchitectureClear" -and $_.Type -eq "exe"}
        $Version = $WiresharkD.Version
        $URL = $WiresharkD.uri
        Add-Content -Path "$FWFile" -Value "$URL"
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$WiresharkArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Host -ForegroundColor Magenta "Download $Product $WiresharkArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CurrentVersion"
        If ($CurrentVersion -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            If ($WhatIf -eq '0') {
                If (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                If ($Repository -eq '1') {
                    If ($CurrentVersion) {
                        Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                        If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                        Copy-Item -Path "$PSScriptRoot\$Product\*.exe" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                        Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                    }
                }
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                Start-Transcript $LogPS | Out-Null
                Set-Content -Path "$VersionPath" -Value "$Version"
            }
            Write-Host "Starting download of $Product $WiresharkArchitectureClear $Version"
            If ($WhatIf -eq '0') {
                Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                Write-Verbose "Stop logging"
                Stop-Transcript | Out-Null
            }
            Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
            Write-Output ""
        }
        Else {
            Write-Host -ForegroundColor Cyan "No new version available"
            Write-Output ""
        }
    }

    #// Mark: Download Zoom
    If ($Zoom -eq 1) {
        If ($ZoomInstallerClear -eq 'Machine Based') {
            $Product = "Zoom VDI"
            $PackageName = "ZoomInstaller"
            $ZoomD = Get-EvergreenApp -Name Zoom | Where-Object {$_.Platform -eq "VDI"}
            $URLVersion = "https://support.zoom.us/hc/en-us/articles/360041602711"
            $webRequest = Invoke-WebRequest -UseBasicParsing -Uri ($URLVersion) -SessionVariable websession
            $regexAppVersion = "(\d\.\d\.\d)"
            $Version = $webRequest.RawContent | Select-String -Pattern $regexAppVersion -AllMatches | ForEach-Object { $_.Matches.Value } | Sort-Object -Descending | Select-Object -First 1
            $ZoomVersion = $Version
            $URL = $ZoomD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerType = "msi"
            $Source = "$PackageName" + "." + "$InstallerType"
            $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
            Write-Host -ForegroundColor Magenta "Download $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion"
            If ($CurrentVersion -lt $Version) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                }
                Write-Host "Starting download of $Product $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
                $PackageNameP = "ZoomADMX"
                $ZoomDP = Get-ZoomAdmx
                $VersionP = $ZoomDP.version
                $URL = $ZoomDP.uri
                Add-Content -Path "$FWFile" -Value "$URL"
                $InstallerTypeP = "zip"
                $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
                $FolderP = "Zoom_" + + "$VersionP"
                Write-Host "Starting download of $Product ADMX files $VersionP"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                    expand-archive -path "$PSScriptRoot\$Product\$SourceP" -destinationpath "$PSScriptRoot\$Product"
                    Remove-Item -Path "$PSScriptRoot\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                    If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                    If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\ZoomMeetings_HKCU.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\ZoomMeetings_HKLM.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\ZoomMeetingsGlobalPolicy.reg" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                    If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                    If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\ZoomMeetings_HKCU.adml" -PathType leaf)) {
                        Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\ZoomMeetings_HKCU.adml" -ErrorAction SilentlyContinue
                        Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\ZoomMeetings_HKLM.adml" -ErrorAction SilentlyContinue
                    }
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\en-US\ZoomMeetings_HKCU.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\en-US\ZoomMeetings_HKLM.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\$Product\$FolderP" -Force -Recurse -ErrorAction SilentlyContinue
                }
                Write-Host -ForegroundColor Green "Download of the new ADMX files version $VersionP finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
        If ($ZoomInstallerClear -eq 'User Based') {
            $Product = "Zoom"
            $PackageName = "ZoomInstaller"
            $ZoomD = Get-EvergreenApp -Name Zoom | Where-Object {$_.Type -eq "Msi"}
            $Version = $ZoomD.version
            $URL = $ZoomD.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $InstallerType = "msi"
            $Source = "$PackageName" + "." + "$InstallerType"
            $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
            Write-Host -ForegroundColor Magenta "Download $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion"
            If ($CurrentVersion -lt $Version) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product\Version.txt" -Value "$Version"
                }
                Write-Host "Starting download of $Product $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $Source -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version finished!"
                Write-Output ""
                $PackageNameP = "ZoomADMX"
                $ZoomDP = Get-ZoomAdmx
                $URL = $ZoomDP.uri
                Add-Content -Path "$FWFile" -Value "$URL"
                $InstallerTypeP = "zip"
                $SourceP = "$PackageNameP" + "." + "$InstallerTypeP"
                $FolderP = "Zoom_" + "$Version"
                Write-Host "Starting download of $Product ADMX files $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product\" $SourceP -includeStats
                    expand-archive -path "$PSScriptRoot\$Product\$SourceP" -destinationpath "$PSScriptRoot\$Product"
                    Remove-Item -Path "$PSScriptRoot\$Product\$SourceP" -Force -ErrorAction SilentlyContinue
                    If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                    If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\ZoomMeetings_HKCU.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\ZoomMeetings_HKLM.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\ZoomMeetingsGlobalPolicy.reg" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                    If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                    If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\ZoomMeetings_HKCU.adml" -PathType leaf)) {
                        Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\ZoomMeetings_HKCU.adml" -ErrorAction SilentlyContinue
                        Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\ZoomMeetings_HKLM.adml" -ErrorAction SilentlyContinue
                    }
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\en-US\ZoomMeetings_HKCU.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                    Move-Item -Path "$PSScriptRoot\$Product\$FolderP\en-US\ZoomMeetings_HKLM.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                    Remove-Item -Path "$PSScriptRoot\$Product\$FolderP" -Force -Recurse -ErrorAction SilentlyContinue
                }
                Write-Host -ForegroundColor Green "Download of the new ADMX files version $VersionP finished!"
                Write-Output ""
            }
            Else {
                Write-Host -ForegroundColor Cyan "No new version available"
                Write-Output ""
            }
        }
        If ($ZoomCitrixClient -eq 1) {
            $Product2 = "Zoom Citrix Client"
            $PackageName2 = "ZoomCitrixHDXMediaPlugin"
            $ZoomCitrix = Get-EvergreenApp -Name Zoom | Where-Object {$_.Platform -eq "Citrix"}
            $URL = $ZoomCitrix.uri
            Add-Content -Path "$FWFile" -Value "$URL"
            $Source2 = "$PackageName2" + "." + "$InstallerType"
            $CurrentVersion2 = Get-Content -Path "$PSScriptRoot\$Product2\Version.txt" -EA SilentlyContinue
            Write-Host -ForegroundColor Magenta "Download $Product2" -Verbose
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $CurrentVersion2"
            If (!($CurrentVersion2 -lt $Version)) {
                Write-Host -ForegroundColor Green "Update available"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path "$PSScriptRoot\$Product2")) {New-Item -Path "$PSScriptRoot\$Product2" -ItemType Directory | Out-Null}
                    $LogPS = "$PSScriptRoot\$Product2\" + "$Product2 $Version.log"
                    If ($Repository -eq '1') {
                        If ($CurrentVersion) {
                            Write-Host "Copy $Product Installer Version $CurrentVersion to Repository folder"
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product")) { New-Item -Path "$PSScriptRoot\_Repository\$Product" -ItemType Directory | Out-Null }
                            If (!(Test-Path -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion")) { New-Item -Path "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ItemType Directory | Out-Null }
                            Copy-Item -Path "$PSScriptRoot\$Product\*.msi" -Destination "$PSScriptRoot\_Repository\$Product\$CurrentVersion" -ErrorAction SilentlyContinue
                            Write-Host -ForegroundColor Green "Copy of the current version $CurrentVersion finished!"
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product2\*" -Recurse
                    Start-Transcript $LogPS | Out-Null
                    Set-Content -Path "$PSScriptRoot\$Product2\Version.txt" -Value "$Version"
                }
                Write-Host "Starting download of $Product2 $Version"
                If ($WhatIf -eq '0') {
                    Get-Download $URL "$PSScriptRoot\$Product2\" $Source2 -includeStats
                    Write-Verbose "Stop logging"
                    Stop-Transcript | Out-Null
                }
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

If ($Install -eq "1") {

    If ($Installer -eq 0) {
        Write-Host "Change User Mode to Install."
        Write-Output ""
        Change User /Install | Out-Null
    }

    Write-Host -ForegroundColor DarkGray "Starting installs..."
    Write-Output ""

    # Logging
    # Global variables
    # $StartDir = $PSScriptRoot # the directory path of the script currently being executed
    $LogDir = "$PSScriptRoot\_Install Logs"
    $LogFileName = ("$ENV:COMPUTERNAME - $Date.log")
    $LogFile = Join-path $LogDir $LogFileName
    $FWFileName = ("Firewall - $Date.log")
    $FWFile = Join-path $LogDir $FWFileName
    $LogTemp = "$env:windir\Logs\Evergreen"

    # Create the log directories if they don't exist
    If (!(Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType directory | Out-Null }
    If (!(Test-Path $LogTemp)) { New-Item -Path $LogTemp -ItemType directory | Out-Null }

    # Create new log file (overwrite existing one)
    New-Item $LogFile -ItemType "file" -force | Out-Null
    DS_WriteLog "I" "START SCRIPT - " $LogFile
    DS_WriteLog "-" "" $LogFile

    # Install script part (AddScript)

    If ($WhatIf -eq '1') {
        Write-Host -BackgroundColor Magenta "WhatIf Mode, nothing will be installed !!!"
        Write-Host -BackgroundColor Green "The Install log will be created !!!"
        Write-Output ""
    }

    #// Mark: Install 1Password
    If ($1Password -eq 1) {
        $Product = "1Password"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $1PasswordD.Version
        }
        $1PasswordV = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*1Password*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$1PasswordV) {
            $1PasswordV = (Get-ItemProperty HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -like "*1Password*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $1PasswordInstaller = "1Password-Setup.exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $1PasswordV"
        If ($1PasswordV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$1PasswordInstaller" -ArgumentList --Silent
                }
                $p = Get-Process 1Password-Setup -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\1Password.lnk") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\1Password.lnk" -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install 7 Zip
    If ($7ZIP -eq 1) {
        $Product = "7-Zip"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$7ZipArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $7ZipD.Version
        }
        $SevenZip = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$SevenZip) {
            $SevenZip = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $7ZipInstaller = "7-Zip_" + "$7ZipArchitectureClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $7ZipArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $SevenZip"
        If ($SevenZip -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $7ZipArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$7ZipInstaller" -ArgumentList /S
                }
                $p = Get-Process 7-Zip_$7ZipArchitectureClear -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\7-Zip") {Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\7-Zip" -Recurse -Force}
                    }
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $7ZipArchitectureClear (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Adobe Pro DC
    If ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $AdobeProD.Version
        }
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$Adobe) {
            $Adobe = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $Adobe"
        If ($Adobe -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                $mspArgs = "/P `"$PSScriptRoot\$Product\Adobe_Pro_DC_Update.msp`" /quiet /qn"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath msiexec.exe -ArgumentList $mspArgs -Wait
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
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
                    If ($WhatIf -eq '0') {
                        Stop-Service AdobeARMservice
                        Set-Service AdobeARMservice -StartupType Disabled
                    }
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                Write-Host "Customize Scheduled Task"
                If ($WhatIf -eq '0') {
                    Get-ScheduledTask -TaskName Adobe* -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
                }
                Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Adobe Reader DC
    If ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$AdobeArchitectureClear" + "_$AdobeLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $AdobeReaderD.Version
        }
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Adobe Acrobat*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$Adobe) {
            $Adobe = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Adobe Acrobat*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $AdobeReaderInstaller = "Adobe_Reader_DC_" + "$AdobeArchitectureClear" + "$AdobeLanguageClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $AdobeArchitectureClear $AdobeLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $Adobe"
        If ($Adobe -lt $Version) {
            DS_WriteLog "I" "Installing $Product $AdobeArchitectureClear $AdobeLanguageClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/sAll"
                "/rs"
                "/msi EULA_ACCEPT=YES ENABLE_OPTIMIZATION=YES DISABLEDESKTOPSHORTCUT=1 UPDATE_MODE=0 DISABLE_ARM_SERVICE_INSTALL=1 DISABLE_CACHE=1 DISABLE_PDFMAKER=YES ALLUSERS=1"
            )
            Try {
                Write-Host "Starting install of $Product $AdobeArchitectureClear $AdobeLanguageClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$AdobeReaderInstaller" -ArgumentList $Options
                }
                $p = Get-Process Adobe_Reader_DC_$AdobeArchitectureClear$AdobeLanguageClear -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Acrobat DC.lnk") {Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Acrobat DC.lnk" -Force}
                    }
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
                    If ($WhatIf -eq '0') {
                        Stop-Service AdobeARMservice
                        Set-Service AdobeARMservice -StartupType Disabled
                    }
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error customizing (Error: $($Error[0]))"
                DS_WriteLog "E" "Error customizing (Error: $($Error[0]))" $LogFile
            }
            Try {
                Write-Host "Customize Scheduled Task"
                If ($WhatIf -eq '0') {
                    Get-ScheduledTask -TaskName Adobe* -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
                }
                Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Autodesk DWG TrueView
    If ($AutodeskDWGTrueView -eq 1) {
        $Product = "Autodesk DWG TrueView"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $AutodeskDWGTrueViewD.Version
        }
        $DWGTrueView = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*DWG TrueView*"}).DisplayName | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$DWGTrueView) {
            $DWGTrueView = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*DWG TrueView*"}).DisplayName | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($DWGTrueView) {$DWGTrueViewV = $DWGTrueView.split(" ")[2]}
        $AutodeskDWGTrueViewInstaller = "AutodeskDWGTrueView.exe"
        $InstallMSI = "$PSScriptRoot\$Product\DWGTrueView_2022_English_64bit_dlm\x64\dwgviewr\DWGVIEWR.msi"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $DWGTrueViewV"
        If ($DWGTrueViewV -lt $Version) {
            DS_WriteLog "I" "Unpacking $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "-suppresslaunch"
                "-d"
                "`"$PSScriptRoot\$Product`""
            )
            Try {
                Write-Host "Starting unpacking of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$AutodeskDWGTrueViewInstaller" -ArgumentList $Options
                }
                $p = Get-Process AutodeskDWGTrueView -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Unpacking of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error unpacking $Product (Error: $($Error[0]))"
                DS_WriteLog "E" "Error unpacking $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "I" "Installing $Product" $LogFile
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qb ADSK_EULA_STATUS=#1 REBOOT=ReallySuppress ADSK_SETUP_EXE=1"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    If (Test-Path -Path "$env:PUBLIC\Desktop\DWG TrueView*.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\DWG TrueView*.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\DWG TrueView*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\DWG TrueView*" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install BIS-F
    If ($BISF -eq 1) {
        $Product = "BIS-F"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $BISFD.Version
        }
        $BISFV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Base Image*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$BISFV) {
            $BISFV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Base Image*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $BISFLog = "$LogTemp\BISF.log"
        $InstallMSI = "$PSScriptRoot\$Product\setup-BIS-F.msi"
        Write-Host -ForegroundColor Magenta "Install $Product"
        If ($BISFV) {$BISFV = $BISFV -replace ".{6}$"}
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $BISFV"
        If ($BISFV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $BISFLog"
                "/norestart"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
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
                    If ($WhatIf -eq '0') {
                        ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw -ErrorAction SilentlyContinue) -replace "DisableTaskOffload' -Value '1'","DisableTaskOffload' -Value '0'") | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -ErrorAction SilentlyContinue
                        ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw -ErrorAction SilentlyContinue) -replace 'nx AlwaysOff','nx OptOut') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -ErrorAction SilentlyContinue
                        ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw -ErrorAction SilentlyContinue) -replace 'rss=disable','rss=enable') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -ErrorAction SilentlyContinue
                    }
                    Write-Host -ForegroundColor Green "Customize scripts $Product finished!"
                } Catch {
                    Write-Host -ForegroundColor Red "Error when customizing scripts (Error: $($Error[0]))"
                    DS_WriteLog "E" "Error when customizing scripts (Error: $($Error[0]))" $LogFile
                }
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            Write-Host "Starting copy of $Product ADMX files $Version"
            $BISFInstallFolder = "${env:ProgramFiles(x86)}\Base Image Script Framework (BIS-F)\ADMX"
            If ($WhatIf -eq '0') {
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\BaseImageScriptFramework.admx" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse -ErrorAction SilentlyContinue
                }
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$BISFInstallFolder\BaseImageScriptFramework.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\BaseImageScriptFramework.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\BaseImageScriptFramework.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$BISFInstallFolder\en-US\BaseImageScriptFramework.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Copy of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Cisco Webex Teams
    If ($CiscoWebexTeams -eq 1) {
        $Product = "Cisco Webex Teams"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$CiscoWebexTeamsArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $WebexTeamsD.Version
        }
        $CiscoWebexTeamsV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Webex"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$CiscoWebexTeamsV) {
            $CiscoWebexTeamsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Webex"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $WebexLog = "$LogTemp\WebexTeams.log"
        $InstallMSI = "$PSScriptRoot\$Product\webexteams-" + "$CiscoWebexTeamsArchitectureClear" + ".msi"
        Write-Host -ForegroundColor Magenta "Install $Product $CiscoWebexTeamsArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CiscoWebexTeamsV"
        If ($CiscoWebexTeamsV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/quiet"
                "/L*V $WebexLog"
                "ACCEPT_EULA=TRUE ALLUSERS=1 AUTOSTART_WITH_WINDOWS=false"
            )
            Try {
                Write-Host "Starting install of $Product $CiscoWebexTeamsArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Webex") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Webex" -Recurse -Force}
                    }
                }
                Get-Content $WebexLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $WebexLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }
    
    #// Mark: Install Citrix Files
    If ($CitrixFiles -eq 1) {
        $Product = "Citrix Files"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $CitrixFilesD.Version
        }
        $CitrixFilesV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Files*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $CitrixFilesLog = "$LogTemp\CitrixFiles.log"
        $CitrixFilesInstaller = "CitrixFilesForWindows.msi"
        $InstallMSI = "$PSScriptRoot\Citrix\$Product\$CitrixFilesInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product"
        If (!$CitrixFilesV) {
            $CitrixFilesV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Files*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $CitrixFilesV"
        If ($CitrixFilesV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/quiet"
                "/norestart"
                "/L*V $CitrixFilesLog"
                )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Citrix") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Citrix" -Recurse -Force}
                    }
                }
                Get-Content $CitrixFilesLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $CitrixFilesLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Citrix Hypervisor Tools
    If ($Citrix_Hypervisor_Tools -eq 1) {
        $Product = "Citrix Hypervisor Tools"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version_" + "$CitrixHypervisorToolsArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $CitrixHypervisor.Version
        }
        $HypTools = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Hypervisor*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $CitrixHypLog = "$LogTemp\CitrixHypervisor.log"
        $HypToolsInstaller = "managementagent" + "$CitrixHypervisorToolsArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\Citrix\$Product\$HypToolsInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $CitrixHypervisorToolsArchitectureClear"
        If (!$HypTools) {
            $HypTools = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Hypervisor*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($HypTools) {$HypTools = $HypTools.Insert(3,'.0')}
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $HypTools"
        If ($HypTools -lt $Version) {
            DS_WriteLog "I" "Installing $Product $CitrixHypervisorToolsArchitectureClear" $LogFile
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
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $CitrixHypLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $CitrixHypLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Citrix WorkspaceApp
    If ($Citrix_WorkspaceApp -eq 1) {
        $Product = "Citrix WorkspaceApp $CitrixWorkspaceAppReleaseClear"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $WSACD.Version
        }
        $WSA = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Workspace*" -and $_.UninstallString -like "*Trolley*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$WSA) {
            $WSA = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Workspace*" -and $_.UninstallString -like "*Trolley*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $UninstallWSACR = "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $WSA"
        If ($WSA -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            # Citrix WSA Uninstallation
            Write-Host "Uninstall Citrix Workspace App / Receiver"
            DS_WriteLog "I" "Uninstall Citrix Workspace App / Receiver" $LogFile
            Try {
                If ($WhatIf -eq '0') {
                    Start-process $UninstallWSACR -ArgumentList '/silent /disableCEIP' -NoNewWindow -Wait
                }
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
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\Citrix\$Product\CitrixWorkspaceApp.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                Write-Host "Customize $Product"
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Citrix Workspace.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Citrix Workspace.lnk" -Force}
                    }
                    reg add "HKLM\SOFTWARE\Wow6432Node\Policies\Citrix" /v EnableX1FTU /t REG_DWORD /d 0 /f | Out-Null
                    reg add "HKCU\Software\Citrix\Splashscreen" /v SplashscreenShown /d 1 /f | Out-Null
                    reg add "HKLM\SOFTWARE\Policies\Citrix" /f /v EnableFTU /t REG_DWORD /d 0 | Out-Null
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install ControlUp Agent
    If ($ControlUpAgent -eq 1) {
        $Product = "ControlUp Agent"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ControlUpAgentFrameworkClear" + "_$ControlUpAgentArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $ControlUpAgentD.Version
        }
        $ControlUpAgentV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "ControlUpAgent"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$ControlUpAgentV) {
            $ControlUpAgentV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "ControlUpAgent"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $ControlUpAgentLog = "$LogTemp\ControlUpAgent.log"
        $ControlUpAgentInstaller = "ControlUpAgent-" + "$ControlUpAgentFrameworkClear" + "-$ControlUpAgentArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$ControlUpAgentInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ControlUpAgentFrameworkClear $ControlUpAgentArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $ControlUpAgentV"
        If ($ControlUpAgentV -lt $Version) {
            DS_WriteLog "I" "Installing $Product $ControlUpAgentFrameworkClear $ControlUpAgentArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $ControlUpAgentLog"
                )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $ControlUpAgentLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $ControlUpAgentLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install deviceTRUST
    If ($deviceTRUST -eq 1) {
        $Product = "deviceTRUST"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + "_$deviceTRUSTArchitectureClear"+ ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $deviceTRUSTD.Version
        }
        $deviceTRUSTClientV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Client*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$deviceTRUSTClientV) {
            $deviceTRUSTClientV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Client*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($deviceTRUSTClientV.length -ne "8") {$deviceTRUSTClientV = $deviceTRUSTClientV -replace ".{2}$"}
        $deviceTRUSTHostV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Host*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$deviceTRUSTHostV) {
            $deviceTRUSTHostV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Host*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($deviceTRUSTHostV.length -ne "8") {$deviceTRUSTHostV = $deviceTRUSTHostV -replace ".{2}$"}
        $deviceTRUSTConsoleV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Console*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$deviceTRUSTConsoleV) {
            $deviceTRUSTConsoleV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*deviceTRUST Console*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($deviceTRUSTConsoleV.length -ne "8") {$deviceTRUSTConsoleV = $deviceTRUSTConsoleV -replace ".{2}$"}
        $deviceTRUSTLog = "$LogTemp\deviceTRUST.log"
        $deviceTRUSTClientLog = "$LogTemp\deviceTRUST.txt"
        $deviceTRUSTClientInstaller = "dtclient-release" + ".exe"
        $deviceTRUSTHostInstaller = "dthost-" + "$deviceTRUSTArchitectureClear" + "-release" + ".msi"
        $deviceTRUSTConsoleInstaller = "dtconsole-" + "$deviceTRUSTArchitectureClear" + "-release" + ".msi"
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
            Write-Host "Current Version:  $deviceTRUSTClientV"
            If ($deviceTRUSTClientV -lt $Version) {
                # deviceTRUST Client
                DS_WriteLog "I" "Installing $Product Client" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                Try {
                    $Options = @(
                        "/INSTALL"
                        "/QUIET"
                        "/NORESTART"
                        "/LOG $deviceTRUSTClientLog"
                    )
                    Write-Host "Starting install of $Product Client $Version"
                    If ($WhatIf -eq '0') {
                        Start-Process -FilePath "$PSScriptRoot\$Product\$deviceTRUSTClientInstaller" -ArgumentList $Options -PassThru -Wait -ErrorAction Stop | Out-Null
                        If ($CleanUpStartMenu) {
                            If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\DWG TrueView*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\DWG TrueView*" -Recurse -Force}
                        }
                    }
                    Get-Content $deviceTRUSTClientLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $deviceTRUSTClientLog -ErrorAction SilentlyContinue
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
            Write-Host "Current Version:  $deviceTRUSTHostV"
            If ($deviceTRUSTHostV -lt $Version) {
                # deviceTRUST Host
                DS_WriteLog "I" "Installing $Product Host" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $InstallMSI = "$PSScriptRoot\$Product\$deviceTRUSTHostInstaller"
                Try {
                    Write-Host "Starting install of $Product Host $Version"
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                    }
                    Get-Content $deviceTRUSTLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $deviceTRUSTLog -ErrorAction SilentlyContinue
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
            Write-Host "Current Version:  $deviceTRUSTConsoleV"
            If ($deviceTRUSTConsoleV -lt $Version) {
                # deviceTRUST Console
                DS_WriteLog "I" "Installing $Product Console" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $InstallMSI = "$PSScriptRoot\$Product\$deviceTRUSTConsoleInstaller"
                Try {
                    Write-Host "Starting install of $Product Console $Version"
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                    }
                    Get-Content $deviceTRUSTLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $deviceTRUSTLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Filezilla
    If ($Filezilla -eq 1) {
        $Product = "Filezilla"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $FilezillaD.Version
        }
        $FilezillaV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Filezilla*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$FilezillaV) {
            $FilezillaV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Filezilla*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $FilezillaV"
        If ($FilezillaV -lt $Version) {
            $Options = @(
                "/S"
                "/user=all"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\Filezilla-win64.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\FileZilla FTP Client") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\FileZilla FTP Client" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Foxit PDF Editor
    If ($FoxitPDFEditor -eq 1) {
        $Product = "Foxit PDF Editor"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitPDFEditorLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $FoxitPDFEditorD.Version
        }
        $FoxitPDFEditorV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Foxit PDF Editor"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$FoxitPDFEditorV) {
            $FoxitPDFEditorV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Foxit PDF Editor"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $FoxitPDFEditorLog = "$LogTemp\FoxitPDFEditor.log"
        $FoxitPDFEditorInstaller = "FoxitPDFEditor-Setup-" + "$FoxitPDFEditorLanguageClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$FoxitPDFEditorInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $FoxitPDFEditorLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $FoxitPDFEditorV"
        If ($FoxitPDFEditorV -lt $Version) {
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/quiet"
                "/L*V $FoxitPDFEditorLog"
                "/NORESTART"
                "AUTO_UPDATE=0 LAUNCHCHECKDEFAULT=0 DESKTOP_SHORTCUT=0"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $FoxitPDFEditorLanguageClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $FoxitPDFEditorLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $FoxitPDFEditorLog -ErrorAction SilentlyContinue
                If ($WhatIf -eq '0') {
                    If (Test-Path -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Foxit PDF Editor") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Foxit PDF Editor" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Foxit Reader
    If ($Foxit_Reader -eq 1) {
        $Product = "Foxit Reader"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitReaderLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $Foxit_ReaderD.Version
        }
        $FReader = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Foxit Reader*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$FReader) {
            $FReader = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Foxit Reader*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $FoxitLog = "$LogTemp\FoxitReader.log"
        $FoxitReaderInstaller = "FoxitReader-Setup-" + "$FoxitReaderLanguageClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$FoxitReaderInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $FoxitReaderLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $FReader"
        If ($FReader -lt $Version) {
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $FoxitLog"
                "/NORESTART"
                "AUTO_UPDATE=0 LAUNCHCHECKDEFAULT=0 DESKTOP_SHORTCUT=0"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $FoxitReaderLanguageClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $FoxitLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $FoxitLog -ErrorAction SilentlyContinue
                If ($WhatIf -eq '0') {
                    If (Test-Path -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Foxit PDF Reader") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Foxit PDF Reader" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install GIMP
    If ($GIMP -eq 1) {
        $Product = "GIMP"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $GIMPD.Version
        }
        $GIMPV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*GIMP*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$GIMPV) {
            $GIMPV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*GIMP*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $GIMPV"
        If ($GIMPV -ne $Version) {
            $Options = @(
                "/VERYSILENT"
                "/NORESTART"
                "/ALLUSERS"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\gimp-setup.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\GIMP*.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\GIMP*.lnk" -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Git for Windows
    If ($GitForWindows -eq 1) {
        $Product = "Git for Windows"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$GitForWindowsArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $GitForWindowsD.Version
        }
        $GitForWindowsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Git"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$GitForWindowsV) {
            $GitForWindowsV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Git"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $GitForWindowsInstaller = "GitForWindows_" + "$GitForWindowsArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $GitForWindowsArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $GitForWindowsV"
        If ($GitForWindowsV -lt $Version) {
            $Options = @(
                "/suppressmsgboxes"
                "/norestart"
                "/noicons"
                "/verysilent"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $GitForWindowsArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$GitForWindowsInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Git for Windows.lnk") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Git for Windows.lnk" -Force}
                    }
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $GitForWindowsArchitectureClear (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Google Chrome
    If ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$GoogleChromeArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $ChromeD.Version
        }
        $ChromeSplit = $Version.split(".")
        $ChromeStrings = ([regex]::Matches($Version, "\." )).count
        $ChromeStringLast = ([regex]::Matches($ChromeSplit[$ChromeStrings], "." )).count
        If ($ChromeStringLast -lt "3") {
            $ChromeSplit[$ChromeStrings] = "0" + $ChromeSplit[$ChromeStrings]
        }
        Switch ($ChromeStrings) {
            1 {
                $NewVersion = $ChromeSplit[0] + "." + $ChromeSplit[1]
            }
            2 {
                $NewVersion = $ChromeSplit[0] + "." + $ChromeSplit[1] + "." + $ChromeSplit[2]
            }
            3 {
                $NewVersion = $ChromeSplit[0] + "." + $ChromeSplit[1] + "." + $ChromeSplit[2] + "." + $ChromeSplit[3]
            }
        }
        $Chrome = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $ChromeLog = "$LogTemp\GoogleChrome.log"
        If (!$Chrome) {
            $Chrome = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $CurrentChromeSplit = $Chrome.split(".")
        $CurrentChromeStrings = ([regex]::Matches($Chrome, "\." )).count
        $CurrentChromeStringLast = ([regex]::Matches($CurrentChromeSplit[$CurrentChromeStrings], "." )).count
        If ($CurrentChromeStringLast -lt "3") {
            $CurrentChromeSplit[$CurrentChromeStrings] = "0" + $CurrentChromeSplit[$CurrentChromeStrings]
        }
        Switch ($CurrentChromeStrings) {
            1 {
                $NewCurrentVersion = $CurrentChromeSplit[0] + "." + $CurrentChromeSplit[1]
            }
            2 {
                $NewCurrentVersion = $CurrentChromeSplit[0] + "." + $CurrentChromeSplit[1] + "." + $CurrentChromeSplit[2]
            }
            3 {
                $NewCurrentVersion = $CurrentChromeSplit[0] + "." + $CurrentChromeSplit[1] + "." + $CurrentChromeSplit[2] + "." + $CurrentChromeSplit[3]
            }
        }
        $ChromeInstaller = "googlechromestandaloneenterprise_" + "$GoogleChromeArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$ChromeInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $GoogleChromeArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $Chrome"
        If ($NewCurrentVersion -lt $NewVersion) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $ChromeLog"
            )
            Try {
                Write-Host "Starting install of $Product $GoogleChromeArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $ChromeLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $ChromeLog -ErrorAction SilentlyContinue
                If ($WhatIf -eq '0') {
                    If (Test-Path -Path "$env:PUBLIC\Desktop\Google Chrome.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Google Chrome.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk" -Recurse -Force}
                    }
                }
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                # Disable update service and scheduled task
                $Service = Get-Service -Name gupdate -ErrorAction SilentlyContinue
                If ($Service.Length -gt 0) {
                    Write-Host "Customize Service"
                    If ($WhatIf -eq '0') {
                        Stop-Service gupdate
                        Set-Service gupdate -StartupType Disabled
                        Stop-Service gupdatem
                        Set-Service gupdatem -StartupType Disabled
                    }
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                Write-Host "Customize Scheduled Task"
                If ($WhatIf -eq '0') {
                    Get-ScheduledTask -TaskName GoogleUpdateTaskMachine* -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
                    Get-ScheduledTask -TaskName GPUpdate* -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
                }
                Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Greenshot
    If ($Greenshot -eq 1) {
        $Product = "Greenshot"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $GreenshotD.Version
        }
        $GreenshotV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Greenshot*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$GreenshotV) {
            $GreenshotV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Greenshot*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $GreenshotV"
        If ($GreenshotV -lt $Version) {
            $Options = @(
                "/VERYSILENT"
                "/NORESTART"
                "/NORESTARTAPPLICATIONS"
                "/SUPPRESSMSGBOXES"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\Greenshot-INSTALLER-x86.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Greenshot") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Greenshot" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install ImageGlass
    If ($ImageGlass -eq 1) {
        $Product = "ImageGlass"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ImageGlassArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $ImageGlassD.Version
        }
        $ImageGlassV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "ImageGlass"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $ChromeLog = "$LogTemp\ImageGlass.log"
        If (!$ImageGlassV) {
            $ImageGlassV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "ImageGlass"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $ImageGlassInstaller = "ImageGlass_" + "$ImageGlassArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$ImageGlassInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ImageGlassArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $ImageGlassV"
        If ($ImageGlassV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/QUIET"
                "/L* $ImageGlassLog"
                "/NORESTART"
            )
            Try {
                Write-Host "Starting install of $Product $ImageGlassArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $ImageGlassLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $ImageGlassLog -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 10
                If ($WhatIf -eq '0') {
                    If (Test-Path -Path "$env:PUBLIC\Desktop\ImageGlass.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\ImageGlass.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\ImageGlass") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\ImageGlass" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install IrfanView
    If ($IrfanView -eq 1) {
        $Product = "IrfanView"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$IrfanViewArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $IrfanViewD.Version
        }
        $IrfanViewV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*IrfanView*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$IrfanViewV) {
            $IrfanViewV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*IrfanView*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $IrfanViewInstaller = "IrfanView" + "$IrfanViewArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $IrfanViewLanguageLongClear $IrfanViewArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $IrfanViewV"
        If ($IrfanViewV -lt $Version) {
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
                Write-Host "Starting install of $Product $IrfanViewArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$IrfanViewInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else {
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\IrfanView") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\IrfanView" -Recurse -Force}
                    }
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $IrfanViewArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            If ($IrfanViewLanguageLongClear -ne "English") {
                Write-Host "Copy $Product language pack $IrfanViewLanguageLongClear"
                If ($WhatIf -eq '0') {
                    Switch ($IrfanViewArchitectureClear) {
                        x64 { 
                            If (Test-Path -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\Languages\*.lng") {
                                Move-Item -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\Languages\*" -Destination "$Env:ProgramFiles\IrfanView\Languages" -ErrorAction SilentlyContinue
                            }
                            If (Test-Path -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\*.lng") {
                                Move-Item -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\*" -Destination "$Env:ProgramFiles\IrfanView\Languages" -ErrorAction SilentlyContinue
                            }
                        }
                        x86 {
                            If (Test-Path -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\Languages\*.lng") {
                                Move-Item -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\Languages\*" -Destination "${Env:ProgramFiles(x86)}\IrfanView\Languages" -ErrorAction SilentlyContinue
                            }
                            If (Test-Path -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\*.lng") {
                                Move-Item -Path "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear\*" -Destination "${Env:ProgramFiles(x86)}\IrfanView\Languages" -ErrorAction SilentlyContinue
                            }
                        }
                    }
                    Remove-Item "$PSScriptRoot\$Product\$IrfanViewLanguageLongClear" -Recurse
                }
                Write-Host -ForegroundColor Green "Copy $Product language pack $IrfanViewLanguageLongClear finished!"
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install KeePass
    If ($KeePass -eq 1) {
        $Product = "KeePass"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $KeePassD.Version
        }
        $KeePassV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*KeePass*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$KeePassV) {
            $KeePassV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*KeePass*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $KeePassLog = "$LogTemp\KeePass.log"
        $InstallMSI = "$PSScriptRoot\$Product\KeePass.msi"
        Write-Host -ForegroundColor Magenta "Install $Product $KeePassLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $KeePassV"
        If ($KeePassV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/quiet"
                "/L*V $KeePassLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $KeePassLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $KeePassLog -ErrorAction SilentlyContinue
                If ($WhatIf -eq '0') {
                    If (Test-Path -Path "$env:PUBLIC\Desktop\KeePass.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\KeePass.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\KeePass") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\KeePass" -Recurse -Force}
                    }
                }
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Write-Host "Copy $Product language pack $KeePassLanguageClear"
            If ($WhatIf -eq '0') {
                Move-Item -Path "$PSScriptRoot\$Product\$KeePassLanguageClear.lngx" -Destination "${Env:ProgramFiles(x86)}\KeePass2x\Languages" -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Copy $Product language pack $KeePassLanguageClear finished!"
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install LogMeIn GoToMeeting
    If ($LogMeInGoToMeeting -eq 1) {
        If ($LogMeInGoToMeetingInstallerClear -eq "Machine Based") {
            $Product = "LogMeIn GoToMeeting XenApp"
            # Check, if a new version is available
            $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $LogMeInGoToMeetingD.Version
            }
            $LogMeInGoToMeetingV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*GoToMeeting*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            If (!$LogMeInGoToMeetingV) {
                $LogMeInGoToMeetingV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*GoToMeeting*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            }
            $LogMeInGoToMeetingLog = "$LogTemp\LogMeInGoToMeeting.log"
            $InstallMSI = "$PSScriptRoot\$Product\GoToMeeting-Setup.msi"
            Write-Host -ForegroundColor Magenta "Install $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $LogMeInGoToMeetingV"
            If ($LogMeInGoToMeetingV -lt $Version) {
                DS_WriteLog "I" "Installing $Product" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $Arguments = @(
                    "/i"
                    "`"$InstallMSI`""
                    "/quiet"
                    "/L*V $LogMeInGoToMeetingLog"
                )
                Try {
                    Write-Host "Starting install of $Product $Version"
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                    }
                    Get-Content $LogMeInGoToMeetingLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $LogMeInGoToMeetingLog -ErrorAction SilentlyContinue
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
            If ($CleanUp -eq '1') {
                If ($WhatIf -eq '0') {
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                }
                Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
        }
        If ($LogMeInGoToMeetingInstallerClear -eq "User Based") {
            $Product = "LogMeIn GoToMeeting"
            # Check, if a new version is available
            $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $LogMeInGoToMeetingD.Version
            }
            $LogMeInGoToMeetingV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*GoToMeeting*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            If (!$LogMeInGoToMeetingV) {
                $LogMeInGoToMeetingV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*GoToMeeting*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            }
            $LogMeInGoToMeetingLog = "$LogTemp\LogMeInGoToMeeting.log"
            $InstallMSI = "$PSScriptRoot\$Product\GoToMeeting-Setup.msi"
            Write-Host -ForegroundColor Magenta "Install $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $LogMeInGoToMeetingV"
            If ($LogMeInGoToMeetingV -lt $Version) {
                DS_WriteLog "I" "Installing $Product" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $Arguments = @(
                    "/i"
                    "`"$InstallMSI`""
                    "/quiet"
                    "/L*V $LogMeInGoToMeetingLog"
                )
                Try {
                    Write-Host "Starting install of $Product $Version"
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                    }
                    Get-Content $LogMeInGoToMeetingLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $LogMeInGoToMeetingLog -ErrorAction SilentlyContinue
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
            If ($CleanUp -eq '1') {
                If ($WhatIf -eq '0') {
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                }
                Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
        }
    }

    #// Mark: Install Microsoft .Net Framework
    If ($MSDotNetFramework -eq 1) {
        $Product = "Microsoft Dot Net Framework"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSDotNetFrameworkArchitectureClear" + "_$MSDotNetFrameworkChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSDotNetFrameworkD.Version
        }
        $MSDotNetFrameworkV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Windows Desktop Runtime*" -and $_.URLInfoAbout -like "https://dot.net/core"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSDotNetFrameworkV) {
            $MSDotNetFrameworkV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Windows Desktop Runtime*" -and $_.URLInfoAbout -like "https://dot.net/core"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSDotNetFrameworkInstaller = "NetFramework-runtime_" + "$MSDotNetFrameworkArchitectureClear" + "_$MSDotNetFrameworkChannelClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $MSDotNetFrameworkArchitectureClear $MSDotNetFrameworkChannelClear Channel"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSDotNetFrameworkV"
        If ($MSDotNetFrameworkV -ne $Version) {
            $Options = @(
                "/install"
                "/quiet"
                "/norestart"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $MSDotNetFrameworkArchitectureClear $MSDotNetFrameworkChannelClear Channel $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$MSDotNetFrameworkInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSDotNetFrameworkArchitectureClear $MSDotNetFrameworkChannelClear Channel (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft 365 Apps
    If ($MS365Apps -eq 1) {
        $Product = "Microsoft 365 Apps"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MS365AppsD.Version
        }
        $MS365AppsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft 365*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MS365AppsV) {
            $MS365AppsV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft 365*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MS365AppsInstaller = "setup_" + "$MS365AppsChannelClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $MS365AppsChannelClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MS365AppsV"
        If ($MS365AppsV -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            # Download Apps 365 install files
            If (!(Test-Path -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Office\Data\$Version")) {
                Write-Host "Starting download of $Product install files"
                $DApps365 = @(
                    "/download install.xml"
                )
                If ($WhatIf -eq '0') {
                    set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                    Start-Process ".\$MS365AppsInstaller" -ArgumentList $DApps365 -wait -NoNewWindow
                    set-location $PSScriptRoot
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version install files finished!"
            }
            # MS365Apps Uninstallation
            $Options = @(
                "/configure remove.xml"
            )
            Write-Host "Uninstall Microsoft Office or Microsoft 365 Apps"
            DS_WriteLog "I" "Uninstall Microsoft Office or Microsoft 365 Apps" $LogFile
            Try {
                If ($WhatIf -eq '0') {
                    set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                    Start-Process -FilePath ".\$MS365AppsInstaller" -ArgumentList $Options -NoNewWindow -wait
                    set-location $PSScriptRoot
                }
                Write-Host -ForegroundColor Green "Uninstall Microsoft Office or Microsoft 365 Apps finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error uninstalling Microsoft Office or Microsoft 365 Apps (Error: $($Error[0]))"
                DS_WriteLog "E" "Error uninstalling Microsoft Office or Microsoft 365 Apps (Error: $($Error[0]))" $LogFile
            }
            # MS365Apps Installation
            $Options = @(
                "/configure install.xml"
            )
            Try {
                DS_WriteLog "I" "Install $Product" $LogFile
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                    Start-Process -FilePath ".\$MS365AppsInstaller" -ArgumentList $Options -NoNewWindow -wait
                    set-location $PSScriptRoot
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools" -Recurse -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Access.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Access.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Excel.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Excel.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\OneNote.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\OneNote.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Outlook.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Publisher.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Publisher.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Project.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Project.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visio.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visio.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Word.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Word.lnk" -Force}
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Outlook.lnk") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft AVD Remote Desktop
    If ($MSAVDRemoteDesktop -eq 1) {
        $Product = "Microsoft AVD Remote Desktop"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSAVDRemoteDesktopChannelClear" + "$MSAVDRemoteDesktopArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSAVDRemoteDesktopD.Version
        }
        $MSAVDRemoteDesktopV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Remotedesktop"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $MSAVDRemoteDesktopLog = "$LogTemp\MSAVDRemoteDesktop.log"
        If (!$MSAVDRemoteDesktopV) {
            $MSAVDRemoteDesktopV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Remotedesktop"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSAVDRemoteDesktopInstaller = "RemoteDesktop_" + "$MSAVDRemoteDesktopArchitectureClear" + "_$MSAVDRemoteDesktopChannelClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$MSAVDRemoteDesktopInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $MSAVDRemoteDesktopChannelClear $MSAVDRemoteDesktopArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSAVDRemoteDesktopV"
        If ($MSAVDRemoteDesktopV -ne $Version) {
            DS_WriteLog "I" "Install $Product $MSAVDRemoteDesktopChannelClear $MSAVDRemoteDesktopArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $MSAVDRemoteDesktopLog"
            )
            try {
                Write-Host "Starting install of $Product $MSAVDRemoteDesktopChannelClear $MSAVDRemoteDesktopArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Remote Desktop.lnk") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Remote Desktop.lnk" -Force}
                    }
                }
                Get-Content $MSAVDRemoteDesktopLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $MSAVDRemoteDesktopLog -ErrorAction SilentlyContinue
            } catch {
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Azure CLI
    If ($MSAzureCLI -eq 1) {
        $Product = "Microsoft Azure CLI"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSAzureCLID.Version
        }
        $MSAzureCLIV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Azure CLI"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $MSAzureCLILog = "$LogTemp\MSAzureCLI.log"
        If (!$MSAzureCLIV) {
            $MSAzureCLIV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Azure CLI"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSAzureCLIInstaller = "AzureCLI.msi"
        $InstallMSI = "$PSScriptRoot\$Product\$MSAzureCLIInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSAzureCLIV"
        If ($MSAzureCLIV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $MSAzureCLILog"
            )
            try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $MSAzureCLILog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $MSAzureCLILog -ErrorAction SilentlyContinue
            } catch {
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Azure Data Studio
    If ($MSAzureDataStudio -eq 1) {
        $Product = "Microsoft Azure Data Studio"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSAzureDataStudioChannelClear" + "-$MSAzureDataStudioPlatformClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSAzureDataStudioD.Version
        }
        $MSAzureDataStudioV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Azure Data Studio*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSAzureDataStudioV) {
            $MSAzureDataStudioV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Azure Data Studio*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If (!$MSAzureDataStudioV) {
            $MSAzureDataStudioV = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Azure Data Studio*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSAzureDataStudioInstaller = "AzureDataStudio-Setup-" + "$MSAzureDataStudioChannelClear" + "-$MSAzureDataStudioPlatformClear" + "." + "exe"
        $MSAzureDataStudioProcess = "AzureDataStudio-Setup-" + "$MSAzureDataStudioChannelClear" + "-$MSAzureDataStudioPlatformClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MSAzureDataStudioChannelClear $ArchitectureClear $MSAzureDataStudioModeClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSAzureDataStudioV"
        If ($MSAzureDataStudioV -ne $Version) {
            Write-Host -ForegroundColor Green "Update available"
            DS_WriteLog "I" "Install $Product $Product $MSAzureDataStudioChannelClear $ArchitectureClear $MSAzureDataStudioModeClear" $LogFile
            $Options = @(
                "/VERYSILENT"
                "/NORESTART"
                "/MERGETASKS=!runcode"
            )
            Try {
                Write-Host "Starting install of $Product $MSAzureDataStudioChannelClear $ArchitectureClear $MSAzureDataStudioModeClear $Version"
                If ($WhatIf -eq '0') {
                    $null = Start-Process "$PSScriptRoot\$Product\$MSAzureDataStudioInstaller" -ArgumentList $Options -NoNewWindow -PassThru
                    while (Get-Process -Name $MSAzureDataStudioProcess -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Azure Data Studio") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Azure Data Studio" -Recurse -Force}
                    }
                }
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSAzureDataStudioChannelClear $ArchitectureClear $MSAzureDataStudioModeClear (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Edge
    If ($MSEdge -eq 1) {
        $Product = "Microsoft Edge"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSEdgeArchitectureClear" + "_$MSEdgeChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $EdgeD.Version
        }
        $Edge = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $EdgeLog = "$LogTemp\MSEdge.log"
        If (!$Edge) {
            $Edge = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $EdgeInstaller = "MicrosoftEdgeEnterprise_" + "$MSEdgeArchitectureClear" + "_$MSEdgeChannelClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$EdgeInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $MSEdgeChannelClear $MSEdgeArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $Edge"
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
                Write-Host "Starting install of $Product $MSEdgeChannelClear $MSEdgeArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk" -Force}
                    }
                }
                Get-Content $EdgeLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $EdgeLog -ErrorAction SilentlyContinue
            } catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Write-Host "Customize $Product"
            Try {
                # Disable Microsoft Edge auto update
                Write-Host "Customize $Product registry"
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate)) {
                        New-Item -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -ErrorAction SilentlyContinue | Out-Null
                        New-ItemProperty -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -Name UpdateDefault -Value 0 -PropertyType DWORD -ErrorAction SilentlyContinue | Out-Null
                    }
                    Else {
                        $EdgeUpdateState = Get-ItemProperty -path "HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate" | Select-Object -Expandproperty "UpdateDefault"
                        If ($EdgeUpdateState -ne "0") {Set-ItemProperty -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -Name UpdateDefault -Value 0 | Out-Null}
                    }
                    # Disable Citrix API Hooks (MS Edge) on Citrix VDA
                    $(
                        $RegPath = "HKLM:SYSTEM\CurrentControlSet\services\CtxUvi"
                        If (Test-Path $RegPath) {
                            $RegName = "UviProcessExcludes"
                            $EdgeRegvalue = "msedge.exe"
                            # Get current values in UviProcessExcludes
                            $CurrentValues = Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $RegName
                            # Add the msedge.exe value to existing values in UviProcessExcludes
                            Set-ItemProperty -Path $RegPath -Name $RegName -Value "$CurrentValues$EdgeRegvalue;" -ErrorAction SilentlyContinue
                        }
                    ) | Out-Null
                }
                Write-Host -ForegroundColor Green "Customize $Product registry finished!"
                $Service = Get-Service -Name edgeupdate -ErrorAction SilentlyContinue
                If ($Service.Length -gt 0) {
                    Write-Host "Customize Service"
                    If ($WhatIf -eq '0') {
                        Stop-Service edgeupdate
                        Set-Service -Name edgeupdate -StartupType Manual
                        Stop-Service edgeupdatem
                        Set-Service -Name edgeupdatem -StartupType Manual
                    }
                    Write-Host -ForegroundColor Green "Stop and Disable Service $Product finished!"
                }
                Write-Host "Customize Scheduled Task"
                If ($WhatIf -eq '0') {
                    Start-Sleep -s 5
                    Get-ScheduledTask -TaskName MicrosoftEdgeUpdate* -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Edge WebView2
    If ($MSEdgeWebView2 -eq 1) {
        $Product = "Microsoft Edge WebView2"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSEdgeWebView2ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $EdgeWebView2D.Version
        }
        $EdgeWebView2 = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Microsoft Edge WebView2*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$EdgeWebView2) {
            $EdgeWebView2 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Microsoft Edge WebView2*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $EdgeWebView2Installer = "MicrosoftEdgeWebView2_" + "$MSEdgeWebView2ArchitectureClear" + ".exe"
        $EdgeWebView2Process = "MicrosoftEdgeWebView2_" + "$MSEdgeWebView2ArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MSEdgeWebView2ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $EdgeWebView2"
        If ($EdgeWebView2 -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/SILENT"
                "/INSTALL"
            )
            Try {
                Write-Host "Starting install of $Product $MSEdgeWebView2ArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $null = Start-Process "$PSScriptRoot\$Product\$EdgeWebView2Installer" -ArgumentList $Options -NoNewWindow -PassThru
                    while (Get-Process -Name $EdgeWebView2Process -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                }
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } catch {
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft FSLogix
    If ($MSFSLogix -eq 1) {
        $Product = "Microsoft FSLogix"
        $OS = (Get-WmiObject Win32_OperatingSystem).Caption
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\$MSFSLogixChannelClear\Version_" + "$MSFSLogixChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSFSLogixD.Version
        }
        $MSFSLogixV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSFSLogixV) {
            $MSFSLogixV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}) {
            $UninstallFSL = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps"}).UninstallString.replace("/uninstall","")
        }
        If (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps RuleEditor"}) {
            $UninstallFSLRE = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft FSLogix Apps RuleEditor"}).UninstallString.replace("/uninstall","")
        }
        Write-Host -ForegroundColor Magenta "Install $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSFSLogixV"
        If ($MSFSLogixV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            # FSLogix Uninstall
            If ($MSFSLogixV) {
                Write-Host "Uninstall $Product"
                DS_WriteLog "I" "Uninstall $Product" $LogFile
                Try {
                    If ($WhatIf -eq '0') {
                        Start-process $UninstallFSL -ArgumentList '/uninstall /quiet /norestart' -NoNewWindow -Wait
                        Start-process $UninstallFSLRE -ArgumentList '/uninstall /quiet /norestart' -NoNewWindow -Wait
                    }
                    Write-Host -ForegroundColor Green "Uninstall $Product finished!"
                } Catch {
                    Write-Host -ForegroundColor Red "Error uninstalling $Product (Error: $($Error[0]))"
                    DS_WriteLog "E" "Error uninstalling $Product (Error: $($Error[0]))" $LogFile
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
                Write-Host -ForegroundColor Red "Server needs to reboot, start script again after reboot"
                If ($WhatIf -eq '0') {
                    Write-Output ""
                    Write-Host -ForegroundColor Red "Hit any key to reboot server"
                    Read-Host
                    Restart-Computer
                }
            }
            # FSLogix Install
            Try {
                Write-Host "Starting install of $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$MSFSLogixChannelClear\FSLogixAppsSetup.exe" -ArgumentList '/install /norestart /quiet' -NoNewWindow -Wait
                }
                Write-Host -ForegroundColor Green "Install $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear finished!"
                Write-Host "Starting install of $Product Rule Editor $MSFSLogixChannelClear $MSFSLogixArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$MSFSLogixChannelClear\FSLogixAppsRuleEditorSetup.exe" -ArgumentList '/install /norestart /quiet' -NoNewWindow -Wait
                }
                Write-Host -ForegroundColor Green "Install $Product Rule Editor $MSFSLogixChannelClear $MSFSLogixArchitectureClear finished!"
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Try {
                Start-Sleep -s 20
                # Application post deployment tasks (Thx to Kasper https://github.com/kaspersmjohansen)
                Write-Host "Applying $Product post setup customizations"
                Write-Host "Post setup customizations for $OS"
                If ($OS -Like "*Windows Server 2019*" -or $OS -eq "Microsoft Windows 10 Enterprise for Virtual Desktops" -or $OS -Like "*Windows Server 2021*") {
                    If ((Test-RegistryValue2 -Path "HKLM:SOFTWARE\FSLogix\Apps" -Value "RoamSearch") -ne $true) {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        If ($WhatIf -eq '0') {
                            New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "0" -Type DWORD -ErrorAction SilentlyContinue | Out-Null
                        }
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                    If ((Get-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" | Select-Object -ExpandProperty "RoamSearch" -ErrorAction SilentlyContinue) -ne "0") {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        If ($WhatIf -eq '0') {
                            Set-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "0" -Type DWORD -ErrorAction SilentlyContinue
                        }
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                }
                If ($OS -Like "*Windows 10*" -and $OS -ne "Microsoft Windows 10 Enterprise for Virtual Desktops") {
                    If ((Test-RegistryValue2 -Path "HKLM:SOFTWARE\FSLogix\Apps" -Value "RoamSearch") -ne $true) {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        If ($WhatIf -eq '0') {
                            New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "1" -Type DWORD -ErrorAction SilentlyContinue | Out-Null
                        }
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                    If ((Get-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" | Select-Object -ExpandProperty "RoamSearch" -ErrorAction SilentlyContinue) -ne "1") {
                        Write-Host "Deactivate FSLogix RoamSearch"
                        If ($WhatIf -eq '0') {
                            Set-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "1" -Type DWORD -ErrorAction SilentlyContinue
                        }
                        Write-Host -ForegroundColor Green "Deactivate FSLogix RoamSearch finished!"
                    }
                }
                Write-Host -ForegroundColor Green "Post setup customizations for $OS finished!"
                # Implement user based group policy processing fix
                If ($WhatIf -eq '0') {
                    If (!(Test-Path -Path HKLM:SOFTWARE\FSLogix\Profiles)) {
                        New-Item -Path "HKLM:SOFTWARE\FSLogix" -Name Profiles -ErrorAction SilentlyContinue | Out-Null
                    }
                }
                If ((Test-RegistryValue -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Value "GroupPolicyState") -ne $true) {
                    Write-Host "Deactivate FSLogix GroupPolicy"
                    If ($WhatIf -eq '0') {
                        New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Name "GroupPolicyState" -Value "0" -Type DWORD -ErrorAction SilentlyContinue | Out-Null
                    }
                    Write-Host -ForegroundColor Green "Deactivate FSLogix GroupPolicy finished!"
                }
                If ((Get-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" | Select-Object -ExpandProperty "GroupPolicyState" -ErrorAction SilentlyContinue) -ne "0") {
                    Write-Host "Deactivate FSLogix GroupPolicy"
                    If ($WhatIf -eq '0') {
                        Set-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Name "GroupPolicyState" -Value "0" -Type DWORD -ErrorAction SilentlyContinue
                    }
                    Write-Host -ForegroundColor Green "Deactivate FSLogix GroupPolicy finished!"
                }
                If (!(Get-ScheduledTask -TaskName "Restart Windows Search Service on Event ID 2" -ErrorAction SilentlyContinue)) {
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
                    If ($WhatIf -eq '0') {
                        Register-ScheduledTask @RegSchTaskParameters -ErrorAction SilentlyContinue
                    }
                    Write-Host -ForegroundColor Green "Implement scheduled task to restart Windows Search service on Event ID 2 finished!"
                }
                Write-Host -ForegroundColor Green "Applying $Product post setup customizations finished!"
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSFSLogixChannelClear $MSFSLogixArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Office
    If ($MSOffice -eq 1) {
        $Product = "Microsoft Office " + $MSOfficeChannelClear
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSOfficeD.Version
        }
        $MSOfficeV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSOfficeV) {
            $MSOfficeV = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSOfficeV"
        If ($MSOfficeV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            # Download MS Office install files
            If (!(Test-Path -Path "$PSScriptRoot\$Product\Office\Data\$Version")) {
                Write-Host "Starting download of $Product install files"
                $DOffice = @(
                    "/download install.xml"
                )
                If ($WhatIf -eq '0') {
                    set-location $PSScriptRoot\$Product
                    Start-Process ".\setup.exe" -ArgumentList $DOffice -wait -NoNewWindow
                    set-location $PSScriptRoot
                }
                Write-Host -ForegroundColor Green "Download of the new version $Version install files finished!"
            }
            # MS Office Uninstallation
            $Options = @(
                "/configure remove.xml"
            )
            Write-Host "Uninstall Microsoft Office or Microsoft 365 Apps"
            DS_WriteLog "I" "Uninstall Microsoft Office or Microsoft 365 Apps" $LogFile
            Try {
                If ($WhatIf -eq '0') {
                    set-location $PSScriptRoot\$Product
                    Start-Process -FilePath ".\setup.exe" -ArgumentList $Options -NoNewWindow -wait
                    set-location $PSScriptRoot
                }
                Write-Host -ForegroundColor Green "Uninstall Microsoft Office or Microsoft 365 Apps finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error uninstalling Microsoft Office or Microsoft 365 Apps (Error: $($Error[0]))"
                DS_WriteLog "E" "Error uninstalling Microsoft Office or Microsoft 365 Apps (Error: $($Error[0]))" $LogFile
            }
            # MS Office Installation
            $Options = @(
                "/configure install.xml"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host "Starting install of $Product $Version"
            Try {
                If ($WhatIf -eq '0') {
                    set-location $PSScriptRoot\$Product
                    Start-Process -FilePath ".\setup.exe" -ArgumentList $Options -NoNewWindow -wait
                    set-location $PSScriptRoot
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft OneDrive
    If ($MSOneDrive -eq 1) {
        $Product = "Microsoft OneDrive"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSOneDriveD.Version
        }
        $MSOneDriveV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSOneDriveV) {
            $MSOneDriveV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If (!$MSOneDriveV) {
            $MSOneDriveV = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $OneDriveInstaller = "OneDriveSetup-" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear" + ".exe"
        $OneDriveProcess = "OneDriveSetup-" + "$MSOneDriveRingClear" + "_$MSOneDriveArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSOneDriveV"
        If ($MSOneDriveV -ne $Version) {
            Write-Host -ForegroundColor Green "Update available"
            DS_WriteLog "I" "Install $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear" $LogFile
            If ($MSOneDriveInstallerClear -eq 'Machine Based') {
                $Options = @(
                    "/allusers"
                    "/SILENT"
                )
            }
            If ($MSOneDriveInstallerClear -eq 'User Based') {
                $Options = @(
                    "/SILENT"
                )
            }
            Try {
                Write-Host "Starting install of $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $null = Start-Process "$PSScriptRoot\$Product\$OneDriveInstaller" -ArgumentList $Options -NoNewWindow -PassThru
                    while (Get-Process -Name $OneDriveProcess -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk" -Force}
                    }
                }
                # OneDrive starts automatically after setup. kill!
                #Stop-Process -Name "OneDrive" -Force
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSOneDriveRingClear Ring $MSOneDriveArchitectureClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            Write-Host "Customize $Product"
            Try {
                # Disable Microsoft OneDrive auto update
                Write-Host "Customize Scheduled Task"
                If ($WhatIf -eq '0') {
                    Start-Sleep -s 5
                    Get-ScheduledTask -TaskName OneDrive* -ErrorAction SilentlyContinue | Disable-ScheduledTask -ErrorAction SilentlyContinue | Out-Null
                }
                Write-Host -ForegroundColor Green "Disable Scheduled Task $Product finished!"
                Write-Host -ForegroundColor Green "Customize $Product finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error customizing (Error: $($Error[0]))"
                DS_WriteLog "E" "Error customizing (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            Write-Host "Starting copy of $Product $MSOneDriveRingClear ADMX files $Version"
            If ($WhatIf -eq '0') {
                $OneDriveUninstall = (Get-ItemProperty -Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object {$_.DisplayIcon -like "*OneDriveSetup.exe*"})
                If (!$OneDriveUninstall) {
                    $OneDriveUninstall = (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object {$_.DisplayIcon -like "*OneDriveSetup.exe*"})
                }
                If (!$OneDriveUninstall) {
                    $OneDriveUninstall = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object {$_.DisplayIcon -like "*OneDriveSetup.exe*"})
                }
                $OneDriveInstallFolder = $OneDriveUninstall.DisplayIcon.Substring(0, $OneDriveUninstall.DisplayIcon.IndexOf("\OneDriveSetup.exe"))
                $sourceadmx = "$($OneDriveInstallFolder)\adm"
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\OneDrive.admx" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse -ErrorAction SilentlyContinue
                }
                If (Test-Path -Path "$PSScriptRoot\_ADMX\$Product") {Remove-Item -Path "$PSScriptRoot\_ADMX\$Product" -Force -Recurse}
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product" -ItemType Directory | Out-Null }
                Move-Item -Path "$sourceadmx\OneDrive.admx" -Destination "$PSScriptRoot\_ADMX\$Product" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\en-US")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\en-US\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\en-US\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\en-US" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\de-DE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\de-DE\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\de-DE\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\de\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\de-DE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\es-ES")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\es-ES\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\es-ES\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\es\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\es-ES" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\fr-FR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\fr-FR\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\fr-FR\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\fr\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\fr-FR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\it-IT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\it-IT\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\it-IT\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\it\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\it-IT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ja-JP")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ja-JP\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ja-JP\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\ja\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ja-JP" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ko-KR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ko-KR\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ko-KR\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\ko\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ko-KR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\nl-NL")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\nl-NL\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\nl-NL\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\nl\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\nl-NL" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pl-PL")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pl-PL" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pl-PL\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pl-PL\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\pl\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pl-PL" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pt-BR")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pt-BR\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-BR\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\pt-BR\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-BR" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\pt-PT")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-PT" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\pt-PT\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\pt-PT\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\pt-PT\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\pt-PT" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\ru-RU")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\ru-RU\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\ru-RU\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\ru\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\ru-RU" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\sv-SE")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\sv-SE" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\sv-SE\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\sv-SE\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\sv\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\sv-SE" -ErrorAction SilentlyContinue
                If (!(Test-Path -Path "$PSScriptRoot\_ADMX\$Product\zh-CN")) { New-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN" -ItemType Directory | Out-Null }
                If ((Test-Path "$PSScriptRoot\_ADMX\$Product\zh-CN\OneDrive.adml" -PathType leaf)) {
                    Remove-Item -Path "$PSScriptRoot\_ADMX\$Product\zh-CN\OneDrive.adml" -ErrorAction SilentlyContinue
                }
                Move-Item -Path "$sourceadmx\zh-CN\OneDrive.adml" -Destination "$PSScriptRoot\_ADMX\$Product\zh-CN" -ErrorAction SilentlyContinue
            }
            Write-Host -ForegroundColor Green "Copy of the new ADMX files version $Version finished!"
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Power BI Desktop
    If ($MSPowerBIDesktop -eq 1) {
        $Product = "Microsoft Power BI Desktop"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSPowerBIDesktopArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSPowerBIDesktopD.Version
        }
        $MSPowerBIDesktopV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Microsoft PowerBI Desktop*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSPowerBIDesktopV) {
            $MSPowerBIDesktopV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Microsoft PowerBI Desktop*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSPowerBIDesktopInstaller = "PBIDesktopSetup_" + "$MSPowerBIDesktopArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MSPowerBIDesktopArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSPowerBIDesktopV"
        If ($MSPowerBIDesktopV -ne $Version) {
            DS_WriteLog "I" "Install $Product $MSPowerBIDesktopArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "-quiet"
                "-norestart"
                "ACCEPT_EULA=1 INSTALLDESKTOPSHORTCUT=0 ENABLECXP=0"
            )
            Try {
                Write-Host "Starting install of $Product $MSPowerBIDesktopArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$MSPowerBIDesktopInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Power BI Desktop") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Power BI Desktop" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Power BI Report Builder
    If ($MSPowerBIReportBuilder -eq 1) {
        $Product = "Microsoft Power BI Report Builder"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSPowerBIReportBuilderD.Version
        }
        If ($Version) {
            $VersionSplit = $Version.split("0")
            $Version = $VersionSplit[0] + $VersionSplit[1] + $VersionSplit[2] + $VersionSplit[3] + $VersionSplit[4] + $VersionSplit[5] + $VersionSplit[6]
        }
        $MSPowerBIReportBuilderV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Power BI Report Builder*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSPowerBIReportBuilderV) {
            $MSPowerBIReportBuilderV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Power BI Report Builder*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSPowerBIReportBuilderLog = "$LogTemp\MSPowerBIReportBuilder.log"
        $MSPowerBIReportBuilderInstaller = "PBIReportBuilderSetup.msi"
        $InstallMSI = "$PSScriptRoot\$Product\$MSPowerBIReportBuilderInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSPowerBIReportBuilderV"
        If ($MSPowerBIReportBuilderV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/norestart"
                "/L*V $MSPowerBIReportBuilderLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Power BI Report Builder") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Power BI Report Builder" -Recurse -Force}
                    }
                }
                Get-Content $MSPowerBIReportBuilderLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $MSPowerBIReportBuilderLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft PowerShell
    If ($MSPowerShell -eq 1) {
        $Product = "Microsoft PowerShell"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSPowerShellArchitectureClear" + "_$MSPowerShellReleaseClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSPowershellD.Version
        }
        $MSPowerShellV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PowerShell*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSPowerShellV) {
            $MSPowerShellV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PowerShell*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($MSPowerShellV) {$MSPowerShellV = $MSPowerShellV -replace ".{2}$"}
        $MSPowerShellLog = "$LogTemp\MSPowerShell.log"
        $MSPowerShellInstaller = "PowerShell" + "$MSPowerShellArchitectureClear" + "_$MSPowerShellReleaseClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$MSPowerShellInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $MSPowerShellArchitectureClear $MSPowerShellReleaseClear Release"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSPowerShellV"
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
                Write-Host "Starting install of $Product $MSPowerShellArchitectureClear $MSPowerShellReleaseClear Release $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PowerShell") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PowerShell" -Recurse -Force}
                    }
                }
                Get-Content $MSPowerShellLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $MSPowerShellLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft PowerToys
    If ($MSPowerToys -eq 1) {
        $Product = "Microsoft PowerToys"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSPowerToysD.Version
        }
        $MSPowerToysV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PowerToys*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSPowerToysV) {
            $MSPowerToysV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PowerToys*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSPowerToysV"
        If ($MSPowerToysV -lt $Version) {
            $Options = @(
                "--silent"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\PowerToysSetup-x64.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PowerToys*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PowerToys*" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft SQL Server Management Studio
    If ($MSSQLServerManagementStudio -eq 1) {
        $Product = "Microsoft SQL Server Management Studio"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSSQLServerManagementStudioLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSSQLServerManagementStudioD.Version
        }
        If ($Version) {
            $VersionSplit = $Version.split(".")
            $VersionSplit2 = $VersionSplit[2].Substring(0,3)
            $Version = $VersionSplit[0] + "." + $VersionSplit[1] + "." + $VersionSplit2
        }
        $MSSQLServerManagementStudioV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*SQL Server Management Studio*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSSQLServerManagementStudioV) {
            $MSSQLServerManagementStudioV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*SQL Server Management Studio*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($MSSQLServerManagementStudioV) {
            $MSSQLServerManagementStudioVSplit = $MSSQLServerManagementStudioV.split(".")
            $MSSQLServerManagementStudioVSplit2 = $MSSQLServerManagementStudioVSplit[2].Substring(0,3)
            $MSSQLServerManagementStudioV = $MSSQLServerManagementStudioVSplit[0] + "." + $MSSQLServerManagementStudioVSplit[1] + "." + $MSSQLServerManagementStudioVSplit2
        }
        $MSSQLServerManagementStudioInstaller = "SSMS-Setup_" + "$MSSQLServerManagementStudioLanguageClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $MSSQLServerManagementStudioLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSSQLServerManagementStudioV"
        If ($MSSQLServerManagementStudioV -ne $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/install"
                "/quiet"
                "/norestart"
            )
            Try {
                Write-Host "Starting install of $Product $MSSQLServerManagementStudioLanguageClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$MSSQLServerManagementStudioInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft SQL Server Tools*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft SQL Server Tools*" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Teams
    If ($MSTeams -eq 1) {
        If ($MSTeamsInstallerClear -eq 'Machine Based') {
            $Product = "Microsoft Teams Machine Based"
            # Check, if a new version is available
            $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
            $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $TeamsD.Version
            }
            If (Test-Path -Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\") {
                $Teams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            }
            If (!$Teams) {
                If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\") {
                    $Teams = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
                }
            }
            $TeamsInstaller = "Teams_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear" + ".msi"
            $TeamsLog = "$LogTemp\MSTeams.log"
            $InstallMSI = "$PSScriptRoot\$Product\$TeamsInstaller"
            If ($Teams) {$Teams = $Teams.Insert(5,'0')}
            Write-Host -ForegroundColor Magenta "Install $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $Teams"
            If ($Teams -ne $Version) {
                DS_WriteLog "I" "Install $Product" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                #Uninstalling MS Teams
                If ($Teams) {
                    Write-Host "Uninstall $Product"
                    DS_WriteLog "I" "Uninstall $Product" $LogFile
                    Try {
                        If (Test-Path -Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\") {
                            $UninstallTeams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).UninstallString
                        }
                        If (!$UninstallTeams) {
                            If (Test-Path -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\") {
                                $UninstallTeams = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).UninstallString
                            }
                        }
                        $UninstallTeams = $UninstallTeams -Replace("MsiExec.exe /I","")
                        If ($WhatIf -eq '0') {
                            Start-Process -FilePath msiexec.exe -ArgumentList "/X $UninstallTeams /qn /L*V $TeamsLog"
                            Start-Sleep 20
                        }
                        Get-Content $TeamsLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                        Remove-Item $TeamsLog -ErrorAction SilentlyContinue
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
                    Write-Host "Customize System for $Product Machine Based Install"
                    If ($WhatIf -eq '0') {
                        If (!(Test-Path 'HKLM:\Software\Citrix\')) {New-Item -Path "HKLM:Software\Citrix" | Out-Null}
                        New-Item -Path "HKLM:Software\Citrix\PortICA" | Out-Null
                    }
                    Write-Host -ForegroundColor Green "Customize System for $Product Machine Based Install finished!"
                }
                Try {
                    Write-Host "Starting install of $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring $Version"
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                        Start-Sleep 5
                    }
                    Get-Content $TeamsLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $TeamsLog -ErrorAction SilentlyContinue
                    If ($WhatIf -eq '0') {
                        If (Test-Path "$env:PUBLIC\Desktop\Microsoft Teams.lnk") { Remove-Item -Path "$env:PUBLIC\Desktop\Microsoft Teams.lnk" -Force }
                        If ($CleanUpStartMenu) {
                            If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk" -Force}
                        }
                    }
                } Catch {
                    DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
                }
                Try {
                    Write-Host "Customize $Product"
                    #reg add "HKLM\SOFTWARE\Citrix\CtxHook\AppInit_Dlls\SfrHook" /v Teams.exe /t REG_DWORD /d 204 /f | Out-Null
                    If ($MSTeamsNoAutoStart -eq 1) {
                        #Prevents MS Teams from starting at logon, better do this with WEM or similar
                        Write-Host "Customize $Product Autorun"
                        If ($WhatIf -eq '0') {
                            If (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run") {
                                If (Test-RegistryValue2 -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Value "Teams") {
                                    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -Force
                                }
                            }
                        }
                        Write-Host -ForegroundColor Green "Customize $Product Autorun finished!"
                    }
                    Write-Host "Register $Product Add-In for Outlook"
                    # Register Teams add-in for Outlook - https://microsoftteams.uservoice.com/forums/555103-public/suggestions/38846044-fix-the-teams-meeting-addin-for-outlook
                    If ($WhatIf -eq '0') {
                        $appDLLs = (Get-ChildItem -Path "${Env:ProgramFiles(x86)}\Microsoft\TeamsMeetingAddin" -Include "Microsoft.Teams.AddinLoader.dll" -Recurse).FullName
                        $appX64DLL = $appDLLs[0]
                        $appX86DLL = $appDLLs[1]
                        Start-Process -FilePath "$env:WinDir\SysWOW64\regsvr32.exe" -ArgumentList "/s /n /i:user `"$appX64DLL`"" -ErrorAction SilentlyContinue
                        Start-Process -FilePath "$env:WinDir\SysWOW64\regsvr32.exe" -ArgumentList "/s /n /i:user `"$appX86DLL`"" -ErrorAction SilentlyContinue
                    }
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
            If ($CleanUp -eq '1') {
                If ($WhatIf -eq '0') {
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                }
                Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
        }
        If ($MSTeamsInstallerClear -eq 'User Based') {
            $Product = "Microsoft Teams User Based"
            # Check, if a new version is available
            $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
            $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $TeamsD.Version
            }
            If (Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\") {
                $Teams = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Teams*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            }
            If (!$Teams) {
                If (Test-Path -Path "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\") {
                    $Teams = (Get-ItemProperty HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Teams*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
                }
            }
            $TeamsInstaller = "Teams_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear" + ".exe"
            $TeamsProcess = "Teams_" + "$MSTeamsArchitectureClear" + "_$MSTeamsRingClear"
            Write-Host -ForegroundColor Magenta "Install $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $Teams"
            If ($Teams -ne $Version) {
                DS_WriteLog "I" "Install $Product" $LogFile
                Write-Host -ForegroundColor Green "Update available"
                $Options = @(
                    "/s"
                )
                Try {
                    Write-Host "Starting install of $Product $MSTeamsArchitectureClear $MSTeamsRingClear Ring $Version"
                    If ($WhatIf -eq '0') {
                        $null = Start-Process -FilePath "$PSScriptRoot\$Product\$TeamsInstaller" -ArgumentList $Options -PassThru -NoNewWindow
                        while (Get-Process -Name $TeamsProcess -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                    }
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                    If ($WhatIf -eq '0') {
                        If (Test-Path "$env:USERPROFILE\Desktop\Microsoft Teams.lnk") {Remove-Item -Path "$env:USERPROFILE\Desktop\Microsoft Teams.lnk" -Force}
                        If ($CleanUpStartMenu) {
                            If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk" -Force}
                        }
                    }
                } Catch {
                    DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
                }
                Try {
                    Write-Host "Customize $Product"
                    If ($MSTeamsNoAutoStart -eq 1) {
                        #Prevents MS Teams from starting at logon, better do this with WEM or similar
                        Write-Host "Customize $Product Autorun"
                        If ($WhatIf -eq '0') {
                            If (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run") {
                                If (Test-RegistryValue2 -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Value "Teams") {
                                    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -Force
                                }
                            }
                        }
                        Write-Host -ForegroundColor Green "Customize $Product Autorun finished!"
                    }
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
            If ($CleanUp -eq '1') {
                If ($WhatIf -eq '0') {
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                }
                Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
            }
        }
    }

    #// Mark: Install Microsoft Visual Studio 2019
    If ($MSVisualStudio -eq 1) {
        $Product = "Microsoft Visual Studio 2019"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSVisualStudioD.Version
        }
        $MSVisualStudioV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visual Studio $MSVisualStudioEditionClear*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSVisualStudioV) {
            $MSVisualStudioV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visual Studio $MSVisualStudioEditionClear*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product $MSVisualStudioEditionClear Edition"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSVisualStudioV"
        If ($MSVisualStudioV -ne $Version) {
            $MSVisualStudioEditionInstall = "Microsoft.VisualStudio.Product." + "$MSVisualStudioEditionClear"
            If ($MSVisualStudioV) {
                $Options = @(
                    "update"
                    "--quiet"
                    "--norestart"
                    "--productid $MSVisualStudioEditionInstall"
                    "--channelid VisualStudio.16.Release"
                )
            }
            Else {
                $Options = @(
                    "--quiet"
                    "--norestart"
                    "--productid $MSVisualStudioEditionInstall"
                    "--channelid VisualStudio.16.Release"
                )
            }
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $MSVisualStudioEditionClear Edition $Version"
                If ($WhatIf -eq '0') {
                    $null = Start-Process -FilePath "$PSScriptRoot\$Product\VS-Setup.exe" -ArgumentList $Options -PassThru -NoNewWindow
                    while (Get-Process -Name setup -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio 2019") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio 2019" -Recurse -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio 2019.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio 2019.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio Installer.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio Installer.lnk" -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Blend for Visual Studio 2019.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Blend for Visual Studio 2019.lnk" -Force}
                    }
                }
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSVisualStudioEditionClear Edition (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product $MSVisualStudioEditionClear Edition (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product $MSVisualStudioEditionClear Edition"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Start-Sleep -Seconds 20
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft Visual Studio Code
    If ($MSVisualStudioCode -eq 1) {
        $Product = "Microsoft Visual Studio Code"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSVisualStudioCodeChannelClear" + "-$MSVisualStudioCodePlatformClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MSVisualStudioCodeD.Version
        }
        $MSVisualStudioCodeV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visual Studio Code*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MSVisualStudioCodeV) {
            $MSVisualStudioCodeV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visual Studio Code*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If (!$MSVisualStudioCodeV) {
            $MSVisualStudioCodeV = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visual Studio Code*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MSVisualStudioCodeInstall = "VSCode-Setup-" + "$MSVisualStudioCodeChannelClear" + "-$MSVisualStudioCodePlatformClear" + "." + "exe"
        $MSVisualStudioCodeProcess = "VSCode-Setup-" + "$MSVisualStudioCodeChannelClear" + "-$MSVisualStudioCodePlatformClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MSVisualStudioCodeChannelClear $MSVisualStudioCodeArchitectureClear $MSVisualStudioCodeModeClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MSVisualStudioCodeV"
        If ($MSVisualStudioCodeV -ne $Version) {
            Write-Host -ForegroundColor Green "Update available"
            DS_WriteLog "I" "Install $Product $Product $MSVisualStudioCodeChannelClear $MSVisualStudioCodeArchitectureClear $MSVisualStudioCodeModeClear" $LogFile
            $Options = @(
                "/VERYSILENT"
                "/MERGETASKS=!runcode"
            )
            Try {
                Write-Host "Starting install of $Product $MSVisualStudioCodeChannelClear $MSVisualStudioCodeArchitectureClear $MSVisualStudioCodeModeClear $Version"
                If ($WhatIf -eq '0') {
                    $null = Start-Process "$PSScriptRoot\$Product\$MSVisualStudioCodeInstall" -ArgumentList $Options -NoNewWindow -PassThru
                    while (Get-Process -Name $MSVisualStudioCodeProcess -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code*") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code*" -Recurse -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio Code*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio Code*" -Recurse -Force}
                    }
                }
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MSVisualStudioCodeChannelClear $MSVisualStudioCodeArchitectureClear $MSVisualStudioCodeModeClear (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install MindView 7
    If ($MindView7 -eq 1) {
        $Product = "MindView 7"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MindView7LanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $MindView7D.Version
        }
        $MindView7V = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*MindView 7*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$MindView7V) {
            $MindView7V = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*MindView 7*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $MindView7Installer = "MindView7-" + "$MindView7LanguageClear" + "." + "exe"
        $MindView7Process = "MindView7-" + "$MindView7LanguageClear"
        Write-Host -ForegroundColor Magenta "Install $Product $MindView7LanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $MindView7V"
        If ($MindView7V -lt $Version) {
            Write-Host -ForegroundColor Green "Update available"
            DS_WriteLog "I" "Install $Product $Product $MindView7LanguageClear" $LogFile
            $Options = @(
                "/quiet"
            )
            Try {
                Write-Host "Starting install of $Product $MindView7LanguageClear $Version"
                If ($WhatIf -eq '0') {
                    $null = Start-Process "$PSScriptRoot\$Product\$MindView7Installer" -ArgumentList $Options -NoNewWindow -PassThru
                    while (Get-Process -Name $MindView7Process -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                    If (Test-Path -Path "$env:PUBLIC\Desktop\MindView*.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\MindView*.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\MindView*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\MindView*" -Recurse -Force}
                    }
                }
                Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $MindView7LanguageClear (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Mozilla Firefox
    If ($Firefox -eq 1) {
        $Product = "Mozilla Firefox"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FirefoxChannelClear" + "$FirefoxArchitectureClear" + "$FFLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $FirefoxD.Version
        }
        $FirefoxV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Firefox*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $FirefoxLog = "$LogTemp\Firefox.log"
        If (!$FirefoxV) {
            $FirefoxV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Firefox*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $FirefoxInstaller = "Firefox_Setup_" + "$FirefoxChannelClear" + "$FirefoxArchitectureClear" + "_$FFLanguageClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$FirefoxInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $FirefoxChannelClear $FirefoxArchitectureClear $FFLanguageClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $FirefoxV"
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
            DS_WriteLog "I" "Install $Product $FirefoxChannelClear $FirefoxArchitectureClear $FFLanguageClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $FirefoxChannelClear $FirefoxArchitectureClear $FFLanguageClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Firefox.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Firefox.lnk" -Force}
                    }
                }
                Get-Content $FirefoxLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $FirefoxLog -ErrorAction SilentlyContinue
            } Catch {
                DS_WriteLog "E" "Error installing $Product $FirefoxChannelClear $FirefoxArchitectureClear $FFLanguageClear (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install mRemoteNG
    If ($mRemoteNG -eq 1) {
        $Product = "mRemoteNG"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $mRemoteNGD.Version
        }
        $mRemoteNGV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "mRemoteNG"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$mRemoteNGV) {
            $mRemoteNGV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "mRemoteNG"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $mRemoteLog = "$LogTemp\mRemote.log"
        If ($mRemoteNGV) {$mRemoteNGV = $mRemoteNGV -replace ".{6}$"}
        $InstallMSI = "$PSScriptRoot\$Product\mRemoteNG.msi"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $mRemoteNGV"
        If ($mRemoteNGV -lt $Version) {
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
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                }
                Get-Content $mRemoteLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $mRemoteLog -ErrorAction SilentlyContinue
            } Catch {
                DS_WriteLog "E" "Error installing $Product (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            If ($WhatIf -eq '0') {
                If (Test-Path -Path "$env:PUBLIC\Desktop\mRemoteNG.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\mRemoteNG.lnk" -Force}
                If ($CleanUpStartMenu) {
                    If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\mRemoteNG") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\mRemoteNG" -Recurse -Force}
                }
            }
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Nmap
    If ($Nmap -eq 1) {
        $Product = "Nmap"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $NMapD.Version
        }
        $NmapV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Nmap*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$NmapV) {
            $NmapV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Nmap*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $NmapInstaller = "Nmap-setup.exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $NmapV"
        If ($NmapV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$NmapInstaller"
                }
                $p = Get-Process Nmap-setup -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Nmap") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Nmap" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Notepad ++
    If ($NotePadPlusPlus -eq 1) {
        $Product = "NotepadPlusPlus"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$NotepadPlusPlusArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $NotepadD.Version
        }
        $Notepad = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$Notepad) {
            $Notepad = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $NotepadPlusPlusInstaller = "NotePadPlusPlus_" + "$NotepadPlusPlusArchitectureClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $NotepadPlusPlusArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $Notepad"
        If ($Notepad -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $NotepadPlusPlusArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$NotepadPlusPlusInstaller" -ArgumentList /S -NoNewWindow
                }
                $p = Get-Process NotePadPlusPlus_$NotepadPlusPlusArchitectureClear -ErrorAction SilentlyContinue
		        If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Notepad++.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Notepad++.lnk" -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install openJDK
    If ($OpenJDK -eq 1) {
        $Product = "open JDK"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$openJDKArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $OpenJDKD.Version
        }
        $OpenJDKV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $openJDKLog = "$LogTemp\OpenJDK.log"
        If (!$OpenJDKV) {
            $OpenJDKV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $OpenJDKInstaller = "OpenJDK" + "$openJDKArchitectureClear" + ".msi"
        If ($Version) {$Version = $Version -replace ".-"}
        $InstallMSI = "$PSScriptRoot\$Product\$OpenJDKInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $openJDKArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $OpenJDKV"
        If ($OpenJDKV -lt $Version) {
            DS_WriteLog "I" "Installing $Product $openJDKArchitectureClear" $LogFile
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
                Write-Host "Starting install of $Product $openJDKArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                }
                Get-Content $openJDKLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $openJDKLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Open-Shell Menu
    If ($OpenShellMenu -eq 1) {
        $Product = "Open-Shell Menu"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $OpenShellMenuD.Version
        }
        $OpenShellMenuV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Open-Shell*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$OpenShellMenuV) {
            $OpenShellMenuV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Open-Shell*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $OpenShellMenuInstaller = "OpenShellSetup.exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $OpenShellMenuV"
        If ($OpenShellMenuV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/quiet"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$OpenShellMenuInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process OpenShellSetup -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Open-Shell") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Open-Shell" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install OracleJava8
    If ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$OracleJava8ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $OracleJava8D.Version
        }
        $OracleJava = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java 8*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$OracleJava) {
            $OracleJava = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java 8*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($OracleJava) {
            $OracleJavaSplit = $OracleJava.split(".")
            $OracleJavaSplit2 = $OracleJavaSplit[2].split("0")
            $OracleJava = "1." + $OracleJavaSplit[0] + "." + $OracleJavaSplit[1] + "_" + $OracleJavaSplit2[0] + "-b" + $OracleJavaSplit[3]
        }
        $OracleJavaInstaller = "OracleJava8_" + "$OracleJava8ArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $OracleJava8ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $OracleJava"
        If ($OracleJava -lt $Version) {
            DS_WriteLog "I" "Install $Product $OracleJava8ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/s INSTALL_SILENT=Enable AUTO_UPDATE=Disable REBOOT=Disable SPONSORS=Disable REMOVEOUTOFDATEJRES=1 WEB_ANALYTICS=Disable"
            )
            Try {
                Write-Host "Starting install of $Product $OracleJava8ArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$OracleJavaInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process OracleJava8_$OracleJava8ArchitectureClear -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Java") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Java" -Recurse -Force}
                    }
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $OracleJava8ArchitectureClear (Error: $($Error[0]))"
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Paint.Net
    If ($PaintDotNet -eq 1) {
        $Product = "Paint Dot Net"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $PaintDotNetD.Version
        }
        If ($Version) {
            $VersionSplit = $Version.split(".")
            $VersionSplit2 = $VersionSplit[1] -split("",3)
            $Version = $VersionSplit[0] + "." + $VersionSplit2[1] + "." + $VersionSplit2[2]
        }
        $PaintDotNetV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Paint.Net*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$PaintDotNetV) {
            $PaintDotNetV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Paint.Net*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $PaintDotNetV"
        If ($PaintDotNetV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/auto DESKTOPSHORTCUT=0 CHECKFORUPDATES=0"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\paint.net.install.exe" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process paint.net.install -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\paint.net.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\paint.net.lnk" -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install pdfforge PDFCreator
    If ($PDFForgeCreator -eq 1) {
        $Product = "pdfforge PDFCreator"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$pdfforgePDFCreatorChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $PDFForgeCreatorD.Version
        }
        $PDFForgeCreatorV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "PDFCreator*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$PDFForgeCreatorV) {
            $PDFForgeCreatorV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "PDFCreator*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $PDFForgeCreatorInstaller = "PDFForgeCreatorWebSetup_" + "$pdfforgePDFCreatorChannelClear" + ".exe"
        $PDFForgeCreatorProcess = "PDFForgeCreatorWebSetup" + "$pdfforgePDFCreatorChannelClear"
        Write-Host -ForegroundColor Magenta "Install $Product $pdfforgePDFCreatorChannelClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $PDFForgeCreatorV"
        If ($PDFForgeCreatorV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/NORESTART /NoIcons"
            )
            Try {
                Write-Host "Starting install of $Product $pdfforgePDFCreatorChannelClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$PDFForgeCreatorInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process $PDFForgeCreatorProcess -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PDFCreator.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PDFCreator.lnk" -Force}
                    }
                }
            } Catch {
                Write-Host -ForegroundColor Red "Error installing $Product $pdfforgePDFCreatorChannelClear (Error: $($Error[0]))"
                DS_WriteLog "E" "Error installing $Product $pdfforgePDFCreatorChannelClear (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install PDF Split & Merge
    If ($PDFsam -eq 1) {
        $Product = "PDF Split & Merge"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $PDFsamD.Version
        }
        $PDFsamV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "PDFsam*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $PDFsamLog = "$LogTemp\PDFsam.log"
        If (!$PDFsamV) {
            $PDFsamV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "PDFsam*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $PDFsamInstaller = "PDFsam.msi"
        $InstallMSI = "$PSScriptRoot\$Product\$PDFsamInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $PDFsamV"
        If ($PDFsamV -lt $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $PDFsamLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    If (Test-Path -Path "$env:PUBLIC\Desktop\PDFsam Basic.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\PDFsam Basic.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PDFsam*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PDFsam*" -Recurse -Force}
                    }
                }
                Get-Content $PDFsamLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $PDFsamLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install PeaZip
    If ($PeaZip -eq 1) {
        $Product = "PeaZip"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$PeaZipArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $PeaZipD.Version
        }
        $PeaZipV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "PeaZip*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$PeaZipV) {
            $PeaZipV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "PeaZip*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $PeaZipInstaller = "PeaZip" + "$PeaZipArchitectureClear" + ".exe"
        $PeaZipProcess = "PeaZip" + "$PeaZipArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $PeaZipArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $PeaZipV"
        If ($PeaZipV -lt $Version) {
            DS_WriteLog "I" "Install $Product $PeaZipArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/VERYSILENT"
            )
            Try {
                Write-Host "Starting install of $Product $PeaZipArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$PeaZipInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process $PeaZipProcess -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PeaZip") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PeaZip" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install PuTTY
    If ($PuTTY -eq 1) {
        $Product = "PuTTY"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$PuTTYArchitectureClear" + "_$PuttyChannelClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $PuTTYD.Version
        }
        $PuTTYV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PuTTY*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $PuTTYLog = "$LogTemp\PuTTY.log"
        If (!$PuTTYV) {
            $PuTTYV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*PuTTY*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $PuTTYInstaller = "PuTTY-" + "$PuTTYArchitectureClear" + "-$PuttyChannelClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$PuTTYInstaller"
        If ($PuTTYV) {
            $PuTTYV = $PuTTYV.Split("\.",3)
            $PuTTYV = $PuTTYV[0] + "." + $PuTTYV[1]
        }
        Write-Host -ForegroundColor Magenta "Install $Product $PuttyChannelClear $PuTTYArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $PuTTYV"
        If ($PuTTYV -ne $Version) {
            DS_WriteLog "I" "Installing $Product $PuttyChannelClear $PuTTYArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $PuTTYLog"
            )
            Try {
                Write-Host "Starting install of $Product $PuttyChannelClear $PuTTYArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    Get-Content $PuTTYLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $PuTTYLog -ErrorAction SilentlyContinue
                    If (Test-Path -Path "$env:PUBLIC\Desktop\PuTTY.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\PuTTY.lnk" -Force}
                    If (Test-Path -Path "$env:PUBLIC\Desktop\PuTTY (64-bit).lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\PuTTY (64-bit).lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PuTTY*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\PuTTY*" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Remote Desktop Manager
    If ($RemoteDesktopManager -eq 1) {
        Switch ($RemoteDesktopManagerType) {
            0 {
                $Product = "RemoteDesktopManager Free"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
                If (!($Version)) {
                    $Version = $RemoteDesktopManagerFreeD
                }
                $RemoteDesktopManagerFree = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Remote Desktop Manager*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
                $RemoteDesktopManagerLog = "$LogTemp\RemoteDesktopManager.log"
                $InstallMSI = "$PSScriptRoot\$Product\Setup.RemoteDesktopManagerFree.msi"
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $RemoteDesktopManagerFree"
                If ($RemoteDesktopManagerFree -lt $Version) {
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
                        If ($WhatIf -eq '0') {
                            Install-MSI $InstallMSI $Arguments
                            If ($CleanUpStartMenu) {
                                If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Remote Desktop Manager*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Remote Desktop Manager*" -Recurse -Force}
                            }
                        }
                        Get-Content $RemoteDesktopManagerLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                        Remove-Item $RemoteDesktopManagerLog -ErrorAction SilentlyContinue
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
                If ($CleanUp -eq '1') {
                    If ($WhatIf -eq '0') {
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    }
                    Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
            }
            1 {
                $Product = "RemoteDesktopManager Enterprise"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
                If (!($Version)) {
                    $Version = $RemoteDesktopManagerEnterpriseD
                }
                $RemoteDesktopManagerEnterprise = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Remote Desktop Manager*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
                $RemoteDesktopManagerLog = "$LogTemp\RemoteDesktopManager.log"
                $InstallMSI = "$PSScriptRoot\$Product\Setup.RemoteDesktopManagerEnterprise.msi"
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $RemoteDesktopManagerEnterprise"
                If ($RemoteDesktopManagerEnterprise -lt $Version) {
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
                        If ($WhatIf -eq '0') {
                            Install-MSI $InstallMSI $Arguments
                        }
                        Get-Content $RemoteDesktopManagerLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                        Remove-Item $RemoteDesktopManagerLog -ErrorAction SilentlyContinue
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
                If ($CleanUp -eq '1') {
                    If ($WhatIf -eq '0') {
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    }
                    Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                    DS_WriteLog "-" "" $LogFile
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
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $ShareXD.Version
        }
        $ShareXV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*ShareX*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$ShareXV) {
            $ShareXV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*ShareX*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $ShareXInstaller = "ShareX-setup" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $ShareXV"
        If ($ShareXV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/VERYSILENT"
                "/UPDATE"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$ShareXInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process ShareX-setup -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\ShareX") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\ShareX" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Slack
    If ($Slack -eq 1) {
        $Product = "Slack"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$SlackArchitectureClear" + "_$SlackPlatformClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $SlackD.Version
        }
        $SlackV = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Slack"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$SlackV) {
            $SlackV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Slack"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If (!$SlackV) {
            $SlackV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Slack"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If (!$SlackV) {
        }
        Else {
            If ($SlackV.length -ne "6") {$SlackV = $SlackV -replace ".{2}$"}
        }
        $SlackLog = "$LogTemp\Slack.log"
        $SlackInstaller = "Slack.setup" + "_$SlackArchitectureClear" + "_$SlackPlatformClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$SlackInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $SlackArchitectureClear $SlackPlatformClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $SlackV"
        If ($SlackV -ne $Version) {
            DS_WriteLog "I" "Installing $Product $SlackArchitectureClear $SlackPlatformClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/norestart"
                "/L*V $SlackLog"
            )
            Try {
                Write-Host "Starting install of $Product $SlackArchitectureClear $SlackPlatformClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Slack*") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Slack*" -Recurse -Force}
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Slack*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Slack*" -Recurse -Force}
                    }
                }
                Get-Content $SlackLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $SlackLog -ErrorAction SilentlyContinue
            } Catch {
                DS_WriteLog "E" "Error installing $Product $SlackArchitectureClear $SlackPlatformClear (Error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Host -ForegroundColor Cyan "No update available for $Product"
            Write-Output ""
        }
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }            
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Sumatra PDF
    If ($SumatraPDF -eq 1) {
        $Product = "Sumatra PDF"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$SumatraPDFArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $SumatraPDFD.Version
        }
        $SumatraPDFV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "SumatraPDF"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$SumatraPDFV) {
            $SumatraPDFV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "SumatraPDF"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $SumatraPDFInstaller = "SumatraPDF-Install-" + "$SumatraPDFArchitectureClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $SumatraPDFArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $SumatraPDFV"
        If ($SumatraPDFV -ne $Version) {
            DS_WriteLog "I" "Install $Product $SumatraPDFArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "-quiet"
                "-s"
            )
            Try {
                Write-Host "Starting install of $Product $SumatraPDFArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$SumatraPDFInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    If ($WhatIf -eq '0') {
                        If (Test-Path "$env:USERPROFILE\Desktop\SumatraPDF.lnk") {Remove-Item -Path "$env:USERPROFILE\Desktop\SumatraPDF.lnk" -Force}
                        If ($CleanUpStartMenu) {
                            If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\SumatraPDF.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\SumatraPDF.lnk" -Force}
                        }
                    }
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If (Test-Path -Path "$env:LOCALAPPDATA\SumatraPDF\SumatraPDF-settings.txt") {
                    # Disable auto update
                    Write-Host "Disable auto update"
                    Try {
                        If ($WhatIf -eq '0') {
                            (Get-Content "$env:LOCALAPPDATA\SumatraPDF\SumatraPDF-settings.txt" -ErrorAction SilentlyContinue) | ForEach-Object { $_ -replace "CheckForUpdates = true" , "CheckForUpdates = false" } | Set-Content "$env:LOCALAPPDATA\SumatraPDF\SumatraPDF-settings.txt"
                        }
                        Write-Host -ForegroundColor Green "Disable auto update $Product finished!"
                    } Catch {
                        Write-Host -ForegroundColor Red "Error disable auto update (Error: $($Error[0]))"
                        DS_WriteLog "E" "Error disable auto update (Error: $($Error[0]))" $LogFile
                    }
                }
                DS_WriteLog "-" "" $LogFile
                Write-Output ""
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install TeamViewer
    If ($TeamViewer -eq 1) {
        $Product = "TeamViewer"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $TeamViewerD.Version
        }
        $TeamViewerV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TeamViewer*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $TeamViewerInstaller = "TeamViewer-setup" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $TeamViewerV"
        If ($TeamViewerV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/S"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$TeamViewerInstaller" -ArgumentList $Options -NoNewWindow
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                $p = Get-Process TeamViewer-setup -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                    If ($WhatIf -eq '0') {
                        If (Test-Path -Path "$env:PUBLIC\Desktop\Teamviewer.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Teamviewer.lnk" -Force}
                        If ($CleanUpStartMenu) {
                            If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TeamViewer.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TeamViewer.lnk" -Force}
                        }
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install TechSmith Camtasia
    If ($TechSmithCamtasia -eq 1) {
        $Product = "TechSmith Camtasia"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version.txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $TechSmithCamtasiaD.Version
        }
        $TechSmithCamtasiaV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Camtasia*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $TechSmithCamtasiaLog = "$LogTemp\TechSmithCamtasia.log"
        If (!$TechSmithCamtasiaV) {
            $TechSmithCamtasiaV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Camtasia*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($TechSmithCamtasiaV) {$TechSmithCamtasiaSplit = $TechSmithCamtasiaV.split(".")}
        If ($TechSmithCamtasiaSplit) {$TechSmithCamtasiaStrings = ([regex]::Matches($TechSmithCamtasiaV, "\." )).count}
        Switch ($TechSmithCamtasiaStrings) {
            1 {
                $TechSmithCamtasiaVN = $TechSmithCamtasiaSplit[0] + "." + $TechSmithCamtasiaSplit[1]
            }
            2 {
                $TechSmithCamtasiaVN = $TechSmithCamtasiaSplit[0] + "." + $TechSmithCamtasiaSplit[1] + "." + $TechSmithCamtasiaSplit[2]
            }
            3 {
                $TechSmithCamtasiaVN = $TechSmithCamtasiaSplit[0] + "." + $TechSmithCamtasiaSplit[1] + "." + $TechSmithCamtasiaSplit[2]
            }
        }
        $TechSmithCamtasiaInstaller = "camtasia-setup.msi"
        $InstallMSI = "$PSScriptRoot\$Product\$TechSmithCamtasiaInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $TechSmithCamtasiaVN"
        If ($TechSmithCamtasiaVN -ne $Version) {
            DS_WriteLog "I" "Installing $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $TechSmithCamtasiaLog"
            )
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    Get-Content $TechSmithCamtasiaLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $TechSmithCamtasiaLog -ErrorAction SilentlyContinue
                    If (Test-Path -Path "$env:PUBLIC\Desktop\Camtasia*.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Camtasia*.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TechSmith") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TechSmith" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install TechSmith SnagIt
    If ($TechSmithSnagIt -eq 1) {
        $Product = "TechSmith SnagIt"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $TechSmithSnagitD.Version
        }
        $TechSmithSnagItV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "SnagIt*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $TechSmithSnagItLog = "$LogTemp\TechSmithSnagIt.log"
        If (!$TechSmithSnagItV) {
            $TechSmithSnagItV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "SnagIt*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $TechSmithSnagItInstaller = "snagit-setup_" + "$ArchitectureClear" + ".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$TechSmithSnagItInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $TechSmithSnagItV"
        If ($TechSmithSnagItV -ne $Version) {
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Arguments = @(
                "/i"
                "`"$InstallMSI`""
                "/qn"
                "/L*V $TechSmithSnagItLog"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Start-Sleep 25
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TechSmith") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TechSmith" -Recurse -Force}
                    }
                }
                Get-Content $TechSmithSnagItLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                Remove-Item $TechSmithSnagItLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Total Commander
    If ($TotalCommander -eq 1) {
        $Product = "Total Commander"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$TotalCommanderArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $TotalCommanderD.Version
        }
        $TotalCommanderV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Total Commander*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$TotalCommanderV) {
            $TotalCommanderV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "Total Commander*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $TotalCommanderInstaller = "TotalCommander_" + "$TotalCommanderArchitectureClear" + ".exe"
        $TotalCommanderProcess = "TotalCommander_" + "$TotalCommanderArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $TotalCommanderArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $TotalCommanderV"
        If ($TotalCommanderV -lt $Version) {
            DS_WriteLog "I" "Install $Product $TotalCommanderArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/A1H1L1G1U1"
            )
            Try {
                Write-Host "Starting install of $Product $TotalCommanderArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$TotalCommanderInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process $TotalCommanderProcess -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Total Commander") {Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Total Commander" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install TreeSize
    If ($TreeSize -eq 1) {
        Switch ($TreeSizeType) {
            0 {
                $Product = "TreeSize Free"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
                If (!($Version)) {
                    $Version = $TreeSizeFreeD.Version
                }
                $Version = $Version.Insert(3,'.')
                $TreeSizeV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $TreeSizeV"
                If ($TreeSizeV -lt $Version) {
                    DS_WriteLog "I" "Install $Product" $LogFile
                    Write-Host -ForegroundColor Green "Update available"
                    Try {
                        Write-Host "Starting install of $Product $Version"
                        If ($WhatIf -eq '0') {
                            Start-Process "$PSScriptRoot\$Product\TreeSize_Free.exe" -ArgumentList /VerySilent -NoNewWindow
                        }
                        $p = Get-Process TreeSize_Free -ErrorAction SilentlyContinue
                        If ($p) {
                            $p.WaitForExit()
                            Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                        }
                        If ($WhatIf -eq '0') {
                            If ($CleanUpStartMenu) {
                                If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TreeSize*") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\TreeSize*" -Recurse -Force}
                            }
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
                If ($CleanUp -eq '1') {
                    If ($WhatIf -eq '0') {
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    }
                    Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
            }
            1 {
                $Product = "TreeSize Professional"
                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
                If (!($Version)) {
                    $Version = $TreeSizeProfD.Version
                }
                $Version = $Version.Insert(3,'.')
                $TreeSizeV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
                Write-Host -ForegroundColor Magenta "Install $Product"
                Write-Host "Download Version: $Version"
                Write-Host "Current Version:  $TreeSizeV"
                If ($TreeSizeV -lt $Version) {
                    DS_WriteLog "I" "Install $Product" $LogFile
                    Write-Host -ForegroundColor Green "Update available"
                    Try {
                        Write-Host "Starting install of $Product $Version"
                        If ($WhatIf -eq '0') {
                            Start-Process "$PSScriptRoot\$Product\TreeSize_Professional.exe" -ArgumentList /VerySilent -NoNewWindow
                        }
                        $p = Get-Process TreeSize_Professional -ErrorAction SilentlyContinue
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
                If ($CleanUp -eq '1') {
                    If ($WhatIf -eq '0') {
                        Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    }
                    Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
                    DS_WriteLog "-" "" $LogFile
                    Write-Output ""
                }
            }
        }
    }

    #// Mark: Install uberAgent
    If ($uberAgent -eq 1) {
        $Product = "uberAgent"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $uberAgentD.Version
        }
        $uberAgentV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*uberAgent*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$uberAgentV) {
            $uberAgentV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*uberAgent*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $uberAgentInstaller = "silent-install.cmd"
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $uberAgentV"
        If ($uberAgentV -lt $Version) {
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$uberAgentInstaller" -NoNewWindow -Wait
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install VLC Player
    If ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $VLCD.Version
        }
        $VLC = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        $VLCLog = "$LogTemp\VLC.log"
        If (!$VLC) {
            $VLC = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($VLC) {$VLC = $VLC -replace ".{2}$"}
        $VLCInstaller = "VLC-Player_" + "$ArchitectureClear" +".msi"
        $InstallMSI = "$PSScriptRoot\$Product\$VLCInstaller"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $VLC"
        If ($VLC -lt $Version) {
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
                If ($WhatIf -eq '0') {
                    Install-MSI $InstallMSI $Arguments
                    Get-Content $VLCLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $VLCLog -ErrorAction SilentlyContinue
                    If (Test-Path -Path "$env:PUBLIC\Desktop\VLC media player.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\VLC media player.lnk" -Force}
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\VideoLAN") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\VideoLAN" -Recurse -Force}
                    }
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install VMWareTools
    If ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $VMWareToolsD.Version
        }
        $VMWT = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$VMWT) {
            $VMWT = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        If ($VMWT) {$VMWT = $VMWT -replace ".{9}$"}
        $VMWareToolsInstaller = "VMWareTools_" + "$ArchitectureClear" +".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $VMWT"
        If ($VMWT -lt $Version) {
            $Options = @(
                "/s"
                "/v"
                "/qn REBOOT=Y"
            )
            DS_WriteLog "I" "Install $Product" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$VMWareToolsInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                    Write-Host -ForegroundColor Yellow "System needs to reboot after installation!"
                }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install WinMerge
    If ($WinMerge -eq 1) {
        $Product = "WinMerge"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $WinMergeD.Version
        }
        $WinMergeV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "WinMerge*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$WinMergeV) {
            $WinMergeV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "WinMerge*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $WinMergeInstaller = "WinMerge_" + "$ArchitectureClear" + ".exe"
        $WinMergeProcess = "WinMerge_" + "$ArchitectureClear"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $WinMergeV"
        If ($WinMergeV -lt $Version) {
            DS_WriteLog "I" "Install $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/VERYSILENT"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    Start-Process "$PSScriptRoot\$Product\$WinMergeInstaller" -ArgumentList $Options -NoNewWindow
                }
                $p = Get-Process $WinMergeProcess -ErrorAction SilentlyContinue
                If ($p) {
                    $p.WaitForExit()
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\WinMerge") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\WinMerge" -Recurse -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install WinSCP
    If ($WinSCP -eq 1) {
        $Product = "WinSCP"
        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $WinSCPD.Version
        }
        $WSCP = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*WinSCP*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        Write-Host -ForegroundColor Magenta "Install $Product"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $WSCP"
        If ($WSCP -lt $Version) {
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
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\WinSCP.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                    If ($WhatIf -eq '0') {
                        If (Test-Path -Path "$env:PUBLIC\Desktop\WinSCP.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\WinSCP.lnk" -Force}
                        If ($CleanUpStartMenu) {
                            If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\WinSCP.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\WinSCP.lnk" -Force}
                        }
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Wireshark
    If ($Wireshark -eq 1) {
        $Product = "Wireshark"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
        If (!($Version)) {
            $Version = $WiresharkD.Version
        }
        $WiresharkV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Wireshark*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        If (!$WiresharkV) {
            $WiresharkV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Wireshark*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
        }
        $WiresharkInstaller = "Wireshark-" + "$ArchitectureClear" + ".exe"
        Write-Host -ForegroundColor Magenta "Install $Product $ArchitectureClear"
        Write-Host "Download Version: $Version"
        Write-Host "Current Version:  $WiresharkV"
        If ($WiresharkV -lt $Version) {
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            Write-Host -ForegroundColor Green "Update available"
            $Options = @(
                "/S"
                "/desktopicon=no"
                "/quicklaunchdicon=no"
            )
            Try {
                Write-Host "Starting install of $Product $ArchitectureClear $Version"
                If ($WhatIf -eq '0') {
                    $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$WiresharkInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                }
                else{
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($inst) {
                    Wait-Process -InputObject $inst
                    Write-Host -ForegroundColor Green "Install of the new version $Version finished!"
                }
                If ($WhatIf -eq '0') {
                    If ($CleanUpStartMenu) {
                        If (Test-Path -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Wireshark.lnk") {Remove-Item -Path "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Wireshark.lnk" -Force}
                    }
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }

    #// Mark: Install Zoom
    If ($Zoom -eq 1) {
        If ($ZoomInstallerClear -eq 'Machine Based') {
            $Product = "Zoom VDI"
            # Check, if a new version is available
            $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
            $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $ZoomVersion
            }
            $ZoomV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Zoom Client for VDI*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            If ($ZoomV.length -ne "5") {$ZoomV = $ZoomV -replace ".{4}$"}
            $ZoomLog = "$LogTemp\Zoom.log"
            $ZoomInstaller = "ZoomInstaller" + ".msi"
            $InstallMSI = "$PSScriptRoot\$Product\$ZoomInstaller"
            Write-Host -ForegroundColor Magenta "Install $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $ZoomV"
            If ($ZoomV -lt $Version) {
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
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                        Start-Sleep 25
                        Get-Content $ZoomLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                        Remove-Item $ZoomLog -ErrorAction SilentlyContinue
                        If (Test-Path -Path "$env:PUBLIC\Desktop\Zoom VDI.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Zoom VDI.lnk" -Force}
                    }
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
        If ($ZoomInstallerClear -eq 'User Based') {
            $Product = "Zoom"
            # Check, if a new version is available
            $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
            $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $ZoomD.Version
            }
            $ZoomV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Zoom*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            If ($ZoomV.length -ne "5") {$ZoomV = $ZoomV -replace ".{4}$"}
            $ZoomLog = "$LogTemp\Zoom.log"
            $ZoomInstaller = "ZoomInstaller" + ".msi"
            $InstallMSI = "$PSScriptRoot\$Product\$ZoomInstaller"
            Write-Host -ForegroundColor Magenta "Install $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $ZoomV"
            If ($ZoomV -lt $Version) {
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
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                        Start-Sleep 25
                        Get-Content $ZoomLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                        Remove-Item $ZoomLog -ErrorAction SilentlyContinue
                        If (Test-Path -Path "$env:PUBLIC\Desktop\Zoom.lnk") {Remove-Item -Path "$env:PUBLIC\Desktop\Zoom.lnk" -Force}
                    }
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
    If ($Zoom -eq 1) {
        If ($ZoomCitrixClient -eq 1) {
            $Product = "Zoom Citrix Client"
            # Check, if a new version is available
            $VersionPath = "$PSScriptRoot\$Product\Version" + ".txt"
            $Version = Get-Content -Path "$VersionPath" -ErrorAction SilentlyContinue
            If (!($Version)) {
                $Version = $ZoomD.Version
            }
            $ZoomV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Zoom Plugin*"}).DisplayVersion | Sort-Object -Property Version -Descending | Select-Object -First 1
            If ($ZoomV.length -ne "5") {$ZoomV = $ZoomV -replace ".{4}$"}
            $ZoomInstaller = "ZoomCitrixHDXMediaPlugin" + ".msi"
            $ZoomLog = "$LogTemp\Zoom.log"
            $InstallMSI = "$PSScriptRoot\$Product\$ZoomInstaller"
            Write-Host -ForegroundColor Magenta "Install $Product"
            Write-Host "Download Version: $Version"
            Write-Host "Current Version:  $ZoomV"
            If ($ZoomV -lt $Version) {
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
                    If ($WhatIf -eq '0') {
                        Install-MSI $InstallMSI $Arguments
                        Start-Sleep 25
                    }
                    Get-Content $ZoomLog -ErrorAction SilentlyContinue | Add-Content $LogFile -Encoding ASCI -ErrorAction SilentlyContinue
                    Remove-Item $ZoomLog -ErrorAction SilentlyContinue
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
        If ($CleanUp -eq '1') {
            If ($WhatIf -eq '0') {
                Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            }
            Write-Host -ForegroundColor Green "CleanUp for $Product install files successfully."
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
    }
    If ($Installer -eq 0) {
        Write-Host "Disable Change User Mode Install."
        Change User /Execute | Out-Null
    }
}