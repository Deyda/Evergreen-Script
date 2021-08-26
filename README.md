# Evergreen Script by Manuel Winkel / [Deyda.net](https://www.deyda.net) / [@deyda84](https://twitter.com/Deyda84)
Download and Install several Software the lazy way with the [Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein](https://github.com/aaronparker/Evergreen).

![https://www.deyda.net/index.php/en/evergreen-script/](/img/EvergreenLeaf.png)

To update or download a software package just switch from 0 to 1 in the section "Select software" (With PowerShell parameter -list) or select your Software out of the GUI.

![Parameter -list](/img/script-list.png)

Don't forget to set the Software Version, Update Ring, Architecture and so on.

![Parameter -list1](/img/script-list1.png)

![Parameter -list2](/img/script-list2.png)

![Parameter -list3](/img/script-list3.png)


A new folder for every single package will be created, together with a version file and a log file. 

If a new version is available the script checks the version number and will update the package.

I'm no powershell expert, so I'm sure there is much room for improvements!

So let me hear your feedback, I will try to include everything as much as I can.

![Script](/img/script.png)

## Purpose/Change:
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
## Parameter

### -list

Don't start the GUI to select the Software Packages and use the hardcoded list in the script (From line 517). 

If neither parameter -Download or -Install is also used, both processes will be executed.

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
    
    # Select Machine Type (If this is selectable at install or download)
    # 0 = Virtual
    # 1 = Physical
    $Machine = 0

    # Software Release / Ring / Channel / Type ?!
    # Citrix Workspace App
    # 0 = Current Release
    # 1 = Long Term Service Release
    $CitrixWorkspaceAppRelease = 1

    # ControlUp Agent
    # 0 = .Net 3.5 Framework
    # 1 = .Net 4.5 Framework
    $ControlUpAgentFramework = 1

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

    # Microsoft Azure Data Studio
    # 0 = Insider Channel
    # 1 = Stable Channel
    $MSAzureDataStudioChannel = 1

    # Microsoft Edge
    # 0 = Developer Channel
    # 1 = Beta Channel
    # 2 = Stable Channel
    $MSEdgeChannel = 2

    # Microsoft FSLogix
    # 0 = Preview Channel
    # 1 = Production Channel
    $MSFSLogixChannel = 1

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
    # 1 = Exploration Ring
    # 2 = Preview Ring
    # 3 = General Ring
    $MSTeamsRing = 3

    # Microsoft Teams AutoStart
    # 0 = AutoStart Microsoft Teams
    # 1 = No AutoStart (Delete HKLM Registry Entry)
    $MSTeamsNoAutoStart = 0

    # Microsoft Visual Studio
    # 0 = Enterprise Edition
    # 1 = Professional Edition
    # 2 = Community Edition
    $MSVisualStudioEdition = 1

    # Microsoft Visual Studio Code
    # 0 = Insider Channel
    # 1 = Stable Channel
    $MSVisualStudioCodeChannel = 1

    # Microsoft AVD Remote Desktop
    # 0 = Insider Channel
    # 1 = Public Channel
    $MSAVDRemoteDesktopChannel = 1

    # Mozilla Firefox
    # 0 = Current
    # 1 = ESR
    $FirefoxChannel = 0

    # PuTTY
    # 0 = Pre-Release
    # 1 = Stable
    $PuttyChannel = 1

    # Remote Desktop Manager
    # 0 = Free
    # 1 = Enterprise
    $RemoteDesktopManagerType = 0

    # TreeSize
    # 0 = Free
    # 1 = Professional
    $TreeSizeType = 0

    # Zoom
    # 0 = Installer
    # 1 = Installer + Citrix Plugin
    $ZoomCitrixClient = 1

    # Select Software
    # 0 = Not selected
    # 1 = Selected
    $1Password = 0
    $7ZIP = 0
    $AdobeProDC = 0 # Only Update @ the moment
    $AdobeReaderDC = 0
    $BISF = 0
    $CiscoWebexTeams = 0
    $CitrixFiles = 0
    $Citrix_Hypervisor_Tools = 0
    $Citrix_WorkspaceApp = 0
    $ControlUpAgent = 0
    $ControlUpConsole = 0
    $deviceTRUST = 0
    $Filezilla = 0
    $Firefox = 0
    $FoxitPDFEditor = 0
    $Foxit_Reader = 0
    $GIMP = 0
    $GitForWindows = 0
    $GoogleChrome = 0
    $Greenshot = 0
    $ImageGlass = 0
    $IrfanView = 0
    $KeePass = 0
    $LogMeInGoToMeeting = 0
    $mRemoteNG = 0
    $MSDotNetFramework = 0
    $MS365Apps = 0 # Automatically created install.xml is used. Please replace this file if you want to change the installation.
    $MSAVDRemoteDesktop = 0
    $MSAzureCLI = 0
    $MSAzureDataStudio = 0
    $MSEdge = 0
    $MSFSLogix = 0
    $MSOffice2019 = 0 # Automatically created install.xml is used. Please replace this file if you want to change the installation.
    $MSOneDrive = 0
    $MSPowerBIDesktop = 0
    $MSPowerShell = 0
    $MSPowerToys = 0
    $MSSQLServerManagementStudio = 0
    $MSSysinternals = 0
    $MSTeams = 0
    $MSVisualStudio = 0
    $MSVisualStudioCode = 0
    $NMap = 0
    $NotePadPlusPlus = 0
    $OpenJDK = 0
    $OracleJava8 = 0
    $PaintDotNet = 0
    $Putty = 0
    $RDAnalyzer = 0
    $RemoteDesktopManager = 0
    $ShareX = 0
    $Slack = 0
    $SumatraPDF = 0
    $TeamViewer = 0
    $TechSmithCamtasia = 0
    $TechSmithSnagit = 0
    $TreeSize = 0
    $uberAgent = 0
    $VLCPlayer = 0
    $VMWareTools = 0
    $WinSCP = 0
    $Wireshark = 0
    $Zoom = 0
For example, to automate the process via Scheduled Task or to integrate this into [BIS-F](https://eucweb.com/download-bis-f) (Thx Matthias Schlimm for your work).

### -download

Only download the selected software packages in list Mode (-list).

### -install

Only install the selected software packages in list Mode (-list).

### -file

Path to GUI file (LastSettings.txt) for software selection in list Mode.

## Example

.\Evergreen.ps1 -list -download

Download the selected Software out of the list.


.\Evergreen.ps1 -list -install

Install the selected Software out of the list.


.\Evergreen.ps1 -list

Download and install the selected Software out of the list.


.\Evergreen.ps1 -list -file LastSetting.txt

Download and install the selected Software out of the file LastSettings.txt.


.\Evergreen.ps1

Starts the GUI to select the mode (Install and/or Download) and the software (Release, Update Ring, Language, etc.).

![GUI-Mode](/img/GUI.png)

## Shortcut
In GitHub I have placed a sample lnk file under [shortcut](https://github.com/Deyda/Evergreen/tree/main/shortcut), as well as the Evergreen Script logo as an icon file.

Change the path after the -file parameter to the location of your Evergreen Script folder.

![https://www.deyda.net](/img/Logo_DEYDA_with_url.png)
