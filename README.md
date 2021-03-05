# Evergreen Script by Manuel Winkel / [Deyda.net](https://www.deyda.net) / [@deyda84](https://twitter.com/Deyda84)
Download and Install several Software the lazy way with the [Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein](https://github.com/aaronparker/Evergreen).

![https://www.deyda.net/index.php/en/evergreen-script/](/img/EvergreenLeaf.png)

To update or download a software package just switch from 0 to 1 in the section "Select software" (With PowerShell parameter -list) or select your Software out of the GUI.

![Parameter -list](/img/script-list.png)

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
    2022-03-02        Add Microsoft Teams Citrix Api Hook / Correction En dash Error
    2022-03-05        Adjustment regarding merge #122 (Get-AdobeAcrobatReader)
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

    # Software Release / Ring / Channel / Type ?!
    # Citrix Workspace App
    # 0 = Current Release
    # 1 = Long Term Service Release
    $CitrixWorkspaceAppRelease = 1

    # Microsoft 365 Apps
    # 0 = Current (Preview) Channel
    # 1 = Current Channel
    # 2 = Monthly Enterprise Channel
    # 3 = Semi-Annual Enterprise (Preview) Channel
    # 4 = Semi-Annual Enterprise Channel
    $MS365AppsChannel = 4

    # Microsoft OneDrive
    # 0 = Insider Ring
    # 1 = Production Ring
    # 2 = Enterprise Ring
    $MSOneDriveRing = 2

    # Microsoft Teams
    # 0 = Preview Ring
    # 1 = General Ring
    $MSTeamsRing = 1

    # Mozilla Firefox
    # 0 = Current
    # 1 = ESR
    $FirefoxChannel = 0

    # TreeSize
    # 0 = Free
    # 1 = Professional
    $TreeSizeType = 0

    # Select software
    # 0 = Not selected
    # 1 = Selected
    $7ZIP = 0
    $AdobeProDC = 0 # Only Update @ the moment
    $AdobeReaderDC = 0
    $BISF = 0
    $Citrix_Hypervisor_Tools = 0
    $Citrix_WorkspaceApp = 0
    $Filezilla = 0
    $Firefox = 0
    $Foxit_Reader = 0  # No Silent Install
    $FSLogix = 0
    $GoogleChrome = 0
    $Greenshot = 0
    $KeePass = 0
    $mRemoteNG = 0
    $MS365Apps = 0 # Automatically created install.xml is used. Please replace this file if you want to change the installation.
    $MSEdge = 0
    $MSOffice2019 = 0 # Automatically created install.xml is used. Please replace this file if you want to change the installation.
    $MSOneDrive = 0
    $MSTeams = 0
    $NotePadPlusPlus = 0
    $OpenJDK = 0
    $OracleJava8 = 0
    $TreeSize = 0
    $VLCPlayer = 0
    $VMWareTools = 0
    $WinSCP = 0
For example, to automate the process via Scheduled Task or to integrate this into [BIS-F](https://eucweb.com/download-bis-f) (Thx Matthias Schlimm for your work).

### -download

Only download the selected software packages in list Mode (-list).

### -install

Only install the selected software packages in list Mode (-list).


## Example

.\Evergreen.ps1 -list -download

Download the selected Software out of the list.


.\Evergreen.ps1 -list -install

Install the selected Software out of the list.


.\Evergreen.ps1 -list

Download and install the selected Software out of the list.


.\Evergreen.ps1

Starts the GUI to select the mode (Install and/or Download) and the software (Release, Update Ring, Language, etc.).

![GUI-Mode](/img/GUI.png)

## Notes

### Evergreen PowerShell Module
If Download is selected, the module is checked each time the script is run and reinstalled if a new version is available.

### 7-ZIP
Line 746 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For 7-ZIP this is an exe file.

### Adobe Pro DC
Line 780 defines which package is downloaded (You can change the version).   

For Adobe Pro DC this is the update package (msp file).

Only update @ the moment, no installer!

After the update, the Adobe service and scheduled task will be stopped and disabled.

### Adobe Reader DC
Line 813 defines which package is downloaded (You can change the architecture in line 538 and the language in line 533 for non GUI start).

        English
        Danish
        Dutch
        French
        Finnish
        German
        Italian
        Japanese
        Korean
        Norwegian
        Spanish
The architecture can only be changed to x64 for the English package at the moment.

After the installation, the Adobe service and scheduled task will be stopped and disabled.

### BIS-F
Line 846 defines which package is downloaded.

For BIS-F this is the msi file.

After the installation, the scripts will be adjusted regarding task offload, RSS activation and DEP deactivation.

### Citrix Hypervisor Tools
Line 879 defines which package is downloaded (You can change the architecture in line 538 and the version in line 879 for non-GUI startup).

For Citrix Hypervisor Tools this is the x64 msi file (LTSR Path).

For Windows 7, Windows Server 2008 SP2 and Windows Server 2008 R2 SP1, you can change to version 7.2.0.1555 in line 879.

### Citrix WorkspaceApp
Line 913 defines which package is downloaded (You can change the release on line 544 for non GUI start).

Before the installation of the new receiver, the old one is uninstalled via Receiver CleanUp Tool.

The installation is executed with the following parameters (from line 2157):
        /forceinstall
        /silent
        /EnableCEIP=false
        /FORCE_LAA=1
        /AutoUpdateCheck=disabled
        /ALLOWADDSTORE=S
        /ALLOWSAVEPWD=S
        /includeSSON
        /ENABLE_SSON=Yes
After the installation, various registry keys are set (from line 2177).

As always, after installing the new WorkspaceApp, the system should be rebooted.

### Filezilla
Line 954 defines which package is downloaded.

For Filezilla this is the exe file.

Filezilla is installed with the parameter /user=all for all users.

### Foxit Reader
Line 987 defines which package is downloaded (You can change the language in line 533 for non GUI start).

        Danish
        Dutch
        English
        Finnish
        French
        German
        Italian
        Korean
        Norwegian
        Polish
        Portuguese
        Russian
        Spanish
        Swedish
For Foxit Reader this is an exe file.

Unfortunately, a silent install is not possible at the moment.

### GoogleChrome
Line 1054 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For Google Chrome this is the msi file.

After the installation the Chrome services and scheduled tasks will be stopped and disabled.

### Greenshot
Line 1021 defines which package is downloaded.

For Greenshot this is an exe file.

### KeePass
Line 1088 defines which package is downloaded.

For KeePass this is the msi file.

### Microsoft 365 Apps
Line 1120 defines which package is downloaded (You can change the channel in line 552 for non GUI start).

For Microsoft 365 Apps this is the exe setup file.

During the download not only the setup.exe is downloaded, but also the following xml files are created, if they are not already present in the folder:

remove.xml (from line 1132)

install.xml (from line 1157)

Afterwards the install.xml is used to download the required install files.

Before installing the new Microsoft 365 Apps version, the previous installation is removed (remove.xml).

After that the reinstall of the software starts (install.xml).

An install.xml with the special features of the own installation can be stored and used in advance (e.g. Languages, App Exclusion or Inclusion (Visio & Project)).

By default, the following is defined in install.xml (64Bit / Match OS Language / Semi Annual Channel):

    <Configuration>
      <Add Channel="SemiAnnual" OfficeClientEdition="64" SourcePath="<Path to Evergreen Folder>\MS 365 Apps (Semi Annual Channel)">
        <Product ID="O365ProPlusRetail">
          <Language ID="MatchOS" Fallback="en-us"/>
          <ExcludeApp ID="Teams"/>
          <ExcludeApp ID="Lync"/>
          <ExcludeApp ID="Groove"/>
          <ExcludeApp ID="OneDrive"/>
        </Product>
      </Add>
      <Display AcceptEULA="TRUE" Level="None"/>
      <Logging Level="Standard" Path="%temp%"/>
      <Property Value="1" Name="SharedComputerLicensing"/>
      <Property Value="TRUE" Name="FORCEAPPSHUTDOWN"/>
      <Updates Enabled="FALSE"/>
    </Configuration>

### Microsoft Edge
Line 1230 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For Microsoft Edge this is the msi file.

Microsoft Edge is installed with the parameter that don't create icons (Desktop and Quickstart).

After the installation, the scheduled tasks of Microsoft Edge are disabled and the Citrix API Hooks are set in the registry.

### Microsoft FSLogix
Line 1265 defines which package is downloaded.

For FSLogix this is the zip package.

With FSLogix installation, the old installation, if present, is uninstalled first and a restart is requested. 

Then the script must be started again, so that the new version is installed cleanly.

Not only the FSLogix Agent is installed, but also the FSLogix AppRule Editor.

### Microsoft Office 2019
Line 1305 defines which package is downloaded.

For Microsoft Office 2019 this is the exe setup file for Office 2019 Enterprise.

During the download not only the setup.exe is downloaded, but also the following xml files are created, if they are not already present in the folder:

remove.xml (from line 1316)

install.xml (from line 1341)

Afterwards the install.xml is used to download the required install files.

Before installing the new Microsoft Office 2019 version, the previous installation is removed (remove.xml).

After that the reinstall of the software starts (install.xml).

An install.xml with the special features of the own installation can be stored and used in advance (e.g. Languages or architecture).

By default, the following is defined in install.xml (64Bit / Match OS Language):

    <Configuration>
      <Add Channel="PerpetualVL2019" OfficeClientEdition="64" SourcePath="<Path to Evergreen Folder>\MS2019">
        <Product ID="ProPlus2019Volume">
          <Language ID="MatchOS" Fallback="en-us"/>
          <ExcludeApp ID="Teams"/>
          <ExcludeApp ID="Lync"/>
          <ExcludeApp ID="Groove"/>
          <ExcludeApp ID="OneDrive"/>
        </Product>
      </Add>
      <Display AcceptEULA="TRUE" Level="None"/>
      <Logging Level="Standard" Path="%temp%"/>
      <Property Value="1" Name="SharedComputerLicensing"/>
      <Property Value="TRUE" Name="FORCEAPPSHUTDOWN"/>
      <Updates Enabled="FALSE"/>
    </Configuration>
    
### Microsoft OneDrive
Line 1414 defines which package is downloaded (You can change the update ring in line 558 for non GUI start).

For Microsoft OneDrive this is the Production Ring exe file.

Microsoft OneDrive is installed with the Machine Based Install parameter.

### Microsoft Teams
Line 1448 defines which package is downloaded (You can change the architecture in line 538 and update ring in line 563 for non GUI start).
 
For Microsoft Teams this is the x64 msi file (General Ring).

Microsoft Teams is installed with the Machine Based Install parameters.

The registry key "Disable Microsoft Teams Autostart" can be enabled in the script (from line 2816).

### Mozilla Firefox
Line 1482 defines which package is downloaded (You can change the architecture in line 538, language in line 533 and the channel in line 568 for non GUI start).

       Danish
       Dutch
       English
       Finnish
       French
       German
       Italian
       Japanese
       Korean
       Norwegian
       Polish
       Portuguese
       Russian
       Spanish
       Swedish

For Firefox this is the english x64 msi file (Latest Firefox Version).

Firefox is installed with the parameter that disables the creation of the icons and the the maintenance service.
### mRemoteNG
Line 1516 defines which package is downloaded.

For mRemoteNG this is the msi file.

### NotePad++
Line 1549 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For Notepad++ this is the x64 exe file.

### OpenJDK
Line 1583 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For OpenJDK this is the x64 msi file.

### Oracle Java 8
Line 1617 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For Oracle Java 8 this is the x64 msi file.

### TreeSize
Line 1653 and 1684 defines which package is downloaded (You can change the version in line 573 for non GUI start).

For TreeSize this is the exe file.

### VLC Player
Line 1719 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For VLC Player this is the x64 msi file.

### VMWare Tools
Line 1753 defines which package is downloaded (You can change the architecture in line 538 for non GUI start).

For VMWare Tools this is the x64 exe file.

With VMWare Tools installation, the old installation, if present, is uninstalled first and a restart is requested. 

Then the script must be started again, so that the new version is installed cleanly.
    
### WinSCP
Line 1787 defines which package is downloaded.

For WinSCP this is the exe file.

## Shortcut
In GitHub I have placed a sample lnk file under [shortcut](https://github.com/Deyda/Evergreen/tree/main/shortcut), as well as the Evergreen Script logo as an icon file.

Change the path after the -file parameter to the location of your Evergreen Script folder.

![https://www.deyda.net](/img/Logo_DEYDA_with_url.png)
