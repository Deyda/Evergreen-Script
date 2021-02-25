# Evergreen Script by Manuel Winkel / [Deyda.net](https://www.deyda.net) / [@deyda84](https://twitter.com/Deyda84)
Download and Install several Software the lazy way with the [Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein](https://github.com/aaronparker/Evergreen).

![www.deyda.net](/img/Logo_DEYDA_with_url.png)

To update or download a software package just switch from 0 to 1 in the section "Select software" (With PowerShell parameter -list) or select your Software out of the GUI.

A new folder for every single package will be created, together with a version file and a log file. 

If a new version is available the script checks the version number and will update the package.

I'm no powershell expert, so I'm sure there is much room for improvements!

So let me hear your feedback, I will try to include everything as much as I can.

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

## Parameter

### -list

Don't start the GUI to select the Software Packages and use the hardcoded list in the script (From line 549)

    # Select software
    # 0 = Not selected
    # 1 = Selected
    
    $7ZIP = 0
    $AdobeProDC = 0 # Only Update @ the moment
    $AdobeReaderDC = 0
    $BISF = 0
    $Citrix_Hypervisor_Tools = 0
    $Citrix_WorkspaceApp_CR = 0
    $Citrix_WorkspaceApp_LTSR = 0
    $Filezilla = 0
    $Firefox = 0
    $Foxit_Reader = 0  # No Silent Install
    $FSLogix = 0
    $GoogleChrome = 0
    $Greenshot = 0
    $KeePass = 0
    $mRemoteNG = 0
    $MS365Apps = 0 # 64Bit / Match OS Language / Semi Annual Channel
    $MSEdge = 0
    $MSOffice2019 = 0 # 64Bit / Match OS Language
    $MSOneDrive = 0
    $MSTeams = 0
    $NotePadPlusPlus = 0
    $OpenJDK = 0
    $OracleJava8 = 0
    $TreeSizeFree = 0
    $VLCPlayer = 0
    $VMWareTools = 0
    $WinSCP = 0

![Parameter -list](/img/script-list.png)

For example, to automate the process via Scheduled Task or to integrate this into [BIS-F](https://eucweb.com/download-bis-f) (Thx Matthias Schlimm for your work).

### -download

Only download the selected software packages in list Mode (-list).

### -install

Only install the selected software packages in list Mode (-list).


## Example

.\Evergreen.ps1 -list -download

Downlod the selected Software out of the list.


.\Evergreen.ps1 -list -install

Install the selected Software out of the list.


.\Evergreen.ps1 -list

Download and install the selected Software out of the list.


.\Evergreen.ps1

Start the GUI to select the mode (Install and/or Download) and the Software.

## Notes

### Evergreen PowerShell Module
If Download is selected, the module is checked each time the script is run and reinstalled if a new version is available.

### 7-ZIP
Line 605 defines which package is downloaded (You can change the architecture).

For 7-ZIP this is the x64 exe file.

### Adobe Pro DC
Line 638 defines which package is downloaded (You can change the version).

For Adobe Pro DC this is the update package (msp file).

Only update @ the moment, no installer!

After update stop & disable Adobe service & scheduled task.

### Adobe Reader DC
Line 671 defines which package is downloaded (You can change the language).

For Adobe Reader DC this is the english exe file.

After installation stop & disable Adobe service & scheduled task.

### BIS-F
Line 703 defines which package is downloaded.

For BIS-F this is the msi file.

After installation, customization of the scripts regarding Task Offload, RSS to enable and DEP to disable.

### Citrix Hypervisor Tools
Line 736 defines which package is downloaded (You can change the architecture and the version).

For Citrix Hypervisor Tools this is the x64 msi file (LTSR Path).

For Windows 7, Windows Server 2008 SP2, Windows Server 2008 R2 SP1 you can switch to the version 7.2.0.1555.

### Citrix WorkspaceApp Current Release
Line 770 defines which package is downloaded (You can change the release or use the Citrix_Workspace_LTSR switch).

For Citrix Workspace App this is the exe file (CR Path).

Before the installation of the new receiver, the old one is uninstalled via Receiver CleanUp Tool.

The installation is executed with the following parameters (from line 1986):
        /forceinstall
        /silent
        /EnableCEIP=false
        /FORCE_LAA=1
        /AutoUpdateCheck=disabled
        /EnableCEIP=false
        /ALLOWADDSTORE=S
        /ALLOWSAVEPWD=S
        /includeSSON
        /ENABLE_SSON=Yes
        
After the installation various registry keys are still set (from line 2006).

As always, after installing the new Workspace Agent, the system should be rebooted.

### Citrix WorkspaceApp Long Term Service Release
Line 812 defines which package is downloaded (You can change the release or use the Citrix_Workspace_CR switch).

For Citrix Workspace App this is the exe file (LTSR Path).

Before the installation of the new receiver, the old one is uninstalled via Receiver CleanUp Tool.

The installation is executed with the following parameters (from line 2046):
        /forceinstall
        /silent
        /EnableCEIP=false
        /FORCE_LAA=1
        /AutoUpdateCheck=disabled
        /EnableCEIP=false
        /ALLOWADDSTORE=S
        /ALLOWSAVEPWD=S
        /includeSSON
        /ENABLE_SSON=Yes
        
After the installation various registry keys are still set (from line 2066).

As always, after installing the new Workspace Agent, the system should be rebooted.

### Filezilla
Line 854 defines which package is downloaded.

For Filezilla this is the exe file.

Filezilla is installed with the parameter /user=all for all users.

### Firefox
Line 887 defines which package is downloaded (You can change the architecture, language and the channel(to ESR)).

For Firefox this is the english x64 msi file (Latest Firefox Version).

Firefox is installed with the parameter that don't create icons or the maintenance service.

### Foxit_Reader
Line 920 defines which package is downloaded (You can change the language).

For Foxit Reader this is the english exe file.

Unfortunately, a silent install is not possible at the moment.

### FSLogix
Line 953 defines which package is downloaded.

For FSLogix this is the zip package.

With FSLogix installation, the old installation, if present, is uninstalled first and a restart is requested. 

Then the script must be started again, so that the new version is installed cleanly.

Not only the FSLogix Agent is installed, but also the FSLogix AppRule Editor.

### GoogleChrome
Line 1024 defines which package is downloaded (You can change the architecture).

For Google Chrome this is the x64 msi file.

After installation stop & disable Chrome services & scheduled tasks.

### Greenshot
Line 992 defines which package is downloaded.

For Greenshot this is the exe file.

### KeePass
Line 1054 defines which package is downloaded.

For KeePass this is the msi file.

### mRemoteNG
Line 1087 defines which package is downloaded.

For mRemoteNG this is the msi file.

### Microsoft 365 Apps
Line 1120 defines which package is downloaded (You can change the channel).

For Microsoft 365 Apps this is the exe setup file for the Semi-Annual Channel.

During the download not only the setup.exe is downloaded, but also the following xml files are created, if they are not already present in the folder:

remove.xml (from line 1132)

install.xml (from line 1157)

Afterwards the install.xml is used to download the required install files.

Before installing the new Microsoft 365 Apps version, the previous installation is removed (remove.xml).

After that the reinstall of the software starts (install.xml).

An install.xml with the special features of the own installation can be stored and used in advance (e.g. Languages, App Extension or Inclusion (Visio & Project)).

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
Line 1226 defines which package is downloaded (You can change the architecture).

For Microsoft Edge this is the x64 msi file.

Microsoft Edge is installed with the parameter that don't create icons.

After installation disable Microsoft Edge scheduled tasks and set Citrix API Hooks in the registry.

### Microsoft Office 2019
Line 1257 defines which package is downloaded (You can change the channel).

For Microsoft Office 2019 this is the exe setup file for Office 2019 Enterprise.

During the download not only the setup.exe is downloaded, but also the following xml files are created, if they are not already present in the folder:

remove.xml (from line 1269)

install.xml (from line 1294)

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
Line 1364 defines which package is downloaded (You can change the update ring).

For Microsoft OneDrive this is the Production Ring exe file.

Microsoft OneDrive is installed with the Machine Based Install parameter.

### Microsoft Teams
Line 1397 defines which package is downloaded (You can change the architecture and update ring).

For Microsoft Teams this is the x64 msi file (General Ring).

Microsoft Teams is installed with the Machine Based Install parameters.

After installation disable Microsoft Teams autostart registry key can be enabled (from line 2820).

### NotePad++
Line 1430 defines which package is downloaded (You can change the architecture).

For Notepad++ this is the x64 exe file.

### OpenJDK
Line 1463 defines which package is downloaded (You can change the architecture).

For OpenJDK this is the x64 msi file.

### Oracle Java 8
Line 1496 defines which package is downloaded (You can change the architecture).

For Oracle Java 8 this is the x64 msi file.

### TreeSize Free
Line 1529 defines which package is downloaded.

For TreeSize Free this is the exe file.

### VLC Player
Line 1562 defines which package is downloaded (You can change the architecture).

For VLC Player this is the x64 msi file.

### VMWare Tools
Line 1595 defines which package is downloaded (You can change the architecture).

For VMWare Tools this is the x64 exe file.

With VMWare Tools installation, the old installation, if present, is uninstalled first and a restart is requested. 

Then the script must be started again, so that the new version is installed cleanly.
    
### WinSCP
Line 1628 defines which package is downloaded.

For WinSCP this is the exe file.
