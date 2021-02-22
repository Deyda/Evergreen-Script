# Evergreen Script by Manuel Winkel / [Deyda.net](https://www.deyda.net) / [@deyda84](https://twitter.com/Deyda84)
Download and Install several Software the lazy way with the Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. 

To update or download a software package just switch from 0 to 1 in the section "Select software" (With PowerShell parameter -list) or select your Software out of the GUI.
A new folder for every single package will be created, together with a version file and a log file. If a new version is available the script checks the version number and will update the package.

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
  2021-02-19        Implementation of new GUI / Choice of architecture option in 7-Zip / Choice of language option in Adobe Reader DC / Choice of architecture option in Citrix Hypervisor Tools / Choice of release option in Citrix Workspace App

## Parameter
### download

Only download the selected software packages.

### install

Only install the selected software packages.

### list

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
    $MS365Apps = 1 # 64Bit / Match OS Language / Semi Annual Channel
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

For example, to automate the process via Scheduled Task or to integrate this into [BIS-F](https://eucweb.com/download-bis-f) (Thx Matthias Schlimm for your work).

## Example

& '.\Evergreen.ps1 -download

Downlod the selected Software.

.EXAMPLE

& '.\Evergreen.ps1

Download and install the selected Software.

.EXAMPLE

& '.\Evergreen.ps1 -list

Download and install the selected Software out of the script.
#>
# Evergreen PowerShell Module
Download, install and update the newest version of several software packages based on the powerful Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein.
https://github.com/aaronparker/Evergreen

## How To
The idea is to select a client or server that periodically checks for updates and if updates are available, downloads them. This can be done every day or once a week by launching the script "Evergreen.ps1 -list" via scheduled task. You decide which software do download by giving it a "0" or "1" in the script.

The "Evergreen.ps1" script must be launched on your clients. If you have a golden master like in Citrix MCS/PVS environments it's sufficient to launch the script only on this machine. This can be done manually or automatic, like you prefer.

If it is run manually, do not use the -list parameter and you will be taken to the GUI to select the software and mode (Download or Install).
