#requires -version 3
<#
.SYNOPSIS
Download and Install several Software with the Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. 
.DESCRIPTION
To update or download a software package just switch from 0 to 1 in the section "Select software" (With parameter -list) or select your Software out of the GUI.
A new folder for every single package will be created, together with a version file, a download date file and a log file. If a new version is available
the script checks the version number and will update the package.
.NOTES
  Version:          0.5
  Author:           Manuel Winkel <www.deyda.net>
  Creation Date:    2021-01-29
  Purpose/Change:
  2021-01-29        Initial Version
  2021-01-30        Error solved: No installation without parameters / Add WinSCP Install
  2021-01-31        Error solved: Installation Workspace App -> Wrong Variable / Error solved: Detection acute version 7-Zip -> Limitation of the results
  2021-02-01        Add Gui Mode as Standard
  2021-02-02        Add Install OpenJDK / Add Install VMWare Tools / Add Install Oracle Java 8 / Add Install Adobe Reader DC
<#


.PARAMETER download

Only download the software packages.

.PARAMETER install

Only install the software packages.

.PARAMETER list

Don't start the GUI to select the Software Packages and use the hardcoded list in the script.

.EXAMPLE

& '.\Evergreen.ps1 -download

Downlod the selected Software.

.EXAMPLE

& '.\Evergreen.ps1

Download and install the selected Software.

.EXAMPLE

& '.\Evergreen.ps1 -list

Download and install the selected Software out of the script.
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

# Do you run the script as admin?
# ========================================================================================================================================
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

# Script Version
# ========================================================================================================================================
$eVersion = "0.5"
Write-Verbose "Evergreen Download and Install Script by Manuel Winkel (www.deyda.net) - Version $eVersion" -Verbose
Write-Output ""

if ($myWindowsPrincipal.IsInRole($adminRole)) {
    # OK, runs as admin
    Write-Verbose "OK, script is running with Admin rights" -Verbose
    Write-Output ""
}
else {
    # Script doesn't run as admin, stop!
    Write-Verbose "Error! Script is NOT running with Admin rights!" -Verbose
    BREAK
}

# FUNCTION GUI
# ========================================================================================================================================
function gui_mode{
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Set Variable
    $Script:install = $true
    $Script:download = $true

    # Set the size of your form
    $Form = New-Object system.Windows.Forms.Form
    $Form.ClientSize = New-Object System.Drawing.Point(300,900)
    $Form.text = "Evergreen - Update your Software - Version $eVersion"
    $Form.TopMost = $true
    $Form.AutoSize = $true

    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font

    # Download / Install Headline
    $Headline1 = New-Object system.Windows.Forms.Label
    $Headline1.text = "Select Mode"
    $Headline1.AutoSize = $true
    $Headline1.width = 25
    $Headline1.height = 10
    $Headline1.location = New-Object System.Drawing.Point(11,4)
    $form.Controls.Add($Headline1)

    # Download Checkbox
    $DownloadBox = New-Object system.Windows.Forms.CheckBox
    $DownloadBox.text = "Download"
    $DownloadBox.AutoSize = $false
    $DownloadBox.width = 95
    $DownloadBox.height = 20
    $DownloadBox.location = New-Object System.Drawing.Point(11,35)
    $form.Controls.Add($DownloadBox)

    # Install Checkbox
    $InstallBox = New-Object system.Windows.Forms.CheckBox
    $InstallBox.text = "Install"
    $InstallBox.AutoSize = $false
    $InstallBox.width = 95
    $InstallBox.height = 20
    $InstallBox.location = New-Object System.Drawing.Point(108,35)
    $form.Controls.Add($InstallBox)

    # Software Headline
    $Headline2 = New-Object system.Windows.Forms.Label
    $Headline2.text = "Select Software"
    $Headline2.AutoSize = $true
    $Headline2.width = 25
    $Headline2.height = 10
    $Headline2.location = New-Object System.Drawing.Point(11,70)
    $form.Controls.Add($Headline2)

    # 7Zip Checkbox
    $7ZipBox = New-Object system.Windows.Forms.CheckBox
    $7ZipBox.text = "7 Zip"
    $7ZipBox.width = 95
    $7ZipBox.height = 20
    $7ZipBox.autosize = $true
    $7ZipBox.location = New-Object System.Drawing.Point(11,95)
    $form.Controls.Add($7ZipBox)

    # AdobeProDC Checkbox
    $AdobeProDCBox = New-Object system.Windows.Forms.CheckBox
    $AdobeProDCBox.text = "Adobe Pro DC #Only Update @ the moment"
    $AdobeProDCBox.width = 95
    $AdobeProDCBox.height = 20
    $AdobeProDCBox.autosize = $true
    $AdobeProDCBox.location = New-Object System.Drawing.Point(11,120)
    $form.Controls.Add($AdobeProDCBox)

    # AdobeReaderDC Checkbox
    $AdobeReaderDCBox = New-Object system.Windows.Forms.CheckBox
    $AdobeReaderDCBox.text = "Adobe Reader DC"
    $AdobeReaderDCBox.width = 95
    $AdobeReaderDCBox.height = 20
    $AdobeReaderDCBox.autosize = $true
    $AdobeReaderDCBox.location = New-Object System.Drawing.Point(11,145)
    $form.Controls.Add($AdobeReaderDCBox)

    # BISF Checkbox
    $BISFBox = New-Object system.Windows.Forms.CheckBox
    $BISFBox.text = "BIS-F"
    $BISFBox.width = 95
    $BISFBox.height = 20
    $BISFBox.autosize = $true
    $BISFBox.location = New-Object System.Drawing.Point(11,170)
    $form.Controls.Add($BISFBox)

    # FSLogix Checkbox
    $FSLogixBox = New-Object system.Windows.Forms.CheckBox
    $FSLogixBox.text = "FSLogix"
    $FSLogixBox.width = 95
    $FSLogixBox.height = 20
    $FSLogixBox.autosize = $true
    $FSLogixBox.location = New-Object System.Drawing.Point(11,195)
    $form.Controls.Add($FSLogixBox)

    # GoogleChrome Checkbox
    $GoogleChromeBox = New-Object system.Windows.Forms.CheckBox
    $GoogleChromeBox.text = "Google Chrome"
    $GoogleChromeBox.width = 95
    $GoogleChromeBox.height = 20
    $GoogleChromeBox.autosize = $true
    $GoogleChromeBox.location = New-Object System.Drawing.Point(11,220)
    $form.Controls.Add($GoogleChromeBox)

    # KeePass Checkbox
    $KeePassBox = New-Object system.Windows.Forms.CheckBox
    $KeePassBox.text = "KeePass"
    $KeePassBox.width = 95
    $KeePassBox.height = 20
    $KeePassBox.autosize = $true
    $KeePassBox.location = New-Object System.Drawing.Point(11,245)
    $form.Controls.Add($KeePassBox)

    # mRemoteNG Checkbox
    $mRemoteNGBox = New-Object system.Windows.Forms.CheckBox
    $mRemoteNGBox.text = "mRemoteNG"
    $mRemoteNGBox.width = 95
    $mRemoteNGBox.height = 20
    $mRemoteNGBox.autosize = $true
    $mRemoteNGBox.location = New-Object System.Drawing.Point(11,270)
    $form.Controls.Add($mRemoteNGBox)

    # MS365Apps Checkbox
    $MS365AppsBox = New-Object system.Windows.Forms.CheckBox
    $MS365AppsBox.text = "Microsoft 365 Apps # Office Deployment Toolkit for installing Office 365 / Only Download @ the moment"
    $MS365AppsBox.width = 95
    $MS365AppsBox.height = 20
    $MS365AppsBox.autosize = $true
    $MS365AppsBox.location = New-Object System.Drawing.Point(11,295)
    $form.Controls.Add($MS365AppsBox)

    # MSEdge Checkbox
    $MSEdgeBox = New-Object system.Windows.Forms.CheckBox
    $MSEdgeBox.text = "Microsoft Edge"
    $MSEdgeBox.width = 95
    $MSEdgeBox.height = 20
    $MSEdgeBox.autosize = $true
    $MSEdgeBox.location = New-Object System.Drawing.Point(11,320)
    $form.Controls.Add($MSEdgeBox)

    # MSOffice2019 Checkbox
    $MSOffice2019Box = New-Object system.Windows.Forms.CheckBox
    $MSOffice2019Box.text = "Microsoft Office 2019 # OfficeDeployment Toolkit for installing Office 2019 / Only Download @ the moment"
    $MSOffice2019Box.width = 95
    $MSOffice2019Box.height = 20
    $MSOffice2019Box.autosize = $true
    $MSOffice2019Box.location = New-Object System.Drawing.Point(11,345)
    $form.Controls.Add($MSOffice2019Box)

    # MSOneDrive Checkbox
    $MSOneDriveBox = New-Object system.Windows.Forms.CheckBox
    $MSOneDriveBox.text = "Microsoft OneDrive (Machine-Based Install)"
    $MSOneDriveBox.width = 95
    $MSOneDriveBox.height = 20
    $MSOneDriveBox.autosize = $true
    $MSOneDriveBox.location = New-Object System.Drawing.Point(11,370)
    $form.Controls.Add($MSOneDriveBox)

    # MSTeams Checkbox
    $MSTeamsBox = New-Object system.Windows.Forms.CheckBox
    $MSTeamsBox.text = "Microsoft Teams (Machine-Based Install)"
    $MSTeamsBox.width = 95
    $MSTeamsBox.height = 20
    $MSTeamsBox.autosize = $true
    $MSTeamsBox.location = New-Object System.Drawing.Point(11,395)
    $form.Controls.Add($MSTeamsBox)

    # NotePadPlusPlus Checkbox
    $NotePadPlusPlusBox = New-Object system.Windows.Forms.CheckBox
    $NotePadPlusPlusBox.text = "NotePad++"
    $NotePadPlusPlusBox.width = 95
    $NotePadPlusPlusBox.height = 20
    $NotePadPlusPlusBox.autosize = $true
    $NotePadPlusPlusBox.location = New-Object System.Drawing.Point(11,420)
    $form.Controls.Add($NotePadPlusPlusBox)

    # OpenJDK Checkbox
    $OpenJDKBox = New-Object system.Windows.Forms.CheckBox
    $OpenJDKBox.text = "Open JDK"
    $OpenJDKBox.width = 95
    $OpenJDKBox.height = 20
    $OpenJDKBox.autosize = $true
    $OpenJDKBox.location = New-Object System.Drawing.Point(11,445)
    $form.Controls.Add($OpenJDKBox)

    # OracleJava8 Checkbox
    $OracleJava8Box = New-Object system.Windows.Forms.CheckBox
    $OracleJava8Box.text = "Oracle Java 8"
    $OracleJava8Box.width = 95
    $OracleJava8Box.height = 20
    $OracleJava8Box.autosize = $true
    $OracleJava8Box.location = New-Object System.Drawing.Point(11,470)
    $form.Controls.Add($OracleJava8Box)

    # TreeSizeFree Checkbox
    $TreeSizeFreeBox = New-Object system.Windows.Forms.CheckBox
    $TreeSizeFreeBox.text = "TreeSize Free"
    $TreeSizeFreeBox.width = 95
    $TreeSizeFreeBox.height = 20
    $TreeSizeFreeBox.autosize = $true
    $TreeSizeFreeBox.location = New-Object System.Drawing.Point(11,495)
    $form.Controls.Add($TreeSizeFreeBox)

    # VLCPlayer Checkbox
    $VLCPlayerBox = New-Object system.Windows.Forms.CheckBox
    $VLCPlayerBox.text = "VLC Player"
    $VLCPlayerBox.width = 95
    $VLCPlayerBox.height = 20
    $VLCPlayerBox.autosize = $true
    $VLCPlayerBox.location = New-Object System.Drawing.Point(11,520)
    $form.Controls.Add($VLCPlayerBox)

    # VMWareTools Checkbox
    $VMWareToolsBox = New-Object system.Windows.Forms.CheckBox
    $VMWareToolsBox.text = "VMWare Tools"
    $VMWareToolsBox.width = 95
    $VMWareToolsBox.height = 20
    $VMWareToolsBox.autosize = $true
    $VMWareToolsBox.location = New-Object System.Drawing.Point(11,545)
    $form.Controls.Add($VMWareToolsBox)

    # WinSCP Checkbox
    $WinSCPBox = New-Object system.Windows.Forms.CheckBox
    $WinSCPBox.text = "WinSCP"
    $WinSCPBox.width = 95
    $WinSCPBox.height = 20
    $WinSCPBox.autosize = $true
    $WinSCPBox.location = New-Object System.Drawing.Point(11,570)
    $form.Controls.Add($WinSCPBox)

    # WorkspaceApp_Current_Release Checkbox
    $WorkspaceApp_Current_ReleaseBox = New-Object system.Windows.Forms.CheckBox
    $WorkspaceApp_Current_ReleaseBox.text = "Citrix WorkspaceApp Current Release"
    $WorkspaceApp_Current_ReleaseBox.width = 95
    $WorkspaceApp_Current_ReleaseBox.height = 20
    $WorkspaceApp_Current_ReleaseBox.autosize = $true
    $WorkspaceApp_Current_ReleaseBox.location = New-Object System.Drawing.Point(11,595)
    $form.Controls.Add($WorkspaceApp_Current_ReleaseBox) 

    # WorkspaceApp_LTSR_Release Checkbox
    $WorkspaceApp_LTSR_ReleaseBox = New-Object system.Windows.Forms.CheckBox
    $WorkspaceApp_LTSR_ReleaseBox.text = "Citrix WorkspaceApp LTSR Release"
    $WorkspaceApp_LTSR_ReleaseBox.width = 95
    $WorkspaceApp_LTSR_ReleaseBox.height = 20
    $WorkspaceApp_LTSR_ReleaseBox.autosize = $true
    $WorkspaceApp_LTSR_ReleaseBox.location = New-Object System.Drawing.Point(11,620)
    $form.Controls.Add($WorkspaceApp_LTSR_ReleaseBox)

    # SelectAll Checkbox
    $SelectAllBox = New-Object system.Windows.Forms.CheckBox
    $SelectAllBox.text = "Select All"
    $SelectAllBox.width = 95
    $SelectAllBox.height = 20
    $SelectAllBox.autosize = $true
    $SelectAllBox.location = New-Object System.Drawing.Point(11,775)
    $SelectAllBox.Add_CheckStateChanged({
        $7ZipBox.Checked = $SelectAllBox.Checked
        $AdobeProDCBox.Checked = $SelectAllBox.Checked
        $AdobeReaderDCBox.Checked = $SelectAllBox.Checked
        $BISFBox.Checked = $SelectAllBox.Checked
        $FSLogixBox.Checked = $SelectAllBox.Checked
        $GoogleChromeBox.Checked = $SelectAllBox.Checked
        $KeePassBox.Checked = $SelectAllBox.Checked
        $mRemoteNGBox.Checked = $SelectAllBox.Checked
        $MS365AppsBox.Checked = $SelectAllBox.Checked
        $MSEdgeBox.Checked = $SelectAllBox.Checked
        $MSOffice2019Box.Checked = $SelectAllBox.Checked
        $MSOneDriveBox.Checked = $SelectAllBox.Checked
        $MSTeamsBox.Checked = $SelectAllBox.Checked
        $NotePadPlusPlusBox.Checked = $SelectAllBox.Checked
        $OpenJDKBox.Checked = $SelectAllBox.Checked
        $OracleJava8Box.Checked = $SelectAllBox.Checked
        $TreeSizeFreeBox.Checked = $SelectAllBox.Checked
        $VLCPlayerBox.Checked = $SelectAllBox.Checked
        $VMWareToolsBox.Checked = $SelectAllBox.Checked
        $WinSCPBox.Checked = $SelectAllBox.Checked
        $WorkspaceApp_Current_ReleaseBox.Checked = $SelectAllBox.Checked
        $WorkspaceApp_LTSR_ReleaseBox.Checked = $SelectAllBox.Checked
    })
    $form.Controls.Add($SelectAllBox)

    # OK Button
    $OKButton = New-Object system.Windows.Forms.Button
    $OKButton.text = "OK"
    $OKButton.width = 60
    $OKButton.height = 30
    $OKButton.location = New-Object System.Drawing.Point(70,800)
    $OKButton.Add_Click({
        if ($DownloadBox.checked -eq $true) {$Script:install = $false}
        else {$Script:install = $true}
        if ($InstallBox.checked -eq $true) {$Script:download = $false}
        else {$Script:download = $true}
        if ($7ZipBox.checked -eq $true) {$Script:7ZIP = 1}
        else {$Script:7ZIP = 0}
        if ($AdobeProDCBox.checked -eq $true) {$Script:AdobeProDC = 1}
        else {$Script:AdobeProDC = 0}
        if ($AdobeReaderDCBox.checked -eq $true) {$Script:AdobeReaderDC = 1}
        else {$Script:AdobeReaderDC = 0}
        if ($BISFBox.checked -eq $true) {$Script:BISF = 1}
        else {$Script:BISF = 0}
        if ($FSLogixBox.checked -eq $true) {$Script:FSLogix = 1}
        else {$Script:FSLogix = 0}
        if ($GoogleChromeBox.checked -eq $true) {$Script:GoogleChrome = 1}
        else {$Script:GoogleChrome = 0}
        if ($KeePassBox.checked -eq $true) {$Script:KeePass = 1}
        else {$Script:KeePass = 0}
        if ($mRemoteNGBox.checked -eq $true) {$Script:mRemoteNG = 1}
        else {$Script:mRemoteNG = 0}
        if ($MS365AppsBox.checked -eq $true) {$Script:MS365Apps = 1}
        else {$Script:MS365Apps = 0}
        if ($MSEdgeBox.checked -eq $true) {$Script:MSEdge = 1}
        else {$Script:MSEdge = 0}
        if ($MSOffice2019Box.checked -eq $true) {$Script:MSOffice2019 = 1}
        else {$Script:MSOffice2019 = 0}
        if ($MSOneDriveBox.checked -eq $true) {$Script:MSOneDrive = 1}
        else {$Script:MSOneDrive = 0}
        if ($MSTeamsBox.checked -eq $true) {$Script:MSTeams = 1}
        else {$Script:MSTeams = 0}
        if ($NotePadPlusPlusBox.checked -eq $true) {$Script:NotePadPlusPlus = 1}
        else {$Script:NotePadPlusPlus = 0}
        if ($OpenJDKBox.checked -eq $true) {$Script:OpenJDK = 1}
        else {$Script:OpenJDK = 0}
        if ($OracleJava8Box.checked -eq $true) {$Script:OracleJava8 = 1}
        else {$Script:OracleJava8 = 0}
        if ($TreeSizeFreeBox.checked -eq $true) {$Script:TreeSizeFree = 1}
        else {$Script:TreeSizeFree = 0}
        if ($VLCPlayerBox.checked -eq $true) {$Script:VLCPlayer = 1}
        else {$Script:VLCPlayer = 0}
        if ($VMWareToolsBox.checked -eq $true) {$Script:VMWareTools = 1}
        else {$Script:VMWareTools = 0}
        if ($WinSCPBox.checked -eq $true) {$Script:WinSCP = 1}
        else {$Script:WinSCP = 0}
        if ($WorkspaceApp_Current_ReleaseBox.checked -eq $true) {$Script:WorkspaceApp_Current_Release = 1}
        else {$Script:WorkspaceApp_Current_Release = 0}
        if ($WorkspaceApp_LTSR_ReleaseBox.checked -eq $true) {$Script:WorkspaceApp_LTSR_Release = 1}
        else {$Script:WorkspaceApp_LTSR_Release = 0}
        
        Write-Verbose "GUI MODE" -Verbose
        $Form.Close()
    })
    $form.Controls.Add($OKButton)

    # Cancel Button
    $CancelButton = New-Object system.Windows.Forms.Button
    $CancelButton.text = "Cancel"
    $CancelButton.width = 60
    $CancelButton.height = 30
    $CancelButton.location = New-Object System.Drawing.Point(170,800)
    $CancelButton.Add_Click({
        $Script:install = $true
        $Script:download = $true
        Write-Verbose "GUI MODE Canceled - Nothing happens" -Verbose
        $Form.Close()
    })
    $form.Controls.Add($CancelButton)

    # Activate the form
    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
}
# ========================================================================================================================================

Write-Verbose "Setting Variables" -Verbose
Write-Output ""

# Define and Reset Variables
$Date = $Date = Get-Date -UFormat "%m.%d.%Y"
$Script:install = $install
$Script:download = $download

if ($list -eq $True) {
    # Select software
    $7ZIP = 0
    $AdobeProDC = 0 #Only Update @ the moment
    $AdobeReaderDC = 0
    $BISF = 0
    $FSLogix = 0
    $GoogleChrome = 0
    $KeePass = 0
    $mRemoteNG = 0
    $MS365Apps = 0 # Office Deployment Toolkit for installing Office 365 / Only Download @ the moment
    $MSEdge = 0
    $MSOffice2019 = 0 # Office Deployment Toolkit for installing Office 2019 / Only Download @ the moment
    $MSOneDrive = 0
    $MSTeams = 0
    $NotePadPlusPlus = 0
    $OpenJDK = 0
    $OracleJava8 = 0
    $TreeSizeFree = 0
    $VLCPlayer = 0
    $VMWareTools = 0
    $WinSCP = 0
    $WorkspaceApp_Current_Release = 1
    $WorkspaceApp_LTSR_Release = 0
}
else {
    Clear-Variable -name 7ZIP,AdobeProDC,AdobeReaderDC,BISF,FSLogix,GoogleChrome,KeePass,mRemoteNG,MS365Apps,MSEdge,MSOffice2019,MSTeams,NotePadPlusPlus,MSOneDrive,OpenJDK,OracleJava8,TreeSizeFree,VLCPlayer,VMWareTools,WinSCP,WorkspaceApp_Current_Release,WorkspaceApp_LTSR_Release -ErrorAction SilentlyContinue
    gui_mode
}


# Disable progress bar while downloading
$ProgressPreference = 'SilentlyContinue'

if ($install -eq $False) {
    # Install/Update Evergreen module
    Write-Output ""
    Write-Verbose "Installing/updating Evergreen module... please wait" -Verbose
    Write-Output ""
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies}
    if (!(Get-Module -ListAvailable -Name Evergreen)) {Install-Module Evergreen -Force | Import-Module Evergreen}
    Update-Module Evergreen -force

    Write-Output "Starting downloads..."
    Write-Output ""

    # Download 7-ZIP
    if ($7ZIP -eq 1) {
        $Product = "7-Zip"
        $PackageName = "7-Zip_x64"
        $7ZipD = Get-7zip | Where-Object { $_.Architecture -eq "x64" -and $_.URI -like "*exe*" }
        $Version = $7ZipD.Version
        $URL = $7ZipD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        $PackageName = "Adobe_Pro_DC_Update"
        $AdobeProD = Get-AdobeAcrobatProDC | Where-Object { $_.Type -eq "Updater" }
        $Version = $AdobeProD.Version
        $URL = $AdobeProD.uri
        $InstallerType = "msp"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        $PackageName = "Adobe_Reader_DC"
        $AdobeReaderD = Get-AdobeAcrobatReaderDC | Where-Object {$_.Type -eq "Installer" -and $_.Language -eq "English"}
        $Version = $AdobeReaderD.Version
        $URL = $AdobeReaderD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($BISF -eq 1) {
        $Product = "BIS-F"
        $PackageName = "setup-BIS-F"
        $BISFD = Get-BISF | Where-Object { $_.URI -like "*msi*" }
        $Version = $BISFD.Version
        $URL = $BISFD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($FSLogix -eq 1) {
        $Product = "FSLogix"
        $PackageName = "FSLogixAppsSetup"
        $FSLogixD = Get-MicrosoftFSLogixApps
        $Version = $FSLogixD.Version
        $URL = $FSLogixD.uri
        $InstallerType = "zip"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Install\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        $ChromeURL = Get-GoogleChrome | Where-Object { $_.Architecture -eq "x64" } | Select-Object -ExpandProperty URI
        $Version = (Get-GoogleChrome | Where-Object { $_.Architecture -eq "x64" }).Version
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($KeePass -eq 1) {
        $Product = "KeePass"
        $PackageName = "KeePass"
        $KeePassD = Get-KeePass | Where-Object { $_.URI -like "*msi*" }
        $Version = $KeePassD.Version
        $URL = $KeePassD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($mRemoteNG -eq 1) {
        $Product = "mRemoteNG"
        $PackageName = "mRemoteNG"
        $mRemoteNGD = Get-mRemoteNG | Where-Object { $_.URI -like "*msi*" }
        $Version = $mRemoteNGD.Version
        $URL = $mRemoteNGD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($MS365Apps -eq 1) {
        $Product = "MS 365 Apps (Semi Annual Channel)"
        $PackageName = "setup"
        $MS365AppsD = Get-Microsoft365Apps | Where-Object {$_.Channel -eq "Semi-Annual Channel"}
        $Version = $MS365AppsD.Version
        $URL = $MS365AppsD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
    if (!($CurrentVersion -eq $Version)) {
        if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
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
    if ($MSEdge -eq 1) {
        $Product = "MS Edge"
        $EdgeURL = Get-MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "x64" }
        $EdgeURL = $EdgeURL | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = (Get-MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "x64" }).Version
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue 
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($MSOffice2019 -eq 1) {
        $Product = "MS Office 2019"
        $PackageName = "setup"
        $MSOffice2019D = Get-Microsoft365Apps | Where-Object {$_.Channel -eq "Office 2019 Enterprise"}
        $Version = $MSOffice2019D.Version
        $URL = $MSOffice2019D.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
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

    # Download MSOneDrive
    if ($MSOneDrive -eq 1) {
        $Product = "MS OneDrive"
        $PackageName = "OneDriveSetup"
        $MSOneDriveD = Get-MicrosoftOneDrive | Where-Object { $_.Ring -eq "Production" -and $_.Type -eq "Exe" } | Sort-Object -Property Version -Descending | Select-Object -Last 1
        $Version = $MSOneDriveD.Version
        $URL = $MSOneDriveD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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

    # Download MS Teams
    if ($MSTeams -eq 1) {
        $Product = "MS Teams"
        $PackageName = "Teams_windows_x64"
        $TeamsD = Get-MicrosoftTeams | Where-Object { $_.Architecture -eq "x64" }
        $Version = $TeamsD.Version
        $URL = $TeamsD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($NotePadPlusPlus -eq 1) {
        $Product = "NotePadPlusPlus"
        $PackageName = "NotePadPlusPlus_x64"
        $NotepadD = Get-NotepadPlusPlus | Where-Object { $_.Architecture -eq "x64" -and $_.URI -match ".exe" }
        $Version = $NotepadD.Version
        $URL = $NotepadD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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

    # Download openJDK
    if ($OpenJDK -eq 1) {
        $Product = "open JDK"
        $PackageName = "OpenJDK"
        $OpenJDKD = Get-OpenJDK | Where-Object { $_.Architecture -eq "x64" -and $_.URI -like "*msi*" } | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $OpenJDKD.Version
        $URL = $OpenJDKD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        $PackageName = "Oracle Java 8"
        $OracleJava8D = Get-OracleJava8 | Where-Object { $_.Architecture -eq "x64" }
        $Version = $OracleJava8D.Version
        $URL = $OracleJava8D.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($TreeSizeFree -eq 1) {
        $Product = "TreeSizeFree"
        $PackageName = "TreeSizeFree"
        $TreeSizeFreeD = Get-JamTreeSizeFree
        $Version = $TreeSizeFreeD.Version
        $URL = $TreeSizeFreeD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        $PackageName = "VLC-Player"
        $VLCD = Get-VideoLanVlcPlayer | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" -and $_.Type -eq "MSI" }
        $Version = $VLCD.Version
        $URL = $VLCD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"
        $PackageName = "VMWareTools"
        $VMWareToolsD = Get-VMwareTools | Where-Object { $_.Architecture -eq "x64" }
        $Version = $VMWareToolsD.Version
        $URL = $VMWareToolsD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($WinSCP -eq 1) {
        $Product = "WinSCP"
        $PackageName = "WinSCP"
        $WinSCPD = Get-WinSCP | Where-Object {$_.URI -like "*Setup*"}
        $Version = $WinSCPD.Version
        $URL = $WinSCPD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($WorkspaceApp_Current_Release -eq 1) {
        $Product = "WorkspaceApp"
        $PackageName = "CitrixWorkspaceApp"
        $WSACD = Get-CitrixWorkspaceApp | Where-Object { $_.Title -like "*Workspace*" -and "*Current*" -and $_.Platform -eq "Windows" -and $_.Title -like "*Current*" }
        $Version = $WSACD.Version
        $URL = $WSACD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\Current\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product CR" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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
    if ($WorkspaceApp_LTSR_Release -eq 1) {
        $Product = "WorkspaceApp"
        $PackageName = "CitrixWorkspaceApp"
        $WSALD = Get-CitrixWorkspaceApp | Where-Object { $_.Title -like "*Workspace*" -and "*LTSR*" -and $_.Platform -eq "Windows" -and $_.Title -like "*LTSR*" }
        $Version = $WSALD.Version
        $URL = $WSALD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Windows\LTSR\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product LTSR" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
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

if ($download -eq $False) {

    # FUNCTION Logging
    #========================================================================================================================================
    Function DS_WriteLog {
        
        [CmdletBinding()]
        Param (
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
    # Global variables
    #$StartDir = $PSScriptRoot # the directory path of the script currently being executed
    $LogDir = "$PSScriptRoot\_Install Logs"
    $LogFileName = ("$ENV:COMPUTERNAME.log")
    $LogFile = Join-path $LogDir $LogFileName

    # Create the log directory if it does not exist
    if (!(Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType directory | Out-Null }

    # Create new log file (overwrite existing one)
    New-Item $LogFile -ItemType "file" -force | Out-Null
    DS_WriteLog "I" "START SCRIPT - " $LogFile
    DS_WriteLog "-" "" $LogFile
    #========================================================================================================================================
    
    # define Error handling
    # note: do not change these values
    $global:ErrorActionPreference = "Stop"
    if ($verbose){ $global:VerbosePreference = "Continue" }


    # Install 7-ZIP
    if ($7ZIP -eq 1) {
        $Product = "7-Zip"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $SevenZip = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Select-Object -First 1
        if ($SevenZip -ne $Version) {
            # 7-Zip
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\7-Zip_x64.exe" –ArgumentList /S –NoNewWindow -Wait
                $p = Get-Process 7-Zip_x64
                if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
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
    if ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion
        if ($Adobe -ne $Version) {
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
    if ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*"}).DisplayVersion
        if ($Adobe -ne $Version) {
            # Adobe Reader DC
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            $Options = @(
                "/sAll"
                "/rs"
            )
            try	{
                Start-Process "$PSScriptRoot\$Product\Adobe_Reader_DC.exe" –ArgumentList $Options –NoNewWindow
                $p = Get-Process Adobe_Reader_DC
                if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
                # Update Dienst und Task deaktivieren
                Stop-Service AdobeARMservice
                Set-Service AdobeARMservice -StartupType Disabled
                Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
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

    # Install BIS-F
    if ($BISF -eq 1) {
        $Product = "BIS-F"
        
        # FUNCTION MSI Installation
        #========================================================================================================================================
        function Install-MSiFile {
            [CmdletBinding()]
            Param(
                [parameter(mandatory=$true,ValueFromPipeline=$true,ValueFromPipelinebyPropertyName=$true)]
                [ValidateNotNullorEmpty()]
                [string]$msiFile,
    
                [parameter()]
                [ValidateNotNullorEmpty()]
                [string]$targetDir
            )
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

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
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
        # Customize scripts, it's best practise to enable Task Offload and RSS and to disable DEP
        write-Verbose "Customize scripts" -Verbose
        DS_WriteLog "I" "Customize scripts" $LogFile
        $BISFDir = "C:\Program Files (x86)\Base Image Script Framework (BIS-F)\Framework\SubCall"
        try {
            ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace "DisableTaskOffload' -Value '1'","DisableTaskOffload' -Value '0'") | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
            ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'nx AlwaysOff','nx OptOut') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
            ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'rss=disable','rss=enable') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
        } catch {
            DS_WriteLog "E" "Error when customizing scripts (error: $($Error[0]))" $LogFile
        }
        DS_WriteLog "-" "" $LogFile
        write-Output ""
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
                $p = Get-Process FSLogixAppsSetup
		        if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

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
    IF ($KeePass -eq 1) {
        $Product = "KeePass"

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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path!"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path!"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
                "DONOTCREATEDESKTOPSHORTCUT=TRUE"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================
        
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
                 Disable-ScheduledTask -TaskName MicrosoftEdgeUpdateTaskMachineCore | Out-Null
                 Disable-ScheduledTask -TaskName MicrosoftEdgeUpdateTaskMachineUA | Out-Null
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

    # Install MSOneDrive
    IF ($MSOneDrive -eq 1) {
        $Product = "MS OneDrive"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $MSOneDriveV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion
        IF ($MSOneDriveV -ne $Version) {
            # Installation MSOneDrive
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            $Options = @(
                "/ALLUSERS"
                "/SILENT"
            )
            try	{
                $null = Start-Process "$PSScriptRoot\$Product\OneDriveSetup.exe" –ArgumentList $Options –NoNewWindow -PassThru
                while (Get-Process -Name "OneDriveSetup" -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                Write-Verbose "Installation $Product finished!" -Verbose
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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "ALLUSER=1"
                "ALLUSERS=1"
                "OPTIONS='noAutoStart=true'"
                "/qn"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================
        
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
                $p = Get-Process NotePadPlusPlus_x64
		        if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
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

    # Install OpenJDK
    IF ($OpenJDK -eq 1) {
        $Product = "open JDK"

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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
                "/log $PSScriptRoot\$Product\verbose.log"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $OpenJDK = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion
        IF ($Version) {$Version = $Version -replace ".-"}
        IF ($OpenJDK -ne $Version) {
            # OpenJDK
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\OpenJDK.msi" | Install-MSIFile
                Start-Sleep 25
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

    # Install OracleJava8
    if ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        IF ($Version) {$Version = $Version -replace "^.{2}"}
        IF ($Version) {$Version = $Version -replace "-.*(0)"}
        IF ($Version) {$Version = $Version -replace "._"}
        IF ($Version) {$Version = $Version -replace "\."}
        $OracleJava = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java*"}).DisplayVersion
        IF ($OracleJava) {$OracleJava = $OracleJava -replace "^.{2}"}
        IF ($OracleJava) {$OracleJava = $OracleJava -replace "\."}
        if ($OracleJava -ne $Version) {
            # Oracle Java 8
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            $Options = @(
                "/s"
            )
            try	{
                Start-Process "$PSScriptRoot\$Product\Oracle Java 8.exe" –ArgumentList $Options –NoNewWindow -PassThru
                $p = Get-Process "Oracle Java 8"
                if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
                # Update Dienst und Task deaktivieren
                Stop-Service AdobeARMservice
                Set-Service AdobeARMservice -StartupType Disabled
                Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
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
                $p = Get-Process TreeSizeFree
		        if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
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
            if (!(Test-Path $msiFile)) {
                throw "Path to MSI file ($msiFile) is invalid. Please check name and path"
            }
            $arguments = @(
                "/i"
                "`"$msiFile`""
                "/qn"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

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
                DS_WriteLog "E" "An error occurred installing $Product (error: $($Error[0]))" $LogFile 
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

    # Install VMWareTools
    IF ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $VMWT = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion
        IF ($VMWT) {$VMWT = $VMWT -replace ".{9}$"}
        IF ($VMWT -ne $Version) {
            # VMWareTools Installation
            $Options = @(
                "/s"
                "/v"
                "/qn REBOOT=Y"
            )
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\VMWareTools.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product finished!" -Verbose
                    Write-Output ""
                    Write-Verbose "Server needs to reboot, start script again after reboot" -Verbose
                    Write-Output ""
                }
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

    # Install WinSCP
    IF ($WinSCP -eq 1) {
        $Product = "WinSCP"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $WSCP = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*WinSCP*"}).DisplayVersion
        IF ($WSCP -ne $Version) {
            # WinSCP Installation
            $Options = @(
                "/VERYSILENT"
                "/ALLUSERS"
                "/NORESTART"
                "/NOCLOSEAPPLICATIONS"
            )
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\WinSCP.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
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
                    Write-Verbose "Installation $Product finished!" -Verbose
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
                    Write-Verbose "Installation $Product finished!" -Verbose
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