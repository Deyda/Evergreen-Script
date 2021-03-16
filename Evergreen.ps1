#requires -version 3
<#
.SYNOPSIS
Download and Install several Software with the Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein. 
.DESCRIPTION
To update or download a software package just switch from 0 to 1 in the section "Select software" (With parameter -list) or select your Software out of the GUI.
A new folder for every single package will be created, together with a version file, a download date file and a log file. If a new version is available
the script checks the version number and will update the package.
.NOTES
  Version:          0.9
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

# Do you run the script as admin?
# ========================================================================================================================================
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

# Script Version
# ========================================================================================================================================
$eVersion = "0.9"
Write-Verbose "Evergreen Script - Update your Software, the lazy way - Manuel Winkel (www.deyda.net) - Version $eVersion" -Verbose
$host.ui.RawUI.WindowTitle = “Evergreen Script - Update your Software, the lazy way - Manuel Winkel (www.deyda.net) - Version $eVersion”
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

#// MARK: XAML Code
$inputXML = @"
<Window x:Class="GUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:GUI"
        mc:Ignorable="d"
        Title="Evergreen Script - Update your Software, the lazy way" Height="470" Width="840">
    <Grid x:Name="Evergreen_GUI">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="13*"/>
            <ColumnDefinition Width="234*"/>
            <ColumnDefinition Width="586*"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="Image_Logo" Height="100" Margin="467,0,19,0" VerticalAlignment="Top" Width="100" Source="$PSScriptRoot\img\Logo_DEYDA_no_cta.png" Grid.Column="2" ToolTip="www.deyda.net"/>
        <Button x:Name="Button_Start" Content="Start" HorizontalAlignment="Left" Margin="258,375,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
        <Button x:Name="Button_Cancel" Content="Cancel" HorizontalAlignment="Left" Margin="353,375,0,0" VerticalAlignment="Top" Width="75" Grid.Column="2"/>
        <Label x:Name="Label_SelectMode" Content="Select Mode" HorizontalAlignment="Left" Margin="15.5,10,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Download" Content="Download" HorizontalAlignment="Left" Margin="15.5,41,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Install" Content="Install" HorizontalAlignment="Left" Margin="102.5,41,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <Label x:Name="Label_SelectLanguage" Content="Select Language" HorizontalAlignment="Left" Margin="73,10,0,0" VerticalAlignment="Top" Grid.Column="2"/>
        <ComboBox x:Name="Box_Language" HorizontalAlignment="Left" Margin="86,37,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="2" ToolTip="If this is selectable at download!">
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
        <Label x:Name="Label_SelectArchitecture" Content="Select Architecture" HorizontalAlignment="Left" Margin="241,10,0,0" VerticalAlignment="Top" Grid.Column="2"/>
        <ComboBox x:Name="Box_Architecture" HorizontalAlignment="Left" Margin="275,37,0,0" VerticalAlignment="Top" SelectedIndex="0" RenderTransformOrigin="0.864,0.591" Grid.Column="2" ToolTip="If this is selectable at download!">
            <ListBoxItem Content="x64"/>
            <ListBoxItem Content="x86"/>
        </ComboBox>
        <Label x:Name="Label_Explanation" Content="When software download can be filtered on language or architecture." HorizontalAlignment="Left" Margin="58,59,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2"/>
        <Label x:Name="Label_Software" Content="Select Software" HorizontalAlignment="Left" Margin="15.5,70,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_7Zip" Content="7 Zip" HorizontalAlignment="Left" Margin="15.5,101,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_AdobeProDC" Content="Adobe Pro DC" HorizontalAlignment="Left" Margin="15.5,121,0,0" VerticalAlignment="Top" Grid.Column="1" ToolTip="Update Only!"/>
        <CheckBox x:Name="Checkbox_AdobeReaderDC" Content="Adobe Reader DC" HorizontalAlignment="Left" Margin="15.5,141,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_BISF" Content="BIS-F" HorizontalAlignment="Left" Margin="15.5,161,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_CitrixHypervisorTools" Content="Citrix Hypervisor Tools" HorizontalAlignment="Left" Margin="15.5,181,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_CitrixWorkspaceApp" Content="Citrix Workspace App" HorizontalAlignment="Left" Margin="15.5,201,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <ComboBox x:Name="Box_CitrixWorkspaceApp" HorizontalAlignment="Left" Margin="173,198,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.ColumnSpan="2" Grid.Column="1">
            <ListBoxItem Content="Current Release"/>
            <ListBoxItem Content="Long Term Service Release"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_Filezilla" Content="Filezilla" HorizontalAlignment="Left" Margin="15.5,221,0,0" VerticalAlignment="Top" Grid.Column="1" />
        <CheckBox x:Name="Checkbox_FoxitReader" Content="Foxit Reader" HorizontalAlignment="Left" Margin="15.5,241,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1" ToolTip="No silent installation"/>
        <CheckBox x:Name="Checkbox_GoogleChrome" Content="Google Chrome" HorizontalAlignment="Left" Margin="15.5,261,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Greenshot" Content="Greenshot" HorizontalAlignment="Left" Margin="15.5,281,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_KeePass" Content="KeePass" HorizontalAlignment="Left" Margin="16.5,301,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_mRemoteNG" Content="mRemoteNG" HorizontalAlignment="Left" Margin="16.5,321,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSEdge" Content="Microsoft Edge" HorizontalAlignment="Left" Margin="243,101,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_MSFSlogix" Content="Microsoft FSLogix" HorizontalAlignment="Left" Margin="243,121,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_MSOffice2019" Content="Microsoft Office 2019" HorizontalAlignment="Left" Margin="243,141,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_MSOneDrive" Content="Microsoft OneDrive" HorizontalAlignment="Left" Margin="243,161,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2" ToolTip="Machine Based Install"/>
        <ComboBox x:Name="Box_MSOneDrive" HorizontalAlignment="Left" Margin="407,154,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="2" ToolTip="Machine Based Install">
            <ListBoxItem Content="Insider Ring"/>
            <ListBoxItem Content="Production Ring"/>
            <ListBoxItem Content="Enterprise Ring"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_MSTeams" Content="Microsoft Teams" HorizontalAlignment="Left" Margin="243,181,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2" ToolTip="Machine Based Install"/>
        <ComboBox x:Name="Box_MSTeams" HorizontalAlignment="Left" Margin="407,176,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="2" ToolTip="Machine Based Install">
            <ListBoxItem Content="Preview Ring"/>
            <ListBoxItem Content="General Ring"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_Firefox" Content="Mozilla Firefox" HorizontalAlignment="Left" Margin="243,201,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <ComboBox x:Name="Box_Firefox" HorizontalAlignment="Left" Margin="407,198,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="Current"/>
            <ListBoxItem Content="ESR"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_NotepadPlusPlus" Content="Notepad ++" HorizontalAlignment="Left" Margin="243,221,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_OpenJDK" Content="Open JDK" HorizontalAlignment="Left" Margin="243,241,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_OracleJava8" Content="Oracle Java 8" HorizontalAlignment="Left" Margin="243,261,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_TreeSize" Content="TreeSize" HorizontalAlignment="Left" Margin="243,281,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <ComboBox x:Name="Box_TreeSize" HorizontalAlignment="Left" Margin="407,277,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="2">
            <ListBoxItem Content="Free"/>
            <ListBoxItem Content="Professional"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_VLCPlayer" Content="VLC Player" HorizontalAlignment="Left" Margin="243,301,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_VMWareTools" Content="VMWare Tools" HorizontalAlignment="Left" Margin="243,321,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_WinSCP" Content="WinSCP" HorizontalAlignment="Left" Margin="243,341,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_SelectAll" Content="Select All" HorizontalAlignment="Left" Margin="9,386,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="2"/>
        <Label x:Name="Label_author" Content="Manuel Winkel / @deyda84 / www.deyda.net / 2021" HorizontalAlignment="Left" Margin="309,404,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="2"/>
        <CheckBox x:Name="Checkbox_MS365Apps" Content="Microsoft 365 Apps" HorizontalAlignment="Left" Margin="17,341,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <ComboBox x:Name="Box_MS365Apps" HorizontalAlignment="Left" Margin="173,337,0,0" VerticalAlignment="Top" SelectedIndex="4" Grid.Column="1" Grid.ColumnSpan="2">
            <ListBoxItem Content="Current (Preview)"/>
            <ListBoxItem Content="Current"/>
            <ListBoxItem Content="Monthly Enterprise"/>
            <ListBoxItem Content="Semi-Annual Enterprise (Preview)"/>
            <ListBoxItem Content="Semi-Annual Enterprise"/>
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
    try{
        $Form=[Windows.Markup.XamlReader]::Load( $reader )
    }
    catch{
        Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
        throw
    }

    # Load XAML Objects In PowerShell  
    $xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
        try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
        catch{throw}
    } | out-null
 
    Function Get-FormVariables{
        #if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
        #write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
        get-variable WPF*
    }

    Get-FormVariables | out-null

    # Set Variable
    $Script:install = $true
    $Script:download = $true

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
        switch ($LastSetting[8]) {
            1 { $WPFCheckbox_7ZIP.IsChecked = "True"}
        }
        switch ($LastSetting[9]) {
            1 { $WPFCheckbox_AdobeProDC.IsChecked = "True"}
        }
        switch ($LastSetting[10]) {
            1 { $WPFCheckbox_AdobeReaderDC.IsChecked = "True"}
        }
        switch ($LastSetting[11]) {
            1 { $WPFCheckbox_BISF.IsChecked = "True"}
        }
        switch ($LastSetting[12]) {
            1 { $WPFCheckbox_CitrixHypervisorTools.IsChecked = "True"}
        }
        switch ($LastSetting[13]) {
            1 { $WPFCheckbox_CitrixWorkspaceApp.IsChecked = "True"}
        }
        switch ($LastSetting[14]) {
            1 { $WPFCheckbox_Filezilla.IsChecked = "True"}
        }
        switch ($LastSetting[15]) {
            1 { $WPFCheckbox_Firefox.IsChecked = "True"}
        }
        switch ($LastSetting[16]) {
            1 { $WPFCheckbox_FoxitReader.IsChecked = "True"}
        }
        switch ($LastSetting[17]) {
            1 { $WPFCheckbox_MSFSLogix.IsChecked = "True"}
        }
        switch ($LastSetting[18]) {
            1 { $WPFCheckbox_GoogleChrome.IsChecked = "True"}
        }
        switch ($LastSetting[19]) {
            1 { $WPFCheckbox_Greenshot.IsChecked = "True"}
        }
        switch ($LastSetting[20]) {
            1 { $WPFCheckbox_KeePass.IsChecked = "True"}
        }
        switch ($LastSetting[21]) {
            1 { $WPFCheckbox_mRemoteNG.IsChecked = "True"}
        }
        switch ($LastSetting[22]) {
            1 { $WPFCheckbox_MS365Apps.IsChecked = "True"}
        }
        switch ($LastSetting[23]) {
            1 { $WPFCheckbox_MSEdge.IsChecked = "True"}
        }
        switch ($LastSetting[24]) {
            1 { $WPFCheckbox_MSOffice2019.IsChecked = "True"}
        }
        switch ($LastSetting[25]) {
            1 { $WPFCheckbox_MSOneDrive.IsChecked = "True"}
        }
        switch ($LastSetting[26]) {
            1 { $WPFCheckbox_MSTeams.IsChecked = "True"}
        }
        switch ($LastSetting[27]) {
            1 { $WPFCheckbox_NotePadPlusPlus.IsChecked = "True"}
        }
        switch ($LastSetting[28]) {
            1 { $WPFCheckbox_OpenJDK.IsChecked = "True"}
        }
        switch ($LastSetting[29]) {
            1 { $WPFCheckbox_OracleJava8.IsChecked = "True"}
        }
        switch ($LastSetting[30]) {
            1 { $WPFCheckbox_TreeSize.IsChecked = "True"}
        }
        switch ($LastSetting[31]) {
            1 { $WPFCheckbox_VLCPlayer.IsChecked = "True"}
        }
        switch ($LastSetting[32]) {
            1 { $WPFCheckbox_VMWareTools.IsChecked = "True"}
        }
        switch ($LastSetting[33]) {
            1 { $WPFCheckbox_WinSCP.IsChecked = "True"}
        }
        switch ($LastSetting[34]) {
            True { $WPFCheckbox_Download.IsChecked = "True"}
        }
        switch ($LastSetting[35]) {
            True { $WPFCheckbox_Install.IsChecked = "True"}
        }
    }
    
    #// MARK: Event Handler
    # Checkbox SelectAll
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
        $WPFCheckbox_MSFSLogix.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GoogleChrome.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Greenshot.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_KeePass.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_mRemoteNG.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MS365Apps.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSEdge.IsChecked = $WPFCheckbox_SelectAll.IsChecked
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
    })

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
        $WPFCheckbox_MSFSLogix.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_GoogleChrome.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_Greenshot.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_KeePass.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_mRemoteNG.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MS365Apps.IsChecked = $WPFCheckbox_SelectAll.IsChecked
        $WPFCheckbox_MSEdge.IsChecked = $WPFCheckbox_SelectAll.IsChecked
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
    })

    # Button Start                                                                    
    $WPFButton_Start.Add_Click({
        if ($WPFCheckbox_Download.IsChecked -eq $True) {$Script:install = $false}
        else {$Script:install = $true}
        if ($WPFCheckbox_Install.IsChecked -eq $True) {$Script:download = $false}
        else {$Script:download = $true}
        if ($WPFCheckbox_7Zip.IsChecked -eq $true) {$Script:7ZIP = 1}
        else {$Script:7ZIP = 0}
        if ($WPFCheckbox_AdobeProDC.IsChecked -eq $true) {$Script:AdobeProDC = 1}
        else {$Script:AdobeProDC = 0}
        if ($WPFCheckbox_AdobeReaderDC.IsChecked -eq $true) {$Script:AdobeReaderDC = 1}
        else {$Script:AdobeReaderDC = 0}
        if ($WPFCheckbox_BISF.IsChecked -eq $true) {$Script:BISF = 1}
        else {$Script:BISF = 0}
        if ($WPFCheckbox_CitrixHypervisorTools.IsChecked -eq $true) {$Script:Citrix_Hypervisor_Tools = 1}
        else {$Script:Citrix_Hypervisor_Tools = 0}
        if ($WPFCheckbox_CitrixWorkspaceApp.IsChecked -eq $true) {$Script:Citrix_WorkspaceApp = 1}
        else {$Script:Citrix_WorkspaceApp = 0}
        if ($WPFCheckbox_Filezilla.IsChecked -eq $true) {$Script:Filezilla = 1}
        else {$Script:Filezilla = 0}
        if ($WPFCheckbox_Firefox.IsChecked -eq $true) {$Script:Firefox = 1}
        else {$Script:Firefox = 0}
        if ($WPFCheckbox_MSFSLogix.IsChecked -eq $true) {$Script:FSLogix = 1}
        else {$Script:FSLogix = 0}
        if ($WPFCheckbox_FoxitReader.Ischecked -eq $true) {$Script:Foxit_Reader = 1}
        else {$Script:Foxit_Reader = 0}
        if ($WPFCheckbox_GoogleChrome.ischecked -eq $true) {$Script:GoogleChrome = 1}
        else {$Script:GoogleChrome = 0}
        if ($WPFCheckbox_Greenshot.ischecked -eq $true) {$Script:Greenshot = 1}
        else {$Script:Greenshot = 0}
        if ($WPFCheckbox_KeePass.ischecked -eq $true) {$Script:KeePass = 1}
        else {$Script:KeePass = 0}
        if ($WPFCheckbox_mRemoteNG.ischecked -eq $true) {$Script:mRemoteNG = 1}
        else {$Script:mRemoteNG = 0}
        if ($WPFCheckbox_MS365Apps.ischecked -eq $true) {$Script:MS365Apps = 1}
        else {$Script:MS365Apps = 0}
        if ($WPFCheckbox_MSEdge.ischecked -eq $true) {$Script:MSEdge = 1}
        else {$Script:MSEdge = 0}
        if ($WPFCheckbox_MSOffice2019.ischecked -eq $true) {$Script:MSOffice2019 = 1}
        else {$Script:MSOffice2019 = 0}
        if ($WPFCheckbox_MSOneDrive.ischecked -eq $true) {$Script:MSOneDrive = 1}
        else {$Script:MSOneDrive = 0}
        if ($WPFCheckbox_MSTeams.ischecked -eq $true) {$Script:MSTeams = 1}
        else {$Script:MSTeams = 0}
        if ($WPFCheckbox_NotePadPlusPlus.ischecked -eq $true) {$Script:NotePadPlusPlus = 1}
        else {$Script:NotePadPlusPlus = 0}
        if ($WPFCheckbox_OpenJDK.ischecked -eq $true) {$Script:OpenJDK = 1}
        else {$Script:OpenJDK = 0}
        if ($WPFCheckbox_OracleJava8.ischecked -eq $true) {$Script:OracleJava8 = 1}
        else {$Script:OracleJava8 = 0}
        if ($WPFCheckbox_TreeSize.ischecked -eq $true) {$Script:TreeSize = 1}
        else {$Script:TreeSize = 0}
        if ($WPFCheckbox_VLCPlayer.ischecked -eq $true) {$Script:VLCPlayer = 1}
        else {$Script:VLCPlayer = 0}
        if ($WPFCheckbox_VMWareTools.ischecked -eq $true) {$Script:VMWareTools = 1}
        else {$Script:VMWareTools = 0}
        if ($WPFCheckbox_WinSCP.ischecked -eq $true) {$Script:WinSCP = 1}
        else {$Script:WinSCP = 0}
        $Script:Language = $WPFBox_Language.SelectedIndex
        $Script:Architecture = $WPFBox_Architecture.SelectedIndex
        $Script:FirefoxChannel = $WPFBox_Firefox.SelectedIndex
        $Script:CitrixWorkspaceAppRelease = $WPFBox_CitrixWorkspaceApp.SelectedIndex
        $Script:MS365AppsChannel = $WPFBox_MS365Apps.SelectedIndex
        $Script:MSOneDriveRing = $WPFBox_MSOneDrive.SelectedIndex
        $Script:MSTeamsRing = $WPFBox_MSTeams.SelectedIndex
        $Script:TreeSizeType = $WPFBox_TreeSize.SelectedIndex
        $Language,$Architecture,$CitrixWorkspaceAppRelease,$MS365AppsChannel,$MSOneDriveRing,$MSTeamsRing,$FirefoxChannel,$TreeSizeType,$7ZIP,$AdobeProDC,$AdobeReaderDC,$BISF,$Citrix_Hypervisor_Tools,$Citrix_WorkspaceApp,$Filezilla,$Firefox,$Foxit_Reader,$FSLogix,$GoogleChrome,$Greenshot,$KeePass,$mRemoteNG,$MS365Apps,$MSEdge,$MSOffice2019,$MSOneDrive,$MSTeams,$NotePadPlusPlus,$OpenJDK,$OracleJava8,$TreeSize,$VLCPlayer,$VMWareTools,$WinSCP,$WPFCheckbox_Download.IsChecked,$WPFCheckbox_Install.IsChecked | out-file -filepath "$PSScriptRoot\LastSetting.txt"
        Write-Verbose "GUI MODE" -Verbose
        $Form.Close()
    })

    # Button Cancel                                                                    
    $WPFButton_Cancel.Add_Click({
        $Script:install = $true
        $Script:download = $true
        Write-Verbose "GUI MODE Canceled - Nothing happens" -Verbose
        $Form.Close()
    })

    # Image Logo
    $WPFImage_Logo.Add_MouseLeftButtonUp({
        [system.Diagnostics.Process]::start('https://www.deyda.net')
    })

    # Shows the form
    $Form.ShowDialog() | out-null
}

#===========================================================================

Write-Verbose "Setting Variables" -Verbose
Write-Output ""

#// MARK: Define and Reset Variables
$Date = $Date = Get-Date -UFormat "%m.%d.%Y"
$Script:install = $install
$Script:download = $download

if ($list -eq $True) {
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
    
}
else {
    Clear-Variable -name 7ZIP,AdobeProDC,AdobeReaderDC,BISF,Citrix_Hypervisor_Tools,Filezilla,Firefox,Foxit_Reader,FSLogix,Greenshot,GoogleChrome,KeePass,mRemoteNG,MS365Apps,MSEdge,MSOffice2019,MSTeams,NotePadPlusPlus,MSOneDrive,OpenJDK,OracleJava8,TreeSize,VLCPlayer,VMWareTools,WinSCP,Citrix_WorkspaceApp,Architecture,FirefoxChannel,CitrixWorkspaceAppRelease,Language,MS365AppsChannel,MSOneDriveRing,MSTeamsRing,TreeSizeType -ErrorAction SilentlyContinue
    gui_mode
}

# Disable progress bar while downloading
$ProgressPreference = 'SilentlyContinue'

#// MARK: Variable definition (Architecture,Language etc)
switch ($Architecture) {
    0 { $ArchitectureClear = 'x64'}
    1 { $ArchitectureClear = 'x86'}
}

switch ($Language) {
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
switch ($LanguageClear) {
    Polish { $AdobeLanguageClear = 'English'}
    Portuguese { $AdobeLanguageClear = 'English'}
    Russian { $AdobeLanguageClear = 'English'}
    Swedish { $AdobeLanguageClear = 'English'}
}

$AdobeArchitectureClear = 'x86'
switch ($LanguageClear) {
    English { $AdobeArchitectureClear = $ArchitectureClear}
}

switch ($CitrixWorkspaceAppRelease) {
    0 { $CitrixWorkspaceAppReleaseClear = 'Current Release'}
    1 { $CitrixWorkspaceAppReleaseClear = 'LTSR'}
}

switch ($FirefoxChannel) {
    0 { $FirefoxChannelClear = 'LATEST'}
    1 { $FirefoxChannelClear = 'ESR'}
}

switch ($LanguageClear) {
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

$FoxitReaderLanguageClear = $LanguageClear
switch ($LanguageClear) {
    Japanese { $FoxitReaderLanguageClear = 'English'}
}

switch ($MS365AppsChannel) {
    0 { $MS365AppsChannelClear = 'CurrentPreview'}
    1 { $MS365AppsChannelClear = 'Current'}
    2 { $MS365AppsChannelClear = 'MonthlyEnterprise'}
    3 { $MS365AppsChannelClear = 'SemiAnnualPreview'}
    4 { $MS365AppsChannelClear = 'SemiAnnual'}
}

switch ($MS365AppsChannel) {
    0 { $MS365AppsChannelClearDL = 'Monthly (Targeted)'}
    1 { $MS365AppsChannelClearDL = 'Monthly'}
    2 { $MS365AppsChannelClearDL = 'Monthly Enterprise'}
    3 { $MS365AppsChannelClearDL = 'Semi-Annual Channel (Targeted)'}
    4 { $MS365AppsChannelClearDL = 'Semi-Annual Channel'}
}

switch ($Architecture) {
    0 { $MS365AppsArchitectureClear = '64'}
    1 { $MS365AppsArchitectureClear = '32'}
}

switch ($Language) {
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

switch ($MSOneDriveRing) {
    0 { $MSOneDriveRingClear = 'Insider'}
    1 { $MSOneDriveRingClear = 'Production'}
    2 { $MSOneDriveRingClear = 'Enterprise'}
}

switch ($MSTeamsRing) {
    0 { $MSTeamsRingClear = 'Preview'}
    1 { $MSTeamsRingClear = 'General'}
}

if ($install -eq $False) {
    #// Mark: Install/Update Evergreen module
    Write-Output ""
    Write-Verbose "Installing/updating Evergreen module... please wait" -Verbose
    Write-Output ""
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies}
    if (!(Get-Module -ListAvailable -Name Evergreen)) {Install-Module Evergreen -Force | Import-Module Evergreen}
    Update-Module Evergreen -force

    Write-Output "Starting downloads..."
    Write-Output ""

    #// Mark: Download 7-ZIP
    if ($7ZIP -eq 1) {
        $Product = "7-Zip"
        $PackageName = "7-Zip_" + "$ArchitectureClear"
        $7ZipD = Get-7zip | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.URI -like "*exe*" }
        $Version = $7ZipD.Version
        $URL = $7ZipD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download Adobe Pro DC Update
    if ($AdobeProDC -eq 1) {
        $Product = "Adobe Pro DC"
        $PackageName = "Adobe_Pro_DC_Update"
        $AdobeProD = Get-AdobeAcrobat | Where-Object { $_.Type -eq "Updater" -and $_.Track -eq "DC" }
        $Version = $AdobeProD.Version
        $URL = $AdobeProD.uri
        $InstallerType = "msp"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Include *.msp, *.log, Version.txt, Download* -Recurse
            Start-Transcript $LogPS
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

    #// Mark: Download Adobe Reader DC
    if ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"
        $PackageName = "Adobe_Reader_DC_"
        $AdobeReaderD = Get-AdobeAcrobatReaderDC | Where-Object {$_.Architecture -eq "$AdobeArchitectureClear" -and $_.Language -eq "$AdobeLanguageClear"}
        $Version = $AdobeReaderD.Version
        $URL = $AdobeReaderD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "$AdobeArchitectureClear" + "$AdobeLanguageClear" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$AdobeArchitectureClear" + "_$AdobeLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $AdobeArchitectureClear $AdobeLanguageClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Include *.msp, *.log, Version.txt, Download* -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $AdobeArchitectureClear $AdobeLanguageClear $Version" -Verbose
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

    #// Mark: Download BIS-F
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
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Exclude *.ps1, *.lnk -Recurse
            Start-Transcript $LogPS
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

    #// Mark: Download Citrix Hypervisor Tools
    if ($Citrix_Hypervisor_Tools -eq 1) {
        $Product = "Citrix Hypervisor Tools"
        $PackageName = "managementagent" + "$ArchitectureClear"
        $CitrixHypervisor = Get-CitrixVMTools | Where-Object {$_.Architecture -eq "$ArchitectureClear"} | Select-Object -Last 1
        $Version = $CitrixHypervisor.Version
        $URL = $CitrixHypervisor.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\Citrix\$Product\" + ($Source))
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

    #// Mark: Download Citrix WorkspaceApp
    if ($Citrix_WorkspaceApp -eq 1) {
        $Product = "Citrix WorkspaceApp $CitrixWorkspaceAppReleaseClear"
        $PackageName = "CitrixWorkspaceApp"
        $WSACD = Get-CitrixWorkspaceApp -WarningAction:SilentlyContinue | Where-Object { $_.Title -like "*Workspace*" -and "*$CitrixWorkspaceAppReleaseClear*" -and $_.Platform -eq "Windows" -and $_.Title -like "*$CitrixWorkspaceAppReleaseClear*" }
        $Version = $WSACD.Version
        $URL = $WSACD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -EA SilentlyContinue
        if (!(Test-Path -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility")) { New-Item -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility" -ItemType Directory | Out-Null }
        if (!(Test-Path -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.exe")) {
            Write-Verbose "Download Citrix Receiver Cleanup Utility" -Verbose
            Invoke-WebRequest -Uri https://fileservice.citrix.com/downloadspecial/support/article/CTX137494/downloads/ReceiverCleanupUtility.zip -OutFile ("$PSScriptRoot\Citrix\ReceiverCleanupUtility\" + "ReceiverCleanupUtility.zip")
            Expand-Archive -path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.zip" -destinationpath "$PSScriptRoot\Citrix\ReceiverCleanupUtility\"
            Remove-Item -Path "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.zip" -Force
            Write-Verbose "Download Citrix Receiver Cleanup Utility finished" -Verbose
        }
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\Citrix\$Product")) { New-Item -Path "$PSScriptRoot\Citrix\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\Citrix\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\Citrix\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $Version" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\Citrix\$Product\" + ($Source))
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

    #// Mark: Download Filezilla
    if ($Filezilla -eq 1) {
        $Product = "Filezilla"
        $PackageName = "Filezilla-win64"
        $FilezillaD = Get-Filezilla | Where-Object { $_.URI -like "*win64*"}
        $Version = $FilezillaD.Version
        $URL = $FilezillaD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
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

    #// Mark: Download Foxit Reader
    if ($Foxit_Reader -eq 1) {
        $Product = "Foxit Reader"
        $PackageName = "FoxitReader-Setup-" + "$FoxitReaderLanguageClear"
        $Foxit_ReaderD = Get-FoxitReader | Where-Object {$_.Language -eq "$FoxitReaderLanguageClear"}
        $Version = $Foxit_ReaderD.Version
        $URL = $Foxit_ReaderD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitReaderLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $FoxitReaderLanguageClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $FoxitReaderLanguageClear $Version" -Verbose
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

    #// Mark: Download Greenshot
    if ($Greenshot -eq 1) {
        $Product = "Greenshot"
        $PackageName = "Greenshot-INSTALLER-x86"
        $GreenshotD = Get-Greenshot | Where-Object { $_.Architecture -eq "x86" -and $_.URI -like "*INSTALLER*" -and $_.Type -like "exe"}
        $Version = $GreenshotD.Version
        $URL = $GreenshotD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
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

    #// Mark: Download Google Chrome
    if ($GoogleChrome -eq 1) {
        $Product = "Google Chrome"
        $PackageName = "googlechromestandaloneenterprise_" + "$ArchitectureClear"
        $ChromeD = Get-GoogleChrome | Where-Object { $_.Architecture -eq "$ArchitectureClear" }
        $Version = $ChromeD.Version
        $URL = $ChromeD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download KeePass
    if ($KeePass -eq 1) {
        $Product = "KeePass"
        $PackageName = "KeePass"
        $KeePassD = Get-KeePass | Where-Object { $_.URI -like "*msi*" }
        $Version = $KeePassD.Version
        $URL = $KeePassD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"-EA SilentlyContinue 
        Write-Verbose "Download $Product" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
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

    #// Mark: Download Microsoft 365 Apps
    if ($MS365Apps -eq 1) {
        $Product = "Microsoft 365 Apps"
        $PackageName = "setup_" + "$MS365AppsChannelClear"
        $MS365AppsD = Get-Microsoft365Apps | Where-Object {$_.Channel -eq "$MS365AppsChannelClearDL"}
        $Version = $MS365AppsD.Version
        $URL = $MS365AppsD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product $MS365AppsChannelClear setup file" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!(Test-Path -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear")) {New-Item -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear" -ItemType Directory | Out-Null}
        if (!(Test-Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\remove.xml" -PathType leaf)) {
            Write-Verbose "Create remove.xml" -Verbose
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
            Write-Verbose "Create remove.xml finished!" -Verbose
        }
        if (!(Test-Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\install.xml" -PathType leaf)) {
            Write-Verbose "Create install.xml" -Verbose
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
            Write-Verbose "Create install.xml finished!" -Verbose
        }
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            $LogPS = "$PSScriptRoot\$Product\$MS365AppsChannelClear\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\$MS365AppsChannelClear\*" -Recurse -Exclude install.xml,remove.xml
            Start-Transcript $LogPS
            Set-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt" -Value "$Version"
            Write-Verbose "Starting Download of $Product $MS365AppsChannelClear $Version setup file" -Verbose
            Invoke-WebRequest -Uri $URL -OutFile ("$PSScriptRoot\$Product\$MS365AppsChannelClear\" + ($Source))
            Write-Verbose "Stop logging" -Verbose
            Stop-Transcript
            Write-Output ""
        }
        else {
            Write-Verbose "No new version available" -Verbose
            Write-Output ""
        }
        # Download Apps 365 install files
        if (!(Test-Path -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Office\Data\$Version")) {
            Write-Verbose "Download $Product $MS365AppsChannelClear $MS365AppsArchitectureClear $MS365AppsLanguageClear $Version install files" -Verbose
            $DApps365 = @(
                "/download install.xml"
            )
            set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
            Start-Process ".\$Source" -ArgumentList $DApps365 -wait -NoNewWindow
            set-location $PSScriptRoot
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft Edge
    if ($MSEdge -eq 1) {
        $Product = "Microsoft Edge"
        $PackageName = "MicrosoftEdgeEnterprise_" + "$ArchitectureClear"
        $EdgeD = Get-MicrosoftEdge | Where-Object { $_.Platform -eq "Windows" -and $_.Channel -eq "stable" -and $_.Architecture -eq "$ArchitectureClear" }
        #$EdgeURL = $EdgeURL | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $EdgeD.Version
        $URL = $EdgeD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue 
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download Microsoft FSLogix
    if ($FSLogix -eq 1) {
        $Product = "Microsoft FSLogix"
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
            Write-Verbose "Update available" -Verbose
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


    #// Mark: Download Microsoft Office 2019
    if ($MSOffice2019 -eq 1) {
        $Product = "Microsoft Office 2019"
        $PackageName = "setup"
        $MSOffice2019D = Get-Microsoft365Apps | Where-Object {$_.Channel -eq "Office 2019 Enterprise"}
        $Version = $MSOffice2019D.Version
        $URL = $MSOffice2019D.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
        Write-Verbose "Download $Product setup file" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
        if (!(Test-Path "$PSScriptRoot\$Product\remove.xml" -PathType leaf)) {
            Write-Verbose "Create remove.xml" -Verbose
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
            Write-Verbose "Create remove.xml finished!" -Verbose
        }
        if (!(Test-Path "$PSScriptRoot\$Product\install.xml" -PathType leaf)) {
            Write-Verbose "Create install.xml" -Verbose
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
            Write-Verbose "Create install.xml finished!" -Verbose
        }
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse -Exclude install.xml,remove.xml
            Start-Transcript $LogPS
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
        # Download MS Office 2019 install files
        if (!(Test-Path -Path "$PSScriptRoot\$Product\Office\Data\$Version")) {
            Write-Verbose "Download $Product $MS365AppsArchitectureClear $MS365AppsLanguageClear $Version install files" -Verbose
            $DOffice2019 = @(
                "/download install.xml"
            )
            set-location $PSScriptRoot\$Product
            Start-Process ".\setup.exe" -ArgumentList $DOffice2019 -wait -NoNewWindow
            set-location $PSScriptRoot
            Write-Output ""
        }
    }

    #// Mark: Download Microsoft OneDrive
    if ($MSOneDrive -eq 1) {
        $Product = "Microsoft OneDrive"
        $PackageName = "OneDriveSetup-" + "$MSOneDriveRingClear"
        $MSOneDriveD = Get-MicrosoftOneDrive | Where-Object { $_.Ring -eq "$MSOneDriveRingClear" -and $_.Type -eq "Exe" } | Sort-Object -Property Version -Descending | Select-Object -Last 1
        $Version = $MSOneDriveD.Version
        $URL = $MSOneDriveD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSOneDriveRingClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $MSOneDriveRingClear Ring" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $MSOneDriveRingClear Ring $Version" -Verbose
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

    #// Mark: Download Microsoft Teams
    if ($MSTeams -eq 1) {
        $Product = "Microsoft Teams"
        $PackageName = "Teams_" + "$ArchitectureClear" + "_$MSTeamsRingClear"
        $TeamsD = Get-MicrosoftTeams | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.Ring -eq "$MSTeamsRingClear"}
        $Version = $TeamsD.Version
        $URL = $TeamsD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear $MSTeamsRingClear Ring" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Include *.msi, *.log, Version.txt, Download* -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $MSTeamsRingClear Ring $Version" -Verbose
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

    #// Mark: Download Mozilla Firefox
    if ($Firefox -eq 1) {
        $Product = "Mozilla Firefox"
        $PackageName = "Firefox_Setup_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "_$FFLanguageClear"
        $FirefoxD = Get-MozillaFirefox | Where-Object { $_.Type -eq "msi" -and $_.Architecture -eq "$ArchitectureClear" -and $_.Channel -like "*$FirefoxChannelClear*" -and $_.Language -eq "$FFLanguageClear"}
        $Version = $FirefoxD.Version
        $URL = $FirefoxD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "$FFLanguageClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear $Version" -Verbose
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

    #// Mark: Download mRemoteNG
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
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
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

    #// Mark: Download Notepad ++
    if ($NotePadPlusPlus -eq 1) {
        $Product = "NotePadPlusPlus"
        $PackageName = "NotePadPlusPlus_" + "$ArchitectureClear"
        $NotepadD = Get-NotepadPlusPlus | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.Type -eq "exe" }
        $Version = $NotepadD.Version
        $URL = $NotepadD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Get-ChildItem "$PSScriptRoot\$Product\" -Exclude lang | Remove-Item -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download openJDK
    if ($OpenJDK -eq 1) {
        $Product = "open JDK"
        $PackageName = "OpenJDK" + "$ArchitectureClear"
        $OpenJDKD = Get-OpenJDK | Where-Object { $_.Architecture -eq "$ArchitectureClear" -and $_.URI -like "*msi*" } | Sort-Object -Property Version -Descending | Select-Object -First 1
        $Version = $OpenJDKD.Version
        $URL = $OpenJDKD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download OracleJava8
    if ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        $PackageName = "OracleJava8_" + "$ArchitectureClear"
        $OracleJava8D = Get-OracleJava8 | Where-Object { $_.Architecture -eq "$ArchitectureClear" }
        $Version = $OracleJava8D.Version
        $URL = $OracleJava8D.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download TreeSize
    if ($TreeSize -eq 1) {
        switch ($TreeSizeType) {
            0 {
                $Product = "TreeSize Free"
                $PackageName = "TreeSize_Free"
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
                    Write-Verbose "Update available" -Verbose
                    if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS
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
            1 {
                $Product = "TreeSize Professional"
                $PackageName = "TreeSize_Professional"
                $TreeSizeProfD = Get-JamTreeSizeProfessional
                $Version = $TreeSizeProfD.Version
                $URL = $TreeSizeProfD.uri
                $InstallerType = "exe"
                $Source = "$PackageName" + "." + "$InstallerType"
                $CurrentVersion = Get-Content -Path "$PSScriptRoot\$Product\Version.txt" -EA SilentlyContinue
                Write-Verbose "Download $Product" -Verbose
                Write-Host "Download Version: $Version"
                Write-Host "Current Version: $CurrentVersion"
                if (!($CurrentVersion -eq $Version)) {
                    Write-Verbose "Update available" -Verbose
                    if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
                    $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
                    Remove-Item "$PSScriptRoot\$Product\*" -Recurse
                    Start-Transcript $LogPS
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
        }
    }

    #// Mark: Download VLC Player
    if ($VLCPlayer -eq 1) {
        $Product = "VLC Player"
        $PackageName = "VLC-Player_" + "$ArchitectureClear"
        $VLCD = Get-VideoLanVlcPlayer | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "$ArchitectureClear" -and $_.Type -eq "MSI" }
        $Version = $VLCD.Version
        $URL = $VLCD.uri
        $InstallerType = "msi"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download VMWareTools
    if ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"
        $PackageName = "VMWareTools_" + "$ArchitectureClear"
        $VMWareToolsD = Get-VMwareTools | Where-Object { $_.Architecture -eq "$ArchitectureClear" }
        $Version = $VMWareToolsD.Version
        $URL = $VMWareToolsD.uri
        $InstallerType = "exe"
        $Source = "$PackageName" + "." + "$InstallerType"
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $CurrentVersion = Get-Content -Path "$VersionPath" -EA SilentlyContinue
        Write-Verbose "Download $Product $ArchitectureClear" -Verbose
        Write-Host "Download Version: $Version"
        Write-Host "Current Version: $CurrentVersion"
        if (!($CurrentVersion -eq $Version)) {
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) { New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null }
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
            Set-Content -Path "$VersionPath" -Value "$Version"
            Write-Verbose "Starting Download of $Product $ArchitectureClear $Version" -Verbose
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

    #// Mark: Download WinSCP
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
            Write-Verbose "Update available" -Verbose
            if (!(Test-Path -Path "$PSScriptRoot\$Product")) {New-Item -Path "$PSScriptRoot\$Product" -ItemType Directory | Out-Null}
            $LogPS = "$PSScriptRoot\$Product\" + "$Product $Version.log"
            Remove-Item "$PSScriptRoot\$Product\*" -Recurse
            Start-Transcript $LogPS
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
    $LogFileName = ("$ENV:COMPUTERNAME - $Date.log")
    $LogFile = Join-path $LogDir $LogFileName
    $LogTemp = "$env:windir\Logs\Evergreen"

    # Create the log directories if they don't exist
    if (!(Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType directory | Out-Null }
    if (!(Test-Path $LogTemp)) { New-Item -Path $LogTemp -ItemType directory | Out-Null }

    # Create new log file (overwrite existing one)
    New-Item $LogFile -ItemType "file" -force | Out-Null
    DS_WriteLog "I" "START SCRIPT - " $LogFile
    DS_WriteLog "-" "" $LogFile
    #========================================================================================================================================
    
    # define Error handling
    # note: do not change these values
    $global:ErrorActionPreference = "Stop"
    if ($verbose){ $global:VerbosePreference = "Continue" }


    #// Mark: Install 7-ZIP
    if ($7ZIP -eq 1) {
        $Product = "7-Zip"

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $SevenZip = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Select-Object -First 1
        If ($SevenZip -eq $NULL) {
            $SevenZip = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*7-Zip*"}).DisplayVersion | Select-Object -First 1
        }
        $7ZipInstaller = "7-Zip_" + "$ArchitectureClear" + ".exe"
        if ($SevenZip -ne $Version) {
            # 7-Zip
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\$7ZipInstaller" -ArgumentList /S
                $p = Get-Process 7-Zip_$ArchitectureClear
                if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
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

    #// Mark: Install Adobe Pro DC
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
                $inst = Start-Process -FilePath msiexec.exe -ArgumentList $mspArgs -Wait
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
                # Update Dienst und Task deaktivieren
                Write-Verbose "Customize Service and Scheduled Task" -Verbose
                Stop-Service AdobeARMservice
                Set-Service AdobeARMservice -StartupType Disabled
                Write-Verbose "Stop and Disable Service $Product finished!" -Verbose
                Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
                Write-Verbose "Disable Scheduled Task $Product finished!" -Verbose
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

    #// Mark: Install Adobe Reader DC
    if ($AdobeReaderDC -eq 1) {
        $Product = "Adobe Reader DC"

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$AdobeArchitectureClear" + "_$AdobeLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Adobe = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Adobe Acrobat Reader*" | Sort-Object -Property DisplayVersion | Select-Object -Last 1 }).DisplayVersion
        $AdobeReaderInstaller = "Adobe_Reader_DC_" + "$AdobeArchitectureClear" + "$AdobeLanguageClear" + ".exe"
        if ($Adobe -ne $Version) {
            # Adobe Reader DC
            Write-Verbose "Installing $Product $AdobeArchitectureClear $AdobeLanguageClear" -Verbose
            DS_WriteLog "I" "Installing $Product $AdobeArchitectureClear $AdobeLanguageClear" $LogFile
            $Options = @(
                "/sAll"
                "/rs"
                "/msi EULA_ACCEPT=YES ENABLE_OPTIMIZATION=YES DISABLEDESKTOPSHORTCUT=1 UPDATE_MODE=0 DISABLE_ARM_SERVICE_INSTALL=1"
            )
            try	{
                Start-Process "$PSScriptRoot\$Product\$AdobeReaderInstaller" -ArgumentList $Options
                $p = Get-Process Adobe_Reader_DC_$AdobeArchitectureClear$AdobeLanguageClear
                if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product $AdobeArchitectureClear $AdobeLanguageClear finished!" -Verbose
                }
                # Update Dienst und Task deaktivieren
                Write-Verbose "Customize Service and Scheduled Task" -Verbose
                Stop-Service AdobeARMservice
                Set-Service AdobeARMservice -StartupType Disabled
                Write-Verbose "Stop and Disable Service $Product finished!" -Verbose
                Disable-ScheduledTask -TaskName "Adobe Acrobat Update Task" | Out-Null
                Write-Verbose "Disable Scheduled Task $Product finished!" -Verbose
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

    #// Mark: Install BIS-F
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
                "/L*V $BISFLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product finished!" -Verbose
                    DS_WriteLog "I" "Installation $Product finished!" $LogFile
            }
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
        $BISFLog = "$LogTemp\BISF.log"
        IF ($BISF) {$BISF = $BISF -replace ".{6}$"}
        IF ($BISF -ne $Version) {
            # Base Image Script Framework
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\setup-BIS-F.msi" | Install-MSIFile
                Get-Content $BISFLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $BISFLog
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            # Customize scripts, it's best practise to enable Task Offload and RSS and to disable DEP
            Write-Verbose "Customize scripts $Product" -Verbose
            DS_WriteLog "I" "Customize scripts $Product" $LogFile
            $BISFDir = "C:\Program Files (x86)\Base Image Script Framework (BIS-F)\Framework\SubCall"
            try {
                ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace "DisableTaskOffload' -Value '1'","DisableTaskOffload' -Value '0'") | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
                ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'nx AlwaysOff','nx OptOut') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
                ((Get-Content "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1" -Raw) -replace 'rss=disable','rss=enable') | Set-Content -Path "$BISFDir\Preparation\97_PrepBISF_PRE_BaseImage.ps1"
                Write-Verbose "Customize scripts $Product finished!" -Verbose
            } catch {
                DS_WriteLog "E" "Error when customizing scripts (error: $($Error[0]))" $LogFile
            }
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
        DS_WriteLog "-" "" $LogFile
        write-Output ""
    }

    #// Mark: Install Citrix Hypervisor Tools
    IF ($Citrix_Hypervisor_Tools -eq 1) {
        $Product = "Citrix Hypervisor Tools"

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
                "/quiet"
                "/norestart"
                "/L*V $CitrixHypLog"
                )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if ($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $ArchitectureClear finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\Citrix\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $HypTools = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Hypervisor*"}).DisplayVersion
        $CitrixHypLog = "$LogTemp\CitrixHypervisor.log"
        If ($HypTools -eq $NULL) {
            $HypTools = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Hypervisor*"}).DisplayVersion
        }
        If ($HypTools) {$HypTools = $HypTools.Insert(3,'.0')}
        $HypToolsInstaller = "managementagent" + "$ArchitectureClear" + ".msi"
        IF ($HypTools -ne $Version) {
            # Citrix Hypervisor Tools
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try {
                "$PSScriptRoot\Citrix\$Product\$HypToolsInstaller" | Install-MSIFile
                Get-Content $CitrixHypLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $CitrixHypLog
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

    #// Mark: Install Citrix WorkspaceApp
    IF ($Citrix_WorkspaceApp -eq 1) {
        $Product = "Citrix WorkspaceApp $CitrixWorkspaceAppReleaseClear"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\Citrix\$Product\Version.txt"
        $WSA = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Citrix Workspace*" -and $_.UninstallString -like "*Trolley*"}).DisplayVersion
        $UninstallWSACR = "$PSScriptRoot\Citrix\ReceiverCleanupUtility\ReceiverCleanupUtility.exe"
        IF ($WSA -ne $Version) {
            # Citrix WSA Uninstallation
            Write-Verbose "Uninstalling Citrix Workspace App / Receiver" -Verbose
            DS_WriteLog "I" "Uninstalling Citrix Workspace App / Receiver" $LogFile
            try	{
                Start-process $UninstallWSACR -ArgumentList '/silent /disableCEIP' -NoNewWindow -Wait
            } catch {
                DS_WriteLog "E" "Error Uninstalling Citrix Workspace App / Receiver (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Verbose "Uninstalling and Cleanup Citrix Workspace App / Receiver finished!" -Verbose

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
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\Citrix\$Product\CitrixWorkspaceApp.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product finished!" -Verbose
                }
                Write-Verbose "Customize $Product" -Verbose
                reg add "HKLM\SOFTWARE\Wow6432Node\Policies\Citrix" /v EnableX1FTU /t REG_DWORD /d 0 /f | Out-Null
                reg add "HKCU\Software\Citrix\Splashscreen" /v SplashscrrenShown /d 1 /f | Out-Null
                reg add "HKLM\SOFTWARE\Policies\Citrix" /f /v EnableFTU /t REG_DWORD /d 0 | Out-Null
                Write-Verbose "Customizing $Product finished!" -Verbose
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

    #// Mark: Install Filezilla
    IF ($Filezilla -eq 1) {
        $Product = "Filezilla"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Filezilla = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Filezilla*"}).DisplayVersion
        IF ($Filezilla -ne $Version) {
            # Filezilla
            $Options = @(
                "/S"
                "/user=all"
            )
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\Filezilla-win64.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
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

    #// Mark: Install Foxit Reader
    IF ($Foxit_Reader -eq 1) {
        $Product = "Foxit Reader"

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FoxitReaderLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $FReader = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Foxit Reader*"}).DisplayVersion
        $FoxitReaderInstaller = "FoxitReader-Setup-" + "$FoxitReaderLanguageClear" + ".exe"
        IF ($FReader -ne $Version) {
            # Foxit Reader
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
            Write-Verbose "Installing $Product $FoxitReaderLanguageClear" -Verbose
            DS_WriteLog "I" "Installing $Product $FoxitReaderLanguageClear" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$FoxitReaderInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Remove-Item -Path "$env:PUBLIC\Desktop\Foxit Reader.lnk" -Force
                    Write-Verbose "Installation $Product $FoxitReaderLanguageClear finished!" -Verbose
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

    #// Mark: Install Greenshot
    IF ($Greenshot -eq 1) {
        $Product = "Greenshot"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $Greenshot = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Greenshot*"}).DisplayVersion
        IF ($Greenshot -ne $Version) {
            # Greenshot
            $Options = @(
                "/VERYSILENT"
                "/NORESTART"
                "/SUPPRESSMSGBOXES"
            )
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\Greenshot-INSTALLER-x86.exe" -ArgumentList $Options -PassThru -ErrorAction Stop
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

    #// Mark: Install Google Chrome
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
                "/L*V $ChromeLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if ($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $ArchitectureClear finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Chrome = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion
        $ChromeLog = "$LogTemp\GoogleChrome.log"
        If ($Chrome -eq $NULL) {
            $Chrome = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Google Chrome"}).DisplayVersion
        }
        $ChromeInstaller = "googlechromestandaloneenterprise_" + "$ArchitectureClear" + ".msi"
        IF ($Chrome -ne $Version) {
            # Google Chrome
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try {
                "$PSScriptRoot\$Product\$ChromeInstaller" | Install-MSIFile
                Get-Content $ChromeLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $ChromeLog
                # Update Dienst und Task deaktivieren
                Write-Verbose "Customize Service and Scheduled Task" -Verbose
                Stop-Service gupdate
                Set-Service gupdate -StartupType Disabled
                Stop-Service gupdatem
                Set-Service gupdatem -StartupType Disabled
                Write-Verbose "Stop and Disable Service $Product finished!" -Verbose
                Disable-ScheduledTask -TaskName "GoogleUpdateTaskMachineCore" | Out-Null
                Disable-ScheduledTask -TaskName "GoogleUpdateTaskMachineUA" | Out-Null
                Disable-ScheduledTask -TaskName "GPUpdate on Startup" | Out-Null
                Write-Verbose "Disable Scheduled Task $Product finished!" -Verbose
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

    #// Mark: Install KeePass
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
                "/L*V $KeePassLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product finished!" -Verbose
                DS_WriteLog "I" "Installation $Product finished!" $LogFile
            }
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
        $KeePassLog = "$LogTemp\KeePass.log"
        IF ($KeePass) {$KeePass = $KeePass -replace ".{2}$"}
        IF ($KeePass -ne $Version) {
            # KeePass
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\KeePass.msi" | Install-MSIFile
                Get-Content $KeePassLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $KeePassLog
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

    #// Mark: Install Microsoft Apps 365
    IF ($MS365Apps -eq 1) {
        $Product = "Microsoft 365 Apps"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\$MS365AppsChannelClear\Version.txt"
        $MS365AppsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft 365 Apps*"}).DisplayVersion
        If ($MS365AppsV -eq $NULL) {
            $MS365AppsV = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft 365 Apps*"}).DisplayVersion
        }
        $MS365AppsInstaller = "setup_" + "$MS365AppsChannelClear" + ".exe"
        IF ($MS365AppsV -ne $Version) {
            # MS365Apps Uninstallation
            $Options = @(
                "/configure remove.xml"
            )
            Write-Verbose "Uninstalling Office 2019 or Microsoft 365 Apps" -Verbose
            DS_WriteLog "I" "Uninstalling Office 2019 or Microsoft 365 Apps" $LogFile
            try	{
                set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                Start-Process -FilePath ".\$MS365AppsInstaller" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Verbose "Uninstallation Office 2019 or Microsoft 365 Apps finished!" -Verbose
            } catch {
                DS_WriteLog "E" "Error uninstalling Office 2019 or Microsoft 365 Apps (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            # MS365Apps Installation
            $Options = @(
                "/configure install.xml"
            )
            Write-Verbose "Installing $Product $MS365AppsChannelClear $MS365AppsArchitectureClear $MS365AppsLanguageClear" -Verbose
            DS_WriteLog "I" "Installing $Product $MS365AppsChannelClear $MS365AppsArchitectureClear $MS365AppsLanguageClear" $LogFile
            try	{
                set-location $PSScriptRoot\$Product\$MS365AppsChannelClear
                Start-Process -FilePath ".\$MS365AppsInstaller" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Verbose "Installation $Product $MS365AppsChannelClear $MS365AppsArchitectureClear $MS365AppsLanguageClear finished!" -Verbose
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

    #// Mark: Install Microsoft Edge
    IF ($MSEdge -eq 1) {
        $Product = "Microsoft Edge"

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
                "REBOOT=ReallySuppress"
                "DONOTCREATEDESKTOPSHORTCUT=TRUE"
                "DONOTCREATETASKBARSHORTCUT=true"
                "/L*V $EdgeLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $ArchitectureClear finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================
        
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Edge = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion
        $EdgeLog = "$LogTemp\MSEdge.log"
        If ($Edge -eq $NULL) {
            $Edge = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq "Microsoft Edge"}).DisplayVersion
        }
        $EdgeInstaller = "MicrosoftEdgeEnterprise_" + "$ArchitectureClear" + ".msi"
        IF ($Edge -ne $Version) {
            # MS Edge
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try {
                "$PSScriptRoot\$Product\$EdgeInstaller" | Install-MSIFile
                Get-Content $EdgeLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $EdgeLog
                #Disable Microsoft Edge auto update
                Write-Verbose "Disable Edge Update" -Verbose
                If (!(Test-Path -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate)) {
                    New-Item -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate
                    New-ItemProperty -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -Name UpdateDefault -Value 0 -PropertyType DWORD
                }
                else {
                    $EdgeUpdateState = Get-ItemProperty -path "HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate" | Select-Object -Expandproperty "UpdateDefault"
                    If ($EdgeUpdateState -ne "0") {New-ItemProperty -Path HKLM:SOFTWARE\Policies\Microsoft\EdgeUpdate -Name UpdateDefault -Value 0 -PropertyType DWORD}
                }
                #Configure Microsoft Edge update service to manual startup
                Set-Service -Name edgeupdate -StartupType Disabled
                Set-Service -Name edgeupdatem -StartupType Disabled
                # Execute the Microsoft Edge browser replacement task to make sure that the legacy Microsoft Edge browser is tucked away
                # This is only needed on Windows 10 versions where Microsoft Edge is not included in the OS.
                #Start-Process -FilePath "${env:ProgramFiles(x86)}\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe" -Wait -ArgumentList "/browserreplacement"
                Write-Verbose "Disable Edge Update finished !" -Verbose
                #Disable update tasks
                Write-Verbose "Disable Scheduled Task" -Verbose
                Start-Sleep -s 5
                Get-ScheduledTask -TaskName MicrosoftEdgeUpdate* | Disable-ScheduledTask | Out-Null
                Write-Verbose "Disable Scheduled Task $Product finished!" -Verbose
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            # Disable Citrix API Hooks (MS Edge) on Citrix VDA
            $(
                $RegPath = "HKLM:SYSTEM\CurrentControlSet\services\CtxUvi"
                IF (Test-Path $RegPath) {
                    Write-Verbose "Disable Citrix API Hooks" -Verbose
                    $RegName = "UviProcessExcludes"
                    $EdgeRegvalue = "msedge.exe"
                    # Get current values in UviProcessExcludes
                    $CurrentValues = Get-ItemProperty -Path $RegPath | Select-Object -ExpandProperty $RegName
                    # Add the msedge.exe value to existing values in UviProcessExcludes
                    Set-ItemProperty -Path $RegPath -Name $RegName -Value "$CurrentValues$EdgeRegvalue;"
                    Write-Verbose "Disable Citrix API Hooks for $Product finished!" -Verbose
                }
            ) | Out-Null
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }

    #// Mark: Install Microsoft FSLogix
    IF ($FSLogix -eq 1) {
        $Product = "Microsoft FSLogix"
        $OS = (Get-WmiObject Win32_OperatingSystem).Caption
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
                    Start-process $UninstallFSL -ArgumentList '/uninstall /quiet /norestart' -NoNewWindow -Wait
                    Start-process $UninstallFSLRE -ArgumentList '/uninstall /quiet /norestart' -NoNewWindow -Wait
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
                $inst = Start-Process "$PSScriptRoot\$Product\Install\FSLogixAppsSetup.exe" -ArgumentList '/install /norestart /quiet'  -NoNewWindow
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product Setup finished!" -Verbose
                }
                $inst = Start-Process "$PSScriptRoot\$Product\Install\FSLogixAppsRuleEditorSetup.exe" -ArgumentList '/install /norestart /quiet'  -NoNewWindow
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product Rule Editor finished!" -Verbose
                }
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile
            }
            # Application post deployment tasks (Thx to Kasper https://github.com/kaspersmjohansen)
            Write-Verbose "Applying $Product post setup customizations" -Verbose
            Write-Verbose "Post setup customizations for $OS" -Verbose
            If ($OS -Like "*Windows Server 2019*" -or $OS -eq "Microsoft Windows 10 Enterprise for Virtual Desktops") {
                Write-Verbose "Deactivate FSLogix RoamSearch" -Verbose
                New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "0" -Type DWORD
            }
            If ($OS -Like "*Windows 10*" -and $OS -ne "Microsoft Windows 10 Enterprise for Virtual Desktops") {
                Write-Verbose "Activate FSLogix RoamSearch" -Verbose
                New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Apps" -Name "RoamSearch" -Value "1" -Type DWORD
            }
            # Implement user based group policy processing fix
            Write-Verbose "Deactivate FSLogix GroupPolicy" -Verbose
            If (!(Test-Path -Path HKLM:SOFTWARE\FSLogix\Profiles)) {
                New-Item -Path "HKLM:SOFTWARE\FSLogix" -Name Profiles
                New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Name "GroupPolicyState" -Value "0" -Type DWORD
            }
            else {
                $FSLGroupState = Get-ItemProperty -path "HKLM:SOFTWARE\FSLogix\Profiles" | Select-Object -Expandproperty "GroupPolicyState"
                If ($FSLGroupState -eq "1") {New-ItemProperty -Path "HKLM:SOFTWARE\FSLogix\Profiles" -Name "GroupPolicyState" -Value "0" -Type DWORD}
            }
            If (!(Get-ScheduledTask -TaskName "Restart Windows Search Service on Event ID 2")) {
                Write-Verbose "Implement scheduled task to restart Windows Search service on Event ID 2" -Verbose
                # Implement scheduled task to restart Windows Search service on Event ID 2
                # Define CIM object variables
                # This is needed for accessing the non-default trigger settings when creating a schedule task using Powershell
                $Class = cimclass MSFT_TaskEventTrigger root/Microsoft/Windows/TaskScheduler
                $Trigger = $class | New-CimInstance -ClientOnly
                $Trigger.Enabled = $true
                $Trigger.Subscription = "<QueryList><Query Id=`"0`" Path=`"Application`"><Select Path=`"Application`">*[System[Provider[@Name='Microsoft-Windows-Search-ProfileNotify'] and EventID=2]]</Select></Query></QueryList>"
                # Define additional variables containing scheduled task action and scheduled task principal
                $A = New-ScheduledTaskAction –Execute powershell.exe -Argument "Restart-Service Wsearch"
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
            }
            Write-Verbose "Applying $Product post setup customizations finished !" -Verbose
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
        }
        # Stop, if no new version is available
        Else {
            Write-Verbose "No Update available for $Product" -Verbose
            Write-Output ""
        }
    }


    #// Mark: Install Microsoft Office 2019
    IF ($MSOffice2019 -eq 1) {
        $Product = "Microsoft Office 2019"

        # Check, if a new version is available
        $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
        $MSOffice2019V = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Microsoft Office*"}).DisplayVersion
        IF ($MSOffice2019V -ne $Version) {
            # MS Office 2019 Uninstallation
            $Options = @(
                "/configure remove.xml"
            )
            Write-Verbose "Uninstalling Office 2019 or Microsoft 365 Apps" -Verbose
            DS_WriteLog "I" "Uninstalling Office 2019 or Microsoft 365 Apps" $LogFile
            try	{
                set-location $PSScriptRoot\$Product
                Start-Process -FilePath ".\setup.exe" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Verbose "Uninstallation Office 2019 or Microsoft 365 Apps finished!" -Verbose
            } catch {
                DS_WriteLog "E" "Error uninstalling Office 2019 or Microsoft 365 Apps (error: $($Error[0]))" $LogFile
            }
            DS_WriteLog "-" "" $LogFile
            Write-Output ""
            # MS Office 2019 Installation
            $Options = @(
                "/configure install.xml"
            )
            Write-Verbose "Installing $Product $MS365AppsArchitectureClear $MS365AppsLanguageClear" -Verbose
            DS_WriteLog "I" "Installing $Product $MS365AppsArchitectureClear $MS365AppsLanguageClear" $LogFile
            try	{
                set-location $PSScriptRoot\$Product
                Start-Process -FilePath ".\setup.exe" -ArgumentList $Options -NoNewWindow -wait
                set-location $PSScriptRoot
                Write-Verbose "Installation $Product $MS365AppsArchitectureClear $MS365AppsLanguageClear finished!" -Verbose
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

    #// Mark: Install Microsoft OneDrive
    IF ($MSOneDrive -eq 1) {
        $Product = "Microsoft OneDrive"

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$MSOneDriveRingClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $MSOneDriveV = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OneDrive*"}).DisplayVersion
        IF ($MSOneDriveV -ne $Version) {
            # Installation MSOneDrive
            Write-Verbose "Installing $Product $MSOneDriveRingClear" -Verbose
            DS_WriteLog "I" "Installing $Product $MSOneDriveRingClear" $LogFile
            $Options = @(
                "/ALLUSERS"
                "/SILENT"
            )
            try	{
                $null = Start-Process "$PSScriptRoot\$Product\OneDriveSetup.exe" -ArgumentList $Options -NoNewWindow -PassThru
                while (Get-Process -Name "OneDriveSetup" -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                Write-Verbose "Installation $Product $MSOneDriveRingClear finished!" -Verbose
                # OneDrive starts automatically after setup. kill!
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

    #// Mark: Install Microsoft Teams
    IF ($MSTeams -eq 1) {
        $Product = "Microsoft Teams"

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
                "REBOOT=ReallySuppress"
                "ALLUSER=1"
                "ALLUSERS=1"
                "OPTIONS='noAutoStart=true'"
                "/qn"
                "/L*V $TeamsLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $ArchitectureClear $MSTeamsRingClear Ring finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $ArchitectureClear $MSTeamsRingClear Ring finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================
        
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + "_$MSTeamsRingClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Teams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).DisplayVersion
        $TeamsInstaller = "Teams_" + "$ArchitectureClear" + "_$MSTeamsRingClear" + ".msi"
        $TeamsLog = "$LogTemp\MSTeams.log"
        IF ($Teams) {$Teams = $Teams.Insert(5,'0')}
        IF ($Teams -ne $Version) {
            #Uninstalling MS Teams
            Write-Verbose "Uninstalling $Product" -Verbose
            DS_WriteLog "I" "Uninstalling $Product" $LogFile
            try {
                $UninstallTeams = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).UninstallString
                $UninstallTeams = $UninstallTeams -Replace("MsiExec.exe /I","")
                Start-Process -FilePath msiexec.exe -ArgumentList "/X $UninstallTeams /qn /L*V $TeamsLog"
                Start-Sleep 20
                Get-Content $TeamsLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $TeamsLog
                Write-Verbose "Uninstalling $Product finished!" -Verbose
                DS_WriteLog "I" "Uninstalling $Product finished!" $LogFile
            } catch {
                DS_WriteLog "E" "Error installing $Product (error: $($Error[0]))" $LogFile       
            }
            DS_WriteLog "-" "" $LogFile
            #MS Teams Installation
            #Registry key for Teams machine-based install with Citrix VDA (Thx to Kasper https://github.com/kaspersmjohansen)
            If (!(Test-Path 'HKLM:\Software\Citrix\PortICA\')) {
                If (!(Test-Path 'HKLM:\Software\Citrix\')) {New-Item -Path "HKLM:Software\Citrix"}
                New-Item -Path "HKLM:Software\Citrix\PortICA"
            }
            Write-Verbose "Installing $Product $ArchitectureClear $MSTeamsRingClear Ring" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear $MSTeamsRingClear Ring" $LogFile
            try {
                "$PSScriptRoot\$Product\$TeamsInstaller" | Install-MSIFile
                Start-Sleep 5
                Get-Content $TeamsLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $TeamsLog
                reg add "HKLM\SOFTWARE\Citrix\CtxHook\AppInit_Dlls\SfrHook" /v Teams.exe /t REG_DWORD /d 204 /f | Out-Null
                <# Prevents MS Teams from starting at logon, better do this with WEM or similar
                Write-Verbose "Customize $Product Autorun" -Verbose
                Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" -Name "Teams" -Force
                Write-Verbose "Customize $Product Autorun finished!" -Verbose#>
                #Remove public desktop shortcut (Thx to Kasper https://github.com/kaspersmjohansen)
                Remove-Item -Path "$env:PUBLIC\Desktop\Microsoft Teams.lnk" -Force
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

    #// Mark: Install Mozilla Firefox
    IF ($Firefox -eq 1) {
        $Product = "Mozilla Firefox"

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
                "/q"
                "DESKTOP_SHORTCUT=false"
                "TASKBAR_SHORTCUT=false"
                "INSTALL_MAINTENANCE_SERVICE=false"
                "/L*V $FirefoxLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if ($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "$FFLanguageClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Firefox = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Firefox*"}).DisplayVersion
        $FirefoxLog = "$LogTemp\Firefox.log"
        If ($Firefox -eq $NULL) {
            $Firefox = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Firefox*"}).DisplayVersion
        }
        $FirefoxInstaller = "Firefox_Setup_" + "$FirefoxChannelClear" + "$ArchitectureClear" + "_$FFLanguageClear" + ".msi"
        IF ($Firefox -ne $Version) {
            # Firefox
            Write-Verbose "Installing $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear" -Verbose
            DS_WriteLog "I" "Installing $Product $FirefoxChannelClear $ArchitectureClear $FFLanguageClear" $LogFile
            try {
                "$PSScriptRoot\$Product\$FirefoxInstaller" | Install-MSIFile
                Get-Content $FirefoxLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $FirefoxLog
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

    #// Mark: Install mRemoteNG
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
                "/L*V $mRemoteLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            if($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product finished!" -Verbose
                DS_WriteLog "I" "Installation $Product finished!" $LogFile
            }
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
        $mRemoteLog = "$LogTemp\mRemote.log"
        IF ($mRemoteNG) {$mRemoteNG = $mRemoteNG -replace ".{6}$"}
        IF ($mRemoteNG -ne $Version) {
            # mRemoteNG
            Write-Verbose "Installing $Product" -Verbose
            DS_WriteLog "I" "Installing $Product" $LogFile
            try {
                "$PSScriptRoot\$Product\mRemoteNG.msi" | Install-MSIFile
                Get-Content $mRemoteLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $mRemoteLog
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

    #// Mark: Install Notepad ++
    IF ($NotePadPlusPlus -eq 1) {
        $Product = "NotepadPlusPlus"

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $Notepad = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion
        If ($Notepad -eq $NULL) {
            $Notepad = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Notepad++*"}).DisplayVersion
        }
        $NotepadPlusPlusInstaller = "NotePadPlusPlus_" + "$ArchitectureClear" + ".exe"
        IF ($Notepad -ne $Version) {
            # Installation Notepad++
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try	{
                Start-Process "$PSScriptRoot\$Product\$NotepadPlusPlusInstaller" -ArgumentList /S -NoNewWindow
                $p = Get-Process NotePadPlusPlus_$ArchitectureClear
		        if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
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

    #// Mark: Install OpenJDK
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
                "/L*V $openJDKLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $ArchitectureClear finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $OpenJDK = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion
        $openJDKLog = "$LogTemp\OpenJDK.log"
        If ($OpenJDK -eq $NULL) {
            $OpenJDK = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*OpenJDK*"}).DisplayVersion
        }
        $OpenJDKInstaller = "OpenJDK" + "$ArchitectureClear" + ".msi"
        IF ($Version) {$Version = $Version -replace ".-"}
        IF ($OpenJDK -ne $Version) {
            # OpenJDK
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try {
                "$PSScriptRoot\$Product\$OpenJDKInstaller" | Install-MSIFile
                Start-Sleep 25
                Get-Content $openJDKLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $openJDKLog
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

    #// Mark: Install OracleJava8
    if ($OracleJava8 -eq 1) {
        $Product = "Oracle Java 8"
        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        IF ($Version) {$Version = $Version -replace "^.{2}"}
        IF ($Version) {$Version = $Version -replace "\."}
        IF ($Version) {$Version = $Version -replace "_"}
        IF ($Version) {$Version = $Version -replace "-b"}
        $OracleJava = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java 8*"}).DisplayVersion
        If ($OracleJava -eq $NULL) {
            $OracleJava = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Java 8*"}).DisplayVersion
        }
        IF ($OracleJava) {$OracleJava = $OracleJava -replace "\."}
        $OracleJavaInstaller = "OracleJava8_" + "$ArchitectureClear" +".exe"
        if ($OracleJava -ne $Version) {
            # Oracle Java 8
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            $Options = @(
                "/s INSTALL_SILENT=Enable AUTO_UPDATE=Disable REBOOT=Disable SPONSORS=Disable REMOVEOUTOFDATEJRES=1 WEB_ANALYTICS=Disable"
            )
            try	{
                Start-Process "$PSScriptRoot\$Product\$OracleJavaInstaller" -ArgumentList $Options -NoNewWindow
                $p = Get-Process OracleJava8_$ArchitectureClear
                if ($p) {
                    $p.WaitForExit()
                    Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
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

    #// Mark: Install TreeSize
    IF ($TreeSize -eq 1) {
        switch ($TreeSizeType) {
            0 {
                $Product = "TreeSize Free"

                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
                $Version = $Version.Insert(3,'.')
                $TreeSize = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion
                IF ($TreeSize -ne $Version) {
                    # Installation Tree Size Free
                    Write-Verbose "Installing $Product" -Verbose
                    DS_WriteLog "I" "Installing $Product" $LogFile
                    try	{
                        Start-Process "$PSScriptRoot\$Product\TreeSize_Free.exe" -ArgumentList /VerySilent -NoNewWindow -Wait
                        $p = Get-Process TreeSize_Free
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
            1 {
                $Product = "TreeSize Professional"

                # Check, if a new version is available
                $Version = Get-Content -Path "$PSScriptRoot\$Product\Version.txt"
                $Version = $Version.Insert(3,'.')
                $TreeSize = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*TreeSize*"}).DisplayVersion
                IF ($TreeSize -ne $Version) {
                    # Installation Tree Size Free
                    Write-Verbose "Installing $Product" -Verbose
                    DS_WriteLog "I" "Installing $Product" $LogFile
                    try	{
                        Start-Process "$PSScriptRoot\$Product\TreeSize_Professional.exe" -ArgumentList /VerySilent -NoNewWindow -Wait
                        $p = Get-Process TreeSize_Professional
                        if ($p) {
                            $p.WaitForExit()
                            Write-Verbose "Installation $Product finished!" -Verbos
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
        }
    }

    #// Mark: Install VLC Player
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
                "/L*V $VLCLog"
            )
            if ($targetDir) {
                if (!(Test-Path $targetDir)) {
                    throw "Path to installation directory $($targetDir) is invalid. Please check path and file name!"
                }
                $arguments += "INSTALLDIR=`"$targetDir`""
            }
            $inst = $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -NoNewWindow -PassThru
            if($inst -ne $null) {
                Wait-Process -InputObject $inst
                Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
                DS_WriteLog "I" "Installation $Product $ArchitectureClear finished!" $LogFile
            }
            if ($process.ExitCode -eq 0) {
            }
            else {
                Write-Verbose "Installer Exit Code  $($process.ExitCode) for file  $($msifile)"
            }
        }
        #========================================================================================================================================

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $VLC = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion
        $VLCLog = "$LogTemp\VLC.log"
        If ($VLC -eq $NULL) {
            $VLC = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VLC*"}).DisplayVersion
        }
        IF ($VLC) {$VLC = $VLC -replace ".{2}$"}
        $VLCInstaller = "VLC-Player_" + "$ArchitectureClear" +".msi"
        IF ($VLC -ne $Version) {
            # VLC Player
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try {
                "$PSScriptRoot\$Product\$VLCInstaller" | Install-MSIFile
                Get-Content $VLCLog | Add-Content $LogFile -Encoding ASCI
                Remove-Item $VLCLog
                Remove-Item -Path "$env:PUBLIC\Desktop\VLC media player.lnk" -Force
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

    #// Mark: Install VMWareTools
    IF ($VMWareTools -eq 1) {
        $Product = "VMWare Tools"

        # Check, if a new version is available
        $VersionPath = "$PSScriptRoot\$Product\Version_" + "$ArchitectureClear" + ".txt"
        $Version = Get-Content -Path "$VersionPath"
        $VMWT = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion
        If ($VMWT -eq $NULL) {
            $VMWT = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*VMWare*"}).DisplayVersion
        }
        IF ($VMWT) {$VMWT = $VMWT -replace ".{9}$"}
        $VMWareToolsInstaller = "VMWareTools_" + "$ArchitectureClear" +".exe"
        IF ($VMWT -ne $Version) {
            # VMWareTools Installation
            $Options = @(
                "/s"
                "/v"
                "/qn REBOOT=Y"
            )
            Write-Verbose "Installing $Product $ArchitectureClear" -Verbose
            DS_WriteLog "I" "Installing $Product $ArchitectureClear" $LogFile
            try	{
                $inst = Start-Process -FilePath "$PSScriptRoot\$Product\$VMWareToolsInstaller" -ArgumentList $Options -PassThru -ErrorAction Stop
                if($inst -ne $null) {
                    Wait-Process -InputObject $inst
                    Write-Verbose "Installation $Product $ArchitectureClear finished!" -Verbose
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

    #// Mark: Install WinSCP
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
}