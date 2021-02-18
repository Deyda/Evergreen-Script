#XAML Code kann zwischen @" und "@ ersetzt werden:
[xml]$XAML = @"
<Window x:Class="GUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:GUI"
        mc:Ignorable="d"
        Title="Evergreen - Update your Software, the lazy way" Height="450" Width="800">
    <Grid x:Name="Evergreen_GUI">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="13*"/>
            <ColumnDefinition Width="31*"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="Image_Logo" HorizontalAlignment="Left" Height="100" Margin="448,0,0,0" VerticalAlignment="Top" Width="100" Source="Logo_DEYDA_no_cta.png" Grid.Column="1"/>
        <Button x:Name="Button_Start" Content="Start" HorizontalAlignment="Left" Margin="258,366,0,0" VerticalAlignment="Top" Width="75" Grid.Column="1"/>
        <Button x:Name="Button_Cancel" Content="Cancel" HorizontalAlignment="Left" Margin="353,366,0,0" VerticalAlignment="Top" Width="75" Grid.Column="1"/>
        <Label x:Name="Label_SelectMode" Content="Select Mode" HorizontalAlignment="Left" Margin="28,10,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="Checkbox_Download" Content="Download" HorizontalAlignment="Left" Margin="28,41,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="Checkbox_Install" Content="Install" HorizontalAlignment="Left" Margin="115,41,0,0" VerticalAlignment="Top"/>
        <Label x:Name="Label_Software" Content="Select Software" HorizontalAlignment="Left" Margin="28,70,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="Checkbox_7Zip" Content="7 Zip" HorizontalAlignment="Left" Margin="28,101,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="Checkbox_AdobeProDC" Content="Adobe Pro DC" HorizontalAlignment="Left" Margin="28,121,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="Checkbox_AdobeReaderDC" Content="Adobe Reader DC" HorizontalAlignment="Left" Margin="28,141,0,0" VerticalAlignment="Top" />
        <CheckBox x:Name="Checkbox_BISF" Content="BIS-F" HorizontalAlignment="Left" Margin="28,161,0,0" VerticalAlignment="Top" />
        <CheckBox x:Name="Checkbox_CitrixHypervisorTools" Content="Citrix Hypervisor Tools" HorizontalAlignment="Left" Margin="28,181,0,0" VerticalAlignment="Top" />
        <CheckBox x:Name="Checkbox_CitrixWorkspaceApp" Content="Citrix Workspace App" HorizontalAlignment="Left" Margin="28,201,0,0" VerticalAlignment="Top" />
        <ComboBox x:Name="Box_CitrixWorksapceApp" HorizontalAlignment="Left" Margin="180,197,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.ColumnSpan="2">
            <ListBoxItem Content="Current Release"/>
            <ListBoxItem Content="Long Term Service Release"/>
        </ComboBox>
        <Label x:Name="Label_SelectLanguage" Content="Select Language" HorizontalAlignment="Left" Margin="73,10,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <ComboBox x:Name="Box_Language" HorizontalAlignment="Left" Margin="86,37,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="1">
            <ListBoxItem Content="Danish"/>
            <ListBoxItem Content="Dutch"/>
            <ListBoxItem Content="English"/>
            <ListBoxItem Content="French"/>
            <ListBoxItem Content="German"/>
            <ListBoxItem Content="Finnish"/>
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
        <Label x:Name="Label_Explanation" Content="When software download can be filtered on language or architecture." HorizontalAlignment="Left" Margin="64,59,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Filezilla" Content="Filezilla" HorizontalAlignment="Left" Margin="28,221,0,0" VerticalAlignment="Top" />
        <CheckBox x:Name="Checkbox_FoxitReader" Content="Foxit Reader" HorizontalAlignment="Left" Margin="28,241,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133"/>
        <CheckBox x:Name="Checkbox_GoogleChrome" Content="Google Chrome" HorizontalAlignment="Left" Margin="28,261,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133"/>
        <CheckBox x:Name="Checkbox_Greenshot" Content="Greenshot" HorizontalAlignment="Left" Margin="28,281,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133"/>
        <CheckBox x:Name="Checkbox_KeePass" Content="KeePass" HorizontalAlignment="Left" Margin="29,301,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133"/>
        <CheckBox x:Name="Checkbox_mRemoteNG" Content="mRemoteNG" HorizontalAlignment="Left" Margin="29,321,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133"/>
        <CheckBox x:Name="Checkbox_MS365Apps" Content="Microsoft 365 Apps" HorizontalAlignment="Left" Margin="124,101,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <ComboBox x:Name="Box_MS365Apps" HorizontalAlignment="Left" Margin="275,97,0,0" VerticalAlignment="Top" SelectedIndex="4" Grid.Column="1">
            <ListBoxItem Content="Current Channel (Preview) &quot;CurrentPreview&quot;"/>
            <ListBoxItem Content="Current Channel &quot;Current&quot;"/>
            <ListBoxItem Content="Monthly Enterprise Channel &quot;MonthlyEnterprise&quot;"/>
            <ListBoxItem Content="Semi-Annual Enterprise Channel (Preview) &quot;SemiAnnualPreview&quot;"/>
            <ListBoxItem Content="Semi-Annual Enterprise Channel &quot;SemiAnnual&quot;"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_MSEdge" Content="Microsoft Edge" HorizontalAlignment="Left" Margin="124,121,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSFSlogix" Content="Microsoft FSLogix" HorizontalAlignment="Left" Margin="124,141,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSOffice2019" Content="Microsoft Office 2019" HorizontalAlignment="Left" Margin="124,161,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSOneDrive" Content="Microsoft OneDrive" HorizontalAlignment="Left" Margin="124,181,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_MSTeams" Content="Microsoft Teams" HorizontalAlignment="Left" Margin="124,201,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <ComboBox x:Name="Box_MSTeams" HorizontalAlignment="Left" Margin="275,197,0,0" VerticalAlignment="Top" SelectedIndex="1" Grid.Column="1">
            <ListBoxItem Content="Preview Ring"/>
            <ListBoxItem Content="General Ring"/>
        </ComboBox>
        <Label x:Name="Label_SelectArchitecture" Content="Select Architecture" HorizontalAlignment="Left" Margin="241,10,0,0" VerticalAlignment="Top" Grid.Column="1"/>
        <ComboBox x:Name="Box_Architecture" HorizontalAlignment="Left" Margin="275,37,0,0" VerticalAlignment="Top" SelectedIndex="0" RenderTransformOrigin="0.864,0.591" Grid.Column="1">
            <ListBoxItem Content="x64"/>
            <ListBoxItem Content="x86"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_NotepadPlusPlus" Content="Notepad ++" HorizontalAlignment="Left" Margin="124,241,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_OpenJDK" Content="Open JDK" HorizontalAlignment="Left" Margin="124,261,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_OracleJava8" Content="Oracle Java 8" HorizontalAlignment="Left" Margin="124,281,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_TreeSize" Content="TreeSize" HorizontalAlignment="Left" Margin="124,301,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <ComboBox x:Name="Box_TreeSize" HorizontalAlignment="Left" Margin="275,297,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1">
            <ListBoxItem Content="Free"/>
            <ListBoxItem Content="Professional"/>
        </ComboBox>
        <CheckBox x:Name="Checkbox_VMWareTools" Content="VMWare Tools" HorizontalAlignment="Left" Margin="124,321,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_WinSCP" Content="WinSCP" HorizontalAlignment="Left" Margin="124,341,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_SelectAll" Content="Select All" HorizontalAlignment="Left" Margin="9,371,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <CheckBox x:Name="Checkbox_Firefox" Content="Mozilla Firefox" HorizontalAlignment="Left" Margin="124,221,0,0" VerticalAlignment="Top"  RenderTransformOrigin="0.517,1.133" Grid.Column="1"/>
        <ComboBox x:Name="Box_Firefox" HorizontalAlignment="Left" Margin="275,218,0,0" VerticalAlignment="Top" SelectedIndex="0" Grid.Column="1">
            <ListBoxItem Content="Current"/>
            <ListBoxItem Content="ESR"/>
        </ComboBox>
        <ComboBox x:Name="Box_MSOneDrive" HorizontalAlignment="Left" Margin="275,177,0,0" VerticalAlignment="Top" SelectedIndex="2" Grid.Column="1">
            <ListBoxItem Content="Insider Ring"/>
            <ListBoxItem Content="Production Ring"/>
            <ListBoxItem Content="Enterprise Ring"/>
        </ComboBox>
        <Label x:Name="Label_author" Content="Manuel Winkel / @deyda84 / www.deyda.net / 2021" HorizontalAlignment="Left" Margin="309,396,0,0" VerticalAlignment="Top" FontSize="10" Grid.Column="1"/>

    </Grid>
</Window>
"@ -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"','' #-replace wird benötigt, wenn XAML aus Visual Studio kopiert wird.

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}


# Create variables based on form control names.
# Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
Get-Variable var_*

$Null = $window.ShowDialog()