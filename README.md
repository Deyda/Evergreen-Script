# Evergreen
Download, install and update the newest version of several software packages based on the powerful Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein.
https://github.com/aaronparker/Evergreen

I'm no powershell expert, so I'm sure there is much room for improvements! 

## How To
The idea is to select a client or server that periodically checks for updates and if updates are available, downloads them. This can be done every day or once a week by launching the script "Evergreen.ps1 -list" via scheduled task. You decide which software do download by giving it a "0" or "1" in the script.

The "Evergreen.ps1" script must be launched on your clients. If you have a golden master like in Citrix MCS/PVS environments it's sufficient to launch the script only on this machine. This can be done manually or automatic, like you prefer.

If it is run manually, do not use the -list parameter and you will be taken to the GUI to select the software and mode (Download or Install).

## Version check
The updater always checks for the latest version of the Evergreen module, so you don't have to do this. Sometimes the software version found with Evergreen differs from the installed version in the registry, that's stupid, but we can't influence that. Don't blame the Evergreen module!

Let me show you an example:

*MS Teams*

Let's check the installed version:
```
(Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Teams Machine*"}).DisplayVersion
```
The result is: **1.3.0.28779**

Let's check the version with Evergreen:
```
(Get-MicrosoftTeams | Where-Object {$_.Architecture -eq "x64"}).Version
```
The result is: **1.3.00.28779**

So there is one "0" more! We have to insert a "0" to the installed version to be able to compare the versions: 
```
IF ($Teams) {$Teams = $Teams.Insert(5,'0')}
```

# Evergreen-Admx

After deploying several Windows Virtual Desktop environments I decided I no longer wanted to manually download the Admx files I needed, and I wanted a way to keep them up-to-date.

This script solves both problems.
*  Checks for newer versions of the Admx files that are present and processes the new version if found
*  Optionally copies the new Admx files to the Policy Store or Definition folder, or a folder of your chosing

The name I chose for this script is an ode to the Evergreen module (https://github.com/aaronparker/Evergreen) by Aaron Parker (@stealthpuppy).

## How to use

Quick start:
*  Download the script to a location of your chosing (for example: C:\Scripts\EvergreenAdmx)
*  Run or schedule the script

You can also install the script from the PowerShell Gallery ([EvergreenAdmx][poshgallery-evergreenadmx]):
```powershell
Install-Script -Name EvergreenAdmx
```

I have scheduled the script to run daily:

`
Evergreen-Admx.ps1 -WindowsVersion "20H2" -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions"
`
The above execution will keep the central Policy Store up-to-date on a daily basis.

A sample .xml file that you can import in Task Scheduler is provided with this script.

This script processes all the products by default. Simply comment out any products you don't need and the script will skip those.
This will change in a future release.

```
SYNTAX
    D:\Personal Data\amensc\Gits\EvergreenAdmx\Evergreen-Admx.ps1 [[-WindowsVersion] <String>] [[-WorkingDirectory] <String>] [[-PolicyStore] <String>] [[-Languages] <String[]>] [-UseProductFolders] [<CommonParameters>]

DESCRIPTION
    Script to download latest Admx files for several products.
    Optionally copy the latest Admx files to a folder of your chosing, for example a Policy Store.

PARAMETERS
    -WindowsVersion <String>
        The Windows 10 version to get the Admx files for.
        If omitted the newest version supported by this script will be used.

    -WorkingDirectory <String>
        Optionally provide a Working Directory for the script.
        The script will store Admx files in a subdirectory called "admx".
        The script will store downloaded files in a subdirectory called "downloads".
        If omitted the script will treat the script's folder as the working directory.
        
    -PolicyStore <String>
        Optionally provide a Policy Store location to copy the Admx files to after processing.

    -Languages <String[]>
        Optionally provide an array of languages to process. Entries must be in 'xy-XY' format.
        If omitted the script will process 'en-US'.
        
    -UseProductFolders [<SwitchParameter>]
        When specified the extracted Admx files are copied to their respective product folders in a subfolder of 'Admx' in the WorkingDirectory.

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).
    
EXAMPLES
    PS C:\>.\Evergreen-Admx.ps1 -WindowsVersion "20H2" -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders
```

## Admx files

Also see [Change Log][change-log] for a list of supported products.

Now supports
*  Adobe Acrobat Reader DC
*  Base Image Script Framework (BIS-F)
*  Citrix Workspace App
*  FSLogix
*  Google Chrome
*  Microsoft Desktop Optimization Pack
*  Microsoft Edge (Chromium)
*  Microsoft Office
*  Microsoft OneDrive
*  Microsoft Windows 10 (1903/1909/2004/20H2)
*  Mozilla Firefox
*  Zoom Desktop Client

## Notes

I have not tested this script on Windows Core.
Some of the Admx files can only be obtained by installing the package that was downloaded. For instance, the Windows 10 Admx files are in an msi file, the OneDrive Admx files are in the installation folder after installing OneDrive.
So this is what the script does for these packages: installing the package, copying the Admx files, uninstalling the package.

[github-release-badge]: https://img.shields.io/github/release/msfreaks/EvergreenAdmx.svg?style=flat-square
[github-release]: https://github.com/msfreaks/EvergreenAdmx/releases/latest
[code-quality-badge]: https://app.codacy.com/project/badge/Grade/c0efab02b66442399bb16b0493cdfbef?style=flat-square
[code-quality]: https://www.codacy.com/gh/msfreaks/EvergreenAdmx/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=msfreaks/EvergreenAdmx&amp;utm_campaign=Badge_Grade
[license-badge]: https://img.shields.io/github/license/msfreaks/EvergreenAdmx.svg?style=flat-square
[license]: https://github.com/msfreaks/EvergreenAdmx/blob/master/LICENSE
[twitter-follow-badge]: https://img.shields.io/twitter/follow/menschab?style=flat-square
[twitter-follow]: https://twitter.com/menschab?ref_src=twsrc%5Etfw
[change-log]: https://github.com/msfreaks/EvergreenAdmx/blob/main/CHANGELOG.md
[poshgallery-evergreenadmx]: https://www.powershellgallery.com/packages/EvergreenAdmx/
