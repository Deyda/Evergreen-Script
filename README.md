# Evergreen
Download, install and update the newest version of several software packages based on the powerful Evergreen module from Aaron Parker, Bronson Magnan and Trond Eric Haarvarstein.
https://github.com/aaronparker/Evergreen

I'm no powershell expert, so I'm sure there is much room for improvements! 

## How To
The idea is to select a client or server that periodically checks for updates and if updates are available, downloads them. This can be done every day or once a week by launching the script "Evergreen Software Updater.ps1" via scheduled task. You decide which software do download by giving it a "0" or "1" in the script. 

The "Evergreen Software Installer.ps1" script must be launched on your clients. If you have a golden master like in Citrix MCS/PVS environments it's sufficient to launch the script only on this machine. This can be done manually or automatic, like you prefer. 
Again, you decide which package gets installed by "0" or "1". 

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
