![MODES Logo, Office, VSCode and an M in a hexagon, overlapped](https://davecra.files.wordpress.com/2020/04/modes-1.png)
# The Microsoft Office Development Environment Script (MODES)
This script is provided to make it easy to setup your Office Development Envrionment for Office Web Add-in Development (OfficeJS). If you are new to Office Web Add-in Development, this script will help you quickly install everything you need to get started.

This supports the following platforms:
* Windows platform

NOTE: Other platform support is TBD.

## Install on Windows
To install, download the modes.ps1 to your desktop, right-click and select "Run in Powershell".

If this does not work, try these steps:
1) Make sure the modes.ps1 is on your desktop
2) Press Windows+R, type "powershell" and press enter 
3) Once in powershell, type "cd $env:USERPROFILE\desktop" and press enter
4) Type ".\modes.ps1" and press enter
5) When asked to run the script, type A and press enter

**Important Note:** This repository is in-progress and the PowerShell script is not signed and may not run on your desktop.

## What this installs
The following items will occur when this script is run:

- It will download the latest version of Visual Studio Code (VSCode) from Microsoft
- It will download the latest version of Node from NodeJS.org
- It will download v1.4.9 of NPM from NodeJS.org
- It will install VSCode silently
- It will create the Node directory in your user profile
- It will place node.exe in the that directory
- It will unzip NPM into that directory
- It will update the PATH environment variable for your user profile
- It will update NPM
- It will install Yeoman via NPM
- It will install the Office Yeoman generator via NPM
- It will install http-server via NPM
- It will then delete the 3 files downloaded above
