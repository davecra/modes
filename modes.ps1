Add-Type -assembly “system.io.compression.filesystem”
<# 
=================================================================
The Microsoft Office Development Environment Script (for Windows)
=================================================================

This script will download and install all the tools needed to 
be able to quickly get started developing OfficeJS Web Add-in 
solutions.

NAME:		modes.ps1
CREATED-BY: David E. Craig (davecra)
VERSION:	1.0.1
DATE:		04-24-2020 @ 04:24:23PM EST	
#>
# VSCODE URL and OUTPUT - always the latest
$vscode_url = "https://aka.ms/win32-x64-user-stable"
$vscode_output = $env:USERPROFILE + "\downloads\vscode.exe"
# NODE URL and OUTPUT - latest version
$node_url = "https://nodejs.org/dist/latest/win-x64/node.exe"
$node_output = $env:USERPROFILE + "\downloads\node.exe"
$node_install_path = $env:USERPROFILE + "\node"
# NPM URL and OUTPUT - specific version 12.6.2 (no way to get latest?)
$npm_url = "https://nodejs.org/dist/npm/npm-1.4.9.zip"
$npm_output = $env:USERPROFILE + "\downloads\npm.zip"
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host -ForegroundColor White -BackgroundColor DarkGreen "=================================================================
The Microsoft Office Development Environment Script (for Windows)
=================================================================" 
Write-Host ""
Write-Host "Created by David E. Craig (davecra)" -ForegroundColor DarkYellow
Write-Host "Visit: https://theofficecontext.com, or" -ForegroundColor DarkYellow
Write-Host "       https://github.com/davecra" -ForegroundColor DarkYellow
Write-Host ""
Write-Host ""
Write-Host "This script will setup the Microsoft Office development environment 
on your Windows PC. It will perform the following steps:"
Write-Host ""
Write-Host " 1)  Download the latest version of VSCode from Microsoft"
Write-Host " 2)  Download the latest version of Node from the NodeJS.org site"
Write-Host " 3)  Download the latest version of Node Package Manager NPM v1.4.9 from the NodeJS.org site"
Write-Host " 4)  Install VSCode (silently)"
Write-Host " 5)  Create a user profile Node folder and setup Node and NPM there"
Write-Host " 6)  Update the local PATH statement to point to Node/NPM"
Write-Host " 7)  Update NPM to the latest"
Write-Host " 8)  Install the Yeoman generator via Node Package Manager (NPM)"
Write-Host " 9)  Install the http-server module via Node Package Manager (NPM)"
Write-Host "10)  Install the Office Yeoman Generator for creating Microsoft OfficeJS projects."
Write-Host ""
Read-Host -Prompt "Press <enter> if you wish to proceed"
#################
# Download VSCode
Write-Host "Downloading VSCode..."
Invoke-WebRequest -Uri $vscode_url -OutFile $vscode_output
#################
# Download Node
Write-Host "Downloading Node..."
Invoke-WebRequest -Uri $node_url -OutFile $node_output
#################
# Download NPM
Write-Host "Downloading NPM v1.4.9..."
Invoke-WebRequest -Uri $npm_url -OutFile $npm_output
#################
# Install VSCode
Write-Host "Installing VS Code (user only) - silently..."
Start-Process -FilePath $vscode_output -ArgumentList ('/VERYSILENT', '/MERGETASKS=!runcode') -Wait
#################
# Install Node
Write-Host "Setting up Node/NPM (user only)..."
# copy node.exe
md $node_install_path
copy $node_output $node_install_path
# extract npm.zip
Write-Host "Updating the PATH statement..."
[io.compression.zipfile]::ExtractToDirectory($npm_output, $node_install_path)
# setup the PATH for the user (permanent)
[Environment]::SetEnvironmentVariable(
    "Path",
    [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::User) + ";" + $node_install_path,
    [EnvironmentVariableTarget]::User)
# setup the PATH for this session
$env:Path += ";" + $node_install_path     
#################
# Updating NPM to latest
Write-Host "Updating NPM to latest..."
cd $node_install_path
Start-Process -FilePath npm -ArgumentList ('install', '-g', 'npm') -WindowStyle Hidden -Wait
#################
# Install Yeoman 
Write-Host "Installing Yeoman and the Office generator..."
Start-Process -FilePath npm -ArgumentList('install', '-g', 'yo', 'generator-office') -WindowStyle Hidden -Wait
#################
# Install http-server
Write-Host "Instaling http-server..."
Start-Process -FilePath npm -ArgumentList('install', '-g', 'http-server') -WindowStyle Hidden -Wait
#################
# Clean up
Write-Host "Cleaning up..."
del $vscode_output
del $node_output
del $npm_Output
$yo_file = $node_install_path + "\yo.ps1"
ren $yo_file "yo.ps1.old" #breaks yo office if we do not delete it
#################
# Complete
Write-Host "The installation is complete."
Read-Host -Prompt "Press <enter> to exit"