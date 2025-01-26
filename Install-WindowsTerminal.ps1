<#
.SYNOPSIS
This script aims to install the latest version of Windows Terminal including its Preview version. 
Windows Terminal can be installed for the current user or for all users on the system.
        
.NOTES
Name: Install-Windows Terminal
Author: Dave Tapley (JE Fuller) @davetapley
Contribution: @ahpooch
Version: 1.0
DateModified: 01.26.2025
    
.EXAMPLE
# Example 1: Installing for current user.
& .\Install-WindowsTerminal

# Example 2: Installing for all users.
& .\Install-WindowsTerminal -Scope AllUsers

# Example 3: Installing Preview version for current user.
& .\Install-WindowsTerminal -Preview

# Example 4: Installing Preview version for all users.
& .\Install-WindowsTerminal -Scope AllUsers -Preview
    
.LINK
https://github.com/JEFuller/install-windows-terminal    
https://github.com/ahpooch/install-windows-terminal
#>
    
[CmdletBinding()]
param(
    [Parameter(
        Mandatory = $false
    )]
    [ValidateSet("CurrentUser", "AllUsers")]
    [string]$Scope = "CurrentUser",
    [Switch]$Preview
)

if ($null -ne $Preview) {
    $PackageName = "Microsoft.WindowsTerminalPreview"
    $Package = Get-AppxPackage -Name $PackageName
}
else {
    $PackageName = "Microsoft.WindowsTerminal"
    $Package = Get-AppxPackage -Name $PackageName
}
if ($null -ne $Package) {
    Write-Host "$PackageName is already installed."
    exit
}

### Checking elevation if installing for all users.
Function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
if ($Scope -eq "AllUsers" -and !(Test-IsAdmin)) {
    Throw "Installation with -Scope AllUsers requires elevation!"
    exit
}

### Checking OS version and throwing an error if it is not compatible.
[version]$RequiredVersion = [version]"10.0.19041.0"
[version]$CurrentOSVersion = [version](Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Version)
if ($CurrentOSVersion -lt $RequiredVersion) {
    Throw "Cannot install package because this package is not compatible with the device. See https://github.com/microsoft/terminal/issues/2177.`nThe package requires OS version 10.0.19041.0 or higher (Windows Server 2022 at least). The device is currently running OS version $CurrentOSVersion"
    exit
}

### Setting variables
# Microsoft.VCLibs.x64.14.00.Desktop.appx
$vclibs_x64_filename = "Microsoft.VCLibs.x64.14.00.Desktop.appx"
$vclibs_x64_url = "https://aka.ms/$vclibs_x64_filename"
# Microsoft.UI.Xaml.2.8.x64.appx
$uixaml_x64_filename = "Microsoft.UI.Xaml.2.8.x64.appx"
$uixaml_x64_url = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/$uixaml_x64_filename"
# Windows Terminal
$windows_terminal_filename = "Microsoft.WindowsTerminal_1.21.3231.0_8wekyb3d8bbwe.msixbundle"
$windows_terminal_url = "https://github.com/microsoft/terminal/releases/download/v1.21.3231.0/$windows_terminal_filename"
# Windows Terminal Preview
$windows_terminal_preview_filename = "Microsoft.WindowsTerminalPreview_1.22.3232.0_8wekyb3d8bbwe.msixbundle"
$windows_terminal_preview_url = "https://github.com/microsoft/terminal/releases/download/v1.22.3232.0/$windows_terminal_preview_filename"

### Check if files present in current folder
$vclibs_x64_present = $false
$uixaml_x64_present = $false
$windows_terminal_present = $false
$windows_terminal_preview_present = $false

if (Test-Path -Path .\$vclibs_x64_filename) {
    $vclibs_x64_present = $true
}
else {
    # Download Microsoft.VCLibs.x64.14.00.Desktop.appx
    Invoke-WebRequest -Uri $vclibs_x64_url -OutFile "$env:temp\$vclibs_x64_filename"
}

if (Test-Path -Path .\$uixaml_x64_filename) {
    $uixaml_x64_present = $true
}
else {
    # Download Microsoft.UI.Xaml.2.8.x64.appx
    Invoke-WebRequest -Uri $uixaml_x64_url -OutFile "$env:temp\$uixaml_x64_filename"
}

if ($null -ne $Preview) {
    if (Test-Path -Path .\$windows_terminal_preview_filename) {
        $windows_terminal_preview_present = $true
    }
    else {
        # Download Windows Terminal Preview
        Invoke-WebRequest -Uri $windows_terminal_preview_url -OutFile "$env:temp\$windows_terminal_preview_filename"
    }
    
}
else {
    if (Test-Path -Path .\$windows_terminal_filename) {
        $windows_terminal_present = $true
    }
    else {
        # Download Windows Terminal
        Invoke-WebRequest -Uri $windows_terminal_url -OutFile "$env:temp\$windows_terminal_filename"
    }
}

### Installing files depending on conditions
# Install for CurrentUser
if ($Scope -eq "CurrentUser") {
    if ($vclibs_x64_present) {
        Add-AppxPackage .\$vclibs_x64_filename
    }
    else {
        Add-AppxPackage "$env:temp\$vclibs_x64_filename"
    }
    
    if ($uixaml_x64_present) {
        Add-AppxPackage .\$uixaml_x64_filename
    }
    else {
        Add-AppxPackage "$env:temp\$uixaml_x64_filename"
    }

    if ($null -ne $Preview) {
        if ($windows_terminal_preview_present) {
            Add-AppxPackage .\$windows_terminal_preview_filename
        }
        else {
            Add-AppxPackage "$env:temp\$windows_terminal_preview_filename"
        }
    }
    else {
        if ($windows_terminal_present) {
            Add-AppxPackage .\$windows_terminal_filename
        }
        else {
            Add-AppxPackage "$env:temp\$windows_terminal_filename"
        }
    }
}
# Install for AllUsers
else {
    if ($vclibs_x64_present) {
        Add-ProvisionedAppPackage -Online -PackagePath ".\$vclibs_x64_filename" -SkipLicense | Out-Null
    }
    else {
        Add-ProvisionedAppPackage -Online -PackagePath "$env:temp\$vclibs_x64_filename" -SkipLicense | Out-Null
    }

    if ($uixaml_x64_present) {
        Add-ProvisionedAppPackage -Online -PackagePath ".\$uixaml_x64_filename" -SkipLicense | Out-Null
    }
    else {
        Add-ProvisionedAppPackage -Online -PackagePath "$env:temp\$uixaml_x64_filename" -SkipLicense | Out-Null
    }

    if ($null -ne $Preview) {
        if ($windows_terminal_preview_present) { 
            Add-ProvisionedAppPackage -Online -PackagePath ".\$windows_terminal_preview_filename" -SkipLicense | Out-Null
        }
        else {
            Add-ProvisionedAppPackage -Online -PackagePath "$env:temp\$windows_terminal_preview_filename" -SkipLicense | Out-Null
        }
    }
    else {
        if ($windows_terminal_present) {
            Add-ProvisionedAppPackage -Online -PackagePath ".\$windows_terminal_filename" -SkipLicense | Out-Null
        }
        else {
            Add-ProvisionedAppPackage -Online -PackagePath "$env:temp\$windows_terminal_filename" -SkipLicense | Out-Null
        }
    }
}
Write-Host "Microsoft.WindowsTerminal installed successfull."

### Removing temp files
if ($vclibs_x64_present -eq $false) {
    Remove-Item -Path "$env:temp\$vclibs_x64_filename" -ErrorAction SilentlyContinue
}
if ($uixaml_x64_present -eq $false) {
    Remove-Item -Path "$env:temp\$uixaml_x64_filename" -ErrorAction SilentlyContinue
}
if ($null -ne $Preview) {
    if ($windows_terminal_preview_present -eq $false) {
        Remove-Item -Path "$env:temp\$windows_terminal_preview_filename" -ErrorAction SilentlyContinue
    }
}
else {
    if ($windows_terminal_present -eq $false) {
        Remove-Item -Path "$env:temp\$windows_terminal_filename" -ErrorAction SilentlyContinue
    }
}
Write-Host "Temporary files removed."


