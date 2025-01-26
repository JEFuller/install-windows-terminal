# install-windows-terminal
Easily install the new Windows Terminal or Windows Terminal Preview on Windows Server 2022 and newer.

Inspired by: https://github.com/microsoft/terminal/discussions/13983#discussioncomment-7554301

## Current versions provided by scripts
Current `Windows Terminal` version provided by `Install-WindowsTerminal.ps1` script : 1.21.3231.0  
Current `Windows Terminal Preview` version provided by `Install-WindowsTerminalPreview.ps1` script : 1.22.3232.0  

## Usage
Install-WindowsTerminal can download all files by itself. Just use one of examples below.  
```Powershell
# Example 1: Installing for current user.
& .\Install-WindowsTerminal

# Example 2: Installing for all users.
& .\Install-WindowsTerminal -Scope AllUsers

# Example 3: Installing Preview version for current user.
& .\Install-WindowsTerminal -Preview

# Example 4: Installing Preview version for all users.
& .\Install-WindowsTerminal -Scope AllUsers -Preview
```

Or it can be offline installer. Just downoad files by yourself and place it in the same folder as Install-WindowsTerminal.ps1.  
### Download Microsoft.VCLibs.x64.14.00.Desktop.appx  
https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx  
### Download Microsoft.UI.Xaml.2.8.x64.appx  
https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx  
### Download Windows Terminal  
https://github.com/microsoft/terminal/releases/download/v1.21.3231.0/Microsoft.WindowsTerminal_1.21.3231.0_8wekyb3d8bbwe.msixbundle  
### Download Windows Terminal Preview  
https://github.com/microsoft/terminal/releases/download/v1.22.3232.0/Microsoft.WindowsTerminalPreview_1.22.3232.0_8wekyb3d8bbwe.msixbundle  
