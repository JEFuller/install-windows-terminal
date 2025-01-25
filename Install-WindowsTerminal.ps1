$Package = Get-AppxPackage -Name Microsoft.WindowsTerminal
if ($null -ne $Package) {
    Write-Host "Microsoft.WindowsTerminal is already installed."
    Pause
    exit
}

Push-Location $Env:TEMP

Write-Host "Downloading Dependencies..."
### Download Microsoft.VCLibs.x64.14.00.Desktop.appx
Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -OutFile Microsoft.VCLibs.x64.14.00.Desktop.appx

### Download Microsoft.UI.Xaml.2.8.x64.appx
#Invoke-WebRequest -Uri https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.5/Microsoft.UI.Xaml.2.8.x64.appx -OutFile Microsoft.UI.Xaml.2.8.x64.appx
Invoke-WebRequest -Uri https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx -OutFile Microsoft.UI.Xaml.2.8.x64.appx

### Download Windows Terminal
# Windows Terminal Latest 1.21.3231.0
#Invoke-WebRequest -Uri https://github.com/microsoft/terminal/releases/download/v1.21.3231.0/Microsoft.WindowsTerminal_1.21.3231.0_8wekyb3d8bbwe.msixbundle -OutFile Microsoft.WindowsTerminal_1.21.3231.0_8wekyb3d8bbwe.msixbundle
# Windows Terminal Preview Pre-Release 1.22.3232.0
Invoke-WebRequest -Uri https://github.com/microsoft/terminal/releases/download/v1.22.3232.0/Microsoft.WindowsTerminalPreview_1.22.3232.0_8wekyb3d8bbwe.msixbundle -OutFile Microsoft.WindowsTerminalPreview_1.22.3232.0_8wekyb3d8bbwe.msixbundle

Write-Host "Installing Dependencies..."
Add-AppxPackage .\Microsoft.VCLibs.x64.14.00.Desktop.appx
Add-AppxPackage .\Microsoft.UI.Xaml.2.8.x64.appx
Add-AppxPackage .\Microsoft.WindowsTerminalPreview_1.22.3232.0_8wekyb3d8bbwe.msixbundle

Write-Host "Installed Microsoft.WindowsTerminal."
Pop-Location
Pause
