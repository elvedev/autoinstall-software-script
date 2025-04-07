Import-Module BitsTransfer;

function Set-ConsoleSize {
    [CmdletBinding()]
    Param(
         [Parameter(Mandatory=$False,Position=0)]
         [int]$Height = 40,
         [Parameter(Mandatory=$False,Position=1)]
         [int]$Width = 120
         )
    $console = $host.ui.rawui
    $ConBuffer  = $console.BufferSize
    $ConSize = $console.WindowSize
    
    $currWidth = $ConSize.Width
    $currHeight = $ConSize.Height
    
    # if height is too large, set to max allowed size
    if ($Height -gt $host.UI.RawUI.MaxPhysicalWindowSize.Height) {
        $Height = $host.UI.RawUI.MaxPhysicalWindowSize.Height
    }
    
    # if width is too large, set to max allowed size
    if ($Width -gt $host.UI.RawUI.MaxPhysicalWindowSize.Width) {
        $Width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
    }
    
    # If the Buffer is wider than the new console setting, first reduce the width
    If ($ConBuffer.Width -gt $Width ) {
       $currWidth = $Width
    }
    # If the Buffer is higher than the new console setting, first reduce the height
    If ($ConBuffer.Height -gt $Height ) {
        $currHeight = $Height
    }
    # initial resizing if needed
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($currWidth,$currHeight)
    
    # Set the Buffer
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Width, 2000)
    
    # Now set the WindowSize
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($Width,$Height)

}
function Install-App {
    param (
        $Link,
        $FileName
    )

    $LocalTempDir = $env:TEMP;
    Write-Host "Start downloading $FileName...";
    Start-BitsTransfer -Source $Link -Destination "$LocalTempDir/$FileName";
    if ($FileName -like "*.msi"){
        $ExitCode = (Start-Process msiexec.exe -Wait -Verb RunAs -ArgumentList "/q /i $LocalTempDir\$FileName" -PassThru).ExitCode;
    }else{
        Start-Process -Wait -FilePath "$LocalTempDir\$FileName";
        $ExitCode = 0;
    }

    if (0 -eq $ExitCode) {
        & Write-Host "$FileName is downloaded :)";
    }
    else{
        & Write-Host "Error (ExitCode): $ExitCode";
    }

    & Remove-Item -LiteralPath "$LocalTempDir\$FileName";
}

function Check-GoogleChrome {
    $chrome = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" |
    Where-Object { $_.DisplayName -like "*Google Chrome*" }

    if ($chrome) {
        return "Google Chrome is already installed."
    } else {
        return "Google Chrome is not installed."
    }
}

function Get-MenuSelection {
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$MenuItems,
        [Parameter(Mandatory = $true)]
        [String]$MenuPrompt
    )
    cls
    # store initial cursor position
    $cursorPosition = $host.UI.RawUI.CursorPosition
    $pos = 0 # current item selection
    #==============
    # 1. Draw menu
    #==============
    function Write-Menu {
        param (
            [int]$selectedItemIndex
        )
        # reset the cursor position
        $Host.UI.RawUI.CursorPosition = $cursorPosition
        # Padding the menu prompt to center it
        $prompt = $MenuPrompt
        $maxLineLength = ($MenuItems | Measure-Object -Property Length -Maximum).Maximum + 4
        while ($prompt.Length -lt $maxLineLength + 4) {
            $prompt = "                         $prompt                 "
        }
        Write-Host $prompt -ForegroundColor Green
        # Write the menu lines
        Write-Host ""
        Write-Host "$isInstalledGoogleChrome"
        Write-Host ""
        for ($i = 0; $i -lt $MenuItems.Count; $i++) {
            $line = "$($MenuItems[$i])"
            if ($selectedItemIndex -eq $i) {
                Write-Host $line -ForegroundColor Blue
            }
            else {
                Write-Host $line
            }
        }
    }

    Write-Menu -selectedItemIndex $pos
    $key = $null
    while ($key -ne 13) {
        #============================
        # 2. Read the keyboard input
        #============================
        $press = $host.ui.rawui.readkey("NoEcho,IncludeKeyDown")
        $key = $press.virtualkeycode
        if ($key -eq 38) {
            $pos--
        }
        if ($key -eq 40) {
            $pos++
        }
        #handle out of bound selection cases
        if ($pos -lt 0) { $pos = 0 }
        if ($pos -eq $MenuItems.count) { $pos = $MenuItems.count - 1 }

        #==============
        # 1. Draw menu
        #==============
        Write-Menu -selectedItemIndex $pos
    }

    return $MenuItems[$pos]
}

Set-ConsoleSize -Height 30 -Width 60
$ChromeLink = 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi';
$OfficeLink = 'https://go.microsoft.com/fwlink/?linkid=2264705&clcid=0x409&culture=en-us&country=us';
$JottaCloudLink = 'https://sw.jotta.cloud/desktop/download/windows/jottacloud/release'
$ChromeFileName = 'ChromeInstaller.msi';
$OfficeFileName = 'OfficeInstaller.exe';
$JottaCloudFileName = 'JottaCloud.exe';
$isInstalledGoogleChrome = Check-GoogleChrome;

$values = 'Basic (McAfee + JottaCloud)', 'Full (Basic + MS Office)'
$selection = Get-MenuSelection -MenuItems $values -MenuPrompt 'AutoInstall'
cls
if ($selection -eq 'Basic (McAfee + JottaCloud)'){
    if($isInstalledGoogleChrome -eq 'Google Chrome is not installed.'){
        Install-App $ChromeLink $ChromeFileName;
    }
    Install-App $JottaCloudLink $JottaCloudFileName;
}
elseif ($selection -eq 'Full (Basic + MS Office)') {
    if($isInstalledGoogleChrome -eq 'Google Chrome is not installed.'){
        Install-App $ChromeLink $ChromeFileName;
    }
    Install-App $JottaCloudLink $JottaCloudFileName;
    Install-App $OfficeLink $OfficeFileName;
}
else{

}