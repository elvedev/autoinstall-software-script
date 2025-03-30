Import-Module BitsTransfer;

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


$ChromeLink = 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi';
$OfficeLink = 'https://go.microsoft.com/fwlink/?linkid=2264705&clcid=0x409&culture=en-us&country=us';
$ChromeFileName = 'ChromeInstaller.msi';
$OfficeFileName = 'OfficeInstaller.exe';

Install-App $ChromeLink $ChromeFileName;
Install-App $OfficeLink $OfficeFileName;