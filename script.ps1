$LocalTempDir = $env:TEMP;
$testfolder = 'D:\testfolder';
$ChromeLink = 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi';
$OfficeLink = 'https://go.microsoft.com/fwlink/?linkid=2264705&clcid=0x409&culture=en-us&country=us';
$ChromeFileName = 'ChromeInstaller.msi';
$OfficeFileName = 'OfficeInstaller.exe';
Write-Host 'Start downloading Google Chrome...';

Import-Module BitsTransfer;
# Start-BitsTransfer -Source $ChromeLink -Destination "$testfolder/$ChromeFileName"; 
# Start-BitsTransfer -Source $OfficeLink -Destination "$testfolder/$OfficeFileName";
# $ExitCode = (Start-Process msiexec.exe -Wait -Verb RunAs -ArgumentList "/q /i $testfolder\$ChromeFileName" -PassThru).ExitCode;
# if (0 -eq $ExitCode) {
#     & Write-Host 'Google Chrome is downloaded :)';
# }else{
#     & Write-Host "Error (ExitCode): $ExitCode";
# }
& "$testfolder\$OfficeFileName" /help;

# or
# Invoke-WebRequest $url -OutFile $path_to_file

# $ChromeInstaller = "ChromeInstaller.exe"; 
# (new-object    System.Net.WebClient).DownloadFile('http://dl.google.com/chrome/install/375.126/chrome_installer.exe', "$LocalTempDir\$ChromeInstaller"); 
# & "$LocalTempDir\$ChromeInstaller" /silent /install;
# $Process2Monitor =  "ChromeInstaller"; 
# Do { 
#         $ProcessesFound = Get-Process | ?{$Process2Monitor -contains $_.Name} | Select-Object -ExpandProperty Name; 
#         If ($ProcessesFound) { 
#             "Still running: $($ProcessesFound -join ', ')" | Write-Host; 
#             Start-Sleep -Seconds 2 
#         } else {
#             Remove-Item "$LocalTempDir\$ChromeInstaller" -ErrorAction SilentlyContinue -Verbose 
#         } 
#     } 
# Until (!$ProcessesFound)