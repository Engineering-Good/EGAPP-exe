#QC Script
"`n"
Write-Host "Welcome to Automated Script for QC Testing" -ForegroundColor yellow
"`n"
Write-Host "Written by Hong Liang" -ForegroundColor yellow
Write-Host "Last updated 08/08/20 by Hong Liang" -ForegroundColor yellow
"`n"


Write-Host "Script file is in: $PSScriptRoot"
Write-Host "Current directory: $(Get-Location)"

# Starting the programs process
$counter = 0



Write-Host "Starting SW Test 1/5 - Chrome..."

$testChromeLocation = Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
$testChromeLocationAlt = Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe"

If ($testChromeLocation -eq $true) {
    Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    Write-Host "SW Test Passed. Chrome started." -ForegroundColor green
    $counter++
        }
elseif ($testChromeLocationAlt -eq $true) {
    Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe"
    Write-Host "SW Test Passed. Chrome started." -ForegroundColor green
    $counter++
        }
else {
     Write-Host "Chrome is not installed or does not exist in the standard location." -ForegroundColor red}
"`n"

Start-Sleep 5

Write-Host "Starting SW Test 2/5 - Acrobat..."
$testProgram = Test-Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
If ($testProgram -eq $false) {
    Write-Host "Acrobat is not installed or does not exist in the standard location." -ForegroundColor red}
else {
    Start-Process "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    Start-Sleep 3
    Write-Host "SW Test Passed. Acrobat started." -ForegroundColor green
    $counter++}
"`n"



Write-Host "Starting SW Test 3/5 - LibreOffice..."
$testProgram = Test-Path "C:\Program Files\LibreOffice\program\soffice.exe"
If ($testProgram -eq $false) {
    Write-Host "LibreOffice is not installed or does not exist in the standard location." -ForegroundColor red}
else {
    Start-Process "C:\Program Files\LibreOffice\program\soffice.exe"
    Start-Sleep 5
    Write-Host "SW Test Passed. LibreOffice started." -ForegroundColor green
    $counter++}
"`n"



$ping = Test-NetConnection
#write-host $ping.PingSucceeded
$pingResult = $ping.PingSucceeded
#write-host $pingResult

if ($pingResult -eq $false)
{
do {
    Write-Host "Please connect to WiFi manually before we can continue." -ForegroundColor green
    $ping = Test-NetConnection
    #write-host $ping.PingSucceeded
    $pingResult = $ping.PingSucceeded
   # write-host $pingResult
   pause
    }
    while ($pingResult -eq $false)
    }

Write-Host "Starting SW Test 4/5 - Joining Zoom Meeting..."
$testZoomLocation = Test-Path "C:\Program Files (x86)\Zoom\bin\Zoom.exe"
$AppDataPath = [Environment]::GetFolderPath('ApplicationData')
$ZoomPath = Join-Path -Path $AppDataPath -ChildPath "\Zoom\bin\Zoom.exe"
$testZoomLocationAlt = Test-Path $ZoomPath

If ($testZoomLocation -eq $true) {

        Start-Process "zoommtg://zoom.us/join?confno=3966517262&pwd=2020&zc=0&uname=User"
        Start-Sleep 5
        Write-Host "SW Test Passed. Zoom started." -ForegroundColor green
        $counter++
        }
elseif ($testZoomLocationAlt -eq $true) {

        Start-Process "zoommtg://zoom.us/join?confno=3966517262&pwd=2020&zc=0&uname=User"
        Start-Sleep 5
        Write-Host "SW Test Passed. Zoom started." -ForegroundColor green
        $counter++
        }
else {
        Write-Host "Zoom is not installed or does not exist in the standard location." -ForegroundColor red}
   
"`n"



Write-Host "Starting SW Test 5/5 Copying Joseph Schooling Video to Desktop..."
Set-Location $PSScriptRoot
$DesktopPath = [Environment]::GetFolderPath("Desktop")
Copy-Item "Team Singapore Surprise.mp4" -Destination $DesktopPath
$VidPath = Join-Path -Path $DesktopPath -ChildPath "\Team Singapore Surprise.mp4"
$testProgram = Test-Path $VidPath
If ($testProgram -eq $false) {
    Write-Host "Joseph Schooling Video is not found on Desktop. Please copy it manually." -ForegroundColor red}
else {
    Start-Process $VidPath
    Start-Sleep 5
    Write-Host "SW Test Passed. Joseph Schooling Video played." -ForegroundColor green
    $counter++}
"`n"

Installs\keyboardtestutility.exe

"`n"
If ($counter -eq 5 -and $winActivationStatus.LicenseStatus -eq 1) {
    Write-Host "QC Software passed." -ForegroundColor green}
else {
     Write-Host "QC Software Done. Please review the log to find if anything failed." -Foreground red}

"`n"

Pause

