#Requires -RunAsAdministrator

Write-Host Hello there! This is the PCMark10 Benchmark script! Extracted and Edited by 9Guest from [ThisQCandBenchmark.ps1] -ForegroundColor Yellow
Write-Host Written by Hong Liang
Write-Host Last updated 03/10/2020

$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "SwitchUACLevel.psm1"
Import-Module $modulePath

Write-Host Supressing UAC notifications -ForegroundColor Yellow
Set-UACLevel 0 | Out-Null

Write-Host Trying to Stop Some Applications from Running ==> Red console lines may appear due to inactive application process -ForegroundColor Yellow
Stop-Process -Name "chrome" -Force
Stop-Process -Name "AcroRd32" -Force
Stop-Process -Name "soffice" -Force
Stop-Process -Name "zoom" -Force

#kill all known programs for testing

Write-Host "Setting Time Zone to Singapore Time" -ForegroundColor Yellow
Set-TimeZone -Id "Singapore Standard Time" -PassThru
net start w32time
W32tm /resync /force
"`n"

$biosInfo = Get-CimInstance Win32_bios | Select-Object -Property SerialNumber
$processorInfo = Get-WmiObject -Class Win32_Processor | Select-Object -Property Name
$ramInfo = (Get-CimInstance -ClassName Win32_ComputerSystem).totalphysicalmemory 
$ramInfo = ([Math]::Round(($ramInfo)/1GB,0))

#check if benchmark is needed. If processor is 4th gen and above, no benchmark needed.

if ($processorInfo -match 'i3-[4567891]' -or $processorInfo -match 'i5-[4567891]' -or $processorInfo -match 'i7-[4567891]'){
    write-host "CPU is" $processorInfo.Name  ", which is above acceptable range. Benchmark not needed." -ForegroundColor green
    Start-Sleep 5

    Write-Host Remember to Enter [Above Acceptable Range] in the form to upload the data to google sheet -ForegroundColor cyan
    }
else 
    {
write-host "CPU is" $processorInfo.Name  ", which is below acceptable range. Benchmark is needed." -ForegroundColor red
Start-Sleep 5
    
#Start-process of PCMark

Write-Host "`nInstalling PCMark 10..."

$pcMark10ResultsPathPDF = $biosInfo.SerialNumber + "_" + $processorInfo.Name + "_" + $ramInfo + "GB_bmarkExpress.pdf"
$pcMark10ResultsPathXML = $biosInfo.SerialNumber + "_" + $processorInfo.Name + "_" + $ramInfo + "GB_bmarkExpress.xml"
$pcMark10ResultsPathLog = $biosInfo.SerialNumber

#Get the installation file of PC Mark 10

$path = Join-Path -Path $PSScriptRoot -ChildPath "Installs\PCMark10\pcmark10-setup.exe"

cmd /c start /wait $path /quiet /silent
& 'C:\Program Files\UL\PCMark 10\PCMark10Cmd.exe'`
 --register PCM10-TPRO-20210801-227PQ-FD6M2-DUJNH-VM5V7 `
 --definition=pcm10_express.pcmdef `
 --out=$PSScriptRoot\PCMark10ResultsLog\$pcMark10ResultsPathLog `
 --export-pdf=$PSScriptRoot\PCMark10Results\$pcMark10ResultsPathPDF `
 --export-xml=$PSScriptRoot\PCMark10Results\$pcMark10ResultsPathXML `
 --systeminfo on `
 --systeminfomonitor on `
 --online on

"`n"

Start-Process $PSScriptRoot\PCMark10Results\$pcMark10ResultsPathPDF

Write-Host Uploading Benchmark results to Google Drive...

# Set the Google Auth parameters. Fill in your RefreshToken, ClientID, and ClientSecret
$params = @{
    Uri = 'https://accounts.google.com/o/oauth2/token'
    Body = @(
        "refresh_token=1//0fONxwhKN3xj7CgYIARAAGA8SNwF-L9IrOd2-JNVNjaxuIL9GOex9VHUnZ7HB5_v0DbnbKYBl-GOsjldWLmpny-LLZgavgbuhYTk", # Replace $RefreshToken with your refresh token
        "client_id=418256764249-lv1r71ok8q62m0dttdqff8lmnvht4kkg.apps.googleusercontent.com",         # Replace $ClientID with your client ID
        "client_secret=wJZEpdwZiXWaQDqzfBZwvQim", # Replace $ClientSecret with your client secret
        "grant_type=refresh_token"
    ) -join '&'
    Method = 'Post'
    ContentType = 'application/x-www-form-urlencoded'
}
$accessToken = (Invoke-RestMethod @params).access_token

# Change this to the file you want to upload
$SourceFile = "$PSScriptRoot\PCMark10Results\$pcMark10ResultsPathPDF"

# Get the source file contents and details, encode in base64
$sourceItem = Get-Item $sourceFile
$sourceBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($sourceItem.FullName))
$sourceMime = [System.Web.MimeMapping]::GetMimeMapping($sourceItem.FullName)

# If uploading to a Team Drive, set this to 'true'
$supportsTeamDrives = 'true'

# Set the file metadata
$uploadMetadata = @{
    originalFilename = $sourceItem.Name
    name = $sourceItem.Name
    description = $sourceItem.VersionInfo.FileDescription
    parents = @('1Sds2y-5OT_V1-rR0piyoLfqxlQ7qHIG_') # Include to upload to a specific folder
    teamDriveId = ‘1NhFWy8A_c6yWUWU7PA46wcdUIsHRQaVb’            # Include to upload to a specific teamdrive
}

# Set the upload body
$uploadBody = @"
--boundary
Content-Type: application/json; charset=UTF-8

$($uploadMetadata | ConvertTo-Json)

--boundary
Content-Transfer-Encoding: base64
Content-Type: $sourceMime

$sourceBase64
--boundary--
"@

# Set the upload headers
$uploadHeaders = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type" = 'multipart/related; boundary=boundary'
    "Content-Length" = $uploadBody.Length
}

# Perform the upload
$response = Invoke-RestMethod -Uri "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsTeamDrives=$supportsTeamDrives" -Method Post -Headers $uploadHeaders -Body $uploadBody

Write-Host "`nBenchmark results uploaded Google Drive TAC folder."

Write-Host "`nUninstalling PCMark 10..."

cmd /c start /wait $path /uninstall

#popd
}

# Revert UAC Settings to default
Write-Host Restoring UAC settings to Default -ForegroundColor Yellow
Set-UACLevel 2 | Out-Null

Read-Host -Prompt "Press Enter to exit"
Write-Host "Remember to save the computer's details for uploading data to Google Sheet" -ForegroundColor Cyan

PAUSE
PAUSE
