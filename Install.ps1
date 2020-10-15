#Requires -RunAsAdministrator

Write-Host Hello there! This is the QC software install script! (Extracted and Editted by 9Guest from InstallAndQC.ps1)-ForegroundColor Yellow
Write-Host Last updated by Hong Liang 03/10/2020


$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "SwitchUACLevel.psm1"
Import-Module $modulePath

Write-Host Supressing UAC notifications -ForegroundColor Yellow
Set-UACLevel 0 | Out-Null


#kill all known programs for testing
Stop-Process -Name "chrome" -Force
Stop-Process -Name "AcroRd32" -Force
Stop-Process -Name "soffice" -Force
Stop-Process -Name "zoom" -Force


Write-Host "Connecting to engineeringGood WiFi..."
Push-Location $PSScriptRoot
netsh wlan add profile filename="Wi-Fi-eG.xml"
Start-Sleep 5
Write-Host "Wifi profile added."

# Get list of applications to be installed
$filePaths = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath "MSI_list.txt")

# Define folder for logs
$logFolderPath = Join-Path $PSScriptRoot -ChildPath "logs"

foreach ($filePath in $filePaths) {
    $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $filePath
    if ([IO.File]::Exists($fullPath)) {
        
        $DataStamp = get-date -Format yyyyMMddTHHmmss
        $logFile = '{0}-{1}.log' -f $filePath,$DataStamp
        $logFilePath = Join-Path -Path $logFolderPath -ChildPath $logFile

        $extn = [IO.Path]::GetExtension($fullPath)
   
        if ($extn -eq ".msi") {
            New-Item -ItemType File -Force -Path $logFilePath | Out-Null
            $MSIArguments = @(
                "/i"
                ('"{0}"' -f $fullPath)
                "/qn"
                "/norestart"
                "/L*v"
                $logFilePath
            )
            Write-Host Installing: $fullPath
            $proc = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -PassThru -Wait #-NoNewWindow
            if ($proc.ExitCode -eq 0) {
                Write-Host Successfully Installed: $fullPath -ForegroundColor Green
            } else {
                Write-Host Error: $proc.ExitCode -ForegroundColor Red
            }
	
        } elseif ($extn -eq ".msp") {
            New-Item -ItemType File -Force -Path $logFilePath | Out-Null
            $MSIArguments = @(
                "/p"
                ('"{0}"' -f $fullPath)
                "/qn"
                "/norestart"
                "/L*v"
                $logFilePath
            )
            Write-Host Installing: $fullPath
            $proc = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -PassThru -Wait #-NoNewWindow
            if ($proc.ExitCode -eq 0) {
                Write-Host Successfully Installed: $fullPath -ForegroundColor Green
            } else {
                Write-Host Error: $proc.ExitCode -ForegroundColor Red
            }
        } else {
            Write-Host Error: Unidentified Extension -ForegroundColor Red
        }
    } else {
        Write-Host File $fullPath does not exist
    }
}


# Revert UAC Settings to default
Write-Host Restoring UAC settings to Default -ForegroundColor Yellow
Set-UACLevel 2 | Out-Null

Read-Host -Prompt "Press Enter to exit"
Write-Host "That's all for Installation. Please Remember to fill in the form to upload data of this PC to Google Sheet" -ForegroundColor Yellow
PAUSE
PAUSE

