#System Info
##############################################################################################################################
Write-Host "System Info" -ForegroundColor Yellow
"`n"
$processorInfo = Get-WmiObject -Class Win32_Processor | Select-Object -Property Name
Write-Host "CPU Name: "$processorInfo.Name""
$ramInfo = (Get-CimInstance -ClassName Win32_ComputerSystem).totalphysicalmemory 
$ramInfo = ([Math]::Round(($ramInfo)/1GB,0))
Write-Host "Total Installed RAM: $ramInfo GB"


$systemInfo = Get-CimInstance Win32_ComputerSystem | Select-Object -Property Name, Manufacturer, Model
$biosInfo = Get-CimInstance Win32_bios | Select-Object -Property SerialNumber
$windowsInfo = Get-CimInstance -ClassName win32_operatingsystem | Select-Object -Property OSArchitecture, Caption
Write-Host "Laptop Brand: "$systemInfo.Manufacturer""
Write-Host "Laptop Model: "$systemInfo.Model""
Write-Host "Laptop S/N: "$biosInfo.SerialNumber""
Write-Host "Windows Version: "$windowsInfo.Caption""
Write-Host "Windows Type: "$windowsInfo.OSArchitecture""
Start-Sleep 5

Function Get-DiskInfo {
$disk = Get-WMIObject Win32_Logicaldisk -ComputerName $computer |
            Select-Object  @{Name="Computer";Expression={$computer}}, 
                DeviceID,
                @{Name="Size in GB";Expression={$_.Size/(1000*1000*1000) -as [int]}}
            
        #Write-Host $Computer -ForegroundColor Magenta
        $disk
}
Function Get-VRamInfo {
    $vram = Get-WmiObject win32_videocontroller -ComputerName $computer | 
                Select-Object @{Name="Computer";Expression={$computer}},
                    @{Name="Video RAM in MB";Expression={$_.adapterram / (1000*1000) -as [int]}},
                    @{Name="Size in GB";Expression={$_.adapterram/(1000*1000*1000) -as [int]}},
                    Name
            #Write-Host $computer -ForegroundColor Cyan
            $vram
    }
    
    $computer = '.'
    
    Get-VRamInfo | Format-Table
    Get-DiskInfo | Format-Table

#Stop-Service WinRM

###############################################################################################################################
# Display Licences
Write-Host Checking for existing licenses -ForegroundColor Yellow
$licList = Get-CimInstance -Class SoftwareLicensingProduct |
    where {$_.name -match ‘windows’ -AND $_.LicenseFamily -AND $_.LicenseStatus -ne 0} |
        Select-Object -Property Name, `
                    @{Label= “License Status”; Expression={switch (foreach {$_.LicenseStatus}) `
                    { 0 {“Unlicensed”} `
                    1 {“Licensed”} `
                    2 {“Out-Of-Box Grace Period”} `
                    3 {“Out-Of-Tolerance Grace Period”} `
                    4 {“Non-Genuine Grace Period”} `
                    5 {“Notification”} `
                    6 {“Extended Grace Period”} `
                    } } }
    if ($licList -eq $null) {
        Write-Host No license found -ForegroundColor Red
    } else {
        Write-Host The following licenses were found: -ForegroundColor Green
        if ([System.Environment]::Is64BitOperatingSystem) {Write-Host 64-bit Windows detected} else {Write-Host 32-bit Windows detected}
    }
    $licList | Format-List



###############################################################################################################################
# Check Office

Write-Host Checking Office and Windows Activation Status -ForegroundColor Yellow

$checkOfficex64 = Test-Path "C:\Program Files\Microsoft Office\Office16"
$checkOfficex86 = Test-Path "C:\Program Files (x86)\Microsoft Office\Office16"


If ($checkOfficex64 -eq $true) {
    Set-Location "C:\Program Files\Microsoft Office\Office16"}
elseif ($checkOfficex86 -eq $true) {
    Set-Location "C:\Program Files (x86)\Microsoft Office\Office16"}
else {
    Write-Host "Office is not installed."}

$officeActivation = cscript ospp.vbs /dstatus
If ($officeActivation -match "---LICENSED---"){
    Write-host "Office is activated." -ForegroundColor green}
else {
    Write-Host "Office is not activated." -Foreground red}
"`n"


$winActivationStatus = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | where { $_.PartialProductKey } | select Description, LicenseStatus
if ($winActivationStatus.LicenseStatus -eq 1) {
Write-Host "Windows is activated." -ForegroundColor Green}
else {
Write-Host "WINDOWS IS NOT ACTIVATED!!!!!!!!!!!!!!!! WIN KEY NEEDED" -ForegroundColor Red}
###############################################################################################################################
"`n"

Write-Host Running battery checks -ForegroundColor Yellow

Try{          
    $BattAssembly = [Windows.Devices.Power.Battery,Windows.Devices.Power.Battery,ContentType=WindowsRuntime] 
}
Catch
{
    Write-Error "Unable to load the Windows.Devices.Power.Battery class"
}
        
Try{
    $Report = [Windows.Devices.Power.Battery]::AggregateBattery.GetReport() 
}
Catch{
    Write-Error "Unable to retrieve Battery Report information"
    Break
}

If ($Report.Status -ne "NotPresent")
{
        
    if ($Report.DesignCapacityInMilliwattHours -ne 0) {
        $batteryHealth = $Report.FullChargeCapacityInMilliwattHours / $Report.DesignCapacityInMilliwattHours
    } else {
        $batteryHealth = 0
    }

    $data = @{
        Status = $Report.Status
        "Battery Health" = ($batteryHealth * 100).toString('F2') + "%"
        "Charge Rate (%/min)" = ($Report.ChargeRateInMilliwatts / $Report.FullChargeCapacityInMilliwattHours / 60 * 100).toString('F2') + "%"
    }
        
    New-Object PSObject -Property $data | Format-List
        
    if ($batteryHealth -eq 0) {
        Write-Host Battery is dead -ForegroundColor DarkRed
    } elseif ($batteryHealth -lt 0.2) {
        Write-Host Battery is very weak -ForegroundColor Red
    } elseif ($batteryHealth -lt 0.6) {
        Write-Host Battery is weak -ForegroundColor Yellow
    } else {
        Write-Host Battery health is decent -ForegroundColor Green
    }
       
}
Else
{
    Write-Host "Unable to detect working battery, please check." -ForegroundColor Red
}

Write-Host "`n"

#########################################################################################################################
Write-Host Checking Harddrive Health -ForegroundColor Yellow
"`n"
Get-Disk | Get-StorageReliabilityCounter | Format-List -Property *

$attention = 'Attention! Check Above List. If there is any line is red in color, Please REPLACE the Hard Disk which gives ERROR'
Write-Host $attention -ForegroundColor Red 


Start-Sleep -s 15
pause
