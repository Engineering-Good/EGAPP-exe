Function Set-WallPaper($Value)
 {
    Set-ItemProperty -path 'HKCU:\Control Panel\Desktop\' -name wallpaper -value $value
    rundll32.exe user32.dll, UpdatePerUserSystemParameters
 }

$path = Join-Path -Path $PSScriptRoot -ChildPath "1.bmp"

Set-WallPaper -value $path