# Verify whether 7-Zip is available.
$7zip_path = "C:\Program Files\7-Zip\7z.exe"
if ((Test-Path -Path $7zip_path) -ne $true) {
    return
}
Set-Alias 7zip $7zip_path



# Close Thunderbird, if it's open.
if ((Get-Process -ErrorAction SilentlyContinue -Name "Thunderbird") -ne $null) {
     Stop-Process -Name "Thunderbird"
}



# Backup Thunderbird VIA 7-Zip
Set-Location -Path "$env:APPDATA"

$desktop_path = [Environment]::GetFolderPath("Desktop")
7zip a "$desktop_path\Thunderbird.7z" ".\Thunderbird"



# Restart Thunderbird
Start-Process -FilePath "C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"
