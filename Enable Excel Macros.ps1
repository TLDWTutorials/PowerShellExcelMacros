#Define registry path for Excel Trust Center Settings
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"

#Set VBAWarnings key (1 = enable all, 4 = disable w/o notification)
Set-ItemProperty -Path $regPath -Name "VBAWarnings" -Value 4

#Optional: Close and restart Excel to apply the changes
Stop-Process -Name EXCEL -Force
Start-Process excel

#Create a message to acknowledge user input
Write-Host "Macro security settings have been updated. Press Enter to close this windows. Don't forget to change your wallpaper!"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

