# Create-Shortcut.ps1
# Creates a premium desktop shortcut with a camera icon and custom global hotkey.

try {
    $WshShell = New-Object -ComObject WScript.Shell
    $desktop = [System.Environment]::GetFolderPath('Desktop')
    $shortcutPath = Join-Path $desktop "OneNote Screenshot Organizer.lnk"
    
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = "c:\Users\anupa\OneDrive\Documents\LECTURE SCREENSHOT ORGANIZER AGENT\Start-ScreenshotOrganizer.bat"
    $Shortcut.WorkingDirectory = "c:\Users\anupa\OneDrive\Documents\LECTURE SCREENSHOT ORGANIZER AGENT"
    $Shortcut.Hotkey = "Ctrl+Alt+O"
    
    # 265 in shell32.dll is a beautiful retro camera/photo icon on Windows
    $Shortcut.IconLocation = "shell32.dll,265"
    $Shortcut.Save()
    
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "[SUCCESS] Desktop Shortcut Created Successfully!" -ForegroundColor Green
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "  Location: On your Desktop (named 'OneNote Screenshot Organizer')" -ForegroundColor Gray
    Write-Host "  Global Hotkey: Press [Ctrl + Alt + O] from anywhere to launch it!" -ForegroundColor Yellow
    Write-Host "==========================================================================" -ForegroundColor Cyan
} catch {
    Write-Error $_
}
