# download_to_desktop.ps1
# Securely downloads the latest GitHub Desktop installer directly to the user's Desktop.

try {
    $ProgressPreference = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $desktop = [System.Environment]::GetFolderPath('Desktop')
    $outFile = Join-Path $desktop "GitHubDesktopSetup.exe"
    
    Write-Host "Downloading GitHub Desktop to your Desktop..."
    Invoke-WebRequest -Uri "https://central.github.com/deployments/desktop/desktop/latest/win32" -OutFile $outFile -TimeoutSec 180
    
    if (Test-Path $outFile) {
        Write-Host "Download complete! Setup saved to: $outFile"
    } else {
        Write-Host "Failed to save file to Desktop."
    }
} catch {
    Write-Error $_
}
