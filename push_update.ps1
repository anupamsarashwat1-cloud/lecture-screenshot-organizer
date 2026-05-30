# push_update.ps1
# Stages the new assets and pushes them directly to GitHub.

try {
    $ghExePath = Get-ChildItem -Path "$PSScriptRoot\tools\gh" -Filter "gh.exe" -Recurse | Select-Object -First 1 -ExpandProperty DirectoryName
    $gitCmdPath = "$PSScriptRoot\tools\MinGit\cmd"
    $env:PATH = "$gitCmdPath;$ghExePath;$env:PATH"
    
    Write-Host "Staging files..."
    git add .
    
    Write-Host "Committing changes..."
    git commit -m "Add terminal demo screenshot to README"
    
    Write-Host "Pushing updates to GitHub..."
    git push origin main
    
    Write-Host "SUCCESS! Updates pushed to GitHub."
} catch {
    Write-Error $_
}
