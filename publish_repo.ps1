# publish_repo.ps1
# Automates the downloading of MinGit and GitHub CLI, generates the authorization code,
# and pushes the repository to GitHub once the user authorizes it in their browser.

$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$toolsDir = Join-Path $PSScriptRoot "tools"
if (-not (Test-Path $toolsDir)) {
    New-Item -ItemType Directory -Path $toolsDir | Out-Null
}

$gitDir = Join-Path $toolsDir "MinGit"
$ghDir = Join-Path $toolsDir "gh"
$authFile = Join-Path $PSScriptRoot "auth_code.txt"

# 1. Download and extract MinGit (25MB portable Git)
if (-not (Test-Path $gitDir)) {
    Write-Host "Downloading MinGit..."
    $gitZip = Join-Path $toolsDir "mingit.zip"
    Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/download/v2.45.1.windows.1/MinGit-2.45.1-64-bit.zip" -OutFile $gitZip
    Write-Host "Extracting MinGit..."
    Expand-Archive -Path $gitZip -DestinationPath $gitDir -Force
    Remove-Item $gitZip -ErrorAction SilentlyContinue
}

# 2. Download and extract GitHub CLI (10MB portable gh)
if (-not (Test-Path $ghDir)) {
    Write-Host "Downloading GitHub CLI..."
    $ghZip = Join-Path $toolsDir "gh.zip"
    Invoke-WebRequest -Uri "https://github.com/cli/cli/releases/download/v2.49.0/gh_2.49.0_windows_amd64.zip" -OutFile $ghZip
    Write-Host "Extracting GitHub CLI..."
    Expand-Archive -Path $ghZip -DestinationPath $ghDir -Force
    Remove-Item $ghZip -ErrorAction SilentlyContinue
}

# 3. Add tools to current session PATH
$ghExePath = Get-ChildItem -Path $ghDir -Filter "gh.exe" -Recurse | Select-Object -First 1 -ExpandProperty DirectoryName
$gitCmdPath = Join-Path $gitDir "cmd"

$env:PATH = "$gitCmdPath;$ghExePath;$env:PATH"

Write-Host "Verifying tool versions..."
git --version
gh --version

# 4. Start GitHub CLI browser-based device login
Write-Host "Starting GitHub authorization..."
$authFile | Remove-Item -ErrorAction SilentlyContinue

$stdoutFile = Join-Path $PSScriptRoot "gh_stdout.txt"
$stderrFile = Join-Path $PSScriptRoot "gh_stderr.txt"
Remove-Item $stdoutFile, $stderrFile -ErrorAction SilentlyContinue

# Launch gh auth login using Start-Process with redirection (no blocking streams!)
$ghExe = Join-Path $ghExePath "gh.exe"
$process = Start-Process $ghExe -ArgumentList "auth login --hostname github.com --git-protocol https --web" -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru

# Monitor stdout and stderr files for the one-time code
$code = $null
$attempts = 0
while ($null -eq $code -and $attempts -lt 150) {
    $content = ""
    if (Test-Path $stdoutFile) {
        $content += Get-Content $stdoutFile -Raw
    }
    if (Test-Path $stderrFile) {
        $content += Get-Content $stderrFile -Raw
    }
    
    if ($content -match "one-time code:\s*([A-Z0-9]{4}-[A-Z0-9]{4})") {
        $code = $Matches[1]
        $authData = @"
URL: https://github.com/login/device
CODE: $code
"@
        $authData | Out-File -FilePath $authFile -Encoding utf8
        Write-Host "SAVED AUTH CODE: $code"
        break
    }
    Start-Sleep -Milliseconds 200
    $attempts++
}

if ($null -eq $code) {
    Write-Error "Failed to generate authorization code. Please check gh_stderr.txt"
    if (Test-Path $stderrFile) {
        Get-Content $stderrFile | Write-Error
    }
    exit
}

# Wait for process to fully exit (meaning user authorized in browser)
Write-Host "Waiting for user to authorize in browser..."
$process.WaitForExit()

Write-Host "GitHub authorization completed successfully!"

# 6. Initialize Git and Push repository
Write-Host "Initializing Git local repository..."
git init

# Configure user identity
git config --global --add safe.directory "*"
git config user.name "UPSC & VLSI Aspirant"
git config user.email "upsc.vlsi.notes@example.com"

# Stage and commit files
git add .
git commit -m "Initial commit: Lecture Screenshot Organizer Agent"
git branch -M main

# Create repository on GitHub and push local commits
Write-Host "Creating and pushing repository to GitHub..."
gh repo create "lecture-screenshot-organizer" --public --source=. --push -y

# Signal success
"SUCCESS" | Out-File -FilePath (Join-Path $PSScriptRoot "success.txt") -Encoding utf8
Write-Host "Project pushed to GitHub successfully!"
