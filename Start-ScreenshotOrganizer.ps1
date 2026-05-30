# Start-ScreenshotOrganizer.ps1
# A premium, zero-dependency background automation agent to auto-paste lecture screenshots into Microsoft OneNote.

# Set output encoding to UTF8 for beautiful terminal characters
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Load required Windows Forms and Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Clear screen and draw a stunning premium ASCII Art header
Clear-Host
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "   * LECTURE SCREENSHOT ORGANIZER FOR ONENOTE *" -ForegroundColor Yellow
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "  This agent monitors your clipboard for screenshots (e.g. via Win+Shift+S)" -ForegroundColor Gray
Write-Host "  and automatically appends them to your active OneNote page in real time." -ForegroundColor Gray
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host ""

# State variables
$lastImageHash = ""
$oneNote = $null
$activePageId = ""
$activePageTitle = ""
$lastResolvedPageId = ""
$hasAttemptedLaunch = $false
$logPath = Join-Path $PSScriptRoot "error_log.txt"

# Sound helper - plays a subtle Windows notification sound
function Play-SuccessSound {
    try {
        [System.Media.SystemSounds]::Asterisk.Play()
    } catch {}
}

# Logger helper for debugging
function Log-Error($message, $exception) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMsg = "[$timestamp] ERROR: $message`nException: $($exception.ToString())`n"
    $logMsg | Out-File -FilePath $logPath -Append -Encoding utf8
    
    Write-Host "`n[ERROR] $message" -ForegroundColor Red
    Write-Host "$($exception.Message)" -ForegroundColor DarkRed
    Write-Host "Detailed log written to: error_log.txt" -ForegroundColor Gray
}

# Main event loop
Write-Host "Running Screenshot Watcher..." -ForegroundColor Green
Write-Host "TIP: Use [Win + Shift + S] to crop a lecture slide anytime!" -ForegroundColor Magenta
Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkGray

while ($true) {
    try {
        # 1. Ensure connection to OneNote
        if ($null -eq $oneNote) {
            $oneNote = New-Object -ComObject OneNote.Application -ErrorAction Stop
        }
        
        # 2. Resolve the Active Page (highly robust multi-stage detection)
        $activePageId = ""
        $activePageTitle = ""
        
        # Stage A: Try getting page from current active window
        $window = $oneNote.Windows.CurrentWindow
        if ($null -eq $window) {
            # Stage B: Fall back to first open window in collection
            $window = $oneNote.Windows | Select-Object -First 1
        }
        
        if ($null -ne $window) {
            $activePageId = $window.CurrentPageId
        }
        
        # Stage C: If window-based detection yields nothing, scan hierarchy for the most recently modified page
        if ([string]::IsNullOrEmpty($activePageId)) {
            [ref]$hierarchyRef = ""
            $oneNote.GetHierarchy($null, 4, [ref]$hierarchyRef) # 4 = hsPages
            $hierarchyXml = [xml]$hierarchyRef.Value
            
            $ns = New-Object System.Xml.XmlNamespaceManager($hierarchyXml.NameTable)
            $ns.AddNamespace("one", $hierarchyXml.DocumentElement.NamespaceURI)
            
            $pages = $hierarchyXml.SelectNodes("//one:Page", $ns)
            $latestPage = $null
            $latestTime = [System.DateTime]::MinValue
            
            foreach ($page in $pages) {
                $timeStr = $page.lastModifiedTime
                if ($null -ne $timeStr) {
                    $parsedTime = [System.DateTime]::Parse($timeStr)
                    if ($parsedTime -gt $latestTime) {
                        $latestTime = $parsedTime
                        $latestPage = $page
                    }
                }
            }
            
            if ($null -ne $latestPage) {
                $activePageId = $latestPage.ID
                $activePageTitle = $latestPage.name
            }
        } else {
            # Window was found, so fetch page details
            [ref]$pageXmlRef = ""
            $oneNote.GetPageContent($activePageId, [ref]$pageXmlRef, 1)
            $pageXml = [xml]$pageXmlRef.Value
            $activePageTitle = $pageXml.Page.title
        }
        
        # Handle blank/unnamed page title
        if ([string]::IsNullOrEmpty($activePageTitle)) {
            $activePageTitle = "Untitled Page"
        }
        
        # 3. If still no active GUI window, force-launch it!
        $window = $oneNote.Windows.CurrentWindow
        if ($null -eq $window) {
            $window = $oneNote.Windows | Select-Object -First 1
        }
        
        if ($null -eq $window -and -not [string]::IsNullOrEmpty($activePageId)) {
            if (-not $hasAttemptedLaunch) {
                $hasAttemptedLaunch = $true
                
                # BREATHTAKING FIX: If OneNote is running as an invisible background process,
                # calling NavigateTo will instantly force Windows to make the OneNote GUI visible
                # and jump straight to the page where we are pasting!
                try {
                    Write-Host "`n[SYSTEM] Restoring OneNote Desktop GUI window..." -ForegroundColor Yellow
                    $oneNote.NavigateTo($activePageId)
                    Start-Sleep -Seconds 3
                    continue
                } catch {
                    # If NavigateTo fails, fall back to launching the EXE directly
                    $exePaths = @(
                        "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
                        "C:\Program Files\Microsoft Office\Office16\ONENOTE.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE"
                    )
                    foreach ($path in $exePaths) {
                        if (Test-Path $path) {
                            Start-Process $path
                            break
                        }
                    }
                }
            }
        }
        
        # Reset launch attempt flag when a page and GUI window are successfully resolved
        if ($null -ne $window) {
            $hasAttemptedLaunch = $false
        }
        
        # Log resolved page changes dynamically
        if ($activePageId -ne $lastResolvedPageId) {
            $lastResolvedPageId = $activePageId
            Write-Host "`nActive OneNote Page: " -NoNewline -ForegroundColor Gray
            Write-Host "$activePageTitle" -ForegroundColor Green
        }
        
        # 4. Check clipboard for image
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $img = $null
            # Retry clip reading 3 times in case of Windows locking the clipboard temporarily
            for ($i = 0; $i -lt 3; $i++) {
                try {
                    $img = [System.Windows.Forms.Clipboard]::GetImage()
                    if ($null -ne $img) { break }
                } catch {
                    Start-Sleep -Milliseconds 100
                }
            }
            
            if ($null -ne $img) {
                # Calculate image hash to check if it's a new screenshot
                $ms = New-Object System.IO.MemoryStream
                $img.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
                $bytes = $ms.ToArray()
                $hash = [System.Security.Cryptography.MD5]::Create().ComputeHash($bytes)
                $hashString = [System.BitConverter]::ToString($hash)
                $ms.Dispose()
                
                if ($hashString -ne $lastImageHash) {
                    $lastImageHash = $hashString
                    $timestamp = Get-Date -Format "HH:mm:ss"
                    
                    Write-Host "[ $timestamp ] New screenshot detected in clipboard!" -ForegroundColor Cyan
                    Write-Host "Processing and pasting to OneNote..." -ForegroundColor Gray
                    
                    # Convert to Base64
                    $base64 = [Convert]::ToBase64String($bytes)
                    
                    # Get fresh content of the active page to discover the dynamic schema namespace URI
                    [ref]$pageXmlRef = ""
                    $oneNote.GetPageContent($activePageId, [ref]$pageXmlRef, 1)
                    $pageXml = [xml]$pageXmlRef.Value
                    
                    # Safely extract the active OneNote schema namespace URI
                    $namespaceUri = $pageXml.DocumentElement.NamespaceURI
                    if ([string]::IsNullOrEmpty($namespaceUri)) {
                        $namespaceUri = "http://schemas.microsoft.com/office/onenote/2013/onenote"
                    }
                    
                    # Construct a highly robust, minimal XML update snippet
                    # Omitting the objectID on the Outline forces OneNote to append it as new content!
                    $timeString = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $xmlContent = @"
<?xml version="1.0"?>
<one:Page xmlns:one="$namespaceUri" ID="$activePageId">
    <one:Outline>
        <one:Position x="72" y="72" />
        <one:OEChildren>
            <one:OE>
                <one:Image format="png">
                    <one:Data>$base64</one:Data>
                </one:Image>
            </one:OE>
            <one:OE>
                <one:T><![CDATA[Lecture Captured at $timeString]]></one:T>
            </one:OE>
        </one:OEChildren>
    </one:Outline>
</one:Page>
"@
                    
                    # Send update to OneNote
                    $oneNote.UpdatePageContent($xmlContent)
                    
                    Play-SuccessSound
                    Write-Host "SUCCESS: Pasted screenshot into " -NoNewline -ForegroundColor Green
                    Write-Host "'$activePageTitle'" -ForegroundColor Green
                    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkGray
                }
            }
        }
    } catch {
        # Handle exceptions gracefully and log them for diagnostics
        Log-Error "An unexpected error occurred during execution." $_
        $oneNote = $null
        Write-Host "Attempting to recover and reconnect in 5 seconds..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }
    
    # Check the clipboard twice per second for instant responsiveness
    Start-Sleep -Milliseconds 500
}
