# diagnostics/test_onenote.ps1
# A diagnostic script to verify OneNote COM connection and list all pages in the hierarchy.

try {
    $oneNote = New-Object -ComObject OneNote.Application
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "✅ COM Object Created Successfully!" -ForegroundColor Green
    Write-Host "==========================================================================" -ForegroundColor Cyan
    
    [ref]$xml = ""
    $oneNote.GetHierarchy($null, 4, [ref]$xml) # 4 = hsPages
    
    $hierarchyXml = [xml]$xml.Value
    $ns = New-Object System.Xml.XmlNamespaceManager($hierarchyXml.NameTable)
    $ns.AddNamespace("one", $hierarchyXml.DocumentElement.NamespaceURI)
    
    $pages = $hierarchyXml.SelectNodes("//one:Page", $ns)
    Write-Host "  Notebook Hierarchy scanned successfully." -ForegroundColor Gray
    Write-Host "  Total Pages Found: $($pages.Count)" -ForegroundColor Yellow
    
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
        Write-Host "  Most Recently Modified Page detected:" -ForegroundColor Gray
        Write-Host "    Name: $($latestPage.name)" -ForegroundColor Green
        Write-Host "    ID: $($latestPage.ID)" -ForegroundColor DarkGray
        Write-Host "    Last Modified: $($latestPage.lastModifiedTime)" -ForegroundColor DarkGray
    } else {
        Write-Host "  No pages found in hierarchy. Please create a page in OneNote first." -ForegroundColor Yellow
    }
    Write-Host "==========================================================================" -ForegroundColor Cyan
} catch {
    Write-Error $_
}
