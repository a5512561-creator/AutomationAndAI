param(
    [string]$NotebookName = "Switch-DD member weekly",
    [string]$MemberListPath = "",
    [string]$OutputPath = "",
    [int]$ContentPages = 2
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir

if (-not $MemberListPath) {
    $MemberListPath = Join-Path $projectRoot "config\member_list.txt"
}
if (-not $OutputPath) {
    $outDir = Join-Path $projectRoot "output"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputPath = Join-Path $outDir "onenote_pages_$ts.json"
}

if (-not (Test-Path $MemberListPath)) {
    Write-Error "Member list not found: $MemberListPath"
    exit 1
}
$members = Get-Content $MemberListPath -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
Write-Host "Members ($($members.Count)): $($members -join ', ')"

Write-Host "Connecting to OneNote..."
Add-Type -AssemblyName "Microsoft.Office.Interop.OneNote" -ErrorAction SilentlyContinue
$onenote = New-Object -ComObject OneNote.Application

Write-Host "Reading notebook hierarchy..."
$scope = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages
[string]$xml = ''
$onenote.GetHierarchy('', $scope, [ref]$xml)
Write-Host "XML size: $($xml.Length) chars"

if ($xml.Length -lt 500) {
    Write-Error "GetHierarchy returned too little data. Is OneNote open with the notebook synced?"
    exit 1
}

[xml]$doc = $xml
$nsUri = $doc.DocumentElement.NamespaceURI
if (-not $nsUri) { $nsUri = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$notebook = $doc.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $ns)
if (-not $notebook) {
    $allNbs = $doc.SelectNodes("//one:Notebook", $ns)
    foreach ($nb in $allNbs) {
        if ($nb.name -like "*$NotebookName*") { $notebook = $nb; break }
    }
}
if (-not $notebook) {
    Write-Error "Notebook '$NotebookName' not found."
    $allNbs = $doc.SelectNodes("//one:Notebook", $ns)
    Write-Host "Available notebooks:"
    foreach ($nb in $allNbs) { Write-Host "  - $($nb.name)" }
    exit 1
}
Write-Host "Found notebook: $($notebook.name)"

$sections = @{}
foreach ($sec in $notebook.SelectNodes(".//one:Section", $ns)) {
    $sections[$sec.name] = $sec
}
Write-Host "Total sections: $($sections.Count)"

$jsonMembers = New-Object System.Collections.ArrayList
foreach ($member in $members) {
    $sec = $sections[$member]
    if (-not $sec) {
        Write-Host "  WARNING: Section '$member' not found"
        $entry = @{ name = $member; pages = @(); pageContents = @() }
        [void]$jsonMembers.Add($entry)
        continue
    }

    $pageList = New-Object System.Collections.ArrayList
    $datePages = New-Object System.Collections.ArrayList
    foreach ($page in $sec.SelectNodes("one:Page", $ns)) {
        $p = @{
            title = $page.name
            dateTime = $page.dateTime
            lastModifiedTime = $page.lastModifiedTime
        }
        [void]$pageList.Add($p)
        # 接受 YYYY/M/D 或 YYYY/MM/DD（OneNote 可能顯示 2026/3/5 無前導零）
        if ($page.name -match '^\d{4}/\d{1,2}/\d{1,2}$') {
            [void]$datePages.Add($page)
        }
    }
    Write-Host "  $member : $($pageList.Count) pages ($($datePages.Count) date-named)"

    $contentList = New-Object System.Collections.ArrayList
    if ($ContentPages -gt 0 -and $datePages.Count -gt 0) {
        # 依「日期」排序取最新 N 頁，避免字串排序造成 2026/9/15 排在 2026/10/1 後面
        $sorted = $datePages | ForEach-Object {
            if ($_.name -match '^(\d{4})/(\d{1,2})/(\d{1,2})$') {
                try {
                    $dt = [DateTime]::new([int]$matches[1], [int]$matches[2], [int]$matches[3])
                    [PSCustomObject]@{ Node = $_; SortDate = $dt }
                } catch { $null }
            }
        } | Where-Object { $_ } | Sort-Object -Property SortDate -Descending
        $topN = $sorted | Select-Object -First $ContentPages | ForEach-Object { $_.Node }
        foreach ($dp in $topN) {
            [string]$pxml = ''
            try {
                $onenote.GetPageContent($dp.ID, [ref]$pxml, 0)
                $cEntry = @{
                    title = $dp.name
                    lastModifiedTime = $dp.lastModifiedTime
                    xml = $pxml
                }
                [void]$contentList.Add($cEntry)
                Write-Host "    content: $($dp.name) ($($pxml.Length) chars)"
            } catch {
                Write-Host "    ERROR reading content for $($dp.name): $_"
            }
        }
    }

    $entry = @{
        name = $member
        pages = $pageList.ToArray()
        pageContents = $contentList.ToArray()
    }
    [void]$jsonMembers.Add($entry)
}

$output = @{
    members = $jsonMembers.ToArray()
    exportTime = (Get-Date -Format "o")
}

$json = $output | ConvertTo-Json -Depth 10
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutputPath, $json, $utf8NoBom)

Write-Host ""
Write-Host "JSON saved to: $OutputPath"
