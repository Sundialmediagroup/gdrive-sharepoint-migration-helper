#Requires -Version 7.2
# Creates an identical tree of SharePoint-incompatible test names in each empty target
# folder or shared drive. Test tooling only: never point it at production content.
[CmdletBinding(SupportsShouldProcess)]
param([Parameter(Mandatory)][string[]]$TargetFolderId)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$api = 'https://www.googleapis.com/drive/v3/'
$folder = 'application/vnd.google-apps.folder'
$cmdlet = $PSCmdlet

function Invoke-Google {
    param([string]$Uri, [string]$Method = 'Get', [byte[]]$Body, [string]$ContentType = 'application/json; charset=utf-8')
    if (!$env:GDRIVE_ACCESS_TOKEN) { throw 'Set GDRIVE_ACCESS_TOKEN to a short-lived Google OAuth access token.' }
    $request = @{ Uri = $Uri; Method = $Method; Headers = @{ Authorization = 'Bearer ' + $env:GDRIVE_ACCESS_TOKEN }; TimeoutSec = 120 }
    if ($null -ne $Body) { $request.Body = $Body; $request.ContentType = $ContentType }
    Invoke-RestMethod @request
}

function New-TestItem {
    param([string]$Parent, [string]$Name, [string]$MimeType)
    $metadata = ConvertTo-Json @{ name = $Name; parents = @($Parent); mimeType = $MimeType } -Compress
    if ($MimeType.StartsWith('application/vnd.google-apps.')) {
        return Invoke-Google ($api + 'files?supportsAllDrives=true&fields=id,name') Post ([Text.Encoding]::UTF8.GetBytes($metadata))
    }
    $boundary = 'drive-test-' + [guid]::NewGuid()
    $body = "--$boundary`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$metadata`r`n" +
        "--$boundary`r`nContent-Type: $MimeType`r`n`r`nDriveMigration test file.`r`n--$boundary--`r`n"
    Invoke-Google 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name' Post ([Text.Encoding]::UTF8.GetBytes($body)) "multipart/related; boundary=$boundary"
}

function Add-TestTree {
    param([string]$Parent, [object[]]$Nodes, [string]$Location)
    foreach ($node in $Nodes) {
        $mime = if ($node['Mime']) { $node['Mime'] } else { 'text/plain' }
        $path = "$Location > $($node.Name)"
        $id = '(WhatIf)'
        if ($cmdlet.ShouldProcess($path, "Create $mime")) {
            $created = New-TestItem $Parent $node.Name $mime
            $id = $created.id
            if ($created.name -cne $node.Name) { Write-Warning "Google stored '$($created.name)' instead of '$($node.Name)'." }
        }
        [pscustomobject]@{ Path = $path; MimeType = $mime; Id = $id }
        if ($node['Children']) { Add-TestTree $id $node['Children'] $path }
    }
}

# Six long folder levels plus a long leaf exceed 400 characters under a typical site prefix.
$deep = @{ Name = ('Annual compliance report ' * 4).Trim() + '.pdf' }
for ($level = 6; $level -ge 1; $level--) {
    $deep = @{ Name = "Level $level - " + ('Department records archive ' * 2).Trim(); Mime = $folder; Children = @($deep) }
}
$tree = @(
    @{ Name = 'budget:2026?.pdf' },
    @{ Name = 'Q1/Q2 Report.pdf' },
    @{ Name = '"Final" <draft>.pdf' },
    @{ Name = 'a|b*c.txt' },
    @{ Name = 'CON.txt' },
    @{ Name = '.lock' },
    @{ Name = 'desktop.ini' },
    @{ Name = '~$budget.xlsx' },
    @{ Name = 'my_vti_file.pdf' },
    @{ Name = 'notes.txt.' },
    @{ Name = '  padded.txt  ' },
    @{ Name = ('Quarterly financial summary ' * 11).Trim() + '.pdf' },
    @{ Name = 'Report.pdf' },
    @{ Name = 'report.pdf' },
    @{ Name = 'Notes'; Mime = 'application/vnd.google-apps.document' },
    @{ Name = 'Notes.docx' },
    @{ Name = 'Budget 2026'; Mime = 'application/vnd.google-apps.spreadsheet' },
    @{ Name = 'forms'; Mime = $folder; Children = @(@{ Name = 'readme.txt' }) },
    $deep
)

# Check every target before creating anything, so both trees start identical.
if (@($TargetFolderId | Sort-Object -Unique).Count -ne $TargetFolderId.Count) { throw 'Target IDs must be distinct.' }
$targets = foreach ($targetId in $TargetFolderId) {
    $target = Invoke-Google ($api + 'files/' + [uri]::EscapeDataString($targetId) + '?supportsAllDrives=true&fields=id,name,mimeType,driveId')
    if ($target.mimeType -ne $folder) { throw "Target $targetId is not a folder or shared drive." }
    $corpus = '&corpora=user'
    if ($target.PSObject.Properties['driveId'] -and $target.driveId) { $corpus = '&corpora=drive&driveId=' + [uri]::EscapeDataString($target.driveId) }
    $query = [uri]::EscapeDataString("'$($target.id)' in parents and trashed = false")
    $existing = Invoke-Google ($api + "files?q=$query&pageSize=1&fields=files(id)&supportsAllDrives=true&includeItemsFromAllDrives=true" + $corpus)
    if (@($existing.files).Count) { throw "Target '$($target.name)' is not empty. Use an empty test folder or shared drive." }
    $target
}
foreach ($target in $targets) { Add-TestTree $target.id $tree $target.name }
