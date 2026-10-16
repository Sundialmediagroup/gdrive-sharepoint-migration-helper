#Requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force
$module = Get-Module DriveMigration
$script:checks = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (!$Condition) { throw $Message }
    $script:checks++
}

function New-ItemFixture {
    param([string]$Id, [string]$Name, [string]$Parent = 'test-root', [string]$Mime = 'application/pdf', [bool]$CanRename = $true)
    [pscustomobject]@{ id = $Id; name = $Name; parents = @($Parent); mimeType = $Mime; version = '1'; capabilities = @{ canRename = $CanRename } }
}

function New-TestPlan {
    param([AllowEmptyCollection()][object[]]$Items, [int]$MaxPathLength = 400)
    New-DriveRenamePlan -Items $Items -RootId test-root -DestinationUrl 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents' -MaxPathLength $MaxPathLength
}

$plan = New-TestPlan @(
    (New-ItemFixture a ('x' * 280 + '.pdf')),
    (New-ItemFixture b 'budget:2026?.xlsx'),
    (New-ItemFixture c 'CON.txt'),
    (New-ItemFixture d 'Report' -Mime 'application/vnd.google-apps.document'),
    (New-ItemFixture e 'report.docx'),
    (New-ItemFixture f '  .lock  '),
    (New-ItemFixture g 'a_vti_b'),
    (New-ItemFixture h 'Case.pdf'),
    (New-ItemFixture i 'case.pdf'),
    (New-ItemFixture j 'no:permission.pdf' -CanRename $false),
    (New-ItemFixture k 'shortcut' -Mime 'application/vnd.google-apps.shortcut')
)
Assert-True ($plan.Rows[0].NewName.Length -le 120 -and $plan.Rows[0].NewName.EndsWith('.pdf')) 'Long name or extension failed.'
Assert-True ($plan.Rows[1].NewName -eq 'budget_2026_.xlsx') 'Invalid character replacement failed.'
Assert-True ($plan.Rows[2].NewName -eq '_CON.txt') 'Reserved device name failed.'
Assert-True ($plan.Rows[3].TargetName -eq 'Report.docx' -and $plan.Rows[3].NewName -eq 'Report') 'Google extension projection failed.'
Assert-True ($plan.Rows[4].TargetName -match '~[0-9a-f]{10}\.docx$') 'Post-conversion collision failed.'
Assert-True ($plan.Rows[5].NewName -eq '_.lock') 'Whitespace or reserved name failed.'
Assert-True ($plan.Rows[6].NewName -eq 'a_vti-b') 'Embedded reserved token failed.'
Assert-True ($plan.Rows[8].NewName -match '~[0-9a-f]{10}\.pdf$') 'Case-insensitive collision failed.'
Assert-True ($plan.Rows[9].Warnings -like '*BLOCKED:*') 'Rename permission check failed.'
Assert-True ($plan.Rows[10].Warnings -like '*Unsupported*') 'Unsupported item warning failed.'
Assert-True ($plan.Rows[0].PathLength -eq ('/sites/IT/Shared Documents/' + $plan.Rows[0].TargetName).Length) 'Decoded URL budget failed.'

$tree = @(
    (New-ItemFixture folder1 ('Alpha' * 30) -Mime 'application/vnd.google-apps.folder'),
    (New-ItemFixture folder2 ('Beta' * 40) -Parent folder1 -Mime 'application/vnd.google-apps.folder'),
    (New-ItemFixture leaf ('Report' * 30 + '.pdf') -Parent folder2)
)
$deep = New-TestPlan $tree -MaxPathLength 160
Assert-True (@($deep.Rows | Where-Object PathLength -gt 160).Count -eq 0) 'Ancestor shortening failed.'
Assert-True ($deep.Rows[0].Rename -and $deep.Rows[1].Rename) 'Ancestor renames missing.'
$stable = New-TestPlan @($tree | Sort-Object id -Descending) -MaxPathLength 160
Assert-True (($deep.Rows | ConvertTo-Json -Depth 10 -Compress) -ceq ($stable.Rows | ConvertTo-Json -Depth 10 -Compress)) 'Plan depends on API listing order.'
$empty = New-TestPlan @()
Assert-True (@($empty.Rows).Count -eq 0) 'Empty source failed.'
$failed = $false
try { New-TestPlan @((New-ItemFixture bad name -Parent missing)) | Out-Null } catch { $failed = $true }
Assert-True $failed 'Orphan inventory was accepted.'
$failed = $false
try { New-TestPlan $tree -MaxPathLength 80 | Out-Null } catch { $failed = $true }
Assert-True $failed 'Impossible hierarchy was accepted.'
$emoji = [char]::ConvertFromUtf32(0x1f600)
$unicode = New-TestPlan @((New-ItemFixture unicode (($emoji * 100) + '.pdf')))
$encoder = [Text.UTF8Encoding]::new($false, $true)
Assert-True ($encoder.GetByteCount($unicode.Rows[0].NewName) -gt 0) 'Unicode was split inside a surrogate pair.'

# Report-targeted mode: only reported items and path-budget ancestors change; collisions are left to Migration Manager.
$deepLeafPath = '/' + ('Alpha' * 30) + '/' + ('Beta' * 40) + '/' + ('Report' * 30 + '.pdf')
$reported = New-DriveRenamePlan -RootId test-root -DestinationUrl 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents' -Items (@(
    (New-ItemFixture con 'CON.txt'),
    (New-ItemFixture budget 'budget:2026?.pdf'),
    (New-ItemFixture upper 'Report.pdf'),
    (New-ItemFixture lower 'report.pdf'),
    (New-ItemFixture fine 'good.pdf')
) + $tree) -ReportFailures @(
    [pscustomobject]@{ FullPath = '/CON.txt'; ResultCode = 'MJOBERROR'; FailureReason = 'Item level failure' },
    [pscustomobject]@{ FullPath = '/good.pdf'; ResultCode = 'MJOBERROR'; FailureReason = 'Item level failure' },
    [pscustomobject]@{ FullPath = '/missing.pdf'; ResultCode = 'MJOBERROR'; FailureReason = 'Item level failure' },
    [pscustomobject]@{ FullPath = $deepLeafPath; ResultCode = 'MITEMPATHLENGTH'; FailureReason = 'Item path length exceeds the limit (400)' }
)
$byId = @{}; foreach ($row in $reported.Rows) { $byId[$row.Id] = $row }
$status = @{}; foreach ($finding in $reported.ReportFindings) { $status[$finding.FullPath] = $finding.Status }
Assert-True ($byId['con'].NewName -eq '_CON.txt') 'Reported reserved name was not renamed.'
Assert-True (!$byId['budget'].Rename -and !$byId['fine'].Rename) 'Unreported or non-naming items were renamed.'
$pair = @($byId['upper'], $byId['lower'])
Assert-True (!@($pair | Where-Object Rename).Count -and @($pair | Where-Object { $_.TargetName -match '^[Rr]eport\(1\)\.pdf$' -and $_.Warnings -like '*(n) suffix*' }).Count -eq 1) 'Unreported collision was not left to Migration Manager with a projected (n) suffix.'
Assert-True (@($reported.Rows | Where-Object PathLength -gt 400).Count -eq 0 -and $byId['folder2'].Rename) 'Reported path length was not fixed through an ancestor.'
Assert-True ($status['/CON.txt'] -eq 'Fixed' -and $status[$deepLeafPath] -eq 'Fixed' -and $status['/good.pdf'] -eq 'NotFixed' -and $status['/missing.pdf'] -eq 'NotFound') 'Report finding status was wrong.'

& $module {
    $script:ListCalls = [Collections.Generic.List[string]]::new()
    $script:Incomplete = $false
    function script:Invoke-DriveApi {
        param($Path, $Method = 'Get', $Body)
        if ($Path.StartsWith('files/test-root?')) {
            return [pscustomobject]@{ id = 'test-root'; mimeType = $script:FolderMime; driveId = 'shared-drive' }
        }
        $script:ListCalls.Add($Path)
        if ($script:Incomplete) { return [pscustomobject]@{ files = @(); incompleteSearch = $true } }
        if ($Path -notmatch 'pageToken=') { return [pscustomobject]@{ files = @(); nextPageToken = 'page-two'; incompleteSearch = $false } }
        [pscustomobject]@{ files = @([pscustomobject]@{ id = 'a'; name = 'file.pdf'; parents = @('test-root'); mimeType = 'application/pdf'; version = '1'; capabilities = @{ canRename = $true } }); incompleteSearch = $false }
    }
}
$inventory = @(Get-DriveInventory test-root)
Assert-True ($inventory.Count -eq 1) 'Pagination stopped on an empty intermediate page.'
$calls = & $module { @($script:ListCalls.ToArray()) }
Assert-True ($calls.Count -eq 2 -and $calls[1] -match 'pageToken=page-two') 'Page token was not followed.'
Assert-True ($calls[0] -match 'corpora=drive&driveId=shared-drive') 'Shared drive corpus missing.'
& $module { $script:Incomplete = $true }
$failed = $false
try { Get-DriveInventory test-root | Out-Null } catch { $failed = $true }
Assert-True $failed 'Incomplete inventory was accepted.'

$temp = Join-Path ([IO.Path]::GetTempPath()) ('drive-migration-test-' + [guid]::NewGuid())
[void](New-Item -Path $temp -ItemType Directory)
Push-Location $temp
try {
    $script:fixture = @((New-ItemFixture a 'bad:name.pdf'), (New-ItemFixture b 'good.pdf'))
    $applyPlan = New-TestPlan $script:fixture
    & $module {
        param($Fixture)
        $script:Fixture = $Fixture
        $script:PatchCount = 0
        $script:FailAfterPatch = $false
        function script:Get-DriveInventory { param($RootId) $script:Fixture }
        function script:Invoke-DriveApi {
            param($Path, $Method = 'Get', $Body)
            $id = ($Path -split '[/?]')[1]
            $item = $script:Fixture | Where-Object id -eq $id
            if ($Method -eq 'Patch') {
                $script:PatchCount++
                $item.name = $Body.name
                $item.version = [string](1 + [int]$item.version)
                if ($script:FailAfterPatch) { $script:FailAfterPatch = $false; throw 'Simulated lost response after successful PATCH.' }
            }
            $item
        }
    } $script:fixture
    # Relative on purpose: Push-Location does not move the .NET current directory.
    $journal = 'journal.jsonl'
    Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -WhatIf | Out-Null
    Assert-True (!(Test-Path $journal) -and $script:fixture[0].name -eq 'bad:name.pdf') 'WhatIf wrote data.'
    Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Confirm:$false | Out-Null
    Assert-True ($script:fixture[0].name -eq 'bad_name.pdf') 'Apply failed.'
    Assert-True (Test-Path -LiteralPath (Join-Path $temp 'journal.jsonl')) 'Relative journal path ignored the PowerShell location.'
    Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Confirm:$false | Out-Null
    Assert-True ((& $module { $script:PatchCount }) -eq 1) 'Resume repeated an already completed PATCH.'
    Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Rollback -Confirm:$false | Out-Null
    Assert-True ($script:fixture[0].name -eq 'bad:name.pdf') 'Rollback failed.'
    & $module { $script:FailAfterPatch = $true }
    $failed = $false
    try { Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Confirm:$false | Out-Null } catch { $failed = $true }
    Assert-True ($failed -and $script:fixture[0].name -eq 'bad_name.pdf') 'Lost-response fixture failed.'
    Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Rollback -Confirm:$false | Out-Null
    Assert-True ($script:fixture[0].name -eq 'bad:name.pdf') 'Write-ahead intent did not recover a lost response.'
    $script:fixture[1].name = 'user-edited.pdf'
    $before = & $module { $script:PatchCount }
    $failed = $false
    try { Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Confirm:$false | Out-Null } catch { $failed = $true }
    Assert-True ($failed -and (& $module { $script:PatchCount }) -eq $before) 'Drift preflight allowed writes.'
    $script:fixture[1].name = 'good.pdf'
    $script:fixture[1].version = '2'
    $failed = $false
    try { Invoke-DriveRenamePlan $applyPlan $journal -ApproveSourceRenames -Confirm:$false | Out-Null } catch { $failed = $true }
    Assert-True $failed 'Version drift was accepted.'
} finally {
    Pop-Location
    Remove-Item -LiteralPath $temp -Recurse -Force
    Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force
}
# Large files: download, manual-download import, and PnP upload with mocked Google and SharePoint calls.
$module = Get-Module DriveMigration
$savedToken = $env:GDRIVE_ACCESS_TOKEN
$env:GDRIVE_ACCESS_TOKEN = 'offline-test-token'
$failed = $false
try { & $module { Invoke-DriveDownload 'https://evil.example.com/file' 'unused' } } catch { $failed = $_.Exception.Message -like '*Refusing*' }
Assert-True $failed 'Google token could be sent to a non-Google host.'

$largeItems = @(
    (New-ItemFixture projects 'Projects' -Mime 'application/vnd.google-apps.folder'),
    (New-ItemFixture bigdoc 'Big Report' -Parent projects -Mime 'application/vnd.google-apps.document'),
    (New-ItemFixture sheet 'Budget' -Mime 'application/vnd.google-apps.spreadsheet'),
    (New-ItemFixture scan 'scan.pdf'),
    (New-ItemFixture badname 'bad:name.pdf'),
    (New-ItemFixture survey 'Survey' -Mime 'application/vnd.google-apps.form')
)
$largePlan = New-DriveRenamePlan -Items $largeItems -RootId test-root -DestinationUrl 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents' -ReportFailures @()
$selected = @(Select-DrivePlanRow -Plan $largePlan -ReportFailures @([pscustomobject]@{ FullPath = '/Projects/Big Report' }))
Assert-True ($selected.Count -eq 1 -and $selected[0].Id -eq 'bigdoc') 'Report path did not select the plan row.'
$failed = $false
try { Select-DrivePlanRow -Plan $largePlan -Id projects | Out-Null } catch { $failed = $true }
Assert-True $failed 'A folder was accepted as a large file.'

$temp = Join-Path ([IO.Path]::GetTempPath()) ('drive-largefile-test-' + [guid]::NewGuid())
$staging = Join-Path $temp 'staging'
$inbox = Join-Path $temp 'inbox'
[void](New-Item -Path $inbox -ItemType Directory -Force)
try {
    & $module {
        param($Items)
        $script:Items = $Items
        $script:Downloads = [Collections.Generic.List[string]]::new()
        function script:Invoke-DriveApi {
            param($Path, $Method = 'Get', $Body)
            $id = ($Path -split '[/?]')[1]
            if ($Path -match 'fields=exportLinks') {
                return '{"exportLinks":{"application/vnd.openxmlformats-officedocument.wordprocessingml.document":"https://docs.google.com/export/bigdoc"}}' | ConvertFrom-Json
            }
            $script:Items | Where-Object id -eq $id
        }
        function script:Invoke-DriveDownload {
            param($Uri, $OutFile)
            $script:Downloads.Add($Uri)
            if ($Uri -match '/files/bigdoc/export') { throw 'Google download failed (HTTP 403 exportSizeLimitExceeded).' }
            # The spreadsheet export returns an HTML page instead of a workbook.
            $content = if ($Uri -match 'sheet') { '<html>' } elseif ($Uri -match 'alt=media') { '%PDF-1.7' } else { 'PK' + [char]3 + [char]4 }
            [IO.File]::WriteAllText($OutFile, $content)
        }
    } $largeItems
    $downloaded = @(Save-DriveLargeFile -Plan $largePlan -Rows @(Select-DrivePlanRow -Plan $largePlan -Id bigdoc, sheet, scan, badname, survey) -StagingPath $staging)
    $state = @{}; foreach ($result in $downloaded) { $state[$result.Id] = $result }
    Assert-True ($state['bigdoc'].Status -eq 'Downloaded' -and $state['bigdoc'].Method -eq 'ExportLink') 'Export limit did not fall back to the export link.'
    Assert-True (Test-Path -LiteralPath (Join-Path $staging 'Projects/Big Report.docx')) 'Download was not rebuilt into its original folder.'
    Assert-True ($state['sheet'].Status -eq 'ManualDownload' -and !@(Get-ChildItem $staging -Recurse -Filter *.partial).Count) 'Invalid download content was kept.'
    Assert-True ($state['scan'].Status -eq 'Downloaded' -and $state['scan'].Method -eq 'Media') 'Binary download failed.'
    Assert-True ($state['badname'].Status -eq 'InvalidTarget' -and $state['survey'].Status -eq 'ManualReview') 'Invalid or non-exportable items were downloaded.'
    $before = (& $module { $script:Downloads.Count })
    Save-DriveLargeFile -Plan $largePlan -Rows @(Select-DrivePlanRow -Plan $largePlan -Id scan) -StagingPath $staging | Out-Null
    Assert-True ((& $module { $script:Downloads.Count }) -eq $before) 'An already staged file was downloaded again.'
    Remove-Item -LiteralPath $staging -Recurse -Force
    $budget = @(Save-DriveLargeFile -Plan $largePlan -Rows @(Select-DrivePlanRow -Plan $largePlan -Id scan, bigdoc) -StagingPath $staging -MaxDownloadBytes 1)
    Assert-True ($budget[0].Status -eq 'Downloaded' -and $budget[1].Status -eq 'Deferred') 'Daily download budget was not enforced.'
    Save-DriveLargeFile -Plan $largePlan -Rows @(Select-DrivePlanRow -Plan $largePlan -Id bigdoc) -StagingPath $staging | Out-Null
    $retry = & $module { @((Test-DriveRetryable 403 '{"message":"User rate limit exceeded."}'), (Test-DriveRetryable 403 '"reason": "userRateLimitExceeded"'), (Test-DriveRetryable 429 ''), (Test-DriveRetryable 403 '"reason": "exportSizeLimitExceeded"'), (Test-DriveRetryable 404 '')) }
    Assert-True (($retry -join ',') -eq 'True,True,True,False,False') 'Google quota errors were not classified as retryable.'
    $sizes = & $module { @((Test-UploadedSize 12357 4937 'a.xlsx'), (Test-UploadedSize 0 4937 'a.docx'), (Test-UploadedSize 5 5 'a.mp4'), (Test-UploadedSize 4 5 'a.mp4')) }
    Assert-True (($sizes -join ',') -eq 'True,False,True,False') 'Upload size check mishandled SharePoint-rewritten Office files or truncated files.'

    [IO.File]::WriteAllText((Join-Path $inbox 'Budget (1).xlsx'), 'PK' + [char]3 + [char]4)
    [IO.File]::WriteAllText((Join-Path $inbox 'scan.pdf'), '%PDF-1.7')
    [IO.File]::WriteAllText((Join-Path $inbox 'random.txt'), 'x')
    $imported = @(Import-DriveManualDownload -Plan $largePlan -Rows @($largePlan.Rows | Where-Object MimeType -ne 'application/vnd.google-apps.folder') -InboxPath $inbox -StagingPath $staging)
    $byFile = @{}; foreach ($result in $imported | Where-Object File) { $byFile[(Split-Path $result.File -Leaf)] = $result.Status }
    Assert-True ($byFile['Budget (1).xlsx'] -eq 'Imported' -and (Test-Path -LiteralPath (Join-Path $staging 'Budget.xlsx'))) 'Manual download was not matched into staging.'
    Assert-True ($byFile['scan.pdf'] -eq 'AlreadyStaged' -and $byFile['random.txt'] -eq 'Unmatched') 'Import overwrote staging or matched an unrelated file.'
    Assert-True (@($imported | Where-Object { $_.Id -eq 'badname' -and $_.Status -eq 'ManualReview' }).Count -eq 1) 'Invalid target was offered for import.'

    $script:SpFiles = @{ '/sites/IT/Shared Documents/scan.pdf' = 1L }
    $script:SpFolders = @('/sites/IT/Shared Documents')
    $script:AddCalls = [Collections.Generic.List[string]]::new()
    function global:Get-PnPConnection { [CmdletBinding()] param() [pscustomobject]@{ Url = 'https://contoso.sharepoint.com/sites/IT' } }
    function global:Get-PnPFolder { [CmdletBinding()] param($Url) if ($Url -notin $script:SpFolders) { throw 'Folder not found.' } [pscustomobject]@{ ServerRelativeUrl = $Url } }
    function global:Get-PnPFile {
        [CmdletBinding()] param($Url, [switch]$AsListItem)
        if ($script:SpFiles.ContainsKey($Url)) { return @{ File_x0020_Size = $script:SpFiles[$Url] } }
        # Real PnP throws instead of returning nothing when the file's folder is missing as well.
        if ($Url.Substring(0, $Url.LastIndexOf('/')) -notin $script:SpFolders) { throw 'Unexpected response from the server. The status code is "NotFound".' }
    }
    function global:Resolve-PnPFolder {
        [CmdletBinding()] param($SiteRelativePath)
        $script:SpFolders += "/sites/IT/$SiteRelativePath"
    }
    function global:Add-PnPFile {
        [CmdletBinding()] param($Path, $Folder, $NewFileName)
        # Real PnP 3.4.1 throws NotFound instead of creating a missing folder.
        if ("/sites/IT/$Folder" -notin $script:SpFolders) { throw 'Unexpected response from the server. The status code is "NotFound".' }
        $script:AddCalls.Add("$Folder|$NewFileName")
        # Like SharePoint, store Office files with extra document properties (a different size).
        $script:SpFiles["/sites/IT/$Folder/$NewFileName"] = (Get-Item -LiteralPath $Path).Length + $(if ($NewFileName -match '\.(docx|xlsx|pptx)$') { 7000 } else { 0 })
    }
    $uploadJournal = Join-Path $temp 'upload.jsonl'
    $upload = @{ Plan = $largePlan; StagingPath = $staging; JournalPath = $uploadJournal; ApproveUpload = $true }
    Publish-DriveLargeFile @upload -WhatIf | Out-Null
    Assert-True (!$script:AddCalls.Count -and !(Test-Path $uploadJournal)) 'Upload WhatIf wrote data.'
    $uploaded = @{}; foreach ($result in Publish-DriveLargeFile @upload -Confirm:$false) { $uploaded[$result.Id] = $result.Status }
    Assert-True ($uploaded['sheet'] -eq 'Uploaded' -and $script:AddCalls -contains 'Shared Documents|Budget.xlsx') 'Upload did not target the original folder.'
    Assert-True ($uploaded['scan'] -eq 'Exists' -and $script:SpFiles['/sites/IT/Shared Documents/scan.pdf'] -eq 1) 'Upload overwrote an existing SharePoint file.'
    Assert-True ($uploaded['bigdoc'] -eq 'FolderMissing') 'Upload created a folder Migration Manager did not.'
    $calls = $script:AddCalls.Count
    $uploaded = @{}; foreach ($result in Publish-DriveLargeFile @upload -CreateMissingFolders -Confirm:$false) { $uploaded[$result.Id] = $result.Status }
    Assert-True ($uploaded['sheet'] -eq 'AlreadyUploaded' -and $uploaded['bigdoc'] -eq 'Uploaded' -and $script:AddCalls.Count -eq $calls + 1) 'Upload resume repeated work or ignored -CreateMissingFolders.'
    & $module { $script:SharePointMaxFileBytes = 1 }
    $tooLarge = @(Publish-DriveLargeFile @upload -JournalPath (Join-Path $temp 'upload2.jsonl') -Confirm:$false)
    Assert-True (@($tooLarge | Where-Object Status -ne 'TooLarge').Count -eq 0) 'Oversized file was not refused.'
    [IO.File]::WriteAllText((Join-Path $staging 'stray.txt'), 'x')
    $failed = $false
    try { Publish-DriveLargeFile @upload -Confirm:$false | Out-Null } catch { $failed = $_.Exception.Message -like '*match no uploadable*' }
    Assert-True $failed 'An unplanned staged file was accepted for upload.'

    # Metadata: capture Google dates/sharing, then reapply to uploaded items and their folders.
    & $module {
        function script:Invoke-DriveApi {
            param($Path, $Method = 'Get', $Body)
            $id = ($Path -split '[/?]')[1]
            if ($Path -notmatch '/permissions\?') {
                if ($id -eq 'bigdoc') { return '{"createdTime":"2020-01-02T03:04:05.000Z","modifiedTime":"2021-05-06T07:08:09.000Z","owners":[{"emailAddress":"owner@g.com"}],"lastModifyingUser":{"emailAddress":"c@g.com"}}' | ConvertFrom-Json }
                return '{"createdTime":"2019-01-01T00:00:00.000Z"}' | ConvertFrom-Json
            }
            $json = switch ($id) {
                'test-root' { '{"permissions":[{"type":"user","role":"writer","emailAddress":"a@g.com"}]}' }
                'projects' { '{"permissions":[{"type":"user","role":"writer","emailAddress":"a@g.com"},{"type":"user","role":"reader","emailAddress":"b@g.com"}]}' }
                'bigdoc' {
                    '{"permissions":[{"type":"user","role":"writer","emailAddress":"a@g.com","permissionDetails":[{"inherited":true}]},' +
                    '{"type":"user","role":"owner","emailAddress":"owner@g.com","permissionDetails":[{"inherited":false}]},' +
                    '{"type":"user","role":"writer","emailAddress":"c@g.com","permissionDetails":[{"inherited":false}]},' +
                    '{"type":"anyone","role":"reader","permissionDetails":[{"inherited":false}]},' +
                    '{"type":"user","role":"reader","emailAddress":"x@ext.com","permissionDetails":[{"inherited":false}]}]}'
                }
                'sheet' { throw 'Google Drive request failed (HTTP 403). Check token, permissions, and connectivity; rerun using the same plan/journal.' }
                default { '{"permissions":[]}' }
            }
            $json | ConvertFrom-Json
        }
    }
    $metadata = Get-DriveItemMetadata -Plan $largePlan
    $meta = @{}; foreach ($item in $metadata.Items) { $meta[$item.Id] = $item }
    $direct = { param($Id) @($meta[$Id].Permissions | Where-Object Inherited -eq $false | ForEach-Object Identity) -join ',' }
    Assert-True ((& $direct projects) -eq 'b@g.com' -and (& $direct bigdoc) -eq 'owner@g.com,c@g.com,anyone,x@ext.com') 'Inherited Google sharing was not separated from direct sharing.'
    Assert-True ($meta['bigdoc'].Created -eq '2020-01-02T03:04:05.0000000Z' -and $meta['bigdoc'].Author -eq 'owner@g.com' -and $meta['bigdoc'].Editor -eq 'c@g.com') 'Dates or people were not captured in UTC.'
    Assert-True (!$meta['sheet'].PermissionsReadable -and $metadata.RootPermissions[0].Identity -eq 'a@g.com') 'Unreadable or root permissions were not recorded.'

    $script:SpFolders += '/sites/IT/Shared Documents/Projects'
    $script:ItemCalls = [Collections.Generic.List[string]]::new()
    $script:PermissionCalls = [Collections.Generic.List[string]]::new()
    function global:Get-PnPFolder {
        [CmdletBinding()] param($Url, $Includes)
        if ($Url -notin $script:SpFolders) { throw 'Folder not found.' }
        [pscustomobject]@{ ServerRelativeUrl = $Url; ListItemAllFields = [pscustomobject]@{ Id = 100 + $script:SpFolders.IndexOf($Url) } }
    }
    $script:Modified = @{}
    $script:IgnoreDates = $false
    function global:Get-PnPFile {
        [CmdletBinding()] param($Url, [switch]$AsListItem)
        $id = 200 + $Url.Length
        $modified = if ($script:Modified.ContainsKey($id)) { $script:Modified[$id] } else { [datetime]::UtcNow }
        if ($script:SpFiles.ContainsKey($Url)) { @{ File_x0020_Size = $script:SpFiles[$Url]; Id = $id; Modified = $modified } }
    }
    function global:New-PnPUser { [CmdletBinding()] param($LoginName) [pscustomobject]@{ LoginName = $LoginName } }
    function global:Set-PnPListItem {
        [CmdletBinding()] param($List, $Identity, [hashtable]$Values, $UpdateType)
        $script:ItemCalls.Add("$List|$Identity|$UpdateType|$($Values['Author'])|$($Values['Editor'])|$($Values['Created'].ToString('o'))")
        # Like SharePoint without override rights, optionally keep its own dates.
        if (!$script:IgnoreDates -and $Values.ContainsKey('Modified')) { $script:Modified[$Identity] = $Values['Modified'] }
    }
    function global:Invoke-PnPGraphMethod {
        [CmdletBinding()] param($Url, $Method, $Content)
        if ($Url -eq 'v1.0/sites/contoso.sharepoint.com:/sites/IT') { return [pscustomobject]@{ id = 'site-1' } }
        if ($Url -like 'v1.0/sites/site-1/drives*') {
            return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'other'; webUrl = 'https://contoso.sharepoint.com/sites/IT/Other' }, [pscustomobject]@{ id = 'drive-1'; webUrl = 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents' }) }
        }
        if ($Url -like 'v1.0/drives/drive-1/root:/*') { return [pscustomobject]@{ id = 'item:' + [uri]::UnescapeDataString($Url.Substring(26)) } }
        if ($Url -like 'v1.0/drives/drive-1/items/*/invite' -and $Method -eq 'Post') {
            $item = $Url.Substring(26, $Url.Length - 33)
            $script:PermissionCalls.Add("$item|$($Content.recipients[0].email)|$($Content.roles[0])|$($Content.requireSignIn)|$($Content.sendInvitation)|$($Content.retainInheritedPermissions)")
            return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'perm'; roles = $Content.roles }) }
        }
        throw "Unexpected Graph call $Method $Url"
    }
    $identityMap = Join-Path $temp 'identity-map.csv'
    "GoogleIdentity,Microsoft365Identity`nb@g.com,bee@m.com`n@g.com,@m.com" | Set-Content -LiteralPath $identityMap
    $metadataJournal = Join-Path $temp 'metadata.jsonl'
    $apply = @{ Plan = $largePlan; Metadata = $metadata; IdentityMapPath = $identityMap; UploadJournalPath = $uploadJournal; JournalPath = $metadataJournal; ApproveMetadata = $true }
    Set-DriveSharePointMetadata @apply -WhatIf | Out-Null
    Assert-True (!$script:ItemCalls.Count -and !$script:PermissionCalls.Count -and !(Test-Path $metadataJournal)) 'Metadata WhatIf wrote data.'
    $results = @(Set-DriveSharePointMetadata @apply -Confirm:$false)
    $byId = @{}; foreach ($result in $results) { $byId[$result.Id] = $result }
    Assert-True ($byId['test-root'].Status -eq 'RootPermissions' -and @($results | Where-Object Status -eq 'Applied' | ForEach-Object Id) -join ',' -eq 'projects,sheet,bigdoc') 'Metadata targets or order were wrong (folders must come first).'
    Assert-True ($script:PermissionCalls[0] -eq 'item:Projects|bee@m.com|read|True|False|True' -and $script:PermissionCalls[1] -eq 'item:Projects/Big Report.docx|c@m.com|write|True|False|True' -and $script:PermissionCalls.Count -eq 2) 'Direct sharing was not granted through Graph invites with sign-in, no email, and inherited access kept.'
    Assert-True (@($script:ItemCalls | Where-Object { $_ -like '*|UpdateOverwriteVersion|owner@m.com|c@m.com|2020-01-02T03:04:05.0000000Z' }).Count -eq 1) 'Dates and people were not applied without a new version.'
    Assert-True ($byId['bigdoc'].Detail -like '*anyone sharing*' -and $byId['bigdoc'].Detail -like '*x@ext.com unmapped*' -and $byId['sheet'].Detail -like '*unreadable*') 'Unmappable sharing was not reported.'
    $calls = $script:PermissionCalls.Count + $script:ItemCalls.Count
    $rerun = @(Set-DriveSharePointMetadata @apply -Confirm:$false)
    Assert-True (($script:PermissionCalls.Count + $script:ItemCalls.Count) -eq $calls -and @($rerun | Where-Object Status -eq 'AlreadyApplied').Count -eq 3) 'Metadata rerun repeated work.'
    $script:IgnoreDates = $true
    $script:Modified = @{}
    $ignored = @(Set-DriveSharePointMetadata @apply -JournalPath (Join-Path $temp 'metadata2.jsonl') -Confirm:$false)
    Assert-True (@($ignored | Where-Object { $_.Id -eq 'bigdoc' -and $_.Status -eq 'PartlyApplied' -and $_.Detail -like '*ignored*' }).Count -eq 1) 'Dates SharePoint ignored were reported as applied.'
    $other = $metadata.PSObject.Copy(); $other.PlanHash = 'other'
    $failed = $false
    try { Set-DriveSharePointMetadata @apply -Metadata $other -Confirm:$false | Out-Null } catch { $failed = $true }
    Assert-True $failed 'Metadata from another plan was accepted.'
} finally {
    $env:GDRIVE_ACCESS_TOKEN = $savedToken
    foreach ($name in 'Get-PnPConnection', 'Get-PnPFolder', 'Get-PnPFile', 'Add-PnPFile', 'Resolve-PnPFolder', 'New-PnPUser', 'Set-PnPListItem', 'Invoke-PnPGraphMethod') { Remove-Item -LiteralPath "Function:\$name" -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $temp -Recurse -Force
    Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force
}
# Wizard helpers: Migration Manager source path -> Google folder ID, and the per-failure summary.
$module = Get-Module DriveMigration
& $module {
    function script:Invoke-DriveApi {
        param($Path, $Method = 'Get', $Body)
        $text = [uri]::UnescapeDataString($Path)
        if ($text -like 'about*') { return [pscustomobject]@{ user = [pscustomobject]@{ emailAddress = 'a@x.com' } } }
        if ($text -like 'files/root*') { return [pscustomobject]@{ id = 'root-id' } }
        if ($text -like 'drives*') { return [pscustomobject]@{ drives = @([pscustomobject]@{ id = 'drv'; name = 'Finance' }, [pscustomobject]@{ id = 'other'; name = 'finance' }) } }
        $children = @{ 'root-id|Team' = 'team'; 'team|Sub' = 'sub'; "drv|Year's" = 'year' }
        $match = [regex]::Match($text, "^files\?q='([^']+)' in parents and name = '((?:[^'\\]|\\.)*)'")
        $key = $match.Groups[1].Value + '|' + ($match.Groups[2].Value -replace "\\'", "'")
        [pscustomobject]@{ files = @(if ($children.ContainsKey($key)) { [pscustomobject]@{ id = $children[$key]; name = $key.Split('|')[1] } }) }
    }
}
Assert-True ((Resolve-DriveSourceRoot 'a@x.com/Team/Sub') -eq 'sub' -and (Resolve-DriveSourceRoot "/Finance/Year's") -eq 'year') 'Migration Manager source path was not resolved to a Google folder.'
$failed = $false
try { Resolve-DriveSourceRoot 'b@x.com/Team' | Out-Null } catch { $failed = $_.Exception.Message -like '*Sign in as b@x.com*' }
Assert-True $failed 'Another user''s My Drive path was resolved with the wrong Google account.'
Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force

$summaryPlan = New-TestPlan @(
    (New-ItemFixture docs 'Docs' -Mime 'application/vnd.google-apps.folder'),
    (New-ItemFixture memo 'Memo' -Parent docs -Mime 'application/vnd.google-apps.document'),
    (New-ItemFixture clip 'clip.mp4' -Mime 'video/mp4'),
    (New-ItemFixture lost 'lost.pdf')
)
$runRows = @(
    [pscustomobject]@{ Step = 'Download'; Source = 'Google Drive: /Docs/Memo (ID memo)'; Result = 'Success'; Reason = 'Downloaded from Google (Export)' },
    [pscustomobject]@{ Step = 'ApplyMetadata'; Source = 'Google Drive: /Docs/Memo (ID memo)'; Result = 'Failure'; Reason = 'SharePoint ignored the Created/Modified values' },
    [pscustomobject]@{ Step = 'Upload'; Source = 'Google Drive: /Docs/Memo (ID memo)'; Result = 'Success'; Reason = 'Uploaded to SharePoint' },
    [pscustomobject]@{ Step = 'Download'; Source = 'Google Drive: /clip.mp4 (ID clip)'; Result = 'Failure'; Reason = 'Google download failed (HTTP 404)' },
    [pscustomobject]@{ Step = 'Download'; Source = 'Google Drive: /lost.pdf (ID lost)'; Result = 'Success'; Reason = 'Downloaded from Google (Media)' },
    [pscustomobject]@{ Step = 'Upload'; Source = 'Google Drive: /lost.pdf (ID lost)'; Result = 'Failure'; Reason = 'A file already exists at the destination; it was not overwritten' },
    [pscustomobject]@{ Step = 'Upload'; Source = 'Google Drive: /clip.mp4 (ID clip)'; Result = 'Preview'; Reason = '' }
)
$failures = @('/Docs/Memo', '/clip.mp4', '/lost.pdf', '/Docs', '/missing.txt') | ForEach-Object { [pscustomobject]@{ FullPath = $_; ResultCode = 'MJOBERROR'; FailureReason = 'Item level failure' } }
$final = @{}; foreach ($row in New-FailedFileSummary -Plan $summaryPlan -ReportFailures $failures -RunReports $runRows) { $final[$row.FullPath] = $row }
Assert-True ($final['/Docs/Memo'].Result -eq 'Success' -and $final['/Docs/Memo'].Reason -like 'Uploaded to SharePoint; metadata: Failure*' -and $final['/Docs/Memo'].Destination -eq 'https://contoso.sharepoint.com/sites/IT/Shared Documents/Docs/Memo.docx') 'Summary did not report the upload result with the metadata outcome.'
Assert-True ($final['/clip.mp4'].Reason -eq 'Download: Google download failed (HTTP 404)' -and $final['/lost.pdf'].Reason -like 'Upload: A file already exists*' -and $final['/lost.pdf'].Result -eq 'Failure') 'Summary did not report the step that failed.'
Assert-True ($final['/Docs'].Reason -like 'A folder failed*' -and $final['/missing.txt'].Reason -like 'Not found*' -and $final['/Docs'].MigrationManagerError -eq 'MJOBERROR: Item level failure') 'Summary mishandled folder or unknown failures.'

# Settings: the .secrets folder is created, the Entra ID saved, and a copied Google JSON found automatically.
$temp = Join-Path ([IO.Path]::GetTempPath()) ('drive-settings-test-' + [guid]::NewGuid())
try {
    [void](New-Item -Path (Join-Path $temp 'Downloads') -ItemType Directory -Force)
    $download = Join-Path $temp 'Downloads\client_secret_123.json'
    '{"installed":{}}' | Set-Content -LiteralPath $download
    $settingsFile = Join-Path $temp 'tool\.secrets\migration-settings.json'
    $saved = Save-MigrationSettings -SettingsPath $settingsFile -EntraClientId 'id-1' -GoogleClientSecretPath $download -CopyGoogleClient
    Assert-True ((Test-Path -LiteralPath (Join-Path $temp 'tool\.secrets\client_secret_123.json')) -and $saved.GoogleClientSecretPath -eq '' -and
        (Get-Content -LiteralPath $settingsFile -Raw | ConvertFrom-Json).EntraClientId -eq 'id-1') 'Settings folder, Entra ID, or copied Google JSON were not saved portably.'
    $kept = Save-MigrationSettings -SettingsPath (Join-Path $temp 'other\migration-settings.json') -EntraClientId 'id-2' -GoogleClientSecretPath $download
    '{"installed":{}}' | Set-Content -LiteralPath (Join-Path $temp 'tool\.secrets\client_secret_456.json')
    $ambiguous = Save-MigrationSettings -SettingsPath $settingsFile -EntraClientId 'id-1' -GoogleClientSecretPath (Join-Path $temp 'tool\.secrets\client_secret_123.json')
    Assert-True ($kept.GoogleClientSecretPath -eq $download -and $ambiguous.GoogleClientSecretPath -like '*client_secret_123.json') 'A Google JSON that cannot be found automatically was saved without its path.'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}

# Run report: the entry script writes one row per item with source, destination, result, and reason.
# Import mode is local-only, so the real script runs here without network calls.
$temp = Join-Path ([IO.Path]::GetTempPath()) ('drive-report-test-' + [guid]::NewGuid())
[void](New-Item -Path (Join-Path $temp 'inbox') -ItemType Directory -Force)
Push-Location $temp
try {
    $reportPlan = New-TestPlan @((New-ItemFixture docs 'Docs' -Mime 'application/vnd.google-apps.folder'), (New-ItemFixture memo 'Memo' -Parent docs -Mime 'application/vnd.google-apps.document'))
    $reportPlan | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath .\plan.json -Encoding utf8
    [IO.File]::WriteAllText((Join-Path $temp 'inbox\Memo.docx'), 'PK' + [char]3 + [char]4)
    [IO.File]::WriteAllText((Join-Path $temp 'inbox\stray.txt'), 'x')
    $entry = Join-Path $PSScriptRoot 'Copy-LargeDriveFiles.ps1'
    $summary = & $entry -Mode Import -PlanPath .\plan.json -StagingPath .\stage -InboxPath .\inbox -StatusPath .\report.csv -WarningAction SilentlyContinue
    $report = @(Import-Csv -LiteralPath .\report.csv)
    $memo = $report | Where-Object Status -eq 'Imported'
    $stray = $report | Where-Object Status -eq 'Unmatched'
    Assert-True ($memo.Source -eq 'Google Drive: /Docs/Memo (ID memo)' -and $memo.Destination -eq 'https://contoso.sharepoint.com/sites/IT/Shared Documents/Docs/Memo.docx' -and $memo.Result -eq 'Success') 'Run report lacked the source, destination, or success.'
    Assert-True ($stray.Result -eq 'Failure' -and $stray.Reason -eq 'No plan item matches this file name' -and $stray.Source -like 'Local file: *stray.txt' -and $summary.Succeeded -eq 1 -and $summary.Failed -eq 1) 'Run report lacked a failure reason.'
    # A run that stops partway still leaves a report, ending with why it stopped.
    [IO.File]::WriteAllText((Join-Path $temp 'blocker'), 'x')
    [IO.File]::WriteAllText((Join-Path $temp 'inbox\Memo.docx'), 'PK' + [char]3 + [char]4)
    $failed = $false
    try { & $entry -Mode Import -PlanPath .\plan.json -StagingPath .\blocker -InboxPath .\inbox -StatusPath .\stopped.csv -WarningAction SilentlyContinue | Out-Null } catch { $failed = $true }
    $stoppedReport = @(Import-Csv -LiteralPath .\stopped.csv)
    Assert-True ($failed -and $stoppedReport[-1].Status -eq 'Stopped' -and $stoppedReport[-1].Result -eq 'Failure' -and $stoppedReport[-1].Reason -like 'Run stopped before finishing:*') 'A stopped run left no report or hid the error.'
} finally {
    Pop-Location
    Remove-Item -LiteralPath $temp -Recurse -Force
}
"Passed $script:checks offline checks. No Google or Microsoft calls were made."