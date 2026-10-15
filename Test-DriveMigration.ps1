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
"Passed $script:checks offline checks. No Google or Microsoft calls were made."