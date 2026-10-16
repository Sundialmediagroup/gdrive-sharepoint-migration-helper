#Requires -Version 7.2
# Rebuilds files that Migration Manager could not copy (for example Google-native files over the
# export limit) into their original SharePoint folders, using an existing plan's TargetPath values.
#   Download:      Google -> local staging tree that mirrors the destination (drive.readonly token).
#   Import:        manual browser downloads -> the same staging tree, matched to plan items by name.
#   Upload:        staging tree -> SharePoint with PnP.PowerShell; never overwrites, journaled for resume.
#   Metadata:      Google dates, owner, last editor, and sharing for the plan tree -> local JSON + CSV.
#   ApplyMetadata: that file + an identity map -> uploaded SharePoint items and their folders (PnP).
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][ValidateSet('Download', 'Import', 'Upload', 'Metadata', 'ApplyMetadata')][string]$Mode,
    [Parameter(Mandatory)][string]$PlanPath,
    [string]$StagingPath,
    [string[]]$Id,
    [string[]]$MigrationReport,
    [string]$ReportSourcePath,
    [string]$InboxPath,
    [string]$JournalPath = '.\upload-journal.jsonl',
    [string]$StatusPath,
    # Download only: bytes this run may pull from Google (project egress limit is 1 TB per day).
    [ValidateRange(1, [long]::MaxValue)][long]$MaxDownloadBytes = 1000000000000,
    [switch]$ApproveUpload,
    [switch]$CreateMissingFolders,
    # Metadata modes: defaults to the plan path with a .metadata.json extension.
    [string]$MetadataPath,
    [string]$IdentityMapPath,
    [string]$MetadataJournalPath = '.\metadata-journal.jsonl',
    [switch]$SkipFolders,
    [switch]$ApproveMetadata
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force
# Resolve against the PowerShell location; [IO.Path]::GetFullPath ignores Set-Location.
$resolve = { param($Path) $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) }

$PlanPath = & $resolve $PlanPath
$plan = Get-Content -LiteralPath $PlanPath -Raw | ConvertFrom-Json
if ($plan.SchemaVersion -ne 1) { throw 'Unsupported plan schema.' }
$MetadataPath = if ($MetadataPath) { & $resolve $MetadataPath } else { [IO.Path]::ChangeExtension($PlanPath, '.metadata.json') }
if ($Mode -in @('Download', 'Import', 'Upload')) {
    if (!$StagingPath) { throw "$Mode requires -StagingPath." }
    $StagingPath = & $resolve $StagingPath
}
if (!$StatusPath) { $StatusPath = ".\large-files-$($Mode.ToLowerInvariant())-$([datetime]::Now.ToString('yyyyMMdd-HHmmss')).csv" }
$StatusPath = & $resolve $StatusPath
if (Test-Path -LiteralPath $StatusPath) { throw 'StatusPath already exists; choose a new one.' }

# Check inputs before any work, so a report is only written for runs that actually start.
$rows = @()
if ($Mode -in @('Download', 'Import')) {
    $selection = @{ Plan = $plan; Id = $Id }
    if ($MigrationReport) { $selection.ReportFailures = @(Import-MigrationReport -Path $MigrationReport -SourcePath $ReportSourcePath) }
    if ($Id -or $MigrationReport) {
        $rows = @(Select-DrivePlanRow @selection)
    } elseif ($Mode -eq 'Import') {
        $rows = @($plan.Rows | Where-Object MimeType -ne 'application/vnd.google-apps.folder')
    } else {
        throw 'Download requires -Id or -MigrationReport; it never downloads a whole tree.'
    }
}
if ($Mode -eq 'Import' -and !$InboxPath) { throw 'Import requires -InboxPath (the folder holding the browser downloads).' }
if ($Mode -eq 'Upload' -and !$ApproveUpload) { throw 'Upload requires -ApproveUpload after reviewing the staging folder.' }
if ($Mode -eq 'ApplyMetadata') {
    if (!$ApproveMetadata) { throw 'ApplyMetadata requires -ApproveMetadata after reviewing the metadata CSV and identity map.' }
    if (!$IdentityMapPath) { throw 'ApplyMetadata requires -IdentityMapPath (CSV: GoogleIdentity,Microsoft365Identity).' }
}
if ($Mode -in @('Upload', 'ApplyMetadata') -and !(Get-Command Add-PnPFile -ErrorAction SilentlyContinue)) {
    throw 'Install PnP.PowerShell (Install-Module PnP.PowerShell -Scope CurrentUser) and run Connect-PnPOnline for the destination site first.'
}
$permissionsPath = [IO.Path]::ChangeExtension($MetadataPath, '.permissions.csv')
if ($Mode -eq 'Metadata') {
    if ([IO.Path]::GetExtension($MetadataPath) -ne '.json') { throw 'MetadataPath must end in .json.' }
    if ((Test-Path -LiteralPath $MetadataPath) -or (Test-Path -LiteralPath $permissionsPath)) { throw 'Metadata file already exists. Choose a new -MetadataPath; keep the old one with its plan.' }
}

$common = @{ Plan = $plan; WhatIf = [bool]$WhatIfPreference }
if ($PSBoundParameters.ContainsKey('Confirm')) { $common.Confirm = [bool]$PSBoundParameters['Confirm'] }
$results = [Collections.Generic.List[object]]::new()
$stopped = $null
try {
    & {
        switch ($Mode) {
            'Metadata' {
                # Read-only on Google; writes the metadata JSON and a flat CSV for review.
                $metadata = Get-DriveItemMetadata -Plan $plan
                if ($PSCmdlet.ShouldProcess($MetadataPath, 'Write captured Google metadata; no Google changes')) {
                    $metadata | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $MetadataPath -Encoding utf8
                    $root = [pscustomobject]@{ Id = $plan.RootId; TargetPath = '(source root)'; Created = ''; Modified = ''; Author = ''; Editor = ''; PermissionsReadable = $metadata.RootPermissionsReadable; Permissions = $metadata.RootPermissions }
                    @(foreach ($item in @($root) + @($metadata.Items)) {
                        $permissions = @($item.Permissions)
                        if (!$permissions.Count) { $permissions = @([pscustomobject]@{ Type = ''; Role = ''; Identity = ''; Inherited = ''; Deleted = '' }) }
                        foreach ($permission in $permissions) {
                            [pscustomobject]@{
                                Id = $item.Id; TargetPath = $item.TargetPath; Created = $item.Created; Modified = $item.Modified; Author = $item.Author; Editor = $item.Editor
                                PermissionsReadable = $item.PermissionsReadable; Type = $permission.Type; Role = $permission.Role; Identity = $permission.Identity
                                Inherited = $permission.Inherited; Deleted = $permission.Deleted
                            }
                        }
                    }) | Export-Csv -LiteralPath $permissionsPath -NoTypeInformation -Encoding utf8
                }
                foreach ($item in $metadata.Items) {
                    $direct = @($item.Permissions | Where-Object { $_.Inherited -eq $false -and $_.Role -ne 'owner' }).Count
                    [pscustomobject]@{
                        Id = $item.Id; TargetPath = $item.TargetPath
                        Status = if ($item.PermissionsReadable) { 'Captured' } else { 'PermissionsUnreadable' }
                        Detail = if ($item.PermissionsReadable) { "$direct direct sharing grant(s)" } else { 'Google hid the sharing list; the account cannot share this item' }
                    }
                }
            }
            'Download' { Save-DriveLargeFile @common -StagingPath $StagingPath -Rows $rows -MaxDownloadBytes $MaxDownloadBytes }
            'Import' { Import-DriveManualDownload @common -StagingPath $StagingPath -Rows $rows -InboxPath (& $resolve $InboxPath) }
            'Upload' { Publish-DriveLargeFile @common -StagingPath $StagingPath -JournalPath (& $resolve $JournalPath) -ApproveUpload -CreateMissingFolders:$CreateMissingFolders }
            'ApplyMetadata' {
                $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
                Set-DriveSharePointMetadata @common -Metadata $metadata -IdentityMapPath (& $resolve $IdentityMapPath) -UploadJournalPath (& $resolve $JournalPath) `
                    -JournalPath (& $resolve $MetadataJournalPath) -ApproveMetadata -SkipFolders:$SkipFolders
            }
        }
    } | ForEach-Object { $results.Add($_) }
} catch {
    $stopped = $_
}

# Run report: one row per item with its Google source, SharePoint destination, result, and reason.
$runUtc = [datetime]::UtcNow.ToString('o')
$siteRoot = ([uri]$plan.DestinationUrl).GetLeftPart([UriPartial]::Authority)
$rowById = @{}
foreach ($row in $plan.Rows) { $rowById[$row.Id] = $row }
$sourcePath = @{ $plan.RootId = '' }
foreach ($row in $plan.Rows | Sort-Object Depth) { $sourcePath[$row.Id] = $sourcePath[$row.ParentId] + '/' + $row.OriginalName }
$successReasons = @{
    Captured = 'Google metadata captured'; Downloaded = 'Downloaded from Google'; Imported = 'Manual download copied into staging'
    Uploaded = 'Uploaded to SharePoint'; Applied = 'Dates, people, and sharing applied'
    AlreadyStaged = 'Already in staging from an earlier run'; AlreadyUploaded = 'Already uploaded in an earlier run'; AlreadyApplied = 'Already applied in an earlier run'
}
$failureReasons = @{
    Unmatched = 'No plan item matches this file name'; Missing = 'No manual download found for this item'
    Exists = 'A file already exists at the destination; it was not overwritten'; NotFound = 'Item not found in SharePoint'
    NoMetadata = 'Item missing from the metadata file'; Stopped = 'Run stopped'
}
$value = { param($Object, $Name) if ($Object.PSObject.Properties[$Name]) { [string]$Object.$Name } else { '' } }
$report = @(foreach ($result in $results) {
    $itemId = & $value $result 'Id'
    $file = & $value $result 'File'
    $targetPath = & $value $result 'TargetPath'
    $status = & $value $result 'Status'
    $detail = & $value $result 'Detail'
    $succeeded = $successReasons.ContainsKey($status)
    $reason = if ($succeeded) {
        $text = $successReasons[$status]
        if (& $value $result 'Method') { $text += " ($(& $value $result 'Method'))" }
        if (& $value $result 'Granted') { $text += "; shared with $(& $value $result 'Granted')" }
        if ($detail) { $text += "; note: $detail" }
        $text
    } elseif ($detail) { $detail } elseif ($failureReasons.ContainsKey($status)) { $failureReasons[$status] } else { $status }
    [pscustomobject]@{
        RunUtc = $runUtc; Step = $Mode
        Source = if ($rowById.ContainsKey($itemId)) { "Google Drive: $($sourcePath[$itemId]) (ID $itemId)" } elseif ($itemId -eq $plan.RootId) { "Google Drive: source root (ID $itemId)" } elseif ($file) { "Local file: $file" } else { '' }
        Destination = if ($targetPath.StartsWith('/')) { $siteRoot + $targetPath } else { '' }
        Result = if ($status -eq 'WhatIf') { 'Preview' } elseif ($succeeded) { 'Success' } else { 'Failure' }
        Reason = $reason; Status = $status; LocalPath = & $value $result 'LocalPath'
    }
})
if ($stopped) {
    $report += [pscustomobject]@{
        RunUtc = $runUtc; Step = $Mode; Source = ''; Destination = ''; Result = 'Failure'
        Reason = "Run stopped before finishing: $($stopped.Exception.Message)"; Status = 'Stopped'; LocalPath = ''
    }
}
foreach ($row in $report | Where-Object Result -eq 'Failure') { Write-Warning "$($row.Status): $($row.Destination) $($row.Reason)" }
$written = !$WhatIfPreference -and $report.Count
if ($written) { $report | Export-Csv -LiteralPath $StatusPath -NoTypeInformation -Encoding utf8 }
if ($stopped) {
    if ($written) { Write-Warning "Run report written to $StatusPath." }
    throw $stopped
}
[pscustomobject]@{
    Mode = $Mode; Items = $report.Count
    Succeeded = @($report | Where-Object Result -eq 'Success').Count; Failed = @($report | Where-Object Result -eq 'Failure').Count
    Summary = (@($report | Group-Object Status | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ')
    ReportPath = if ($written) { $StatusPath } else { '' }
    MetadataPath = if ($Mode -eq 'Metadata' -and !$WhatIfPreference) { $MetadataPath } else { '' }
}
