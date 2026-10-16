#Requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:FolderMime = 'application/vnd.google-apps.folder'
$script:Exports = @{
    'application/vnd.google-apps.document' = '.docx'
    'application/vnd.google-apps.spreadsheet' = '.xlsx'
    'application/vnd.google-apps.presentation' = '.pptx'
    'application/vnd.google-apps.drawing' = '.png'
}

function Get-NameParts {
    param([string]$Name, [string]$MimeType)
    if ($script:Exports.ContainsKey($MimeType)) {
        $extension = $script:Exports[$MimeType]
        $stem = $Name
        if ($stem.EndsWith($extension, [StringComparison]::OrdinalIgnoreCase)) {
            $stem = $stem.Substring(0, $stem.Length - $extension.Length)
        }
        return @($stem, $extension)
    }
    if ($MimeType -eq $script:FolderMime) { return @($Name, '') }
    $dot = $Name.LastIndexOf('.')
    if ($dot -gt 0) { return @($Name.Substring(0, $dot), $Name.Substring($dot)) }
    return @($Name, '')
}

function Get-ShortHash {
    param([string]$Id)
    $bytes = [Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($Id))
    return [Convert]::ToHexString($bytes).Substring(0, 10).ToLowerInvariant()
}

function Get-SafeName {
    param([string]$Name, [string]$MimeType, [string]$Id, [int]$Limit, [switch]$Suffix)
    $clean = $Name.Normalize([Text.NormalizationForm]::FormC)
    $clean = [regex]::Replace($clean, '["*:<>?/\\|\x00-\x1f\x7f]', '_').Trim().TrimEnd('.')
    $clean = [regex]::Replace($clean, '(?i)_vti_', '_vti-')
    if ($clean -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(?:\.|$)' -or
        $clean -in @('.lock', 'desktop.ini', 'forms') -or $clean.StartsWith('~$') -or
        ($clean.Length -gt 0 -and [int]$clean[0] -in @(0x309b, 0x1027))) {
        $clean = "_$clean"
    }
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = 'item' }
    $parts = Get-NameParts $clean $MimeType
    $stem, $extension = $parts
    $tag = ''
    if ($Suffix -or ($stem.Length + $extension.Length) -gt $Limit) { $tag = '~' + (Get-ShortHash $Id) }
    $budget = $Limit - $extension.Length - $tag.Length
    if ($budget -lt 1) { throw "Name budget too small for item $Id; shorten the destination or flatten the source hierarchy." }
    if ($stem.Length -gt $budget) {
        $stem = $stem.Substring(0, $budget)
        if ([char]::IsHighSurrogate($stem[$stem.Length - 1])) { $stem = $stem.Substring(0, $stem.Length - 1) }
    }
    $stem = $stem.TrimEnd(' ', '.')
    if (!$stem) { $stem = '_' }
    $target = $stem + $tag + $extension
    $source = $target
    if ($script:Exports.ContainsKey($MimeType)) { $source = $stem + $tag }
    return [pscustomobject]@{ Source = $source; Target = $target }
}

function Get-KeptName {
    param([string]$Name, [string]$MimeType)
    $stem, $extension = Get-NameParts $Name $MimeType
    return [pscustomobject]@{ Source = $Name; Target = $stem + $extension }
}

function New-DriveRenamePlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items,
        [Parameter(Mandatory)][string]$RootId,
        [Parameter(Mandatory)][string]$DestinationUrl,
        [ValidateRange(80, 400)][int]$MaxPathLength = 400,
        [ValidateRange(24, 255)][int]$MaxNameLength = 120,
        # Migration Manager failure rows (FullPath, ResultCode, FailureReason). When supplied, only
        # reported items, sibling collisions, and path-budget ancestors are renamed.
        [AllowEmptyCollection()][object[]]$ReportFailures
    )
    $uri = [uri]$DestinationUrl
    if (!$uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.Query -or $uri.Fragment) {
        throw 'Destination must be an HTTPS library/folder URL, not a sharing link or browser view URL.'
    }
    $prefix = [uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
    if (!$prefix -or $prefix -match '/Forms/[^/]+\.aspx$') { throw 'Provide the actual library/folder URL.' }
    $byId = @{}
    $children = @{}
    foreach ($item in $Items) {
        if ($byId.ContainsKey($item.id) -or $item.id -eq $RootId) { throw 'Duplicate ID or cycle in inventory.' }
        if (@($item.parents).Count -ne 1) { throw "Item $($item.id) must have exactly one parent." }
        $byId[$item.id] = $item
        $parent = [string]$item.parents[0]
        if (!$children.ContainsKey($parent)) { $children[$parent] = [Collections.Generic.List[object]]::new() }
        $children[$parent].Add($item)
    }
    $ordered = [Collections.Generic.List[object]]::new()
    $queue = [Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($RootId)
    $depth = @{ $RootId = 0 }
    while ($queue.Count) {
        $parent = $queue.Dequeue()
        if (!$children.ContainsKey($parent)) { continue }
        foreach ($item in ($children[$parent] | Sort-Object id)) {
            $depth[$item.id] = $depth[$parent] + 1
            $ordered.Add($item)
            if ($item.mimeType -eq $script:FolderMime) { $queue.Enqueue($item.id) }
        }
    }
    if ($ordered.Count -ne $Items.Count) { throw 'Inventory has unreachable items, cycles, or a non-folder parent.' }
    $limits = @{}
    foreach ($item in $ordered) { $limits[$item.id] = $MaxNameLength }
    $targeted = $PSBoundParameters.ContainsKey('ReportFailures')
    $fixIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $keptLength = @{}
    $findings = @()
    if ($targeted) {
        $relative = @{ $RootId = '' }
        foreach ($item in $ordered) { $relative[$item.id] = $relative[[string]$item.parents[0]] + '/' + $item.name }
        $findings = @(foreach ($failure in $ReportFailures) {
            $found = @($ordered | Where-Object { $relative[$_.id] -ceq $failure.FullPath })
            if (!$found.Count) { $found = @($ordered | Where-Object { $relative[$_.id] -eq $failure.FullPath }) }
            if ($found.Count -eq 1) { [void]$fixIds.Add($found[0].id) }
            [pscustomobject]@{
                FullPath = $failure.FullPath; ResultCode = $failure.ResultCode; FailureReason = $failure.FailureReason
                Id = if ($found.Count -eq 1) { $found[0].id } else { '' }
                Status = if (!$found.Count) { 'NotFound' } elseif ($found.Count -gt 1) { 'Ambiguous' } else { 'Pending' }
            }
        })
        # Unreported items keep their names unless the path budget forces a shorter limit.
        foreach ($item in $ordered) {
            $keptLength[$item.id] = (Get-KeptName $item.name $item.mimeType).Target.Length
            if (!$fixIds.Contains($item.id)) { $limits[$item.id] = $keptLength[$item.id] }
        }
    }
    $rows = @()
    for ($pass = 0; $pass -le ($MaxNameLength * [Math]::Max(1, $Items.Count)); $pass++) {
        $paths = @{ $RootId = $prefix }
        $used = @{}
        $rows = @(foreach ($item in $ordered) {
            $parent = [string]$item.parents[0]
            if (!$used.ContainsKey($parent)) { $used[$parent] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
            $kept = $targeted -and !$fixIds.Contains($item.id) -and $limits[$item.id] -ge $keptLength[$item.id]
            $safe = if ($kept) { Get-KeptName $item.name $item.mimeType } else { Get-SafeName $item.name $item.mimeType $item.id $limits[$item.id] }
            $warnings = [Collections.Generic.List[string]]::new()
            if (!$used[$parent].Add($safe.Target)) {
                if ($kept) {
                    # Migration Manager keeps both and appends "(n)"; project that length instead of renaming.
                    $mime = if ($item.mimeType -eq $script:FolderMime) { $script:FolderMime } else { '' }
                    $stem, $extension = Get-NameParts $safe.Target $mime
                    for ($n = 1; !$used[$parent].Add("$stem($n)$extension"); $n++) {}
                    $safe = [pscustomobject]@{ Source = $safe.Source; Target = "$stem($n)$extension" }
                    $warnings.Add('Name collision left to Migration Manager, which adds a (n) suffix to one item.')
                } else {
                    $safe = Get-SafeName $item.name $item.mimeType $item.id $limits[$item.id] -Suffix
                    if (!$used[$parent].Add($safe.Target)) { throw "Unresolvable name collision at $($item.id)." }
                }
            }
            $paths[$item.id] = $paths[$parent] + '/' + $safe.Target
            if ($item.mimeType -in @('application/vnd.google-apps.shortcut', 'application/vnd.google-apps.site', 'application/vnd.google-apps.map')) {
                $warnings.Add('Unsupported by Migration Manager; arrange separate disposition.')
            } elseif ($item.mimeType -eq 'application/vnd.google-apps.form') {
                $warnings.Add('Requires a Microsoft Forms destination; path projection is not applicable to Forms.')
            } elseif ($item.mimeType.StartsWith('application/vnd.google-apps.') -and
                $item.mimeType -ne $script:FolderMime -and !$script:Exports.ContainsKey($item.mimeType)) {
                $warnings.Add('Google-native type requires manual conversion/support review.')
            }
            if ($safe.Source -cne $item.name -and !$item.capabilities.canRename) { $warnings.Add('BLOCKED: source account cannot rename this item.') }
            [pscustomobject]@{
                Id = $item.id; ParentId = $parent; MimeType = $item.mimeType
                OriginalName = $item.name; NewName = $safe.Source; TargetName = $safe.Target
                TargetPath = $paths[$item.id]; PathLength = $paths[$item.id].Length
                Version = [string]$item.version; Depth = $depth[$item.id]
                Rename = ($safe.Source -cne $item.name); Warnings = ($warnings -join ' ')
            }
        })
        $tooLong = @($rows | Where-Object PathLength -gt $MaxPathLength | Sort-Object PathLength -Descending)
        if (!$tooLong.Count) { break }
        $leaf = $tooLong[0]
        $chain = [Collections.Generic.List[string]]::new()
        $current = $leaf.Id
        while ($current -ne $RootId) { $chain.Add($current); $current = [string]$byId[$current].parents[0] }
        $candidate = $chain | Where-Object { $limits[$_] -gt 24 } | Sort-Object { $limits[$_] } -Descending | Select-Object -First 1
        if (!$candidate) { throw "Hierarchy cannot fit the path budget at $($leaf.Id). Use a shorter destination or split into subfolder migration tasks; no plan was produced." }
        $limits[$candidate] = [Math]::Max(24, $limits[$candidate] - [Math]::Max(1, $leaf.PathLength - $MaxPathLength))
    }
    # A reported item is Fixed when it or any ancestor is renamed; otherwise naming is not the cause.
    $rowById = @{}
    foreach ($row in $rows) { $rowById[$row.Id] = $row }
    foreach ($finding in $findings | Where-Object Status -eq 'Pending') {
        $finding.Status = 'NotFixed'
        for ($current = $finding.Id; $current -ne $RootId; $current = $rowById[$current].ParentId) {
            if ($rowById[$current].Rename) { $finding.Status = 'Fixed'; break }
        }
    }
    return [pscustomobject]@{
        SchemaVersion = 1; CreatedUtc = [datetime]::UtcNow.ToString('o'); RootId = $RootId
        DestinationUrl = $DestinationUrl; MaxPathLength = $MaxPathLength; MaxNameLength = $MaxNameLength
        Rows = @($rows); ReportFindings = @($findings)
    }
}

# Google Drive usage limits: https://developers.google.com/workspace/drive/api/guides/limits
$script:MaxAttempts = 8

function Test-DriveRetryable {
    # Quota errors are "403: User rate limit exceeded" (reason rateLimitExceeded/userRateLimitExceeded) or 429.
    param([int]$Status, [string]$Message)
    return $Status -in @(429, 500, 502, 503, 504) -or ($Status -eq 403 -and $Message -match 'rate ?limit ?exceeded')
}

function Wait-DriveBackoff {
    # Google's recommended wait: min(2^n seconds + up to 1,000 ms random, 64 seconds).
    param([int]$Attempt)
    [Threading.Thread]::Sleep([int][Math]::Min([Math]::Pow(2, $Attempt) * 1000 + (Get-Random -Maximum 1001), 64000))
}

function Invoke-DriveApi {
    param([string]$Path, [ValidateSet('Get', 'Patch')][string]$Method = 'Get', [object]$Body)
    if (!$env:GDRIVE_ACCESS_TOKEN) { throw 'Set GDRIVE_ACCESS_TOKEN to a short-lived Google OAuth access token. Never store it in the plan.' }
    for ($attempt = 0; $attempt -lt $script:MaxAttempts; $attempt++) {
        try {
            $request = @{
                Uri = 'https://www.googleapis.com/drive/v3/' + $Path
                Method = $Method; Headers = @{ Authorization = 'Bearer ' + $env:GDRIVE_ACCESS_TOKEN }
                TimeoutSec = 120; ErrorAction = 'Stop'
            }
            if ($null -ne $Body) { $request.Body = ConvertTo-Json $Body -Compress; $request.ContentType = 'application/json; charset=utf-8' }
            return Invoke-RestMethod @request
        } catch {
            $status = 0
            if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            $message = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { '' }
            if ($attempt -eq $script:MaxAttempts - 1 -or !(Test-DriveRetryable $status $message)) {
                throw "Google Drive request failed (HTTP $status). Check token, permissions, and connectivity; rerun using the same plan/journal."
            }
            Wait-DriveBackoff $attempt
        }
    }
}

function Get-DriveInventory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RootId)
    if ($RootId -eq 'root') { throw 'Use the actual My Drive folder ID, not the root alias. Resolve it with files/root?fields=id first.' }
    $root = Invoke-DriveApi ('files/' + [uri]::EscapeDataString($RootId) + '?supportsAllDrives=true&fields=id,mimeType,driveId')
    if ($root.mimeType -ne $script:FolderMime) { throw 'RootId must identify a folder, My Drive root, or shared-drive root.' }
    $RootId = $root.id
    $corpus = '&corpora=user'
    if ($root.PSObject.Properties['driveId'] -and $root.driveId) {
        $corpus = '&corpora=drive&driveId=' + [uri]::EscapeDataString($root.driveId)
    }
    $queue = [Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($RootId)
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    [void]$seen.Add($RootId)
    while ($queue.Count) {
        $parent = $queue.Dequeue()
        $page = ''
        do {
            $query = [uri]::EscapeDataString("'$($parent.Replace("'", "\'"))' in parents and trashed = false")
            $fields = [uri]::EscapeDataString('nextPageToken,incompleteSearch,files(id,name,mimeType,parents,version,capabilities(canRename))')
            $url = "files?q=$query&fields=$fields&pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true" + $corpus
            if ($page) { $url += '&pageToken=' + [uri]::EscapeDataString($page) }
            $result = Invoke-DriveApi $url
            if ($result.PSObject.Properties['incompleteSearch'] -and $result.incompleteSearch) { throw 'Google returned an incomplete inventory. No safe plan can be made.' }
            foreach ($item in $result.files) {
                if (!$seen.Add($item.id)) { throw "Repeated item or cycle: $($item.id)" }
                $item
                if ($item.mimeType -eq $script:FolderMime) { $queue.Enqueue($item.id) }
            }
            $page = if ($result.PSObject.Properties['nextPageToken']) { $result.nextPageToken } else { '' }
        } while ($page)
    }
}

function Write-DriveJournal {
    param([string]$Path, [object]$Entry)
    $line = (ConvertTo-Json $Entry -Compress -Depth 8) + [Environment]::NewLine
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($line)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally { $stream.Dispose() }
}

function Get-PlanFingerprint {
    param([object]$Plan)
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $Plan -Depth 20 -Compress))))
}

function Read-DriveJournal {
    param([string]$Path, [string]$Fingerprint)
    if (!(Test-Path -LiteralPath $Path)) { return }
    $entries = @(Get-Content -LiteralPath $Path | Where-Object { $_.Trim() } | ForEach-Object { ConvertFrom-Json $_ })
    if (@($entries | Where-Object PlanHash -ne $Fingerprint).Count) { throw 'Journal belongs to another plan. Use the original plan or a new journal.' }
    $entries
}

function Invoke-DriveRenamePlan {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][string]$JournalPath,
        [Parameter(Mandatory)][switch]$ApproveSourceRenames,
        [switch]$Rollback
    )
    if (!$ApproveSourceRenames) { throw 'Explicit source rename approval is required.' }
    if ($Plan.SchemaVersion -ne 1) { throw 'Unsupported plan schema.' }
    if (@($Plan.Rows | Where-Object Warnings -like '*BLOCKED:*').Count -and !$Rollback) { throw 'Plan contains blocked renames. Resolve permissions and re-plan.' }
    $journalFullPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($JournalPath)
    if (!(Test-Path -LiteralPath (Split-Path $journalFullPath -Parent))) { throw 'Journal parent directory must exist.' }
    $fingerprint = Get-PlanFingerprint $Plan
    $entries = @(Read-DriveJournal $journalFullPath $fingerprint)
    if (@($entries | Where-Object Operation -notin @('Apply', 'Rollback')).Count) { throw 'Journal is not a rename journal. Use a separate journal per operation type.' }
    $attempted = @{}
    foreach ($entry in $entries) { if ($entry.Operation -eq 'Apply' -and $entry.State -eq 'Intent') { $attempted[$entry.Id] = $true } }
    $live = @(Get-DriveInventory $Plan.RootId)
    $liveById = @{}
    foreach ($item in $live) { $liveById[$item.id] = $item }
    if ($live.Count -ne @($Plan.Rows).Count) { throw 'Source tree changed. Freeze changes and create a new plan before applying, or restore the original tree before rollback.' }
    foreach ($row in $Plan.Rows) {
        $item = $liveById[$row.Id]
        if (!$item -or @($item.parents).Count -ne 1 -or $item.parents[0] -cne $row.ParentId -or $item.mimeType -cne $row.MimeType) { throw "Source structure changed at $($row.Id)." }
        $isOriginal = $item.name -ceq $row.OriginalName
        $isOurs = $attempted.ContainsKey($row.Id) -and $item.name -ceq $row.NewName
        if (!$isOriginal -and !$isOurs) { throw "Source name changed at $($row.Id). No writes performed during preflight." }
        if (!$Rollback -and !$attempted.ContainsKey($row.Id) -and [string]$item.version -cne $row.Version) { throw "Source version changed at $($row.Id). Re-plan." }
    }
    $changes = @($Plan.Rows | Where-Object Rename | Sort-Object Depth, Id)
    if ($Rollback) { $changes = @($changes | Where-Object { $attempted.ContainsKey($_.Id) } | Sort-Object Depth, Id -Descending) }
    if (!$changes.Count) { return }
    $operation = if ($Rollback) { 'Rollback' } else { 'Apply' }
    if (!$PSCmdlet.ShouldProcess("Google source root $($Plan.RootId)", "$operation $($changes.Count) planned renames; journal $journalFullPath")) { return }
    foreach ($row in $changes) {
        $desired = if ($Rollback) { $row.OriginalName } else { $row.NewName }
        $expected = if ($Rollback) { $row.NewName } else { $row.OriginalName }
        $path = 'files/' + [uri]::EscapeDataString($row.Id) + '?supportsAllDrives=true&fields=id,name,parents,version,capabilities(canRename)'
        $current = Invoke-DriveApi $path
        if ($current.name -ceq $desired) { continue }
        if ($current.name -cne $expected -or @($current.parents).Count -ne 1 -or $current.parents[0] -cne $row.ParentId -or !$current.capabilities.canRename) {
            throw "Concurrent change or lost rename permission at $($row.Id). Stopped; keep the journal for recovery."
        }
        $entry = [ordered]@{ PlanHash = $fingerprint; Operation = $operation; State = 'Intent'; Id = $row.Id; OldName = $current.name; NewName = $desired; Utc = [datetime]::UtcNow.ToString('o') }
        Write-DriveJournal $journalFullPath $entry
        $updated = Invoke-DriveApi $path -Method Patch -Body @{ name = $desired }
        if ($updated.name -cne $desired) { throw "Rename response mismatch at $($row.Id). Stopped." }
        $entry.State = 'Done'
        Write-DriveJournal $journalFullPath $entry
        [pscustomobject]@{ Id = $row.Id; Operation = $operation; Name = $desired }
    }
}

function Import-MigrationReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Path, [string]$SourcePath)
    # Migration Manager's ProjectError CSV; FullPath is relative to the reported source root.
    $failures = @(foreach ($file in $Path) {
        $file = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($file)
        if ((Get-Content -LiteralPath $file -TotalCount 1) -notmatch '(^|,)"?FullPath"?(,|$)') {
            throw "$file has no FullPath column. Use the Migration Manager error report (ProjectError.csv), not MigrationSummary.csv."
        }
        Import-Csv -LiteralPath $file
    })
    if (!$SourcePath) {
        $sources = @($failures | ForEach-Object SourcePath | Sort-Object -Unique)
        if ($sources.Count -gt 1) { throw 'The report covers several sources; pass -ReportSourcePath.' }
        $SourcePath = $sources | Select-Object -First 1
    }
    $matched = @($failures | Where-Object SourcePath -eq $SourcePath)
    if ($failures.Count -and !$matched.Count) { throw "No report rows for source '$SourcePath'. Check -ReportSourcePath." }
    $matched
}

# Large files: Google download or manual-download import into a local staging tree that mirrors
# each plan row's TargetPath, then upload to SharePoint with PnP.PowerShell.
$script:ExportMimes = @{
    'application/vnd.google-apps.document' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    'application/vnd.google-apps.spreadsheet' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    'application/vnd.google-apps.presentation' = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    'application/vnd.google-apps.drawing' = 'image/png'
}
$script:SharePointMaxFileBytes = 250GB
$script:SharePointMaxPathLength = 400
$script:HttpClient = $null

function Get-PlanPrefix {
    param([object]$Plan)
    return [uri]::UnescapeDataString(([uri]$Plan.DestinationUrl).AbsolutePath).TrimEnd('/')
}

function Get-TargetProblem {
    # Returns why SharePoint would reject the row's decoded target path, or '' when it looks acceptable.
    param([object]$Plan, [object]$Row)
    $prefix = Get-PlanPrefix $Plan
    if (!$Row.TargetPath.StartsWith($prefix + '/', [StringComparison]::Ordinal)) { return 'Target path is outside the plan destination.' }
    if ($Row.TargetPath.Length -gt $script:SharePointMaxPathLength) { return "Target path is $($Row.TargetPath.Length) characters; SharePoint allows $script:SharePointMaxPathLength." }
    $segments = $Row.TargetPath.Substring($prefix.Length + 1).Split('/')
    if ($segments.Count -ne [int]$Row.Depth) { return "A name in this path contains '/'; place this file manually." }
    foreach ($segment in $segments) {
        if (!$segment -or $segment -cne $segment.Trim() -or $segment.EndsWith('.') -or $segment -match '["*:<>?\\|\x00-\x1f\x7f]' -or
            $segment -match '(?i)_vti_' -or $segment -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\.|$)' -or
            $segment -in @('.lock', 'desktop.ini') -or $segment.StartsWith('~$')) {
            return "'$segment' is not a valid SharePoint name; Migration Manager may have renamed it. Place this file manually."
        }
    }
    return ''
}

function Get-StagedPath {
    param([object]$Plan, [object]$Row, [string]$StagingPath)
    $relative = $Row.TargetPath.Substring((Get-PlanPrefix $Plan).Length + 1)
    return [IO.Path]::Combine([string[]](@($StagingPath) + $relative.Split('/')))
}

function Test-FileSignature {
    # Rejects HTML error pages and truncated files saved under an Office or PNG name.
    param([string]$Path, [string]$Name)
    $signatures = @{ '.docx' = 'PK'; '.xlsx' = 'PK'; '.pptx' = 'PK'; '.png' = [char]0x89 + 'PNG' }
    $extension = [IO.Path]::GetExtension($Name).ToLowerInvariant()
    if (!$signatures.ContainsKey($extension)) { return $true }
    $bytes = [byte[]]::new($signatures[$extension].Length)
    $stream = [IO.File]::OpenRead($Path)
    try { $read = $stream.Read($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
    return $read -eq $bytes.Length -and [Text.Encoding]::Latin1.GetString($bytes) -ceq $signatures[$extension]
}

function Get-MatchKey {
    param([string]$Name)
    return [regex]::Replace($Name.Normalize([Text.NormalizationForm]::FormC), '["*:<>?/\\|\x00-\x1f\x7f]', '_').Trim().TrimEnd('.').ToLowerInvariant()
}

function Get-PlanPathIndex {
    # Report paths use the names Migration Manager saw: original names, or new names after Apply.
    param([object]$Plan)
    $original = @{ $Plan.RootId = '' }
    $renamed = @{ $Plan.RootId = '' }
    foreach ($row in $Plan.Rows | Sort-Object Depth) {
        $original[$row.Id] = $original[$row.ParentId] + '/' + $row.OriginalName
        $renamed[$row.Id] = $renamed[$row.ParentId] + '/' + $row.NewName
    }
    [pscustomobject]@{ Original = $original; Renamed = $renamed }
}

function Find-PlanPath {
    # Plan rows (files or folders) at a report FullPath: exact case first, then case-insensitive.
    param([object]$Plan, [object]$Index, [string]$FullPath)
    $found = @($Plan.Rows | Where-Object { $Index.Original[$_.Id] -ceq $FullPath -or $Index.Renamed[$_.Id] -ceq $FullPath })
    if (!$found.Count) { $found = @($Plan.Rows | Where-Object { $Index.Original[$_.Id] -eq $FullPath -or $Index.Renamed[$_.Id] -eq $FullPath }) }
    $found
}

function Select-DrivePlanRow {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Plan, [string[]]$Id, [AllowEmptyCollection()][object[]]$ReportFailures)
    $byId = @{}
    foreach ($row in $Plan.Rows) { $byId[$row.Id] = $row }
    $selected = [ordered]@{}
    foreach ($itemId in $Id) {
        if (!$byId.ContainsKey($itemId)) { throw "Id $itemId is not in the plan." }
        if ($byId[$itemId].MimeType -eq $script:FolderMime) { throw "Id $itemId is a folder; select files only." }
        $selected[$itemId] = $byId[$itemId]
    }
    if ($ReportFailures) {
        $index = Get-PlanPathIndex $Plan
        foreach ($failure in $ReportFailures) {
            $found = @(Find-PlanPath $Plan $index $failure.FullPath)
            if ($found.Count -ne 1) { Write-Warning "$(if ($found.Count) { 'Ambiguous' } else { 'NotFound' }): report path $($failure.FullPath)"; continue }
            if ($found[0].MimeType -eq $script:FolderMime) { Write-Warning "Skipped folder in report: $($failure.FullPath)"; continue }
            $selected[$found[0].Id] = $found[0]
        }
    }
    @($selected.Values)
}

function Invoke-DriveDownload {
    param([string]$Uri, [string]$OutFile)
    if (!$env:GDRIVE_ACCESS_TOKEN) { throw 'Set GDRIVE_ACCESS_TOKEN to a short-lived Google OAuth access token. Never store it in the plan.' }
    $target = [uri]$Uri
    if ($target.Scheme -ne 'https' -or $target.Host -notmatch '(^|\.)(googleapis|google)\.com$') { throw "Refusing to send the Google token to $($target.Host)." }
    if (!$script:HttpClient) {
        $script:HttpClient = [Net.Http.HttpClient]::new()
        $script:HttpClient.Timeout = [Threading.Timeout]::InfiniteTimeSpan
    }
    for ($attempt = 0; $attempt -lt $script:MaxAttempts; $attempt++) {
        $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Get, $target)
        $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $env:GDRIVE_ACCESS_TOKEN)
        $status = 0
        $message = ''
        try {
            # Streams to disk; .NET drops the Authorization header on cross-host redirects.
            $response = $script:HttpClient.SendAsync($request, [Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
            try {
                $status = [int]$response.StatusCode
                if ($response.IsSuccessStatusCode) {
                    $file = [IO.File]::Create($OutFile)
                    try { $response.Content.ReadAsStream().CopyTo($file); $file.Flush($true) } finally { $file.Dispose() }
                    return
                }
                $message = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            } finally { $response.Dispose() }
        } catch [Net.Http.HttpRequestException], [IO.IOException] {
            $status = 0
        } finally { $request.Dispose() }
        if ($attempt -eq $script:MaxAttempts - 1 -or ($status -ne 0 -and !(Test-DriveRetryable $status $message))) {
            $reason = [regex]::Match($message, '"reason"\s*:\s*"([A-Za-z]+)"').Groups[1].Value
            throw "Google download failed (HTTP $status $reason)."
        }
        Wait-DriveBackoff $attempt
    }
}

function Save-DriveLargeFile {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$StagingPath,
        # Google allows 1 TB of egress per project per day before charges apply; this run stays under it.
        [ValidateRange(1, [long]::MaxValue)][long]$MaxDownloadBytes = 1000000000000
    )
    $staging = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($StagingPath)
    $downloadedBytes = 0L
    $deferred = 'Download budget for this run reached (Google: 1 TB per project per day); rerun after 24 hours. Staged files are skipped.'
    foreach ($row in $Rows) {
        $result = [pscustomobject]@{ Id = $row.Id; TargetPath = $row.TargetPath; LocalPath = ''; Method = ''; Status = ''; Detail = '' }
        $native = $row.MimeType.StartsWith('application/vnd.google-apps.')
        $result.Detail = Get-TargetProblem $Plan $row
        if ($result.Detail) {
            $result.Status = 'InvalidTarget'
        } elseif ($row.Warnings -like '*(n) suffix*') {
            $result.Status = 'CollisionReview'
            $result.Detail = 'Migration Manager chose which item got the (n) suffix; place this file manually.'
        } elseif ($native -and !$script:ExportMimes.ContainsKey($row.MimeType)) {
            $result.Status = 'ManualReview'
            $result.Detail = 'No Office export exists for this Google type.'
        } else {
            $result.LocalPath = Get-StagedPath $Plan $row $staging
            if (Test-Path -LiteralPath $result.LocalPath) {
                $result.Status = 'AlreadyStaged'
            } elseif ($downloadedBytes -ge $MaxDownloadBytes) {
                $result.Status = 'Deferred'
                $result.Detail = $deferred
            } elseif (!$PSCmdlet.ShouldProcess($result.LocalPath, "Download Google item $($row.Id)")) {
                $result.Status = 'WhatIf'
            } else {
                $folder = [IO.Path]::GetDirectoryName($result.LocalPath)
                $partial = [IO.Path]::Combine($folder, '~download-' + [guid]::NewGuid().ToString('N') + '.partial')
                try {
                    $escaped = [uri]::EscapeDataString($row.Id)
                    $live = Invoke-DriveApi "files/$escaped`?supportsAllDrives=true&fields=id,name,mimeType,parents,trashed,size"
                    $trashed = $live.PSObject.Properties['trashed'] -and $live.trashed
                    if ($trashed -or $live.mimeType -cne $row.MimeType -or @($live.parents).Count -ne 1 -or $live.parents[0] -cne $row.ParentId -or $live.name -cne $row.NewName) {
                        throw 'Source item was moved, renamed, or trashed, or the plan renames were not applied.'
                    }
                    # Google-native files report no size, so exports are counted only after download.
                    $size = if ($live.PSObject.Properties['size']) { [long]$live.size } else { 0L }
                    if ($downloadedBytes + $size -gt $MaxDownloadBytes) {
                        $result.Status = 'Deferred'
                        $result.Detail = $deferred
                        $result
                        continue
                    }
                    [void][IO.Directory]::CreateDirectory($folder)
                    if ($native) {
                        $mime = $script:ExportMimes[$row.MimeType]
                        try {
                            Invoke-DriveDownload "https://www.googleapis.com/drive/v3/files/$escaped/export?mimeType=$([uri]::EscapeDataString($mime))" $partial
                            $result.Method = 'Export'
                        } catch {
                            # files.export stops at 10 MB; try the item's own export link before giving up.
                            if ($_.Exception.Message -notmatch 'HTTP 403') { throw }
                            $links = Invoke-DriveApi "files/$escaped`?supportsAllDrives=true&fields=exportLinks"
                            $link = if ($links.PSObject.Properties['exportLinks'] -and $links.exportLinks.PSObject.Properties[$mime]) { $links.exportLinks.$mime } else { '' }
                            if (!$link) { throw "Export failed ($($_.Exception.Message)) and Google offered no export link." }
                            Invoke-DriveDownload $link $partial
                            $result.Method = 'ExportLink'
                        }
                    } else {
                        Invoke-DriveDownload "https://www.googleapis.com/drive/v3/files/$escaped`?alt=media&supportsAllDrives=true" $partial
                        $result.Method = 'Media'
                    }
                    if (!(Test-FileSignature $partial $row.TargetName)) { throw 'Downloaded content is not the expected file type.' }
                    $downloadedBytes += [IO.FileInfo]::new($partial).Length
                    [IO.File]::Move($partial, $result.LocalPath)
                    $result.Status = 'Downloaded'
                } catch {
                    if ($_.Exception.Message -match 'HTTP 401|GDRIVE_ACCESS_TOKEN') { throw }
                    if ([IO.File]::Exists($partial)) { [IO.File]::Delete($partial) }
                    $result.Status = 'ManualDownload'
                    $result.Detail = $_.Exception.Message
                }
            }
        }
        $result
    }
}

function Import-DriveManualDownload {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$InboxPath,
        [Parameter(Mandatory)][string]$StagingPath
    )
    $inbox = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($InboxPath)
    $staging = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($StagingPath)
    if (!(Test-Path -LiteralPath $inbox -PathType Container)) { throw "Inbox folder $inbox not found." }
    # Browser downloads use the Google name plus the export extension, with characters Windows rejects replaced.
    $rowsByKey = @{}
    $eligible = [Collections.Generic.List[object]]::new()
    foreach ($row in $Rows) {
        $problem = Get-TargetProblem $Plan $row
        if ($row.Warnings -like '*(n) suffix*') { $problem = 'Migration Manager chose which item got the (n) suffix; place this file manually.' }
        if ($problem) {
            [pscustomobject]@{ File = ''; Id = $row.Id; TargetPath = $row.TargetPath; LocalPath = ''; Status = 'ManualReview'; Detail = $problem }
            continue
        }
        $eligible.Add($row)
        $keys = [Collections.Generic.HashSet[string]]::new()
        foreach ($name in @($row.TargetName, $row.NewName, $row.OriginalName)) {
            $stem, $extension = Get-NameParts $name $row.MimeType
            [void]$keys.Add((Get-MatchKey ($stem + $extension)))
        }
        foreach ($key in $keys) {
            if (!$rowsByKey.ContainsKey($key)) { $rowsByKey[$key] = [Collections.Generic.List[object]]::new() }
            $rowsByKey[$key].Add($row)
        }
    }
    $files = @(Get-ChildItem -LiteralPath $inbox -File -Recurse | Sort-Object FullName)
    $found = @{}
    $claims = @{}
    foreach ($file in $files) {
        $key = Get-MatchKey $file.Name
        if (!$rowsByKey.ContainsKey($key)) {
            # Browsers add " (1)" when the Downloads folder already holds that name.
            $key = Get-MatchKey ([regex]::Replace($file.Name, ' \(\d+\)(?=\.[^.]*$|$)', ''))
        }
        $candidates = @()
        if ($rowsByKey.ContainsKey($key)) { $candidates = @($rowsByKey[$key]) }
        $found[$file.FullName] = $candidates
        if ($candidates.Count -eq 1) { $claims[$candidates[0].Id] = 1 + [int]$claims[$candidates[0].Id] }
    }
    foreach ($file in $files) {
        $candidates = @($found[$file.FullName])
        $result = [pscustomobject]@{ File = $file.FullName; Id = ''; TargetPath = ''; LocalPath = ''; Status = ''; Detail = '' }
        if (!$candidates.Count) {
            $result.Status = 'Unmatched'
        } elseif ($candidates.Count -gt 1) {
            $result.Status = 'Ambiguous'
            $result.Detail = "Matches $($candidates.Count) plan items; copy it to the right staging path by hand."
        } else {
            $row = $candidates[0]
            $result.Id = $row.Id
            $result.TargetPath = $row.TargetPath
            $result.LocalPath = Get-StagedPath $Plan $row $staging
            if ($claims[$row.Id] -gt 1) {
                $result.Status = 'Ambiguous'
                $result.Detail = 'Several downloads match this item; copy the right one by hand.'
            } elseif (Test-Path -LiteralPath $result.LocalPath) {
                $result.Status = 'AlreadyStaged'
            } elseif ([IO.Path]::GetExtension($file.Name) -ne [IO.Path]::GetExtension($row.TargetName) -or !(Test-FileSignature $file.FullName $row.TargetName)) {
                $result.Status = 'WrongType'
                $result.Detail = "Expected a valid $([IO.Path]::GetExtension($row.TargetName)) file."
            } elseif ($PSCmdlet.ShouldProcess($result.LocalPath, "Copy $($file.Name) into staging")) {
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($result.LocalPath))
                [IO.File]::Copy($file.FullName, $result.LocalPath, $false)
                $result.Status = 'Imported'
            } else {
                $result.Status = 'WhatIf'
            }
        }
        $result
    }
    foreach ($row in $eligible) {
        $local = Get-StagedPath $Plan $row $staging
        if (!$claims.ContainsKey($row.Id) -and !(Test-Path -LiteralPath $local)) {
            [pscustomobject]@{ File = ''; Id = $row.Id; TargetPath = $row.TargetPath; LocalPath = $local; Status = 'Missing'; Detail = 'No download found; save it to LocalPath by hand or rerun Import.' }
        }
    }
}

function Get-SharePointFileItem {
    # Get-PnPFile -AsListItem returns nothing for a missing file, but throws "NotFound" when its folder is missing too.
    param([string]$Url)
    try { return Get-PnPFile -Url $Url -AsListItem -ErrorAction Stop }
    catch {
        if ($_.Exception.Message -match 'NotFound|File Not Found|does not exist|-2147024894') { return $null }
        throw
    }
}

function Test-UploadedSize {
    # SharePoint writes document properties into Office files on upload, so their size changes (pilot:
    # 4,937-byte .xlsx stored as 12,357). For those, require a non-empty file: an interrupted chunked
    # upload leaves a 0-byte placeholder. Every other type must match exactly.
    param([long]$StoredBytes, [long]$LocalBytes, [string]$Name)
    $office = @('.docx', '.docm', '.dotx', '.xlsx', '.xlsm', '.xltx', '.pptx', '.pptm', '.potx', '.doc', '.xls', '.ppt')
    if ([IO.Path]::GetExtension($Name).ToLowerInvariant() -in $office) { return $StoredBytes -gt 0 -or $LocalBytes -eq 0 }
    return $StoredBytes -eq $LocalBytes
}

function Publish-DriveLargeFile {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][string]$StagingPath,
        [Parameter(Mandatory)][string]$JournalPath,
        [Parameter(Mandatory)][switch]$ApproveUpload,
        [switch]$CreateMissingFolders
    )
    if (!$ApproveUpload) { throw 'Explicit upload approval is required.' }
    if ($Plan.SchemaVersion -ne 1) { throw 'Unsupported plan schema.' }
    $staging = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($StagingPath)
    $journalFullPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($JournalPath)
    if (!(Test-Path -LiteralPath $staging -PathType Container)) { throw "Staging folder $staging not found." }
    if (!(Test-Path -LiteralPath (Split-Path $journalFullPath -Parent))) { throw 'Journal parent directory must exist.' }
    $connection = Get-PnPConnection -ErrorAction Stop
    $web = [uri]::UnescapeDataString(([uri]$connection.Url).AbsolutePath).TrimEnd('/')
    if (!(Get-PlanPrefix $Plan).StartsWith($web + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Connected site $($connection.Url) does not contain the plan destination. Connect-PnPOnline to the destination site first."
    }
    $rowsByPath = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Plan.Rows | Where-Object MimeType -ne $script:FolderMime) {
        if (Get-TargetProblem $Plan $row) { continue }
        $path = Get-StagedPath $Plan $row $staging
        if ($rowsByPath.ContainsKey($path)) { throw "Two plan items map to $path." }
        $rowsByPath[$path] = $row
    }
    $staged = @(Get-ChildItem -LiteralPath $staging -File -Recurse -Force | Sort-Object FullName)
    $unknown = @($staged | Where-Object { !$rowsByPath.ContainsKey($_.FullName) })
    if ($unknown.Count) { throw "Staging holds $($unknown.Count) file(s) that match no uploadable plan item, for example $($unknown[0].FullName). Remove them; nothing was uploaded." }
    $fingerprint = Get-PlanFingerprint $Plan
    $entries = @(Read-DriveJournal $journalFullPath $fingerprint)
    if (@($entries | Where-Object Operation -ne 'Upload').Count) { throw 'Journal is not an upload journal. Use a separate journal per operation type.' }
    $done = @{}
    $intended = @{}
    foreach ($entry in $entries) { if ($entry.State -eq 'Done') { $done[$entry.Id] = $true } else { $intended[$entry.Id] = [long]$entry.Bytes } }
    foreach ($file in $staged) {
        $row = $rowsByPath[$file.FullName]
        $folderPath = $row.TargetPath.Substring(0, $row.TargetPath.LastIndexOf('/'))
        $result = [pscustomobject]@{ Id = $row.Id; TargetPath = $row.TargetPath; LocalPath = $file.FullName; Bytes = $file.Length; Status = ''; Detail = '' }
        if ($file.Length -gt $script:SharePointMaxFileBytes) {
            $result.Status = 'TooLarge'
            $result.Detail = 'SharePoint accepts files up to 250 GB.'
        } elseif ($done.ContainsKey($row.Id)) {
            $result.Status = 'AlreadyUploaded'
        } elseif ($existing = Get-SharePointFileItem $row.TargetPath) {
            if (!$intended.ContainsKey($row.Id)) {
                $result.Status = 'Exists'
                $result.Detail = 'A file already exists at the target; it was not overwritten.'
            } elseif ((Test-UploadedSize ([long]$existing['File_x0020_Size']) $file.Length $row.TargetName)) {
                # The upload finished but its response was lost: complete the journal instead of uploading again.
                if (!$WhatIfPreference) { Write-DriveJournal $journalFullPath ([ordered]@{ PlanHash = $fingerprint; Operation = 'Upload'; State = 'Done'; Id = $row.Id; TargetPath = $row.TargetPath; Bytes = $file.Length; Sha256 = ''; Utc = [datetime]::UtcNow.ToString('o') }) }
                $result.Status = 'AlreadyUploaded'
            } else {
                $result.Status = 'PartialUpload'
                $result.Detail = 'An earlier upload in this journal stopped partway. Check the SharePoint file, delete it if incomplete, then rerun.'
            }
        } else {
            $folderExists = $false
            try { $folderExists = [bool](Get-PnPFolder -Url $folderPath -ErrorAction Stop) } catch { $folderExists = $false }
            if (!$folderExists -and !$CreateMissingFolders) {
                $result.Status = 'FolderMissing'
                $result.Detail = 'Destination folder not found. Check where Migration Manager put it, or pass -CreateMissingFolders.'
            } elseif (!$PSCmdlet.ShouldProcess($row.TargetPath, "Upload $($file.Length) bytes to SharePoint")) {
                $result.Status = 'WhatIf'
            } else {
                $entry = [ordered]@{
                    PlanHash = $fingerprint; Operation = 'Upload'; State = 'Intent'; Id = $row.Id; TargetPath = $row.TargetPath
                    Bytes = $file.Length; Sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash; Utc = [datetime]::UtcNow.ToString('o')
                }
                Write-DriveJournal $journalFullPath $entry
                # Add-PnPFile does not reliably create missing folders (PnP 3.4.1 throws NotFound), so create them explicitly.
                if (!$folderExists) { Resolve-PnPFolder -SiteRelativePath $folderPath.Substring($web.Length + 1) -ErrorAction Stop | Out-Null }
                # Add-PnPFile overwrites silently, so the existence check above guards it. Above 10 MB it uploads in chunks.
                Add-PnPFile -Path $file.FullName -Folder $folderPath.Substring($web.Length + 1) -NewFileName $row.TargetName -ErrorAction Stop | Out-Null
                $uploaded = Get-SharePointFileItem $row.TargetPath
                if (!$uploaded -or !(Test-UploadedSize ([long]$uploaded['File_x0020_Size']) $file.Length $row.TargetName)) { throw "Upload of $($row.Id) could not be verified. Stopped; keep the journal." }
                $entry.State = 'Done'
                $entry.Utc = [datetime]::UtcNow.ToString('o')
                Write-DriveJournal $journalFullPath $entry
                $result.Status = 'Uploaded'
            }
        }
        $result
    }
}

# Metadata: capture Google dates, people, and sharing, then reapply them to uploaded SharePoint items.
function Get-OptionalValue {
    param([object]$Object, [string]$Name)
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
}

function ConvertTo-UtcText {
    # Invoke-RestMethod turns ISO timestamps into DateTime; store them as unambiguous UTC text.
    param([object]$Value)
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o') }
    return [string]$Value
}

function ConvertTo-PermissionRecord {
    param([object]$Permission)
    $type = [string](Get-OptionalValue $Permission 'type')
    $identity = switch ($type) {
        'domain' { [string](Get-OptionalValue $Permission 'domain') }
        'anyone' { 'anyone' }
        default { [string](Get-OptionalValue $Permission 'emailAddress') }
    }
    $details = @(Get-OptionalValue $Permission 'permissionDetails' | Where-Object { $null -ne $_ })
    $inherited = $null
    if ($details.Count) { $inherited = !@($details | Where-Object { !(Get-OptionalValue $_ 'inherited') }).Count }
    [pscustomobject]@{
        Type = $type; Role = [string](Get-OptionalValue $Permission 'role'); Identity = $identity.ToLowerInvariant()
        Deleted = [bool](Get-OptionalValue $Permission 'deleted'); Inherited = $inherited
    }
}

function Get-DriveItemMetadata {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Plan)
    if ($Plan.SchemaVersion -ne 1) { throw 'Unsupported plan schema.' }
    $files = @{}
    $records = @{}
    $readable = @{}
    foreach ($id in @($Plan.RootId) + @($Plan.Rows | ForEach-Object Id)) {
        $escaped = [uri]::EscapeDataString($id)
        $files[$id] = Invoke-DriveApi "files/$escaped`?supportsAllDrives=true&fields=id,createdTime,modifiedTime,owners(emailAddress),lastModifyingUser(emailAddress)"
        $records[$id] = [Collections.Generic.List[object]]::new()
        $readable[$id] = $true
        try {
            $page = ''
            do {
                $fields = [uri]::EscapeDataString('nextPageToken,permissions(type,role,emailAddress,domain,deleted,permissionDetails(inherited))')
                $url = "files/$escaped/permissions?supportsAllDrives=true&pageSize=100&fields=$fields"
                if ($page) { $url += '&pageToken=' + [uri]::EscapeDataString($page) }
                $result = Invoke-DriveApi $url
                foreach ($permission in @(Get-OptionalValue $result 'permissions')) {
                    if ($null -ne $permission) { $records[$id].Add((ConvertTo-PermissionRecord $permission)) }
                }
                $page = [string](Get-OptionalValue $result 'nextPageToken')
            } while ($page)
        } catch {
            # Google hides the permission list from accounts that cannot share the item; record that instead of guessing.
            if ($_.Exception.Message -notmatch 'HTTP 40[34]') { throw }
            $readable[$id] = $false
            $records[$id].Clear()
        }
    }
    $keys = @{}
    foreach ($id in @($Plan.RootId) + @($Plan.Rows | ForEach-Object Id)) {
        $keys[$id] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($record in $records[$id]) { [void]$keys[$id].Add("$($record.Type)|$($record.Role)|$($record.Identity)") }
    }
    foreach ($record in $records[$Plan.RootId]) { $record.Inherited = $false }
    foreach ($row in $Plan.Rows) {
        # My Drive items carry no inheritance details: a grant the parent also has is treated as inherited.
        # When the parent's list is unreadable the origin stays unknown (null) and Apply skips it.
        foreach ($record in $records[$row.Id] | Where-Object { $null -eq $_.Inherited }) {
            if ($readable[$row.ParentId]) { $record.Inherited = $keys[$row.ParentId].Contains("$($record.Type)|$($record.Role)|$($record.Identity)") }
        }
    }
    [pscustomobject]@{
        SchemaVersion = 1; CreatedUtc = [datetime]::UtcNow.ToString('o'); PlanHash = Get-PlanFingerprint $Plan; RootId = $Plan.RootId
        RootPermissionsReadable = $readable[$Plan.RootId]; RootPermissions = @($records[$Plan.RootId])
        Items = @(foreach ($row in $Plan.Rows) {
            $file = $files[$row.Id]
            [pscustomobject]@{
                Id = $row.Id; TargetPath = $row.TargetPath; IsFolder = ($row.MimeType -eq $script:FolderMime)
                Created = ConvertTo-UtcText (Get-OptionalValue $file 'createdTime')
                Modified = ConvertTo-UtcText (Get-OptionalValue $file 'modifiedTime')
                Author = [string](Get-OptionalValue (Get-OptionalValue $file 'owners' | Select-Object -First 1) 'emailAddress')
                Editor = [string](Get-OptionalValue (Get-OptionalValue $file 'lastModifyingUser') 'emailAddress')
                PermissionsReadable = $readable[$row.Id]; Permissions = @($records[$row.Id])
            }
        })
    }
}

function Read-IdentityMap {
    # CSV with GoogleIdentity,Microsoft365Identity. A row like "@google-domain.com,@tenant.com" maps a whole domain.
    param([string]$Path)
    $map = @{}
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -and (!$rows[0].PSObject.Properties['GoogleIdentity'] -or !$rows[0].PSObject.Properties['Microsoft365Identity'])) {
        throw 'Identity map needs the columns GoogleIdentity and Microsoft365Identity.'
    }
    foreach ($row in $rows) {
        $google = $row.GoogleIdentity.Trim().ToLowerInvariant()
        $microsoft = $row.Microsoft365Identity.Trim()
        if (!$google -or !$microsoft) { continue }
        if ($google.StartsWith('@') -ne $microsoft.StartsWith('@')) { throw "Domain rows must map @domain to @domain: $google." }
        if ($map.ContainsKey($google)) { throw "Identity $google is mapped twice." }
        $map[$google] = $microsoft
    }
    $map
}

function Resolve-MappedIdentity {
    param([hashtable]$Map, [string]$Identity)
    $key = $Identity.ToLowerInvariant()
    if (!$key) { return '' }
    if ($Map.ContainsKey($key)) { return $Map[$key] }
    $at = $key.LastIndexOf('@')
    if ($at -gt 0 -and $Map.ContainsKey($key.Substring($at))) { return $key.Substring(0, $at) + $Map[$key.Substring($at)] }
    return ''
}

function Get-GraphDriveId {
    # The Graph drive behind a document library, matched by its URL.
    param([string]$SiteHost, [string]$Web, [string]$Library)
    $site = Invoke-PnPGraphMethod -Url "v1.0/sites/$SiteHost`:$(if ($Web) { $Web } else { '/' })" -Method Get -ErrorAction Stop
    $drives = Invoke-PnPGraphMethod -Url "v1.0/sites/$($site.id)/drives?`$select=id,webUrl" -Method Get -ErrorAction Stop
    $wanted = "$Web/$Library"
    $match = @(@($drives.value) | Where-Object { [uri]::UnescapeDataString(([uri]$_.webUrl).AbsolutePath).TrimEnd('/') -eq $wanted })
    if ($match.Count -ne 1) { throw "Could not find the Graph drive for library $wanted." }
    return $match[0].id
}

function Grant-DriveItemAccess {
    # Direct access like SharePoint's Share dialog: sign-in required, no invitation email, existing access kept.
    param([string]$DriveId, [string]$ItemId, [string]$User, [ValidateSet('read', 'write')][string]$Role)
    $body = @{ recipients = @(@{ email = $User }); roles = @($Role); requireSignIn = $true; sendInvitation = $false; retainInheritedPermissions = $true }
    $response = Invoke-PnPGraphMethod -Url "v1.0/drives/$DriveId/items/$ItemId/invite" -Method Post -Content $body -ErrorAction Stop
    $failed = @(@(Get-OptionalValue $response 'value') | Where-Object { $null -ne $_ -and (Get-OptionalValue $_ 'error') })
    if ($failed.Count) { throw "SharePoint refused to share with ${User}: $(Get-OptionalValue $failed[0].error 'message')" }
}

function Set-DriveSharePointMetadata {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][object]$Metadata,
        [Parameter(Mandatory)][string]$IdentityMapPath,
        [Parameter(Mandatory)][string]$UploadJournalPath,
        [Parameter(Mandatory)][string]$JournalPath,
        [Parameter(Mandatory)][switch]$ApproveMetadata,
        [switch]$SkipFolders
    )
    if (!$ApproveMetadata) { throw 'Explicit metadata approval is required.' }
    if ($Plan.SchemaVersion -ne 1 -or $Metadata.SchemaVersion -ne 1) { throw 'Unsupported plan or metadata schema.' }
    $fingerprint = Get-PlanFingerprint $Plan
    if ($Metadata.PlanHash -ne $fingerprint) { throw 'Metadata file was captured for another plan.' }
    $map = Read-IdentityMap ($PSCmdlet.GetUnresolvedProviderPathFromPSPath($IdentityMapPath))
    $journalFullPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($JournalPath)
    if (!(Test-Path -LiteralPath (Split-Path $journalFullPath -Parent))) { throw 'Journal parent directory must exist.' }
    $connection = Get-PnPConnection -ErrorAction Stop
    $web = [uri]::UnescapeDataString(([uri]$connection.Url).AbsolutePath).TrimEnd('/')
    $prefix = Get-PlanPrefix $Plan
    if (!$prefix.StartsWith($web + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Connected site $($connection.Url) does not contain the plan destination. Connect-PnPOnline to the destination site first."
    }
    # A document library is a top-level folder of its site, so the first segment names the list.
    $library = $prefix.Substring($web.Length + 1).Split('/')[0]
    $siteHost = ([uri]$connection.Url).Host
    $driveId = $null
    $uploads = @(Read-DriveJournal ($PSCmdlet.GetUnresolvedProviderPathFromPSPath($UploadJournalPath)) $fingerprint)
    if (@($uploads | Where-Object Operation -ne 'Upload').Count) { throw 'UploadJournalPath is not an upload journal.' }
    $entries = @(Read-DriveJournal $journalFullPath $fingerprint)
    if (@($entries | Where-Object Operation -ne 'Metadata').Count) { throw 'Journal is not a metadata journal. Use a separate journal per operation type.' }
    $applied = @{}
    foreach ($entry in $entries) { $applied[$entry.Id] = $true }
    $rowById = @{}
    foreach ($row in $Plan.Rows) { $rowById[$row.Id] = $row }
    $metaById = @{}
    foreach ($item in $Metadata.Items) { $metaById[$item.Id] = $item }
    # Only items this tool uploaded, plus their folders so folder-level sharing flows down by inheritance.
    $targets = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $uploads | Where-Object State -eq 'Done') {
        for ($current = [string]$entry.Id; $current -and $current -ne $Plan.RootId -and $rowById.ContainsKey($current); $current = $rowById[$current].ParentId) {
            if ($rowById[$current].MimeType -ne $script:FolderMime -or !$SkipFolders) { [void]$targets.Add($current) }
        }
    }
    $ordered = @($targets | ForEach-Object { $rowById[$_] } | Sort-Object @{ Expression = { $_.MimeType -ne $script:FolderMime } }, Depth, Id)
    $rootGrants = @($Metadata.RootPermissions | Where-Object { $_.Type -in @('user', 'group') -and !$_.Deleted -and $_.Role -ne 'owner' })
    if ($rootGrants.Count) {
        [pscustomobject]@{
            Id = $Plan.RootId; TargetPath = $prefix; Status = 'RootPermissions'; Granted = ''
            Detail = "$($rootGrants.Count) Google grant(s) on the source root or shared drive. Give them access at the destination library, folder, or site membership; not applied automatically."
        }
    }
    $unique = @($ordered | Where-Object { $meta = $metaById[$_.Id]; $meta -and @($meta.Permissions | Where-Object { $_.Inherited -eq $false -and $_.Role -ne 'owner' }).Count }).Count
    if ($unique -gt 5000) { Write-Warning "$unique items would get unique permissions; Microsoft recommends at most 5,000 per library. Consider simplifying sharing first." }
    $ensured = @{}
    foreach ($row in $ordered) {
        $meta = $metaById[$row.Id]
        $isFolder = $row.MimeType -eq $script:FolderMime
        $result = [pscustomobject]@{ Id = $row.Id; TargetPath = $row.TargetPath; Status = ''; Granted = ''; Detail = '' }
        if ($applied.ContainsKey($row.Id)) { $result.Status = 'AlreadyApplied'; $result; continue }
        if (!$meta) { $result.Status = 'NoMetadata'; $result.Detail = 'Item missing from the metadata file.'; $result; continue }
        $itemId = $null
        if ($isFolder) {
            try { $itemId = (Get-PnPFolder -Url $row.TargetPath -Includes ListItemAllFields -ErrorAction Stop).ListItemAllFields.Id } catch { $itemId = $null }
        } else {
            $item = Get-SharePointFileItem $row.TargetPath
            if ($item) { $itemId = $item.Id }
        }
        if (!$itemId) { $result.Status = 'NotFound'; $result.Detail = 'Item not found in SharePoint.'; $result; continue }
        $notes = [Collections.Generic.List[string]]::new()
        $values = @{}
        if ($meta.Created) { $values.Created = ([datetime]$meta.Created).ToUniversalTime() }
        if ($meta.Modified) { $values.Modified = ([datetime]$meta.Modified).ToUniversalTime() }
        foreach ($field in @(@('Author', $meta.Author), @('Editor', $meta.Editor))) {
            if (!$field[1]) { continue }
            $mapped = Resolve-MappedIdentity $map $field[1]
            if ($mapped) { $values[$field[0]] = $mapped } else { $notes.Add("$($field[0]) $($field[1]) unmapped") }
        }
        if (!$meta.PermissionsReadable) { $notes.Add('Google permissions unreadable; sharing not applied') }
        $grants = [ordered]@{}
        foreach ($permission in $meta.Permissions) {
            if ($null -eq $permission.Inherited) { $notes.Add("$($permission.Identity): origin unknown, not applied"); continue }
            # Google lists every My Drive item's owner as a direct grant; ownership maps to Author, not sharing.
            if ($permission.Inherited -or $permission.Role -eq 'owner') { continue }
            if ($permission.Type -notin @('user', 'group')) { $notes.Add("$($permission.Type) sharing ($($permission.Role)) has no SharePoint equivalent"); continue }
            if ($permission.Deleted) { $notes.Add("$($permission.Identity): deleted account"); continue }
            $mapped = Resolve-MappedIdentity $map $permission.Identity
            if (!$mapped) { $notes.Add("$($permission.Identity) unmapped"); continue }
            $role = if ($permission.Role -in @('reader', 'commenter')) { 'read' } else { 'write' }
            if (!$grants.Contains($mapped) -or $role -eq 'write') { $grants[$mapped] = $role }
        }
        $result.Granted = (@($grants.Keys | ForEach-Object { "$_=$($grants[$_])" }) -join '; ')
        $result.Detail = $notes -join '; '
        if (!$PSCmdlet.ShouldProcess($row.TargetPath, "Set $($values.Count) date/people field(s) and grant $($grants.Count) permission(s)")) { $result.Status = 'WhatIf'; $result; continue }
        foreach ($user in @($values['Author'], $values['Editor']) | Where-Object { $_ -and !$ensured.ContainsKey($_) }) {
            New-PnPUser -LoginName $user -ErrorAction Stop | Out-Null
            $ensured[$user] = $true
        }
        if ($grants.Count) {
            # Graph sharing invites need only edit rights (delegated Files.ReadWrite.All), not Full Control.
            if (!$driveId) { $driveId = Get-GraphDriveId $siteHost $web $library }
            $relative = $row.TargetPath.Substring($web.Length + $library.Length + 2)
            $encoded = ($relative.Split('/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            $driveItem = Invoke-PnPGraphMethod -Url "v1.0/drives/$driveId/root:/$encoded" -Method Get -ErrorAction Stop
            foreach ($user in $grants.Keys) { Grant-DriveItemAccess $driveId $driveItem.id $user $grants[$user] }
        }
        $datesIgnored = $false
        if ($values.Count) {
            Set-PnPListItem -List $library -Identity $itemId -Values $values -UpdateType UpdateOverwriteVersion -ErrorAction Stop | Out-Null
            # Pilot: with Edit rights and AllSites.Write, SharePoint accepted the call but kept its own dates. Check.
            if (!$isFolder -and $values.ContainsKey('Modified')) {
                $after = Get-SharePointFileItem $row.TargetPath
                $stored = if ($after) { [datetime]$after['Modified'] } else { $null }
                # CSOM may return UTC values without a UTC kind, so accept either reading.
                $datesIgnored = !$stored -or ([Math]::Abs(($stored - $values.Modified).TotalMinutes) -gt 1 -and [Math]::Abs(($stored.ToUniversalTime() - $values.Modified).TotalMinutes) -gt 1)
                if ($datesIgnored) { $notes.Add('SharePoint ignored the Created/Modified/Author/Editor values (overriding them needs more than Edit rights); the originals stay in the metadata file') }
            }
        }
        Write-DriveJournal $journalFullPath ([ordered]@{ PlanHash = $fingerprint; Operation = 'Metadata'; State = 'Done'; Id = $row.Id; TargetPath = $row.TargetPath; Granted = $result.Granted; Utc = [datetime]::UtcNow.ToString('o') })
        $result.Detail = $notes -join '; '
        $result.Status = if ($datesIgnored) { 'PartlyApplied' } else { 'Applied' }
        $result
    }
}

function Resolve-DriveSourceRoot {
    # Turns a Migration Manager source path into a Google folder ID:
    # "/Shared drive name[/folder/...]" or "user@example.com[/folder/...]" (that user's My Drive).
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SourcePath)
    $segments = @($SourcePath.Trim('/').Split('/') | Where-Object { $_ })
    if (!$segments.Count) { throw 'The source path is empty.' }
    $quote = { param($Text) "'" + $Text.Replace('\', '\\').Replace("'", "\'") + "'" }
    if ($SourcePath.StartsWith('/')) {
        $query = [uri]::EscapeDataString("name = $(& $quote $segments[0])")
        $drives = @(@((Invoke-DriveApi "drives?q=$query&fields=drives(id,name)&pageSize=100").drives) | Where-Object { $_ -and $_.name -ceq $segments[0] })
        if ($drives.Count -ne 1) { throw "Found $($drives.Count) shared drives named '$($segments[0])' that this account can see." }
        $current = $drives[0].id
    } else {
        $signedIn = (Invoke-DriveApi 'about?fields=user(emailAddress)').user.emailAddress
        if ($signedIn -ne $segments[0]) { throw "The source is the My Drive of $($segments[0]), but Google sign-in is $signedIn. Sign in as $($segments[0]) or enter the folder ID." }
        $current = (Invoke-DriveApi 'files/root?fields=id').id
    }
    foreach ($name in $segments | Select-Object -Skip 1) {
        $query = [uri]::EscapeDataString("'$current' in parents and name = $(& $quote $name) and mimeType = '$script:FolderMime' and trashed = false")
        $found = @(@((Invoke-DriveApi "files?q=$query&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives").files) | Where-Object { $_ -and $_.name -ceq $name })
        if ($found.Count -ne 1) { throw "Found $($found.Count) folders named '$name' under the source path." }
        $current = $found[0].id
    }
    $current
}

function Save-MigrationSettings {
    # Saves the Entra app (client) ID and the Google OAuth client JSON to the settings folder (.secrets),
    # creating it if needed. A lone client_secret*.json in that folder is found automatically, so its
    # path is stored only when the JSON lives elsewhere; that keeps the folder portable between computers.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SettingsPath,
        [Parameter(Mandatory)][string]$EntraClientId,
        [Parameter(Mandatory)][string]$GoogleClientSecretPath,
        [switch]$CopyGoogleClient
    )
    $folder = [IO.Path]::GetFullPath($PSCmdlet.GetUnresolvedProviderPathFromPSPath((Split-Path $SettingsPath -Parent)))
    [void](New-Item -ItemType Directory -Path $folder -Force)
    $json = (Resolve-Path -LiteralPath $GoogleClientSecretPath).Path
    if ($CopyGoogleClient -and (Split-Path $json -Parent) -ne $folder) {
        $copy = Join-Path $folder (Split-Path $json -Leaf)
        if (!(Test-Path -LiteralPath $copy)) { Copy-Item -LiteralPath $json -Destination $copy }
        $json = $copy
    }
    $found = (Split-Path $json -Parent) -eq $folder -and (Split-Path $json -Leaf) -like 'client_secret*.json' -and
        @(Get-ChildItem -LiteralPath $folder -Filter 'client_secret*.json').Count -eq 1
    $settings = [ordered]@{ EntraClientId = $EntraClientId; GoogleClientSecretPath = if ($found) { '' } else { $json } }
    $settings | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder (Split-Path $SettingsPath -Leaf)) -Encoding utf8
    [pscustomobject]$settings
}

function New-FailedFileSummary {
    # One row per Migration Manager failure: what it was, and where it ended up after this tool's runs.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Plan,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ReportFailures,
        # Run report rows (Copy-LargeDriveFiles CSVs), any order.
        [AllowEmptyCollection()][object[]]$RunReports = @()
    )
    $siteRoot = ([uri]$Plan.DestinationUrl).GetLeftPart([UriPartial]::Authority)
    $index = Get-PlanPathIndex $Plan
    $rank = @{ Metadata = 0; Download = 1; Import = 2; Upload = 3; ApplyMetadata = 4 }
    $byItem = @{}
    foreach ($row in $RunReports | Where-Object { $_.Result -ne 'Preview' -and $_.Source -match '\(ID ([^)]+)\)$' }) {
        $itemId = [regex]::Match($row.Source, '\(ID ([^)]+)\)$').Groups[1].Value
        if (!$byItem.ContainsKey($itemId)) { $byItem[$itemId] = [Collections.Generic.List[object]]::new() }
        $byItem[$itemId].Add($row)
    }
    foreach ($failure in $ReportFailures) {
        $summary = [ordered]@{
            FullPath = $failure.FullPath; MigrationManagerError = "$($failure.ResultCode): $($failure.FailureReason)"
            Source = ''; Destination = ''; Result = 'Failure'; Reason = ''
        }
        $matched = @(Find-PlanPath $Plan $index $failure.FullPath)
        if ($matched.Count -ne 1) {
            $summary.Reason = if ($matched.Count) { 'The path matches several Google items; check it by hand.' } else { 'Not found in the Google source; check it by hand.' }
        } elseif ($matched[0].MimeType -eq $script:FolderMime) {
            $summary.Reason = 'A folder failed, not a file: fix its name or path length (Prepare-DriveMigration report mode), then migrate again.'
        } else {
            $row = $matched[0]
            $summary.Source = "Google Drive: $($failure.FullPath) (ID $($row.Id))"
            $summary.Destination = $siteRoot + $row.TargetPath
            $steps = if ($byItem.ContainsKey($row.Id)) { @($byItem[$row.Id] | Sort-Object { $rank[$_.Step] }) } else { @() }
            $upload = @($steps | Where-Object Step -eq 'Upload' | Select-Object -Last 1)
            $metadata = @($steps | Where-Object Step -eq 'ApplyMetadata' | Select-Object -Last 1)
            if ($upload.Count -and $upload[0].Result -eq 'Success') {
                $summary.Result = 'Success'
                $summary.Reason = $upload[0].Reason
                if ($metadata.Count) { $summary.Reason += "; metadata: $($metadata[0].Result) - $($metadata[0].Reason)" } else { $summary.Reason += '; metadata: not applied' }
            } else {
                $failed = @($steps | Where-Object { $_.Step -ne 'Metadata' -and $_.Result -eq 'Failure' } | Select-Object -Last 1)
                $summary.Reason = if ($failed.Count) { "$($failed[0].Step): $($failed[0].Reason)" } elseif ($steps.Count) { "Stopped after $($steps[-1].Step): $($steps[-1].Reason)" } else { 'Not processed by this tool.' }
            }
        }
        [pscustomobject]$summary
    }
}

Export-ModuleMember -Function New-DriveRenamePlan, Get-DriveInventory, Invoke-DriveRenamePlan, Import-MigrationReport,
    Select-DrivePlanRow, Save-DriveLargeFile, Import-DriveManualDownload, Publish-DriveLargeFile,
    Get-DriveItemMetadata, Set-DriveSharePointMetadata, Resolve-DriveSourceRoot, New-FailedFileSummary, Save-MigrationSettings