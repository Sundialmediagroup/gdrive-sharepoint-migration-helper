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

function Invoke-DriveApi {
    param([string]$Path, [ValidateSet('Get', 'Patch')][string]$Method = 'Get', [object]$Body)
    if (!$env:GDRIVE_ACCESS_TOKEN) { throw 'Set GDRIVE_ACCESS_TOKEN to a short-lived Google OAuth access token. Never store it in the plan.' }
    for ($attempt = 0; $attempt -lt 6; $attempt++) {
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
            $rateLimited = $status -eq 403 -and $_.ErrorDetails -and $_.ErrorDetails.Message -match 'rateLimitExceeded|userRateLimitExceeded'
            if ($attempt -eq 5 -or (!$rateLimited -and $status -notin @(429, 500, 502, 503, 504))) {
                throw "Google Drive request failed (HTTP $status). Check token, permissions, and connectivity; rerun using the same plan/journal."
            }
            [Threading.Thread]::Sleep([int]([Math]::Pow(2, $attempt) * 1000 + (Get-Random -Maximum 500)))
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
    $fingerprint = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $Plan -Depth 20 -Compress))))
    $entries = @()
    if (Test-Path -LiteralPath $journalFullPath) {
        $entries = @(Get-Content -LiteralPath $journalFullPath | Where-Object { $_.Trim() } | ForEach-Object { ConvertFrom-Json $_ })
        if (@($entries | Where-Object PlanHash -ne $fingerprint).Count) { throw 'Journal belongs to another plan. Use the original plan or a new journal.' }
    }
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

Export-ModuleMember -Function New-DriveRenamePlan, Get-DriveInventory, Invoke-DriveRenamePlan