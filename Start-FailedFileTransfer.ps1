#Requires -Version 7.4
# Guided run for end users: reads Migration Manager's error report and moves the files it could not
# copy into their original SharePoint folders, then reapplies Google sharing. Asks only for what it
# cannot find in .secrets\migration-settings.json; prints setup scripts when an app is missing.
# Every step writes a run report; the run ends with failed-files-summary.csv (one row per failure).
# Every question has a matching parameter. With -NonInteractive nothing is asked: missing answers stop the
# run, and SharePoint changes happen only with -Upload / -ApplySharing (used by the /start-transfer skill).
[CmdletBinding()]
param(
    [string]$SettingsPath = (Join-Path $PSScriptRoot '.secrets\migration-settings.json'),
    [string]$WorkRoot = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'DriveMigration'),
    [string]$ReportPath,
    [string]$SourcePath,
    [string]$DestinationLink,
    [string]$GoogleClientSecretPath,
    [string]$EntraClientId,
    [string]$SourceFolderId,
    [string]$IdentityMapPath,
    [string]$GoogleDomain,
    [string]$MicrosoftDomain,
    [switch]$NonInteractive,
    [switch]$NewRun,
    [switch]$SaveSettings,
    [switch]$ImportManualDownloads,
    [switch]$CreateMissingFolders,
    [switch]$Upload,
    [switch]$ApplySharing
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force
$copyScript = Join-Path $PSScriptRoot 'Copy-LargeDriveFiles.ps1'
if ($NonInteractive -and $ApplySharing -and !$IdentityMapPath -and !($GoogleDomain -and $MicrosoftDomain)) {
    throw '-ApplySharing needs -IdentityMapPath, or -GoogleDomain and -MicrosoftDomain.'
}

function Read-Answer {
    # A value given as a parameter wins; unattended runs use the default or stop with the parameter to pass.
    param([string]$Prompt, [string]$Default = '', [string]$Value = '', [string]$Name = '')
    if ($Value) { Write-Host "${Prompt}: $Value"; return $Value }
    if ($NonInteractive) {
        if ($Default) { Write-Host "${Prompt}: $Default"; return $Default }
        throw "Missing -$Name (the run is unattended)."
    }
    $answer = (Read-Host $(if ($Default) { "$Prompt [$Default]" } else { $Prompt })).Trim().Trim('"').Trim("'")
    if ($answer) { $answer } else { $Default }
}

function Read-YesNo {
    # Unattended runs answer with -Unattended, which callers take from the matching switch.
    param([string]$Prompt, [bool]$Default = $true, [bool]$Unattended = $false)
    if ($NonInteractive) { Write-Host "$Prompt $(if ($Unattended) { 'yes' } else { 'no' })"; return $Unattended }
    $answer = Read-Host "$Prompt $(if ($Default) { '[Y/n]' } else { '[y/N]' })"
    if (!$answer.Trim()) { return $Default }
    return $answer.Trim() -match '^(y|yes)$'
}

function Write-Step { param([string]$Text) Write-Host ''; Write-Host "== $Text" -ForegroundColor Cyan }

function Save-SetupScript {
    param([string]$Path, [string]$Content, [string]$Who)
    Set-Content -LiteralPath $Path -Value $Content -Encoding utf8
    Write-Host $Content -ForegroundColor Gray
    Write-Host "Setup script saved to $Path. $Who, then run this script again." -ForegroundColor Yellow
}

function Get-EntraSetupScript {
    param([string]$Tenant)
    @"
# Creates the Entra ID app for the Drive migration tool. Run once per organisation, in PowerShell 7.4+,
# as an admin who can grant admin consent (for example a Global Administrator).
# Delegated permissions only, no Full Control. People using the app never get more than their own access:
#   SharePoint      AllSites.Write       upload files, set dates
#   Microsoft Graph Files.ReadWrite.All  share files with people, like SharePoint's Share button
#   Microsoft Graph User.Read            sign in
if (!(Get-Module -ListAvailable PnP.PowerShell)) { Install-Module PnP.PowerShell -Scope CurrentUser }
Import-Module PnP.PowerShell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'Drive Migration' -Tenant '$Tenant' ``
    -SharePointDelegatePermissions 'AllSites.Write' -GraphDelegatePermissions 'Files.ReadWrite.All', 'User.Read'
# Accept the consent prompt. Share the printed 'Application (client) ID' with the people running the migration,
# or save it as EntraClientId in .secrets\migration-settings.json next to the migration scripts.
# The people running it need Member (Edit) access to the destination sites, not Owner.
"@
}

function Get-GoogleSetupScript {
    @'
# Creates the Google sign-in client for the Drive migration tool. Run once per Google Workspace,
# as a Workspace admin (or a Google Cloud user who may create projects).
# Part 1 - command line, in PowerShell with the Google Cloud CLI (https://cloud.google.com/sdk/docs/install):
$project = 'drive-migration-' + (Get-Random -Maximum 999999)
gcloud projects create $project --name='Drive Migration'
gcloud config set project $project
gcloud services enable drive.googleapis.com
"Created Google Cloud project $project"

# Part 2 - Google Cloud console, with that project selected (Google offers no command for desktop sign-in clients):
#  1. https://console.cloud.google.com/auth/overview  ->  Get started: app name 'Drive Migration', your support email,
#     Audience: Internal (only your Workspace accounts can sign in; no Google review needed).
#  2. https://console.cloud.google.com/auth/scopes  ->  Add or remove scopes  ->  add
#     https://www.googleapis.com/auth/drive.readonly   (read-only: list, read sharing, download)
#     Add https://www.googleapis.com/auth/drive.metadata as well only if you also use the rename steps.
#  3. https://console.cloud.google.com/auth/clients  ->  Create client  ->  Application type: Desktop app
#     ->  Create  ->  Download JSON (client_secret_....json).
#  4. Only if your Workspace limits third-party apps: admin.google.com  ->  Security  ->  Access and data control
#     ->  API controls  ->  Manage third-party app access  ->  add the client ID and set it to Trusted.
# Save the JSON in the .secrets folder next to the migration scripts (or give its path to the wizard).
# Google treats desktop clients as unable to keep a secret, but store the file in a protected team location.
'@
}

function Resolve-SharePointFolder {
    # Any SharePoint folder link (browser address, library view, or "Copy link") -> folder URL and site URL.
    param([string]$Link)
    $candidates = [Collections.Generic.List[string]]::new()
    $candidates.Add($Link)
    $uri = [uri]$Link
    $id = [regex]::Match($uri.Query, '(?:^\?|&)id=([^&]+)').Groups[1].Value
    if ($id) { $candidates.Add($uri.GetLeftPart([UriPartial]::Authority) + [uri]::UnescapeDataString($id)) }
    foreach ($candidate in $candidates) {
        $token = 'u!' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($candidate)).TrimEnd('=').Replace('/', '_').Replace('+', '-')
        try {
            $item = Invoke-PnPGraphMethod -Url "v1.0/shares/$token/driveItem?`$select=name,webUrl,folder,parentReference" -Method Get -ErrorAction Stop
        } catch { continue }
        if (!$item.PSObject.Properties['folder'] -or !$item.folder) { throw 'That link points to a file. Give the link of the folder Migration Manager copied into.' }
        $root = Invoke-PnPGraphMethod -Url "v1.0/drives/$($item.parentReference.driveId)/root?`$select=sharepointIds" -Method Get -ErrorAction Stop
        return [pscustomobject]@{ FolderUrl = $item.webUrl; SiteUrl = $root.sharepointIds.siteUrl }
    }
    throw 'SharePoint could not resolve that link. Open the folder in the browser and use Copy link, or check that your account can open it.'
}

Write-Host 'Failed-file transfer: moves files Migration Manager could not copy into their SharePoint folders.' -ForegroundColor Cyan
Write-Host 'Google is only read; nothing in Google is changed. You confirm each SharePoint change before it happens.'

# --- Settings and prerequisites
$settings = [ordered]@{ EntraClientId = ''; GoogleClientSecretPath = '' }
if (Test-Path -LiteralPath $SettingsPath) {
    foreach ($property in (Get-Content -LiteralPath $SettingsPath -Raw | ConvertFrom-Json).PSObject.Properties) { $settings[$property.Name] = [string]$property.Value }
}
if (!(Get-Module -ListAvailable PnP.PowerShell)) {
    if (!(Read-YesNo 'PnP.PowerShell (needed for SharePoint) is not installed. Install it for your account now?' -Unattended $true)) { throw 'PnP.PowerShell is required.' }
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
}
Import-Module PnP.PowerShell

# --- 1. Error report
Write-Step 'Step 1 of 7: Migration Manager error report'
Write-Host 'In the Microsoft 365 admin center, open the migration, download the task report, and give its path here.'
do {
    $reportInput = Read-Answer 'Path to ProjectError.csv, its folder, or the downloaded .zip' -Value $ReportPath -Name ReportPath
    $ReportPath = ''
    $reportFile = $null
    if ($reportInput -and (Test-Path -LiteralPath $reportInput)) {
        if ($reportInput -like '*.zip') {
            $unzipped = Join-Path ([IO.Path]::GetTempPath()) ('mm-report-' + [guid]::NewGuid())
            Expand-Archive -LiteralPath $reportInput -DestinationPath $unzipped
            $reportInput = $unzipped
        }
        $reportFile = if (Test-Path -LiteralPath $reportInput -PathType Leaf) { Get-Item -LiteralPath $reportInput } else { Get-ChildItem -LiteralPath $reportInput -Recurse -Filter 'ProjectError*.csv' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
    }
    if (!$reportFile) {
        $message = 'No ProjectError.csv found there. MigrationSummary.csv has no file paths; use the error report.'
        if ($NonInteractive) { throw $message }
        Write-Warning $message
    }
} while (!$reportFile)
$allRows = @(Import-Csv -LiteralPath $reportFile.FullName)
$sources = @($allRows | Group-Object SourcePath | Sort-Object Name)
if (!$sources.Count) { Write-Host 'The report lists no failures. Nothing to do.'; return }
for ($i = 0; $i -lt $sources.Count; $i++) { Write-Host ("  {0}. {1}  ({2} failures)" -f ($i + 1), $sources[$i].Name, $sources[$i].Count) }
if ($SourcePath) {
    if ($SourcePath -notin @($sources | ForEach-Object Name)) { throw "-SourcePath '$SourcePath' is not in the report. Sources: $((@($sources | ForEach-Object Name)) -join '; ')" }
    $sourcePath = $SourcePath
} elseif ($sources.Count -eq 1) {
    $sourcePath = $sources[0].Name
} elseif ($NonInteractive) {
    throw "The report covers several sources; pass -SourcePath: $((@($sources | ForEach-Object Name)) -join '; ')"
} else {
    $sourcePath = $sources[[int](Read-Answer 'Which source? Number' '1') - 1].Name
}
$failures = @(Import-MigrationReport -Path $reportFile.FullName -SourcePath $sourcePath)
Write-Host "$($failures.Count) failures for $sourcePath. Error codes: $((@($failures | Group-Object ResultCode | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '))"
$safeName = ($sourcePath -replace '[^A-Za-z0-9._-]+', '_').Trim('_')
$work = Join-Path $WorkRoot $safeName
$runFile = Join-Path $work 'run.json'
# One work folder per source, so a second run continues with the same plan, downloads, and journals.
$resume = (Test-Path -LiteralPath $runFile) -and (Read-YesNo "A previous run for $sourcePath exists in $work. Continue it (finished files are skipped)?" -Unattended (!$NewRun))
if (!$resume -and (Test-Path -LiteralPath $work)) { $work = "$work-" + [datetime]::Now.ToString('yyyyMMdd-HHmmss'); $runFile = Join-Path $work 'run.json' }
[void](New-Item -ItemType Directory -Path $work -Force)
Copy-Item -LiteralPath $reportFile.FullName -Destination (Join-Path $work 'ProjectError.csv') -Force
Write-Host "Work folder (plans, downloads, reports): $work"

# --- 2. Destination
Write-Step 'Step 2 of 7: SharePoint destination'
if ($resume) {
    $run = Get-Content -LiteralPath $runFile -Raw | ConvertFrom-Json
    $folderLink = $run.FolderUrl
    Write-Host "Continuing with $([uri]::UnescapeDataString($run.FolderUrl))"
} else {
    Write-Host "Paste the SharePoint folder Migration Manager copied $sourcePath into (browser address or Copy link)."
    do {
        $folderLink = Read-Answer 'SharePoint folder link' -Value $DestinationLink -Name DestinationLink
        if ($NonInteractive -and $folderLink -notmatch '^https://[^/]+\.sharepoint\.com/') { throw '-DestinationLink must be a https://<tenant>.sharepoint.com/ link.' }
        $DestinationLink = $null
    } while ($folderLink -notmatch '^https://[^/]+\.sharepoint\.com/')
}
$spHost = ([uri]$folderLink).Host
$tenant = $spHost.Split('.')[0] + '.onmicrosoft.com'

# --- 3. App details
Write-Step 'Step 3 of 7: Google and Microsoft app details'
$secretsFolder = Split-Path $SettingsPath -Parent
$defaultJson = $settings.GoogleClientSecretPath
if (!$defaultJson) {
    # A lone client_secret*.json in the settings folder is used without asking for its path.
    $found = @(Get-ChildItem -LiteralPath $secretsFolder -Filter 'client_secret*.json' -ErrorAction SilentlyContinue)
    if ($found.Count -eq 1) { $defaultJson = $found[0].FullName }
}
$googleJson = Read-Answer "Google OAuth client JSON (type 'setup' if your organisation has none)" $defaultJson -Value $GoogleClientSecretPath -Name GoogleClientSecretPath
if ($googleJson -eq 'setup') { Save-SetupScript (Join-Path $work 'setup-google-oauth.ps1') (Get-GoogleSetupScript) 'Ask a Google Workspace admin to run it'; return }
if (!(Test-Path -LiteralPath $googleJson) -or !(Get-Content -LiteralPath $googleJson -Raw | ConvertFrom-Json).PSObject.Properties['installed']) {
    throw "$googleJson is not a Google Desktop OAuth client JSON. Run the script again and type 'setup' for instructions."
}
$clientId = Read-Answer "Entra app (client) ID for the migration tool (type 'setup' if your organisation has none)" $settings.EntraClientId -Value $EntraClientId -Name EntraClientId
if ($clientId -eq 'setup') { Save-SetupScript (Join-Path $work 'setup-entra-app.ps1') (Get-EntraSetupScript $tenant) 'Ask a Microsoft 365 admin to run it'; return }
if ($clientId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { throw 'That is not an app (client) ID; it looks like 00000000-0000-0000-0000-000000000000.' }
$googleJson = (Resolve-Path -LiteralPath $googleJson).Path
$jsonChanged = !$defaultJson -or $googleJson -ne (Resolve-Path -LiteralPath $defaultJson -ErrorAction SilentlyContinue).Path
if (($jsonChanged -or $clientId -ne $settings.EntraClientId) -and (Read-YesNo "Save these app details in $secretsFolder so you are not asked again?" -Unattended $SaveSettings.IsPresent)) {
    $copyJson = (Split-Path $googleJson -Parent) -ne [IO.Path]::GetFullPath($secretsFolder) -and (Read-YesNo "Also copy the Google JSON into $secretsFolder?" -Unattended $SaveSettings.IsPresent)
    [void](Save-MigrationSettings -SettingsPath $SettingsPath -EntraClientId $clientId -GoogleClientSecretPath $googleJson -CopyGoogleClient:$copyJson)
    Write-Host "Saved to $SettingsPath"
}

$googleUser = if ($sourcePath.StartsWith('/')) { 'an account that is a member of that shared drive' } else { $sourcePath.Split('/')[0] }
$reports = [Collections.Generic.List[string]]::new()
$stopped = $null
Push-Location -LiteralPath $work
try {
    # --- 4. Sign-ins and plan
    Write-Step 'Step 4 of 7: Sign in to Microsoft 365, then Google'
    Write-Host 'A browser window opens for Microsoft 365. Sign in with your work account.'
    if ($resume) {
        $destination = $run
    } else {
        Connect-PnPOnline -Url "https://$spHost" -ClientId $clientId -Interactive -PersistLogin
        $destination = Resolve-SharePointFolder $folderLink
    }
    Connect-PnPOnline -Url $destination.SiteUrl -ClientId $clientId -Interactive -PersistLogin
    Write-Host "Destination: $([uri]::UnescapeDataString($destination.FolderUrl))"
    $signInGoogle = {
        Write-Host "A browser window opens for Google. Sign in as $googleUser."
        & (Join-Path $PSScriptRoot 'Get-GoogleToken.ps1') -ClientSecretPath $googleJson -Scope 'https://www.googleapis.com/auth/drive.readonly' -TimeoutSeconds 600
    }
    if (!$resume) {
        & $signInGoogle
        try {
            $rootId = if ($SourceFolderId) { $SourceFolderId } else { Resolve-DriveSourceRoot -SourcePath $sourcePath }
        } catch {
            Write-Warning $_.Exception.Message
            $answer = Read-Answer 'Paste the Google folder link or ID of the source instead' -Name SourceFolderId
            $rootId = $answer
        }
        if ($rootId -match '/folders/([A-Za-z0-9_-]+)') { $rootId = $Matches[1] }
        Write-Host 'Reading the Google folder tree (names only)...'
        $items = @(Get-DriveInventory -RootId $rootId)
        # Names are kept as they are now in Google: that is what Migration Manager copied.
        New-DriveRenamePlan -Items $items -RootId $rootId -DestinationUrl $destination.FolderUrl -ReportFailures @() |
            ConvertTo-Json -Depth 20 | Set-Content -LiteralPath .\plan.json -Encoding utf8
        [ordered]@{ SourcePath = $sourcePath; RootId = $rootId; FolderUrl = $destination.FolderUrl; SiteUrl = $destination.SiteUrl } |
            ConvertTo-Json | Set-Content -LiteralPath $runFile -Encoding utf8
    }
    $plan = Get-Content -LiteralPath .\plan.json -Raw | ConvertFrom-Json
    $selected = @(Select-DrivePlanRow -Plan $plan -ReportFailures $failures -WarningAction SilentlyContinue)
    Write-Host "$(@($plan.Rows).Count) items in the Google source; $($selected.Count) failed files can be transferred."
    foreach ($row in $selected) { Write-Host "  $($row.TargetPath)" }

    if ($selected.Count) {
        $ids = @($selected | ForEach-Object Id)
        # --- 5. Download
        Write-Step 'Step 5 of 7: Download from Google'
        $prefix = [uri]::UnescapeDataString(([uri]$plan.DestinationUrl).AbsolutePath).TrimEnd('/')
        $unstaged = @($selected | Where-Object { !(Test-Path -LiteralPath (Join-Path (Join-Path $work 'stage') $_.TargetPath.Substring($prefix.Length + 1))) })
        if (!$unstaged.Count -and (Test-Path -LiteralPath .\plan.metadata.json)) {
            Write-Host 'All failed files were downloaded in an earlier run; no Google sign-in needed.'
        } else {
            if ($resume) { & $signInGoogle }
            if (!(Test-Path -LiteralPath .\plan.metadata.json)) {
                $result = & $copyScript -Mode Metadata -PlanPath .\plan.json -Confirm:$false
                if ($result.ReportPath) { $reports.Add($result.ReportPath) }
            }
            $result = & $copyScript -Mode Download -PlanPath .\plan.json -StagingPath .\stage -Id $ids -Confirm:$false
            if ($result.ReportPath) { $reports.Add($result.ReportPath) }
        }
        $lastDownload = Get-ChildItem -LiteralPath $work -Filter 'large-files-download-*.csv' | Sort-Object LastWriteTime | Select-Object -Last 1
        $manual = @(if ($lastDownload) { Import-Csv -LiteralPath $lastDownload.FullName | Where-Object { $_.Status -eq 'ManualDownload' -and !(Test-Path -LiteralPath $_.LocalPath) } })
        if ($manual.Count) {
            $inbox = Join-Path $work 'manual-downloads'
            [void](New-Item -ItemType Directory -Path $inbox -Force)
            Write-Host "$($manual.Count) file(s) could not be downloaded automatically. Open each link, use File > Download (Office format), and save it in:" -ForegroundColor Yellow
            Write-Host "  $inbox"
            foreach ($row in $manual) {
                $itemId = [regex]::Match($row.Source, '\(ID ([^)]+)\)$').Groups[1].Value
                Write-Host "  https://drive.google.com/open?id=$itemId  ->  $(Split-Path $row.LocalPath -Leaf)"
            }
            $importNow = if ($NonInteractive) {
                if (!$ImportManualDownloads) { Write-Host 'Run again with -ImportManualDownloads once they are saved.' }
                $ImportManualDownloads.IsPresent
            } else {
                Invoke-Item -LiteralPath $inbox
                (Read-Answer "Press Enter when the files are saved (or type 'skip')") -ne 'skip'
            }
            if ($importNow) {
                $manualIds = @($manual | ForEach-Object { [regex]::Match($_.Source, '\(ID ([^)]+)\)$').Groups[1].Value })
                $result = & $copyScript -Mode Import -PlanPath .\plan.json -StagingPath .\stage -InboxPath $inbox -Id $manualIds -Confirm:$false
                if ($result.ReportPath) { $reports.Add($result.ReportPath) }
            }
        }

        # --- 6. Upload
        Write-Step 'Step 6 of 7: Upload to SharePoint'
        $uploadArgs = @{ PlanPath = '.\plan.json'; StagingPath = '.\stage'; JournalPath = '.\upload-journal.jsonl'; ApproveUpload = $true }
        $preview = & $copyScript -Mode Upload @uploadArgs -WhatIf
        Write-Host "Preview: $($preview.Summary)"
        if ($preview.Summary -match 'FolderMissing') {
            Write-Host 'Some destination folders do not exist. Usually that means Migration Manager renamed them; check the preview warnings.'
            $uploadArgs.CreateMissingFolders = Read-YesNo 'Create the missing folders?' $false -Unattended $CreateMissingFolders.IsPresent
        }
        if (!(Read-YesNo "Upload to $([uri]::UnescapeDataString($destination.FolderUrl))? Existing files are never overwritten." -Unattended $Upload.IsPresent)) {
            Write-Host 'Nothing uploaded. Run again (with -Upload when unattended) to upload; downloads are kept.'
        } else {
            $result = & $copyScript -Mode Upload @uploadArgs -Confirm:$false
            if ($result.ReportPath) { $reports.Add($result.ReportPath) }

            # --- 7. Sharing, dates, people
            Write-Step 'Step 7 of 7: Sharing, dates, and people'
            if (Read-YesNo 'Give the same people access as in Google (sharing invites, no emails sent)?' -Unattended $ApplySharing.IsPresent) {
                $map = if ($IdentityMapPath -or $NonInteractive) { $IdentityMapPath } else { Read-Answer 'Identity map CSV (GoogleIdentity,Microsoft365Identity), or press Enter to map by email domain' }
                if (!$map) {
                    $googleEmail = if ($env:GDRIVE_ACCESS_TOKEN) { (Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/about?fields=user(emailAddress)' -Headers @{ Authorization = 'Bearer ' + $env:GDRIVE_ACCESS_TOKEN }).user.emailAddress } else { '' }
                    $microsoftEmail = (Invoke-PnPGraphMethod -Url 'v1.0/me?$select=userPrincipalName' -Method Get).userPrincipalName
                    $googleDomain = Read-Answer 'Google email domain' $(if ($googleEmail) { $googleEmail.Split('@')[1] }) -Value $GoogleDomain -Name GoogleDomain
                    $microsoftDomain = Read-Answer 'Microsoft 365 email domain for the same people' $microsoftEmail.Split('@')[1] -Value $MicrosoftDomain -Name MicrosoftDomain
                    $map = Join-Path $work 'identity-map.csv'
                    "GoogleIdentity,Microsoft365Identity`n@$googleDomain,@$microsoftDomain" | Set-Content -LiteralPath $map -Encoding utf8
                }
                $apply = @{ PlanPath = '.\plan.json'; JournalPath = '.\upload-journal.jsonl'; MetadataJournalPath = '.\metadata-journal.jsonl'; IdentityMapPath = $map; ApproveMetadata = $true }
                $preview = & $copyScript -Mode ApplyMetadata @apply -WhatIf
                Write-Host "Preview: $($preview.Summary)"
                if (Read-YesNo 'Apply sharing, dates, and people now?' -Unattended $ApplySharing.IsPresent) {
                    $result = & $copyScript -Mode ApplyMetadata @apply -Confirm:$false
                    if ($result.ReportPath) { $reports.Add($result.ReportPath) }
                }
            }
        }
    }
} catch {
    $stopped = $_
    Write-Warning "Stopped: $($_.Exception.Message)"
    Write-Warning 'Run the script again with the same report to continue; finished files are skipped.'
} finally {
    Remove-Item Env:\GDRIVE_ACCESS_TOKEN -ErrorAction SilentlyContinue
    Pop-Location
}

# --- Summary
$planPath = Join-Path $work 'plan.json'
if (Test-Path -LiteralPath $planPath) {
    $plan = Get-Content -LiteralPath $planPath -Raw | ConvertFrom-Json
    # Include reports from earlier runs in this work folder, so a continued run summarises everything.
    $runRows = @(Get-ChildItem -LiteralPath $work -Filter 'large-files-*.csv' | Sort-Object LastWriteTime | ForEach-Object { Import-Csv -LiteralPath $_.FullName })
    $summary = @(New-FailedFileSummary -Plan $plan -ReportFailures $failures -RunReports $runRows)
    $summaryPath = Join-Path $work 'failed-files-summary.csv'
    $summary | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
    Write-Step 'Summary'
    $summary | Format-Table FullPath, Result, Reason -Wrap | Out-String -Width 160 | Write-Host
    Write-Host "$(@($summary | Where-Object Result -eq 'Success').Count) of $($summary.Count) failures fixed. Summary: $summaryPath"
    Write-Host "Step reports and files: $work"
}
# A non-zero exit lets unattended callers (the /start-transfer skill) see that the run stopped early.
if ($stopped) { exit 1 }
