# Google Drive to SharePoint Migration

For model-neutral architecture, artifact schemas, safety rules, and project status,
read [AI-HANDOFF.md](AI-HANDOFF.md).

## Recommendation

Use **Microsoft Migration Manager for the transfer**, with a reviewed source-name
remediation pass only where necessary. Do not build a download/upload migration
engine: document conversion, permissions, versions, throttling, and delta runs
make that substantially riskier than the supported Microsoft service.

Start with Migration Manager's scan and invalid-character replacement settings.
If the scan succeeds with those settings, leave the source untouched. Long paths
are different: shorten the destination site/library/folder URL first, then use
this script if the remaining source names or hierarchy still exceed the budget.
For very deep trees, split the migration into subfolder tasks rather than making
every name unreadably short.

**PowerShell automation boundary:** the documented SPMT cmdlets migrate file
shares and on-premises SharePoint, not Google Drive. This project does not invent
a Google `Add-SPMTTask` command or call private Microsoft endpoints. It automates
Google inventory, naming plans, approved renames, and recovery. It can generate a
supported Migration Manager source CSV, but creating/starting the Google migration
still uses Microsoft's admin interface. The one exception is `Copy-LargeDriveFiles.ps1`
(section 4): it moves only the individual files Migration Manager could not copy,
such as Google-native files over the export limit, into their original folders.

## Requirements

- PowerShell 7.2 or newer (`pwsh`), with no additional PowerShell modules. Only
  the large-file Upload (section 4) needs PnP.PowerShell and PowerShell 7.4+.
- Google Workspace with the Drive API enabled in an approved Google Cloud project.
- An OAuth access token for an account that can access **all** content in the
  selected tree. Workspace admin status alone does not give access to everyone's
  My Drive. Use an appropriately authorized user, shared-drive membership, or an
  approved impersonation setup; run separately for each source root/account.
- For inventory only, `https://www.googleapis.com/auth/drive.metadata.readonly`.
  For applying or rolling back names, `https://www.googleapis.com/auth/drive.metadata`
  permits metadata edits without granting file-content access. Restricted-scope
  consent and organization policies still apply. `drive.file` is insufficient for
  a general inventory of existing organizational files.
- Microsoft Migration Administrator or SharePoint administrator access, authorized
  Google Workspace migration app installation, and provisioned destination libraries.

Set the token without displaying it or putting it in command history:

```powershell
$secret = Read-Host 'Google OAuth access token' -AsSecureString
$env:GDRIVE_ACCESS_TOKEN = [System.Net.NetworkCredential]::new('', $secret).Password
Remove-Variable secret
```

With a Desktop-type OAuth client, save its downloaded `client_secret_*.json` in
`.secrets\` at the project root (git-ignored; create the folder yourself), then run
`.\Get-GoogleToken.ps1`. It signs in through the browser and sets the variable for
the current session; pass `-ClientSecretPath` to use another file, or `-Scope` to
change the requested scopes. Otherwise obtain the token through your organization's
approved OAuth client/flow. Do not
send credentials or tokens through chat. Tokens are short-lived; refresh the
environment variable and rerun with the **same plan and journal** if one expires.
The scripts do not store tokens or refresh credentials. The token is available
to the current process and child processes until removed.

**Google usage limits.** Drive API quotas are 1,000,000 units per minute per project
and 325,000 per minute per user. A read costs 5 units, a list 100, a download 200,
and an edit 50. The scripts make one request at a time, far below these quotas.
Quota errors (`403 User rate limit exceeded`, `429`) and 5xx responses are retried
with Google's recommended backoff: `2^n` seconds plus up to 1 s of random delay,
capped at 64 s, for up to 8 attempts. Projects that used the API between
November 2025 and April 2026 keep their earlier quotas.

Use the actual folder ID from `https://drive.google.com/drive/folders/FOLDER_ID`.
For the My Drive root, resolve the `root` alias first:

```powershell
$root = Invoke-RestMethod -Uri 'https://www.googleapis.com/drive/v3/files/root?fields=id' `
    -Headers @{ Authorization = "Bearer $env:GDRIVE_ACCESS_TOKEN" }
$root.id
```

## 1. Scan and Plan

First create the Google Workspace project in the Microsoft 365 admin center under
**Setup > Migration resources**, connect Google, scan the intended sources, and
download item-level reports. A "Ready to migrate" status is not proof that every
item is supported. Some smaller tenants use the guided Migration Manager Lite
experience instead of the separate standard workflow tabs.

If renames are needed, agree a maintenance window and freeze edits, moves, uploads,
and other naming scripts from inventory through remediation and verification.
Pilot a small representative folder before a production root.

```powershell
.\Prepare-DriveMigration.ps1 -RootId 'GOOGLE_FOLDER_OR_SHARED_DRIVE_ID' `
    -DestinationUrl 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents/Finance' `
    -MigrationSourcePath '/Finance' -PlanPath .\finance-plan.json -Confirm:$false
```

Planning reads Google metadata and writes local JSON and CSV only. The selected
root itself is **not renamed or included as another folder**: its contents map
directly into the destination URL supplied. Match that mapping in Migration
Manager, including any destination subfolder, or the length calculation is wrong.
The destination must be the actual library/folder URL, not `AllItems.aspx`, a
sharing link, or just a site URL. The script does not contact SharePoint to verify
that URL or check existing destination items. Use a new, empty destination.

Optional `-MigrationSourcePath` produces `finance-plan.sources.csv` for the
Microsoft source import. It is a **Migration Manager path**, not a Google ID:
`/Shared Drive Name`, `user@example.com`, or `user@example.com/folder`.
Confirm it describes the same root as `-RootId`; this cross-service association
cannot be verified by the script. Shared-drive display names containing `/`, or
duplicate shared-drive display names, need separate manual remediation. The
root/drive display name is not changed here.

Review every `Rename` and `Warnings` entry in the CSV against the JSON plan.
Names are untrusted source data: inspect the report as plain text; do not enable
spreadsheet formulas or external content. Protect the reports and journals as
sensitive metadata. Do not edit the JSON or CSV to approve individual rows; the
whole plan is the approval unit. Use a smaller source root or change the planner
limits and generate a fresh plan before any source changes if the proposals are
too aggressive.

### Report-Targeted Plan

After a migration attempt, pass Migration Manager's error report to rename only
what failed instead of applying the full naming policy:

```powershell
.\Prepare-DriveMigration.ps1 -RootId 'GOOGLE_FOLDER_ID' `
    -DestinationUrl 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents/Finance' `
    -MigrationSourcePath '/Finance' -MigrationReport .\ProjectError.csv `
    -PlanPath .\finance-retry-plan.json -Confirm:$false
```

Use `ProjectError.csv` (it has `FullPath`), not `MigrationSummary.csv`. The plan
renames reported items and ancestors needed for the path budget; every other name
is kept. Same-folder collisions such as `Report.pdf`/`report.pdf` are left to
Migration Manager, which keeps both and names one `report(1).pdf`; the plan warns
about them and includes the extra characters in path lengths. The whole selected folder is still read (names
only) for collision and drift checks. Review `finance-retry-plan.findings.csv`:
`NotFixed`, `NotFound`, or `Ambiguous` rows are not naming problems the script can
resolve. Migrate the retry to a new, empty destination, not an incremental rerun.

### Naming Policy

- Replaces invalid characters/control characters, trims edge spaces/trailing dots,
  and avoids reserved names. Conservatively reserves `forms` at all depths.
- Defaults to a **120-character segment budget**, configurable up to 255 using
  `-MaxNameLength`. Set 255 to preserve more already-valid names; 120 is an
  operational naming policy, not a claim that SharePoint's cloud limit is 120.
- Uses a maximum **400-character decoded server-relative destination path**,
  including the site/library/folder prefix, separators, and filename. The domain
  and URL percent-encoding expansion are not counted. UTF-16 string length is
  used conservatively, and truncation preserves Unicode surrogate pairs.
- Shortens ancestors as well as filenames when required. Extensions are preserved;
  shortened/colliding names receive a deterministic ID-derived suffix. Names are
  compared case-insensitively and in normalized Unicode form.
- Reserves `.docx`, `.xlsx`, `.pptx`, and `.png` for Google's document, spreadsheet,
  presentation, and drawing conversions. These extensions are projected only;
  they are not appended to Google-native source names. Verify Microsoft's actual
  conversion results during the pilot, especially names already ending in an
  Office extension.
- Detects collisions within the inventoried source siblings, not against existing
  SharePoint content, other source roots, or independently planned migrations.
  Keep distinct roots in separate destination folders.
- Refuses impossible path budgets instead of silently dropping folders or moving
  files. Use a shorter destination or shallower migration tasks in that case.

SharePoint cloud compatibility is not a guarantee that every desktop application
can open a synced local path. Use `-MaxPathLength 240` or a smaller budget if
needed, but separately assess the actual OneDrive local root and your Office
applications; a server-relative budget alone cannot prove desktop compatibility.

## 2. Apply Approved Names

```powershell
.\Prepare-DriveMigration.ps1 -Mode Apply -PlanPath .\finance-plan.json `
    -JournalPath .\finance-journal.jsonl -ApproveSourceRenames -WhatIf

.\Prepare-DriveMigration.ps1 -Mode Apply -PlanPath .\finance-plan.json `
    -JournalPath .\finance-journal.jsonl -ApproveSourceRenames
```

The second command prompts before changing Google names. `-Confirm:$false` is
available for an already-approved unattended run. A read-only token cannot apply
renames; acquire the metadata-write scope before applying. `-WhatIf` still reads
the source for preflight, but performs no PATCH and writes no journal.

The script checks the complete accessible tree's IDs, parents, names, and versions
before writing, and rechecks each affected item immediately before PATCH. It
refuses drift and permission-blocked plans. This is **not an atomic transaction**:
there is still a read/write race, no server-side lock, and Google cannot reveal
items the account cannot see. Freeze changes and run only one instance per tree.
No claim of tenant-wide completeness is made.

Each rename has a flushed write-ahead intent and completion record. Interrupted
runs can resume using the exact original JSON and journal. Keep both together;
the journal is bound to the plan's fingerprint. If a network error occurs after
Google applied a rename, the intent permits recovery on the next run. Do not
start over with an empty journal, delete a partial journal, or re-plan mid-run.
If a journal is damaged or the tree drifted, stop and reconcile it manually.

Renaming leaves Google IDs and content in place, but can affect path-based
integrations and user workflows, and Google may update metadata such as modified
times. Obtain the content owners' approval. Back up the source independently;
this tool's rollback restores **names only**, not metadata timestamps or history.

### Rollback Before Migration

```powershell
.\Prepare-DriveMigration.ps1 -Mode Rollback -PlanPath .\finance-plan.json `
    -JournalPath .\finance-journal.jsonl -ApproveSourceRenames
```

Rollback reverses journaled rename attempts only, in reverse hierarchy order. It
will not overwrite an unexpected name or restore deleted/moved items. Keep the
tree frozen. Do not roll names back after starting the Microsoft migration without
planning destination reconciliation: later delta runs can create duplicate paths.

## 3. Migrate With Microsoft

1. Re-scan in Migration Manager after names are final. Investigate item-level
   failures, not just task status. Optionally upload the generated sources CSV via
   **Drives > Add Drives > Multiple specific Drives** in the standard experience.
2. Copy the scanned tasks to **Drive migrations**. Assign the exact destination
   used by the plan. For bulk destination updates, download Microsoft's current
   destination CSV template, populate its destination column, and upload it back;
   do not upload this tool's naming report as a destination CSV.
3. Map identities and explicitly configure file-level permission and file-version
   migration if required. These are not fully enabled just by mapping identities.
   Library-root destinations do not inherit migrated source root-folder permissions;
   use a dedicated destination folder if those permissions must be carried over.
4. Assign a Microsoft Forms destination for tasks containing Google Forms. Shortcuts,
   Google Sites/Maps, inaccessible/restricted files, orphaned files, file-size limits,
   and tenant-blocked types require separate review or disposition. This naming
   script is not a complete content/permissions/size assessment.
5. Start **Migrate** in the admin UI. Validate counts, individual failures,
   representative converted content, versions, and effective access. Recreate
   required external sharing links/collaboration deliberately; do not assume they
   are carried over. Record every exclusion.
6. For a staged cutover, permit content edits only after the renames are complete,
   keep names/structure fixed, then freeze Google content again for the final delta
   run. Re-run **Migrate** for that delta. Do not rename/reorganize source or
   destination between runs. Retain Google until business owners sign off.

Clear the token after the operation:

```powershell
Remove-Item Env:\GDRIVE_ACCESS_TOKEN
```

## 4. Large Files Migration Manager Could Not Copy

### Guided Run for End Users

```powershell
pwsh -ExecutionPolicy Bypass -File .\Start-FailedFileTransfer.ps1
```

One script does the whole repair and asks only for what it cannot find:

1. The Migration Manager error report (`ProjectError.csv`, its folder, or the
   downloaded .zip), and which source in it to fix.
2. The SharePoint folder that source was migrated into (browser address or a
   **Copy link** sharing link; the script resolves it).
3. The Google OAuth client JSON and the Entra app (client) ID. Type `setup` at
   either prompt and the script writes a copy-pastable setup script for an admin
   (`setup-google-oauth.ps1` or `setup-entra-app.ps1`) and stops.
4. Microsoft and Google sign-ins in the browser. The Google folder is found from
   the report's source path (`/Shared drive` or `user@domain/folder`).
5. It then captures metadata, downloads the failed files (listing any it cannot,
   with links, for a manual download), previews and uploads them, and offers to
   apply sharing (identity map CSV, or a domain-to-domain mapping it creates).

Every step asks before changing SharePoint and writes its run report. The run ends
with `failed-files-summary.csv`: one row per report failure with its Google
source, SharePoint destination, `Success`/`Failure`, and reason. Folders that failed
(for example path too long) are listed as needing the rename steps instead. Work
files go to `Documents\DriveMigration\<source>`; running again for the same source
offers to continue, skipping finished files.

**With Claude Code:** type `/start-transfer` (optionally followed by the report path)
in Claude Code at the repository root. The skill asks the same questions in the chat,
runs the script unattended (`-NonInteractive`), shows the preview, and uploads or
applies sharing only after you say yes. `.claude/settings.json` pre-approves exactly
`pwsh -NoProfile -File .\Start-FailedFileTransfer.ps1 …` and makes rename
Apply/Rollback always ask. Browser sign-ins still need you.

**Unattended parameters.** Every question has a parameter: `-ReportPath`,
`-SourcePath`, `-DestinationLink`, `-GoogleClientSecretPath`, `-EntraClientId`
(`setup` writes the admin script), `-SourceFolderId`, `-IdentityMapPath` or
`-GoogleDomain`/`-MicrosoftDomain`, plus `-SaveSettings`, `-NewRun`. With
`-NonInteractive` a missing answer stops the run, and SharePoint changes need
`-Upload`, `-CreateMissingFolders`, `-ApplySharing`, `-ImportManualDownloads`;
without them the run stops after the upload preview. Exit code 1 means it stopped
early; run again to continue.

**Where the app details are stored.** Both live in the `.secrets` folder next to
the scripts (git-ignored):

```text
.secrets\
  client_secret_....json      Google OAuth client (Desktop app) downloaded from Google Cloud
  migration-settings.json     the Entra app (client) ID
```

```json
{ "EntraClientId": "00000000-0000-0000-0000-000000000000", "GoogleClientSecretPath": "" }
```

You don't have to create these by hand. When the wizard asks for the JSON and the
client ID, paste them; it then offers to **save them in `.secrets`** (creating the
folder) and to **copy the Google JSON there**. Next time it asks nothing. Leave
`GoogleClientSecretPath` empty when exactly one `client_secret*.json` is in `.secrets`:
it is found automatically, so the folder works on any computer. The wizard only
stores a full path when the JSON is kept somewhere else.

**Recommended setup: once per organisation, not per user.** An admin runs the two
`setup` scripts once, puts the Google JSON and `migration-settings.json` in `.secrets`,
and shares the scripts folder through a protected team location (not the public
repository). Neither value is a password (Microsoft treats client IDs as public;
Google treats desktop clients as unable to keep a secret). End users then only answer
the report and destination questions. The Google client's consent screen and Desktop client cannot be created
from the command line; the setup script covers the command-line part (project and
Drive API) and lists the console steps.

### Individual Steps

Google's API export (`files.export`) stops at 10 MB, so very large Docs, Sheets,
and Slides can fail to convert. `Copy-LargeDriveFiles.ps1` rebuilds those files
into their original SharePoint folders after the main migration. It uses the plan
for the migrated root: each row's `TargetPath` is the original folder and name.

1. **Download** the selected items from Google into a local staging folder that
   mirrors the destination. Google-native files are exported to Office/PNG. When
   `files.export` refuses a file, the script tries that file's `exportLinks` URL.
   Google does not document whether that URL accepts larger files, so check the
   run report. Anything still failing is marked `ManualDownload`.
2. **Import** files you downloaded by hand (in the Google editor, **File >
   Download**) from an inbox folder. Names are matched to plan items, including
   the `_` replacements and ` (1)` suffixes browsers add. Unclear matches are
   reported instead of guessed; the run report shows the exact staging path, so
   you can copy those files into place yourself.
3. **Upload** the staging folder to SharePoint with PnP.PowerShell.

```powershell
# Download: needs a token with https://www.googleapis.com/auth/drive.readonly.
.\Get-GoogleToken.ps1 -Scope 'https://www.googleapis.com/auth/drive.readonly'
.\Copy-LargeDriveFiles.ps1 -Mode Download -PlanPath .\finance-plan.json `
    -StagingPath D:\stage\finance -MigrationReport .\ProjectError.csv

# Import manual downloads (all plan files are candidates unless -Id/-MigrationReport is given).
.\Copy-LargeDriveFiles.ps1 -Mode Import -PlanPath .\finance-plan.json `
    -StagingPath D:\stage\finance -InboxPath "$HOME\Downloads\finance-large"

# Upload: preview, then upload with a confirmation prompt.
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/IT -ClientId 'YOUR_ENTRA_APP_ID' -Interactive
.\Copy-LargeDriveFiles.ps1 -Mode Upload -PlanPath .\finance-plan.json -StagingPath D:\stage\finance `
    -JournalPath .\finance-upload-journal.jsonl -ApproveUpload -WhatIf
.\Copy-LargeDriveFiles.ps1 -Mode Upload -PlanPath .\finance-plan.json -StagingPath D:\stage\finance `
    -JournalPath .\finance-upload-journal.jsonl -ApproveUpload
```

Select items with `-Id` (Google file IDs) or `-MigrationReport` (the failed rows of
`ProjectError.csv`; add `-ReportSourcePath` if it covers several sources).
Download never fetches a whole tree. It stops at Google's egress limit of
**1 TB per project per day** (`-MaxDownloadBytes`, default 10^12 bytes). Remaining
items are marked `Deferred`; rerun after 24 hours and already-staged files are skipped.
The budget counts only this run, so lower it if other jobs use the same Google Cloud
project.

**Run report.** Every run (all five modes) writes `large-files-<mode>-<time>.csv`
(`-StatusPath` overrides the name; existing files are never overwritten), one row
per item:

| Column | Content |
| --- | --- |
| `RunUtc`, `Step` | When the run started and which mode ran |
| `Source` | `Google Drive: /folder/name (ID …)`, or `Local file: …` for unmatched manual downloads |
| `Destination` | Full SharePoint URL of the file or folder |
| `Result` | `Success` or `Failure` |
| `Reason` | Why: e.g. `Downloaded from Google (ExportLink)`, `Uploaded to SharePoint`, `Dates, people, and sharing applied; shared with …`, or the failure cause |
| `Status`, `LocalPath` | Detailed status code and staging file, for troubleshooting |

`Already…` results count as `Success` (done in an earlier run). If a run stops
partway (for example a lost connection), the report is still written with the
items handled so far and a final `Failure` row giving the reason it stopped. `-WhatIf`
previews write no report. Review every `Failure` row.

**Which plan.** Use the plan whose names match what was migrated. If you applied
renames, use that plan. If you migrated without renames, create a report-targeted plan
(it keeps names verbatim). Download checks that each item's current Google name
equals the plan's `NewName`. Rows whose target name contains characters SharePoint
rejects (Migration Manager replaced them during the copy, using its own settings)
and rows where Migration Manager chose which collision got a `(1)` suffix are
reported for manual handling, not guessed.

**SharePoint upload limits handled by the script:**

- Files over **250 GB** are refused (SharePoint's per-file limit).
- Target paths over **400** decoded characters, invalid characters, and reserved
  names are refused.
- `Add-PnPFile` uploads files over 10 MB in 10 MB chunks, so large files are not
  limited by the single-request size.
- `Add-PnPFile` silently **overwrites** existing files and **creates** missing
  folders. The script checks first: it never uploads over an existing file
  (`Exists`), and skips files whose folder is missing (`FolderMissing`), because
  that usually means Migration Manager put the folder under a different name.
  Pass `-CreateMissingFolders` only once you've confirmed the folder really is
  absent; the script then creates it with `Resolve-PnPFolder` (in the live pilot,
  `Add-PnPFile` 3.4.1 failed with NotFound instead of creating it). A check-then-upload race remains, so keep the destination frozen.
- Every uploaded file is checked by size afterwards. A JSONL journal records
  intent and completion. Rerun with the same plan and journal to resume. A file
  left half-uploaded by an interrupted run is reported as `PartialUpload` for
  manual cleanup, never overwritten.
- Staging files that map to no plan item stop the upload before anything is sent.
- Tenant blocked file types and storage quota are not checked; they appear as
  upload errors.

**Not carried over by Upload:** sharing, author, and dates. The uploaded file shows
the uploading account and the upload time until you reapply metadata (below).
Version history is never carried over.

### Reapply Dates, People, and Sharing

Capture Google's metadata **at the same time as the download** (keep the tree frozen),
then apply it after the upload:

```powershell
# Read-only: dates, owner, last editor, and sharing for every item in the plan tree.
.\Copy-LargeDriveFiles.ps1 -Mode Metadata -PlanPath .\finance-plan.json
# -> finance-plan.metadata.json and finance-plan.permissions.csv (review the CSV)

# After Upload, with the same PnP connection: preview, then apply.
.\Copy-LargeDriveFiles.ps1 -Mode ApplyMetadata -PlanPath .\finance-plan.json `
    -IdentityMapPath .\identity-map.csv -JournalPath .\finance-upload-journal.jsonl `
    -MetadataJournalPath .\finance-metadata-journal.jsonl -ApproveMetadata -WhatIf
```

The identity map is a CSV you write, mapping Google accounts to Microsoft 365
sign-in names. A row starting with `@` maps a whole domain; explicit rows win:

```csv
GoogleIdentity,Microsoft365Identity
jane.doe@example-google.com,jdoe@contoso.com
@example-google.com,@contoso.com
```

What Apply does, for each file in the **upload journal** plus the folders above it
(`-SkipFolders` limits it to files):

- **Dates and people.** Tries to set Created, Modified, Author (Google owner), and
  Modified By (Google last editor) with `UpdateOverwriteVersion`. **In the live pilot
  (Member account, `AllSites.Write`) SharePoint accepted the call but kept its own
  values**; the REST `ValidateUpdateListItem` route behaved the same. Overriding
  these system fields needs more than Edit rights. The script reads the dates back
  and reports `PartlyApplied` when they were ignored; the Google values stay in the
  metadata file. People need a mapping; unmapped people are reported.
- **Sharing.** Grants only sharing set **directly** on that item in Google. Inherited
  sharing comes from the folder above, which gets its own direct grants. Folders are
  processed first. Access is granted with Microsoft Graph's sharing invite, as
  SharePoint's **Share** dialog does: Google reader/commenter become `read`;
  writer and organizer become `write`. The Google owner becomes the Author, not a
  sharing grant (Google lists every My Drive item's owner as a direct permission). People must sign in, no invitation
  email is sent, and existing (inherited) access is kept. Grants are only ever added.
  In SharePoint they appear as sharing on the item ("Shared with").
- **Reported, not applied:** domain-wide and "anyone with the link" sharing (no
  direct equivalent), deleted accounts, unmapped people, items whose Google
  permission list was unreadable (the account could not share them), and sharing on
  the source root or shared-drive membership. Grant those at the library, folder, or
  site membership yourself.
- A journal makes reruns skip finished items. More than 5,000 items needing unique
  permissions triggers a warning; Microsoft recommends no more than that per library.

The Microsoft 365 accounts must already exist (create guests for external people
first). Google groups need a mail-enabled Microsoft 365 or security group; whether
SharePoint resolves a group's email directly should be checked in the pilot. Setting
Created/Author through PnP is commonly used but not covered by PnP's documentation;
also confirm it in the pilot. Sharing changes made in Google after the capture are
not picked up: capture again with a new `-MetadataPath`.

**PnP.PowerShell requirements:** `Install-Module PnP.PowerShell -Scope CurrentUser`.
Version 3.x requires PowerShell 7.4 or later and your own Entra ID app registration.
PnP is only needed for Upload and ApplyMetadata; the other modes and the tests run
without it. Connect to the site (or subsite) that contains the destination library.

**Minimum Entra app permissions (all delegated; no Full Control):**

| API | Permission | Used for |
| --- | --- | --- |
| SharePoint | `AllSites.Write` | Upload; Created/Modified/Author/Editor |
| Microsoft Graph | `Files.ReadWrite.All` | Sharing invites in ApplyMetadata |
| Microsoft Graph | `User.Read` | Sign-in |

Delegated permissions never exceed the signed-in account. Run as a dedicated
migration account that is a **Member (Edit)** of the destination sites, not an
Owner. The sites' sharing settings must let members share, and external sharing must
be allowed for people outside the organization. The app needs the platform *Mobile
and desktop applications* with redirect URI `http://localhost`, and admin consent:

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'Drive Migration' -Tenant contoso.onmicrosoft.com `
    -SharePointDelegatePermissions 'AllSites.Write' -GraphDelegatePermissions 'Files.ReadWrite.All', 'User.Read'
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/IT -ClientId 'APP_CLIENT_ID' -Interactive
```

With these permissions, uploads and sharing worked in the live pilot, but original
Created/Modified/Author/Editor values could not be applied (see above).

Keep the staging folder outside this repository on an encrypted, access-controlled
disk with enough free space, and delete it after sign-off. It holds full file
content. Use a short staging path; Windows handles long paths in PowerShell 7, but
other tools may not.

## Offline Tests

```powershell
pwsh -NoProfile -File .\Test-DriveMigration.ps1
```

Tests use synthetic metadata and mocked Google calls. They cover naming and path
budgets, conversion collisions, Unicode truncation, pagination, shared-drive
queries, incomplete inventories, dry runs, resume, lost-response recovery,
rollback, drift rejection, and the large-file download/import/upload flow with
mocked Google and PnP calls. They do not validate a real Google account,
SharePoint destination, Microsoft migration, conversion fidelity, or permission
mapping. A live pilot is required before production use.

## Live Pilot Test Data

`New-DriveTestData.ps1` builds the same 26-item tree of SharePoint-incompatible
names (invalid characters, reserved names, case and conversion collisions, a
300-character name, and a deep path over 400 characters) in each target folder.
Use dedicated, empty test shared drives only; the script refuses non-empty targets.
Two drives allow a comparison: leave one unremediated for Migration Manager's own
handling, and run Plan/Apply on the other.

```powershell
.\New-DriveTestData.ps1 -TargetFolderId 'RAW_TEST_DRIVE_ID', 'FIXED_TEST_DRIVE_ID' -WhatIf
.\New-DriveTestData.ps1 -TargetFolderId 'RAW_TEST_DRIVE_ID', 'FIXED_TEST_DRIVE_ID'
```

The token needs `https://www.googleapis.com/auth/drive.file` to create files, plus
`drive.metadata` (also used by Plan/Apply) to confirm the targets are empty. If
Google rejects creation in a shared drive with HTTP 403/404, use a test-only token
with `https://www.googleapis.com/auth/drive` for seeding. Files contain one line of
placeholder text, so `.pdf`/`.docx`/`.xlsx` uploads are not valid documents; they
test names and paths only. If the run stops partway, empty the drives and rerun.

## Official References

- [Google Workspace migration overview](https://learn.microsoft.com/en-us/sharepointmigration/mm-google-overview)
- [SPMT PowerShell scope and supported task types](https://learn.microsoft.com/en-us/sharepointmigration/overview-spmt-ps-cmdlets)
- [Migration Manager project settings](https://learn.microsoft.com/en-us/sharepointmigration/mm-project-settings)
- [Google migration FAQ and delta behavior](https://learn.microsoft.com/en-us/sharepointmigration/mm-faqs-google)
- [Microsoft source CSV format](https://learn.microsoft.com/en-us/sharepointmigration/mm-google-bulk-upload-source-paths-csv)
- [Destination assignment and CSV template](https://learn.microsoft.com/en-us/sharepointmigration/mm-google-step4-review-destinations)
- [Unsupported content and Google conversions](https://learn.microsoft.com/en-us/sharepointmigration/mm-unsupported-files)
- [SharePoint naming restrictions](https://support.microsoft.com/en-us/office/restrictions-and-limitations-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa)
- [Google Drive file listing](https://developers.google.com/workspace/drive/api/reference/rest/v3/files/list)
- [Google Drive metadata updates and OAuth scopes](https://developers.google.com/workspace/drive/api/reference/rest/v3/files/update)
- [Google Drive API usage limits](https://developers.google.com/workspace/drive/api/guides/limits)
- [Google Drive downloads and the 10 MB export limit](https://developers.google.com/workspace/drive/api/guides/manage-downloads)
- [PnP.PowerShell Add-PnPFile](https://pnp.github.io/powershell/cmdlets/Add-PnPFile.html) and [Entra app registration](https://pnp.github.io/powershell/articles/registerapplication.html)
- [Set-PnPListItem](https://pnp.github.io/powershell/cmdlets/Set-PnPListItem.html) and [Graph driveItem invite](https://learn.microsoft.com/en-us/graph/api/driveitem-invite)
- [Google Drive permissions resource](https://developers.google.com/workspace/drive/api/reference/rest/v3/permissions)