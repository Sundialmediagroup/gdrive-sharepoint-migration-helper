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
still uses Microsoft's admin interface. No files are transferred by these scripts.

## Requirements

- PowerShell 7.2 or newer (`pwsh`), with no additional PowerShell modules.
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

## Offline Tests

```powershell
pwsh -NoProfile -File .\Test-DriveMigration.ps1
```

Tests use synthetic metadata and mocked Google calls. They cover naming and path
budgets, conversion collisions, Unicode truncation, pagination, shared-drive
queries, incomplete inventories, dry runs, resume, lost-response recovery,
rollback, and drift rejection. They do not validate a real Google account,
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