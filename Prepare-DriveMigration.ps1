#Requires -Version 7.2
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [ValidateSet('Plan', 'Apply', 'Rollback')][string]$Mode = 'Plan',
    [string]$RootId,
    [string]$DestinationUrl,
    [string]$MigrationSourcePath,
    [string[]]$MigrationReport,
    [string]$ReportSourcePath,
    [string]$PlanPath = '.\migration-plan.json',
    [string]$JournalPath = '.\rename-journal.jsonl',
    [ValidateRange(80, 400)][int]$MaxPathLength = 400,
    [ValidateRange(24, 255)][int]$MaxNameLength = 120,
    [switch]$ApproveSourceRenames
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'DriveMigration.psm1') -Force
# Resolve against the PowerShell location; [IO.Path]::GetFullPath ignores Set-Location.
$PlanPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PlanPath)
$JournalPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($JournalPath)

if ($Mode -eq 'Plan') {
    if (!$RootId -or !$DestinationUrl) { throw 'Plan requires -RootId and -DestinationUrl.' }
    $reportPath = [IO.Path]::ChangeExtension($PlanPath, '.csv')
    $sourcesPath = [IO.Path]::ChangeExtension($PlanPath, '.sources.csv')
    $findingsPath = [IO.Path]::ChangeExtension($PlanPath, '.findings.csv')
    if ([IO.Path]::GetExtension($PlanPath) -ne '.json') { throw 'PlanPath must end in .json.' }
    if ((Test-Path -LiteralPath $PlanPath) -or (Test-Path -LiteralPath $reportPath) -or ($MigrationSourcePath -and (Test-Path -LiteralPath $sourcesPath)) -or
        ($MigrationReport -and (Test-Path -LiteralPath $findingsPath))) { throw 'Plan/report already exists. Choose a new PlanPath; keep old plans with their journals.' }
    $planOptions = @{ RootId = $RootId; DestinationUrl = $DestinationUrl; MaxPathLength = $MaxPathLength; MaxNameLength = $MaxNameLength }
    if ($MigrationReport) {
        $source = if ($ReportSourcePath) { $ReportSourcePath } else { $MigrationSourcePath }
        $planOptions.ReportFailures = @(Import-MigrationReport -Path $MigrationReport -SourcePath $source)
    }
    $items = @(Get-DriveInventory -RootId $RootId)
    $plan = New-DriveRenamePlan -Items $items @planOptions
    if ($PSCmdlet.ShouldProcess($PlanPath, 'Write local naming plan and CSV report; no Google changes')) {
        $plan | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $PlanPath -Encoding utf8
        $plan.Rows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8
        if ($MigrationSourcePath) {
            [pscustomobject]@{ 'Source Path' = $MigrationSourcePath } | Export-Csv -LiteralPath $sourcesPath -NoTypeInformation -Encoding utf8
        }
        if ($MigrationReport) { $plan.ReportFindings | Export-Csv -LiteralPath $findingsPath -NoTypeInformation -Encoding utf8 }
    }
    foreach ($finding in $plan.ReportFindings | Where-Object Status -ne 'Fixed') {
        Write-Warning "$($finding.Status): $($finding.FullPath) ($($finding.ResultCode)) needs manual review."
    }
    [pscustomobject]@{
        Items = $items.Count; Renames = @($plan.Rows | Where-Object Rename).Count
        Warnings = @($plan.Rows | Where-Object Warnings).Count
        ReportFailures = @($plan.ReportFindings).Count; ReportFixed = @($plan.ReportFindings | Where-Object Status -eq 'Fixed').Count
        PlanPath = $PlanPath; ReportPath = $reportPath
    }
} else {
    if (!$ApproveSourceRenames) { throw 'Apply/Rollback requires -ApproveSourceRenames after reviewing the plan.' }
    $plan = Get-Content -LiteralPath $PlanPath -Raw | ConvertFrom-Json
    $options = @{
        Plan = $plan; JournalPath = $JournalPath; ApproveSourceRenames = $true
        Rollback = ($Mode -eq 'Rollback'); WhatIf = [bool]$WhatIfPreference
    }
    if ($PSBoundParameters.ContainsKey('Confirm')) { $options.Confirm = [bool]$PSBoundParameters['Confirm'] }
    Invoke-DriveRenamePlan @options
}