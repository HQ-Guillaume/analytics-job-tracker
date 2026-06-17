[CmdletBinding()]
param(
    [string]$TrackerPath = "",
    [string]$ConfigDirectory = "config",
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")
. (Join-Path $PSScriptRoot "app\JobTracker.Config.ps1")
. (Join-Path $PSScriptRoot "app\JobTracker.Runtime.ps1")
. (Join-Path $PSScriptRoot "app\JobTracker.Scoring.ps1")
. (Join-Path $PSScriptRoot "app\JobTracker.Deduplication.ps1")
. (Join-Path $PSScriptRoot "app\JobTracker.Excel.ps1")

$configPath = Resolve-JobCrawlerPath -BasePath $PSScriptRoot -Path $ConfigDirectory
$JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $configPath
$JobCrawlerRuntimeConfig = $JobCrawlerConfig.Runtime
$JobCrawlerSourcesConfig = $JobCrawlerConfig.Sources
$JobCrawlerMatchingRules = $JobCrawlerConfig.MatchingRules
$JobCrawlerWorkbookConfig = $JobCrawlerConfig.Workbook

$validation = Test-JobCrawlerConfig -Config $JobCrawlerConfig
if (-not $validation.IsValid) {
    throw ("Invalid crawler config:`n- {0}" -f (($validation.Issues) -join "`n- "))
}

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = Resolve-JobCrawlerPath -BasePath $PSScriptRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.tracker_path" -DefaultValue "output\jobs_tracker.xlsx"))
}

if ((Test-Path -LiteralPath $TrackerPath) -and -not $Force) {
    throw "Tracker already exists: $TrackerPath. Use -Force only if you intentionally want to replace it."
}

$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$RunDate = Get-Date -Format "yyyy-MM-dd"
$DaysBack = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.days_back" -DefaultValue 7)
$Cutoff = [DateTimeOffset]::Now.AddDays(-[Math]::Abs($DaysBack))
$CutoffDate = $Cutoff.ToString("yyyy-MM-dd")
$CrawlMode = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.crawl_mode" -DefaultValue "Default")
$Location = [string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.location" -DefaultValue "France")
$CacheDirectory = Resolve-JobCrawlerPath -BasePath $PSScriptRoot -Path ([string](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_directory" -DefaultValue "output\cache"))
$CacheTtlHours = [int](Get-ConfigPathValue -Object $JobCrawlerRuntimeConfig -Path "defaults.cache_ttl_hours" -DefaultValue 24)
$MinimumMatchScore = [int](Get-ConfigPathValue -Object $JobCrawlerMatchingRules -Path "thresholds.minimum_match_score" -DefaultValue 35)
$LinkedInQueries = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $JobCrawlerSourcesConfig -Path "queries.linkedin" -DefaultValue @()))
$ApiSearchQueries = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $JobCrawlerSourcesConfig -Path "queries.api" -DefaultValue @()))
$SourceRunStats = New-Object System.Collections.Generic.List[object]
$JobCrawlerPreferences = Get-JobCrawlerPreferences
$MasterColumns = Get-JobTrackerMasterColumns
$ColumnLabels = Get-JobTrackerColumnLabels

$summary = @{
    TotalMatched = 0
    ExcludedContractCount = 0
    CurrentCount = 0
    TrackerCount = 0
    DuplicateCount = 0
    RemovedCount = 0
    PreservedAppliedCount = 0
    SourceDiagnostics = "Initialized empty workbook; no crawl run."
    BackupPath = ""
    CrawlCaps = "Initialize only"
    DryRun = "no"
    DiagnosticMode = "no"
    DiagnosticPath = ""
}

Export-TrackerWorkbook -Rows @() -Path $TrackerPath -Summary $summary
Write-Host ("Initialized tracker: {0}" -f (Resolve-Path $TrackerPath).Path)
