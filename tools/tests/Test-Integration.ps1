[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Context.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Runtime.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.OutputMaintenance.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.SourceAdapter.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Scoring.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Deduplication.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Excel.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Pipeline.ps1")
$script:LoadedSourceAdapters = @(Get-JobCrawlerSourceAdapterFiles -SourcesRoot (Join-Path $projectRoot "app\sources"))
foreach ($sourceAdapter in $script:LoadedSourceAdapters) {
    . $sourceAdapter.Path
}

$script:IntegrationTempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-integration-config-{0}" -f ([Guid]::NewGuid().ToString("N")))
$script:IntegrationConfigDirectory = Join-Path $script:IntegrationTempRoot "config"
New-Item -ItemType Directory -Path $script:IntegrationTempRoot -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $projectRoot "config") -Destination $script:IntegrationConfigDirectory -Recurse -Force
Remove-Item -LiteralPath (Join-Path $script:IntegrationConfigDirectory "local") -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -LiteralPath $script:IntegrationConfigDirectory -Filter "local*.json" -File -ErrorAction SilentlyContinue |
    Remove-Item -Force

$integrationProfile = New-JobCrawlerProfileFromBuilder `
    -Label "Integration Analytics" `
    -Id "integration_analytics" `
    -TargetTitles @("Web Analyst", "Digital Analytics Consultant", "Tracking Analyst", "Product Analyst", "CRO Analyst") `
    -ImportantSkills @("Google Analytics", "GA4", "Google Tag Manager", "Piano Analytics", "ContentSquare", "dataLayer", "Tag Commander", "Tealium", "server-side", "RGPD", "CRO") `
    -ExclusionKeywords @("SEO", "SEA", "data engineer", "dbt", "snowflake") `
    -SearchQueries @("web analyst", "digital analytics", "tracking analyst", "product analyst cro") `
    -TargetLocations @("France", "Paris") `
    -ExcludedLocations @("London", "New York") `
    -ExcludedContracts @("CDD", "Apprenticeship", "Internship", "Freelance") `
    -EmployerPreference "annonceur" `
    -Compact
[void](Save-JobCrawlerLocalProfile -ConfigDirectory $script:IntegrationConfigDirectory -Profile $integrationProfile)
[void](Set-JobCrawlerDefaultProfile -ConfigDirectory $script:IntegrationConfigDirectory -ProfileId "integration_analytics")

function Assert-Integration {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Integration test failed: $Message"
    }
}

$script:JobCrawlerConfig = Get-JobCrawlerConfig -ConfigDirectory $script:IntegrationConfigDirectory
$script:JobCrawlerRuntimeConfig = $script:JobCrawlerConfig.Runtime
$script:JobCrawlerSourcesConfig = $script:JobCrawlerConfig.Sources
$script:JobCrawlerMatchingRules = $script:JobCrawlerConfig.MatchingRules
$script:JobCrawlerWorkbookConfig = $script:JobCrawlerConfig.Workbook
$script:JobCrawlerPreferences = Get-JobCrawlerPreferences
$script:MasterColumns = Get-JobTrackerMasterColumns
$script:ColumnLabels = Get-JobTrackerColumnLabels
$script:SeenResultKeys = @{}
$script:FeedbackLearningProfile = $null
$script:SourceRunStats = New-Object System.Collections.Generic.List[object]
$script:DisableCache = $false
$script:CacheTtlHours = 24
$script:Location = "France"
$script:CrawlMode = "Default"
$script:RunDate = Get-Date -Format "yyyy-MM-dd"
$script:RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:Cutoff = [DateTimeOffset]::Now.AddDays(-7)
$script:CutoffDate = $script:Cutoff.ToString("yyyy-MM-dd")
$script:MinimumMatchScore = [int](Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "thresholds.minimum_match_score" -DefaultValue 35)

Assert-Integration -Condition ($script:JobCrawlerConfig.Profile.Id -eq "integration_analytics") -Message "Expected integration test profile to load from local config."
Assert-Integration -Condition (@(Get-ConfigStringArray (Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "queries.linkedin" -DefaultValue @())).Count -gt 0) -Message "Expected profile-level LinkedIn queries to merge into sources config."
Assert-Integration -Condition (@(Get-ConfigPathValue -Object $script:JobCrawlerMatchingRules -Path "positive_signals" -DefaultValue @()).Count -gt 0) -Message "Expected profile-level positive matching signals."

$sources = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $script:JobCrawlerSourcesConfig)
$sourceContract = Test-JobCrawlerSourceContract -SourceDefinitions $sources -LoadedFiles $script:LoadedSourceAdapters
Assert-Integration -Condition $sourceContract.IsValid -Message "Expected configured source functions to be loaded dynamically."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "linkedin" -and $_.CrawlFunction -eq "Get-LinkedInJobs" }).Count -eq 1) -Message "Expected LinkedIn source registry metadata."
Assert-Integration -Condition (@($sources | Where-Object { $_.Key -eq "wttj_public" -and $_.SkipSwitch -eq "DisableWttjPublicFallback" }).Count -eq 1) -Message "Expected WTTJ public fallback to have its own skip switch."
Assert-Integration -Condition ($sources.Count -eq 6) -Message "Expected the public source registry to contain only the six supported crawler sources."

$customSourcesConfig = [PSCustomObject]@{
    source_order = @("custom_board")
    sources      = [PSCustomObject]@{
        custom_board = [PSCustomObject]@{
            label                = "Custom board"
            short_label          = "Custom"
            enabled_by_default   = $true
            requires_credentials = $false
            crawl_function       = "Get-CustomJobs"
        }
    }
}
$customSources = @(Get-JobCrawlerSourceDefinitions -SourcesConfig $customSourcesConfig)
Assert-Integration -Condition ($customSources.Count -eq 1 -and $customSources[0].Key -eq "custom_board" -and $customSources[0].CrawlFunction -eq "Get-CustomJobs") -Message "Expected custom config-defined source metadata without code changes."

$match = Get-JobMatch -Title "Web Analyst" -Text "Google Tag Manager GA4 ContentSquare dataLayer"
$rows = @(
    (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.linkedin.com/jobs/view/222" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"),
    (New-JobResult -Title "Web Analyst H/F" -CompanyName "Radio France" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/222.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager")
)
$merge = Merge-JobsWithTracker -CurrentRows $rows -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($merge.TrackerRows).Count -eq 1) -Message "Expected similar cross-platform rows to merge."
Assert-Integration -Condition ((Get-RowValue -Row $merge.TrackerRows[0] -Name "source_count") -eq "2") -Message "Expected merged source count to be 2."

$oldCurrentRow = New-JobResult -Title "Web Analyst" -CompanyName "Old Company" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/old-job" -Platform "Test" -PublishedAt ([DateTimeOffset]::Now.AddDays(-8)) -SourceText "GA4 Google Tag Manager"
$oldCurrentMerge = Merge-JobsWithTracker -CurrentRows @($oldCurrentRow) -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($oldCurrentMerge.TrackerRows).Count -eq 0 -and [int]$oldCurrentMerge.RemovedCount -eq 0) -Message "Expected current rows outside the published-date retention window to be rejected at merge time without counting as removed tracker rows."

$existingCddRow = New-OrderedJobRecord @{
    status         = "ignored"
    job_title      = "Web Analyst"
    company_name   = "CDD Company"
    location       = "Paris"
    contract_type  = "CDD"
    match_score    = "80"
    match_level    = "High"
    job_url_raw    = "https://example.test/cdd-existing"
    platform       = "LinkedIn"
    published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
}
$currentBlankContractRow = New-JobResult -Title "Web Analyst" -CompanyName "CDD Company" -JobLocation "Paris" -ContractType "" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://example.test/cdd-current" -Platform "Test" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"
$excludedContractMerge = Merge-JobsWithTracker -CurrentRows @($currentBlankContractRow) -ExistingRows @($existingCddRow) -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($excludedContractMerge.TrackerRows).Count -eq 0 -and [int]$excludedContractMerge.RemovedCount -eq 1) -Message "Expected excluded existing contract values not to leak back into current non-application rows."

$nextonDuplicateRows = @(
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Data Analyst Marketing H/F - NEXTON - CDI à Lyon"
        company_name   = "Nexton Consulting"
        location       = "Paris, FR"
        contract_type  = "CDI"
        match_score    = "52"
        match_level    = "Medium"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/nexton-consulting/jobs/data-analyst-senior-digital-analytics-h-f_lyon"
        platform       = "Welcome to the Jungle; LinkedIn"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }),
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Data Analyst Senior Digital Analytics H F"
        company_name   = "NEXTON"
        location       = "Lyon"
        contract_type  = "CDI"
        match_score    = "55"
        match_level    = "Medium"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/nexton-consulting/jobs/data-analyst-senior-digital-analytics-h-f_lyon?utm_source=test"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    })
)
$nextonMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows $nextonDuplicateRows -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($nextonMerge.TrackerRows).Count -eq 1 -and [int]$nextonMerge.DuplicateCount -eq 1) -Message "Expected exact canonical URL duplicates to merge even when titles, company labels, and locations differ."

$nextonAliasRows = @(
    (New-JobResult -Title "Web Analyst" -CompanyName "Nexton Consulting" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://fr.linkedin.com/jobs/view/nexton-web-analyst-111" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager CRO"),
    (New-JobResult -Title "Web Analyst H/F" -CompanyName "NEXTON" -JobLocation "Paris, Ile-de-France" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/nexton-web-analyst-222.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager CRO")
)
$nextonAliasMerge = Merge-JobsWithTracker -CurrentRows $nextonAliasRows -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($nextonAliasMerge.TrackerRows).Count -eq 1 -and [int]$nextonAliasMerge.DuplicateCount -eq 1 -and (Get-RowValue -Row $nextonAliasMerge.TrackerRows[0] -Name "company_name") -eq "nexton") -Message "Expected configured Nexton alias group to merge different-source rows and display the canonical lowercase company name."

$olivierRows = @(
    (New-JobResult -Title "Web Analyst CRO" -CompanyName "L'Olivier Assurance" -JobLocation "Paris" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://fr.linkedin.com/jobs/view/web-analyst-cro-at-lolivier-assurance-111" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager CRO"),
    (New-JobResult -Title "Web Analyst CRO H/F" -CompanyName "Olivier" -JobLocation "Paris, Ile-de-France" -ContractType "CDI" -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url "https://www.welcometothejungle.com/fr/companies/olivier/jobs/web-analyst-cro_paris" -Platform "Welcome to the Jungle" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager CRO")
)
$olivierMerge = Merge-JobsWithTracker -CurrentRows $olivierRows -ExistingRows @() -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($olivierMerge.TrackerRows).Count -eq 1 -and (Get-RowValue -Row $olivierMerge.TrackerRows[0] -Name "platform") -match "LinkedIn" -and (Get-RowValue -Row $olivierMerge.TrackerRows[0] -Name "platform") -match "Welcome to the Jungle" -and (Get-RowValue -Row $olivierMerge.TrackerRows[0] -Name "company_name") -eq "olivier assurance") -Message "Expected company alias hierarchy to merge L'Olivier Assurance and Olivier and display olivier assurance."

$feedbackAliasRows = @(
    (New-OrderedJobRecord @{
        status         = "ignored"
        notes          = "ignore_reason=duplicate; company_alias=Example Retail; detail=same posting"
        job_title      = "Web Analyst"
        company_name   = "Example Retail France"
        location       = "Paris"
        contract_type  = "CDI"
        match_score    = "70"
        match_level    = "Medium"
        job_url_raw    = "https://www.linkedin.com/jobs/view/feedback-alias-1"
        platform       = "LinkedIn"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }),
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Web Analyst H/F"
        company_name   = "Example Retail"
        location       = "Paris, Ile-de-France"
        contract_type  = "CDI"
        match_score    = "72"
        match_level    = "Medium"
        job_url_raw    = "https://www.hellowork.com/fr-fr/emplois/feedback-alias-2.html"
        platform       = "HelloWork"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    })
)
$feedbackAliasMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows $feedbackAliasRows -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($feedbackAliasMerge.TrackerRows).Count -eq 1 -and [int]$feedbackAliasMerge.DuplicateCount -eq 1) -Message "Expected Apply notes company_alias feedback to merge duplicate company labels without hard-coded aliases."

$seniorityDuplicateRows = @(
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Digital Analytics Consultant H F Cdi"
        company_name   = "Fifty Five"
        location       = "Paris"
        contract_type  = "CDI"
        match_score    = "70"
        match_level    = "Medium"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/fifty-five/jobs/digital-analytics-consultant-h-f-cdi_paris"
        platform       = "HelloWork; Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }),
    (New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "Senior Digital Analytics Consultant H F Paris"
        company_name   = "Fifty Five"
        location       = "138Ewzd"
        contract_type  = "CDI"
        match_score    = "72"
        match_level    = "Medium"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/fifty-five/jobs/senior-digital-analytics-consultant-h-f-paris_paris_fifty_138Ewzd"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    })
)
$seniorityDuplicateMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows $seniorityDuplicateRows -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition (@($seniorityDuplicateMerge.TrackerRows).Count -eq 1 -and [int]$seniorityDuplicateMerge.DuplicateCount -eq 1) -Message "Expected seniority wording and WTTJ reference-token locations not to block same-company/same-role deduplication."

$foreignExistingWttjRow = New-OrderedJobRecord @{
    status         = "ignored"
    job_title      = "Head Of Uk Sports Gtm"
    company_name   = "Example Company"
    location       = "London"
    contract_type  = "CDI"
    match_score    = "80"
    match_level    = "Good"
    job_url_raw    = "https://www.welcometothejungle.com/en/companies/acme/jobs/head-of-uk-sports-gtm_london"
    platform       = "Welcome to the Jungle"
    published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
}
$cleanupMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows @($foreignExistingWttjRow) -Path "integration.xlsx" -SkipBackup
Assert-Integration -Condition ([int]$cleanupMerge.RemovedCount -eq 1) -Message "Expected invalid WTTJ existing row cleanup."

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("custom-job-tracker-integration-{0}" -f ([Guid]::NewGuid().ToString("N")))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
try {
    $oldFile = Join-Path $tempRoot "old.txt"
    $newFile = Join-Path $tempRoot "new.txt"
    Set-Content -LiteralPath $oldFile -Value "old" -Encoding UTF8
    Set-Content -LiteralPath $newFile -Value "new" -Encoding UTF8
    (Get-Item -LiteralPath $oldFile).LastWriteTime = (Get-Date).AddDays(-60)
    $prune = Invoke-JobCrawlerCachePrune -Path $tempRoot -Enabled:$true
    Assert-Integration -Condition ([int]$prune.RemovedFiles -ge 1 -and -not (Test-Path -LiteralPath $oldFile) -and (Test-Path -LiteralPath $newFile)) -Message "Expected cache prune to remove old cache files only."

    $managedRoot = Join-Path $tempRoot "project"
    $managedCache = Join-Path $managedRoot "output\cache"
    New-Item -ItemType Directory -Path $managedCache -Force | Out-Null
    $managedOld = Join-Path $managedCache "old-cache.txt"
    $managedNew = Join-Path $managedCache "new-cache.txt"
    Set-Content -LiteralPath $managedOld -Value "old" -Encoding UTF8
    Set-Content -LiteralPath $managedNew -Value "new" -Encoding UTF8
    (Get-Item -LiteralPath $managedOld).LastWriteTime = (Get-Date).AddDays(-20)
    $cleanupRows = @(Invoke-JobCrawlerOutputCleanup -ProjectRoot $managedRoot -CacheDirectory $managedCache -Cache -OlderThanDays 14)
    Assert-Integration -Condition (($cleanupRows | Measure-Object RemovedFiles -Sum).Sum -eq 1 -and -not (Test-Path -LiteralPath $managedOld) -and (Test-Path -LiteralPath $managedNew)) -Message "Expected managed output cleanup to remove old cache files inside the project root only."

    $managedTrackerPath = Get-JobCrawlerTrackerPath -ProjectRoot $managedRoot -Config $script:JobCrawlerConfig
    $managedOutput = Split-Path -Parent $managedTrackerPath
    $managedBackups = Join-Path $managedOutput "backups"
    $managedLogs = Join-Path $managedOutput "launcher_logs"
    New-Item -ItemType Directory -Path $managedBackups -Force | Out-Null
    New-Item -ItemType Directory -Path $managedLogs -Force | Out-Null
    $managedBackup = Join-Path $managedBackups "jobs_tracker_backup.xlsx"
    $managedLog = Join-Path $managedLogs "launcher_run_test.log"
    Set-Content -LiteralPath $managedBackup -Value "backup" -Encoding UTF8
    Set-Content -LiteralPath $managedLog -Value "log" -Encoding UTF8
    $allCleanupRows = @(Invoke-JobCrawlerOutputCleanup -ProjectRoot $managedRoot -CacheDirectory $managedCache -All -OlderThanDays 0)
    Assert-Integration -Condition (($allCleanupRows | Measure-Object RemovedFiles -Sum).Sum -ge 3 -and -not (Test-Path -LiteralPath $managedBackup) -and -not (Test-Path -LiteralPath $managedLog) -and -not (Test-Path -LiteralPath $managedNew)) -Message "Expected all-output cleanup to remove cache, backups, and launcher logs immediately."

    $historyPath = Join-Path $tempRoot "run_history.jsonl"
    for ($i = 0; $i -lt 3; $i++) {
        Write-RunHistoryEntry -Path $historyPath -MaxEntries 2 -Summary @{
            DryRun = "yes"
            DiagnosticMode = "no"
            TotalMatched = $i
            CurrentCount = $i
            TrackerCount = $i
            DuplicateCount = 0
            RemovedCount = 0
            PreservedAppliedCount = 0
            ExcludedContractCount = 0
        }
    }
    Assert-Integration -Condition (@(Get-Content -LiteralPath $historyPath).Count -eq 2) -Message "Expected run history pruning to keep the configured max entries."
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Remove-Item -LiteralPath $script:IntegrationTempRoot -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Integration tests passed."
