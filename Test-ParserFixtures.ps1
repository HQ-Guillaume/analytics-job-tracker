[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Find-AnalyticsJobs.ps1") -SelfTest

function Assert-Fixture {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Parser fixture test failed: $Message"
    }
}

$fixtureRoot = Join-Path $PSScriptRoot "tests\fixtures"

$helloWorkCard = Get-Content -LiteralPath (Join-Path $fixtureRoot "hellowork-card.html") -Raw
$helloWorkStats = Start-SourceStats "HelloWork fixture"
$helloWorkCandidates = @(Get-HelloWorkCardCandidates -Html $helloWorkCard -SearchUrl "https://www.hellowork.com/fr-fr/emploi/recherche.html?k=web%20analyst" -Query "web analyst" -Stats $helloWorkStats)
Assert-Fixture -Condition ($helloWorkCandidates.Count -eq 1) -Message "Expected one HelloWork card candidate."
Assert-Fixture -Condition ($helloWorkCandidates[0].Title -eq "Web Analyst H/F") -Message "Expected HelloWork card title mapping."
Assert-Fixture -Condition ($helloWorkCandidates[0].Company -eq "Example Retailer") -Message "Expected HelloWork card company mapping."
Assert-Fixture -Condition ($helloWorkCandidates[0].Location -eq "Paris - 75") -Message "Expected HelloWork card location mapping."
Assert-Fixture -Condition ($helloWorkCandidates[0].Contract -eq "CDI") -Message "Expected HelloWork card contract mapping."

$helloWorkDetail = Get-Content -LiteralPath (Join-Path $fixtureRoot "hellowork-detail.html") -Raw
$helloWorkMetadata = Get-HelloWorkJobMetadata -Html $helloWorkDetail
Assert-Fixture -Condition ($helloWorkMetadata.Title -eq "Web Analyst H/F") -Message "Expected HelloWork detail title mapping."
Assert-Fixture -Condition ($helloWorkMetadata.Company -eq "Example Retailer") -Message "Expected HelloWork detail company mapping."
Assert-Fixture -Condition ($helloWorkMetadata.Contract -eq "CDI") -Message "Expected HelloWork detail contract mapping."
Assert-Fixture -Condition ($helloWorkMetadata.Description -match "Piano Analytics") -Message "Expected HelloWork detail description mapping."

$apecJob = Get-Content -LiteralPath (Join-Path $fixtureRoot "apec-job.json") -Raw | ConvertFrom-Json
Assert-Fixture -Condition ((Get-ApecContractType $apecJob) -eq "CDI") -Message "Expected APEC contract mapping."
Assert-Fixture -Condition ((Get-ApecJobUrl $apecJob) -match "/detail-offre/123456789W$") -Message "Expected APEC URL mapping."

$linkedinCard = Get-Content -LiteralPath (Join-Path $fixtureRoot "linkedin-card.html") -Raw
Assert-Fixture -Condition ((Get-LinkedInLocationFromHtml $linkedinCard) -eq "Paris, Ile-de-France") -Message "Expected LinkedIn location mapping."
Assert-Fixture -Condition (-not (Test-ShouldSkipEarlyByContract -Text $linkedinCard)) -Message "Expected LinkedIn CDI-free card not to be excluded early."

$dedupeMatch = Get-JobMatch -Title "Web Analyst" -Text "GA4 Google Tag Manager ContentSquare dataLayer"
$dedupeRows = @(
    (New-JobResult -Title "Web Analyst H/F" -CompanyName "Example Retailer" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $dedupeMatch.Score -MatchLevel $dedupeMatch.Level -MatchedKeywords $dedupeMatch.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/123.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager"),
    (New-JobResult -Title "Web Analyst F/H" -CompanyName "Example Retailer" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $dedupeMatch.Score -MatchLevel $dedupeMatch.Level -MatchedKeywords $dedupeMatch.Keywords -Url "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/123W" -Platform "APEC" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager")
)
$dedupeKeys = @($dedupeRows | ForEach-Object { Get-JobDedupeKeyFromRow $_ } | Select-Object -Unique)
Assert-Fixture -Condition ($dedupeKeys.Count -eq 1) -Message "Expected APEC and HelloWork fixture rows to dedupe together."

Write-Host "Parser fixture tests passed."
