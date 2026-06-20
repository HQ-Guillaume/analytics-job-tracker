[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")

function Assert-Rule {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Scoring rule test failed: $Message"
    }
}

Assert-Rule -Condition ((Get-IgnoreReasonFromNotes "ignore_reason=too_data_engineering; detail=dbt") -eq "too_data_engineering") -Message "Structured data-engineering ignore reason was not parsed."
Assert-Rule -Condition ((Get-IgnoreReasonFromNotes "mostly SEO SEA acquisition") -eq "") -Message "Profile-specific free text should not be inferred globally."
Assert-Rule -Condition ((Get-IgnoreReasonFromNotes "not relevant for this profile") -eq "not_relevant_enough") -Message "Generic not-relevant free text was not inferred."
Assert-Rule -Condition ((Get-CompanyAliasFromNotes "ignore_reason=duplicate; company_alias=Example Company; detail=same posting") -eq "Example Company") -Message "Structured company_alias note was not parsed."
Assert-Rule -Condition ((Get-CompanyAliasFromNotes "ignore_reason=duplicate; duplicate_of=Fallback Company; detail=same posting") -eq "Fallback Company") -Message "Structured duplicate_of note was not parsed."
& (Join-Path $projectRoot "app\cli\Find-AnalyticsJobs.ps1") -SelfTest

Write-Host "Scoring rule tests passed."
