function Assert-ScoringCondition {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw "Scoring self-test failed: $Message"
    }
}

function Invoke-ScoringSelfTest {
    $script:JobCrawlerPreferences = Get-JobCrawlerPreferences
    $script:SeenResultKeys = @{}
    $script:FeedbackLearningProfile = $null

    $mojibakeInterim = "Int" + [string][char]0x00C3 + [string][char]0x00A9 + "rim - 6 Mois"
    $expectedInterim = "Int" + [string][char]0x00E9 + "rim - 6 Mois"
    Assert-ScoringCondition -Condition ((Repair-DisplayText $mojibakeInterim) -eq $expectedInterim) -Message "Expected UTF-8 mojibake to be repaired with French accents."
    $terminalMojibakeInterim = "Int" + [string][char]0x251C + [string][char]0x00AE + "rim - 6 Mois"
    Assert-ScoringCondition -Condition ((Repair-DisplayText $terminalMojibakeInterim) -eq $expectedInterim) -Message "Expected terminal mojibake to be repaired with French accents."
    $oemMojibakeLocation = "CDI " + [string][char]0x251C + [string][char]0x00E1 + " Paris"
    $expectedLocation = "CDI " + [string][char]0x00E0 + " Paris"
    Assert-ScoringCondition -Condition ((Repair-DisplayText $oemMojibakeLocation) -eq $expectedLocation) -Message "Expected OEM mojibake to be repaired with French accents."
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType "CDD") -Message "Expected CDD contracts to be excluded."
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType "Freelance") -Message "Expected freelance contracts to be excluded."

    $annonceurMatch = Get-JobMatch -Title "CRM Analyst" -Text "Braze HubSpot segmentation customer journey dashboarding"
    Assert-ScoringCondition -Condition $annonceurMatch.IsMatch -Message "Expected a CRM Analyst role with configured profile skills to match."
    $expandedToolMatch = Get-JobMatch -Title "Marketing Automation Specialist" -Text "Braze HubSpot A/B testing customer lifecycle"
    Assert-ScoringCondition -Condition $expandedToolMatch.IsMatch -Message "Expected configured CRM/lifecycle signals to match."
    Assert-ScoringCondition -Condition ($expandedToolMatch.Keywords -match "Braze" -and $expandedToolMatch.Keywords -match "HubSpot" -and $expandedToolMatch.Keywords -match "A/B testing") -Message "Expected configured profile skill keywords to be reported."
    $positiveFeedbackRow = New-OrderedJobRecord @{
        status           = "interesting"
        job_title        = "CRM Analyst"
        matched_keywords = "Braze; customer lifecycle"
    }
    $ignoredFeedbackRow = New-OrderedJobRecord @{
        status    = "ignored"
        job_title = "SEO Manager"
        notes     = "ignore_reason=profile_exclusions; detail=too marketing"
    }
    $script:FeedbackLearningProfile = New-FeedbackLearningProfile -Rows @($positiveFeedbackRow, $ignoredFeedbackRow)
    $positiveLearning = Get-FeedbackLearningAdjustment -FullText "braze customer lifecycle" -HasCoreTitleSignal:$true -HasProfileSkillSignal:$true -HasProfileContext:$true
    Assert-ScoringCondition -Condition ([int]$positiveLearning.Adjustment -gt 0 -and (($positiveLearning.Reasons -join ";") -match "Braze")) -Message "Expected positive saved tracker feedback to boost similar configured skill signals."
    $negativeLearning = Get-FeedbackLearningAdjustment -FullText "seo sea paid media campaign" -HasCoreTitleSignal:$false -HasProfileSkillSignal:$false -HasProfileContext:$false
    Assert-ScoringCondition -Condition ([int]$negativeLearning.Adjustment -lt 0 -and (($negativeLearning.Reasons -join ";") -match "profile exclusions")) -Message "Expected ignored saved tracker feedback to penalize similar configured exclusion signals."
    $script:FeedbackLearningProfile = $null
    $annonceurResult = New-JobResult `
        -Title "CRM Analyst" `
        -CompanyName "Radio France" `
        -JobLocation "Paris" `
        -ContractType "CDI" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/radio-france-crm-analyst" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Braze HubSpot segmentation customer journey dashboarding"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $annonceurResult -Name "employer_type") -eq "annonceur") -Message "Expected Radio France to be classified as annonceur."
    Assert-ScoringCondition -Condition ((Get-IntegerRowValue -Row $annonceurResult -Name "match_score") -gt (Get-IntegerRowValue -Row $annonceurResult -Name "role_score")) -Message "Expected annonceur/Paris/CDI fit to boost the role score."

    $consultingMatch = Get-JobMatch -Title "CRM Consultant" -Text "Braze HubSpot SQL marketing automation"
    $consultingResult = New-JobResult `
        -Title "CRM Consultant" `
        -CompanyName "fifty-five" `
        -JobLocation "Paris" `
        -ContractType "CDI" `
        -MatchScore $consultingMatch.Score `
        -MatchLevel $consultingMatch.Level `
        -MatchedKeywords $consultingMatch.Keywords `
        -Url "https://example.test/jobs/fifty-five-crm-consultant" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Braze HubSpot SQL marketing automation"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $consultingResult -Name "employer_type") -eq "consulting") -Message "Expected fifty-five to be classified as consulting."
    Assert-ScoringCondition -Condition ((Get-IntegerRowValue -Row $consultingResult -Name "employer_fit") -lt 0) -Message "Expected consulting employer type to be demoted, not excluded."

    $dataEngineeringMatch = Get-JobMatch -Title "Data Engineer" -Text "python dbt snowflake airflow data warehouse data pipeline"
    Assert-ScoringCondition -Condition (-not $dataEngineeringMatch.IsMatch) -Message "Expected excluded engineering role without profile signals to stay below the match threshold."
    $companyNameOnlyToolMatch = Get-JobMatch -Title "People Business Partner" -Text "HubSpot Paris Full-time"
    Assert-ScoringCondition -Condition (-not $companyNameOnlyToolMatch.IsMatch) -Message "Expected a non-profile role not to match only because text contains one configured skill."

    $titleOnlyExcludedContract = New-JobResult `
        -Title "Alternance Assistant CRM" `
        -CompanyName "Example Company" `
        -JobLocation "Paris" `
        -ContractType "" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/alternance-crm" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Braze marketing automation"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $titleOnlyExcludedContract -Name "contract_type") -eq "Apprenticeship") -Message "Expected title-only alternance to be mapped to Apprenticeship."
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType (Get-RowValue -Row $titleOnlyExcludedContract -Name "contract_type")) -Message "Expected title-only alternance to be excluded by contract filtering."
    $titleOverridesGenericContract = New-JobResult `
        -Title "STAGE - Communication CRM" `
        -CompanyName "Example Company" `
        -JobLocation "Paris" `
        -ContractType "Full-time" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/stage-crm" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Braze marketing automation"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $titleOverridesGenericContract -Name "contract_type") -eq "Internship") -Message "Expected explicit STAGE title to override generic Full-time contract."

    $junkLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/fr/companies/acme/jobs/crm-analyst_5Kvvowa"
    Assert-ScoringCondition -Condition ([string]::IsNullOrWhiteSpace($junkLocation)) -Message "Expected random WTTJ URL suffixes not to become city names."
    $parisLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/fr/companies/acme/jobs/crm-analyst_paris"
    Assert-ScoringCondition -Condition ($parisLocation -eq "Paris") -Message "Expected readable WTTJ city suffix to be kept."
    $multiPartUrlLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/en/companies/acme/jobs/crm-analyst_london_ACME_3zgazPX"
    Assert-ScoringCondition -Condition ($multiPartUrlLocation -eq "London") -Message "Expected WTTJ URL city suffixes before reference tokens to be parsed."
    $wttjInitialDataHtml = 'window.__INITIAL_DATA__ = "{\"queries\":[{\"state\":{\"data\":{\"offices\":[{\"city\":\"Saint-Denis\",\"country_code\":\"FR\"}]}}}]}";'
    $wttjInitialDataLocation = Get-WttjLocation -Html $wttjInitialDataHtml -Url "https://www.welcometothejungle.com/fr/companies/acme/jobs/crm-analyst_5Kvvowa" -Title "CRM Analyst"
    Assert-ScoringCondition -Condition ($wttjInitialDataLocation -eq "Saint-Denis, France") -Message "Expected WTTJ embedded office city and country code to be parsed."
    Assert-ScoringCondition -Condition (-not (Test-IsWttjLocationAllowed -JobLocation "New York, United States" -Url "https://www.welcometothejungle.com/fr/companies/acme/jobs/crm-analyst_new-york" -Text "CRM Analyst")) -Message "Expected foreign WTTJ locations to be rejected for France crawls."
    Assert-ScoringCondition -Condition (-not (Test-IsWttjLocationAllowed -JobLocation "" -Url "https://www.welcometothejungle.com/fr/companies/acme/jobs/crm-analyst_5Kvvowa" -Text "CRM Analyst")) -Message "Expected blank WTTJ locations to be rejected for France crawls."
    $invalidExistingWttjRow = New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "CRM Analyst"
        company_name   = "Example Company"
        location       = ""
        contract_type  = "CDI"
        match_score    = "80"
        match_level    = "Good"
        job_url_raw    = "https://www.welcometothejungle.com/en/companies/acme/jobs/crm-analyst_london_ACME_3zgazPX"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }
    $foreignExistingWttjRow = New-OrderedJobRecord @{
        status         = "ignored"
        job_title      = "CRM Analyst Casablanca"
        company_name   = "Example Company"
        location       = "Casablanca"
        contract_type  = "CDI"
        match_score    = "80"
        match_level    = "Good"
        job_url_raw    = "https://www.welcometothejungle.com/fr/companies/acme/jobs/crm-analyst_casablanca"
        platform       = "Welcome to the Jungle"
        published_date = ([DateTimeOffset]::Now.ToString("yyyy-MM-dd"))
    }
    $invalidMerge = Merge-JobsWithTracker -CurrentRows @() -ExistingRows @($invalidExistingWttjRow, $foreignExistingWttjRow) -Path "selftest.xlsx" -SkipBackup
    Assert-ScoringCondition -Condition ([int]$invalidMerge.RemovedCount -eq 2 -and @($invalidMerge.TrackerRows).Count -eq 0) -Message "Expected invalid existing non-applied WTTJ rows to be removed during merge."

    $franceTravailMock = [PSCustomObject]@{
        id                  = "123ABC"
        intitule            = "CRM Analyst"
        description         = "Braze HubSpot segmentation customer journey"
        dateCreation        = ([DateTimeOffset]::Now.ToString("o"))
        typeContrat         = "CDI"
        typeContratLibelle  = "CDI"
        urlPostulation      = "https://candidat.francetravail.fr/offres/recherche/detail/123ABC"
        lieuTravail         = [PSCustomObject]@{ libelle = "75 - Paris" }
        entreprise          = [PSCustomObject]@{ nom = "Example Annonceur" }
    }
    Assert-ScoringCondition -Condition ((Get-FranceTravailContractType $franceTravailMock) -eq "CDI") -Message "Expected France Travail CDI contract mapping."
    Assert-ScoringCondition -Condition ((Get-FranceTravailCompanyName $franceTravailMock) -eq "Example Annonceur") -Message "Expected France Travail company mapping."
    Assert-ScoringCondition -Condition ((Get-FranceTravailLocation $franceTravailMock) -eq "75 - Paris") -Message "Expected France Travail location mapping."

    $adzunaMock = [PSCustomObject]@{
        title         = "CRM Analyst"
        description   = "Braze HubSpot SQL marketing automation"
        created       = ([DateTimeOffset]::Now.ToString("o"))
        redirect_url  = "https://www.adzuna.fr/details/123"
        contract_type = "permanent"
        contract_time = "full_time"
        company       = [PSCustomObject]@{ display_name = "Example Retailer" }
        location      = [PSCustomObject]@{ display_name = "Paris, Ile-de-France" }
    }
    Assert-ScoringCondition -Condition ((Get-AdzunaContractType $adzunaMock) -eq "Permanent") -Message "Expected Adzuna permanent contract mapping."
    Assert-ScoringCondition -Condition ((Get-AdzunaCompanyName $adzunaMock) -eq "Example Retailer") -Message "Expected Adzuna company mapping."
    Assert-ScoringCondition -Condition ((Get-AdzunaLocation $adzunaMock) -eq "Paris, Ile-de-France") -Message "Expected Adzuna location mapping."

    $apecMock = [PSCustomObject]@{
        id              = 123456789
        numeroOffre     = "123456789W"
        intitule        = "CRM Analyst F/H"
        nomCommercial   = "Example Retailer"
        lieuTexte       = "Paris - 75"
        typeContrat     = 101888
        texteOffre      = "Braze HubSpot SQL marketing automation"
        datePublication = ([DateTimeOffset]::Now.ToString("yyyy-MM-ddTHH:mm:ss.000+0000"))
    }
    Assert-ScoringCondition -Condition ((Get-ApecContractType $apecMock) -eq "CDI") -Message "Expected APEC CDI contract mapping."
    Assert-ScoringCondition -Condition ((Get-ApecJobUrl $apecMock) -match "/detail-offre/123456789W$") -Message "Expected APEC detail URL mapping."

    $helloWorkMockHtml = @'
<script type="application/ld+json">{"@context":"https://schema.org","@type":"JobPosting","title":"CRM Analyst H/F","description":"Braze HubSpot segmentation","datePosted":"2026-06-16T09:38:15Z","employmentType":"FULL_TIME","hiringOrganization":{"@type":"Organization","name":"Example Retailer"},"jobLocation":{"@type":"Place","address":{"@type":"PostalAddress","addressLocality":"Paris","addressRegion":"Ile-de-France","addressCountry":"FR"}}}</script>
<script type="application/ld+json">{"JobTitle":"CRM Analyst H/F","Company":"Example Retailer","Localisation":"Paris - 75","ContractType":"CDI","Description":"Braze HubSpot customer journey"}</script>
'@
    $helloWorkMetadata = Get-HelloWorkJobMetadata -Html $helloWorkMockHtml
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Title -eq "CRM Analyst H/F") -Message "Expected HelloWork title metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Company -eq "Example Retailer") -Message "Expected HelloWork company metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Location -eq "Paris - 75") -Message "Expected HelloWork custom location metadata to be preferred."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Contract -eq "CDI") -Message "Expected HelloWork contract metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Description -match "Braze") -Message "Expected HelloWork custom description metadata mapping."
    Assert-ScoringCondition -Condition (Test-IsRecent (ConvertFrom-FrenchRelativeDateText "il y a 2 jours")) -Message "Expected French relative dates to parse as recent."

    $crossPlatformMatch = Get-JobMatch -Title "CRM Analyst" -Text "Braze HubSpot segmentation customer journey"
    $crossPlatformRows = @(
        (New-JobResult -Title "CRM Analyst" -CompanyName "Radio France" -JobLocation "Paris" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.linkedin.com/jobs/view/111" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Braze HubSpot segmentation customer journey"),
        (New-JobResult -Title "Analyste CRM H/F" -CompanyName "Radio France" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://candidat.francetravail.fr/offres/recherche/detail/111" -Platform "France Travail" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Braze HubSpot segmentation customer journey"),
        (New-JobResult -Title "CRM Analyst" -CompanyName "Radio France" -JobLocation "Paris, Ile-de-France" -ContractType "Permanent" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.adzuna.fr/details/111" -Platform "Adzuna" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Braze HubSpot segmentation customer journey"),
        (New-JobResult -Title "CRM Analyst F/H" -CompanyName "Radio France" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/111W" -Platform "APEC" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Braze HubSpot segmentation customer journey"),
        (New-JobResult -Title "CRM Analyst H/F" -CompanyName "Radio France" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/111.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "Braze HubSpot segmentation customer journey")
    )
    $crossPlatformKeys = @($crossPlatformRows | ForEach-Object { Get-JobDedupeKeyFromRow $_ } | Select-Object -Unique)
    Assert-ScoringCondition -Condition ($crossPlatformKeys.Count -eq 1) -Message "Expected same company/title role from several platforms to share one dedupe key."
    $crossPlatformMerged = Merge-SimilarJobRows -Rows $crossPlatformRows -Reason "test cross-platform duplicate"
    $crossPlatformSources = Get-RowValue -Row $crossPlatformMerged -Name "platform"
    Assert-ScoringCondition -Condition ($crossPlatformSources -match "LinkedIn" -and $crossPlatformSources -match "France Travail" -and $crossPlatformSources -match "Adzuna" -and $crossPlatformSources -match "APEC" -and $crossPlatformSources -match "HelloWork") -Message "Expected merged cross-platform row to keep all source names."
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $crossPlatformMerged -Name "source_count") -eq "5") -Message "Expected source_count to count unique platforms."
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $crossPlatformMerged -Name "job_url_raw") -match "apec") -Message "Expected APEC URL to be preferred over LinkedIn, France Travail, HelloWork, and Adzuna for this merge."
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "linkedin" -and (Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "adzuna" -and (Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "hellowork" -and (Get-RowValue -Row $crossPlatformMerged -Name "alternate_urls") -match "francetravail") -Message "Expected alternate URLs to keep non-primary cross-platform links."

    Write-Host "Scoring self-test passed."
}

