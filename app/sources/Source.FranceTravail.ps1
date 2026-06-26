function Get-FranceTravailAccessToken {
    if ([string]::IsNullOrWhiteSpace($FranceTravailClientId) -or [string]::IsNullOrWhiteSpace($FranceTravailClientSecret)) {
        Write-RunStatus "France Travail credentials not set; skipping France Travail source. Set FRANCE_TRAVAIL_CLIENT_ID and FRANCE_TRAVAIL_CLIENT_SECRET to enable it."
        return ""
    }

    $tokenUrl = $env:FRANCE_TRAVAIL_TOKEN_URL
    if ([string]::IsNullOrWhiteSpace($tokenUrl)) {
        $tokenUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.france_travail_token" -DefaultValue "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=/partenaire")
    }

    try {
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $FranceTravailClientId
            client_secret = $FranceTravailClientSecret
            scope         = $FranceTravailScope
        }
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded" -TimeoutSec 45
        return [string]$response.access_token
    }
    catch {
        Write-Warning ("France Travail token request failed: {0}" -f $_.Exception.Message)
        return ""
    }
}

function Get-FranceTravailJobUrl {
    param([AllowNull()]$Job)

    $urlPostulation = Get-ObjectPropertyValue -Object $Job -Names @("urlPostulation")
    if (-not [string]::IsNullOrWhiteSpace($urlPostulation)) {
        return ConvertTo-CleanUrl ([string]$urlPostulation)
    }

    $origin = Get-ObjectPropertyValue -Object $Job -Names @("origineOffre")
    $originUrl = Get-ObjectPropertyValue -Object $origin -Names @("urlOrigine")
    if (-not [string]::IsNullOrWhiteSpace($originUrl)) {
        return ConvertTo-CleanUrl ([string]$originUrl)
    }

    $jobId = Get-ObjectPropertyValue -Object $Job -Names @("id")
    if (-not [string]::IsNullOrWhiteSpace($jobId)) {
        $template = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.france_travail_detail" -DefaultValue "https://candidat.francetravail.fr/offres/recherche/detail/{id}")
        return $template.Replace("{id}", [Uri]::EscapeDataString([string]$jobId))
    }

    return ""
}

function Get-FranceTravailCompanyName {
    param([AllowNull()]$Job)

    $company = Get-ObjectPropertyValue -Object $Job -Names @("entreprise")
    $companyName = Get-ObjectPropertyValue -Object $company -Names @("nom", "name")
    if (-not [string]::IsNullOrWhiteSpace($companyName)) {
        return [string]$companyName
    }

    $origin = Get-ObjectPropertyValue -Object $Job -Names @("origineOffre")
    $originName = Get-ObjectPropertyValue -Object $origin -Names @("origine", "nom")
    if (-not [string]::IsNullOrWhiteSpace($originName) -and -not (Test-IsGenericJobBoardName $originName)) {
        return [string]$originName
    }

    return ""
}

function Get-FranceTravailLocation {
    param([AllowNull()]$Job)

    $workLocation = Get-ObjectPropertyValue -Object $Job -Names @("lieuTravail")
    $location = Get-ObjectPropertyValue -Object $workLocation -Names @("libelle", "commune", "codePostal")
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        return [string]$location
    }

    return ConvertTo-LocationText $workLocation
}

function Get-FranceTravailContractType {
    param([AllowNull()]$Job)

    $contractLabel = Get-ObjectPropertyValue -Object $Job -Names @("typeContratLibelle", "contratLibelle")
    $contractCode = Get-ObjectPropertyValue -Object $Job -Names @("typeContrat", "contrat")
    $contractText = Join-CleanTextParts @($contractLabel, $contractCode)
    $contractType = Get-ContractTypeFromText -Text $contractText
    if (-not [string]::IsNullOrWhiteSpace($contractType)) {
        return $contractType
    }

    return [string]$contractLabel
}

function Get-FranceTravailPublishedAt {
    param([AllowNull()]$Job)

    foreach ($name in @("dateCreation", "dateActualisation")) {
        $value = Get-ObjectPropertyValue -Object $Job -Names @($name)
        $date = ConvertTo-DateTimeOffsetOrNull $value
        if ($null -ne $date) {
            return $date
        }
    }

    return $null
}

function Get-FranceTravailSourceText {
    param([AllowNull()]$Job)

    return Join-CleanTextParts @(
        (Get-ObjectPropertyValue -Object $Job -Names @("intitule", "title")),
        (Get-ObjectPropertyValue -Object $Job -Names @("description")),
        (Get-ObjectPropertyValue -Object $Job -Names @("profil")),
        (Get-ObjectPropertyValue -Object $Job -Names @("competences")),
        (Get-ObjectPropertyValue -Object $Job -Names @("qualitesProfessionnelles"))
    )
}

function Get-FranceTravailSearchDelayMilliseconds {
    $delayVariable = Get-Variable -Name "FranceTravailSearchDelayMilliseconds" -ErrorAction SilentlyContinue
    if ($null -ne $delayVariable -and $null -ne $delayVariable.Value) {
        return [Math]::Max(0, [int]$delayVariable.Value)
    }

    return 300
}

function Start-FranceTravailSearchDelay {
    $delayMilliseconds = Get-FranceTravailSearchDelayMilliseconds
    if ($delayMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $delayMilliseconds
    }
}

function Get-FranceTravailHttpErrorSummary {
    param([AllowNull()][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $statusCode = 0
    $retryAfterMilliseconds = 0
    $body = ""
    $message = ""

    if ($null -ne $ErrorRecord) {
        $message = [string]$ErrorRecord.Exception.Message
        try {
            $response = $ErrorRecord.Exception.Response
            if ($null -ne $response) {
                try {
                    $statusCode = [int]$response.StatusCode
                }
                catch {
                    $statusCode = 0
                }

                try {
                    $retryAfter = [string]$response.Headers["Retry-After"]
                    $retryAfterSeconds = 0
                    if ([int]::TryParse($retryAfter, [ref]$retryAfterSeconds) -and $retryAfterSeconds -gt 0) {
                        $retryAfterMilliseconds = $retryAfterSeconds * 1000
                    }
                }
                catch {
                    $retryAfterMilliseconds = 0
                }

                try {
                    $stream = $response.GetResponseStream()
                    if ($null -ne $stream) {
                        $reader = New-Object System.IO.StreamReader($stream)
                        $body = [string]$reader.ReadToEnd()
                    }
                }
                catch {
                    $body = ""
                }
            }
        }
        catch {
            $statusCode = 0
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($body)) {
        $body = Repair-DisplayText $body
        $body = [regex]::Replace($body, "\s+", " ").Trim()
        if ($body.Length -gt 500) {
            $body = $body.Substring(0, 500)
        }
    }

    return [pscustomobject]@{
        StatusCode = $statusCode
        Message = $message
        Body = $body
        RetryAfterMilliseconds = $retryAfterMilliseconds
    }
}

function Invoke-FranceTravailSearch {
    param(
        [string]$Url,
        [hashtable]$Headers,
        [string]$Query,
        [int]$Page,
        [AllowNull()]$Stats
    )

    $maxRetries = [int](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.max_retries" -DefaultValue 2)
    $retryDelayMilliseconds = [int](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.retry_delay_ms" -DefaultValue 1200)
    $retryBackoffMultiplier = [double](Get-ConfigPathValue -Object $script:JobCrawlerRuntimeConfig -Path "http.retry_backoff_multiplier" -DefaultValue 2.0)
    $attemptCount = [Math]::Max(1, $maxRetries + 1)

    for ($attempt = 1; $attempt -le $attemptCount; $attempt++) {
        try {
            Add-SourceMetric -Stats $Stats -Name "SearchRequests"
            $response = Invoke-RestMethod -Uri $Url -Headers $Headers -Method Get -TimeoutSec 45
            Start-FranceTravailSearchDelay
            return [pscustomobject]@{
                Success = $true
                Response = $response
            }
        }
        catch {
            Add-SourceMetric -Stats $Stats -Name "Errors"
            $summary = Get-FranceTravailHttpErrorSummary $_
            $detail = $summary.Message
            if ($summary.StatusCode -gt 0) {
                $detail = "HTTP {0}; {1}" -f $summary.StatusCode, $detail
            }
            if (-not [string]::IsNullOrWhiteSpace($summary.Body)) {
                $detail = "{0}; body: {1}" -f $detail, $summary.Body
            }

            $isRetryable = ($summary.StatusCode -eq 429 -or $summary.StatusCode -ge 500)
            if ($isRetryable -and $attempt -lt $attemptCount) {
                $delayMilliseconds = $summary.RetryAfterMilliseconds
                if ($delayMilliseconds -le 0) {
                    $delayMilliseconds = [int]($retryDelayMilliseconds * [Math]::Pow($retryBackoffMultiplier, ($attempt - 1)))
                }
                Write-Warning ("France Travail search failed for '{0}' page {1}, attempt {2}/{3}; retrying in {4}ms: {5}" -f $Query, $Page, $attempt, $attemptCount, $delayMilliseconds, $detail)
                Start-Sleep -Milliseconds $delayMilliseconds
                continue
            }

            Write-Warning ("France Travail search failed for '{0}' page {1}, attempt {2}/{3}: {4}" -f $Query, $Page, $attempt, $attemptCount, $detail)
            Start-FranceTravailSearchDelay
            return [pscustomobject]@{
                Success = $false
                Response = $null
            }
        }
    }

    return [pscustomobject]@{
        Success = $false
        Response = $null
    }
}

function Get-FranceTravailJobs {
    $accessToken = Get-FranceTravailAccessToken
    if ([string]::IsNullOrWhiteSpace($accessToken)) {
        return @()
    }

    Set-RunWindowTitle "Custom Job Tracker - France Travail"
    Write-RunStatus "Collecting France Travail jobs through the official API..."
    Write-RunStatus ("France Travail plan: {0} query/queries, up to {1} page(s) each." -f $FranceTravailQueries.Count, $MaxFranceTravailPages)
    $stats = Start-SourceStats "France Travail"
    $results = New-Object System.Collections.Generic.List[object]
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Accept"        = "application/json"
    }
    $searchUrl = [string](Get-ConfigPathValue -Object $script:JobCrawlerSourcesConfig -Path "endpoints.france_travail_jobs" -DefaultValue "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search")
    $pageSize = 150
    $queryIndex = 0

    foreach ($query in $FranceTravailQueries) {
        $queryIndex++
        Write-RunStatus ("France Travail query {0}/{1}: {2}" -f $queryIndex, $FranceTravailQueries.Count, $query)
        for ($page = 0; $page -lt $MaxFranceTravailPages; $page++) {
            $rangeStart = $page * $pageSize
            $rangeEnd = $rangeStart + $pageSize - 1
            $params = @{
                motsCles      = $query
                publieeDepuis = [string][Math]::Abs($DaysBack)
                range         = ("{0}-{1}" -f $rangeStart, $rangeEnd)
                sort          = "1"
            }

            if (-not [string]::IsNullOrWhiteSpace($Location) -and $Location -notmatch "(?i)^france$") {
                $params["lieu"] = $Location
                $params["distance"] = "50"
            }

            $url = "{0}?{1}" -f $searchUrl, (ConvertTo-QueryString $params)
            $searchResult = Invoke-FranceTravailSearch -Url $url -Headers $headers -Query $query -Page ($page + 1) -Stats $stats
            if (-not $searchResult.Success) {
                break
            }
            $response = $searchResult.Response

            $jobArray = @()
            if ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "resultats") {
                $jobArray = @($response.resultats)
            }
            elseif ($null -ne $response) {
                $jobArray = @($response)
            }

            if ($jobArray.Count -eq 0) {
                break
            }

            foreach ($job in $jobArray) {
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $publishedAt = Get-FranceTravailPublishedAt $job
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $title = [string](Get-ObjectPropertyValue -Object $job -Names @("intitule", "title"))
                $sourceText = Get-FranceTravailSourceText $job
                $contractType = Get-FranceTravailContractType $job
                if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text (Join-CleanTextParts @($title, $sourceText)) -Reliable) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $match = Get-JobMatch -Title $title -Text $sourceText
                if (-not $match.IsMatch) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $jobUrl = Get-FranceTravailJobUrl $job
                $companyName = Get-FranceTravailCompanyName $job
                $jobLocation = Get-FranceTravailLocation $job
                $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "France Travail" -PublishedAt $publishedAt -SourceText $sourceText
                if ($null -ne $result) {
                    $results.Add($result) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }

            Write-CountProgress -Activity ("France Travail query {0}/{1}" -f $queryIndex, $FranceTravailQueries.Count) -Current ($page + 1) -Total $MaxFranceTravailPages -Found $results.Count -Every 1
            if ($jobArray.Count -lt $pageSize) {
                break
            }
        }
    }

    Write-RunStatus ("France Travail complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

