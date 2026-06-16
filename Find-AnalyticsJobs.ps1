[CmdletBinding()]
param(
    [int]$DaysBack = 7,
    [string]$Location = "France",
    [string]$TrackerPath = "",
    [string]$WelcomeKitApiKey = $env:WK_API_KEY,
    [string]$FranceTravailClientId = $env:FRANCE_TRAVAIL_CLIENT_ID,
    [string]$FranceTravailClientSecret = $env:FRANCE_TRAVAIL_CLIENT_SECRET,
    [string]$FranceTravailScope = $(if ([string]::IsNullOrWhiteSpace($env:FRANCE_TRAVAIL_SCOPE)) { "api_offresdemploiv2 o2dsoffre" } else { $env:FRANCE_TRAVAIL_SCOPE }),
    [string]$AdzunaAppId = $env:ADZUNA_APP_ID,
    [string]$AdzunaAppKey = $env:ADZUNA_APP_KEY,
    [ValidateSet("Fast", "Default", "Deep")]
    [string]$CrawlMode = "Default",
    [int]$MaxLinkedInSearchPages = 3,
    [int]$MaxLinkedInDetails = 0,
    [int]$MaxFranceTravailPages = 2,
    [int]$MaxAdzunaPages = 1,
    [int]$MaxApecPages = 2,
    [int]$MaxHelloWorkPages = 1,
    [int]$MaxHelloWorkCardsPerQuery = 20,
    [int]$MaxHelloWorkDetails = 50,
    [int]$MaxWelcomeKitPages = 10,
    [int]$MaxWttjCandidatePages = 120,
    [int]$MaxBackups = 5,
    [switch]$SkipFranceTravail,
    [switch]$SkipAdzuna,
    [switch]$SkipApec,
    [switch]$SkipHelloWork,
    [switch]$SkipLinkedIn,
    [switch]$SkipWttj,
    [switch]$DisableWttjPublicFallback,
    [switch]$DisableCache,
    [int]$CacheTtlHours = 24,
    [switch]$SelfTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JobTracker.Common.ps1")

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$BrowserUserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
$Cutoff = [DateTimeOffset]::Now.AddDays(-[Math]::Abs($DaysBack))
$CutoffDate = $Cutoff.ToString("yyyy-MM-dd")
$RunDate = Get-Date -Format "yyyy-MM-dd"
$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$DefaultTrackerPath = Join-Path $PSScriptRoot "output\jobs_tracker.xlsx"
$CacheDirectory = Join-Path $PSScriptRoot "output\cache"

switch ($CrawlMode) {
    "Fast" {
        if (-not $PSBoundParameters.ContainsKey("MaxLinkedInSearchPages")) { $MaxLinkedInSearchPages = 2 }
        if (-not $PSBoundParameters.ContainsKey("MaxLinkedInDetails")) { $MaxLinkedInDetails = 35 }
        if (-not $PSBoundParameters.ContainsKey("MaxFranceTravailPages")) { $MaxFranceTravailPages = 1 }
        if (-not $PSBoundParameters.ContainsKey("MaxAdzunaPages")) { $MaxAdzunaPages = 1 }
        if (-not $PSBoundParameters.ContainsKey("MaxApecPages")) { $MaxApecPages = 1 }
        if (-not $PSBoundParameters.ContainsKey("MaxHelloWorkPages")) { $MaxHelloWorkPages = 1 }
        if (-not $PSBoundParameters.ContainsKey("MaxHelloWorkCardsPerQuery")) { $MaxHelloWorkCardsPerQuery = 12 }
        if (-not $PSBoundParameters.ContainsKey("MaxHelloWorkDetails")) { $MaxHelloWorkDetails = 25 }
        if (-not $PSBoundParameters.ContainsKey("MaxWelcomeKitPages")) { $MaxWelcomeKitPages = 6 }
        if (-not $PSBoundParameters.ContainsKey("MaxWttjCandidatePages")) { $MaxWttjCandidatePages = 50 }
    }
    "Deep" {
        if (-not $PSBoundParameters.ContainsKey("MaxLinkedInSearchPages")) { $MaxLinkedInSearchPages = 4 }
        if (-not $PSBoundParameters.ContainsKey("MaxLinkedInDetails")) { $MaxLinkedInDetails = 160 }
        if (-not $PSBoundParameters.ContainsKey("MaxFranceTravailPages")) { $MaxFranceTravailPages = 3 }
        if (-not $PSBoundParameters.ContainsKey("MaxAdzunaPages")) { $MaxAdzunaPages = 2 }
        if (-not $PSBoundParameters.ContainsKey("MaxApecPages")) { $MaxApecPages = 3 }
        if (-not $PSBoundParameters.ContainsKey("MaxHelloWorkPages")) { $MaxHelloWorkPages = 2 }
        if (-not $PSBoundParameters.ContainsKey("MaxHelloWorkCardsPerQuery")) { $MaxHelloWorkCardsPerQuery = 30 }
        if (-not $PSBoundParameters.ContainsKey("MaxHelloWorkDetails")) { $MaxHelloWorkDetails = 120 }
        if (-not $PSBoundParameters.ContainsKey("MaxWelcomeKitPages")) { $MaxWelcomeKitPages = 15 }
        if (-not $PSBoundParameters.ContainsKey("MaxWttjCandidatePages")) { $MaxWttjCandidatePages = 220 }
    }
    default {
        if (-not $PSBoundParameters.ContainsKey("MaxLinkedInDetails")) { $MaxLinkedInDetails = 80 }
    }
}

if ([string]::IsNullOrWhiteSpace($TrackerPath)) {
    $TrackerPath = $DefaultTrackerPath
}
if ([IO.Path]::GetExtension($TrackerPath).ToLowerInvariant() -ne ".xlsx") {
    throw "This crawler uses only the XLSX tracker file. Use output\jobs_tracker.xlsx for -TrackerPath."
}

$SeenResultKeys = @{}
$LinkedInDelayMilliseconds = 1200
$AdzunaDelayMilliseconds = 2500
$ApecDelayMilliseconds = 300
$HelloWorkSearchDelayMilliseconds = 350
$HelloWorkDetailDelayMilliseconds = 450
$MinimumMatchScore = 35
$JobCrawlerPreferences = $null
$FeedbackLearningProfile = $null
$SourceRunStats = New-Object System.Collections.Generic.List[object]

$WttjUrlCandidatePattern = "(?i)(web[-_\s]*analyst|digital[-_\s]*analyst|web[-_\s]*analytics|digital[-_\s]*analytics|analytics[-_\s]*(consultant|specialist|manager|engineer)|tracking|taggage|tagging|data[-_\s]*analyst|(^|[-_\s])ga4($|[-_\s])|(^|[-_\s])gtm($|[-_\s])|google[-_\s]*(analytics|tag[-_\s]*manager)|piano|contentsquare|content[-_\s]*square|tag[-_\s]*commander|commanders?[-_\s]*act|tealium|server[-_\s]*side|(^|[-_\s])rgpd($|[-_\s])|(^|[-_\s])gdpr($|[-_\s])|(^|[-_\s])cro($|[-_\s]))"

$LinkedInQueries = @(
    "web analyst google tag manager",
    "web analyst google analytics",
    "web analyst cro google analytics",
    "digital analyst google tag manager",
    "digital analyst google analytics",
    "digital analytics consultant",
    "senior digital analytics consultant",
    "tracking analyst ga4 gtm",
    "tracking analytics specialist",
    "analytics consultant google tag manager",
    "performance digital google analytics",
    "charge performance digital google analytics",
    "piano analytics",
    "contentsquare analytics",
    "tag commander analytics",
    "commanders act analytics",
    "tealium analytics",
    "server side tracking",
    "rgpd analytics",
    "tagging plan analytics"
)

$ApiSearchQueries = @(
    "web analyst",
    "digital analyst",
    "web analytics",
    "digital analytics",
    "tracking analyst",
    "tracking analytics",
    "google analytics",
    "google tag manager",
    "piano analytics",
    "contentsquare",
    "tag commander",
    "commanders act",
    "tealium",
    "server side tracking",
    "rgpd analytics",
    "cro analytics",
    "performance digital analytics"
)

function Repair-DisplayText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $clean = [string]$Text
    $clean = $clean.Replace(([string][char]0x00A0), " ")

    $mojibakeScore = {
        param([string]$Value)

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return 0
        }

        $score = 0
        foreach ($marker in @(
            [string][char]0x00C2,
            [string][char]0x00C3,
            [string][char]0x00E2,
            [string][char]0x251C,
            [string][char]0xFFFD
        )) {
            $score += [regex]::Matches($Value, [regex]::Escape($marker)).Count
        }

        return $score
    }

    if (& $mojibakeScore $clean) {
        foreach ($codePage in @(1252, 850, 437)) {
            try {
                $sourceEncoding = [Text.Encoding]::GetEncoding($codePage)
                $decoded = [Text.Encoding]::UTF8.GetString($sourceEncoding.GetBytes($clean))
                if (-not [string]::IsNullOrWhiteSpace($decoded) -and (& $mojibakeScore $decoded) -lt (& $mojibakeScore $clean)) {
                    $clean = $decoded
                    break
                }
            }
            catch {
            }
        }
    }

    $mojibakeReplacements = @(
        @{ From = ([string][char]0x00C3 + [string][char]0x0080); To = ([string][char]0x00C0) },
        @{ From = ([string][char]0x00C3 + [string][char]0x0087); To = ([string][char]0x00C7) },
        @{ From = ([string][char]0x00C3 + [string][char]0x0088); To = ([string][char]0x00C8) },
        @{ From = ([string][char]0x00C3 + [string][char]0x0089); To = ([string][char]0x00C9) },
        @{ From = ([string][char]0x00C3 + [string][char]0x008A); To = ([string][char]0x00CA) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A0); To = ([string][char]0x00E0) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A2); To = ([string][char]0x00E2) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A7); To = ([string][char]0x00E7) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A8); To = ([string][char]0x00E8) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00A9); To = ([string][char]0x00E9) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AA); To = ([string][char]0x00EA) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AB); To = ([string][char]0x00EB) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AE); To = ([string][char]0x00EE) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00AF); To = ([string][char]0x00EF) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00B4); To = ([string][char]0x00F4) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00B9); To = ([string][char]0x00F9) },
        @{ From = ([string][char]0x00C3 + [string][char]0x00BB); To = ([string][char]0x00FB) },
        @{ From = ([string][char]0x251C + [string][char]0x00A1); To = ([string][char]0x00E0) },
        @{ From = ([string][char]0x251C + [string][char]0x00E1); To = ([string][char]0x00E0) },
        @{ From = ([string][char]0x251C + [string][char]0x00A7); To = ([string][char]0x00E7) },
        @{ From = ([string][char]0x251C + [string][char]0x00A8); To = ([string][char]0x00E8) },
        @{ From = ([string][char]0x251C + [string][char]0x00A9); To = ([string][char]0x00E9) },
        @{ From = ([string][char]0x251C + [string][char]0x00AE); To = ([string][char]0x00E9) },
        @{ From = ([string][char]0x251C + [string][char]0x00AA); To = ([string][char]0x00EA) },
        @{ From = ([string][char]0x00E2 + [string][char]0x0080 + [string][char]0x0099); To = "'" },
        @{ From = ([string][char]0x00E2 + [string][char]0x0080 + [string][char]0x0093); To = "-" }
    )

    foreach ($replacement in $mojibakeReplacements) {
        $clean = $clean.Replace([string]$replacement.From, [string]$replacement.To)
    }

    $clean = $clean.Replace(([string][char]0x2018), "'")
    $clean = $clean.Replace(([string][char]0x2019), "'")
    $clean = $clean.Replace(([string][char]0x201C), '"')
    $clean = $clean.Replace(([string][char]0x201D), '"')
    $clean = $clean.Replace(([string][char]0x2013), "-")
    $clean = $clean.Replace(([string][char]0x2014), "-")

    return ([regex]::Replace($clean, "\s+", " ")).Trim()
}

function Set-RunWindowTitle {
    param([string]$Title)

    try {
        $Host.UI.RawUI.WindowTitle = $Title
    }
    catch {
    }
}

function Write-RunStatus {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host ("[{0}] [{1}] {2}" -f $timestamp, $Level, $Message)
}

function Write-CountProgress {
    param(
        [string]$Activity,
        [int]$Current,
        [int]$Total,
        [int]$Found = -1,
        [int]$Every = 10
    )

    if ($Total -le 0) {
        return
    }

    if ($Current -ne 1 -and $Current -ne $Total -and ($Current % $Every) -ne 0) {
        return
    }

    $percent = [int](($Current / [Math]::Max(1, $Total)) * 100)
    $foundText = ""
    if ($Found -ge 0) {
        $foundText = "; {0} matches so far" -f $Found
    }

    Write-RunStatus ("{0}: {1}/{2} ({3}%){4}" -f $Activity, $Current, $Total, $percent, $foundText)
}

function Start-SourceStats {
    param([string]$Source)

    return [ordered]@{
        Source          = $Source
        StartedAt       = [DateTimeOffset]::Now
        FinishedAt      = $null
        DurationSeconds = 0
        SearchRequests  = 0
        DetailRequests  = 0
        CacheHits       = 0
        Candidates      = 0
        SelectedDetails = 0
        SkippedOld      = 0
        SkippedContract = 0
        SkippedNoMatch  = 0
        SkippedByCap    = 0
        Errors          = 0
        Matches         = 0
        Notes           = ""
    }
}

function Add-SourceMetric {
    param(
        [AllowNull()]$Stats,
        [string]$Name,
        [int]$Amount = 1
    )

    if ($null -eq $Stats -or -not $Stats.Contains($Name)) {
        return
    }

    $Stats[$Name] = [int]$Stats[$Name] + $Amount
}

function Set-SourceNote {
    param(
        [AllowNull()]$Stats,
        [AllowNull()][string]$Note
    )

    if ($null -eq $Stats -or [string]::IsNullOrWhiteSpace($Note)) {
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$Stats["Notes"])) {
        $Stats["Notes"] = $Note
    }
    else {
        $Stats["Notes"] = "{0}; {1}" -f $Stats["Notes"], $Note
    }
}

function Complete-SourceStats {
    param([AllowNull()]$Stats)

    if ($null -eq $Stats) {
        return
    }

    $Stats["FinishedAt"] = [DateTimeOffset]::Now
    $Stats["DurationSeconds"] = [int][Math]::Round((([DateTimeOffset]$Stats["FinishedAt"]) - ([DateTimeOffset]$Stats["StartedAt"])).TotalSeconds, 0)
    $script:SourceRunStats.Add([PSCustomObject]$Stats) | Out-Null
    Write-RunStatus ("{0} diagnostics: {1}s, search {2}, details {3}, cache hits {4}, candidates {5}, selected {6}, old {7}, contract {8}, no-match {9}, cap {10}, errors {11}, matches {12}." -f `
            $Stats["Source"],
            $Stats["DurationSeconds"],
            $Stats["SearchRequests"],
            $Stats["DetailRequests"],
            $Stats["CacheHits"],
            $Stats["Candidates"],
            $Stats["SelectedDetails"],
            $Stats["SkippedOld"],
            $Stats["SkippedContract"],
            $Stats["SkippedNoMatch"],
            $Stats["SkippedByCap"],
            $Stats["Errors"],
            $Stats["Matches"])
}

function Get-SourceStatsSummaryText {
    if ($script:SourceRunStats.Count -eq 0) {
        return ""
    }

    $parts = foreach ($stat in @($script:SourceRunStats.ToArray())) {
        "{0}: {1}s, {2} match(es), {3} candidate(s), {4} detail(s), {5} cap-skip, {6} cache-hit(s)" -f `
            $stat.Source,
            $stat.DurationSeconds,
            $stat.Matches,
            $stat.Candidates,
            $stat.DetailRequests,
            $stat.SkippedByCap,
            $stat.CacheHits
    }

    return ($parts -join " | ")
}

function ConvertTo-QueryString {
    param([hashtable]$Params)

    ($Params.GetEnumerator() | ForEach-Object {
        "{0}={1}" -f [Uri]::EscapeDataString([string]$_.Key), [Uri]::EscapeDataString([string]$_.Value)
    }) -join "&"
}

function ConvertFrom-HtmlText {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $text = [regex]::Replace($Html, "(?is)<script\b.*?</script>|<style\b.*?</style>", " ")
    $text = [regex]::Replace($text, "(?is)<br\s*/?>|</p>|</li>|</div>|</h\d>", " ")
    $text = [regex]::Replace($text, "(?is)<[^>]+>", " ")
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return Repair-DisplayText $text
}

function ConvertFrom-HtmlAttribute {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return Repair-DisplayText ([System.Net.WebUtility]::HtmlDecode($Value))
}

function Get-HtmlAttributeValue {
    param(
        [AllowNull()][string]$Html,
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Html) -or [string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $pattern = "(?is)\b{0}\s*=\s*[""'](?<value>[^""']*)[""']" -f [regex]::Escape($Name)
    $match = [regex]::Match($Html, $pattern)
    if (-not $match.Success) {
        return ""
    }

    return ConvertFrom-HtmlAttribute $match.Groups["value"].Value
}

function ConvertTo-MatchText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $decoded = [System.Net.WebUtility]::HtmlDecode($Text)
    $normalized = $decoded.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object Text.StringBuilder
    foreach ($char in $normalized.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString().ToLowerInvariant()
}

function ConvertTo-CleanUrl {
    param([string]$Url)

    $clean = ConvertFrom-HtmlAttribute $Url
    $clean = $clean -replace "\?.*$", ""
    return $clean
}

function ConvertTo-SafeCacheKey {
    param([string]$Value)

    $bytes = [Text.Encoding]::UTF8.GetBytes([string]$Value)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    finally {
        $sha.Dispose()
    }
}

function Get-CacheFilePath {
    param(
        [string]$Scope,
        [string]$Key
    )

    $safeScope = ([regex]::Replace((ConvertTo-MatchText $Scope), "[^a-z0-9_-]+", "_")).Trim("_")
    if ([string]::IsNullOrWhiteSpace($safeScope)) {
        $safeScope = "default"
    }

    return Join-Path (Join-Path $CacheDirectory $safeScope) ("{0}.txt" -f (ConvertTo-SafeCacheKey $Key))
}

function Get-CachedText {
    param(
        [string]$Scope,
        [string]$Key,
        [int]$TtlHours = $CacheTtlHours
    )

    if ($DisableCache -or $TtlHours -le 0) {
        return $null
    }

    $path = Get-CacheFilePath -Scope $Scope -Key $Key
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    $item = Get-Item -LiteralPath $path
    if ($item.LastWriteTime -lt (Get-Date).AddHours(-[Math]::Abs($TtlHours))) {
        return $null
    }

    return [IO.File]::ReadAllText($item.FullName, [Text.Encoding]::UTF8)
}

function Set-CachedText {
    param(
        [string]$Scope,
        [string]$Key,
        [AllowNull()][string]$Text
    )

    if ($DisableCache -or [string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $path = Get-CacheFilePath -Scope $Scope -Key $Key
    $directory = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    [IO.File]::WriteAllText($path, [string]$Text, [Text.Encoding]::UTF8)
}

function ConvertTo-AbsoluteUrl {
    param(
        [string]$BaseUrl,
        [string]$Href
    )

    $cleanHref = ConvertFrom-HtmlAttribute $Href
    if ([string]::IsNullOrWhiteSpace($cleanHref)) {
        return ""
    }

    if ($cleanHref -match "^https?://") {
        return ConvertTo-CleanUrl $cleanHref
    }

    try {
        $baseUri = [Uri]::new($BaseUrl)
        return ConvertTo-CleanUrl ([Uri]::new($baseUri, $cleanHref).AbsoluteUri)
    }
    catch {
        return ConvertTo-CleanUrl $cleanHref
    }
}

function ConvertTo-DateTimeOffsetOrNull {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $parsed = [DateTimeOffset]::MinValue
    $styles = [Globalization.DateTimeStyles]::AllowWhiteSpaces
    if ([DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function ConvertFrom-FrenchRelativeDateText {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $clean = ConvertTo-MatchText (ConvertFrom-HtmlText $Text)
    $now = [DateTimeOffset]::Now
    if ($clean -match "aujourd.?hui|quelques\s+(secondes|minutes)|a\s+l.?instant") {
        return $now
    }
    if ($clean -match "\bhier\b") {
        return $now.AddDays(-1)
    }

    $hoursMatch = [regex]::Match($clean, "il\s+y\s+a\s+(?<value>\d+)\s+h")
    if ($hoursMatch.Success) {
        return $now.AddHours(-[int]$hoursMatch.Groups["value"].Value)
    }

    $dayMatch = [regex]::Match($clean, "il\s+y\s+a\s+(?<value>\d+)\s+j")
    if ($dayMatch.Success) {
        return $now.AddDays(-[int]$dayMatch.Groups["value"].Value)
    }

    $weekMatch = [regex]::Match($clean, "il\s+y\s+a\s+(?<value>\d+)\s+sem")
    if ($weekMatch.Success) {
        return $now.AddDays(-7 * [int]$weekMatch.Groups["value"].Value)
    }

    $dateMatch = [regex]::Match($clean, "(?<day>\d{1,2})[/-](?<month>\d{1,2})[/-](?<year>\d{4})")
    if ($dateMatch.Success) {
        $dateText = "{0}-{1}-{2}" -f $dateMatch.Groups["year"].Value, $dateMatch.Groups["month"].Value.PadLeft(2, "0"), $dateMatch.Groups["day"].Value.PadLeft(2, "0")
        return ConvertTo-DateTimeOffsetOrNull $dateText
    }

    return $null
}

function Test-IsRecent {
    param([AllowNull()]$PublishedAt)

    if ($null -eq $PublishedAt) {
        return $false
    }

    $publishedDateValue = $null
    if ($PublishedAt -is [DateTimeOffset]) {
        $publishedDateValue = $PublishedAt
    }
    elseif ($PublishedAt -is [DateTime]) {
        $publishedDateValue = [DateTimeOffset]$PublishedAt
    }
    else {
        $publishedDateValue = ConvertTo-DateTimeOffsetOrNull ([string]$PublishedAt)
    }

    return ($null -ne $publishedDateValue -and $publishedDateValue -ge $Cutoff)
}

function Add-MatchSignal {
    param(
        [hashtable]$State,
        [string]$Text,
        [string]$Pattern,
        [string]$Keyword,
        [int]$Score
    )

    if ($Text -match $Pattern) {
        $State.Score += $Score
        $State.Keywords[$Keyword] = $true
    }
}

function Get-JobMatch {
    param(
        [string]$Title,
        [string]$Text
    )

    $titleText = ConvertTo-MatchText $Title
    $fullText = ConvertTo-MatchText ("{0} {1}" -f $Title, $Text)
    $coreTitlePattern = "\bweb\s*analyst\b|\bdigital\s*analyst\b|analyste\s+(digital|web)|digital\s+analytics?\s+consultant|analytics?\s+consultant|web\s+analytics?\s+consultant|web\s+analytics|digital\s+analytics|tracking|webtracking|taggage|tagging|\bdata\s*analyst\b|analyste\s+de\s+donnees|performance\s+digital|performance\s+digitale|digital\s+performance|\bcro\b|conversion\s+rate|conversion\s+optimization|optimisation\s+conversion"
    $hasCoreTitleSignal = $titleText -match $coreTitlePattern
    $state = @{
        Score = 0
        Keywords = @{}
    }
    $isGoToMarketContext = $fullText -match "go\s*[- ]?\s*to\s*[- ]?\s*market|\banalytics\s+engineer\s+gtm\b|\bgtm\s+(strategy|strategies|motion|motions|operations|ops|sales|revenue|revops|demand|pipeline|finance|engineer|generation|outbound|inbound)\b|\b(growth|demand|pipeline|revops|revenue|sales|finance|outbound|inbound)\s+gtm\b|\b(head|director|manager|lead|chief)\s+of\s+gtm\b|\b(sdr|bdr)\b.*\bgtm\b"
    $webAnalyticsToolPattern = "google\s+tag\s+manager|google\s+analytics|\bga4\b|piano\s+analytics|contentsquare|content\s+square|tag\s+commander|commanders?\s+act|\btealium\b|tealium\s+iq|matomo|adobe\s+analytics|omniture|data\s*layer|datalayer|tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)|consent\s+mode|cookie\s+consent|server\s*[- ]?\s*side\s+(tracking|tagging|analytics)|tracking\s+server\s*[- ]?\s*side|tagging\s+server\s*[- ]?\s*side|server\s+container|\bsgtm\b"
    $hasWebAnalyticsToolSignal = ($fullText -match $webAnalyticsToolPattern) -or (-not $isGoToMarketContext -and $fullText -match "\bgtm\b")
    $hasDigitalAnalyticsContext = $hasWebAnalyticsToolSignal -or ($titleText -match "\bweb\s*analyst\b|\bdigital\s*analyst\b|analyste\s+(digital|web)|web\s+analytics|digital\s+analytics|tracking|webtracking|taggage|tagging|analytics?\s+consultant|performance\s+digital|performance\s+digitale|digital\s+performance|\bcro\b|conversion")
    $isMarketingOnlyContext = ($fullText -match "\bseo\b|\bsea\b|paid\s+social|paid\s+search|paid\s+media|performance\s+marketing|growth\s+marketing|acquisition\s+marketing|digital\s+marketing|social\s+media|content\s+marketing|campaign\s+manager|media\s+buyer") -and -not $hasDigitalAnalyticsContext
    $isDataWarehouseContext = $fullText -match "\bdbt\b|snowflake|airflow|\betl\b|\belt\b|data\s+warehouse|datawarehouse|data\s+platform|databricks|pyspark|spark|data\s+pipeline|\bpython\b"

    Add-MatchSignal $state $titleText "\bweb\s*analyst\b" "title:web analyst" 60
    Add-MatchSignal $state $titleText "\bdigital\s*analyst\b|analyste\s+digital|analyste\s+web" "title:digital/web analyst" 55
    Add-MatchSignal $state $titleText "digital\s+analytics?\s+consultant|analytics?\s+consultant|web\s+analytics?\s+consultant" "title:analytics consultant" 55
    Add-MatchSignal $state $titleText "tracking|webtracking|taggage|tagging" "title:tracking/tagging" 50
    Add-MatchSignal $state $titleText "web\s+analytics|digital\s+analytics" "title:web/digital analytics" 45
    Add-MatchSignal $state $titleText "\bdata\s*analyst\b|analyste\s+de\s+donnees" "title:data analyst" 25
    Add-MatchSignal $state $titleText "performance\s+digital|performance\s+digitale|digital\s+performance" "title:digital performance" 30
    Add-MatchSignal $state $titleText "\bcro\b|conversion\s+rate|conversion\s+optimization|optimisation\s+conversion" "title:CRO" 25

    Add-MatchSignal $state $fullText "google\s+tag\s+manager" "Google Tag Manager" 35
    if (-not $isGoToMarketContext -and $fullText -match "\bgtm\b") {
        $state.Score += 35
        $state.Keywords["GTM"] = $true
    }
    Add-MatchSignal $state $fullText "google\s+analytics|\bga4\b" "Google Analytics/GA4" 35
    Add-MatchSignal $state $fullText "piano\s+analytics" "Piano Analytics" 35
    Add-MatchSignal $state $fullText "contentsquare|content\s+square" "ContentSquare" 35
    Add-MatchSignal $state $fullText "tag\s+commander|commanders?\s+act" "Tag Commander/Commanders Act" 35
    Add-MatchSignal $state $fullText "\btealium\b|tealium\s+iq" "Tealium" 35
    Add-MatchSignal $state $fullText "data\s*layer|datalayer" "dataLayer" 25
    Add-MatchSignal $state $fullText "plan\s+de\s+(taggage|marquage)|tagging\s+plan|tracking\s+plan" "tagging plan" 30
    Add-MatchSignal $state $fullText "tracking|taggage|tagging|tag\s+management" "tracking/tagging" 20
    Add-MatchSignal $state $fullText "server\s*[- ]?\s*side\s+(tracking|tagging|analytics)|tracking\s+server\s*[- ]?\s*side|tagging\s+server\s*[- ]?\s*side|server\s+container|\bsgtm\b" "server-side tracking" 20
    Add-MatchSignal $state $fullText "consent\s+mode|cmp|cookie\s+consent" "consent tracking" 20
    Add-MatchSignal $state $fullText "\brgpd\b|\bgdpr\b|protection\s+des\s+donnees|protection\s+des\s+donn[eé]es|privacy|conformit[eé]" "RGPD/GDPR" 10
    Add-MatchSignal $state $fullText "a/b\s*test|ab\s*test|experimentation" "A/B testing" 15
    Add-MatchSignal $state $fullText "matomo|adobe\s+analytics|omniture" "other analytics tools" 20
    Add-MatchSignal $state $fullText "looker\s+studio|data\s+studio|dashboard|reporting|tableau\s+de\s+bord" "reporting/dashboard" 10
    Add-MatchSignal $state $fullText "\bkpi\b|conversion|funnel|parcours\s+utilisateur|user\s+journey" "analytics responsibilities" 10

    if ($isMarketingOnlyContext) {
        $state.Score -= 25
        $state.Keywords["possible SEO/SEA/marketing-only role"] = $true
    }
    elseif ($fullText -match "\bseo\b|\bsea\b|paid\s+social|performance\s+marketing|growth\s+marketing|digital\s+marketing") {
        $state.Score -= 8
        $state.Keywords["possible marketing role"] = $true
    }
    if ($titleText -match "time\s+tracking|absence|payroll" -and $titleText -notmatch "web|analytics|tagging|taggage|webtracking|\bcro\b") {
        $state.Score -= 60
        $state.Keywords["possible time-tracking/payroll role"] = $true
    }
    if ($titleText -match "\bseo\s+specialist\b" -and $titleText -notmatch "analytics|web\s+analytics|tracking|tagging|taggage") {
        $state.Score -= 60
        $state.Keywords["possible SEO-only role"] = $true
    }
    if ($titleText -match "people\s+business\s+partner|\bhr\b|human\s+resources|ressources\s+humaines|talent\s+acquisition|recruiter|recruteur|recrutement|charg[eé]\s+de\s+recrutement" -and -not $hasCoreTitleSignal) {
        $state.Score -= 60
        $state.Keywords["possible HR/recruiting role"] = $true
    }
    if ($isGoToMarketContext) {
        $state.Score -= 35
        $state.Keywords["possible go-to-market role"] = $true
    }
    if ($fullText -match "business\s+analyst|risk|risque|finance|banking|bancaire") {
        $state.Score -= 10
        $state.Keywords["possible broad analyst role"] = $true
    }
    if ((($titleText -match "\bdata\s*analyst\b|analyste\s+de\s+donnees|analytics?\s+engineer|data\s+engineer") -or $isDataWarehouseContext) -and -not $hasDigitalAnalyticsContext) {
        $state.Score -= 25
        $state.Keywords["possible data analyst/engineering role"] = $true
    }
    elseif ($isDataWarehouseContext) {
        $state.Score -= 12
        $state.Keywords["possible warehouse/python role"] = $true
    }
    if ($fullText -match "software\s+engineer|data\s+engineer|sales\s+engineer|backend|frontend|devops") {
        $state.Score -= 15
        $state.Keywords["possible engineering role"] = $true
    }

    $learning = Get-FeedbackLearningAdjustment `
        -FullText $fullText `
        -HasCoreTitleSignal:$hasCoreTitleSignal `
        -HasWebAnalyticsToolSignal:$hasWebAnalyticsToolSignal `
        -HasDigitalAnalyticsContext:$hasDigitalAnalyticsContext
    if ($null -ne $learning -and [int]$learning.Adjustment -ne 0) {
        $state.Score += [int]$learning.Adjustment
        foreach ($reason in @($learning.Reasons)) {
            if (-not [string]::IsNullOrWhiteSpace($reason)) {
                $state.Keywords[$reason] = $true
            }
        }
    }

    if ($state.Score -lt 0) {
        $state.Score = 0
    }
    if (-not $hasCoreTitleSignal -and $state.Score -gt 49) {
        $state.Score = 49
        $state.Keywords["no core title signal"] = $true
    }

    $level = "Review"
    if ($state.Score -ge 80) {
        $level = "High"
    }
    elseif ($state.Score -ge 50) {
        $level = "Medium"
    }

    [PSCustomObject]@{
        IsMatch = $state.Score -ge $MinimumMatchScore
        Score = [int]$state.Score
        Level = $level
        Keywords = (($state.Keywords.Keys | Sort-Object) -join "; ")
    }
}

function Get-ContractTypeFromText {
    param(
        [AllowNull()][string]$Text,
        [AllowNull()][string]$RawContractType = $null
    )

    if (-not [string]::IsNullOrWhiteSpace($RawContractType)) {
        switch -Regex ($RawContractType) {
            "^FULL_TIME$" { return "CDI" }
            "^INTERNSHIP$" { return "Internship" }
            "^APPRENTICESHIP$" { return "Apprenticeship" }
            "^TEMPORARY$" { return "CDD" }
            "^FREELANCE$" { return "Freelance" }
            default { return $RawContractType }
        }
    }

    $matchText = ConvertTo-MatchText $Text
    if ($matchText -match "alternance|alternant|apprentissage|apprenticeship") {
        return "Apprenticeship"
    }
    if ($matchText -match "\bstage\b|stagiaire|internship|intern\b") {
        return "Internship"
    }
    if ($matchText -match "\bcdi\b|contrat\s+a\s+duree\s+indeterminee") {
        return "CDI"
    }
    if ($matchText -match "\bcdd\b|contrat\s+a\s+duree\s+determinee|fixed[-\s]*term|temporary|temporaire") {
        return "CDD"
    }
    if ($matchText -match "freelance|contractor|independant") {
        return "Freelance"
    }
    if ($matchText -match "employment\s+type\s+full-time|type\s+d.?emploi\s+temps\s+plein|\bfull-time\b|\btemps\s+plein\b") {
        return "Full-time"
    }
    if ($matchText -match "permanent\s+(contract|position|role)|\bpermanent\b") {
        return "Permanent"
    }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    if ($Text -match "(?i)(\bCDI\b|contrat\s+a\s+duree\s+indeterminee)") {
        return "CDI"
    }
    if ($Text -match "(?i)(alternance|alternant|apprentissage|apprenticeship)") {
        return "Apprenticeship"
    }
    if ($Text -match "(?i)(stage|stagiaire|internship|intern\b)") {
        return "Internship"
    }
    if ($Text -match "(?i)(\bCDD\b|contrat\s+a\s+duree\s+determinee|fixed[-\s]*term|temporary|temporaire)") {
        return "CDD"
    }
    if ($Text -match "(?i)(freelance|contractor|independant)") {
        return "Freelance"
    }
    if ($Text -match "(?i)(Employment\s+type\s+Full-time|Type\s+d.?emploi\s+Temps\s+plein|\bFull-time\b|\bTemps\s+plein\b)") {
        return "Full-time"
    }
    if ($Text -match "(?i)(permanent\s+(contract|position|role)|\bpermanent\b)") {
        return "Permanent"
    }

    return ""
}

function Get-FranceTravailAccessToken {
    if ([string]::IsNullOrWhiteSpace($FranceTravailClientId) -or [string]::IsNullOrWhiteSpace($FranceTravailClientSecret)) {
        Write-RunStatus "France Travail credentials not set; skipping France Travail source. Set FRANCE_TRAVAIL_CLIENT_ID and FRANCE_TRAVAIL_CLIENT_SECRET to enable it."
        return ""
    }

    $tokenUrl = $env:FRANCE_TRAVAIL_TOKEN_URL
    if ([string]::IsNullOrWhiteSpace($tokenUrl)) {
        $tokenUrl = "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=/partenaire"
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
        return "https://candidat.francetravail.fr/offres/recherche/detail/{0}" -f ([Uri]::EscapeDataString([string]$jobId))
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

function Get-AdzunaContractType {
    param([AllowNull()]$Job)

    $contractType = ConvertTo-MatchText (Get-ObjectPropertyValue -Object $Job -Names @("contract_type"))
    $contractTime = ConvertTo-MatchText (Get-ObjectPropertyValue -Object $Job -Names @("contract_time"))

    if ($contractType -match "permanent") {
        return "Permanent"
    }
    if ($contractType -match "contract|freelance") {
        return "Freelance"
    }
    if ($contractTime -match "full\s*time|full_time") {
        return "Full-time"
    }
    if ($contractTime -match "part\s*time|part_time") {
        return "Part-time"
    }

    return ""
}

function Get-AdzunaLocation {
    param([AllowNull()]$Job)

    $location = Get-ObjectPropertyValue -Object $Job -Names @("location")
    $displayName = Get-ObjectPropertyValue -Object $location -Names @("display_name")
    if (-not [string]::IsNullOrWhiteSpace($displayName)) {
        return [string]$displayName
    }

    $area = Get-ObjectPropertyValue -Object $location -Names @("area")
    return ConvertTo-LocationText $area
}

function Get-AdzunaCompanyName {
    param([AllowNull()]$Job)

    $company = Get-ObjectPropertyValue -Object $Job -Names @("company")
    $companyName = Get-ObjectPropertyValue -Object $company -Names @("display_name", "name")
    if (-not [string]::IsNullOrWhiteSpace($companyName)) {
        return [string]$companyName
    }

    return ""
}

function Get-LinkedInContractType {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$DetailText
    )

    $titleContract = Get-ContractTypeFromText -Text $Title
    if (-not [string]::IsNullOrWhiteSpace($titleContract)) {
        return $titleContract
    }

    $detailMatchText = ConvertTo-MatchText $DetailText
    if ($detailMatchText -match "employment\s+type\s+full-time|type\s+d.?emploi\s+temps\s+plein") {
        return "Full-time"
    }
    if ($detailMatchText -match "employment\s+type\s+internship|type\s+d.?emploi\s+stage") {
        return "Internship"
    }
    if ($detailMatchText -match "employment\s+type\s+temporary|employment\s+type\s+contract") {
        return "CDD"
    }

    return Get-ContractTypeFromText -Text $DetailText
}

function ConvertTo-ExcelHyperlinkFormula {
    param(
        [AllowNull()][string]$Url,
        [string]$Label = "Open"
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ""
    }

    $cleanUrl = $Url.Trim()
    if ($cleanUrl -notmatch "^https?://") {
        return $cleanUrl
    }

    $escapedUrl = $cleanUrl.Replace('"', '""')
    $escapedLabel = $Label.Replace('"', '""')
    return '=HYPERLINK("{0}","{1}")' -f $escapedUrl, $escapedLabel
}

function ConvertTo-IdentityText {
    param(
        [AllowNull()][string]$Text,
        [switch]$Title
    )

    $clean = ConvertTo-MatchText $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ""
    }

    if ($Title) {
        $clean = $clean -replace "\b(cdi|cdd|stage|internship|intern|alternance|apprentissage|apprenticeship|contrat|full[- ]?time|permanent|h|f|hf|h/f|f/h|m|x|nb)\b", " "
        $clean = $clean -replace "\b(senior|junior|jr|sr|lead|manager|confirme|confirmee|experimente|experimentee)\b", " "
        $clean = $clean -replace "\b(paris|lyon|lille|bordeaux|nantes|rennes|montpellier|marseille|toulouse|puteaux|levallois|boulogne|casablanca|france)\b", " "
    }
    else {
        $clean = $clean -replace "\b(france|fr|group|groupe|sa|sas|sasu|ltd|limited|inc|plc)\b", " "
    }

    $clean = $clean -replace "[^a-z0-9]+", " "
    return ([regex]::Replace($clean, "\s+", " ")).Trim()
}

function Split-NormalizedTokens {
    param([AllowNull()][string]$Text)

    $clean = ConvertTo-IdentityText -Text $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return @()
    }

    return @($clean -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-IsGenericJobBoardName {
    param([AllowNull()][string]$Name)

    $text = ConvertTo-MatchText $Name
    $text = $text -replace "[^a-z0-9]+", " "
    $text = ([regex]::Replace($text, "\s+", " ")).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    return $text -match "\b(france travail|pole emploi|poles emploi|adzuna|linkedin|indeed|hellowork|hello work|meteojob|jobijoba|monster|apec|jobgether|talent com|confidential|confidentiel|licorne|recrutement)\b"
}

function Get-DedupeCompanyKey {
    param([AllowNull()][string]$CompanyName)

    if (Test-IsGenericJobBoardName $CompanyName) {
        return ""
    }

    $tokens = @(Split-NormalizedTokens $CompanyName)
    if ($tokens.Count -eq 0) {
        return ""
    }

    $noise = @(
        "the", "and", "et", "de", "du", "des", "la", "le", "les",
        "sa", "sas", "sasu", "ltd", "limited", "inc", "plc", "fr", "france",
        "group", "groupe", "company", "companies", "media", "digital",
        "consulting", "consultants", "technology", "technologies", "solutions"
    )
    $weakCompanyTokens = @("confidential", "confidentiel", "jobgether", "licorne", "recrutement", "talent", "emploi", "travail", "adzuna", "linkedin", "indeed", "hellowork", "meteojob", "jobijoba", "monster", "apec")
    $strongTokens = @($tokens | Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ -and $weakCompanyTokens -notcontains $_ })
    if ($strongTokens.Count -eq 0) {
        return ""
    }

    if ($strongTokens.Count -eq 1) {
        return $strongTokens[0]
    }

    return ($strongTokens | Select-Object -First 2) -join " "
}

function ConvertTo-DedupeTitleToken {
    param([AllowNull()][string]$Token)

    if ([string]::IsNullOrWhiteSpace($Token)) {
        return ""
    }

    switch -Regex ($Token) {
        "^(analyste|analystes)$" { return "analyst" }
        "^(consultante|consultants|consultantes)$" { return "consultant" }
        "^(digitale|digitaux|digitales)$" { return "digital" }
        "^(analytics|analytic|analytique|analytiques)$" { return "analytics" }
        "^(chargee|charges|chargees)$" { return "charge" }
        "^(performances)$" { return "performance" }
        default { return $Token }
    }
}

function Get-DedupeLocationKey {
    param([AllowNull()][string]$Location)

    $text = ConvertTo-MatchText $Location
    $text = $text -replace "[^a-z0-9]+", " "
    $text = ([regex]::Replace($text, "\s+", " ")).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "unknown"
    }

    if ($text -match "\b(ile de france|greater paris|ville de paris|paris|puteaux|courbevoie|nanterre|levallois|boulogne|clichy|issy|issy les moulineaux|neuilly|la defense|saint denis|montreuil|gentilly|ivry|bagnolet|suresnes|rueil|meudon|chatillon|montrouge|vincennes)\b") {
        return "paris_metro"
    }

    if ($text -match "\b(remote|teletravail|france|fr)\b" -and $text -notmatch "\b(lyon|lille|bordeaux|nantes|rennes|montpellier|marseille|toulouse|nice|strasbourg|grenoble|dijon|angers|annecy|niort|caen|aix|limoges|poissy|armentieres|champagne|haute savoie|carros|casablanca)\b") {
        return "france"
    }

    $text = $text -replace "\b(et peripherie|metropolitan area|area|region|france|remote|hybrid|on site|sur site)\b", " "
    $tokens = @($text -split "\s+" | Where-Object { $_.Length -gt 1 })
    if ($tokens.Count -eq 0) {
        return "unknown"
    }

    return ($tokens | Select-Object -First 2) -join " "
}

function Get-DedupeTitleKey {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName
    )

    $text = ConvertTo-MatchText $Title
    $text = $text -replace "[^a-z0-9]+", " "
    $text = ([regex]::Replace($text, "\s+", " ")).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    $companyTokens = @(Split-NormalizedTokens $CompanyName | Where-Object { $_.Length -gt 2 })
    foreach ($token in $companyTokens) {
        $text = $text -replace ("\b{0}\b" -f [regex]::Escape($token)), " "
    }

    $noise = @(
        "cdi", "cdd", "stage", "stagiaire", "internship", "intern", "alternance", "apprentissage",
        "contrat", "full", "time", "permanent", "temps", "plein", "h", "f", "hf", "fh", "fm", "mf", "x", "nb",
        "paris", "lyon", "lille", "bordeaux", "nantes", "rennes", "montpellier", "marseille", "toulouse",
        "puteaux", "courbevoie", "nanterre", "levallois", "boulogne", "clichy", "issy", "france",
        "ile", "de", "a", "au", "aux", "en", "la", "le", "les", "du", "des", "et", "emea",
        "media", "sa", "sas", "sasu", "groupe", "group"
    )
    $tokens = @(
        $text -split "\s+" |
            Where-Object { $_.Length -gt 1 -and $noise -notcontains $_ } |
            ForEach-Object { ConvertTo-DedupeTitleToken $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
    if ($tokens.Count -eq 0) {
        return ""
    }

    return ($tokens -join " ")
}

function Test-IsStrongDedupeTitle {
    param([AllowNull()][string]$TitleKey)

    if ([string]::IsNullOrWhiteSpace($TitleKey)) {
        return $false
    }

    $tokens = @($TitleKey -split "\s+" | Where-Object { $_.Length -gt 1 })
    if ($tokens.Count -lt 2) {
        return $false
    }

    return $TitleKey -match "\b(analyst|analyste|analytics|web|digital|tracking|tagging|taggage|data|traffic|performance|cro|conversion|consultant|manager|lead|product|gtm|seo|sea)\b"
}

function Test-UseLocationInDedupeKey {
    param([AllowNull()][string]$TitleKey)

    $tokens = @(([string]$TitleKey) -split "\s+" | Where-Object { $_.Length -gt 1 })
    if ($tokens.Count -ge 3) {
        return $false
    }

    $hasCoreAnalyticsTitle =
        (($TitleKey -match "\bweb\b") -and ($TitleKey -match "\banalyst\b|\banalytics\b")) -or
        (($TitleKey -match "\bdigital\b") -and ($TitleKey -match "\banalyst\b|\banalytics\b")) -or
        (($TitleKey -match "\bdata\b") -and ($TitleKey -match "\banalyst\b") -and ($TitleKey -match "\bweb\b|\bdigital\b|\banalytics\b")) -or
        ($TitleKey -match "\b(tracking|tagging|taggage|gtm|cro)\b")

    if ($hasCoreAnalyticsTitle) {
        return $false
    }

    return $true
}

function Get-JobDedupeKeyFromValues {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$Url
    )

    $companyKey = Get-DedupeCompanyKey $CompanyName
    $titleKey = Get-DedupeTitleKey -Title $Title -CompanyName $CompanyName
    $locationKey = Get-DedupeLocationKey $JobLocation

    if (-not [string]::IsNullOrWhiteSpace($companyKey) -and (Test-IsStrongDedupeTitle $titleKey)) {
        if (-not (Test-UseLocationInDedupeKey $titleKey)) {
            return "company-title|{0}|{1}" -f $companyKey, $titleKey
        }

        return "company-title-location|{0}|{1}|{2}" -f $companyKey, $titleKey, $locationKey
    }

    $urlKey = ""
    if (-not [string]::IsNullOrWhiteSpace($Url)) {
        $urlKey = ($Url.Trim().ToLowerInvariant() -replace "\?.*$", "")
    }
    return "url|{0}" -f $urlKey
}

function Get-StableJobId {
    param([string]$IdentityKey)

    if ([string]::IsNullOrWhiteSpace($IdentityKey)) {
        $IdentityKey = [Guid]::NewGuid().ToString("N")
    }

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($IdentityKey)
        $hash = $sha.ComputeHash($bytes)
        return (($hash | ForEach-Object { $_.ToString("x2") }) -join "").Substring(0, 16)
    }
    finally {
        $sha.Dispose()
    }
}

function Get-JobIdentityKeyFromValues {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$Url
    )

    return Get-JobDedupeKeyFromValues -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -Url $Url
}

function Get-RowValue {
    param(
        [AllowNull()]$Row,
        [string]$Name
    )

    if ($null -eq $Row) {
        return ""
    }

    if (@($Row.PSObject.Properties.Name) -contains $Name) {
        $value = $Row.PSObject.Properties[$Name].Value
        if ($null -ne $value) {
            return [string]$value
        }
    }

    return ""
}

function Get-JobIdentityKeyFromRow {
    param([AllowNull()]$Row)

    return Get-JobIdentityKeyFromValues `
        -Title (Get-RowValue -Row $Row -Name "job_title") `
        -CompanyName (Get-RowValue -Row $Row -Name "company_name") `
        -JobLocation (Get-RowValue -Row $Row -Name "location") `
        -Url (Get-RowValue -Row $Row -Name "job_url_raw")
}

function Get-JobDedupeKeyFromRow {
    param([AllowNull()]$Row)

    return Get-JobDedupeKeyFromValues `
        -Title (Get-RowValue -Row $Row -Name "job_title") `
        -CompanyName (Get-RowValue -Row $Row -Name "company_name") `
        -JobLocation (Get-RowValue -Row $Row -Name "location") `
        -Url (Get-RowValue -Row $Row -Name "job_url_raw")
}

function Test-IsExcludedContractType {
    param([AllowNull()][string]$ContractType)

    $contractText = ConvertTo-MatchText $ContractType
    if ([string]::IsNullOrWhiteSpace($contractText)) {
        return $false
    }

    return $contractText -match "\bcdd\b|apprenticeship|apprentissage|alternance|internship|\bstage\b|stagiaire|temporary|fixed\s+term|freelance|contractor|independant|independent"
}

function Get-EarlyContractType {
    param(
        [AllowNull()][string]$ContractType,
        [AllowNull()][string]$Text = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($ContractType)) {
        return Get-ContractTypeFromText -Text $Text -RawContractType $ContractType
    }

    return Get-ContractTypeFromText -Text $Text
}

function Test-ShouldSkipEarlyByContract {
    param(
        [AllowNull()][string]$ContractType,
        [AllowNull()][string]$Text = "",
        [switch]$Reliable
    )

    $effectiveContract = Get-EarlyContractType -ContractType $ContractType -Text $Text
    if ([string]::IsNullOrWhiteSpace($effectiveContract)) {
        return $false
    }

    if ($Reliable) {
        return Test-IsExcludedContractType $effectiveContract
    }

    $matchText = ConvertTo-MatchText (Join-CleanTextParts @($ContractType, $Text))
    if ($matchText -match "\bcdd\b|apprentissage|alternance|apprenticeship|internship|\bstage\b|stagiaire|freelance|contractor|independant|independent") {
        return Test-IsExcludedContractType $effectiveContract
    }

    return $false
}

function Get-PreferenceObjectValue {
    param(
        [AllowNull()]$Object,
        [string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    if ($Object -is [Collections.IDictionary] -and $Object.Contains($Name)) {
        return $Object[$Name]
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        return $property.Value
    }

    return $DefaultValue
}

function New-DefaultJobCrawlerPreferences {
    return [PSCustomObject]@{
        preferred_employer_type = "annonceur"
        employer_type_weights = [PSCustomObject]@{
            annonceur = 10
            agency = -8
            consulting = -8
            esn = -10
            unknown = 0
        }
        location_fit_weights = [PSCustomObject]@{
            target = 8
            france_other = -4
            foreign = -20
            unknown = 0
        }
        seniority_fit_weights = [PSCustomObject]@{
            target = 0
            senior_ok = 0
            too_junior = -12
            too_managerial = -12
            unknown = 0
        }
        contract_fit_weights = [PSCustomObject]@{
            preferred = 5
            excluded = -100
            unknown = 0
        }
        target_location_patterns = @(
            "paris",
            "ile\s*de\s*france",
            "la\s+defense",
            "puteaux",
            "boulogne",
            "courbevoie",
            "nanterre",
            "clichy",
            "remote",
            "teletravail",
            "france"
        )
        foreign_location_patterns = @(
            "london",
            "madrid",
            "barcelona",
            "casablanca",
            "montreal",
            "brussels",
            "belgium",
            "luxembourg",
            "switzerland",
            "cyprus",
            "canada",
            "morocco",
            "spain",
            "united\s+kingdom"
        )
    }
}

function Get-JobCrawlerPreferences {
    $default = New-DefaultJobCrawlerPreferences
    $path = Join-Path $PSScriptRoot "config\preferences.json"
    if (-not (Test-Path -LiteralPath $path)) {
        return $default
    }

    try {
        return (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)
    }
    catch {
        Write-Warning ("Could not read config\preferences.json, using built-in defaults: {0}" -f $_.Exception.Message)
        return $default
    }
}

function Get-PreferenceWeight {
    param(
        [AllowNull()]$Preferences,
        [string]$GroupName,
        [string]$Key,
        [int]$DefaultValue = 0
    )

    $group = Get-PreferenceObjectValue -Object $Preferences -Name $GroupName -DefaultValue $null
    $rawValue = Get-PreferenceObjectValue -Object $group -Name $Key -DefaultValue $DefaultValue
    $number = 0
    if ([int]::TryParse([string]$rawValue, [ref]$number)) {
        return $number
    }

    return $DefaultValue
}

function Get-PreferenceArray {
    param(
        [AllowNull()]$Preferences,
        [string]$Name,
        [string[]]$DefaultValue = @()
    )

    $rawValue = Get-PreferenceObjectValue -Object $Preferences -Name $Name -DefaultValue $DefaultValue
    if ($null -eq $rawValue) {
        return @()
    }
    if ($rawValue -is [string]) {
        return @([string]$rawValue)
    }

    return @($rawValue)
}

function Test-AnyPatternMatch {
    param(
        [string]$Text,
        [object[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    foreach ($pattern in @($Patterns)) {
        if ([string]::IsNullOrWhiteSpace([string]$pattern)) {
            continue
        }

        try {
            if ($Text -match [string]$pattern) {
                return $true
            }
        }
        catch {
        }
    }

    return $false
}

function Get-EmployerType {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$Text = ""
    )

    $titleText = ConvertTo-MatchText $Title
    $companyText = ConvertTo-MatchText $CompanyName
    $combinedText = ConvertTo-MatchText (Join-CleanTextParts @($Title, $CompanyName, $Text))

    if ([string]::IsNullOrWhiteSpace($companyText) -or $companyText -match "confidential|jobgether|linkedin|indeed|talent\s*com|jobs?\s+via") {
        return "unknown"
    }

    $knownAgencyPattern = "publicis|dentsu|havas|labelium|pixalione|eskimoz|jellyfish|performics|iprospect|allmatik|leonar|ekinox"
    $knownConsultingPattern = "fifty\s*[- ]?\s*five|\b55\b|converteo|artefact|optimal\s+ways|innoha|ekimetrics|wavestone|mc2i|deloitte|pwc|ey|kpmg|accenture"
    $knownEsnPattern = "\bcgi\b|infotel|oventi|keyrus|micropole|business\s+&?\s+decision|devoteam|onepoint|talan|sopra\s+steria|capgemini|\bsqli\b|\bsqly\b|niji|consort|nexton|scalian|amaris|\bsii\b|atos|worldline|inetum|alten|ausy|neosoft"
    $agencyContextPattern = "\bagence\b|agency|paid\s+media\s+agency|marketing\s+agency|media\s+agency"
    $consultingContextPattern = "cabinet\s+(de\s+)?conseil|societe\s+de\s+conseil|\bconseil\b|consulting\s+(firm|agency|company|cabinet)|missions?\s+chez\s+les?\s+clients|chez\s+nos\s+clients"
    $esnContextPattern = "\besn\b|\bssii\b|services\s+numeriques|entreprise\s+de\s+services\s+du\s+numerique"

    if ($companyText -match $knownEsnPattern -or $combinedText -match $esnContextPattern) {
        return "esn"
    }
    if ($companyText -match $knownAgencyPattern -or $combinedText -match $agencyContextPattern) {
        return "agency"
    }
    if ($companyText -match $knownConsultingPattern -or $combinedText -match $consultingContextPattern) {
        return "consulting"
    }
    if ($titleText -match "\bconsultant(e)?\b" -and $combinedText -match "client|mission|conseil|consulting|cabinet") {
        return "consulting"
    }

    return "annonceur"
}

function Get-LocationFitCategory {
    param(
        [AllowNull()][string]$Location,
        [AllowNull()]$Preferences
    )

    $locationText = ConvertTo-MatchText $Location
    if ([string]::IsNullOrWhiteSpace($locationText)) {
        return "unknown"
    }

    $targetPatterns = Get-PreferenceArray -Preferences $Preferences -Name "target_location_patterns" -DefaultValue @("paris", "ile\s*de\s*france", "france", "remote", "teletravail")
    $foreignPatterns = Get-PreferenceArray -Preferences $Preferences -Name "foreign_location_patterns" -DefaultValue @("london", "madrid", "casablanca", "montreal", "belgium", "spain", "canada", "morocco")

    if (Test-AnyPatternMatch -Text $locationText -Patterns $targetPatterns) {
        return "target"
    }
    if (Test-AnyPatternMatch -Text $locationText -Patterns $foreignPatterns) {
        return "foreign"
    }
    if ($locationText -match "\bfrance\b|paris|ile\s*de\s*france") {
        return "target"
    }

    return "france_other"
}

function Get-SeniorityFitCategory {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$Text
    )

    $titleText = ConvertTo-MatchText $Title
    $fullText = ConvertTo-MatchText (Join-CleanTextParts @($Title, $Text))

    if ($titleText -match "\b(stage|stagiaire|intern|internship|apprentice|apprentissage|alternance|assistant|graduate|junior)\b") {
        return "too_junior"
    }
    if ($titleText -match "\b(head|director|directeur|directrice|lead|manager|responsable|principal)\b") {
        return "too_managerial"
    }
    if ($fullText -match "\b(stage|stagiaire|internship|apprentissage|alternance)\b") {
        return "too_junior"
    }
    if ($titleText -match "\b(senior|sr)\b") {
        return "senior_ok"
    }

    return "target"
}

function Get-ContractFitCategory {
    param([AllowNull()][string]$ContractType)

    $contractText = ConvertTo-MatchText $ContractType
    if ([string]::IsNullOrWhiteSpace($contractText)) {
        return "unknown"
    }
    if (Test-IsExcludedContractType $ContractType) {
        return "excluded"
    }
    if ($contractText -match "\bcdi\b|permanent|full\s*time|temps\s+plein") {
        return "preferred"
    }

    return "unknown"
}

function Get-JobFitDimensions {
    param(
        [int]$RoleScore,
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$JobLocation,
        [AllowNull()][string]$ContractType,
        [AllowNull()][string]$Text = "",
        [AllowNull()]$Preferences = $JobCrawlerPreferences
    )

    if ($null -eq $Preferences) {
        $Preferences = New-DefaultJobCrawlerPreferences
    }

    $employerType = Get-EmployerType -Title $Title -CompanyName $CompanyName -Text $Text
    $locationCategory = Get-LocationFitCategory -Location $JobLocation -Preferences $Preferences
    $seniorityCategory = Get-SeniorityFitCategory -Title $Title -Text $Text
    $contractCategory = Get-ContractFitCategory $ContractType

    $employerFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "employer_type_weights" -Key $employerType -DefaultValue 0
    $locationFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "location_fit_weights" -Key $locationCategory -DefaultValue 0
    $seniorityFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "seniority_fit_weights" -Key $seniorityCategory -DefaultValue 0
    $contractFit = Get-PreferenceWeight -Preferences $Preferences -GroupName "contract_fit_weights" -Key $contractCategory -DefaultValue 0

    $notes = New-Object System.Collections.Generic.List[string]
    $notes.Add(("role score {0}" -f $RoleScore)) | Out-Null
    if ($employerFit -ne 0) { $notes.Add(("employer {0}: {1}" -f $employerType, $employerFit)) | Out-Null }
    if ($locationFit -ne 0) { $notes.Add(("location {0}: {1}" -f $locationCategory, $locationFit)) | Out-Null }
    if ($seniorityFit -ne 0) { $notes.Add(("seniority {0}: {1}" -f $seniorityCategory, $seniorityFit)) | Out-Null }
    if ($contractFit -ne 0) { $notes.Add(("contract {0}: {1}" -f $contractCategory, $contractFit)) | Out-Null }

    $finalScore = [Math]::Max(0, $RoleScore + $employerFit + $locationFit + $seniorityFit + $contractFit)
    return [PSCustomObject]@{
        EmployerType      = $employerType
        RoleScore         = $RoleScore
        EmployerFit       = $employerFit
        LocationFit       = $locationFit
        SeniorityFit      = $seniorityFit
        ContractFit       = $contractFit
        LocationCategory  = $locationCategory
        SeniorityCategory = $seniorityCategory
        ContractCategory  = $contractCategory
        FinalScore        = [int]$finalScore
        MatchLevel        = Get-MatchLevelFromScore $finalScore
        Notes             = (($notes.ToArray() | Select-Object -Unique) -join "; ")
    }
}

function Test-IsAgencyConsultingEsnSignal {
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$CompanyName,
        [AllowNull()][string]$Text = ""
    )

    $employerType = Get-EmployerType -Title $Title -CompanyName $CompanyName -Text $Text
    return $employerType -in @("agency", "consulting", "esn")
}

function Test-IsAppliedStatus {
    param([AllowNull()][string]$Status)

    $statusText = ConvertTo-MatchText $Status
    return $statusText -match "^(applied|interview|offer|rejected|withdrawn)$"
}

function New-JobResult {
    param(
        [string]$Title,
        [string]$CompanyName,
        [string]$JobLocation,
        [string]$ContractType,
        [int]$MatchScore,
        [string]$MatchLevel,
        [string]$MatchedKeywords,
        [string]$Url,
        [string]$Platform,
        [AllowNull()]$PublishedAt,
        [AllowNull()][string]$SourceText = ""
    )

    if ([string]::IsNullOrWhiteSpace($Title) -or [string]::IsNullOrWhiteSpace($Url) -or $null -eq $PublishedAt) {
        return $null
    }

    $publishedDateValue = $null
    if ($PublishedAt -is [DateTimeOffset]) {
        $publishedDateValue = $PublishedAt
    }
    elseif ($PublishedAt -is [DateTime]) {
        $publishedDateValue = [DateTimeOffset]$PublishedAt
    }
    else {
        $publishedDateValue = ConvertTo-DateTimeOffsetOrNull ([string]$PublishedAt)
    }
    if ($null -eq $publishedDateValue) {
        return $null
    }

    $textContractType = Get-ContractTypeFromText -Text (Join-CleanTextParts @($Title, $SourceText))
    $effectiveContractType = $ContractType
    if ((Test-IsExcludedContractType $textContractType) -or [string]::IsNullOrWhiteSpace($effectiveContractType)) {
        $effectiveContractType = $textContractType
    }

    $key = "{0}|{1}" -f $Platform, $Url.ToLowerInvariant()
    if ($SeenResultKeys.ContainsKey($key)) {
        return $null
    }

    $identityKey = Get-JobIdentityKeyFromValues -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -Url $Url
    $jobId = Get-StableJobId $identityKey
    $fit = Get-JobFitDimensions -RoleScore $MatchScore -Title $Title -CompanyName $CompanyName -JobLocation $JobLocation -ContractType $effectiveContractType -Text $SourceText
    $adjustedScore = [int]$fit.FinalScore
    $adjustedKeywords = $MatchedKeywords.Trim()
    $fitKeywordNotes = New-Object System.Collections.Generic.List[string]
    if ([int]$fit.EmployerFit -lt 0) {
        $fitKeywordNotes.Add(("employer preference: {0}" -f $fit.EmployerType)) | Out-Null
    }
    if ([int]$fit.LocationFit -lt 0) {
        $fitKeywordNotes.Add(("location fit: {0}" -f $fit.LocationCategory)) | Out-Null
    }
    if ([int]$fit.SeniorityFit -lt 0) {
        $fitKeywordNotes.Add(("seniority fit: {0}" -f $fit.SeniorityCategory)) | Out-Null
    }
    if ([int]$fit.ContractFit -lt 0) {
        $fitKeywordNotes.Add(("contract fit: {0}" -f $fit.ContractCategory)) | Out-Null
    }
    if ($fitKeywordNotes.Count -gt 0) {
        $adjustedKeywords = (Join-CleanTextParts @($adjustedKeywords, (($fitKeywordNotes.ToArray()) -join "; "))) -replace ", ", "; "
    }

    $SeenResultKeys[$key] = $true
    [PSCustomObject]@{
        job_id         = $jobId
        job_title      = $Title.Trim()
        company_name   = $CompanyName.Trim()
        employer_type  = [string]$fit.EmployerType
        location       = $JobLocation.Trim()
        contract_type  = $effectiveContractType.Trim()
        match_score    = $adjustedScore
        match_level    = ([string]$fit.MatchLevel).Trim()
        matched_keywords = $adjustedKeywords
        role_score     = [string]$fit.RoleScore
        employer_fit   = [string]$fit.EmployerFit
        location_fit   = [string]$fit.LocationFit
        seniority_fit  = [string]$fit.SeniorityFit
        contract_fit   = [string]$fit.ContractFit
        fit_notes      = [string]$fit.Notes
        feedback_adjustment = ""
        job_url        = ConvertTo-ExcelHyperlinkFormula -Url $Url -Label "Open"
        job_url_raw    = $Url.Trim()
        platform       = $Platform
        source_count   = "1"
        alternate_urls = ""
        published_date = $publishedDateValue.ToString("yyyy-MM-dd")
    }
}

$MasterColumns = Get-JobTrackerMasterColumns
$ColumnLabels = Get-JobTrackerColumnLabels

function New-OrderedJobRecord {
    param([hashtable]$Values)

    $ordered = [ordered]@{}
    foreach ($column in $MasterColumns) {
        if ($Values.ContainsKey($column) -and $null -ne $Values[$column]) {
            $ordered[$column] = Repair-DisplayText ([string]$Values[$column])
        }
        else {
            $ordered[$column] = ""
        }
    }

    return [PSCustomObject]$ordered
}

function Get-PreferredValue {
    param(
        [AllowNull()][string]$Primary,
        [AllowNull()][string]$Fallback
    )

    if (-not [string]::IsNullOrWhiteSpace($Primary)) {
        return $Primary
    }

    return $Fallback
}

function Import-TrackerRowsFromXlsx {
    param([string]$Path)

    $excel = $null
    $workbook = $null
    $sheet = $null
    $usedRange = $null

    try {
        $fullPath = (Resolve-Path $Path).Path
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($fullPath, $null, $true)

        try {
            $sheet = $workbook.Worksheets.Item("Jobs")
        }
        catch {
            $sheet = $workbook.Worksheets.Item(1)
        }

        $usedRange = $sheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        if ($rowCount -lt 2 -or $columnCount -lt 1) {
            return @()
        }

        $headers = New-Object System.Collections.Generic.List[string]
        for ($column = 1; $column -le $columnCount; $column++) {
            $header = [string]$sheet.Cells.Item(1, $column).Text
            if ([string]::IsNullOrWhiteSpace($header)) {
                $header = "Column$column"
            }
            $headers.Add((ConvertTo-CanonicalColumnName $header.Trim())) | Out-Null
        }

        $rows = New-Object System.Collections.Generic.List[object]
        for ($row = 2; $row -le $rowCount; $row++) {
            $values = @{}
            $hasValue = $false
            for ($column = 1; $column -le $columnCount; $column++) {
                $name = $headers[$column - 1]
                $value = Repair-DisplayText ([string]$sheet.Cells.Item($row, $column).Text)
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $hasValue = $true
                }
                $values[$name] = $value
            }

            if ($hasValue) {
                $rows.Add((New-OrderedJobRecord $values)) | Out-Null
            }
        }

        return @($rows.ToArray())
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit() | Out-Null
        }

        Release-ComObject $usedRange
        Release-ComObject $sheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-SummaryValue {
    param(
        [AllowNull()]$Summary,
        [string]$Name
    )

    if ($null -ne $Summary -and $Summary.ContainsKey($Name)) {
        return [string]$Summary[$Name]
    }

    return ""
}

function Export-TrackerWorkbook {
    param(
        [object[]]$Rows,
        [string]$Path,
        [AllowNull()]$Summary = $null
    )

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    $fullPath = [IO.Path]::GetFullPath($Path)
    $excel = $null
    $workbook = $null
    $jobsSheet = $null
    $summarySheet = $null
    $tableRange = $null
    $table = $null
    $dataRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()

        while ([int]$workbook.Worksheets.Count -gt 1) {
            $workbook.Worksheets.Item([int]$workbook.Worksheets.Count).Delete()
        }

        $jobsSheet = $workbook.Worksheets.Item(1)
        $jobsSheet.Name = "Jobs"
        $summarySheet = $workbook.Worksheets.Add([System.Type]::Missing, $jobsSheet)
        $summarySheet.Name = "Summary"

        $columnIndex = @{}
        for ($index = 0; $index -lt $MasterColumns.Count; $index++) {
            $columnNumber = $index + 1
            $columnName = $MasterColumns[$index]
            $columnIndex[$columnName] = $columnNumber
            $jobsSheet.Cells.Item(1, $columnNumber).Value2 = Get-ColumnLabel $columnName
        }

        $rowCount = @($Rows).Count
        for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
            $row = $Rows[$rowIndex]
            $excelRow = $rowIndex + 2
            foreach ($columnName in $MasterColumns) {
                $excelColumn = [int]$columnIndex[$columnName]
                $cell = $jobsSheet.Cells.Item($excelRow, $excelColumn)
                $value = Repair-DisplayText (Get-RowValue -Row $row -Name $columnName)

                if ($columnName -eq "job_url") {
                    $url = Get-RowValue -Row $row -Name "job_url_raw"
                    if ([string]::IsNullOrWhiteSpace($url) -and $value -match "^https?://") {
                        $url = $value
                    }

                    if ($url -match "^https?://") {
                        $escapedUrl = $url.Replace('"', '""')
                        $cell.Formula = '=HYPERLINK("{0}","Open")' -f $escapedUrl
                    }
                    else {
                        $cell.Value2 = $value
                    }
                }
                elseif ($columnName -in @("match_score", "role_score", "employer_fit", "location_fit", "seniority_fit", "contract_fit", "days_since_published", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
                    $number = 0
                    if ([int]::TryParse($value, [ref]$number)) {
                        $cell.Value2 = [string]$number
                    }
                    else {
                        $cell.Value2 = $value
                    }
                }
                else {
                    $cell.Value2 = $value
                }
            }
        }

        $lastDataRow = [Math]::Max(2, $rowCount + 1)
        $lastColumn = $MasterColumns.Count
        $tableRange = $jobsSheet.Range($jobsSheet.Cells.Item(1, 1), $jobsSheet.Cells.Item($lastDataRow, $lastColumn))
        try {
            $table = $jobsSheet.ListObjects.Add(1, $tableRange, $null, 1)
            $table.Name = "JobsTracker"
            $table.TableStyle = "TableStyleLight9"
            $table.ShowTableStyleRowStripes = $false
            $table.ShowTableStyleColumnStripes = $false
        }
        catch {
            $jobsSheet.Range($jobsSheet.Cells.Item(1, 1), $jobsSheet.Cells.Item(1, $lastColumn)).AutoFilter() | Out-Null
        }

        $headerColor = Get-ExcelColor 38 50 56
        $darkTextColor = Get-ExcelColor 40 47 52
        $mutedTextColor = Get-ExcelColor 100 116 139
        $jobsSheet.Rows.Item(1).Font.Bold = $true
        $jobsSheet.Rows.Item(1).Font.Color = Get-ExcelColor 255 255 255
        $jobsSheet.Rows.Item(1).Interior.Color = $headerColor
        $jobsSheet.Rows.Item(1).VerticalAlignment = -4108
        $jobsSheet.Rows.Item(1).RowHeight = 24
        $jobsSheet.Cells.Font.Name = "Segoe UI"
        $jobsSheet.Cells.Font.Size = 10
        $jobsSheet.Cells.Font.Color = $darkTextColor
        $jobsSheet.Rows.Item(1).Font.Color = Get-ExcelColor 255 255 255

        foreach ($columnName in @("duplicate_reason", "job_title", "matched_keywords", "fit_notes", "job_url_raw", "notes")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).WrapText = $true
            }
        }
        $jobsSheet.Columns.AutoFit() | Out-Null
        $columnSizing = Get-JobTrackerColumnSizing
        foreach ($entry in $columnSizing.GetEnumerator()) {
            if ($columnIndex.ContainsKey($entry.Key)) {
                $column = $jobsSheet.Columns.Item([int]$columnIndex[$entry.Key])
                $minWidth = [double]$entry.Value.Min
                $maxWidth = [double]$entry.Value.Max
                if ([double]$column.ColumnWidth -lt $minWidth) {
                    $column.ColumnWidth = $minWidth
                }
                elseif ([double]$column.ColumnWidth -gt $maxWidth) {
                    $column.ColumnWidth = $maxWidth
                }
            }
        }
        Set-JobTrackerColumnVisibility -Sheet $jobsSheet -ColumnIndex $columnIndex
        foreach ($columnName in @("review_priority", "status", "employer_type", "contract_type", "platform", "source_count", "published_date", "days_since_published", "job_url", "applied_date", "match_level", "match_score", "role_score", "employer_fit", "location_fit", "seniority_fit", "contract_fit", "seen_in_current_crawl", "first_seen_date", "last_seen_date", "is_new")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).HorizontalAlignment = -4108
            }
        }
        foreach ($columnName in @("published_date", "applied_date", "first_seen_date", "last_seen_date")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).NumberFormat = "yyyy-mm-dd"
            }
        }
        foreach ($columnName in @("source_count", "days_since_published", "match_score", "role_score", "employer_fit", "location_fit", "seniority_fit", "contract_fit", "days_since_first_seen", "days_since_last_seen", "feedback_adjustment")) {
            if ($columnIndex.ContainsKey($columnName)) {
                $jobsSheet.Columns.Item([int]$columnIndex[$columnName]).NumberFormat = "0"
            }
        }

        Set-JobTrackerDataValidation -Workbook $workbook -Excel $excel -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
        Set-ReviewPriorityFormulas -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow

        if ($rowCount -gt 0) {
            $dataRange = $jobsSheet.Range($jobsSheet.Cells.Item(2, 1), $jobsSheet.Cells.Item($lastDataRow, $lastColumn))
            $dataRange.VerticalAlignment = -4160
            $dataRange.RowHeight = 30
            $dataRange.Interior.Color = Get-ExcelColor 255 255 255
            Set-StatusRowConditionalFormatting -Range $dataRange -ColumnIndex $columnIndex
            Set-StatusCellConditionalFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
            Set-ReviewPriorityConditionalFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
            Set-IgnoredNotesReminderFormatting -Sheet $jobsSheet -ColumnIndex $columnIndex -LastDataRow $lastDataRow
        }

        $colors = Get-JobTrackerWorkbookColors -DarkTextColor $darkTextColor
        for ($excelRow = 2; $excelRow -le ($rowCount + 1); $excelRow++) {
            $matchLevel = Get-RowValue -Row $Rows[$excelRow - 2] -Name "match_level"
            $seen = Get-RowValue -Row $Rows[$excelRow - 2] -Name "seen_in_current_crawl"

            if ($columnIndex.ContainsKey("match_level")) {
                $matchCell = $jobsSheet.Cells.Item($excelRow, [int]$columnIndex["match_level"])
                $matchCell.Font.Bold = $true
                switch ($matchLevel) {
                    "High" { $matchCell.Font.Color = $colors.GreenText }
                    "Medium" { $matchCell.Font.Color = $colors.AmberText }
                    "Review" { $matchCell.Font.Color = $colors.GrayText; $matchCell.Font.Bold = $false }
                    default { $matchCell.Font.Color = $colors.DarkText }
                }
            }

            if ($seen -eq "no" -and $columnIndex.ContainsKey("seen_in_current_crawl")) {
                $jobsSheet.Cells.Item($excelRow, [int]$columnIndex["seen_in_current_crawl"]).Font.Color = $mutedTextColor
            }
        }

        $jobsSheet.Activate() | Out-Null
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.SplitColumn = 2
        $excel.ActiveWindow.FreezePanes = $true
        $excel.ActiveWindow.DisplayGridlines = $false

        $summarySheet.Cells.Item(1, 1).Value2 = "Analytics Job Tracker"
        $summarySheet.Cells.Item(1, 1).Font.Bold = $true
        $summarySheet.Cells.Item(1, 1).Font.Size = 16
        $summarySheet.Cells.Font.Name = "Segoe UI"
        $summarySheet.Cells.Font.Size = 10
        $summarySheet.Cells.Font.Color = Get-ExcelColor 40 47 52
        $summarySheet.Cells.Item(1, 1).Font.Bold = $true
        $summarySheet.Cells.Item(1, 1).Font.Size = 16
        $currentVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "seen_in_current_crawl") -eq "yes" }).Count
        $newVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "is_new") -eq "yes" }).Count
        $applicationVisibleCount = @($Rows | Where-Object { Test-IsAppliedStatus (Get-RowValue -Row $_ -Name "status") }).Count
        $highVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "match_level") -eq "High" }).Count
        $mediumVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "match_level") -eq "Medium" }).Count
        $reviewVisibleCount = @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "match_level") -eq "Review" }).Count
        $employerTypeSummary = @(
            "annonceur {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "annonceur" }).Count
            "agency {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "agency" }).Count
            "consulting {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "consulting" }).Count
            "ESN {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "employer_type") -eq "esn" }).Count
            "unknown {0}" -f @($Rows | Where-Object { [string]::IsNullOrWhiteSpace((Get-RowValue -Row $_ -Name "employer_type")) -or (Get-RowValue -Row $_ -Name "employer_type") -eq "unknown" }).Count
        ) -join " | "
        $fitDemotionSummary = @(
            "employer {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "employer_fit") -lt 0 }).Count
            "location {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "location_fit") -lt 0 }).Count
            "seniority {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "seniority_fit") -lt 0 }).Count
            "contract {0}" -f @($Rows | Where-Object { (Get-IntegerRowValue -Row $_ -Name "contract_fit") -lt 0 }).Count
        ) -join " | "
        $sourceSummary = @(
            "France Travail {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "France Travail" }).Count
            "Adzuna {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "Adzuna" }).Count
            "APEC {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "APEC" }).Count
            "HelloWork {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "HelloWork" }).Count
            "WTTJ {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "Welcome to the Jungle" }).Count
            "LinkedIn {0}" -f @($Rows | Where-Object { (Get-RowValue -Row $_ -Name "platform") -match "LinkedIn" }).Count
        ) -join " | "
        $summaryPairs = @(
            @("Generated", $RunStamp),
            @("Crawl mode", $CrawlMode),
            @("Retention rule", "Keep non-application jobs only when Published is on or after $CutoffDate."),
            @("Rows in workbook", [string](@($Rows).Count)),
            @("Seen in this crawl", [string]$currentVisibleCount),
            @("New this run", [string]$newVisibleCount),
            @("Application rows kept", [string]$applicationVisibleCount),
            @("Match levels", ("High {0} | Medium {1} | Review {2}" -f $highVisibleCount, $mediumVisibleCount, $reviewVisibleCount)),
            @("Sources", $sourceSummary),
            @("Employer types", $employerTypeSummary),
            @("Fit demotions", $fitDemotionSummary),
            @("Total matched before contract filter", (Get-SummaryValue -Summary $Summary -Name "TotalMatched")),
            @("Excluded CDD/apprenticeship/internship/freelance", (Get-SummaryValue -Summary $Summary -Name "ExcludedContractCount")),
            @("Duplicates merged this run", (Get-SummaryValue -Summary $Summary -Name "DuplicateCount")),
            @("Rows removed by retention", (Get-SummaryValue -Summary $Summary -Name "RemovedCount")),
            @("Source diagnostics", (Get-SummaryValue -Summary $Summary -Name "SourceDiagnostics")),
            @("Backup", (Get-SummaryValue -Summary $Summary -Name "BackupPath")),
            @("Tracker", $fullPath),
            @("Manual fields", "Status, Applied date, Apply notes with ignore_reason templates"),
            @("Reminder", "Close this workbook before launching the crawler.")
        )
        $summaryRow = 3
        foreach ($pair in $summaryPairs) {
            $summarySheet.Cells.Item($summaryRow, 1).Value2 = [string]$pair[0]
            $summarySheet.Cells.Item($summaryRow, 2).Value2 = [string]$pair[1]
            $summaryRow++
        }
        $summarySheet.Rows.Item(1).Font.Color = Get-ExcelColor 38 50 56
        $summarySheet.Columns.Item(1).ColumnWidth = 36
        $summarySheet.Columns.Item(2).ColumnWidth = 90
        $summarySheet.Columns.Item(2).WrapText = $true
        $summarySheet.Range("A3:A$($summaryRow - 1)").Font.Bold = $true
        $summarySheet.Range("A3:B$($summaryRow - 1)").Borders.LineStyle = 1
        $summarySheet.Range("A3:B$($summaryRow - 1)").Borders.Color = Get-ExcelColor 226 232 240

        $workbook.Worksheets.Item("Jobs").Activate() | Out-Null
        $workbook.SaveAs($fullPath, 51)
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit() | Out-Null
        }

        Release-ComObject $table
        Release-ComObject $tableRange
        Release-ComObject $dataRange
        Release-ComObject $summarySheet
        Release-ComObject $jobsSheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function ConvertTo-DateOrNull {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
        return $parsed.Date
    }

    return $null
}

function Get-DaysSince {
    param([AllowNull()][string]$DateText)

    $date = ConvertTo-DateOrNull $DateText
    if ($null -eq $date) {
        return ""
    }

    return [string]([Math]::Max(0, ([DateTime]::Today - $date).Days))
}

function Test-IsKeepForeverStatus {
    param([AllowNull()][string]$Status)

    return Test-IsAppliedStatus $Status
}

function Test-IsRecentTrackerRow {
    param([AllowNull()]$Row)

    $publishedDate = ConvertTo-DateOrNull (Get-RowValue -Row $Row -Name "published_date")
    return ($null -ne $publishedDate -and $publishedDate -ge $Cutoff.Date)
}

function Get-IntegerRowValue {
    param(
        [AllowNull()]$Row,
        [string]$Name
    )

    $value = 0
    if ([int]::TryParse((Get-RowValue -Row $Row -Name $Name), [ref]$value)) {
        return $value
    }

    return 0
}

function Get-SourcePreference {
    param([AllowNull()]$Row)

    $platform = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "platform")
    $url = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "job_url_raw")
    if ($platform -match "welcome|jungle|wttj") {
        return 50
    }
    if ($platform -match "\bapec\b" -or $url -match "apec\.fr") {
        return 45
    }
    if ($platform -match "france\s+travail" -or $url -match "francetravail|pole-emploi") {
        return 40
    }
    if ($platform -match "hellowork" -or $url -match "hellowork") {
        return 35
    }
    if ($platform -match "linkedin") {
        return 30
    }
    if ($platform -match "adzuna") {
        return 15
    }

    return 10
}

function Get-DateSortValue {
    param(
        [AllowNull()]$Row,
        [string]$Name
    )

    $date = ConvertTo-DateOrNull (Get-RowValue -Row $Row -Name $Name)
    if ($null -eq $date) {
        return [DateTime]::MinValue
    }

    return $date
}

function Select-PreferredJobRow {
    param([object[]]$Rows)

    return @($Rows) |
        Sort-Object -Property `
            @{ Expression = { if (Test-IsAppliedStatus (Get-RowValue -Row $_ -Name "status")) { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { if ((ConvertTo-MatchText (Get-RowValue -Row $_ -Name "status")) -eq "interesting") { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { Get-IntegerRowValue -Row $_ -Name "match_score" }; Descending = $true },
            @{ Expression = { Get-DateSortValue -Row $_ -Name "published_date" }; Descending = $true },
            @{ Expression = { Get-SourcePreference $_ }; Descending = $true } |
        Select-Object -First 1
}

function Select-PreferredUrlRow {
    param([object[]]$Rows)

    return @($Rows) |
        Where-Object { -not [string]::IsNullOrWhiteSpace((Get-RowValue -Row $_ -Name "job_url_raw")) } |
        Sort-Object -Property `
            @{ Expression = { Get-SourcePreference $_ }; Descending = $true },
            @{ Expression = { Get-IntegerRowValue -Row $_ -Name "match_score" }; Descending = $true },
            @{ Expression = { Get-DateSortValue -Row $_ -Name "published_date" }; Descending = $true } |
        Select-Object -First 1
}

function Get-UniqueTextValues {
    param(
        [object[]]$Values,
        [string]$SplitPattern = "\s*;\s*"
    )

    $seen = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($value in @($Values)) {
        if ($null -eq $value) {
            continue
        }

        foreach ($part in ([string]$value -split $SplitPattern)) {
            $text = ([string]$part).Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            $key = $text.ToLowerInvariant()
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $result.Add($text) | Out-Null
            }
        }
    }

    return @($result.ToArray())
}

function Join-UniqueTextValues {
    param(
        [object[]]$Values,
        [string]$Delimiter = "; ",
        [string]$SplitPattern = "\s*;\s*"
    )

    return (Get-UniqueTextValues -Values $Values -SplitPattern $SplitPattern) -join $Delimiter
}

function Get-RowUrlValues {
    param([object[]]$Rows)

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($Rows)) {
        $primary = Get-RowValue -Row $row -Name "job_url_raw"
        if (-not [string]::IsNullOrWhiteSpace($primary)) {
            $values.Add($primary) | Out-Null
        }

        $alternate = Get-RowValue -Row $row -Name "alternate_urls"
        foreach ($url in (Get-UniqueTextValues -Values @($alternate))) {
            $values.Add($url) | Out-Null
        }
    }

    return Get-UniqueTextValues -Values $values
}

function Get-RowPlatformValues {
    param([object[]]$Rows)

    return Get-UniqueTextValues -Values @($Rows | ForEach-Object { Get-RowValue -Row $_ -Name "platform" })
}

function Get-SourceCountFromRows {
    param([object[]]$Rows)

    $platforms = @(Get-RowPlatformValues $Rows)
    if ($platforms.Count -gt 0) {
        return [Math]::Max(1, $platforms.Count)
    }

    $urls = @(Get-RowUrlValues $Rows)
    return [Math]::Max(1, $urls.Count)
}

function Get-LatestDateText {
    param(
        [object[]]$Rows,
        [string]$Name
    )

    $dates = foreach ($row in @($Rows)) {
        $date = ConvertTo-DateOrNull (Get-RowValue -Row $row -Name $Name)
        if ($null -ne $date) {
            $date
        }
    }

    if ($null -eq $dates) {
        return ""
    }

    $latest = @($dates | Sort-Object -Descending | Select-Object -First 1)
    if ($latest.Count -eq 0) {
        return ""
    }

    return $latest[0].ToString("yyyy-MM-dd")
}

function Get-EarliestDateText {
    param(
        [object[]]$Rows,
        [string]$Name
    )

    $dates = foreach ($row in @($Rows)) {
        $date = ConvertTo-DateOrNull (Get-RowValue -Row $row -Name $Name)
        if ($null -ne $date) {
            $date
        }
    }

    if ($null -eq $dates) {
        return ""
    }

    $earliest = @($dates | Sort-Object | Select-Object -First 1)
    if ($earliest.Count -eq 0) {
        return ""
    }

    return $earliest[0].ToString("yyyy-MM-dd")
}

function Get-MatchLevelFromScore {
    param([int]$Score)

    if ($Score -ge 80) {
        return "High"
    }
    if ($Score -ge 50) {
        return "Medium"
    }

    return "Review"
}

function Merge-SimilarJobRows {
    param(
        [object[]]$Rows,
        [string]$Reason
    )

    $rowList = @($Rows | Where-Object { $null -ne $_ })
    if ($rowList.Count -eq 0) {
        return $null
    }
    if ($rowList.Count -eq 1 -and [string]::IsNullOrWhiteSpace($Reason)) {
        return $rowList[0]
    }

    $preferred = Select-PreferredJobRow $rowList
    $preferredUrlRow = Select-PreferredUrlRow $rowList
    if ($null -eq $preferredUrlRow) {
        $preferredUrlRow = $preferred
    }

    $values = @{}
    foreach ($column in $MasterColumns) {
        $values[$column] = Get-RowValue -Row $preferred -Name $column
    }

    $urls = @(Get-RowUrlValues $rowList)
    $primaryUrl = Get-RowValue -Row $preferredUrlRow -Name "job_url_raw"
    if ([string]::IsNullOrWhiteSpace($primaryUrl) -and $urls.Count -gt 0) {
        $primaryUrl = $urls[0]
    }
    $alternateUrls = @($urls | Where-Object { $_ -ne $primaryUrl })

    $maxScore = 0
    $maxRoleScore = 0
    $maxEmployerFit = -999
    $maxLocationFit = -999
    $maxSeniorityFit = -999
    $maxContractFit = -999
    foreach ($row in $rowList) {
        $score = Get-IntegerRowValue -Row $row -Name "match_score"
        if ($score -gt $maxScore) {
            $maxScore = $score
        }
        $roleScore = Get-IntegerRowValue -Row $row -Name "role_score"
        if ($roleScore -gt $maxRoleScore) {
            $maxRoleScore = $roleScore
        }
        $employerFit = Get-IntegerRowValue -Row $row -Name "employer_fit"
        if ($employerFit -gt $maxEmployerFit) {
            $maxEmployerFit = $employerFit
        }
        $locationFit = Get-IntegerRowValue -Row $row -Name "location_fit"
        if ($locationFit -gt $maxLocationFit) {
            $maxLocationFit = $locationFit
        }
        $seniorityFit = Get-IntegerRowValue -Row $row -Name "seniority_fit"
        if ($seniorityFit -gt $maxSeniorityFit) {
            $maxSeniorityFit = $seniorityFit
        }
        $contractFit = Get-IntegerRowValue -Row $row -Name "contract_fit"
        if ($contractFit -gt $maxContractFit) {
            $maxContractFit = $contractFit
        }
    }

    $dedupeKey = Get-JobDedupeKeyFromRow $preferred
    $jobId = Get-RowValue -Row $preferred -Name "job_id"
    if ([string]::IsNullOrWhiteSpace($jobId)) {
        $jobId = Get-StableJobId $dedupeKey
    }

    $reasonParts = @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "duplicate_reason" })
    if (-not [string]::IsNullOrWhiteSpace($Reason)) {
        $reasonParts += $Reason
    }

    $values["job_id"] = $jobId
    $values["job_url_raw"] = $primaryUrl
    $values["job_url"] = ConvertTo-ExcelHyperlinkFormula -Url $primaryUrl -Label "Open"
    $values["alternate_urls"] = ($alternateUrls -join "; ")
    $values["source_count"] = [string](Get-SourceCountFromRows $rowList)
    $values["platform"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "platform" })
    $values["matched_keywords"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "matched_keywords" }) -SplitPattern "\s*;\s*|\s*,\s*"
    $employerTypes = @(Get-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "employer_type" }))
    foreach ($employerTypeCandidate in @("annonceur", "consulting", "agency", "esn", "unknown")) {
        if ($employerTypes -contains $employerTypeCandidate) {
            $values["employer_type"] = $employerTypeCandidate
            break
        }
    }
    $values["duplicate_reason"] = Join-CleanTextParts $reasonParts
    $values["match_score"] = [string]$maxScore
    $values["match_level"] = Get-MatchLevelFromScore $maxScore
    $values["role_score"] = [string]$maxRoleScore
    $values["employer_fit"] = $(if ($maxEmployerFit -eq -999) { "" } else { [string]$maxEmployerFit })
    $values["location_fit"] = $(if ($maxLocationFit -eq -999) { "" } else { [string]$maxLocationFit })
    $values["seniority_fit"] = $(if ($maxSeniorityFit -eq -999) { "" } else { [string]$maxSeniorityFit })
    $values["contract_fit"] = $(if ($maxContractFit -eq -999) { "" } else { [string]$maxContractFit })
    $values["fit_notes"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "fit_notes" }) -SplitPattern "\s*;\s*"
    $values["published_date"] = Get-LatestDateText -Rows $rowList -Name "published_date"
    $values["first_seen_date"] = Get-EarliestDateText -Rows $rowList -Name "first_seen_date"
    $values["last_seen_date"] = Get-LatestDateText -Rows $rowList -Name "last_seen_date"
    $values["applied_date"] = Get-EarliestDateText -Rows $rowList -Name "applied_date"
    $values["notes"] = Join-UniqueTextValues -Values @($rowList | ForEach-Object { Get-RowValue -Row $_ -Name "notes" }) -Delimiter " | "

    return New-OrderedJobRecord $values
}

function Group-RowsByDedupeKey {
    param([object[]]$Rows)

    $groups = @{}
    foreach ($row in @($Rows)) {
        $key = Get-JobDedupeKeyFromRow $row
        if ([string]::IsNullOrWhiteSpace($key)) {
            $key = "jobid|{0}" -f (Get-RowValue -Row $row -Name "job_id")
        }

        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }
        $groups[$key].Add($row) | Out-Null
    }

    return $groups
}

function Backup-TrackerFile {
    param([string]$Path)

    return Backup-JobTrackerFile -Path $Path -MaxBackups $MaxBackups
}

function Import-TrackerRows {
    param([string]$Path)

    if (Test-Path $Path) {
        if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne ".xlsx") {
            throw "Unsupported tracker file type. This project uses only output\jobs_tracker.xlsx."
        }

        return @(Import-TrackerRowsFromXlsx -Path $Path)
    }

    return @()
}

function Get-ReviewPriority {
    param(
        [AllowNull()][string]$Status,
        [AllowNull()][string]$MatchLevel,
        [AllowNull()][string]$IsNew
    )

    if (Test-IsKeepForeverStatus $Status) {
        return "Application"
    }
    if ((ConvertTo-MatchText $Status) -eq "ignored") {
        return "Ignored"
    }
    if ($IsNew -eq "yes" -and $MatchLevel -eq "High") {
        return "New High"
    }
    if ($IsNew -eq "yes") {
        return "New"
    }
    if ($MatchLevel -eq "High") {
        return "High"
    }
    return $MatchLevel
}

function ConvertTo-TrackerRecord {
    param(
        $CurrentRow,
        [AllowNull()]$ExistingRow = $null,
        [bool]$SeenInCurrentCrawl = $true,
        [AllowNull()][string]$DuplicateReason = ""
    )

    $existingJobId = Get-RowValue -Row $ExistingRow -Name "job_id"
    $currentJobId = Get-RowValue -Row $CurrentRow -Name "job_id"
    $jobId = Get-PreferredValue -Primary $existingJobId -Fallback (Get-PreferredValue -Primary $currentJobId -Fallback (Get-StableJobId (Get-JobIdentityKeyFromRow $CurrentRow)))

    $status = Get-RowValue -Row $ExistingRow -Name "status"
    if ([string]::IsNullOrWhiteSpace($status)) {
        $status = "new"
    }

    $firstSeen = Get-RowValue -Row $ExistingRow -Name "first_seen_date"
    $isNew = "no"
    $seenBefore = "yes"
    if ([string]::IsNullOrWhiteSpace($firstSeen)) {
        $firstSeen = $RunDate
        $isNew = "yes"
        $seenBefore = "no"
    }

    $lastSeen = Get-RowValue -Row $ExistingRow -Name "last_seen_date"
    if ($SeenInCurrentCrawl) {
        $lastSeen = $RunDate
    }
    elseif ([string]::IsNullOrWhiteSpace($lastSeen)) {
        $lastSeen = $firstSeen
    }

    $matchLevel = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "match_level") -Fallback (Get-RowValue -Row $ExistingRow -Name "match_level")
    $duplicateValue = Get-PreferredValue -Primary $DuplicateReason -Fallback (Get-RowValue -Row $ExistingRow -Name "duplicate_reason")
    $publishedDate = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "published_date") -Fallback (Get-RowValue -Row $ExistingRow -Name "published_date")
    $preferredUrlRow = Select-PreferredUrlRow @($CurrentRow, $ExistingRow)
    $primaryUrl = Get-RowValue -Row $preferredUrlRow -Name "job_url_raw"
    if ([string]::IsNullOrWhiteSpace($primaryUrl)) {
        $primaryUrl = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "job_url_raw") -Fallback (Get-RowValue -Row $ExistingRow -Name "job_url_raw")
    }
    $allUrls = @(Get-RowUrlValues @($CurrentRow, $ExistingRow))
    $alternateUrls = @($allUrls | Where-Object { $_ -ne $primaryUrl })
    $jobTitleValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "job_title") -Fallback (Get-RowValue -Row $ExistingRow -Name "job_title")
    $companyNameValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "company_name") -Fallback (Get-RowValue -Row $ExistingRow -Name "company_name")
    $locationValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "location") -Fallback (Get-RowValue -Row $ExistingRow -Name "location")
    $contractTypeValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "contract_type") -Fallback (Get-RowValue -Row $ExistingRow -Name "contract_type")
    $matchScoreValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "match_score") -Fallback (Get-RowValue -Row $ExistingRow -Name "match_score")
    $matchedKeywordsValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "matched_keywords") -Fallback (Get-RowValue -Row $ExistingRow -Name "matched_keywords")
    $employerTypeValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "employer_type") -Fallback (Get-RowValue -Row $ExistingRow -Name "employer_type")
    $roleScoreValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "role_score") -Fallback (Get-RowValue -Row $ExistingRow -Name "role_score")
    $employerFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "employer_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "employer_fit")
    $locationFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "location_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "location_fit")
    $seniorityFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "seniority_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "seniority_fit")
    $contractFitValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "contract_fit") -Fallback (Get-RowValue -Row $ExistingRow -Name "contract_fit")
    $fitNotesValue = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "fit_notes") -Fallback (Get-RowValue -Row $ExistingRow -Name "fit_notes")

    if ([string]::IsNullOrWhiteSpace($employerTypeValue) -or
        [string]::IsNullOrWhiteSpace($roleScoreValue) -or
        [string]::IsNullOrWhiteSpace($employerFitValue) -or
        [string]::IsNullOrWhiteSpace($locationFitValue) -or
        [string]::IsNullOrWhiteSpace($seniorityFitValue) -or
        [string]::IsNullOrWhiteSpace($contractFitValue) -or
        [string]::IsNullOrWhiteSpace($fitNotesValue)) {
        $roleBaseScore = 0
        if (-not [int]::TryParse($roleScoreValue, [ref]$roleBaseScore)) {
            [void][int]::TryParse($matchScoreValue, [ref]$roleBaseScore)
        }
        $fit = Get-JobFitDimensions `
            -RoleScore $roleBaseScore `
            -Title $jobTitleValue `
            -CompanyName $companyNameValue `
            -JobLocation $locationValue `
            -ContractType $contractTypeValue `
            -Text (Join-CleanTextParts @($matchedKeywordsValue, (Get-RowValue -Row $ExistingRow -Name "notes")))

        if ([string]::IsNullOrWhiteSpace($employerTypeValue)) { $employerTypeValue = [string]$fit.EmployerType }
        if ([string]::IsNullOrWhiteSpace($roleScoreValue)) { $roleScoreValue = [string]$fit.RoleScore }
        if ([string]::IsNullOrWhiteSpace($employerFitValue)) { $employerFitValue = [string]$fit.EmployerFit }
        if ([string]::IsNullOrWhiteSpace($locationFitValue)) { $locationFitValue = [string]$fit.LocationFit }
        if ([string]::IsNullOrWhiteSpace($seniorityFitValue)) { $seniorityFitValue = [string]$fit.SeniorityFit }
        if ([string]::IsNullOrWhiteSpace($contractFitValue)) { $contractFitValue = [string]$fit.ContractFit }
        if ([string]::IsNullOrWhiteSpace($fitNotesValue)) { $fitNotesValue = [string]$fit.Notes }
    }

    return New-OrderedJobRecord @{
        job_id                = $jobId
        status                = $status
        applied_date          = Get-RowValue -Row $ExistingRow -Name "applied_date"
        first_seen_date       = $firstSeen
        last_seen_date        = $lastSeen
        is_new                = $isNew
        seen_before           = $seenBefore
        seen_in_current_crawl = $(if ($SeenInCurrentCrawl) { "yes" } else { "no" })
        days_since_first_seen = Get-DaysSince $firstSeen
        days_since_last_seen  = Get-DaysSince $lastSeen
        duplicate_reason      = $duplicateValue
        feedback_adjustment   = Get-PreferredValue -Primary (Get-RowValue -Row $CurrentRow -Name "feedback_adjustment") -Fallback (Get-RowValue -Row $ExistingRow -Name "feedback_adjustment")
        review_priority       = Get-ReviewPriority -Status $status -MatchLevel $matchLevel -IsNew $isNew
        job_title             = $jobTitleValue
        company_name          = $companyNameValue
        employer_type         = $employerTypeValue
        location              = $locationValue
        contract_type         = $contractTypeValue
        match_score           = $matchScoreValue
        match_level           = $matchLevel
        matched_keywords      = $matchedKeywordsValue
        role_score            = $roleScoreValue
        employer_fit          = $employerFitValue
        location_fit          = $locationFitValue
        seniority_fit         = $seniorityFitValue
        contract_fit          = $contractFitValue
        fit_notes             = $fitNotesValue
        job_url               = ConvertTo-ExcelHyperlinkFormula -Url $primaryUrl -Label "Open"
        job_url_raw           = $primaryUrl
        alternate_urls        = ($alternateUrls -join "; ")
        platform              = Join-UniqueTextValues -Values @((Get-RowValue -Row $CurrentRow -Name "platform"), (Get-RowValue -Row $ExistingRow -Name "platform"))
        source_count          = [string](Get-SourceCountFromRows @($CurrentRow, $ExistingRow))
        published_date        = $publishedDate
        days_since_published  = Get-DaysSince $publishedDate
        notes                 = Get-RowValue -Row $ExistingRow -Name "notes"
    }
}

function ConvertTo-TrackerRecordFromExisting {
    param($ExistingRow)

    return ConvertTo-TrackerRecord -CurrentRow $ExistingRow -ExistingRow $ExistingRow -SeenInCurrentCrawl:$false
}

function Get-FeedbackProfileText {
    param([AllowNull()]$Row)

    return ConvertTo-MatchText (Join-CleanTextParts @(
        (Get-RowValue -Row $Row -Name "job_title"),
        (Get-RowValue -Row $Row -Name "company_name"),
        (Get-RowValue -Row $Row -Name "location"),
        (Get-RowValue -Row $Row -Name "contract_type"),
        (Get-RowValue -Row $Row -Name "matched_keywords"),
        (Get-RowValue -Row $Row -Name "notes")
    ))
}

function Test-FeedbackRowHasAgencyConsultingEsnSignal {
    param([AllowNull()]$Row)

    return Test-IsAgencyConsultingEsnSignal `
        -Title (Get-RowValue -Row $Row -Name "job_title") `
        -CompanyName (Get-RowValue -Row $Row -Name "company_name") `
        -Text (Join-CleanTextParts @(
            (Get-RowValue -Row $Row -Name "matched_keywords"),
            (Get-RowValue -Row $Row -Name "notes")
        ))
}

function Get-FeedbackSeniorityBucket {
    param([string]$Text)

    if ($Text -match "\b(stage|intern|internship|apprentice|apprentissage|alternance|junior|graduate)\b") {
        return "junior"
    }
    if ($Text -match "\b(head|director|directeur|directrice|lead|manager|responsable|principal)\b") {
        return "management"
    }
    if ($Text -match "\b(senior|sr)\b") {
        return "senior"
    }

    return ""
}

function Test-FeedbackTextHasWebAnalyticsSignal {
    param([string]$Text)

    return $Text -match "web\s+analytics|digital\s+analytics|web\s*analyst|digital\s*analyst|tracking|tagging|taggage|webtracking|google\s+tag\s+manager|\bgtm\b|google\s+analytics|\bga4\b|piano|contentsquare|content\s+square|tag\s+commander|commanders?\s+act|\btealium\b|data\s*layer|datalayer|tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)|server\s*[- ]?\s*side|consent\s+mode|\brgpd\b|\bgdpr\b|matomo|adobe\s+analytics"
}

function Get-FeedbackSignalDefinitions {
    return @(
        [PSCustomObject]@{ Key = "google_tag_manager"; Label = "feedback positive: Google Tag Manager"; Pattern = "google\s+tag\s+manager|\bgtm\b" },
        [PSCustomObject]@{ Key = "google_analytics"; Label = "feedback positive: Google Analytics/GA4"; Pattern = "google\s+analytics|\bga4\b" },
        [PSCustomObject]@{ Key = "piano"; Label = "feedback positive: Piano"; Pattern = "piano" },
        [PSCustomObject]@{ Key = "contentsquare"; Label = "feedback positive: ContentSquare"; Pattern = "contentsquare|content\s+square" },
        [PSCustomObject]@{ Key = "tag_commander"; Label = "feedback positive: Tag Commander/Commanders Act"; Pattern = "tag\s+commander|commanders?\s+act" },
        [PSCustomObject]@{ Key = "tealium"; Label = "feedback positive: Tealium"; Pattern = "\btealium\b|tealium\s+iq" },
        [PSCustomObject]@{ Key = "server_side"; Label = "feedback positive: server-side tracking"; Pattern = "server\s*[- ]?\s*side|server\s+container|\bsgtm\b" },
        [PSCustomObject]@{ Key = "rgpd"; Label = "feedback positive: RGPD/GDPR"; Pattern = "\brgpd\b|\bgdpr\b|protection\s+des\s+donn[eé]es|privacy|conformit[eé]" },
        [PSCustomObject]@{ Key = "datalayer"; Label = "feedback positive: dataLayer"; Pattern = "data\s*layer|datalayer" },
        [PSCustomObject]@{ Key = "tagging_plan"; Label = "feedback positive: tagging plan"; Pattern = "tagging\s+plan|tracking\s+plan|plan\s+de\s+(taggage|marquage)" },
        [PSCustomObject]@{ Key = "consent"; Label = "feedback positive: consent tracking"; Pattern = "consent\s+mode|cookie\s+consent|\bcmp\b" },
        [PSCustomObject]@{ Key = "cro"; Label = "feedback positive: CRO"; Pattern = "\bcro\b|conversion\s+rate|conversion\s+optimization|optimisation\s+conversion" }
    )
}

function Get-HashtableIntValue {
    param(
        [AllowNull()]$Table,
        [string]$Key
    )

    if ($null -eq $Table -or [string]::IsNullOrWhiteSpace($Key)) {
        return 0
    }

    if ($Table -is [Collections.IDictionary] -and $Table.Contains($Key)) {
        return [int]$Table[$Key]
    }

    $property = $Table.PSObject.Properties[$Key]
    if ($null -ne $property -and $null -ne $property.Value) {
        return [int]$property.Value
    }

    return 0
}

function Add-FeedbackCount {
    param(
        [hashtable]$Table,
        [string]$Key
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return
    }

    if (-not $Table.ContainsKey($Key)) {
        $Table[$Key] = 0
    }
    $Table[$Key] = [int]$Table[$Key] + 1
}

function New-FeedbackLearningProfile {
    param([object[]]$Rows)

    $positiveCounts = @{}
    $ignoreReasonCounts = @{}
    $positiveRows = 0
    $ignoredRows = 0
    $signals = @(Get-FeedbackSignalDefinitions)

    foreach ($row in @($Rows)) {
        $status = ConvertTo-MatchText (Get-RowValue -Row $row -Name "status")
        if ([string]::IsNullOrWhiteSpace($status)) {
            continue
        }

        $rowText = Get-FeedbackProfileText $row
        if ($status -match "^(applied|interview|offer|interesting)$") {
            $positiveRows++
            foreach ($signal in $signals) {
                if ($rowText -match [string]$signal.Pattern) {
                    Add-FeedbackCount -Table $positiveCounts -Key ([string]$signal.Key)
                }
            }
        }
        elseif ($status -eq "ignored") {
            $ignoredRows++
            $ignoreReason = Get-IgnoreReasonFromNotes (Get-RowValue -Row $row -Name "notes")
            if (-not [string]::IsNullOrWhiteSpace($ignoreReason)) {
                Add-FeedbackCount -Table $ignoreReasonCounts -Key (ConvertTo-IgnoreReasonKey $ignoreReason)
            }
        }
    }

    return [PSCustomObject]@{
        PositiveSignalCounts = $positiveCounts
        IgnoreReasonCounts   = $ignoreReasonCounts
        PositiveRows         = $positiveRows
        IgnoredRows          = $ignoredRows
    }
}

function Get-FeedbackLearningAdjustment {
    param(
        [string]$FullText,
        [bool]$HasCoreTitleSignal,
        [bool]$HasWebAnalyticsToolSignal,
        [bool]$HasDigitalAnalyticsContext
    )

    $profile = $script:FeedbackLearningProfile
    if ($null -eq $profile) {
        return [PSCustomObject]@{ Adjustment = 0; Reasons = @() }
    }

    $adjustment = 0
    $reasons = New-Object System.Collections.Generic.List[string]
    $positiveSignals = $profile.PositiveSignalCounts
    foreach ($signal in @(Get-FeedbackSignalDefinitions)) {
        $count = Get-HashtableIntValue -Table $positiveSignals -Key ([string]$signal.Key)
        if ($count -le 0 -or $FullText -notmatch [string]$signal.Pattern) {
            continue
        }

        $delta = [Math]::Min(8, 2 + (2 * $count))
        $adjustment += $delta
        $reasons.Add([string]$signal.Label) | Out-Null
    }

    if ($adjustment -gt 18) {
        $adjustment = 18
    }

    $ignoreCounts = $profile.IgnoreReasonCounts
    $negativeAdjustment = 0
    $negativeRules = @(
        [PSCustomObject]@{ Key = "too_seo_sea_marketing"; Pattern = "\bseo\b|\bsea\b|paid\s+social|paid\s+search|paid\s+media|performance\s+marketing|growth\s+marketing|acquisition|campaign|media\s+buyer"; Label = "feedback ignored: SEO/SEA/marketing"; Max = 14 },
        [PSCustomObject]@{ Key = "too_data_analyst"; Pattern = "\bdata\s*analyst\b|analyste\s+de\s+donnees|\bpython\b|\bsql\b|notebook|data\s+warehouse|business\s+analyst"; Label = "feedback ignored: data analyst"; Max = 12 },
        [PSCustomObject]@{ Key = "too_data_engineering"; Pattern = "data\s+engineer|analytics?\s+engineer|\bdbt\b|snowflake|airflow|\betl\b|\belt\b|data\s+warehouse|datawarehouse|data\s+platform|databricks|pyspark|spark|pipeline|backend|devops"; Label = "feedback ignored: data engineering"; Max = 16 },
        [PSCustomObject]@{ Key = "too_bi_reporting"; Pattern = "\bbi\b|business\s+intelligence|power\s*bi|tableau|dashboard|reporting|looker|data\s+studio|tableau\s+de\s+bord"; Label = "feedback ignored: BI/reporting"; Max = 10 },
        [PSCustomObject]@{ Key = "too_crm_emailing"; Pattern = "\bcrm\b|emailing|email\s+marketing|marketing\s+automation|salesforce|hubspot|braze|batch|campaign"; Label = "feedback ignored: CRM/emailing"; Max = 12 },
        [PSCustomObject]@{ Key = "too_content_social"; Pattern = "content\s+marketing|social\s+media|community\s+manager|editorial|copywriting|seo\s+content"; Label = "feedback ignored: content/social"; Max = 12 },
        [PSCustomObject]@{ Key = "too_product_analytics"; Pattern = "product\s+analyst|product\s+analytics|amplitude|mixpanel|heap"; Label = "feedback ignored: product analytics"; Max = 8 },
        [PSCustomObject]@{ Key = "too_managerial"; Pattern = "\bhead\b|director|directeur|directrice|lead|manager|responsable|principal"; Label = "feedback ignored: managerial"; Max = 8 },
        [PSCustomObject]@{ Key = "agency_consulting_esn"; Pattern = "consultant|consulting|cabinet|agence|agency|\besn\b|ssii"; Label = "feedback ignored: agency/consulting/ESN"; Max = 8 }
    )

    foreach ($rule in $negativeRules) {
        $count = Get-HashtableIntValue -Table $ignoreCounts -Key ([string]$rule.Key)
        if ($count -le 0 -or $FullText -notmatch [string]$rule.Pattern) {
            continue
        }

        if ($rule.Key -match "too_data|too_bi|too_product" -and $HasWebAnalyticsToolSignal) {
            continue
        }
        if ($rule.Key -eq "too_seo_sea_marketing" -and $HasDigitalAnalyticsContext) {
            continue
        }

        $delta = [Math]::Min([int]$rule.Max, 4 + (3 * $count))
        $negativeAdjustment -= $delta
        $reasons.Add([string]$rule.Label) | Out-Null
    }

    $notAnalyticsCount = Get-HashtableIntValue -Table $ignoreCounts -Key "not_analytics_enough"
    if ($notAnalyticsCount -gt 0 -and -not $HasCoreTitleSignal -and -not $HasWebAnalyticsToolSignal) {
        $negativeAdjustment -= [Math]::Min(12, 4 + (3 * $notAnalyticsCount))
        $reasons.Add("feedback ignored: not analytics enough") | Out-Null
    }

    if ($negativeAdjustment -lt -25) {
        $negativeAdjustment = -25
    }

    return [PSCustomObject]@{
        Adjustment = [int]($adjustment + $negativeAdjustment)
        Reasons    = @($reasons.ToArray() | Select-Object -Unique)
    }
}

function Get-IgnoredFeedbackPenalty {
    param(
        $Row,
        $ExistingRow,
        [AllowNull()][string]$IgnoreReason,
        [bool]$SameCompany,
        [bool]$SameTitle,
        [bool]$KeywordOverlap
    )

    $reason = ConvertTo-IgnoreReasonKey $IgnoreReason
    $rowText = Get-FeedbackProfileText $Row
    $existingText = Get-FeedbackProfileText $ExistingRow
    $hasWebAnalyticsSignal = Test-FeedbackTextHasWebAnalyticsSignal $rowText

    if ($reason -eq "duplicate") {
        return [PSCustomObject]@{ Penalty = 0; Reason = "" }
    }

    switch ($reason) {
        "not_analytics_enough" {
            if (-not $hasWebAnalyticsSignal -or $rowText -match "possible marketing|possible broad analyst|possible data analyst") {
                return [PSCustomObject]@{ Penalty = 22; Reason = "ignored reason: not analytics enough" }
            }
        }
        "too_seo_sea_marketing" {
            if ($rowText -match "\bseo\b|\bsea\b|paid\s+social|paid\s+search|paid\s+media|performance\s+marketing|growth\s+marketing|acquisition|digital\s+marketing|campaign|media\s+buyer") {
                return [PSCustomObject]@{ Penalty = 26; Reason = "ignored reason: SEO/SEA/marketing" }
            }
        }
        "too_data_analyst" {
            if (($rowText -match "\bdata\s*analyst\b|analyste\s+de\s+donnees|\bpython\b|\bsql\b|notebook|data\s+warehouse|possible broad analyst") -and -not $hasWebAnalyticsSignal) {
                return [PSCustomObject]@{ Penalty = 24; Reason = "ignored reason: data analyst" }
            }
        }
        "too_data_engineering" {
            if ($rowText -match "data\s+engineer|analytics?\s+engineer|\bdbt\b|snowflake|airflow|\betl\b|\belt\b|data\s+warehouse|datawarehouse|data\s+platform|databricks|pyspark|spark|pipeline|backend|devops|possible engineering") {
                return [PSCustomObject]@{ Penalty = 28; Reason = "ignored reason: data engineering" }
            }
        }
        "too_bi_reporting" {
            if (($rowText -match "\bbi\b|business\s+intelligence|power\s*bi|tableau|dashboard|reporting|looker|data\s+studio|tableau\s+de\s+bord") -and -not $hasWebAnalyticsSignal) {
                return [PSCustomObject]@{ Penalty = 20; Reason = "ignored reason: BI/reporting" }
            }
        }
        "too_crm_emailing" {
            if ($rowText -match "\bcrm\b|emailing|email\s+marketing|marketing\s+automation|salesforce|hubspot|braze|batch|campaign") {
                return [PSCustomObject]@{ Penalty = 22; Reason = "ignored reason: CRM/emailing" }
            }
        }
        "too_content_social" {
            if ($rowText -match "content\s+marketing|social\s+media|community\s+manager|editorial|copywriting|seo\s+content") {
                return [PSCustomObject]@{ Penalty = 22; Reason = "ignored reason: content/social" }
            }
        }
        "too_product_analytics" {
            if (($rowText -match "product\s+analyst|product\s+analytics|amplitude|mixpanel|heap") -and -not $hasWebAnalyticsSignal) {
                return [PSCustomObject]@{ Penalty = 16; Reason = "ignored reason: product analytics" }
            }
        }
        "too_managerial" {
            $rowBucket = Get-FeedbackSeniorityBucket $rowText
            if ($rowBucket -eq "management") {
                return [PSCustomObject]@{ Penalty = 16; Reason = "ignored reason: too managerial" }
            }
        }
        "agency_consulting_esn" {
            if ((Test-FeedbackRowHasAgencyConsultingEsnSignal $Row) -or $SameCompany -or $SameTitle) {
                return [PSCustomObject]@{ Penalty = 18; Reason = "ignored reason: agency/consulting/ESN preference" }
            }
        }
        "wrong_seniority" {
            $rowBucket = Get-FeedbackSeniorityBucket $rowText
            $existingBucket = Get-FeedbackSeniorityBucket $existingText
            if (-not [string]::IsNullOrWhiteSpace($rowBucket) -and $rowBucket -eq $existingBucket) {
                return [PSCustomObject]@{ Penalty = 14; Reason = "ignored reason: seniority" }
            }
        }
        "wrong_location" {
            $rowLocation = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "location")
            $existingLocation = ConvertTo-IdentityText -Text (Get-RowValue -Row $ExistingRow -Name "location")
            if (-not [string]::IsNullOrWhiteSpace($rowLocation) -and $rowLocation -eq $existingLocation) {
                return [PSCustomObject]@{ Penalty = 12; Reason = "ignored reason: location" }
            }
        }
        "wrong_remote_policy" {
            if ($rowText -match "on\s*site|onsite|hybrid|remote|teletravail") {
                return [PSCustomObject]@{ Penalty = 8; Reason = "ignored reason: remote policy" }
            }
        }
        "wrong_contract" {
            $rowContract = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "contract_type")
            $existingContract = ConvertTo-IdentityText -Text (Get-RowValue -Row $ExistingRow -Name "contract_type")
            if (-not [string]::IsNullOrWhiteSpace($rowContract) -and $rowContract -eq $existingContract) {
                return [PSCustomObject]@{ Penalty = 10; Reason = "ignored reason: contract" }
            }
        }
        "language_issue" {
            if ($rowText -match "english|french|francais|bilingual|native|fluent") {
                return [PSCustomObject]@{ Penalty = 8; Reason = "ignored reason: language" }
            }
        }
        "salary_issue" {
            return [PSCustomObject]@{ Penalty = 0; Reason = "" }
        }
        "company_not_interested" {
            if ($SameCompany) {
                return [PSCustomObject]@{ Penalty = 28; Reason = "ignored reason: company" }
            }
        }
        "industry_not_interested" {
            if ($SameCompany) {
                return [PSCustomObject]@{ Penalty = 12; Reason = "ignored reason: industry/company proxy" }
            }
        }
        "low_quality_posting" {
            if ($SameCompany -or $SameTitle) {
                return [PSCustomObject]@{ Penalty = 10; Reason = "ignored reason: low-quality posting" }
            }
        }
        "other" {
            if ($SameCompany -or $SameTitle) {
                return [PSCustomObject]@{ Penalty = 8; Reason = "ignored reason: other" }
            }
        }
    }

    if ($SameCompany -or $SameTitle -or $KeywordOverlap) {
        return [PSCustomObject]@{ Penalty = 12; Reason = $(if ([string]::IsNullOrWhiteSpace($reason)) { "similar ignored job without reason" } else { "similar ignored job" }) }
    }

    return [PSCustomObject]@{ Penalty = 0; Reason = "" }
}

function Get-FeedbackAdjustment {
    param(
        $Row,
        [object[]]$ExistingRows
    )

    $titleText = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "job_title") -Title
    $companyText = ConvertTo-IdentityText -Text (Get-RowValue -Row $Row -Name "company_name")
    $keywordText = ConvertTo-MatchText (Get-RowValue -Row $Row -Name "matched_keywords")
    $adjustment = 0
    $reasons = New-Object System.Collections.Generic.List[string]

    foreach ($existing in @($ExistingRows)) {
        $status = ConvertTo-MatchText (Get-RowValue -Row $existing -Name "status")
        if ([string]::IsNullOrWhiteSpace($status)) {
            continue
        }

        $existingTitle = ConvertTo-IdentityText -Text (Get-RowValue -Row $existing -Name "job_title") -Title
        $existingCompany = ConvertTo-IdentityText -Text (Get-RowValue -Row $existing -Name "company_name")
        $existingKeywords = ConvertTo-MatchText (Get-RowValue -Row $existing -Name "matched_keywords")
        $sameCompany = -not [string]::IsNullOrWhiteSpace($companyText) -and $companyText -eq $existingCompany
        $sameTitle = -not [string]::IsNullOrWhiteSpace($titleText) -and $titleText -eq $existingTitle
        $keywordOverlap = -not [string]::IsNullOrWhiteSpace($keywordText) -and -not [string]::IsNullOrWhiteSpace($existingKeywords) -and ($keywordText -match "google|gtm|ga4|piano|contentsquare|tag\s+commander|commanders?\s+act|tealium|server-side|server\s+side|rgpd|gdpr|tracking|tagging|cro") -and ($existingKeywords -match "google|gtm|ga4|piano|contentsquare|tag\s+commander|commanders?\s+act|tealium|server-side|server\s+side|rgpd|gdpr|tracking|tagging|cro")

        if ($status -match "^(applied|interview|offer|interesting)$" -and ($sameCompany -or $sameTitle -or $keywordOverlap)) {
            $adjustment += 10
            $reasons.Add("positive history") | Out-Null
        }
        elseif ($status -eq "ignored") {
            $ignoreReason = Get-IgnoreReasonFromNotes (Get-RowValue -Row $existing -Name "notes")
            $ignoredFeedback = Get-IgnoredFeedbackPenalty -Row $Row -ExistingRow $existing -IgnoreReason $ignoreReason -SameCompany:$sameCompany -SameTitle:$sameTitle -KeywordOverlap:$keywordOverlap
            if ([int]$ignoredFeedback.Penalty -gt 0) {
                $adjustment -= [int]$ignoredFeedback.Penalty
                $reasons.Add([string]$ignoredFeedback.Reason) | Out-Null
            }
        }
    }

    if ($adjustment -gt 30) { $adjustment = 30 }
    if ($adjustment -lt -40) { $adjustment = -40 }

    return [PSCustomObject]@{
        Adjustment = $adjustment
        Reason = (($reasons.ToArray() | Select-Object -Unique) -join "; ")
    }
}

function Apply-FeedbackScoring {
    param(
        [object[]]$Rows,
        [object[]]$ExistingRows
    )

    foreach ($row in @($Rows)) {
        $feedback = Get-FeedbackAdjustment -Row $row -ExistingRows $ExistingRows
        $oldScore = 0
        try { $oldScore = [int](Get-RowValue -Row $row -Name "match_score") } catch { $oldScore = 0 }
        $newScore = [Math]::Max(0, $oldScore + [int]$feedback.Adjustment)
        $row.match_score = [string]$newScore
        $row.feedback_adjustment = [string]$feedback.Adjustment
        if (-not [string]::IsNullOrWhiteSpace($feedback.Reason)) {
            $row.matched_keywords = (Join-CleanTextParts @((Get-RowValue -Row $row -Name "matched_keywords"), ("feedback: " + $feedback.Reason))) -replace ", ", "; "
        }

        if ($newScore -ge 80) {
            $row.match_level = "High"
        }
        elseif ($newScore -ge 50) {
            $row.match_level = "Medium"
        }
        else {
            $row.match_level = "Review"
        }
    }

    return @($Rows)
}

function Merge-JobsWithTracker {
    param(
        [object[]]$CurrentRows,
        [object[]]$ExistingRows,
        [string]$Path
    )

    $existingRecords = New-Object System.Collections.Generic.List[object]
    foreach ($existing in @($ExistingRows)) {
        $record = ConvertTo-TrackerRecordFromExisting $existing
        if ($null -ne $record) {
            $existingRecords.Add($record) | Out-Null
        }
    }

    $duplicateCount = 0
    $existingByKey = @{}
    $existingGroups = Group-RowsByDedupeKey -Rows @($existingRecords.ToArray())
    foreach ($key in $existingGroups.Keys) {
        $groupRows = @($existingGroups[$key].ToArray())
        if ($groupRows.Count -gt 1) {
            $duplicateCount += ($groupRows.Count - 1)
        }

        $reason = ""
        if ($groupRows.Count -gt 1) {
            $reason = "merged similar tracker rows"
        }
        $existingByKey[$key] = Merge-SimilarJobRows -Rows $groupRows -Reason $reason
    }

    $currentByKey = @{}
    $currentGroups = Group-RowsByDedupeKey -Rows @($CurrentRows)
    foreach ($key in $currentGroups.Keys) {
        $groupRows = @($currentGroups[$key].ToArray())
        if ($groupRows.Count -gt 1) {
            $duplicateCount += ($groupRows.Count - 1)
        }

        $reasonParts = New-Object System.Collections.Generic.List[string]
        if ($groupRows.Count -gt 1) {
            $reasonParts.Add(("merged {0} similar current postings" -f $groupRows.Count)) | Out-Null
        }

        $platforms = @(Get-UniqueTextValues -Values @($groupRows | ForEach-Object { Get-RowValue -Row $_ -Name "platform" }))
        if ($platforms.Count -gt 1) {
            $reasonParts.Add(("same job found on multiple sources: {0}" -f ($platforms -join "; "))) | Out-Null
        }

        $currentByKey[$key] = Merge-SimilarJobRows -Rows $groupRows -Reason (Join-CleanTextParts $reasonParts.ToArray())
    }

    $trackerByKey = @{}
    foreach ($key in $currentByKey.Keys) {
        $existing = $null
        if ($existingByKey.ContainsKey($key)) {
            $existing = $existingByKey[$key]
        }

        $duplicateReason = $(if ($null -ne $existing) { "same normalized company/title/location from previous crawl" } else { "" })
        $currentDuplicateReason = Get-RowValue -Row $currentByKey[$key] -Name "duplicate_reason"
        if (-not [string]::IsNullOrWhiteSpace($currentDuplicateReason)) {
            $duplicateReason = Join-CleanTextParts @($duplicateReason, $currentDuplicateReason)
        }
        $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $currentByKey[$key] -ExistingRow $existing -SeenInCurrentCrawl:$true -DuplicateReason $duplicateReason
    }

    $removedCount = 0
    $preservedAppliedCount = 0
    foreach ($key in $existingByKey.Keys) {
        if ($trackerByKey.ContainsKey($key)) {
            continue
        }

        $existing = $existingByKey[$key]
        if (Test-IsKeepForeverStatus (Get-RowValue -Row $existing -Name "status")) {
            $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $existing -ExistingRow $existing -SeenInCurrentCrawl:$false -DuplicateReason "kept by application status"
            $preservedAppliedCount++
        }
        elseif (Test-IsExcludedContractType (Get-RowValue -Row $existing -Name "contract_type")) {
            $removedCount++
        }
        elseif (Test-IsRecentTrackerRow $existing) {
            $trackerByKey[$key] = ConvertTo-TrackerRecord -CurrentRow $existing -ExistingRow $existing -SeenInCurrentCrawl:$false -DuplicateReason "not seen this crawl, still inside published-date retention window"
        }
        else {
            $removedCount++
        }
    }

    $trackerRows = @($trackerByKey.Values) |
        Sort-Object -Property `
            @{ Expression = { if ((Get-RowValue -Row $_ -Name "seen_in_current_crawl") -eq "yes") { 1 } else { 0 } }; Descending = $true },
            @{ Expression = { try { [int](Get-RowValue -Row $_ -Name "match_score") } catch { 0 } }; Descending = $true },
            @{ Expression = "published_date"; Descending = $true },
            platform,
            job_title

    $backupPath = Backup-TrackerFile -Path $Path

    return @{
        TrackerRows = @($trackerRows)
        CurrentRows = @($currentByKey.Values)
        RemovedCount = $removedCount
        DuplicateCount = $duplicateCount
        PreservedAppliedCount = $preservedAppliedCount
        BackupPath = $backupPath
    }
}

function Invoke-TextRequest {
    param(
        [string]$Url,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30
    )

    $mergedHeaders = @{
        "User-Agent"      = $BrowserUserAgent
        "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        "Accept-Language" = "fr-FR,fr;q=0.9,en;q=0.8"
    }

    foreach ($key in $Headers.Keys) {
        $mergedHeaders[$key] = $Headers[$key]
    }

    $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -Headers $mergedHeaders -TimeoutSec $TimeoutSec
    return [string]$response.Content
}

function Invoke-CachedTextRequest {
    param(
        [string]$Url,
        [string]$CacheScope,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30,
        [AllowNull()]$Stats = $null
    )

    $cached = Get-CachedText -Scope $CacheScope -Key $Url
    if ($null -ne $cached) {
        Add-SourceMetric -Stats $Stats -Name "CacheHits"
        return $cached
    }

    Add-SourceMetric -Stats $Stats -Name "DetailRequests"
    $text = Invoke-TextRequest -Url $Url -Headers $Headers -TimeoutSec $TimeoutSec
    Set-CachedText -Scope $CacheScope -Key $Url -Text $text
    return $text
}

function Invoke-JsonPostRequest {
    param(
        [string]$Url,
        [AllowNull()]$Body = $null,
        [hashtable]$Headers = @{},
        [int]$TimeoutSec = 30
    )

    $mergedHeaders = @{
        "User-Agent"      = $BrowserUserAgent
        "Accept"          = "application/json, text/plain, */*"
        "Accept-Language" = "fr-FR,fr;q=0.9,en;q=0.8"
        "Content-Type"    = "application/json"
    }

    foreach ($key in $Headers.Keys) {
        $mergedHeaders[$key] = $Headers[$key]
    }

    $jsonBody = ""
    if ($null -ne $Body) {
        $jsonBody = $Body | ConvertTo-Json -Depth 12 -Compress
    }

    $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Post -Headers $mergedHeaders -Body $jsonBody -TimeoutSec $TimeoutSec
    if ([string]::IsNullOrWhiteSpace([string]$response.Content)) {
        return $null
    }

    return ([string]$response.Content | ConvertFrom-Json)
}

function Invoke-CurlTextRequest {
    param([string]$Url)

    $curl = Get-Command "curl.exe" -ErrorAction Stop
    $lines = & $curl.Source -L -s --compressed $Url `
        -H "User-Agent: $BrowserUserAgent" `
        -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" `
        -H "Accept-Language: fr-FR,fr;q=0.9,en;q=0.8" `
        -H "Connection: keep-alive"

    if ($LASTEXITCODE -ne 0) {
        throw "curl.exe failed for $Url with exit code $LASTEXITCODE"
    }

    return ($lines -join "`n")
}

function Get-MetaContent {
    param(
        [string]$Html,
        [string]$Name
    )

    $escapedName = [regex]::Escape($Name)
    $pattern = "<meta[^>]+(?:property|name)=[""']$escapedName[""'][^>]+content=[""'](?<content>[^""']+)[""']"
    $match = [regex]::Match($Html, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
        $pattern = "<meta[^>]+content=[""'](?<content>[^""']+)[""'][^>]+(?:property|name)=[""']$escapedName[""']"
        $match = [regex]::Match($Html, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }

    if ($match.Success) {
        return ConvertFrom-HtmlAttribute $match.Groups["content"].Value
    }

    return ""
}

function Get-TitleFromHtml {
    param([string]$Html)

    $ogTitle = Get-MetaContent -Html $Html -Name "og:title"
    if (-not [string]::IsNullOrWhiteSpace($ogTitle)) {
        return ($ogTitle -replace "\s+-\s+Welcome to the Jungle.*$", "").Trim()
    }

    $titleMatch = [regex]::Match($Html, "(?is)<title[^>]*>(?<title>.*?)</title>")
    if ($titleMatch.Success) {
        return (ConvertFrom-HtmlText $titleMatch.Groups["title"].Value)
    }

    return ""
}

function Get-TitleFromWttjUrl {
    param([string]$Url)

    $slugMatch = [regex]::Match($Url, "/jobs/(?<slug>[^/?#]+)")
    if (-not $slugMatch.Success) {
        return $Url
    }

    $slug = $slugMatch.Groups["slug"].Value
    $slug = ($slug -split "_")[0]
    $slug = $slug -replace "-", " "
    return Repair-DisplayText ([Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($slug))
}

function ConvertFrom-SlugToTitle {
    param([AllowNull()][string]$Slug)

    if ([string]::IsNullOrWhiteSpace($Slug)) {
        return ""
    }

    $clean = $Slug -replace "[-_]+", " "
    $clean = ([regex]::Replace($clean, "\s+", " ")).Trim()
    return Repair-DisplayText ([Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($clean))
}

function Join-CleanTextParts {
    param([object[]]$Parts)

    $clean = New-Object System.Collections.Generic.List[string]
    foreach ($part in @($Parts)) {
        if ($null -eq $part) {
            continue
        }

        $text = ConvertFrom-HtmlText ([string]$part)
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        if (-not $clean.Contains($text)) {
            $clean.Add($text) | Out-Null
        }
    }

    return ($clean.ToArray() -join ", ")
}

function Get-ObjectPropertyValue {
    param(
        [AllowNull()]$Object,
        [string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    if ($Object -is [Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($Object.Contains($name)) {
                return $Object[$name]
            }
        }

        return $null
    }

    $properties = @($Object.PSObject.Properties)
    foreach ($name in $Names) {
        $property = @($properties | Where-Object { $_.Name -eq $name } | Select-Object -First 1)
        if ($property.Count -gt 0) {
            return $property[0].Value
        }
    }

    return $null
}

function ConvertTo-LocationText {
    param(
        [AllowNull()]$Value,
        [int]$Depth = 0
    )

    if ($null -eq $Value -or $Depth -gt 3) {
        return ""
    }

    if ($Value -is [string]) {
        return ConvertFrom-HtmlText $Value
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $locations = foreach ($item in @($Value)) {
            ConvertTo-LocationText -Value $item -Depth ($Depth + 1)
        }
        return Join-CleanTextParts $locations
    }

    $city = Get-ObjectPropertyValue -Object $Value -Names @("city", "locality", "addressLocality", "town")
    $region = Get-ObjectPropertyValue -Object $Value -Names @("region", "state", "addressRegion", "department")
    $country = Get-ObjectPropertyValue -Object $Value -Names @("country", "countryCode", "addressCountry")
    $directLocation = Join-CleanTextParts @($city, $region, $country)
    if (-not [string]::IsNullOrWhiteSpace($directLocation)) {
        return $directLocation
    }

    foreach ($nestedName in @("location", "locations", "address", "addresses", "office", "offices", "place", "places")) {
        $nestedValue = Get-ObjectPropertyValue -Object $Value -Names @($nestedName)
        $nestedLocation = ConvertTo-LocationText -Value $nestedValue -Depth ($Depth + 1)
        if (-not [string]::IsNullOrWhiteSpace($nestedLocation)) {
            return $nestedLocation
        }
    }

    $name = Get-ObjectPropertyValue -Object $Value -Names @("name", "label", "formatted", "full_address", "fullAddress")
    return ConvertTo-LocationText -Value $name -Depth ($Depth + 1)
}

function Get-LocationFromStructuredHtml {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $cityMatch = [regex]::Match($Html, '(?i)"addressLocality"\s*:\s*"(?<value>[^"]+)"')
    $regionMatch = [regex]::Match($Html, '(?i)"addressRegion"\s*:\s*"(?<value>[^"]+)"')
    $countryMatch = [regex]::Match($Html, '(?i)"addressCountry"\s*:\s*"(?<value>[^"]+)"')

    return Join-CleanTextParts @(
        $(if ($cityMatch.Success) { $cityMatch.Groups["value"].Value }),
        $(if ($regionMatch.Success) { $regionMatch.Groups["value"].Value }),
        $(if ($countryMatch.Success) { $countryMatch.Groups["value"].Value })
    )
}

function Get-LocationFromText {
    param([AllowNull()][string]$Text)

    $clean = ConvertFrom-HtmlText $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ""
    }

    $match = [regex]::Match($clean, "(?i)(?:\ba\s+|\b\xE0\s+|\bin\s+)(?<location>\p{L}[\p{L}\p{M}' -]{2,})(?:$|[,.])")
    if ($match.Success) {
        $location = $match.Groups["location"].Value.Trim()
        if (-not (Test-IsJunkLocationText $location)) {
            return $location
        }
    }

    return ""
}

function Test-IsJunkLocationText {
    param([AllowNull()][string]$Location)

    if ([string]::IsNullOrWhiteSpace($Location)) {
        return $true
    }

    $raw = ([string]$Location).Trim()
    $clean = ConvertTo-MatchText $raw
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $true
    }
    if ($clean -match "^(h|f|m|x|nb|stage|internship|cdi|cdd|full\s*time|permanent)$") {
        return $true
    }
    if ($raw -match "^[A-Za-z0-9]{5,14}$" -and $raw -match "\d" -and $raw -match "[A-Z]" -and $raw -match "[a-z]") {
        return $true
    }
    if ($clean -match "^[a-z0-9]{8,}$" -and $clean -match "\d") {
        return $true
    }

    return $false
}

function Get-WttjLocationFromUrl {
    param([AllowNull()][string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ""
    }

    $path = $Url
    try {
        $path = ([Uri]$Url).AbsolutePath
    }
    catch {
    }

    $locationMatch = [regex]::Match($path, "/jobs/[^/?#]*_(?<location>[^/_?#]+)$", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $locationMatch.Success) {
        return ""
    }

    $locationSlug = $locationMatch.Groups["location"].Value
    if ($locationSlug -match "^(h|f|m|x|nb|stage|internship|cdi|cdd)$") {
        return ""
    }

    $location = ConvertFrom-SlugToTitle $locationSlug
    if (Test-IsJunkLocationText $location) {
        return ""
    }

    return $location
}

function Get-WttjLocation {
    param(
        [AllowNull()][string]$Html,
        [AllowNull()][string]$Url,
        [AllowNull()][string]$Title
    )

    $location = Get-LocationFromStructuredHtml $Html
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        return $location
    }

    $location = Get-LocationFromText $Title
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        return $location
    }

    return Get-WttjLocationFromUrl $Url
}

function Get-WelcomeKitLocation {
    param(
        [AllowNull()]$Job,
        [AllowNull()][string]$JobUrl
    )

    foreach ($propertyName in @("location", "locations", "office", "offices", "address", "addresses", "workplace", "workplaces")) {
        $location = ConvertTo-LocationText (Get-ObjectPropertyValue -Object $Job -Names @($propertyName))
        if (-not [string]::IsNullOrWhiteSpace($location)) {
            return $location
        }
    }

    return Get-WttjLocationFromUrl $JobUrl
}

function Get-LinkedInLocationFromHtml {
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    foreach ($pattern in @(
        '(?is)<span[^>]*class="[^"]*job-search-card__location[^"]*"[^>]*>(?<location>.*?)</span>',
        '(?is)<span[^>]*class="[^"]*topcard__flavor[^"]*topcard__flavor--bullet[^"]*"[^>]*>(?<location>.*?)</span>',
        '(?is)<span[^>]*class="[^"]*jobs-unified-top-card__bullet[^"]*"[^>]*>(?<location>.*?)</span>'
    )) {
        $match = [regex]::Match($Html, $pattern)
        if ($match.Success) {
            $location = ConvertFrom-HtmlText $match.Groups["location"].Value
            if (-not [string]::IsNullOrWhiteSpace($location)) {
                return $location
            }
        }
    }

    return Get-LocationFromStructuredHtml $Html
}

function Get-WttjCompanyNameFromUrl {
    param([string]$Url)

    $companyMatch = [regex]::Match($Url, "/companies/(?<slug>[^/]+)/jobs/", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($companyMatch.Success) {
        return ConvertFrom-SlugToTitle $companyMatch.Groups["slug"].Value
    }

    return ""
}

function Get-WttjCandidateScore {
    param([string]$Url)

    $score = 0
    if ($Url -match "(?i)(web[-_\s]*analyst|digital[-_\s]*analyst|web[-_\s]*analytics|digital[-_\s]*analytics|tracking|taggage|tagging|(^|[-_\s])ga4($|[-_\s])|(^|[-_\s])gtm($|[-_\s])|google[-_\s]*(analytics|tag[-_\s]*manager)|piano|contentsquare|content[-_\s]*square|(^|[-_\s])cro($|[-_\s]))") {
        $score += 100
    }
    if ($Url -match "(?i)(data[-_\s]*analyst|analytics[-_\s]*(consultant|specialist|manager|engineer))") {
        $score += 40
    }
    if ($Location -match "(?i)france|paris|french|remote" -and $Url -match "(?i)(/fr/|_paris|_puteaux|_levallois|_boulogne|_lille|_lyon|_fr\b)") {
        $score += 75
    }
    elseif ($Url -match "/fr/") {
        $score += 10
    }

    return $score
}

function Get-WelcomeKitCompanyName {
    param($Job, [string]$JobUrl)

    if ($Job.PSObject.Properties.Name -contains "organization" -and $null -ne $Job.organization) {
        if ($Job.organization.PSObject.Properties.Name -contains "name" -and -not [string]::IsNullOrWhiteSpace($Job.organization.name)) {
            return [string]$Job.organization.name
        }
        if ($Job.organization.PSObject.Properties.Name -contains "slug" -and -not [string]::IsNullOrWhiteSpace($Job.organization.slug)) {
            return ConvertFrom-SlugToTitle $Job.organization.slug
        }
    }

    return Get-WttjCompanyNameFromUrl $JobUrl
}

function Get-WelcomeKitJobUrl {
    param($Job)

    if ($Job.PSObject.Properties.Name -contains "websites" -and $null -ne $Job.websites) {
        foreach ($site in @($Job.websites)) {
            if ($site.PSObject.Properties.Name -contains "url" -and $site.url -match "welcometothejungle") {
                return [string]$site.url
            }
        }

        foreach ($site in @($Job.websites)) {
            if ($site.PSObject.Properties.Name -contains "url" -and -not [string]::IsNullOrWhiteSpace($site.url)) {
                return [string]$site.url
            }
        }
    }

    if ($Job.PSObject.Properties.Name -contains "apply_url" -and -not [string]::IsNullOrWhiteSpace($Job.apply_url)) {
        return [string]$Job.apply_url
    }

    if ($Job.PSObject.Properties.Name -contains "reference" -and -not [string]::IsNullOrWhiteSpace($Job.reference)) {
        return "https://www.welcometothejungle.com/fr/jobs/{0}" -f $Job.reference
    }

    return ""
}

function Get-WelcomeKitJobDetails {
    param(
        [string]$Reference,
        [hashtable]$Headers
    )

    if ([string]::IsNullOrWhiteSpace($Reference)) {
        return $null
    }

    $params = @{
        websites     = "true"
        organization = "true"
    }
    $url = "https://www.welcomekit.co/api/v1/external/jobs/{0}?{1}" -f [Uri]::EscapeDataString($Reference), (ConvertTo-QueryString $params)

    try {
        return Invoke-RestMethod -Uri $url -Headers $Headers -Method Get -TimeoutSec 45
    }
    catch {
        return $null
    }
}

function Get-FranceTravailJobs {
    $accessToken = Get-FranceTravailAccessToken
    if ([string]::IsNullOrWhiteSpace($accessToken)) {
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - France Travail"
    Write-RunStatus "Collecting France Travail jobs through the official API..."
    $stats = Start-SourceStats "France Travail"
    $results = New-Object System.Collections.Generic.List[object]
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Accept"        = "application/json"
    }
    $searchUrl = "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search"
    $pageSize = 150
    $queryIndex = 0

    foreach ($query in $ApiSearchQueries) {
        $queryIndex++
        Write-RunStatus ("France Travail query {0}/{1}: {2}" -f $queryIndex, $ApiSearchQueries.Count, $query)
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
            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 45
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("France Travail search failed for '{0}' page {1}: {2}" -f $query, ($page + 1), $_.Exception.Message)
                break
            }

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

            Write-CountProgress -Activity ("France Travail query {0}/{1}" -f $queryIndex, $ApiSearchQueries.Count) -Current ($page + 1) -Total $MaxFranceTravailPages -Found $results.Count -Every 1
            if ($jobArray.Count -lt $pageSize) {
                break
            }
            Start-Sleep -Milliseconds 400
        }
    }

    Write-RunStatus ("France Travail complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-AdzunaJobs {
    if ([string]::IsNullOrWhiteSpace($AdzunaAppId) -or [string]::IsNullOrWhiteSpace($AdzunaAppKey)) {
        Write-RunStatus "Adzuna credentials not set; skipping Adzuna source. Set ADZUNA_APP_ID and ADZUNA_APP_KEY to enable it."
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - Adzuna"
    Write-RunStatus "Collecting Adzuna jobs through the official API..."
    $stats = Start-SourceStats "Adzuna"
    $results = New-Object System.Collections.Generic.List[object]
    $queryIndex = 0

    foreach ($query in $ApiSearchQueries) {
        $queryIndex++
        Write-RunStatus ("Adzuna query {0}/{1}: {2}" -f $queryIndex, $ApiSearchQueries.Count, $query)
        for ($page = 1; $page -le $MaxAdzunaPages; $page++) {
            $params = @{
                app_id           = $AdzunaAppId
                app_key          = $AdzunaAppKey
                results_per_page = "25"
                what             = $query
                where            = $Location
                max_days_old     = [string][Math]::Abs($DaysBack)
                sort_by          = "date"
                "content-type"   = "application/json"
            }
            $url = "https://api.adzuna.com/v1/api/jobs/fr/search/{0}?{1}" -f $page, (ConvertTo-QueryString $params)

            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ "Accept" = "application/json" } -TimeoutSec 45
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("Adzuna search failed for '{0}' page {1}: {2}" -f $query, $page, $_.Exception.Message)
                break
            }

            $jobArray = @()
            if ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "results") {
                $jobArray = @($response.results)
            }

            if ($jobArray.Count -eq 0) {
                break
            }

            foreach ($job in $jobArray) {
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $publishedAt = ConvertTo-DateTimeOffsetOrNull (Get-ObjectPropertyValue -Object $job -Names @("created"))
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $title = [string](Get-ObjectPropertyValue -Object $job -Names @("title"))
                $description = [string](Get-ObjectPropertyValue -Object $job -Names @("description"))
                $companyName = Get-AdzunaCompanyName $job
                $jobLocation = Get-AdzunaLocation $job
                $contractType = Get-AdzunaContractType $job
                $jobUrl = ConvertTo-CleanUrl ([string](Get-ObjectPropertyValue -Object $job -Names @("redirect_url", "adref")))
                $sourceText = Join-CleanTextParts @($title, $description, $contractType)
                if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text $sourceText -Reliable) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $match = Get-JobMatch -Title $title -Text $sourceText
                if (-not $match.IsMatch) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "Adzuna" -PublishedAt $publishedAt -SourceText $sourceText
                if ($null -ne $result) {
                    $results.Add($result) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }

            Write-CountProgress -Activity ("Adzuna query {0}/{1}" -f $queryIndex, $ApiSearchQueries.Count) -Current $page -Total $MaxAdzunaPages -Found $results.Count -Every 1
            Start-Sleep -Milliseconds $AdzunaDelayMilliseconds
        }
    }

    Write-RunStatus ("Adzuna complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-ApecContractType {
    param([AllowNull()]$Job)

    $rawType = [string](Get-ObjectPropertyValue -Object $Job -Names @("typeContrat", "idNomTypeContrat"))
    switch ($rawType) {
        "101888" { return "CDI" }
        "101887" { return "CDD" }
        "597171" { return "Internship" }
        "20053" { return "Apprenticeship" }
        "101930" { return "Interim" }
        "101889" { return "Interim" }
    }

    $contractText = Join-CleanTextParts @(
        (Get-ObjectPropertyValue -Object $Job -Names @("intitule", "title"))
        (Get-ObjectPropertyValue -Object $Job -Names @("texteOffre", "description"))
        $rawType
    )
    return Get-ContractTypeFromText -Text $contractText
}

function Get-ApecJobUrl {
    param([AllowNull()]$Job)

    $numeroOffre = [string](Get-ObjectPropertyValue -Object $Job -Names @("numeroOffre", "NumeroOffre"))
    if ([string]::IsNullOrWhiteSpace($numeroOffre)) {
        $id = [string](Get-ObjectPropertyValue -Object $Job -Names @("id", "Id"))
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $numeroOffre = "{0}W" -f $id
        }
    }

    if ([string]::IsNullOrWhiteSpace($numeroOffre)) {
        return ""
    }

    return "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/{0}" -f [Uri]::EscapeDataString($numeroOffre.Trim())
}

function New-ApecSearchBody {
    param(
        [string]$Query,
        [int]$Page,
        [int]$PageSize,
        [string]$SortType = "SCORE"
    )

    return [ordered]@{
        lieux                   = @()
        fonctions               = @()
        statutPoste             = @()
        typesContrat            = @()
        typesConvention         = @("143684", "143685", "143686", "143687", "143706")
        niveauxExperience       = @()
        idsEtablissement        = @()
        secteursActivite        = @()
        typesTeletravail        = @()
        idNomZonesDeplacement   = @()
        positionNumbersExcluded = @()
        typeClient              = "CADRE"
        sorts                   = @(@{ type = $SortType; direction = "DESCENDING" })
        pagination              = @{ range = $PageSize; startIndex = ($Page * $PageSize) }
        activeFiltre            = $true
        pointGeolocDeReference  = @{ distance = 0 }
        motsCles                = $Query
    }
}

function Get-ApecJobs {
    Set-RunWindowTitle "Analytics Job Crawler - APEC"
    Write-RunStatus "Collecting APEC jobs from the public search endpoint..."
    Write-RunStatus ("APEC plan: {0} query/queries, up to {1} page(s) each, no detail-page crawl." -f $ApiSearchQueries.Count, $MaxApecPages)

    $stats = Start-SourceStats "APEC"
    $results = New-Object System.Collections.Generic.List[object]
    $headers = @{
        "Accept"  = "application/json, text/plain, */*"
        "Origin"  = "https://www.apec.fr"
        "Referer" = "https://www.apec.fr/candidat/recherche-emploi.html/emploi"
    }
    $searchUrl = "https://www.apec.fr/cms/webservices/rechercheOffre"
    $pageSize = 20
    $queryIndex = 0

    foreach ($query in $ApiSearchQueries) {
        $queryIndex++
        Write-RunStatus ("APEC query {0}/{1}: {2}" -f $queryIndex, $ApiSearchQueries.Count, $query)

        for ($page = 0; $page -lt $MaxApecPages; $page++) {
            $body = New-ApecSearchBody -Query $query -Page $page -PageSize $pageSize -SortType "SCORE"
            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $response = Invoke-JsonPostRequest -Url $searchUrl -Body $body -Headers $headers -TimeoutSec 30
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("APEC search failed for '{0}' page {1}: {2}" -f $query, ($page + 1), $_.Exception.Message)
                break
            }

            $jobArray = @()
            if ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "resultats") {
                $jobArray = @($response.resultats)
            }
            elseif ($null -ne $response -and @($response.PSObject.Properties.Name) -contains "results") {
                $jobArray = @($response.results)
            }

            if ($jobArray.Count -eq 0) {
                break
            }

            foreach ($job in $jobArray) {
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $publishedAt = ConvertTo-DateTimeOffsetOrNull (Get-ObjectPropertyValue -Object $job -Names @("datePublication", "dateValidation", "published_at"))
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $title = Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $job -Names @("intitule", "title")))
                $companyName = Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $job -Names @("nomCommercial", "company", "companyName")))
                $jobLocation = Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $job -Names @("lieuTexte", "location")))
                $contractType = Get-ApecContractType $job
                $jobUrl = Get-ApecJobUrl $job
                $sourceText = Join-CleanTextParts @(
                    $title,
                    $companyName,
                    $jobLocation,
                    $contractType,
                    (Get-ObjectPropertyValue -Object $job -Names @("texteOffre", "description", "intituleSurbrillance"))
                )
                if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text $sourceText -Reliable) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $match = Get-JobMatch -Title $title -Text $sourceText
                if (-not $match.IsMatch) {
                    Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                    continue
                }

                $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "APEC" -PublishedAt $publishedAt -SourceText $sourceText
                if ($null -ne $result) {
                    $results.Add($result) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }

            Write-CountProgress -Activity ("APEC query {0}/{1}" -f $queryIndex, $ApiSearchQueries.Count) -Current ($page + 1) -Total $MaxApecPages -Found $results.Count -Every 1
            if ($jobArray.Count -lt $pageSize) {
                break
            }
            Start-Sleep -Milliseconds $ApecDelayMilliseconds
        }
    }

    Write-RunStatus ("APEC complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-HelloWorkSearchUrl {
    param(
        [string]$Query,
        [int]$Page
    )

    $params = @{
        k = $Query
    }
    if (-not [string]::IsNullOrWhiteSpace($Location) -and $Location -notmatch "(?i)^france$") {
        $params["l"] = $Location
    }
    if ($Page -gt 1) {
        $params["p"] = [string]$Page
    }

    return "https://www.hellowork.com/fr-fr/emploi/recherche.html?{0}" -f (ConvertTo-QueryString $params)
}

function Get-HelloWorkJsonObjects {
    param([AllowNull()][string]$Html)

    $objects = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $objects.ToArray()
    }

    $scripts = [regex]::Matches($Html, '(?is)<script[^>]*type=["'']application/ld\+json["''][^>]*>(?<json>.*?)</script>')
    foreach ($script in $scripts) {
        $jsonText = [System.Net.WebUtility]::HtmlDecode($script.Groups["json"].Value).Trim()
        if ([string]::IsNullOrWhiteSpace($jsonText)) {
            continue
        }

        try {
            $parsed = $jsonText | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            continue
        }

        if ($parsed -is [System.Collections.IEnumerable] -and $parsed -isnot [string] -and $parsed -isnot [pscustomobject]) {
            foreach ($item in @($parsed)) {
                if ($null -ne $item) {
                    $objects.Add($item) | Out-Null
                }
            }
        }
        else {
            $objects.Add($parsed) | Out-Null
        }
    }

    return $objects.ToArray()
}

function Get-HelloWorkJobMetadata {
    param([AllowNull()][string]$Html)

    $title = ""
    $company = ""
    $location = ""
    $contract = ""
    $description = ""
    $datePosted = $null
    $employmentType = ""

    foreach ($jsonObject in (Get-HelloWorkJsonObjects -Html $Html)) {
        $objectType = [string](Get-ObjectPropertyValue -Object $jsonObject -Names @("@type", "type"))
        if ($objectType -eq "JobPosting") {
            $title = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("title", "name")))) $title
            $description = Get-PreferredValue (ConvertFrom-HtmlText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("description")))) $description
            $datePostedValue = Get-ObjectPropertyValue -Object $jsonObject -Names @("datePosted")
            if ($null -eq $datePosted -and $null -ne $datePostedValue) {
                $datePosted = ConvertTo-DateTimeOffsetOrNull $datePostedValue
            }

            $organization = Get-ObjectPropertyValue -Object $jsonObject -Names @("hiringOrganization")
            if ($null -ne $organization) {
                $company = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $organization -Names @("name")))) $company
            }

            $jobLocationValue = Get-ObjectPropertyValue -Object $jsonObject -Names @("jobLocation")
            $location = Get-PreferredValue (ConvertTo-LocationText $jobLocationValue) $location
            $employmentType = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("employmentType")))) $employmentType
        }

        $title = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("JobTitle")))) $title
        $company = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("Company")))) $company
        $location = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("Localisation")))) $location
        $contract = Get-PreferredValue (Repair-DisplayText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("ContractType")))) $contract
        $description = Get-PreferredValue (ConvertFrom-HtmlText ([string](Get-ObjectPropertyValue -Object $jsonObject -Names @("Description")))) $description
    }

    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = Get-TitleFromHtml $Html
    }

    [PSCustomObject]@{
        Title          = $title
        Company        = $company
        Location       = $location
        Contract       = $contract
        Description    = $description
        DatePosted     = $datePosted
        EmploymentType = $employmentType
    }
}

function Get-HelloWorkCardCandidates {
    param(
        [string]$Html,
        [string]$SearchUrl,
        [string]$Query,
        [AllowNull()]$Stats = $null
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $candidates.ToArray()
    }

    $cards = [regex]::Matches($Html, '(?is)<li\b[^>]*>.*?data-cy=["'']serpCard["''].*?</li>')
    $cardIndex = 0
    foreach ($card in $cards) {
        $cardIndex++
        Add-SourceMetric -Stats $Stats -Name "Candidates"
        if ($cardIndex -gt $MaxHelloWorkCardsPerQuery) {
            Add-SourceMetric -Stats $Stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $cards.Count - $MaxHelloWorkCardsPerQuery))
            break
        }

        $cardHtml = $card.Value
        $linkMatch = [regex]::Match($cardHtml, '(?is)<a\b[^>]*data-cy=["'']offerTitle["''][^>]*>.*?</a>')
        if (-not $linkMatch.Success) {
            $linkMatch = [regex]::Match($cardHtml, '(?is)<a\b[^>]*href=["'']/fr-fr/emplois/\d+\.html[^>]*>.*?</a>')
        }
        if (-not $linkMatch.Success) {
            continue
        }

        $linkHtml = $linkMatch.Value
        $jobUrl = ConvertTo-AbsoluteUrl -BaseUrl $SearchUrl -Href (Get-HtmlAttributeValue -Html $linkHtml -Name "href")
        if ([string]::IsNullOrWhiteSpace($jobUrl)) {
            continue
        }

        $title = ""
        $companyName = ""
        $titleAttribute = Get-HtmlAttributeValue -Html $linkHtml -Name "title"
        if ($titleAttribute -match "^(?<title>.+?)\s+-\s+(?<company>.+)$") {
            $title = Repair-DisplayText $matches["title"]
            $companyName = Repair-DisplayText $matches["company"]
        }
        if ([string]::IsNullOrWhiteSpace($title)) {
            $titleMatch = [regex]::Match($linkHtml, '(?is)<p[^>]*class=["''][^"'']*typo-l[^"'']*["''][^>]*>(?<title>.*?)</p>')
            if ($titleMatch.Success) {
                $title = ConvertFrom-HtmlText $titleMatch.Groups["title"].Value
            }
        }
        if ([string]::IsNullOrWhiteSpace($companyName)) {
            $paragraphs = @([regex]::Matches($linkHtml, '(?is)<p[^>]*>(?<text>.*?)</p>'))
            if ($paragraphs.Count -gt 1) {
                $companyName = ConvertFrom-HtmlText $paragraphs[1].Groups["text"].Value
            }
        }

        $location = ""
        $locationMatch = [regex]::Match($cardHtml, '(?is)data-cy=["'']localisationCard["''][^>]*>\s*(?<value>.*?)\s*</div>')
        if ($locationMatch.Success) {
            $location = ConvertFrom-HtmlText $locationMatch.Groups["value"].Value
        }

        $contractType = ""
        $contractMatch = [regex]::Match($cardHtml, '(?is)data-cy=["'']contractCard["''][^>]*>\s*(?<value>.*?)\s*</div>')
        if ($contractMatch.Success) {
            $contractType = ConvertFrom-HtmlText $contractMatch.Groups["value"].Value
        }

        $cardText = ConvertFrom-HtmlText $cardHtml
        if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text (Join-CleanTextParts @($title, $cardText)) -Reliable) {
            Add-SourceMetric -Stats $Stats -Name "SkippedContract"
            continue
        }

        $publishedAt = ConvertFrom-FrenchRelativeDateText $cardText
        if ($null -ne $publishedAt -and -not (Test-IsRecent $publishedAt)) {
            Add-SourceMetric -Stats $Stats -Name "SkippedOld"
            continue
        }

        $actualCandidateText = Join-CleanTextParts @($title, $companyName, $location, $contractType, $cardText)
        $rankingText = Join-CleanTextParts @($actualCandidateText, ("search query {0}" -f $Query))
        $actualMatch = Get-JobMatch -Title $title -Text $actualCandidateText
        $rankingMatch = Get-JobMatch -Title $title -Text $rankingText
        if (-not $actualMatch.IsMatch -and $actualCandidateText -notmatch $WttjUrlCandidatePattern -and $Query -notmatch $WttjUrlCandidatePattern) {
            Add-SourceMetric -Stats $Stats -Name "SkippedNoMatch"
            continue
        }

        $cardScore = $(if ($actualMatch.IsMatch) { $actualMatch.Score } elseif ($rankingMatch.IsMatch) { [Math]::Min(45, [int]$rankingMatch.Score) } else { 10 })
        $candidates.Add([PSCustomObject]@{
            Url          = $jobUrl
            Title        = $title
            Company      = $companyName
            Location     = $location
            Contract     = $contractType
            PublishedAt  = $publishedAt
            CardText     = $actualCandidateText
            Query        = $Query
            CardScore    = [int]$cardScore
            CardPosition = $cardIndex
        }) | Out-Null
    }

    return $candidates.ToArray()
}

function Get-HelloWorkJobs {
    Set-RunWindowTitle "Analytics Job Crawler - HelloWork"
    Write-RunStatus "Collecting HelloWork jobs from public search pages..."
    Write-RunStatus ("HelloWork plan: {0} query/queries, {1} page(s) each, then at most {2} unique detail page(s)." -f $ApiSearchQueries.Count, $MaxHelloWorkPages, $MaxHelloWorkDetails)

    $stats = Start-SourceStats "HelloWork"
    $results = New-Object System.Collections.Generic.List[object]
    $candidateByUrl = @{}
    $queryIndex = 0

    foreach ($query in $ApiSearchQueries) {
        $queryIndex++
        Write-RunStatus ("HelloWork query {0}/{1}: {2}" -f $queryIndex, $ApiSearchQueries.Count, $query)
        for ($page = 1; $page -le $MaxHelloWorkPages; $page++) {
            $searchUrl = Get-HelloWorkSearchUrl -Query $query -Page $page
            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $html = Invoke-TextRequest $searchUrl -Headers @{ "Accept" = "text/html,application/xhtml+xml" } -TimeoutSec 30
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                Write-Warning ("HelloWork search failed for '{0}' page {1}: {2}" -f $query, $page, $_.Exception.Message)
                break
            }

            $candidates = @(Get-HelloWorkCardCandidates -Html $html -SearchUrl $searchUrl -Query $query -Stats $stats)
            foreach ($candidate in $candidates) {
                if (-not $candidateByUrl.ContainsKey($candidate.Url) -or [int]$candidate.CardScore -gt [int]$candidateByUrl[$candidate.Url].CardScore) {
                    $candidateByUrl[$candidate.Url] = $candidate
                }
            }

            Write-CountProgress -Activity ("HelloWork search query {0}/{1}" -f $queryIndex, $ApiSearchQueries.Count) -Current $page -Total $MaxHelloWorkPages -Found $candidateByUrl.Count -Every 1
            if ($candidates.Count -eq 0) {
                break
            }
            Start-Sleep -Milliseconds $HelloWorkSearchDelayMilliseconds
        }
    }

    $selectedCandidates = @($candidateByUrl.Values |
        Sort-Object -Property `
            @{ Expression = "CardScore"; Descending = $true },
            @{ Expression = { if ($null -ne $_.PublishedAt) { $_.PublishedAt } else { [DateTimeOffset]::MinValue } }; Descending = $true },
            @{ Expression = "CardPosition"; Descending = $false } |
        Select-Object -First $MaxHelloWorkDetails)
    Add-SourceMetric -Stats $stats -Name "SelectedDetails" -Amount $selectedCandidates.Count
    Add-SourceMetric -Stats $stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $candidateByUrl.Count - $selectedCandidates.Count))

    Write-RunStatus ("HelloWork candidates selected: {0} unique detail page(s) from {1} candidate(s)." -f $selectedCandidates.Count, $candidateByUrl.Count)
    $detailIndex = 0
    foreach ($candidate in $selectedCandidates) {
        $detailIndex++
        Write-CountProgress -Activity "HelloWork detail pages" -Current $detailIndex -Total $selectedCandidates.Count -Found $results.Count -Every 5

        try {
            $html = Invoke-CachedTextRequest -Url $candidate.Url -CacheScope "hellowork-detail" -Headers @{ "Accept" = "text/html,application/xhtml+xml" } -TimeoutSec 30 -Stats $stats
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            Write-Warning ("HelloWork detail failed for '{0}': {1}" -f $candidate.Url, $_.Exception.Message)
            continue
        }

        $metadata = Get-HelloWorkJobMetadata -Html $html
        $title = Get-PreferredValue $metadata.Title $candidate.Title
        $companyName = Get-PreferredValue $metadata.Company $candidate.Company
        $jobLocation = Get-PreferredValue $metadata.Location $candidate.Location
        $pageTitle = Get-TitleFromHtml $html
        $sourceText = Join-CleanTextParts @($title, $companyName, $jobLocation, $metadata.Contract, $metadata.Description, $candidate.CardText, $pageTitle)
        $contractType = Get-ContractTypeFromText -Text $sourceText
        if ([string]::IsNullOrWhiteSpace($contractType)) {
            $contractType = Get-PreferredValue $metadata.Contract $candidate.Contract
        }
        if ([string]::IsNullOrWhiteSpace($contractType)) {
            $contractType = Get-ContractTypeFromText -Text $sourceText -RawContractType $metadata.EmploymentType
        }

        $publishedAt = $metadata.DatePosted
        if ($null -eq $publishedAt) {
            $publishedAt = $candidate.PublishedAt
        }
        if (-not (Test-IsRecent $publishedAt)) {
            Add-SourceMetric -Stats $stats -Name "SkippedOld"
            continue
        }

        $match = Get-JobMatch -Title $title -Text $sourceText
        if (-not $match.IsMatch) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }

        $result = New-JobResult -Title $title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "HelloWork" -PublishedAt $publishedAt -SourceText $sourceText
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
            Add-SourceMetric -Stats $stats -Name "Matches"
        }

        Start-Sleep -Milliseconds $HelloWorkDetailDelayMilliseconds
    }

    Write-RunStatus ("HelloWork complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-WelcomeKitJobs {
    if ([string]::IsNullOrWhiteSpace($WelcomeKitApiKey)) {
        Write-RunStatus "WelcomeKit API key not set; using WTTJ public sitemap fallback."
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - WTTJ API"
    Write-RunStatus "Collecting Welcome to the Jungle jobs through the official WelcomeKit API..."
    $stats = Start-SourceStats "WelcomeKit"
    $results = New-Object System.Collections.Generic.List[object]
    $headers = @{
        "Authorization" = "Bearer $WelcomeKitApiKey"
        "Accept"        = "application/json"
    }

    for ($page = 1; $page -le $MaxWelcomeKitPages; $page++) {
        Write-RunStatus ("WelcomeKit API page {0}/{1}; {2} matches so far." -f $page, $MaxWelcomeKitPages, $results.Count)
        $params = @{
            status          = "published"
            websites        = "true"
            per_page        = "100"
            page            = [string]$page
            published_after = $CutoffDate
        }
        $url = "https://www.welcomekit.co/api/v1/external/jobs/all?{0}" -f (ConvertTo-QueryString $params)

        try {
            Add-SourceMetric -Stats $stats -Name "SearchRequests"
            $jobs = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -TimeoutSec 45
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            Write-Warning ("WelcomeKit API call failed on page {0}: {1}" -f $page, $_.Exception.Message)
            break
        }

        $jobArray = @($jobs)
        if ($jobArray.Count -eq 0) {
            break
        }

        foreach ($job in $jobArray) {
            Add-SourceMetric -Stats $stats -Name "Candidates"
            $publishedAt = ConvertTo-DateTimeOffsetOrNull $job.published_at
            if (-not (Test-IsRecent $publishedAt)) {
                Add-SourceMetric -Stats $stats -Name "SkippedOld"
                continue
            }

            $combined = ConvertFrom-HtmlText ("{0} {1} {2} {3}" -f $job.name, $job.profile, $job.description, $job.company_description)
            $contractType = Get-ContractTypeFromText -Text $combined -RawContractType ([string]$job.contract_type)
            if (Test-ShouldSkipEarlyByContract -ContractType $contractType -Text $combined -Reliable) {
                Add-SourceMetric -Stats $stats -Name "SkippedContract"
                continue
            }

            $match = Get-JobMatch -Title ([string]$job.name) -Text $combined
            if (-not $match.IsMatch) {
                Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                continue
            }

            $jobUrl = Get-WelcomeKitJobUrl $job
            $jobDetails = $null
            if ($job.PSObject.Properties.Name -contains "reference") {
                Add-SourceMetric -Stats $stats -Name "DetailRequests"
                $jobDetails = Get-WelcomeKitJobDetails -Reference ([string]$job.reference) -Headers $headers
            }
            if ($null -ne $jobDetails) {
                $jobUrl = Get-WelcomeKitJobUrl $jobDetails
                $job = $jobDetails
            }

            $companyName = Get-WelcomeKitCompanyName -Job $job -JobUrl $jobUrl
            $jobLocation = Get-WelcomeKitLocation -Job $job -JobUrl $jobUrl
            $result = New-JobResult -Title ([string]$job.name) -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $jobUrl -Platform "Welcome to the Jungle" -PublishedAt $publishedAt -SourceText $combined
            if ($null -ne $result) {
                $results.Add($result) | Out-Null
                Add-SourceMetric -Stats $stats -Name "Matches"
            }
        }

        if ($jobArray.Count -lt 100) {
            break
        }
    }

    Write-RunStatus ("WelcomeKit API complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-WttjPublicFallbackJobs {
    if ($DisableWttjPublicFallback) {
        return @()
    }

    Set-RunWindowTitle "Analytics Job Crawler - WTTJ"
    Write-RunStatus "Collecting Welcome to the Jungle jobs from public sitemaps..."
    $stats = Start-SourceStats "WTTJ public"
    $results = New-Object System.Collections.Generic.List[object]
    $candidateSeen = @{}
    $candidates = New-Object System.Collections.Generic.List[object]

    try {
        Add-SourceMetric -Stats $stats -Name "SearchRequests"
        $indexXml = Invoke-CurlTextRequest "https://www.welcometothejungle.com/sitemaps/index.xml.gz"
    }
    catch {
        Add-SourceMetric -Stats $stats -Name "Errors"
        Write-Warning ("Could not read WTTJ sitemap index: {0}" -f $_.Exception.Message)
        Complete-SourceStats $stats
        return @()
    }

    $sitemapMatches = [regex]::Matches($indexXml, "<loc>(?<url>https://www\.welcometothejungle\.com/sitemaps/job-listings\.\d+\.xml\.gz)</loc>", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    Write-RunStatus ("WTTJ sitemap scan: {0} job sitemap(s) found." -f $sitemapMatches.Count)
    $sitemapCount = 0
    foreach ($sitemapMatch in $sitemapMatches) {
        $sitemapCount++
        Write-CountProgress -Activity "WTTJ sitemap scan" -Current $sitemapCount -Total $sitemapMatches.Count -Found $candidates.Count -Every 10
        $sitemapUrl = $sitemapMatch.Groups["url"].Value
        try {
            Add-SourceMetric -Stats $stats -Name "SearchRequests"
            $xml = Invoke-CurlTextRequest $sitemapUrl
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            Write-Warning ("Could not read WTTJ sitemap {0}: {1}" -f $sitemapUrl, $_.Exception.Message)
            continue
        }

        $urlMatches = [regex]::Matches($xml, "(?is)<url>.*?</url>")
        foreach ($urlMatch in $urlMatches) {
            $block = $urlMatch.Value
            $locMatch = [regex]::Match($block, "<loc>(?<loc>.*?)</loc>", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $lastmodMatch = [regex]::Match($block, "<lastmod>(?<lastmod>.*?)</lastmod>", [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if (-not $locMatch.Success -or -not $lastmodMatch.Success) {
                continue
            }

            $loc = ConvertFrom-HtmlAttribute $locMatch.Groups["loc"].Value
            if ($candidateSeen.ContainsKey($loc)) {
                continue
            }
            Add-SourceMetric -Stats $stats -Name "Candidates"

            $lastmod = ConvertTo-DateTimeOffsetOrNull $lastmodMatch.Groups["lastmod"].Value
            if (-not (Test-IsRecent $lastmod)) {
                Add-SourceMetric -Stats $stats -Name "SkippedOld"
                continue
            }

            if ($loc -notmatch $WttjUrlCandidatePattern) {
                Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
                continue
            }

            $slugTitle = Get-TitleFromWttjUrl $loc
            $urlText = "{0} {1}" -f $loc, (($slugTitle -replace "[-_]", " "))
            if (Test-ShouldSkipEarlyByContract -Text $urlText) {
                Add-SourceMetric -Stats $stats -Name "SkippedContract"
                continue
            }

            $candidateSeen[$loc] = $true
            $candidates.Add([PSCustomObject]@{
                Url       = $loc
                LastMod   = $lastmod
                SlugTitle = $slugTitle
                Score     = Get-WttjCandidateScore $loc
            }) | Out-Null
        }
    }

    $selectedCandidates = $candidates |
        Sort-Object -Property @{ Expression = "Score"; Descending = $true }, @{ Expression = "LastMod"; Descending = $true } |
        Select-Object -First $MaxWttjCandidatePages

    $selectedCandidateCount = @($selectedCandidates).Count
    Add-SourceMetric -Stats $stats -Name "SelectedDetails" -Amount $selectedCandidateCount
    Add-SourceMetric -Stats $stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $candidates.Count - $selectedCandidateCount))
    Write-RunStatus ("WTTJ candidates selected: {0} page(s) to inspect." -f $selectedCandidateCount)
    $count = 0
    foreach ($candidate in $selectedCandidates) {
        $count++
        Write-CountProgress -Activity "WTTJ candidate pages" -Current $count -Total $selectedCandidateCount -Found $results.Count -Every 10

        $urlText = "{0} {1}" -f $candidate.Url, (($candidate.SlugTitle -replace "[-_]", " "))
        $urlMatchResult = Get-JobMatch -Title $candidate.SlugTitle -Text $urlText
        $urlOnlyMatch = $urlMatchResult.IsMatch

        try {
            $html = Get-CachedText -Scope "wttj-detail" -Key $candidate.Url
            if ($null -ne $html) {
                Add-SourceMetric -Stats $stats -Name "CacheHits"
            }
            else {
                Add-SourceMetric -Stats $stats -Name "DetailRequests"
                $html = Invoke-CurlTextRequest $candidate.Url
                Set-CachedText -Scope "wttj-detail" -Key $candidate.Url -Text $html
            }
        }
        catch {
            Add-SourceMetric -Stats $stats -Name "Errors"
            if ($urlOnlyMatch) {
                $jobLocation = Get-WttjLocation -Html "" -Url $candidate.Url -Title $candidate.SlugTitle
                $fallbackResult = New-JobResult -Title $candidate.SlugTitle -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $urlText) -MatchScore $urlMatchResult.Score -MatchLevel $urlMatchResult.Level -MatchedKeywords $urlMatchResult.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $candidate.LastMod -SourceText $urlText
                if ($null -ne $fallbackResult) {
                    $results.Add($fallbackResult) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }
            continue
        }

        if ($html -match "(?i)<title>ERROR: The request could not be satisfied</title>|AwsWafIntegration|challenge-container") {
            if ($urlOnlyMatch) {
                $jobLocation = Get-WttjLocation -Html $html -Url $candidate.Url -Title $candidate.SlugTitle
                $fallbackResult = New-JobResult -Title $candidate.SlugTitle -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $urlText) -MatchScore $urlMatchResult.Score -MatchLevel $urlMatchResult.Level -MatchedKeywords $urlMatchResult.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $candidate.LastMod -SourceText $urlText
                if ($null -ne $fallbackResult) {
                    $results.Add($fallbackResult) | Out-Null
                    Add-SourceMetric -Stats $stats -Name "Matches"
                }
            }
            continue
        }

        $title = Get-TitleFromHtml $html
        if ([string]::IsNullOrWhiteSpace($title)) {
            $title = $candidate.SlugTitle
        }

        $publishedAt = $candidate.LastMod
        $publishedMatch = [regex]::Match($html, '(?i)"published_at"\s*:\s*"(?<date>[^"]+)"')
        if ($publishedMatch.Success) {
            $parsedPublished = ConvertTo-DateTimeOffsetOrNull $publishedMatch.Groups["date"].Value
            if ($null -ne $parsedPublished) {
                $publishedAt = $parsedPublished
            }
        }

        if (-not (Test-IsRecent $publishedAt)) {
            Add-SourceMetric -Stats $stats -Name "SkippedOld"
            continue
        }

        $pageText = ConvertFrom-HtmlText $html
        $combined = "{0} {1} {2}" -f $title, $candidate.Url, $pageText
        $match = Get-JobMatch -Title $title -Text $combined
        if (-not $match.IsMatch -and -not $urlOnlyMatch) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }
        if (-not $match.IsMatch) {
            $match = $urlMatchResult
        }

        $jobLocation = Get-WttjLocation -Html $html -Url $candidate.Url -Title $title
        $result = New-JobResult -Title $title -CompanyName (Get-WttjCompanyNameFromUrl $candidate.Url) -JobLocation $jobLocation -ContractType (Get-ContractTypeFromText -Text $combined) -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "Welcome to the Jungle" -PublishedAt $publishedAt -SourceText $combined
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
            Add-SourceMetric -Stats $stats -Name "Matches"
        }

        Start-Sleep -Milliseconds 250
    }

    Write-RunStatus ("WTTJ public fallback complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

function Get-LinkedInJobs {
    Set-RunWindowTitle "Analytics Job Crawler - LinkedIn"
    Write-RunStatus "Collecting LinkedIn jobs from public guest endpoints..."
    Write-RunStatus ("LinkedIn plan: {0} search query/queries, up to {1} page(s) each, then up to {2} ranked detail page(s)." -f $LinkedInQueries.Count, $MaxLinkedInSearchPages, $(if ($MaxLinkedInDetails -gt 0) { $MaxLinkedInDetails } else { "all" }))
    $stats = Start-SourceStats "LinkedIn"
    $results = New-Object System.Collections.Generic.List[object]
    $candidateById = @{}
    $seconds = [Math]::Max(86400, [int]([Math]::Abs($DaysBack) * 86400))

    $queryIndex = 0
    foreach ($query in $LinkedInQueries) {
        $queryIndex++
        Write-RunStatus ("LinkedIn query {0}/{1}: {2}" -f $queryIndex, $LinkedInQueries.Count, $query)
        for ($page = 0; $page -lt $MaxLinkedInSearchPages; $page++) {
            $start = $page * 25
            $params = @{
                keywords = $query
                location = $Location
                f_TPR    = "r$seconds"
                start    = [string]$start
            }
            $url = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?{0}" -f (ConvertTo-QueryString $params)

            try {
                Add-SourceMetric -Stats $stats -Name "SearchRequests"
                $html = Invoke-TextRequest $url -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
            }
            catch {
                Start-Sleep -Seconds 8
                try {
                    Add-SourceMetric -Stats $stats -Name "SearchRequests"
                    $html = Invoke-TextRequest $url -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30
                }
                catch {
                    Add-SourceMetric -Stats $stats -Name "Errors"
                    Write-Warning ("LinkedIn search failed for '{0}' page {1}: {2}" -f $query, ($page + 1), $_.Exception.Message)
                    break
                }
            }

            if ([string]::IsNullOrWhiteSpace($html) -or $html -notmatch "jobPosting") {
                Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: no result cards returned." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages)
                break
            }

            $cards = [regex]::Matches($html, '(?is)<li>.*?data-entity-urn="urn:li:jobPosting:(?<id>\d+)".*?</li>')
            if ($cards.Count -eq 0) {
                Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: no readable cards found." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages)
                break
            }
            Write-RunStatus ("LinkedIn query {0}/{1}, page {2}/{3}: {4} card(s) found; {5} unique candidate(s) so far." -f $queryIndex, $LinkedInQueries.Count, ($page + 1), $MaxLinkedInSearchPages, $cards.Count, $candidateById.Count)

            $cardIndex = 0
            foreach ($card in $cards) {
                $cardIndex++
                Add-SourceMetric -Stats $stats -Name "Candidates"
                $cardHtml = $card.Value
                $id = $card.Groups["id"].Value

                $titleMatch = [regex]::Match($cardHtml, '(?is)<h3[^>]*class="[^"]*base-search-card__title[^"]*"[^>]*>(?<title>.*?)</h3>')
                $urlMatch = [regex]::Match($cardHtml, '(?is)<a[^>]+href="(?<url>https?://[^"]+)"')
                $dateMatch = [regex]::Match($cardHtml, '(?is)<time[^>]+datetime="(?<date>[^"]+)"')

                if (-not $titleMatch.Success -or -not $urlMatch.Success -or -not $dateMatch.Success) {
                    continue
                }

                $title = ConvertFrom-HtmlText $titleMatch.Groups["title"].Value
                $companyMatch = [regex]::Match($cardHtml, '(?is)<h4[^>]*class="[^"]*base-search-card__subtitle[^"]*"[^>]*>(?<company>.*?)</h4>')
                $companyName = ""
                if ($companyMatch.Success) {
                    $companyName = ConvertFrom-HtmlText $companyMatch.Groups["company"].Value
                }
                $jobLocation = Get-LinkedInLocationFromHtml $cardHtml

                $jobUrl = ConvertTo-CleanUrl $urlMatch.Groups["url"].Value
                $publishedAt = ConvertTo-DateTimeOffsetOrNull $dateMatch.Groups["date"].Value
                if (-not (Test-IsRecent $publishedAt)) {
                    Add-SourceMetric -Stats $stats -Name "SkippedOld"
                    continue
                }

                $cardText = Join-CleanTextParts @($title, $companyName, $jobLocation, $jobUrl, (ConvertFrom-HtmlText $cardHtml))
                if (Test-ShouldSkipEarlyByContract -Text $cardText) {
                    Add-SourceMetric -Stats $stats -Name "SkippedContract"
                    continue
                }

                $cardMatch = Get-JobMatch -Title $title -Text $cardText
                $cardScore = 10
                if ($cardMatch.IsMatch) {
                    $cardScore = [int]$cardMatch.Score
                }
                elseif ($cardText -match $WttjUrlCandidatePattern) {
                    $cardScore = 45
                }

                $candidate = [PSCustomObject]@{
                    Id           = $id
                    Title        = $title
                    Company      = $companyName
                    Location     = $jobLocation
                    Url          = $jobUrl
                    PublishedAt  = $publishedAt
                    CardText     = $cardText
                    CardScore    = [int]$cardScore
                    QueryIndex   = $queryIndex
                    Page         = $page
                    CardPosition = $cardIndex
                }

                if (-not $candidateById.ContainsKey($id) -or [int]$candidate.CardScore -gt [int]$candidateById[$id].CardScore) {
                    $candidateById[$id] = $candidate
                }
            }

            Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
        }
    }

    $orderedCandidates = @($candidateById.Values |
        Sort-Object -Property `
            @{ Expression = "CardScore"; Descending = $true },
            @{ Expression = "PublishedAt"; Descending = $true },
            @{ Expression = "QueryIndex"; Descending = $false },
            @{ Expression = "Page"; Descending = $false },
            @{ Expression = "CardPosition"; Descending = $false })
    if ($MaxLinkedInDetails -gt 0) {
        $selectedCandidates = @($orderedCandidates | Select-Object -First $MaxLinkedInDetails)
    }
    else {
        $selectedCandidates = @($orderedCandidates)
    }

    Add-SourceMetric -Stats $stats -Name "SelectedDetails" -Amount $selectedCandidates.Count
    Add-SourceMetric -Stats $stats -Name "SkippedByCap" -Amount ([Math]::Max(0, $orderedCandidates.Count - $selectedCandidates.Count))
    Write-RunStatus ("LinkedIn candidates selected: {0} detail page(s) from {1} unique candidate(s)." -f $selectedCandidates.Count, $orderedCandidates.Count)

    $candidateIndex = 0
    foreach ($candidate in $selectedCandidates) {
        $candidateIndex++
        Write-CountProgress -Activity "LinkedIn detail pages" -Current $candidateIndex -Total $selectedCandidates.Count -Found $results.Count -Every 10

        $detailUrl = "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{0}" -f $candidate.Id
        $detailHtml = ""
        try {
            $cacheHitsBefore = [int]$stats["CacheHits"]
            $detailHtml = Invoke-CachedTextRequest -Url $detailUrl -CacheScope "linkedin-detail" -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30 -Stats $stats
            if ([int]$stats["CacheHits"] -eq $cacheHitsBefore) {
                Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
            }
        }
        catch {
            Start-Sleep -Seconds 5
            try {
                $cacheHitsBefore = [int]$stats["CacheHits"]
                $detailHtml = Invoke-CachedTextRequest -Url $detailUrl -CacheScope "linkedin-detail" -Headers @{ "Accept" = "text/html,*/*" } -TimeoutSec 30 -Stats $stats
                if ([int]$stats["CacheHits"] -eq $cacheHitsBefore) {
                    Start-Sleep -Milliseconds $LinkedInDelayMilliseconds
                }
            }
            catch {
                Add-SourceMetric -Stats $stats -Name "Errors"
                $detailHtml = ""
            }
        }

        $detailText = ConvertFrom-HtmlText $detailHtml
        $combined = Join-CleanTextParts @($candidate.Title, $candidate.Url, $candidate.CardText, $detailText)
        $match = Get-JobMatch -Title $candidate.Title -Text $combined
        if (-not $match.IsMatch) {
            Add-SourceMetric -Stats $stats -Name "SkippedNoMatch"
            continue
        }

        $companyName = $candidate.Company
        if ([string]::IsNullOrWhiteSpace($companyName)) {
            $detailCompanyMatch = [regex]::Match($detailHtml, '(?is)<a[^>]*class="[^"]*topcard__org-name-link[^"]*"[^>]*>(?<company>.*?)</a>')
            if ($detailCompanyMatch.Success) {
                $companyName = ConvertFrom-HtmlText $detailCompanyMatch.Groups["company"].Value
            }
        }

        $jobLocation = $candidate.Location
        if ([string]::IsNullOrWhiteSpace($jobLocation)) {
            $jobLocation = Get-LinkedInLocationFromHtml $detailHtml
        }

        $contractType = Get-LinkedInContractType -Title $candidate.Title -DetailText $detailText
        if (Test-IsExcludedContractType $contractType) {
            Add-SourceMetric -Stats $stats -Name "SkippedContract"
            continue
        }

        $result = New-JobResult -Title $candidate.Title -CompanyName $companyName -JobLocation $jobLocation -ContractType $contractType -MatchScore $match.Score -MatchLevel $match.Level -MatchedKeywords $match.Keywords -Url $candidate.Url -Platform "LinkedIn" -PublishedAt $candidate.PublishedAt -SourceText $combined
        if ($null -ne $result) {
            $results.Add($result) | Out-Null
            Add-SourceMetric -Stats $stats -Name "Matches"
        }
    }

    Write-RunStatus ("LinkedIn complete: {0} matching jobs." -f $results.Count)
    Complete-SourceStats $stats
    return $results.ToArray()
}

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
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType "Freelance") -Message "Expected freelance contracts to be excluded."

    $annonceurMatch = Get-JobMatch -Title "Web Analyst CRO" -Text "Google Tag Manager GA4 ContentSquare dataLayer tagging plan"
    Assert-ScoringCondition -Condition $annonceurMatch.IsMatch -Message "Expected a Web Analyst CRO role with web analytics tools to match."
    $expandedToolMatch = Get-JobMatch -Title "Tracking Specialist" -Text "Tag Commander Commanders Act Tealium server-side tracking RGPD"
    Assert-ScoringCondition -Condition $expandedToolMatch.IsMatch -Message "Expected Tag Commander, Commanders Act, Tealium, server-side, and RGPD signals to match."
    Assert-ScoringCondition -Condition ($expandedToolMatch.Keywords -match "Tag Commander" -and $expandedToolMatch.Keywords -match "Tealium" -and $expandedToolMatch.Keywords -match "server-side" -and $expandedToolMatch.Keywords -match "RGPD") -Message "Expected expanded tool/mission keywords to be reported."
    $positiveFeedbackRow = New-OrderedJobRecord @{
        status           = "interesting"
        job_title        = "Web Analyst"
        matched_keywords = "Tealium; server-side tracking"
    }
    $ignoredFeedbackRow = New-OrderedJobRecord @{
        status    = "ignored"
        job_title = "SEO Manager"
        notes     = "ignore_reason=too_seo_sea_marketing; detail=too marketing"
    }
    $script:FeedbackLearningProfile = New-FeedbackLearningProfile -Rows @($positiveFeedbackRow, $ignoredFeedbackRow)
    $positiveLearning = Get-FeedbackLearningAdjustment -FullText "tealium server side tracking" -HasCoreTitleSignal:$true -HasWebAnalyticsToolSignal:$true -HasDigitalAnalyticsContext:$true
    Assert-ScoringCondition -Condition ([int]$positiveLearning.Adjustment -gt 0 -and (($positiveLearning.Reasons -join ";") -match "Tealium")) -Message "Expected positive saved tracker feedback to boost similar tool signals."
    $negativeLearning = Get-FeedbackLearningAdjustment -FullText "seo sea paid media campaign" -HasCoreTitleSignal:$false -HasWebAnalyticsToolSignal:$false -HasDigitalAnalyticsContext:$false
    Assert-ScoringCondition -Condition ([int]$negativeLearning.Adjustment -lt 0 -and (($negativeLearning.Reasons -join ";") -match "SEO/SEA")) -Message "Expected ignored saved tracker feedback to penalize similar marketing-only signals."
    $script:FeedbackLearningProfile = $null
    $annonceurResult = New-JobResult `
        -Title "Web Analyst CRO" `
        -CompanyName "Radio France" `
        -JobLocation "Paris" `
        -ContractType "CDI" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/radio-france-web-analyst" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Google Tag Manager GA4 ContentSquare dataLayer tagging plan"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $annonceurResult -Name "employer_type") -eq "annonceur") -Message "Expected Radio France to be classified as annonceur."
    Assert-ScoringCondition -Condition ((Get-IntegerRowValue -Row $annonceurResult -Name "match_score") -gt (Get-IntegerRowValue -Row $annonceurResult -Name "role_score")) -Message "Expected annonceur/Paris/CDI fit to boost the role score."

    $consultingMatch = Get-JobMatch -Title "Digital Analytics Consultant" -Text "GA4 Google Tag Manager Piano Analytics ContentSquare"
    $consultingResult = New-JobResult `
        -Title "Digital Analytics Consultant" `
        -CompanyName "fifty-five" `
        -JobLocation "Paris" `
        -ContractType "CDI" `
        -MatchScore $consultingMatch.Score `
        -MatchLevel $consultingMatch.Level `
        -MatchedKeywords $consultingMatch.Keywords `
        -Url "https://example.test/jobs/fifty-five-digital-analytics-consultant" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "GA4 Google Tag Manager Piano Analytics ContentSquare"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $consultingResult -Name "employer_type") -eq "consulting") -Message "Expected fifty-five to be classified as consulting."
    Assert-ScoringCondition -Condition ((Get-IntegerRowValue -Row $consultingResult -Name "employer_fit") -lt 0) -Message "Expected consulting employer type to be demoted, not excluded."

    $dataEngineeringMatch = Get-JobMatch -Title "Data Analyst" -Text "python dbt snowflake airflow data warehouse data pipeline"
    Assert-ScoringCondition -Condition (-not $dataEngineeringMatch.IsMatch) -Message "Expected warehouse/python data analyst role without web analytics signals to stay below the match threshold."
    $companyNameOnlyToolMatch = Get-JobMatch -Title "People Business Partner" -Text "Contentsquare Paris Full-time"
    Assert-ScoringCondition -Condition (-not $companyNameOnlyToolMatch.IsMatch) -Message "Expected a non-analytics role not to match only because the company name is an analytics tool."

    $titleOnlyExcludedContract = New-JobResult `
        -Title "Alternance Assistant web analytics" `
        -CompanyName "Example Company" `
        -JobLocation "Paris" `
        -ContractType "" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/alternance-web-analytics" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Google Analytics tagging plan"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $titleOnlyExcludedContract -Name "contract_type") -eq "Apprenticeship") -Message "Expected title-only alternance to be mapped to Apprenticeship."
    Assert-ScoringCondition -Condition (Test-IsExcludedContractType (Get-RowValue -Row $titleOnlyExcludedContract -Name "contract_type")) -Message "Expected title-only alternance to be excluded by contract filtering."
    $titleOverridesGenericContract = New-JobResult `
        -Title "STAGE - Communication digitale et web analytics" `
        -CompanyName "Example Company" `
        -JobLocation "Paris" `
        -ContractType "Full-time" `
        -MatchScore $annonceurMatch.Score `
        -MatchLevel $annonceurMatch.Level `
        -MatchedKeywords $annonceurMatch.Keywords `
        -Url "https://example.test/jobs/stage-web-analytics" `
        -Platform "Test" `
        -PublishedAt ([DateTimeOffset]::Now) `
        -SourceText "Google Analytics tagging plan"
    Assert-ScoringCondition -Condition ((Get-RowValue -Row $titleOverridesGenericContract -Name "contract_type") -eq "Internship") -Message "Expected explicit STAGE title to override generic Full-time contract."

    $junkLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_5Kvvowa"
    Assert-ScoringCondition -Condition ([string]::IsNullOrWhiteSpace($junkLocation)) -Message "Expected random WTTJ URL suffixes not to become city names."
    $parisLocation = Get-WttjLocationFromUrl "https://www.welcometothejungle.com/fr/companies/acme/jobs/web-analyst_paris"
    Assert-ScoringCondition -Condition ($parisLocation -eq "Paris") -Message "Expected readable WTTJ city suffix to be kept."

    $franceTravailMock = [PSCustomObject]@{
        id                  = "123ABC"
        intitule            = "Web Analyst"
        description         = "Google Analytics GA4 Google Tag Manager dataLayer"
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
        title         = "Digital Analyst"
        description   = "Piano Analytics ContentSquare Google Analytics"
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
        intitule        = "Web Analyst F/H"
        nomCommercial   = "Example Retailer"
        lieuTexte       = "Paris - 75"
        typeContrat     = 101888
        texteOffre      = "Google Analytics GA4 Google Tag Manager ContentSquare"
        datePublication = ([DateTimeOffset]::Now.ToString("yyyy-MM-ddTHH:mm:ss.000+0000"))
    }
    Assert-ScoringCondition -Condition ((Get-ApecContractType $apecMock) -eq "CDI") -Message "Expected APEC CDI contract mapping."
    Assert-ScoringCondition -Condition ((Get-ApecJobUrl $apecMock) -match "/detail-offre/123456789W$") -Message "Expected APEC detail URL mapping."

    $helloWorkMockHtml = @'
<script type="application/ld+json">{"@context":"https://schema.org","@type":"JobPosting","title":"Web Analyst H/F","description":"Google Tag Manager GA4 dataLayer","datePosted":"2026-06-16T09:38:15Z","employmentType":"FULL_TIME","hiringOrganization":{"@type":"Organization","name":"Example Retailer"},"jobLocation":{"@type":"Place","address":{"@type":"PostalAddress","addressLocality":"Paris","addressRegion":"Ile-de-France","addressCountry":"FR"}}}</script>
<script type="application/ld+json">{"JobTitle":"Web Analyst H/F","Company":"Example Retailer","Localisation":"Paris - 75","ContractType":"CDI","Description":"Piano Analytics ContentSquare"}</script>
'@
    $helloWorkMetadata = Get-HelloWorkJobMetadata -Html $helloWorkMockHtml
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Title -eq "Web Analyst H/F") -Message "Expected HelloWork title metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Company -eq "Example Retailer") -Message "Expected HelloWork company metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Location -eq "Paris - 75") -Message "Expected HelloWork custom location metadata to be preferred."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Contract -eq "CDI") -Message "Expected HelloWork contract metadata mapping."
    Assert-ScoringCondition -Condition ($helloWorkMetadata.Description -match "Piano Analytics") -Message "Expected HelloWork custom description metadata mapping."
    Assert-ScoringCondition -Condition (Test-IsRecent (ConvertFrom-FrenchRelativeDateText "il y a 2 jours")) -Message "Expected French relative dates to parse as recent."

    $crossPlatformMatch = Get-JobMatch -Title "Web Analyst" -Text "GA4 Google Tag Manager ContentSquare dataLayer"
    $crossPlatformRows = @(
        (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.linkedin.com/jobs/view/111" -Platform "LinkedIn" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Analyste Web H/F" -CompanyName "Radio France" -JobLocation "75 - Paris" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://candidat.francetravail.fr/offres/recherche/detail/111" -Platform "France Travail" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Web Analyst" -CompanyName "Radio France" -JobLocation "Paris, Ile-de-France" -ContractType "Permanent" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.adzuna.fr/details/111" -Platform "Adzuna" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Web Analyst F/H" -CompanyName "Radio France" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/111W" -Platform "APEC" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer"),
        (New-JobResult -Title "Web Analyst H/F" -CompanyName "Radio France" -JobLocation "Paris - 75" -ContractType "CDI" -MatchScore $crossPlatformMatch.Score -MatchLevel $crossPlatformMatch.Level -MatchedKeywords $crossPlatformMatch.Keywords -Url "https://www.hellowork.com/fr-fr/emplois/111.html" -Platform "HelloWork" -PublishedAt ([DateTimeOffset]::Now) -SourceText "GA4 Google Tag Manager ContentSquare dataLayer")
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

$JobCrawlerPreferences = Get-JobCrawlerPreferences

if ($SelfTest) {
    Invoke-ScoringSelfTest
    return
}

Set-RunWindowTitle "Analytics Job Crawler - Starting"
Write-RunStatus ("Starting crawl for jobs published since {0}." -f $CutoffDate)
Write-RunStatus ("Tracker file: {0}" -f $TrackerPath)

$existingTrackerRows = @(Import-TrackerRows -Path $TrackerPath)
Write-RunStatus ("Loaded {0} existing tracker row(s)." -f $existingTrackerRows.Count)
$script:FeedbackLearningProfile = New-FeedbackLearningProfile -Rows $existingTrackerRows
Write-RunStatus ("Feedback profile: {0} positive row(s), {1} ignored row(s)." -f $script:FeedbackLearningProfile.PositiveRows, $script:FeedbackLearningProfile.IgnoredRows)

$allResults = New-Object System.Collections.Generic.List[object]

if (-not $SkipFranceTravail) {
    foreach ($result in @(Get-FranceTravailJobs)) {
        $allResults.Add($result) | Out-Null
    }
}

if (-not $SkipAdzuna) {
    foreach ($result in @(Get-AdzunaJobs)) {
        $allResults.Add($result) | Out-Null
    }
}

if (-not $SkipApec) {
    foreach ($result in @(Get-ApecJobs)) {
        $allResults.Add($result) | Out-Null
    }
}

if (-not $SkipHelloWork) {
    foreach ($result in @(Get-HelloWorkJobs)) {
        $allResults.Add($result) | Out-Null
    }
}

if (-not $SkipWttj) {
    $welcomeKitResults = @(Get-WelcomeKitJobs)
    foreach ($result in $welcomeKitResults) {
        $allResults.Add($result) | Out-Null
    }

    if ($welcomeKitResults.Count -eq 0 -and [string]::IsNullOrWhiteSpace($WelcomeKitApiKey)) {
        foreach ($result in @(Get-WttjPublicFallbackJobs)) {
            $allResults.Add($result) | Out-Null
        }
    }
}

if (-not $SkipLinkedIn) {
    foreach ($result in @(Get-LinkedInJobs)) {
        $allResults.Add($result) | Out-Null
    }
}

$sortedCrawlResults = $allResults |
    Sort-Object -Property @{ Expression = "match_score"; Descending = $true }, @{ Expression = "published_date"; Descending = $true }, platform, job_title
$filteredCrawlResults = @($sortedCrawlResults | Where-Object { -not (Test-IsExcludedContractType (Get-RowValue -Row $_ -Name "contract_type")) })
$excludedContractCount = @($sortedCrawlResults).Count - @($filteredCrawlResults).Count

if ($excludedContractCount -gt 0) {
    Write-RunStatus ("Excluded {0} CDD/apprenticeship/internship/freelance job(s) from this crawl." -f $excludedContractCount)
}

$feedbackAdjustedResults = @(Apply-FeedbackScoring -Rows $filteredCrawlResults -ExistingRows $existingTrackerRows)
$mergeResult = Merge-JobsWithTracker -CurrentRows $feedbackAdjustedResults -ExistingRows $existingTrackerRows -Path $TrackerPath
$finalResults = @($mergeResult.TrackerRows)

$crawlSummary = @{
    TotalMatched = @($sortedCrawlResults).Count
    ExcludedContractCount = $excludedContractCount
    CurrentCount = @($mergeResult.CurrentRows).Count
    TrackerCount = @($finalResults).Count
    DuplicateCount = $mergeResult.DuplicateCount
    RemovedCount = $mergeResult.RemovedCount
    PreservedAppliedCount = $mergeResult.PreservedAppliedCount
    SourceDiagnostics = Get-SourceStatsSummaryText
    BackupPath = $mergeResult.BackupPath
}
Export-TrackerWorkbook -Rows $finalResults -Path $TrackerPath -Summary $crawlSummary

Set-RunWindowTitle "Analytics Job Crawler - Finished"
Write-Host ""
Write-RunStatus ("Wrote tracker with {0} row(s): {1} current job(s), {2} preserved application row(s), {3} removed by retention." -f @($finalResults).Count, @($mergeResult.CurrentRows).Count, $mergeResult.PreservedAppliedCount, $mergeResult.RemovedCount)
Write-RunStatus ("Crawl mode: {0}. Source diagnostics: {1}" -f $CrawlMode, (Get-SourceStatsSummaryText))
if (-not [string]::IsNullOrWhiteSpace($mergeResult.BackupPath)) {
    Write-RunStatus ("Backup: {0}" -f $mergeResult.BackupPath)
}
Write-RunStatus ("Tracker: {0}" -f (Resolve-Path $TrackerPath).Path)

if (@($finalResults).Count -gt 0) {
    $finalResults | Format-Table -AutoSize
}
