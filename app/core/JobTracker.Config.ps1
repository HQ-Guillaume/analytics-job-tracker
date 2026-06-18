# Configuration helpers for the crawler. Values that may reasonably change over
# time live in config/*.json; structural workbook internals stay in code.

$profileModulePath = Join-Path $PSScriptRoot "JobTracker.Profile.ps1"
if (Test-Path -LiteralPath $profileModulePath) {
    . $profileModulePath
}

function Read-JobCrawlerJsonConfig {
    param(
        [string]$Path,
        [AllowNull()]$DefaultValue = $null
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $DefaultValue
    }

    try {
        return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Could not read config file '$Path': $($_.Exception.Message)"
    }
}

function Get-ConfigProperty {
    param(
        [AllowNull()]$Object,
        [string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }
    if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($Name)) {
        return $Object[$Name]
    }
    foreach ($property in @($Object.PSObject.Properties)) {
        if ($property.Name -eq $Name) {
            return $property.Value
        }
    }

    return $DefaultValue
}

function Get-ConfigPathValue {
    param(
        [AllowNull()]$Object,
        [string]$Path,
        [AllowNull()]$DefaultValue = $null
    )

    $current = $Object
    foreach ($part in ($Path -split "\.")) {
        $current = Get-ConfigProperty -Object $current -Name $part -DefaultValue $null
        if ($null -eq $current) {
            return $DefaultValue
        }
    }

    return $current
}

function Get-ConfigStringArray {
    param(
        [AllowNull()]$Value,
        [string[]]$DefaultValue = @()
    )

    if ($null -eq $Value) {
        return @($DefaultValue)
    }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) {
            return @()
        }
        return @([string]$Value)
    }

    return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Get-ConfigPropertyNames {
    param([AllowNull()]$Object)

    if ($null -eq $Object) {
        return @()
    }
    if ($Object -is [System.Collections.IDictionary]) {
        return @($Object.Keys | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    return @($Object.PSObject.Properties | ForEach-Object { [string]$_.Name } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Resolve-JobCrawlerPath {
    param(
        [string]$BasePath,
        [AllowNull()][string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }
    if ([IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path $BasePath $Path
}

function ConvertTo-ConfigHashtable {
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [string]) {
        return $Value
    }
    if ($Value -is [System.Collections.IDictionary]) {
        $hash = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $hash[[string]$key] = ConvertTo-ConfigHashtable $Value[$key]
        }
        return $hash
    }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [pscustomobject])) {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($item in $Value) {
            $items.Add((ConvertTo-ConfigHashtable $item)) | Out-Null
        }
        return @($items.ToArray())
    }
    if ($Value -is [pscustomobject]) {
        $hash = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-ConfigHashtable $property.Value
        }
        return $hash
    }

    return $Value
}

function Merge-ConfigHashtable {
    param(
        [AllowNull()]$Base,
        [AllowNull()]$Override
    )

    if ($null -eq $Override) {
        return $Base
    }
    if ($null -eq $Base) {
        return $Override
    }
    if ($Base -is [System.Collections.IDictionary] -and $Override -is [System.Collections.IDictionary]) {
        $merged = [ordered]@{}
        foreach ($key in $Base.Keys) {
            $merged[$key] = $Base[$key]
        }
        foreach ($key in $Override.Keys) {
            if ($merged.Contains($key)) {
                $merged[$key] = Merge-ConfigHashtable -Base $merged[$key] -Override $Override[$key]
            }
            else {
                $merged[$key] = $Override[$key]
            }
        }
        return $merged
    }

    return $Override
}

function Merge-JobCrawlerConfigObjects {
    param(
        [AllowNull()]$Base,
        [AllowNull()]$Override
    )

    $baseHash = ConvertTo-ConfigHashtable $Base
    $overrideHash = ConvertTo-ConfigHashtable $Override
    $merged = Merge-ConfigHashtable -Base $baseHash -Override $overrideHash
    if ($null -eq $merged) {
        return [PSCustomObject]@{}
    }

    return ($merged | ConvertTo-Json -Depth 100 | ConvertFrom-Json)
}

function Read-JobCrawlerPublicConfig {
    param(
        [string]$Root,
        [string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    return Read-JobCrawlerJsonConfig -Path (Join-Path $Root ("{0}.json" -f $Name)) -DefaultValue $DefaultValue
}

function Merge-JobCrawlerLocalConfig {
    param(
        [string]$Root,
        [string]$Name,
        [AllowNull()]$Value,
        [System.Collections.Generic.List[string]]$AppliedOverrides
    )

    $mergedValue = $Value
    $overridePaths = @(
        (Join-Path $Root ("local.{0}.json" -f $Name)),
        (Join-Path (Join-Path $Root "local") ("{0}.json" -f $Name))
    )

    foreach ($overridePath in $overridePaths) {
        if (Test-Path -LiteralPath $overridePath) {
            $overrideValue = Read-JobCrawlerJsonConfig -Path $overridePath -DefaultValue ([PSCustomObject]@{})
            $mergedValue = Merge-JobCrawlerConfigObjects -Base $mergedValue -Override $overrideValue
            if ($null -ne $AppliedOverrides -and -not $AppliedOverrides.Contains($overridePath)) {
                $AppliedOverrides.Add($overridePath) | Out-Null
            }
        }
    }

    return $mergedValue
}

function Get-JobCrawlerProfileDirectories {
    param([string]$Root)

    return @(
        [PSCustomObject]@{ Path = (Join-Path $Root "profiles"); IsLocal = $false },
        [PSCustomObject]@{ Path = (Join-Path (Join-Path $Root "local") "profiles"); IsLocal = $true }
    )
}

function Get-JobCrawlerProfileSummaries {
    param([string]$ConfigDirectory)

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        return @()
    }

    $profilesById = [ordered]@{}
    foreach ($directory in @(Get-JobCrawlerProfileDirectories -Root $configRoot.Path)) {
        if (-not (Test-Path -LiteralPath $directory.Path)) {
            continue
        }

        foreach ($file in @(Get-ChildItem -LiteralPath $directory.Path -Filter "*.json" -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $profile = Read-JobCrawlerJsonConfig -Path $file.FullName -DefaultValue $null
            if ($null -eq $profile) {
                continue
            }

            $id = [string](Get-ConfigProperty -Object $profile -Name "id" -DefaultValue ([IO.Path]::GetFileNameWithoutExtension($file.Name)))
            $id = ConvertTo-JobCrawlerProfileId $id
            if ([string]::IsNullOrWhiteSpace($id)) {
                continue
            }

            $label = [string](Get-ConfigProperty -Object $profile -Name "label" -DefaultValue $id)
            $description = [string](Get-ConfigProperty -Object $profile -Name "description" -DefaultValue "")
            $existing = $null
            if ($profilesById.Contains($id)) {
                $existing = $profilesById[$id]
            }

            $profilesById[$id] = [PSCustomObject]@{
                Id          = $id
                Label       = $label
                Description = $description
                Path        = $file.FullName
                IsLocal     = [bool]$directory.IsLocal
                HasPublic   = ($null -ne $existing -and [bool]$existing.HasPublic) -or (-not [bool]$directory.IsLocal)
                HasLocal    = ($null -ne $existing -and [bool]$existing.HasLocal) -or [bool]$directory.IsLocal
                Display     = ("{0} ({1})" -f $label, $id)
            }
        }
    }

    return @($profilesById.Values)
}

function Read-JobCrawlerProfile {
    param(
        [string]$Root,
        [string]$ProfileId,
        [System.Collections.Generic.List[string]]$AppliedOverrides
    )

    $selectedId = ConvertTo-JobCrawlerProfileId $ProfileId
    if ([string]::IsNullOrWhiteSpace($selectedId)) {
        return $null
    }

    $profile = $null
    foreach ($directory in @(Get-JobCrawlerProfileDirectories -Root $Root)) {
        $path = Join-Path $directory.Path ("{0}.json" -f $selectedId)
        if (-not (Test-Path -LiteralPath $path)) {
            continue
        }

        $profileValue = Read-JobCrawlerJsonConfig -Path $path -DefaultValue ([PSCustomObject]@{})
        $profile = Merge-JobCrawlerConfigObjects -Base $profile -Override $profileValue
        if ([bool]$directory.IsLocal -and $null -ne $AppliedOverrides -and -not $AppliedOverrides.Contains($path)) {
            $AppliedOverrides.Add($path) | Out-Null
        }
    }

    if ($null -ne $profile) {
        $profile = Expand-JobCrawlerProfile -Profile $profile
    }

    return $profile
}

function Merge-JobCrawlerProfileSection {
    param(
        [AllowNull()]$Base,
        [AllowNull()]$Profile,
        [string]$SectionName
    )

    $profileSection = Get-ConfigProperty -Object $Profile -Name $SectionName -DefaultValue $null
    if ($null -eq $profileSection) {
        return $Base
    }

    return Merge-JobCrawlerConfigObjects -Base $Base -Override $profileSection
}

function Get-JobCrawlerConfig {
    param(
        [string]$ConfigDirectory,
        [string]$ProfileId = ""
    )

    $configRoot = Resolve-Path -LiteralPath $ConfigDirectory -ErrorAction SilentlyContinue
    if ($null -eq $configRoot) {
        throw "Config directory not found: $ConfigDirectory"
    }

    $root = $configRoot.Path
    $appliedOverrides = New-Object System.Collections.Generic.List[string]

    $runtimeBase = Read-JobCrawlerPublicConfig -Root $root -Name "runtime" -DefaultValue ([PSCustomObject]@{})
    $crawlModesBase = Read-JobCrawlerPublicConfig -Root $root -Name "crawl_modes" -DefaultValue ([PSCustomObject]@{})
    $sourcesBase = Read-JobCrawlerPublicConfig -Root $root -Name "sources" -DefaultValue ([PSCustomObject]@{})
    $matchingRulesBase = Read-JobCrawlerPublicConfig -Root $root -Name "matching_rules" -DefaultValue ([PSCustomObject]@{})
    $workbookBase = Read-JobCrawlerPublicConfig -Root $root -Name "workbook" -DefaultValue ([PSCustomObject]@{})
    $preferencesBase = Read-JobCrawlerPublicConfig -Root $root -Name "preferences" -DefaultValue ([PSCustomObject]@{})

    $runtimeForProfileSelection = Merge-JobCrawlerLocalConfig -Root $root -Name "runtime" -Value $runtimeBase -AppliedOverrides $null
    $selectedProfileId = ConvertTo-JobCrawlerProfileId $ProfileId
    if ([string]::IsNullOrWhiteSpace($selectedProfileId)) {
        $selectedProfileId = ConvertTo-JobCrawlerProfileId ([string](Get-ConfigPathValue -Object $runtimeForProfileSelection -Path "defaults.profile_id" -DefaultValue ""))
    }
    if ([string]::IsNullOrWhiteSpace($selectedProfileId)) {
        $profileSummaries = @(Get-JobCrawlerProfileSummaries -ConfigDirectory $root | Sort-Object IsLocal, Label, Id)
        if ($profileSummaries.Count -gt 0) {
            $selectedProfileId = [string]$profileSummaries[0].Id
        }
    }

    $profile = $null
    if (-not [string]::IsNullOrWhiteSpace($selectedProfileId)) {
        $profile = Read-JobCrawlerProfile -Root $root -ProfileId $selectedProfileId -AppliedOverrides $appliedOverrides
        if ($null -eq $profile) {
            throw "Job crawler profile '$selectedProfileId' was not found. Open the GUI and create or select a profile."
        }
    }

    $runtimeConfig = Merge-JobCrawlerProfileSection -Base $runtimeBase -Profile $profile -SectionName "runtime"
    $crawlModesConfig = Merge-JobCrawlerProfileSection -Base $crawlModesBase -Profile $profile -SectionName "crawl_modes"
    $sourcesConfig = Merge-JobCrawlerProfileSection -Base $sourcesBase -Profile $profile -SectionName "sources"
    $matchingRulesConfig = Merge-JobCrawlerProfileSection -Base $matchingRulesBase -Profile $profile -SectionName "matching_rules"
    $workbookConfig = Merge-JobCrawlerProfileSection -Base $workbookBase -Profile $profile -SectionName "workbook"
    $preferencesConfig = Merge-JobCrawlerProfileSection -Base $preferencesBase -Profile $profile -SectionName "preferences"

    $runtimeConfig = Merge-JobCrawlerLocalConfig -Root $root -Name "runtime" -Value $runtimeConfig -AppliedOverrides $appliedOverrides
    $crawlModesConfig = Merge-JobCrawlerLocalConfig -Root $root -Name "crawl_modes" -Value $crawlModesConfig -AppliedOverrides $appliedOverrides
    $sourcesConfig = Merge-JobCrawlerLocalConfig -Root $root -Name "sources" -Value $sourcesConfig -AppliedOverrides $appliedOverrides
    $matchingRulesConfig = Merge-JobCrawlerLocalConfig -Root $root -Name "matching_rules" -Value $matchingRulesConfig -AppliedOverrides $appliedOverrides
    $workbookConfig = Merge-JobCrawlerLocalConfig -Root $root -Name "workbook" -Value $workbookConfig -AppliedOverrides $appliedOverrides
    $preferencesConfig = Merge-JobCrawlerLocalConfig -Root $root -Name "preferences" -Value $preferencesConfig -AppliedOverrides $appliedOverrides

    return [PSCustomObject]@{
        Root           = $root
        LocalOverrides = @($appliedOverrides.ToArray())
        Profile        = [PSCustomObject]@{
            Id          = $selectedProfileId
            Label       = $(if ($null -eq $profile) { "No profile configured" } else { [string](Get-ConfigProperty -Object $profile -Name "label" -DefaultValue $selectedProfileId) })
            Description = $(if ($null -eq $profile) { "Create a profile in the GUI before crawling." } else { [string](Get-ConfigProperty -Object $profile -Name "description" -DefaultValue "") })
            IsConfigured = ($null -ne $profile)
            Raw         = $profile
        }
        Runtime        = $runtimeConfig
        CrawlModes     = $crawlModesConfig
        Sources        = $sourcesConfig
        MatchingRules  = $matchingRulesConfig
        Workbook       = $workbookConfig
        Preferences    = $preferencesConfig
    }
}

function New-JobCrawlerConfigWithProfile {
    param(
        [AllowNull()]$Config,
        [AllowNull()]$Profile
    )

    if ($null -eq $Config) {
        throw "Base crawler config is required."
    }
    if ($null -eq $Profile) {
        throw "Profile is required."
    }

    $expandedProfile = Expand-JobCrawlerProfile -Profile $Profile
    $profileId = ConvertTo-JobCrawlerProfileId ([string](Get-ConfigProperty -Object $expandedProfile -Name "id" -DefaultValue ""))
    if ([string]::IsNullOrWhiteSpace($profileId)) {
        throw "Profile id is required."
    }

    return [PSCustomObject]@{
        Root           = [string](Get-ConfigProperty -Object $Config -Name "Root" -DefaultValue "")
        LocalOverrides = @(Get-ConfigProperty -Object $Config -Name "LocalOverrides" -DefaultValue @())
        Profile        = [PSCustomObject]@{
            Id           = $profileId
            Label        = [string](Get-ConfigProperty -Object $expandedProfile -Name "label" -DefaultValue $profileId)
            Description  = [string](Get-ConfigProperty -Object $expandedProfile -Name "description" -DefaultValue "")
            IsConfigured = $true
            Raw          = $expandedProfile
        }
        Runtime        = Merge-JobCrawlerProfileSection -Base (Get-ConfigProperty -Object $Config -Name "Runtime" -DefaultValue ([PSCustomObject]@{})) -Profile $expandedProfile -SectionName "runtime"
        CrawlModes     = Merge-JobCrawlerProfileSection -Base (Get-ConfigProperty -Object $Config -Name "CrawlModes" -DefaultValue ([PSCustomObject]@{})) -Profile $expandedProfile -SectionName "crawl_modes"
        Sources        = Merge-JobCrawlerProfileSection -Base (Get-ConfigProperty -Object $Config -Name "Sources" -DefaultValue ([PSCustomObject]@{})) -Profile $expandedProfile -SectionName "sources"
        MatchingRules  = Merge-JobCrawlerProfileSection -Base (Get-ConfigProperty -Object $Config -Name "MatchingRules" -DefaultValue ([PSCustomObject]@{})) -Profile $expandedProfile -SectionName "matching_rules"
        Workbook       = Merge-JobCrawlerProfileSection -Base (Get-ConfigProperty -Object $Config -Name "Workbook" -DefaultValue ([PSCustomObject]@{})) -Profile $expandedProfile -SectionName "workbook"
        Preferences    = Merge-JobCrawlerProfileSection -Base (Get-ConfigProperty -Object $Config -Name "Preferences" -DefaultValue ([PSCustomObject]@{})) -Profile $expandedProfile -SectionName "preferences"
    }
}

function Test-JobCrawlerProfileConfigured {
    param([AllowNull()]$Config)

    if ($null -eq $Config -or $null -eq $Config.Profile) {
        return $false
    }

    return [bool](Get-ConfigProperty -Object $Config.Profile -Name "IsConfigured" -DefaultValue $false)
}

function Expand-JobCrawlerProfilePathTemplate {
    param(
        [AllowNull()][string]$Path,
        [AllowNull()]$Config
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    if (-not (Test-JobCrawlerProfileConfigured -Config $Config)) {
        return ([string]$Path) -replace "\{profile_id\}", "unconfigured" -replace "\{profile\}", "unconfigured"
    }

    $profileId = ConvertTo-JobCrawlerProfileId ([string](Get-ConfigProperty -Object $Config.Profile -Name "Id" -DefaultValue ""))
    if ([string]::IsNullOrWhiteSpace($profileId)) {
        $profileId = "profile"
    }

    return ([string]$Path) -replace "\{profile_id\}", $profileId -replace "\{profile\}", $profileId
}

function Get-JobCrawlerTrackerPath {
    param(
        [string]$ProjectRoot,
        [AllowNull()]$Config,
        [AllowNull()][string]$TrackerPath = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($TrackerPath)) {
        return Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path $TrackerPath
    }

    $runtimeConfig = Get-ConfigProperty -Object $Config -Name "Runtime" -DefaultValue $null
    $configuredPath = [string](Get-ConfigPathValue -Object $runtimeConfig -Path "defaults.tracker_path" -DefaultValue "output\profiles\{profile_id}\jobs_tracker.xlsx")
    if ([string]::IsNullOrWhiteSpace($configuredPath)) {
        $configuredPath = "output\profiles\{profile_id}\jobs_tracker.xlsx"
    }

    if ((Test-JobCrawlerProfileConfigured -Config $Config) -and $configuredPath -notmatch "\{profile(_id)?\}") {
        $normalized = $configuredPath.Replace("/", "\").ToLowerInvariant()
        if ($normalized -eq "output\jobs_tracker.xlsx") {
            $profileId = ConvertTo-JobCrawlerProfileId ([string](Get-ConfigProperty -Object $Config.Profile -Name "Id" -DefaultValue "profile"))
            $configuredPath = "output\profiles\{0}\jobs_tracker.xlsx" -f $profileId
        }
    }

    $expandedPath = Expand-JobCrawlerProfilePathTemplate -Path $configuredPath -Config $Config
    return Resolve-JobCrawlerPath -BasePath $ProjectRoot -Path $expandedPath
}

function ConvertTo-ConfigBoolean {
    param(
        [AllowNull()]$Value,
        [bool]$DefaultValue = $false
    )

    if ($null -eq $Value) {
        return $DefaultValue
    }
    if ($Value -is [bool]) {
        return [bool]$Value
    }

    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ($text -in @("1", "true", "yes", "y", "on")) {
        return $true
    }
    if ($text -in @("0", "false", "no", "n", "off")) {
        return $false
    }

    return $DefaultValue
}

function Test-JobCrawlerSourceEnabledByDefault {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey,
        [bool]$DefaultValue = $true
    )

    $value = Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.enabled_by_default" -f $SourceKey) -DefaultValue $null
    return ConvertTo-ConfigBoolean -Value $value -DefaultValue $DefaultValue
}

function Get-JobCrawlerConfiguredSourceOrder {
    param([AllowNull()]$SourcesConfig)

    $sourceOrder = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $SourcesConfig -Path "source_order" -DefaultValue @()))
    if ($sourceOrder.Count -eq 0) {
        $sources = Get-ConfigProperty -Object $SourcesConfig -Name "sources" -DefaultValue $null
        $sourceOrder = @(Get-ConfigPropertyNames -Object $sources)
    }

    if ($sourceOrder.Count -eq 0) {
        # Last-resort compatibility for older local configs that predate sources.json.
        $sourceOrder = @("apec", "hellowork", "wttj_public", "linkedin", "france_travail", "adzuna")
    }

    return @($sourceOrder | Select-Object -Unique)
}

function Get-JobCrawlerSourceDefinitions {
    param([AllowNull()]$SourcesConfig)

    $sourceOrder = @(Get-JobCrawlerConfiguredSourceOrder -SourcesConfig $SourcesConfig)

    $defaults = @{
        apec = @{
            Label = "APEC"; Enabled = $true; Credentials = $false; Function = "Get-ApecJobs"; Skip = "SkipApec"; Enable = ""; FallbackFor = ""
        }
        hellowork = @{
            Label = "HelloWork"; Enabled = $true; Credentials = $false; Function = "Get-HelloWorkJobs"; Skip = "SkipHelloWork"; Enable = ""; FallbackFor = ""
        }
        wttj_public = @{
            Label = "Welcome to the Jungle public"; Enabled = $true; Credentials = $false; Function = "Get-WttjPublicFallbackJobs"; Skip = "DisableWttjPublicFallback"; Enable = ""; FallbackFor = ""
        }
        linkedin = @{
            Label = "LinkedIn public guest"; Enabled = $true; Credentials = $false; Function = "Get-LinkedInJobs"; Skip = "SkipLinkedIn"; Enable = ""; FallbackFor = ""
        }
        france_travail = @{
            Label = "France Travail API"; Enabled = $false; Credentials = $true; Function = "Get-FranceTravailJobs"; Skip = "SkipFranceTravail"; Enable = "EnableFranceTravail"; FallbackFor = ""
        }
        adzuna = @{
            Label = "Adzuna API"; Enabled = $false; Credentials = $true; Function = "Get-AdzunaJobs"; Skip = "SkipAdzuna"; Enable = "EnableAdzuna"; FallbackFor = ""
        }
    }

    $definitions = New-Object System.Collections.Generic.List[object]
    foreach ($sourceKey in $sourceOrder) {
        if ([string]::IsNullOrWhiteSpace($sourceKey)) {
            continue
        }

        $key = [string]$sourceKey
        $fallback = $defaults[$key]
        if ($null -eq $fallback) {
            $fallback = @{ Label = $key; Enabled = $true; Credentials = $false; Function = ""; Skip = ""; Enable = ""; FallbackFor = "" }
        }

        $definition = [PSCustomObject]@{
            Key                = $key
            Label              = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.label" -f $key) -DefaultValue $fallback.Label)
            ShortLabel         = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.short_label" -f $key) -DefaultValue $fallback.Label)
            EnabledByDefault   = Test-JobCrawlerSourceEnabledByDefault -SourcesConfig $SourcesConfig -SourceKey $key -DefaultValue ([bool]$fallback.Enabled)
            RequiresCredential = ConvertTo-ConfigBoolean -Value (Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.requires_credentials" -f $key) -DefaultValue $fallback.Credentials) -DefaultValue ([bool]$fallback.Credentials)
            CrawlFunction      = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.crawl_function" -f $key) -DefaultValue $fallback.Function)
            SkipSwitch         = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.skip_switch" -f $key) -DefaultValue $fallback.Skip)
            EnableSwitch       = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.enable_switch" -f $key) -DefaultValue $fallback.Enable)
            FallbackFor        = [string](Get-ConfigPathValue -Object $SourcesConfig -Path ("sources.{0}.fallback_for" -f $key) -DefaultValue $fallback.FallbackFor)
        }
        $definitions.Add($definition) | Out-Null
    }

    return @($definitions.ToArray())
}

function Get-JobCrawlerSourceQueryList {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey,
        [string[]]$FallbackKeys = @("api")
    )

    foreach ($key in (@($SourceKey) + @($FallbackKeys))) {
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        $queries = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $SourcesConfig -Path ("queries.{0}" -f $key) -DefaultValue @()))
        if ($queries.Count -gt 0) {
            return @($queries)
        }
    }

    return @()
}

function Get-JobCrawlerCredentialValue {
    param(
        [AllowNull()]$SourcesConfig,
        [string]$SourceKey,
        [string]$CredentialKey,
        [AllowNull()][string]$FallbackValue = ""
    )

    $envName = Get-ConfigPathValue -Object $SourcesConfig -Path ("credentials.{0}.{1}.env" -f $SourceKey, $CredentialKey) -DefaultValue ""
    if ([string]::IsNullOrWhiteSpace($envName)) {
        return $FallbackValue
    }

    $value = [Environment]::GetEnvironmentVariable($envName, "Process")
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($envName, "User")
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($envName, "Machine")
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $FallbackValue
    }

    return $value
}

function Get-JobCrawlerCredentialStatuses {
    param([AllowNull()]$SourcesConfig)

    $rows = New-Object System.Collections.Generic.List[object]
    $credentials = Get-ConfigProperty -Object $SourcesConfig -Name "credentials" -DefaultValue $null
    if ($null -eq $credentials) {
        return @()
    }

    foreach ($sourceName in @($credentials.PSObject.Properties.Name)) {
        $sourceCredentials = $credentials.$sourceName
        foreach ($credentialName in @($sourceCredentials.PSObject.Properties.Name)) {
            $envName = Get-ConfigProperty -Object $sourceCredentials.$credentialName -Name "env" -DefaultValue ""
            $defaultValue = Get-ConfigProperty -Object $sourceCredentials.$credentialName -Name "default" -DefaultValue ""
            $status = "missing"
            if (-not [string]::IsNullOrWhiteSpace($envName)) {
                $processValue = [Environment]::GetEnvironmentVariable($envName, "Process")
                $userValue = [Environment]::GetEnvironmentVariable($envName, "User")
                $machineValue = [Environment]::GetEnvironmentVariable($envName, "Machine")
                if (-not [string]::IsNullOrWhiteSpace($processValue) -or -not [string]::IsNullOrWhiteSpace($userValue) -or -not [string]::IsNullOrWhiteSpace($machineValue)) {
                    $status = "set"
                }
                elseif (-not [string]::IsNullOrWhiteSpace($defaultValue)) {
                    $status = "default"
                }
            }
            $rows.Add([PSCustomObject]@{
                Source = $sourceName
                Credential = $credentialName
                EnvironmentVariable = $envName
                Status = $status
            }) | Out-Null
        }
    }

    return @($rows.ToArray())
}

function Test-JobCrawlerConfig {
    param([AllowNull()]$Config)

    $issues = New-Object System.Collections.Generic.List[string]
    foreach ($mode in @("Fast", "Default", "Deep")) {
        $modeConfig = Get-ConfigPathValue -Object $Config.CrawlModes -Path ("modes.{0}" -f $mode) -DefaultValue $null
        if ($null -eq $modeConfig) {
            $issues.Add("Missing crawl mode: $mode") | Out-Null
        }
    }

    $minimumScore = Get-ConfigPathValue -Object $Config.MatchingRules -Path "thresholds.minimum_match_score" -DefaultValue $null
    if ($null -eq $minimumScore) {
        $issues.Add("Missing matching_rules.thresholds.minimum_match_score") | Out-Null
    }

    if (Test-JobCrawlerProfileConfigured -Config $Config) {
        $linkedInQueries = @(Get-JobCrawlerSourceQueryList -SourcesConfig $Config.Sources -SourceKey "linkedin" -FallbackKeys @("api"))
        if ($linkedInQueries.Count -eq 0) {
            $issues.Add("Missing sources.queries.linkedin") | Out-Null
        }

        $searchSourceKeys = @("hellowork", "apec", "france_travail", "adzuna")
        $searchSourceQueryCounts = @($searchSourceKeys | ForEach-Object {
            @(Get-JobCrawlerSourceQueryList -SourcesConfig $Config.Sources -SourceKey $_ -FallbackKeys @("api")).Count
        })
        if (@($searchSourceQueryCounts | Where-Object { $_ -gt 0 }).Count -eq 0) {
            $issues.Add("Missing source search queries. Configure sources.queries.api or per-source query lists.") | Out-Null
        }
    }

    $statusOptions = @(Get-ConfigStringArray (Get-ConfigPathValue -Object $Config.Workbook -Path "status_options" -DefaultValue @()))
    if ($statusOptions.Count -gt 0 -and "new" -notin $statusOptions) {
        $issues.Add("workbook.status_options must include 'new'") | Out-Null
    }

    $configuredSourceOrder = @(Get-JobCrawlerConfiguredSourceOrder -SourcesConfig $Config.Sources)
    if ($configuredSourceOrder.Count -eq 0) {
        $issues.Add("Missing sources metadata") | Out-Null
    }

    foreach ($sourceKey in $configuredSourceOrder) {
        $sourceConfig = Get-ConfigPathValue -Object $Config.Sources -Path ("sources.{0}" -f $sourceKey) -DefaultValue $null
        if ($null -eq $sourceConfig) {
            $issues.Add("Missing sources.$sourceKey metadata") | Out-Null
        }
    }

    foreach ($source in @(Get-JobCrawlerSourceDefinitions -SourcesConfig $Config.Sources)) {
        if ([string]::IsNullOrWhiteSpace([string]$source.CrawlFunction)) {
            $issues.Add("Missing crawl function for source '$($source.Key)'") | Out-Null
        }
        if ([string]::IsNullOrWhiteSpace([string]$source.Label)) {
            $issues.Add("Missing label for source '$($source.Key)'") | Out-Null
        }
    }

    return [PSCustomObject]@{
        IsValid = ($issues.Count -eq 0)
        Issues = @($issues.ToArray())
    }
}
