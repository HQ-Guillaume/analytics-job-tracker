[CmdletBinding()]
param(
    [string]$ProjectRoot = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$resolvedRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$issues = New-Object System.Collections.Generic.List[string]

function Add-ReleaseIssue {
    param([string]$Message)
    $issues.Add($Message) | Out-Null
}

function Get-TrackedFiles {
    $gitFiles = @()
    try {
        $raw = @()
        $raw += git -C $resolvedRoot ls-files
        $raw += git -C $resolvedRoot ls-files --others --exclude-standard
        if ($LASTEXITCODE -eq 0) {
            $gitFiles = @($raw | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
    }
    catch {
        $gitFiles = @()
    }

    if ($gitFiles.Count -gt 0) {
        return @($gitFiles | Where-Object { Test-Path -LiteralPath (Join-Path $resolvedRoot $_) })
    }

    return @(Get-ChildItem -Path $resolvedRoot -File -Recurse -Force |
        Where-Object { $_.FullName -notmatch "\\.git\\" } |
        ForEach-Object { [IO.Path]::GetRelativePath($resolvedRoot, $_.FullName) })
}

function Test-IsTextFile {
    param([string]$Path)

    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    return $extension -in @(".ps1", ".psm1", ".psd1", ".cmd", ".bat", ".md", ".json", ".yml", ".yaml", ".txt", ".html", ".css", ".js", ".xml", ".gitignore", ".gitattributes")
}

$trackedFiles = @(Get-TrackedFiles)
$normalizedFiles = @($trackedFiles | ForEach-Object { ([string]$_).Replace("\", "/") })

foreach ($file in $normalizedFiles) {
    if ($file -match "(^|/)output/") {
        Add-ReleaseIssue "Tracked output file is not public-safe: $file"
    }
    if ($file -match "(^|/)config/local(\.|/)") {
        Add-ReleaseIssue "Tracked local config override is not public-safe: $file"
    }
    if ($file -match "(^|/)config/profiles/[^/]+\.json$") {
        Add-ReleaseIssue "Tracked job-search profile is not public-safe: $file"
    }
    if ($file -match "\.(xlsx|xlsm|xls|pdf|env|key|secret|pfx|pem|p12|log|tmp|zip)$") {
        Add-ReleaseIssue "Tracked file type should not be in a clean release: $file"
    }
    if ($file -match "(?i)(cv_|resume|curriculum)") {
        Add-ReleaseIssue "Tracked CV/resume-like file should not be public: $file"
    }
}

$windowsUserPathPattern = "c:" + "\\users" + "\\[^\\\s]+"
$personalNamePattern = ("\b{0}\b|\b{1}\b|\b{2}\b|{3}" -f ("gui" + "llaume"), ("hai" + "qi"), ("ge" + "ng"), ("cv_" + "hai" + "qi"))
$personalPattern = "(?i){0}|{1}" -f $windowsUserPathPattern, $personalNamePattern
$credentialPattern = "(?i)PAR_[A-Za-z0-9_]{30,}|(?<![A-Za-z0-9])[a-f0-9]{64}(?![A-Za-z0-9])|(?i)(client_secret|app_key|api_key|token)\\s*[:=]\\s*['""][^'""]{12,}['""]"
$profileLeakPattern = "(?i)digital\s+analytics|digital\s+analyst|web\s+analyst|google\s+analytics|google\s+tag\s+manager|piano\s+analytics|contentsquare|tag\s+commander|commanders?\s+act|\btealium\b|broad\s+analyst|not_analytics_enough|web_analytics_tools|digital_analytics_title|digital_analytics_self_test"

foreach ($relativeFile in $trackedFiles) {
    $path = Join-Path $resolvedRoot $relativeFile
    if (-not (Test-Path -LiteralPath $path) -or -not (Test-IsTextFile -Path $path)) {
        continue
    }

    $content = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if ($content -match $personalPattern) {
        Add-ReleaseIssue "Personal path/name detected in tracked file: $relativeFile"
    }
    if ($content -match $credentialPattern) {
        Add-ReleaseIssue "Credential-looking value detected in tracked file: $relativeFile"
    }

    $publicRuntimeFile = ([string]$relativeFile).Replace("\", "/") -match "^(app|config)/|^README\.md$"
    if ($publicRuntimeFile -and $content -match $profileLeakPattern) {
        Add-ReleaseIssue "Digital-analytics profile residue detected in public runtime file: $relativeFile"
    }
}

if ($issues.Count -gt 0) {
    Write-Host "Release safety check failed:"
    foreach ($issue in @($issues.ToArray())) {
        Write-Host "- $issue"
    }
    exit 1
}

Write-Host "Release safety check passed."
Write-Host ("Tracked files checked: {0}" -f $trackedFiles.Count)
