[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$ProjectRoot = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$resolvedRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
& (Join-Path $PSScriptRoot "Test-ReleaseSafety.ps1") -ProjectRoot $resolvedRoot

$distDirectory = Join-Path $resolvedRoot "dist"
if (-not (Test-Path -LiteralPath $distDirectory)) {
    New-Item -ItemType Directory -Path $distDirectory -Force | Out-Null
}

$safeVersion = ($Version -replace "[^A-Za-z0-9._-]", "-").Trim("-")
if ([string]::IsNullOrWhiteSpace($safeVersion)) {
    throw "Version cannot be empty."
}

$packagePath = Join-Path $distDirectory ("custom-job-tracker-{0}.zip" -f $safeVersion)
if (Test-Path -LiteralPath $packagePath) {
    Remove-Item -LiteralPath $packagePath -Force
}

git -C $resolvedRoot archive --format zip --worktree-attributes --output $packagePath HEAD
if ($LASTEXITCODE -ne 0) {
    throw "git archive failed."
}

Write-Host ("Created public release package: {0}" -f $packagePath)
