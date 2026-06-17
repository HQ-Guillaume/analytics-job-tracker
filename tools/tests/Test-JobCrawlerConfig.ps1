[CmdletBinding()]
param(
    [string]$ConfigDirectory = "config",
    [string]$Profile = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $projectRoot "app\core\JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\core\JobTracker.Config.ps1")

$configPath = Resolve-JobCrawlerPath -BasePath $projectRoot -Path $ConfigDirectory
$config = Get-JobCrawlerConfig -ConfigDirectory $configPath -ProfileId $Profile
$result = Test-JobCrawlerConfig -Config $config

if (-not $result.IsValid) {
    Write-Host "Config validation failed:"
    foreach ($issue in @($result.Issues)) {
        Write-Host "- $issue"
    }
    exit 1
}

Write-Host "Config validation passed."
Write-Host ("Config directory: {0}" -f $config.Root)
Write-Host ("Profile: {0} ({1})" -f $config.Profile.Label, $config.Profile.Id)
