[CmdletBinding()]
param(
    [string]$ConfigDirectory = "config"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $projectRoot "JobTracker.Common.ps1")
. (Join-Path $projectRoot "app\JobTracker.Config.ps1")

$configPath = Resolve-JobCrawlerPath -BasePath $projectRoot -Path $ConfigDirectory
$config = Get-JobCrawlerConfig -ConfigDirectory $configPath
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
