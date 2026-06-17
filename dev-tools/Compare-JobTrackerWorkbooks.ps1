[CmdletBinding()]
param(
    [string]$BaselinePath = "",
    [Parameter(Mandatory = $true)]
    [string]$CandidatePath,
    [switch]$IncludeApplicationRows
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $projectRoot "JobTracker.Common.ps1")

if ([string]::IsNullOrWhiteSpace($BaselinePath)) {
    $BaselinePath = Join-Path $projectRoot "output\jobs_tracker.xlsx"
}

function Test-IsApplicationStatus {
    param([AllowNull()][string]$Status)

    return ([string]$Status).Trim().ToLowerInvariant() -in @("applied", "interview", "offer", "rejected", "withdrawn")
}

function Get-CellTextByName {
    param(
        $Sheet,
        [hashtable]$Headers,
        [int]$Row,
        [string]$Name
    )

    if (-not $Headers.ContainsKey($Name)) {
        return ""
    }

    return [string]$Sheet.Cells.Item($Row, [int]$Headers[$Name]).Text
}

function Import-JobWorkbookRows {
    param([string]$Path)

    $excel = $null
    $workbook = $null
    $sheet = $null
    $usedRange = $null

    try {
        $fullPath = (Resolve-Path -LiteralPath $Path).Path
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($fullPath, 0, $true)
        $sheet = $workbook.Worksheets.Item("Jobs")
        $usedRange = $sheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        $headers = Get-WorksheetHeaderMap -Sheet $sheet -ColumnCount $columnCount

        $rows = New-Object System.Collections.Generic.List[object]
        for ($row = 2; $row -le $rowCount; $row++) {
            $jobTitle = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "job_title"
            $jobId = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "job_id"
            if ([string]::IsNullOrWhiteSpace($jobTitle) -and [string]::IsNullOrWhiteSpace($jobId)) {
                continue
            }

            $rows.Add([PSCustomObject]@{
                job_id = $jobId
                status = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "status"
                job_title = $jobTitle
                company_name = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "company_name"
                location = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "location"
                platform = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "platform"
                published_date = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "published_date"
                match_level = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "match_level"
                seen_in_current_crawl = Get-CellTextByName -Sheet $sheet -Headers $headers -Row $row -Name "seen_in_current_crawl"
            }) | Out-Null
        }

        return $rows.ToArray()
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
        if ($null -ne $excel) { $excel.Quit() | Out-Null }
        Release-ComObject $usedRange
        Release-ComObject $sheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$baselineRows = @(Import-JobWorkbookRows -Path $BaselinePath)
$candidateRows = @(Import-JobWorkbookRows -Path $CandidatePath)

if (-not $IncludeApplicationRows) {
    $baselineRows = @($baselineRows | Where-Object { -not (Test-IsApplicationStatus $_.status) })
    $candidateRows = @($candidateRows | Where-Object { -not (Test-IsApplicationStatus $_.status) })
}

$candidateById = @{}
foreach ($row in $candidateRows) {
    if (-not [string]::IsNullOrWhiteSpace($row.job_id)) {
        $candidateById[$row.job_id] = $true
    }
}

$missingRows = @($baselineRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.job_id) -and -not $candidateById.ContainsKey($_.job_id) })
$interestingMissingRows = @($missingRows | Where-Object { $_.status -in @("interesting", "applied", "interview", "offer") -or $_.match_level -eq "High" })

Write-Host ("Baseline rows compared: {0}" -f $baselineRows.Count)
Write-Host ("Candidate rows compared: {0}" -f $candidateRows.Count)
Write-Host ("Rows in baseline but not candidate: {0}" -f $missingRows.Count)
Write-Host ("High/interesting/application-like missing rows: {0}" -f $interestingMissingRows.Count)

if ($missingRows.Count -gt 0) {
    $missingRows |
        Select-Object status, match_level, published_date, job_title, company_name, location, platform, job_id |
        Sort-Object @{ Expression = "match_level"; Descending = $true }, published_date, company_name |
        Format-Table -AutoSize
}
