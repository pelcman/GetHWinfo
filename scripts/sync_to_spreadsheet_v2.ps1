# Sync Hardware Info to Google Spreadsheet via Apps Script (V2 - Simplified)
# Based on working test_connection.ps1 approach

param(
    [Parameter(Mandatory=$false)]
    [string]$WebAppUrl = ""
)

# Configuration file path
$ConfigPath = Join-Path $PSScriptRoot "..\config.json"

# Load configuration if exists
if (Test-Path $ConfigPath) {
    $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    if ([string]::IsNullOrEmpty($WebAppUrl) -and $Config.WebAppUrl) {
        $WebAppUrl = $Config.WebAppUrl
    }
}

# Check parameters
if ([string]::IsNullOrEmpty($WebAppUrl)) {
    Write-Host "ERROR: WebAppUrl is required in config.json"
    exit 1
}

$CsvPath = Join-Path $PSScriptRoot "..\hardware_info.csv"

# Check if CSV exists
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found at: $CsvPath"
    exit 1
}

# Read and parse CSV
Write-Host "Reading CSV file: $CsvPath"
$CsvData = Import-Csv -Path $CsvPath -Encoding UTF8

if ($null -eq $CsvData) {
    Write-Host "ERROR: CSV file is empty or could not be parsed"
    exit 1
}

# Handle single record case
if ($CsvData -isnot [array]) {
    $CsvData = @($CsvData)
}

Write-Host "Found $($CsvData.Count) record(s) in CSV"

# Prepare request body (same format as test_connection.ps1)
$requestBody = @{
    data = $CsvData
} | ConvertTo-Json -Depth 10

Write-Host ""
Write-Host "Sending data to Google Spreadsheet..."
Write-Host "Target URL: $WebAppUrl"
Write-Host "Request body size: $($requestBody.Length) bytes"
Write-Host ""

try {
    # Disable progress bar for better performance
    $ProgressPreference = 'SilentlyContinue'

    Write-Host "Sending request..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $response = Invoke-WebRequest -Uri $WebAppUrl -Method Post -Body $requestBody -ContentType "application/json; charset=utf-8" -TimeoutSec 90 -UseBasicParsing

    $stopwatch.Stop()
    Write-Host "Response received in $($stopwatch.ElapsedMilliseconds)ms (Status: $($response.StatusCode))"

    # Parse JSON response
    $response = $response.Content | ConvertFrom-Json

    if ($response.status -eq "success") {
        Write-Host "SUCCESS: Data synchronized successfully!"
        Write-Host ""
        Write-Host "Updated: $($response.updated) record(s)"
        Write-Host "Added: $($response.added) record(s)"
        Write-Host "Total: $($response.total) record(s) in spreadsheet"

        if ($response.spreadsheetUrl) {
            Write-Host ""
            Write-Host "Spreadsheet URL: $($response.spreadsheetUrl)"
        }
    } else {
        Write-Host "ERROR: $($response.message)"
        exit 1
    }
} catch {
    Write-Host "ERROR: Failed to send data to Web App"
    Write-Host $_.Exception.Message

    if ($_.ErrorDetails.Message) {
        Write-Host ""
        Write-Host "Details:"
        Write-Host $_.ErrorDetails.Message
    }

    exit 1
}

Write-Host ""
Write-Host "Sync completed successfully!"
