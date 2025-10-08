# Sync Hardware Info to Google Spreadsheet via Apps Script (using curl)
# Alternative version using curl instead of Invoke-RestMethod

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

# Read CSV with proper handling of multi-line fields
Write-Host "Reading CSV file: $CsvPath"

# Read raw CSV content
$csvContent = Get-Content -Path $CsvPath -Raw -Encoding UTF8

# Remove BOM if present
if ($csvContent.Length -gt 0 -and $csvContent[0] -eq [char]0xFEFF) {
    $csvContent = $csvContent.Substring(1)
}

# Parse CSV properly
try {
    $CsvData = $csvContent | ConvertFrom-Csv
} catch {
    Write-Host "ERROR: Failed to parse CSV file"
    Write-Host $_.Exception.Message
    exit 1
}

if ($null -eq $CsvData) {
    Write-Host "ERROR: CSV file is empty"
    exit 1
}

# Handle single record case
if ($CsvData -isnot [array]) {
    $CsvData = @($CsvData)
}

Write-Host "Found $($CsvData.Count) record(s) in CSV"

# Prepare request body
$body = @{
    data = $CsvData
} | ConvertTo-Json -Depth 10 -Compress

# Save to temp file
$tempFile = Join-Path $env:TEMP "hw_sync_$(Get-Date -Format 'yyyyMMddHHmmss').json"
$body | Out-File -FilePath $tempFile -Encoding UTF8 -NoNewline

Write-Host ""
Write-Host "Sending data to Google Spreadsheet..."
Write-Host "Target URL: $WebAppUrl"
Write-Host "Data size: $($body.Length) bytes"
Write-Host ""

# Use curl to send the request
try {
    $curlOutput = & curl.exe -s -L -X POST `
        -H "Content-Type: application/json; charset=utf-8" `
        -d "@$tempFile" `
        --max-time 60 `
        "$WebAppUrl" 2>&1

    # Clean up temp file
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue

    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: curl failed with exit code $LASTEXITCODE"
        Write-Host $curlOutput
        exit 1
    }

    Write-Host "Response received from server"

    # Parse response
    try {
        $response = $curlOutput | ConvertFrom-Json
    } catch {
        Write-Host "ERROR: Failed to parse response"
        Write-Host "Raw response:"
        Write-Host $curlOutput
        exit 1
    }

    if ($response.status -eq "success") {
        Write-Host ""
        Write-Host "SUCCESS: Data synchronized successfully!"
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
    Write-Host ""
    Write-Host "ERROR: Failed to send data"
    Write-Host $_.Exception.Message

    # Clean up temp file
    Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    exit 1
}

Write-Host ""
Write-Host "Sync completed successfully!"
