# Sync Hardware Info to Google Spreadsheet via Apps Script
# Simple version - just POST to Web App URL

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
    Write-Host "ERROR: WebAppUrl is required."
    Write-Host ""
    Write-Host "Usage:"
    Write-Host "  .\sync_to_spreadsheet.ps1 -WebAppUrl 'YOUR_WEB_APP_URL'"
    Write-Host ""
    Write-Host "Or create config.json in the root directory:"
    Write-Host '  {'
    Write-Host '    "WebAppUrl": "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"'
    Write-Host '  }'
    Write-Host ""
    Write-Host "Setup Instructions:"
    Write-Host "1. Open your Google Spreadsheet"
    Write-Host "2. Go to Extensions > Apps Script"
    Write-Host "3. Copy the code from 'AppsScript.gs' file"
    Write-Host "4. Deploy as Web App:"
    Write-Host "   - Click Deploy > New deployment"
    Write-Host "   - Type: Web app"
    Write-Host "   - Execute as: Me"
    Write-Host "   - Who has access: Anyone"
    Write-Host "   - Click Deploy"
    Write-Host "5. Copy the Web App URL"
    exit 1
}

$CsvPath = Join-Path $PSScriptRoot "..\hardware_info.csv"

# Check if CSV exists
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found at: $CsvPath"
    Write-Host "Please run systeminfo.ps1 first to generate the CSV file."
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
    Write-Host "ERROR: CSV file is empty or could not be parsed"
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

# Send POST request to Web App
Write-Host ""
Write-Host "Sending data to Google Spreadsheet..."
Write-Host "Target URL: $WebAppUrl"
Write-Host "Data size: $($body.Length) bytes"
Write-Host ""

try {
    # Disable progress bar to speed up the request
    $ProgressPreference = 'SilentlyContinue'

    # Use Invoke-WebRequest with explicit parameters
    $webResponse = Invoke-WebRequest -Uri $WebAppUrl -Method Post -Body $body -ContentType "application/json; charset=utf-8" -TimeoutSec 120 -UseBasicParsing

    Write-Host "Response received from server (Status: $($webResponse.StatusCode))"

    # Parse JSON response
    $response = $webResponse.Content | ConvertFrom-Json

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
