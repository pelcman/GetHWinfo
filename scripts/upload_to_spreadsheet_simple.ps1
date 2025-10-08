# Upload Hardware Info to Google Spreadsheet (Simple Version using REST API)
# This script reads hardware_info.csv and uploads to Google Sheets using REST API

param(
    [Parameter(Mandatory=$false)]
    [string]$SpreadsheetId = "",

    [Parameter(Mandatory=$false)]
    [string]$ApiKey = ""
)

# Configuration file path
$ConfigPath = Join-Path $PSScriptRoot "..\config.json"

# Load configuration if exists
if (Test-Path $ConfigPath) {
    $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    if ([string]::IsNullOrEmpty($SpreadsheetId) -and $Config.SpreadsheetId) {
        $SpreadsheetId = $Config.SpreadsheetId
    }
    if ([string]::IsNullOrEmpty($ApiKey) -and $Config.ApiKey) {
        $ApiKey = $Config.ApiKey
    }
}

# Check parameters
if ([string]::IsNullOrEmpty($SpreadsheetId) -or [string]::IsNullOrEmpty($ApiKey)) {
    Write-Host "ERROR: SpreadsheetId and ApiKey are required."
    Write-Host ""
    Write-Host "Usage:"
    Write-Host "  .\upload_to_spreadsheet_simple.ps1 -SpreadsheetId 'YOUR_SPREADSHEET_ID' -ApiKey 'YOUR_API_KEY'"
    Write-Host ""
    Write-Host "Or create config.json in the root directory:"
    Write-Host '  {'
    Write-Host '    "SpreadsheetId": "YOUR_SPREADSHEET_ID",'
    Write-Host '    "ApiKey": "YOUR_API_KEY"'
    Write-Host '  }'
    Write-Host ""
    Write-Host "Setup Instructions:"
    Write-Host "1. Open/Create Google Spreadsheet"
    Write-Host "   - Get Spreadsheet ID from URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit"
    Write-Host ""
    Write-Host "2. Set up Google Cloud Project:"
    Write-Host "   - Go to https://console.cloud.google.com/"
    Write-Host "   - Create new project or select existing"
    Write-Host "   - Enable 'Google Sheets API'"
    Write-Host "   - Go to 'Credentials' > Create Credentials > API Key"
    Write-Host "   - Copy the API Key"
    Write-Host ""
    Write-Host "3. Share your spreadsheet:"
    Write-Host "   - Open the spreadsheet"
    Write-Host "   - Click Share button"
    Write-Host "   - Change to 'Anyone with the link can edit'"
    Write-Host "   - Or add your service account email"
    exit 1
}

$CsvPath = Join-Path $PSScriptRoot "..\hardware_info.csv"

# Check if CSV exists
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found at: $CsvPath"
    Write-Host "Please run systeminfo.ps1 first to generate the CSV file."
    exit 1
}

# Read CSV
Write-Host "Reading CSV file: $CsvPath"
$CsvData = Import-Csv -Path $CsvPath -Encoding UTF8

if ($CsvData.Count -eq 0) {
    Write-Host "ERROR: CSV file is empty"
    exit 1
}

Write-Host "Found $($CsvData.Count) record(s) in CSV"

# Google Sheets API endpoint
$BaseUrl = "https://sheets.googleapis.com/v4/spreadsheets"

# Function to get spreadsheet data
function Get-SpreadsheetData {
    param($SpreadsheetId, $ApiKey)

    $url = "$BaseUrl/${SpreadsheetId}/values/A:Z?key=$ApiKey"
    try {
        $response = Invoke-RestMethod -Uri $url -Method Get
        return $response.values
    } catch {
        Write-Host "ERROR: Failed to fetch spreadsheet data"
        Write-Host $_.Exception.Message
        return $null
    }
}

# Function to update spreadsheet row
function Update-SpreadsheetRow {
    param($SpreadsheetId, $ApiKey, $Range, $Values)

    $url = "$BaseUrl/${SpreadsheetId}/values/${Range}?valueInputOption=RAW&key=$ApiKey"
    $body = @{
        values = @(,$Values)
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $url -Method Put -Body $body -ContentType "application/json"
        return $true
    } catch {
        Write-Host "ERROR: Failed to update row"
        Write-Host $_.Exception.Message
        return $false
    }
}

# Function to append spreadsheet rows
function Append-SpreadsheetRows {
    param($SpreadsheetId, $ApiKey, $Values)

    $url = "$BaseUrl/${SpreadsheetId}/values/A:Z:append?valueInputOption=RAW&key=$ApiKey"
    $body = @{
        values = $Values
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
        return $true
    } catch {
        Write-Host "ERROR: Failed to append rows"
        Write-Host $_.Exception.Message
        return $false
    }
}

# Get current spreadsheet data
Write-Host "Fetching current spreadsheet data..."
$SheetData = Get-SpreadsheetData -SpreadsheetId $SpreadsheetId -ApiKey $ApiKey

if ($null -eq $SheetData) {
    exit 1
}

# Extract header and data
if ($SheetData.Count -eq 0) {
    Write-Host "Spreadsheet is empty. Creating header row..."
    $Headers = $CsvData[0].PSObject.Properties.Name
    $ExistingData = @()

    # Write header
    $headerArray = @($Headers)
    Update-SpreadsheetRow -SpreadsheetId $SpreadsheetId -ApiKey $ApiKey -Range "A1:Z1" -Values $headerArray
} else {
    $Headers = $SheetData[0]
    $ExistingData = $SheetData | Select-Object -Skip 1
}

Write-Host "Headers: $($Headers -join ', ')"

# Create a hashtable for existing data (keyed by ComputerName)
$ExistingRecords = @{}
$ComputerNameIndex = [array]::IndexOf($Headers, "ComputerName")

if ($ComputerNameIndex -ge 0) {
    for ($i = 0; $i -lt $ExistingData.Count; $i++) {
        $row = $ExistingData[$i]
        if ($row.Count -gt $ComputerNameIndex) {
            $computerName = $row[$ComputerNameIndex]
            if (-not [string]::IsNullOrEmpty($computerName)) {
                $ExistingRecords[$computerName] = $i + 2  # +2 because: +1 for 1-based index, +1 for header
            }
        }
    }
}

Write-Host "Found $($ExistingRecords.Count) existing records in spreadsheet"

# Process each CSV record
$UpdatedCount = 0
$NewRows = @()

foreach ($record in $CsvData) {
    $computerName = $record.ComputerName

    # Convert record to array matching header order
    $rowData = @()
    foreach ($header in $Headers) {
        $value = $record.$header
        if ($null -eq $value) { $value = "" }
        $rowData += $value
    }

    if ($ExistingRecords.ContainsKey($computerName)) {
        # Update existing row
        $rowNumber = $ExistingRecords[$computerName]
        Write-Host "Updating existing record: $computerName (Row $rowNumber)"

        $range = "A${rowNumber}:Z${rowNumber}"
        if (Update-SpreadsheetRow -SpreadsheetId $SpreadsheetId -ApiKey $ApiKey -Range $range -Values $rowData) {
            $UpdatedCount++
        }

        Start-Sleep -Milliseconds 100  # Rate limiting
    } else {
        # Add new row
        Write-Host "Adding new record: $computerName"
        $NewRows += ,$rowData
    }
}

Write-Host "Updated $UpdatedCount existing record(s)"

# Append new rows
if ($NewRows.Count -gt 0) {
    Write-Host "Appending $($NewRows.Count) new record(s)..."
    if (Append-SpreadsheetRows -SpreadsheetId $SpreadsheetId -ApiKey $ApiKey -Values $NewRows) {
        Write-Host "Successfully added $($NewRows.Count) new record(s)"
    }
}

Write-Host ""
Write-Host "Upload completed successfully!"
Write-Host "Spreadsheet URL: https://docs.google.com/spreadsheets/d/$SpreadsheetId/edit"
Write-Host ""
Write-Host "Note: To sort by ComputerName, please:"
Write-Host "1. Open the spreadsheet"
Write-Host "2. Select all data (Click on column A header)"
Write-Host "3. Go to Data > Sort sheet by column A (ComputerName)"
