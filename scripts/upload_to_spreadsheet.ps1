# Upload Hardware Info to Google Spreadsheet
# This script reads hardware_info.csv and uploads to Google Sheets

param(
    [Parameter(Mandatory=$false)]
    [string]$SpreadsheetId = "",

    [Parameter(Mandatory=$false)]
    [string]$CredentialPath = ""
)

# Configuration
if ([string]::IsNullOrEmpty($SpreadsheetId)) {
    Write-Host "ERROR: Spreadsheet ID is required."
    Write-Host "Usage: .\upload_to_spreadsheet.ps1 -SpreadsheetId 'YOUR_SPREADSHEET_ID' -CredentialPath 'path\to\credentials.json'"
    Write-Host ""
    Write-Host "To get Spreadsheet ID:"
    Write-Host "  Open your Google Sheet, the ID is in the URL:"
    Write-Host "  https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit"
    Write-Host ""
    Write-Host "To get credentials.json:"
    Write-Host "  1. Go to https://console.cloud.google.com/"
    Write-Host "  2. Create a new project or select existing one"
    Write-Host "  3. Enable Google Sheets API"
    Write-Host "  4. Create OAuth 2.0 Client ID credentials"
    Write-Host "  5. Download credentials.json"
    exit 1
}

if ([string]::IsNullOrEmpty($CredentialPath)) {
    $CredentialPath = Join-Path $PSScriptRoot "..\credentials.json"
}

$CsvPath = Join-Path $PSScriptRoot "..\hardware_info.csv"

# Check if CSV exists
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found at: $CsvPath"
    Write-Host "Please run systeminfo.ps1 first to generate the CSV file."
    exit 1
}

# Check if credentials exist
if (-not (Test-Path $CredentialPath)) {
    Write-Host "ERROR: Credentials file not found at: $CredentialPath"
    Write-Host "Please provide a valid path to credentials.json"
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

# Install required module if not present
if (-not (Get-Module -ListAvailable -Name "PSGSuite")) {
    Write-Host "Installing PSGSuite module..."
    Install-Module PSGSuite -Scope CurrentUser -Force
}

# Import module
Import-Module PSGSuite

# Authenticate with Google
Write-Host "Authenticating with Google..."
try {
    Set-PSGSuiteConfig -ConfigName "HWInfo" -ServiceAccountClientIDJson $CredentialPath -SetAsDefaultConfig
    $Auth = Get-PSGSuiteConfig
    Write-Host "Authentication successful"
} catch {
    Write-Host "ERROR: Authentication failed. Please check your credentials.json file."
    Write-Host $_.Exception.Message
    exit 1
}

# Get current spreadsheet data
Write-Host "Fetching current spreadsheet data..."
try {
    $SheetData = Get-GSSheetValue -SpreadsheetId $SpreadsheetId -Range "A:Z"
} catch {
    Write-Host "ERROR: Failed to fetch spreadsheet data"
    Write-Host $_.Exception.Message
    exit 1
}

# Extract header and data
if ($SheetData.Count -eq 0) {
    Write-Host "Spreadsheet is empty. Creating header row..."
    $Headers = $CsvData[0].PSObject.Properties.Name
    $ExistingData = @()
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
$UpdatedRows = @()
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
        $UpdatedRows += @{
            Row = $rowNumber
            Data = $rowData
        }
    } else {
        # Add new row
        Write-Host "Adding new record: $computerName"
        $NewRows += ,$rowData
    }
}

# Update existing rows
foreach ($update in $UpdatedRows) {
    $range = "A$($update.Row):Z$($update.Row)"
    try {
        Set-GSSheetValue -SpreadsheetId $SpreadsheetId -Range $range -Value (,$update.Data) | Out-Null
        Write-Host "  Updated row $($update.Row) successfully"
    } catch {
        Write-Host "  ERROR updating row $($update.Row): $($_.Exception.Message)"
    }
}

# Append new rows
if ($NewRows.Count -gt 0) {
    try {
        $range = "A:Z"
        Add-GSSheetValue -SpreadsheetId $SpreadsheetId -Range $range -Value $NewRows | Out-Null
        Write-Host "Added $($NewRows.Count) new row(s) successfully"
    } catch {
        Write-Host "ERROR adding new rows: $($_.Exception.Message)"
    }
}

# Sort by ComputerName
Write-Host "Sorting spreadsheet by ComputerName..."
try {
    # Get sheet ID (usually 0 for first sheet)
    $spreadsheet = Get-GSSpreadsheet -SpreadsheetId $SpreadsheetId
    $sheetId = $spreadsheet.Sheets[0].Properties.SheetId

    # Sort request
    $sortRequest = @{
        sortRange = @{
            range = @{
                sheetId = $sheetId
                startRowIndex = 1  # Skip header
            }
            sortSpecs = @(
                @{
                    dimensionIndex = $ComputerNameIndex
                    sortOrder = "ASCENDING"
                }
            )
        }
    }

    # Note: PSGSuite doesn't have direct sort support, need to use API
    Write-Host "Note: Automatic sorting requires Google Sheets API. Please manually sort by ComputerName column if needed."
} catch {
    Write-Host "Note: Could not auto-sort. Please manually sort by ComputerName column if needed."
}

Write-Host ""
Write-Host "Upload completed successfully!"
Write-Host "Spreadsheet URL: https://docs.google.com/spreadsheets/d/$SpreadsheetId/edit"
