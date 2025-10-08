# Test Google Apps Script Web App Connection
# This script tests if the Web App is accessible

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
    Write-Host "Please set WebAppUrl in config.json or provide it as a parameter."
    exit 1
}

Write-Host "Testing connection to Google Apps Script Web App..."
Write-Host "URL: $WebAppUrl"
Write-Host ""

# Test 1: GET request (should return status message)
Write-Host "[Test 1] Sending GET request..."
try {
    $getResponse = Invoke-RestMethod -Uri $WebAppUrl -Method Get -TimeoutSec 10
    Write-Host "SUCCESS: Web App is accessible"
    Write-Host "Response:"
    $getResponse | ConvertTo-Json | Write-Host
} catch {
    Write-Host "FAILED: Cannot access Web App"
    Write-Host "Error: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Troubleshooting:"
    Write-Host "1. Check if the URL is correct"
    Write-Host "2. Make sure the Web App is deployed"
    Write-Host "3. Verify 'Who has access' is set to 'Anyone'"
    Write-Host "4. Try opening the URL in a browser"
    exit 1
}

Write-Host ""

# Test 2: POST request with minimal data
Write-Host "[Test 2] Sending POST request with test data..."
$testData = @{
    data = @(
        @{
            LastUpdated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            ComputerName = "TEST-CONNECTION"
            AllRegisteredUsers = "TestUser"
            PrimaryUser = "TestUser"
            CurrentUser = "TestUser"
            IPAddress = "192.168.1.1"
            MACAddress = "00-00-00-00-00-00"
            Manufacturer = "Test"
            Model = "Test"
            SerialNumber = "TEST123"
            OSName = "Windows Test"
            OSVersion = "10.0"
            OSArchitecture = "64-bit"
            CPU = "Test CPU"
            CPUCores = "4"
            CPULogicalProcessors = "8"
            TotalMemoryGB = "16"
            RAMSlots = "Test Slot"
            MemoryDetails = "Test Memory"
            GPU = "Test GPU"
            Disk = "Test Disk"
            MotherboardManufacturer = "Test Board"
            MotherboardProduct = "Test Product"
            NetworkAdapter = "Test Adapter"
        }
    )
} | ConvertTo-Json -Depth 10

try {
    $postResponse = Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $testData -ContentType "application/json; charset=utf-8" -TimeoutSec 30

    if ($postResponse.status -eq "success") {
        Write-Host "SUCCESS: Data sent and processed successfully"
        Write-Host "Response:"
        $postResponse | ConvertTo-Json | Write-Host
        Write-Host ""
        Write-Host "Spreadsheet URL: $($postResponse.spreadsheetUrl)"
        Write-Host ""
        Write-Host "Connection test completed successfully!"
        Write-Host "You can now use sync_to_spreadsheet.ps1 to sync your hardware info."
    } else {
        Write-Host "WARNING: Request succeeded but status is not 'success'"
        Write-Host "Response:"
        $postResponse | ConvertTo-Json | Write-Host
    }
} catch {
    Write-Host "FAILED: Error sending data"
    Write-Host "Error: $($_.Exception.Message)"

    if ($_.ErrorDetails.Message) {
        Write-Host ""
        Write-Host "Error Details:"
        Write-Host $_.ErrorDetails.Message
    }

    Write-Host ""
    Write-Host "Troubleshooting:"
    Write-Host "1. Check Apps Script logs (View > Executions in Apps Script editor)"
    Write-Host "2. Make sure you authorized the script (run testUpdateSpreadsheet in editor)"
    Write-Host "3. Check if the spreadsheet has proper permissions"
    exit 1
}

Write-Host ""
Write-Host "All tests passed!"
