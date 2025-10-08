# Hardware Information Collection System

Collect hardware information and sync to Google Spreadsheet automatically.

## Files

- `run.bat` - Run hardware info collection
- `scripts/systeminfo.ps1` - Collect hardware information to CSV
- `scripts/sync_to_spreadsheet.ps1` - Sync CSV to Google Spreadsheet (via Apps Script)
- `AppsScript.gs` - Google Apps Script code for spreadsheet integration
- `config.json` - Configuration file (create from config.json.example)

## Quick Setup Guide

### Step 1: Create Google Spreadsheet

1. Create a new Google Spreadsheet
2. Note the spreadsheet URL

### Step 2: Deploy Apps Script Web App

1. Open your Google Spreadsheet
2. Go to **Extensions** > **Apps Script**
3. Delete any default code
4. Copy all content from `AppsScript.gs` and paste it
5. Click **Deploy** > **New deployment**
6. Settings:
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Click **Deploy**
8. **Copy the Web App URL** (looks like: `https://script.google.com/macros/s/.../exec`)
9. Click **Done**

### Step 3: Create Configuration File

1. Copy `config.json.example` to `config.json`
2. Edit `config.json` and paste your Web App URL:

```json
{
  "WebAppUrl": "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"
}
```

### Step 4: Run the Scripts

#### Collect hardware info only:
```
run.bat
```
This creates `hardware_info.csv`

#### Collect and sync to Google Spreadsheet:
```powershell
.\scripts\systeminfo.ps1
.\scripts\sync_to_spreadsheet.ps1
```

Or create a batch file to do both:
```batch
@echo off
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\systeminfo.ps1"
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\sync_to_spreadsheet.ps1"
pause
```

## Features

### Hardware Information Collected

- **System**: Computer Name, Manufacturer, Model, Serial Number
- **Users**: All local users (excluding system accounts), Primary user, Current user
- **Network**: IP Address, MAC Address, Network Adapters
- **OS**: Name, Version, Architecture
- **CPU**: Name, Cores, Logical Processors
- **Memory**: Total capacity, per-slot details with location
- **GPU**: Graphics cards with VRAM
- **Storage**: All disks with model and capacity
- **Motherboard**: Manufacturer and model
- **Timestamp**: Last updated date/time

### Google Spreadsheet Integration

- **Auto-update**: Existing records (matched by ComputerName) are updated
- **Auto-add**: New computers are added automatically
- **Auto-sort**: Records sorted by ComputerName
- **Auto-format**: Headers formatted, columns auto-resized
- **Multi-line support**: Cell values with line breaks (RAM slots, users, etc.)

## Troubleshooting

### "Authorization required" error
- Open the Apps Script editor
- Run the `testUpdateSpreadsheet()` function manually
- Grant permissions when prompted

### "Web App not found" error
- Check the Web App URL in config.json
- Make sure deployment is set to "Anyone" access
- Try redeploying the Web App

### CSV encoding issues
- All files use UTF-8 encoding
- No Japanese characters in the scripts to avoid encoding problems

### PowerShell execution policy error
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Advanced Usage

### Command-line parameters

```powershell
# Specify Web App URL directly
.\scripts\sync_to_spreadsheet.ps1 -WebAppUrl "YOUR_URL"
```

### Testing Apps Script

In Apps Script editor, run `testUpdateSpreadsheet()` to test with sample data.

### Customizing sheet name

Edit `AppsScript.gs` line 18:
```javascript
const SHEET_NAME = 'Sheet1'; // Change to your preferred name
```

### Scheduled automated sync

Use Windows Task Scheduler to run the scripts periodically:
1. Create a new task
2. Trigger: Daily or weekly
3. Action: Run `run_and_sync.bat` (create this batch file)

## CSV Format

The CSV file contains one row per computer with the following columns:
- LastUpdated
- ComputerName
- AllRegisteredUsers
- PrimaryUser
- CurrentUser
- IPAddress
- MACAddress
- Manufacturer
- Model
- SerialNumber
- OSName
- OSVersion
- OSArchitecture
- CPU
- CPUCores
- CPULogicalProcessors
- TotalMemoryGB
- RAMSlots
- MemoryDetails
- GPU
- Disk
- MotherboardManufacturer
- MotherboardProduct
- NetworkAdapter

Multi-value fields use line breaks (`\n`) as separators.
