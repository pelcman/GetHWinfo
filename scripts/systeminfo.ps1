# Hardware Information Collection Script
# Output: 1 row per PC in CSV format

# Last updated timestamp
$LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# System information
$SystemInfo = Get-WmiObject Win32_ComputerSystem
$OSInfo = Get-WmiObject Win32_OperatingSystem
$BiosInfo = Get-WmiObject Win32_BIOS

# Computer name
$ComputerName = $env:COMPUTERNAME

# All registered users on this PC (exclude system accounts)
$ExcludedUsers = @("DefaultAccount", "Guest", "WDAGUtilityAccount")
$AllUsers = Get-WmiObject Win32_UserAccount -Filter "LocalAccount=True" |
    Where-Object { $ExcludedUsers -notcontains $_.Name } |
    Select-Object -ExpandProperty Name
$RegisteredUsers = ($AllUsers -join "`n")

# Primary registered user
$RegisteredUser = $SystemInfo.UserName
if ([string]::IsNullOrEmpty($RegisteredUser)) {
    $RegisteredUser = $OSInfo.RegisteredUser
}

# Current logged-in user
$CurrentUser = $env:USERNAME

# IP Address (active network adapters only)
$IPAddresses = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.InterfaceAlias -notlike "*Loopback*" -and $_.IPAddress -ne "127.0.0.1"
} | Select-Object -ExpandProperty IPAddress
$IPAddress = ($IPAddresses -join "`n")

# MAC Address
$MACAddresses = Get-NetAdapter | Where-Object {
    $_.Status -eq "Up" -and $_.InterfaceDescription -notlike "*Loopback*"
} | Select-Object -ExpandProperty MacAddress
$MACAddress = ($MACAddresses -join "`n")

# Manufacturer & Model
$Manufacturer = $SystemInfo.Manufacturer
$Model = $SystemInfo.Model

# Serial Number
$SerialNumber = $BiosInfo.SerialNumber

# OS Information
$OSName = $OSInfo.Caption
$OSVersion = $OSInfo.Version
$OSArchitecture = $OSInfo.OSArchitecture

# CPU Information
$CpuInfo = Get-WmiObject Win32_Processor | Select-Object -First 1
$CPUName = $CpuInfo.Name
$CPUCores = $CpuInfo.NumberOfCores
$CPULogicalProcessors = $CpuInfo.NumberOfLogicalProcessors

# Memory Information (GB)
$TotalMemoryGB = [math]::Round($SystemInfo.TotalPhysicalMemory / 1GB, 2)

# Memory Module Details
$MemoryModules = Get-WmiObject Win32_PhysicalMemory
$MemoryDetails = ($MemoryModules | ForEach-Object {
    "$($_.DeviceLocator): $([math]::Round($_.Capacity / 1GB))GB $($_.Speed)MHz"
}) -join "`n"

# RAM capacity per slot/lane
$RAMSlots = @()
$SlotIndex = 1
foreach ($module in $MemoryModules) {
    $capacityGB = [math]::Round($module.Capacity / 1GB, 2)
    $RAMSlots += "Slot$SlotIndex($($module.DeviceLocator)):${capacityGB}GB"
    $SlotIndex++
}
$RAMSlotInfo = ($RAMSlots -join "`n")

# GPU Information
$GpuInfoList = @()
$adapterMemory = Get-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0*" -Name "HardwareInformation.AdapterString", "HardwareInformation.qwMemorySize" -Exclude PSPath -ErrorAction SilentlyContinue
foreach ($adapter in $adapterMemory) {
    if ($adapter."HardwareInformation.AdapterString") {
        $vramGB = [math]::Round($adapter."HardwareInformation.qwMemorySize" / 1GB, 2)
        $GpuInfoList += "$($adapter.'HardwareInformation.AdapterString') ($vramGB GB)"
    }
}
$GPUInfo = ($GpuInfoList -join "`n")
if ([string]::IsNullOrEmpty($GPUInfo)) {
    $GPUInfo = (Get-WmiObject Win32_VideoController | Select-Object -ExpandProperty Name) -join "`n"
}

# Disk Information
$DiskDrives = Get-WmiObject Win32_DiskDrive
$DiskInfo = ($DiskDrives | ForEach-Object {
    $sizeGB = [math]::Round($_.Size / 1GB, 2)
    "$($_.Model) ($sizeGB GB)"
}) -join "`n"

# Motherboard Information
$BaseBoard = Get-WmiObject Win32_BaseBoard
$MotherboardManufacturer = $BaseBoard.Manufacturer
$MotherboardProduct = $BaseBoard.Product

# Network Adapter Information
$NetworkAdapters = Get-WmiObject Win32_NetworkAdapter | Where-Object {
    $_.PhysicalAdapter -eq $true -and $_.MACAddress -ne $null
}
$NetworkInfo = ($NetworkAdapters | ForEach-Object { $_.Name }) -join "`n"

# Create CSV object (1 row per PC)
$HardwareRecord = [PSCustomObject]@{
    "ComputerName" = $ComputerName
    "Users" = $RegisteredUsers
    # "PrimaryUser" = $RegisteredUser
    # "CurrentUser" = $CurrentUser
    # "Manufacturer" = $Manufacturer
    # "Model" = $Model
    "Manufacturer" = $MotherboardManufacturer
    "Motherboard" = $MotherboardProduct
    "CPU" = $CPUName
    "Cores" = $CPUCores
    "Processors" = $CPULogicalProcessors
    "GPU" = $GPUInfo
    "RAM" = $TotalMemoryGB
    # "RAMSlots" = $RAMSlotInfo
    "RAM Details" = $MemoryDetails
    "Disk" = $DiskInfo
    "IP" = $IPAddress
    "MAC" = $MACAddress
    "NetworkAdapter" = $NetworkInfo
    "OSName" = $OSName
    "OSVersion" = $OSVersion
    # "OSArchitecture" = $OSArchitecture
    # "SerialNumber" = $SerialNumber
    "LastUpdated" = $LastUpdated
}

# Output file path
$OutputPath = Join-Path $PSScriptRoot "..\hardware_info.csv"

# CSV output (append if exists, create new if not)
if (Test-Path $OutputPath) {
    # Append to existing file (no header)
    $HardwareRecord | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation -Append
    Write-Host "Hardware info appended: $OutputPath"
} else {
    # Create new file (with header)
    $HardwareRecord | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation
    Write-Host "Hardware info created: $OutputPath"
}

# Display on screen
Write-Host "`n========== Collected Information =========="
$HardwareRecord | Format-List

Write-Host "`nCSV output completed: $OutputPath"
