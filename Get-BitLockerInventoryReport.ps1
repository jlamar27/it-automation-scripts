<#
.SYNOPSIS
    Generates a BitLocker status and device inventory report for Windows devices managed in Microsoft Intune/Azure AD.

.DESCRIPTION
    This script connects to Microsoft Graph to retrieve all Windows managed devices and their BitLocker recovery keys.
    It produces a detailed report including device model, manufacturer, OS version, BitLocker encryption status, and recovery key presence.
    The results are exported to an Excel file with both detailed device data and a summary table.

.REQUIREMENTS
    - Microsoft.Graph PowerShell modules
    - ImportExcel PowerShell module

.PERMISSIONS
    Requires the following Microsoft Graph API permissions:
        - DeviceManagementManagedDevices.Read.All
        - BitLockerKey.Read.All

.PARAMETER exportPath
    The file path where the Excel report will be saved.

.NOTES
    - The script distinguishes between Windows 10 and Windows 11 using the OS build number.
    - BitLocker recovery keys are matched to devices using Azure AD Device ID.
    - The Excel report includes a "Devices" worksheet with detailed data and a "Summary" worksheet with counts by Windows version and encryption status.

.EXAMPLE
    Run the script to generate a BitLocker status report:
        .\ModelAndBitlocker.ps1

.AUTHORe
    Juan Lamar

#>
<# Requires:
   - Microsoft.Graph modules
   - ImportExcel module
   Scopes needed: DeviceManagementManagedDevices.Read.All, BitLockerKey.Read.All
#>

[CmdletBinding()]
param(
  [string]$ExportPath = "$env:USERPROFILE\Documents\BitLockerStatusReportModels.xlsx"
)

Import-Module ImportExcel -ErrorAction Stop

# Connect only if not already connected
try {
  if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All","BitLockerKey.Read.All"
  }
} catch {
  throw "Failed to connect to Microsoft Graph. $_"
}

$devices = Get-MgDeviceManagementManagedDevice -All |
  Where-Object { $_.OperatingSystem -eq "Windows" }

$recoveryKeys = Get-MgInformationProtectionBitlockerRecoveryKey -All

$keysByDeviceId = @{}
foreach ($k in $recoveryKeys) {
  if ([string]::IsNullOrWhiteSpace($k.DeviceId)) { continue }
  ($keysByDeviceId[$k.DeviceId] ??= @()) += $k
}

function Get-WindowsVersionTag {
  param([string]$OsVersion)
  try { if ([version]$OsVersion -ge [version]"10.0.22000") { "Windows 11" } else { "Windows 10" } }
  catch { "Unknown" }
}

$results = foreach ($d in $devices) {
  $aadId  = $d.AzureAdDeviceId
  $verTag = Get-WindowsVersionTag $d.OsVersion
  $isEnc  = if ($null -ne $d.IsEncrypted) { if ($d.IsEncrypted){"Yes"}else{"No"} } else {"Unknown"}
  $keys   = if ($aadId -and $keysByDeviceId.ContainsKey($aadId)) { $keysByDeviceId[$aadId] } else { @() }

  [pscustomobject]@{
    DeviceName          = $d.DeviceName
    AzureADDeviceId     = $aadId
    Model               = $d.Model
    Manufacturer        = $d.Manufacturer
    SerialNumber        = $d.SerialNumber
    OperatingSystem     = $d.OperatingSystem
    OSVersion           = $d.OsVersion
    WindowsVersion      = $verTag
    IsEncrypted         = $isEnc
    RecoveryKeyInAAD    = if ($keys.Count -gt 0) {"Yes"} else {"No"}
    RecoveryKeyCount    = $keys.Count
    RecoveryVolumeTypes = ($keys | Select-Object -Expand VolumeType -Unique) -join ", "
    RecoveryKeyIds      = ($keys | Select-Object -Expand Id) -join ", "
    ComplianceState     = $d.ComplianceState
    LastSyncDateTime    = $d.LastSyncDateTime
    EnrolledDateTime    = $d.EnrolledDateTime
  }
}

$results | Export-Excel -Path $ExportPath -WorksheetName "Devices" -AutoSize -ClearSheet -FreezeTopRow -BoldTopRow -TableName "Devices"

$summary = $results |
  Group-Object WindowsVersion, IsEncrypted |
  Select-Object @{n='WindowsVersion';e={$_.Group[0].WindowsVersion}},
                @{n='IsEncrypted';e={$_.Group[0].IsEncrypted}},
                @{n='Count';e={$_.Count}} |
  Sort-Object WindowsVersion, IsEncrypted

$summary | Export-Excel -Path $ExportPath -WorksheetName "Summary" -AutoSize -ClearSheet

Write-Host "Exported report to: $ExportPath"
