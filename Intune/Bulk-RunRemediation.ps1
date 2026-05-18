<#
.SYNOPSIS
Bulk-trigger Intune Remediations (On-demand) via Microsoft Graph with progress, retries, and reporting.

.REQUIREMENTS
- Microsoft Graph PowerShell SDK
- Permissions (delegated or app):
  DeviceManagementManagedDevices.PrivilegedOperations.All  (required to initiate)  [1](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-initiateondemandproactiveremediation?view=graph-rest-beta)
  DeviceManagementManagedDevices.Read.All                  (to list devices)      [3](https://ourcloudnetwork.com/perform-bulk-on-demand-remediation-with-microsoft-intune/)
  DeviceManagementConfiguration.Read.All                   (to read scripts)      [3](https://ourcloudnetwork.com/perform-bulk-on-demand-remediation-with-microsoft-intune/)

.NOTES
- API: POST /deviceManagement/managedDevices/{managedDeviceId}/initiateOnDemandProactiveRemediation
  Body: { "scriptPolicyId": "<script GUID>" } [1](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-initiateondemandproactiveremediation?view=graph-rest-beta)
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ScriptPackageName,                 # Exact remediation name in Intune

    [string]$DeviceFilter = "operatingSystem eq 'Windows'",

    [string]$DeviceListPath,                    # Optional: txt/csv list of device names (one per line or CSV column DeviceName)

    [int]$ThrottleLimit = 10,                   # PS7+ parallel throttle
    [int]$MaxRetries    = 3,
    [int]$BaseDelaySec  = 2,

    [string]$OutputFolder = ".\IntuneRemediationRuns",
    [switch]$WhatIf
)

# ----------------------------
# Helper: ensure output folder
# ----------------------------
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$csvPath   = Join-Path $OutputFolder "TriggerResults-$timestamp.csv"
$jsonPath  = Join-Path $OutputFolder "TriggerResults-$timestamp.json"

# ----------------------------
# Connect to Graph
# ----------------------------
$scopes = @(
  "DeviceManagementManagedDevices.PrivilegedOperations.All",
  "DeviceManagementManagedDevices.Read.All",
  "DeviceManagementConfiguration.Read.All",
  "DeviceManagementScripts.ReadWrite.All"
)

Connect-MgGraph -Scopes $scopes | Out-Null

# ----------------------------
# Get remediation package (deviceHealthScript)
# ----------------------------
$remediation = Get-MgBetaDeviceManagementDeviceHealthScript -Filter "displayName eq '$ScriptPackageName'"
if (-not $remediation) {
    throw "Remediation package not found: '$ScriptPackageName'"
}

$scriptId = $remediation.Id
Write-Host "Remediation found: $($remediation.DisplayName)  (Id: $scriptId)"

# ----------------------------
# Resolve target devices
# ----------------------------
$devices = @()

if ($DeviceListPath) {
    if (-not (Test-Path $DeviceListPath)) { throw "DeviceListPath not found: $DeviceListPath" }

    if ($DeviceListPath.ToLower().EndsWith(".csv")) {
        $names = (Import-Csv $DeviceListPath) | ForEach-Object { $_.DeviceName } | Where-Object { $_ } | Select-Object -Unique
    } else {
        $names = Get-Content $DeviceListPath | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() } | Select-Object -Unique
    }

    foreach ($n in $names) {
        $d = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$n'" -Top 1
        if ($d) { $devices += $d }
        else { Write-Warning "Device not found in Intune: $n" }
    }
} else {
    $devices = Get-MgDeviceManagementManagedDevice -Filter $DeviceFilter -All
}

if (-not $devices -or $devices.Count -eq 0) {
    throw "No target devices resolved."
}

Write-Host "Targets: $($devices.Count) device(s)"

# ----------------------------
# Build request payload (API expects scriptPolicyId) [1](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-initiateondemandproactiveremediation?view=graph-rest-beta)
# ----------------------------
$bodyObj = @{ scriptPolicyId = $scriptId }
$bodyJson = ($bodyObj | ConvertTo-Json -Depth 3)

# ----------------------------
# Core trigger function with retry
# ----------------------------
function Invoke-TriggerWithRetry {
    param(
        [string]$ManagedDeviceId,
        [string]$DeviceName,
        [string]$BodyJson,
        [int]$MaxRetries,
        [int]$BaseDelaySec
    )

    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$ManagedDeviceId/initiateOnDemandProactiveRemediation"

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            # Successful response is 204 No Content [1](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-initiateondemandproactiveremediation?view=graph-rest-beta)
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $BodyJson -ContentType "application/json" | Out-Null

            return [pscustomobject]@{
                DeviceName   = $DeviceName
                ManagedDeviceId = $ManagedDeviceId
                Triggered    = $true
                Attempts     = $attempt
                Error        = $null
                TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
            }
        }
        catch {
            $err = $_.Exception.Message
            if ($attempt -lt $MaxRetries) {
                $delay = [Math]::Min(60, $BaseDelaySec * [Math]::Pow(2, ($attempt-1)))
                Start-Sleep -Seconds $delay
            } else {
                return [pscustomobject]@{
                    DeviceName   = $DeviceName
                    ManagedDeviceId = $ManagedDeviceId
                    Triggered    = $false
                    Attempts     = $attempt
                    Error        = $err
                    TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
                }
            }
        }
    }
}

# ----------------------------
# Execute (parallel if PS7+)
# ----------------------------
$results = New-Object System.Collections.Generic.List[object]
$total = $devices.Count
$done  = 0

if ($WhatIf) {
    Write-Host "WHATIF mode: no triggers sent."
    $devices | Select-Object deviceName, id | Format-Table -AutoSize
    return
}

if ($PSVersionTable.PSVersion.Major -ge 7) {
    Write-Host "Running in parallel (PS7+) ThrottleLimit=$ThrottleLimit"

    $parallelResults = $devices | ForEach-Object -Parallel {
        $bodyJson     = $using:bodyJson
        $MaxRetries   = $using:MaxRetries
        $BaseDelaySec = $using:BaseDelaySec

        function Invoke-TriggerWithRetryLocal {
            param($ManagedDeviceId, $DeviceName, $BodyJson, $MaxRetries, $BaseDelaySec)

            $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$ManagedDeviceId/initiateOnDemandProactiveRemediation"
            for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
                try {
                    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $BodyJson -ContentType "application/json" | Out-Null
                    return [pscustomobject]@{
                        DeviceName   = $DeviceName
                        ManagedDeviceId = $ManagedDeviceId
                        Triggered    = $true
                        Attempts     = $attempt
                        Error        = $null
                        TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
                    }
                } catch {
                    $err = $_.Exception.Message
                    if ($attempt -lt $MaxRetries) {
                        $delay = [Math]::Min(60, $BaseDelaySec * [Math]::Pow(2, ($attempt-1)))
                        Start-Sleep -Seconds $delay
                    } else {
                        return [pscustomobject]@{
                            DeviceName   = $DeviceName
                            ManagedDeviceId = $ManagedDeviceId
                            Triggered    = $false
                            Attempts     = $attempt
                            Error        = $err
                            TimestampUtc = (Get-Date).ToUniversalTime().ToString("o")
                        }
                    }
                }
            }
        }

        Invoke-TriggerWithRetryLocal -ManagedDeviceId $_.Id -DeviceName $_.DeviceName -BodyJson $bodyJson -MaxRetries $MaxRetries -BaseDelaySec $BaseDelaySec

    } -ThrottleLimit $ThrottleLimit

    foreach ($r in $parallelResults) {
        $results.Add($r) | Out-Null
    }
}
else {
    Write-Host "Running sequentially (PowerShell 5.1)."

    foreach ($d in $devices) {
        $done++
        $pct = [int](($done / $total) * 100)
        Write-Progress -Activity "Triggering on-demand remediation" -Status "$done / $total : $($d.DeviceName)" -PercentComplete $pct

        $r = Invoke-TriggerWithRetry -ManagedDeviceId $d.Id -DeviceName $d.DeviceName -BodyJson $bodyJson -MaxRetries $MaxRetries -BaseDelaySec $BaseDelaySec
        $results.Add($r) | Out-Null
    }

    Write-Progress -Activity "Triggering on-demand remediation" -Completed
}

# ----------------------------
# Reporting
# ----------------------------
$results | Export-Csv -NoTypeInformation -Path $csvPath -Encoding UTF8
$results | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath -Encoding UTF8

$ok   = ($results | Where-Object Triggered).Count
$fail = ($results | Where-Object { -not $_.Triggered }).Count

Write-Host ""
Write-Host "Triggered successfully: $ok"
Write-Host "Failed to trigger:       $fail"
Write-Host "CSV:  $csvPath"
Write-Host "JSON: $jsonPath"

if ($fail -gt 0) {
    Write-Host ""
    Write-Host "Top failures:"
    $results | Where-Object { -not $_.Triggered } | Select-Object -First 10 DeviceName, Attempts, Error | Format-List
}

Write-Host ""
Write-Host "Next: Monitor results in Intune Remediations reports / device status and export output as needed." 
