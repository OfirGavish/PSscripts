
<# Remediate-AvdPerfmon-Forensics.ps1
   Intune Proactive Remediations - Remediation script (forensics burst)
   Creates/refreshes AVD forensic PerfMon DCS and a startup Scheduled Task (SYSTEM).
   Compatible with Windows PowerShell 5.1.
#>

[CmdletBinding()]
param(
  [int] $SampleSeconds = 10,
  [int] $RetentionDays = 3
)

$ErrorActionPreference = 'Stop'

$setName  = 'AVD-Session-Forensics'
$taskName = 'Start-AVD-PerfMon-Forensics'

# Pick D: if present, else C:
if (Test-Path 'D:\') {
    $logRoot = 'D:\PerfLogs\AVD-Forensics'
} else {
    $logRoot = 'C:\PerfLogs\AVD-Forensics'
}

$countersFile = Join-Path $env:ProgramData 'AVD\perfmon-forensics.txt'
$logPrefix    = Join-Path $logRoot 'AVD-Forensics'

# Ensure folders
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
New-Item -Path (Split-Path $countersFile) -ItemType Directory -Force | Out-Null

# Counter set — expanded objects, broad forensics (heavy; time-boxed)
$counterList = @'
System\*
Processor(_Total)\*
Processor Information(_Total)\*
Processor Performance\*
Memory\*
Paging File(_Total)\*
LogicalDisk(_Total)\*
PhysicalDisk(_Total)\*
Network Interface(*)\*
TCPv4\*
TCPv6\*
Terminal Services\*
Terminal Services Session(*)\*
Cache\*
Process(*)\*
Process V2(*)\*
Netlogon\*
Offline Files\*
TCP IP Performance Diagnostics\*
TCP IP Extended Diagnostics\*
TCP IP Transport Layer Packet Drop Counters\*
'@

# Write counters list (ASCII required by logman)
$counterList | Set-Content -Path $countersFile -Encoding ASCII

# Stop & remove existing DCS if present (idempotent)
try { & logman.exe stop $setName 2>$null | Out-Null } catch {}
try { & logman.exe delete $setName 2>$null | Out-Null } catch {}

# Create new DCS
$si = ('00:00:{0:00}' -f $SampleSeconds)  # HH:MM:SS
$createArgs = @(
    'create','counter', $setName,
    '-cf', $countersFile,
    '-si', $si,
    '-f', 'csv',
    '-o', $logPrefix,
    '-v', 'mmddhhmm',
    '-cnf','01:00:00'        # new file every hour
)
& logman.exe @createArgs | Out-Null

# Start DCS
& logman.exe start $setName | Out-Null

# Scheduled Task — start DCS at system boot, runs as SYSTEM
$action    = New-ScheduledTaskAction -Execute 'logman.exe' -Argument ("start `"{0}`"" -f $setName)
$trigger   = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
$settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable
$task      = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings

try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
Register-ScheduledTask -TaskName $taskName -InputObject $task | Out-Null

# Retention cleanup
try {
    Get-ChildItem -Path $logRoot -Filter '*.csv' -Recurse -ErrorAction SilentlyContinue |
      Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays) } |
      Remove-Item -Force -ErrorAction SilentlyContinue
} catch {}

Write-Output ("Forensic PerfMon enabled. DCS '{0}' running; logs at {1}; startup task '{2}' created." -f $setName, $logRoot, $taskName)
