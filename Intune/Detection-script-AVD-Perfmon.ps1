 
<# Detect-AvdPerfmon.ps1
   Intune Proactive Remediation - Detection script
   Exit 0 = compliant, Exit 1 = needs remediation
#>

$ErrorActionPreference = 'Stop'
$setName  = 'AVD-Session-Troubleshooting'
$taskName = 'Start-AVD-PerfMon-DCS'

function Test-DcsRunning {
    try {
        $q = & logman.exe query $setName 2>$null
        if (-not $q) { return $false }
        # If the query output contains "Status: Running", we treat as compliant
        return ($q -match 'Status:\s+Running')
    } catch { return $false }
}

function Test-TaskOk {
    try {
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
        # Must be AtStartup trigger and SYSTEM principal
        $hasStartup = $task.Triggers | Where-Object { $_.TriggerType -eq 'Boot' -or $_.TriggerType -eq 'Logon' -or $_.TriggerType -eq 'AtStartup' }
        $isSystem   = $task.Principal.UserId -eq 'SYSTEM'
        return ($hasStartup -and $isSystem)
    } catch { return $false }
}

$okDcs  = Test-DcsRunning
$okTask = Test-TaskOk

if ($okDcs -and $okTask) {
    Write-Output "PerfMon DCS '$setName' is running and startup task '$taskName' is configured."
    exit 0
} else {
    Write-Output "PerfMon DCS or startup task missing/misconfigured. DCSRunning=$okDcs TaskOk=$okTask"
    exit 1
}
