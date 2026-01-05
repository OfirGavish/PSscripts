
#region ===================== AVD Connectivity Monitor (function) =====================

function Start-AVDConnectivityMonitor {
<#
.SYNOPSIS
  Continuous TCP (and optional UDP) reachability monitor for AVD troubleshooting.

.DESCRIPTION
  Checks 3 categories:
    1) External unrelated website(s)
    2) Microsoft login FQDN(s)
    3) AVD gateway(s) in use (auto-discovered or explicitly provided)
  Writes CSV under D:\AVDNetLogs (or C:\AVDNetLogs) and, if elevated, Error entries to Application log (source: AVD-NetMonitor).
  PowerShell 5.1–compatible.

.PARAMETER ExternalTargets
  One or more FQDNs for general internet reachability. Default: www.bing.com

.PARAMETER MicrosoftTargets
  One or more FQDNs for Microsoft sign-in path. Default: login.microsoftonline.com

.PARAMETER AvdGatewayTargets
  Optional explicit AVD gateway FQDN/IP list. If omitted, function auto-discovers from active TCP 443 flows and RDInfraAgent broker URI.

.PARAMETER EnableUdp
  Include best-effort UDP probes (Shortpath/STUN-like). Default: Disabled.

.PARAMETER UdpPorts
  UDP ports to probe (default 3390, 3478). No handshake; a "no response" is not a failure by itself.

.PARAMETER TcpPorts
  TCP ports to test (default 443).

.PARAMETER IntervalSeconds
  Sampling interval in seconds. Default: 10.

.PARAMETER RunMinutes
  Total duration in minutes. Default: 20.
#>

    [CmdletBinding()]
    param(
        [string[]] $ExternalTargets   = @('www.bing.com'),
        [string[]] $MicrosoftTargets  = @('login.microsoftonline.com'),
        [string[]] $AvdGatewayTargets = @(),  # leave empty to auto-discover
        [switch]   $EnableUdp,
        [int[]]    $UdpPorts          = @(3390,3478),
        [int[]]    $TcpPorts          = @(443),
        [int]      $IntervalSeconds   = 10,   # requested default
        [int]      $RunMinutes        = 20    # requested default
    )

    $ErrorActionPreference = 'Stop'

    # --- Paths & CSV header ---
    if (Test-Path 'D:\') { $logRoot = 'D:\AVDNetLogs' } else { $logRoot = 'C:\AVDNetLogs' }
    $sessionId = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $csvPath   = Join-Path $logRoot ("AVD-Connectivity_{0}.csv" -f $sessionId)
    New-Item -Path $logRoot -ItemType Directory -Force | Out-Null

    $header = 'Timestamp,ComputerName,Category,Target,Protocol,Port,Success,SuccessDetail,LatencyMs,RemoteAddress,ResolvedIP,ErrorText'
    $header | Out-File -FilePath $csvPath -Encoding ASCII

    # --- Elevation & Event Log setup ---
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
                ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $EnableEventLog = $IsAdmin
    $eventSource = 'AVD-NetMonitor'
    $eventLog    = 'Application'
    if ($EnableEventLog) {
        try { New-EventLog -LogName $eventLog -Source $eventSource -ErrorAction SilentlyContinue } catch { $EnableEventLog = $false }
    }

    function Write-FailureEvent([string]$TargetText, [string]$CategoryText, [string]$MessageText) {
        if ($EnableEventLog) {
            try {
                $msg = ("[Category={0}] {1}" -f $CategoryText, $MessageText)
                Write-EventLog -LogName $eventLog -Source $eventSource -EntryType Error -EventId 60100 -Message $msg
            } catch { }
        }
    }

    # --- Utilities (PS 5.1-safe) ---
    function Resolve-FirstIp([string]$TargetHost) {
        try {
            $ip = [System.Net.Dns]::GetHostAddresses($TargetHost) | Select-Object -First 1
            if ($ip) { return $ip.IPAddressToString } else { return $null }
        } catch { return $null }
    }

    function Test-TcpOnce([string]$TargetHost, [int]$TargetPort) {
        try {
            $res = Test-NetConnection -ComputerName $TargetHost -Port $TargetPort -WarningAction SilentlyContinue -InformationLevel Detailed
            $latency = $null
            if ($res.PingReplyDetails) { $latency = $res.PingReplyDetails.RoundtripTime }
            $detail = if ($res.TcpTestSucceeded) { 'TcpHandshakeOk' } else { 'TcpHandshakeFailedOrTimeout' }
            return [pscustomobject]@{
                Success       = [bool]$res.TcpTestSucceeded
                SuccessDetail = $detail
                LatencyMs     = $latency
                RemoteAddress = $res.RemoteAddress
                ErrorText     = $null
            }
        } catch {
            return [pscustomobject]@{
                Success       = $false
                SuccessDetail = 'TcpException'
                LatencyMs     = $null
                RemoteAddress = $null
                ErrorText     = $_.Exception.Message
            }
        }
    }

    function Test-UdpOnce([string]$TargetHost, [int]$TargetPort, [int]$TimeoutMs = 1200) {
        # Best-effort UDP probe (no handshake). "No response" is expected; success = "probe sent".
        try {
            $resolved = Resolve-FirstIp $TargetHost
            $client   = New-Object System.Net.Sockets.UdpClient
            $client.Client.ReceiveTimeout = $TimeoutMs
            $client.Connect($TargetHost, [int]$TargetPort)
            $payload = 0x01,0x02,0x03,0x04
            $start   = Get-Date
            [void]$client.Send($payload, $payload.Length)
            $remote  = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any,0)
            $received = $false
            try {
                $bytes = $client.Receive([ref]$remote)  # likely times out
                $received = $true
            } catch { $received = $false }
            $elapsed = (Get-Date) - $start
            $client.Close()

            $detail = if ($received) { 'UdpResponseReceived' } else { 'UdpProbeSentNoResponse' }
            $remoteAddr = $null
            if ($remote -and $remote.Address) { $remoteAddr = $remote.Address.ToString() }

            return [pscustomobject]@{
                Success       = $true
                SuccessDetail = $detail
                LatencyMs     = [int]$elapsed.TotalMilliseconds
                RemoteAddress = $remoteAddr
                ErrorText     = $null
            }
        } catch {
            return [pscustomobject]@{
                Success       = $false
                SuccessDetail = 'UdpSendFailed'
                LatencyMs     = $null
                RemoteAddress = $null
                ErrorText     = $_.Exception.Message
            }
        }
    }

    function Discover-AvdGatewayTargets {
        $targets = @()

        # 1) Live outbound TCP 443 flows (fast & pragmatic)
        try {
            $conns = Get-NetTCPConnection -State Established -RemotePort 443 | Sort-Object -Property CreationTime -Descending
            if ($conns) {
                $ips = $conns | Select-Object -ExpandProperty RemoteAddress -Unique
                if ($ips) { $targets += $ips }
            }
        } catch { }

        # 2) RDInfraAgent broker URI (common on AVD hosts)
        try {
            $key = 'HKLM:\SOFTWARE\Microsoft\RDInfraAgent'
            $prop = Get-ItemProperty -Path $key -Name 'BrokerResourceIdURI' -ErrorAction SilentlyContinue
            if ($prop -and $prop.BrokerResourceIdURI) {
                $uri = $prop.BrokerResourceIdURI
                try {
                    $u = [Uri]$uri
                    if ($u -and $u.Host) { $targets += $u.Host }
                } catch { }
            }
        } catch { }

        # Deduplicate & return
        $out = @()
        foreach ($t in $targets) { if ($t -and $t.Trim() -ne '') { $out += $t } }
        $out = $out | Select-Object -Unique
        if (-not $out -or $out.Count -eq 0) { $out = @('wvd.microsoft.com') } # fallback
        return $out
    }

    # --- Banner ---
    $extList = ($ExternalTargets -join ', ')
    $msList  = ($MicrosoftTargets -join ', ')
    $avdList = if ($AvdGatewayTargets -and $AvdGatewayTargets.Count -gt 0) { ($AvdGatewayTargets -join ', ') } else { 'autodiscover' }
    Write-Host ("Monitoring → External: {0} | Microsoft: {1} | AVD: {2}" -f $extList, $msList, $avdList)
    if (-not $IsAdmin) { Write-Warning "Not running elevated; Event Log writes disabled. CSV logging continues." }

    $stopAt = (Get-Date).AddMinutes($RunMinutes)

    while ((Get-Date) -lt $stopAt) {

        # Build final target sets (auto-discover AVD if needed)
        if ($AvdGatewayTargets -and $AvdGatewayTargets.Count -gt 0) { $avdTargets = $AvdGatewayTargets } else { $avdTargets = Discover-AvdGatewayTargets }

        $sets = @(
            @{ Category='external';  Targets=$ExternalTargets },
            @{ Category='microsoft'; Targets=$MicrosoftTargets },
            @{ Category='avd-gateway'; Targets=$avdTargets }
        )

        foreach ($set in $sets) {
            foreach ($target in $set.Targets) {

                # Resolve DNS once per target per loop
                $resolved = Resolve-FirstIp $target
                if (-not $resolved) { $resolved = '' }

                # --- TCP tests ---
                foreach ($tp in $TcpPorts) {
                    $tcp = Test-TcpOnce -TargetHost $target -TargetPort $tp

                    $latOut    = if ($tcp.LatencyMs) { $tcp.LatencyMs } else { '' }
                    $remoteOut = if ($tcp.RemoteAddress) { $tcp.RemoteAddress } else { '' }
                    $errOut    = if ($tcp.ErrorText) { '"' + $tcp.ErrorText.Replace('"','''') + '"' } else { '' }

                    $line = ('{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}' -f `
                        (Get-Date).ToString('o'), $env:COMPUTERNAME, $set.Category, $target, 'TCP', $tp,
                        $tcp.Success, $tcp.SuccessDetail, $latOut, $remoteOut, $resolved, $errOut)
                    Add-Content -Path $csvPath -Value $line

                    if (-not $tcp.Success) {
                        $msg = ("TCP:{0} failed ({1}). {2}" -f $tp, $tcp.SuccessDetail, $tcp.ErrorText)
                        Write-FailureEvent -TargetText $target -CategoryText $set.Category -MessageText $msg
                    }
                }

                # --- UDP probes (optional/best-effort) ---
                if ($EnableUdp) {
                    foreach ($up in $UdpPorts) {
                        $udp = Test-UdpOnce -TargetHost $target -TargetPort $up
                        $latOut    = if ($udp.LatencyMs) { $udp.LatencyMs } else { '' }
                        $remoteOut = if ($udp.RemoteAddress) { $udp.RemoteAddress } else { '' }
                        $errOut    = if ($udp.ErrorText) { '"' + $udp.ErrorText.Replace('"','''') + '"' } else { '' }

                        $line = ('{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}' -f `
                            (Get-Date).ToString('o'), $env:COMPUTERNAME, $set.Category, $target, 'UDP', $up,
                            $udp.Success, $udp.SuccessDetail, $latOut, $remoteOut, $resolved, $errOut)
                        Add-Content -Path $csvPath -Value $line

                        if (-not $udp.Success) {
                            $msg = ("UDP:{0} failed ({1}). {2}" -f $up, $udp.SuccessDetail, $udp.ErrorText)
                            Write-FailureEvent -TargetText $target -CategoryText $set.Category -MessageText $msg
                        }
                    }
                }
            }
        }

        Start-Sleep -Seconds $IntervalSeconds
    }

    Write-Host ("Done. CSV: {0}" -f $csvPath)
    return $csvPath
}

#endregion ================== AVD Connectivity Monitor (function) ==================
