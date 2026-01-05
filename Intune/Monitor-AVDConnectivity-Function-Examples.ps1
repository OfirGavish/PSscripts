
# Example 1: defaults (10s interval, 20 minutes)
$csv = Start-AVDConnectivityMonitor

# Example 2: specify targets, include UDP probes
$csv = Start-AVDConnectivityMonitor `
  -ExternalTargets 'www.bbc.com' `
  -MicrosoftTargets 'login.microsoftonline.com' `
  -EnableUdp

# Example 3: pin a known AVD Gateway and keep UDP off
$csv = Start-AVDConnectivityMonitor `
  -AvdGatewayTargets 'global.rdgatewayavd.microsoft.com'
