#Create a function to get Graph API Access token
function Get-GraphAPIAccessToken {
    $resource= "?resource=https://graph.microsoft.com/"
    $url = $env:IDENTITY_ENDPOINT + $resource
    $Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $Headers.Add("X-IDENTITY-HEADER", $env:IDENTITY_HEADER)
    $Headers.Add("Metadata", "True")
    $accessToken = Invoke-RestMethod -Uri $url -Method 'GET' -Headers $Headers
    return $accessToken.access_token
}

# Get access token for Graph API
$graphApiToken = Get-GraphAPIAccessToken

# Create header for using Graph API
$graphApiHeader = @{ Authorization = "Bearer $graphApiToken" }

#Got authentication token, running query
$devicesuri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
$devices = (Invoke-RestMethod -Uri $devicesuri -Headers $graphApiHeader -Method Get).Value

$count = 1
$today = Get-Date

# Create an array to store the alerts
$alert = ""

foreach ($device in $devices) {
    if ($device.operatingSystem -match "Windows") 
    {
        if ($count -ge 3500)
        {
            # Get access token for Graph API
            $graphApiToken = Get-GraphAPIAccessToken

            # Create header for using Graph API
            $graphApiHeader = @{ Authorization = "Bearer $graphApiToken" }
        }
        $deviceuri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($device.id)"
        $devicenew = Invoke-RestMethod -Uri $deviceuri -Headers $graphApiHeader -Method Get

        if ($devicenew.deviceActionResults) {
            # Filter the deviceActionResults for "collectlogs" actions
            $collectLogsActions = $devicenew.deviceActionResults | Where-Object { $_.actionName -eq "collectlogs" }

            if ($collectLogsActions.Count -gt 0) {
                # Sort the filtered actions by lastUpdatedDateTime in descending order
                $mostRecentAction = $collectLogsActions | Sort-Object -Property lastUpdatedDateTime -Descending | Select-Object -First 1

                # Parse the most recent lastUpdatedDateTime
                $lastUpdatedDateTime = [DateTime]::Parse($mostRecentAction.lastUpdatedDateTime)

                # Calculate the timespan (difference in days between today and lastUpdatedDateTime)
                $timespan = ($today - $lastUpdatedDateTime).Days

                # Check if the timespan is less than 30 days
                if ($timespan -lt 30) {
                    $alert += "$($devicenew.deviceName) has diagnostics logs ready to be collected (Last Updated: $($lastUpdatedDateTime.ToString('yyyy-MM-dd HH:mm:ss'))) `n"
                }
            }
        else {
            Write-Host "device: $($devicenew.deviceName) has device actions but non are regarding Diagnostics"
        }
        }
        $count ++
    }
}

Write-Host $alert
