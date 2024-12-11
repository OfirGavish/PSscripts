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


$today = Get-Date
$alert = ""


$count = 0

Foreach ($device in $devices)
{
     if ($device.operatingSystem -match "Windows") 
    {
        if ($count -ge 3500)
        {
            # Get access token for Graph API
            $graphApiToken = Get-GraphAPIAccessToken

            # Create header for using Graph API
            $graphApiHeader = @{ Authorization = "Bearer $graphApiToken" }
        }
        $deviceid = $device.id
        $compuri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceid"
        $devicenew = Invoke-RestMethod -Uri $compuri -Headers $graphApiHeader -Method Get

        if ($devicenew.deviceActionResults) 
        {
            # Filter the deviceActionResults for "collectlogs" actions
            $collectLogsActions = $devicenew.deviceActionResults | Where-Object { $_.actionName -eq "collectlogs" -and $_.actionState -eq "done"}

            if ($collectLogsActions.Count -gt 0) 
            {
                # Sort the filtered actions by lastUpdatedDateTime in descending order
                $mostRecentAction = $collectLogsActions | Sort-Object -Property lastUpdatedDateTime -Descending | Select-Object -First 1

                # Parse the most recent lastUpdatedDateTime
                $lastUpdatedDateTime = [DateTime]::Parse($mostRecentAction.lastUpdatedDateTime)

                # Calculate the timespan (difference in days between today and lastUpdatedDateTime)
                $timespan = ($today - $lastUpdatedDateTime).Days

                # Check if the timespan is less than 30 days
                if ($timespan -lt 30) 
                {
                    $alert += "$($devicenew.deviceName) has diagnostics logs ready to be <a href=`"https://intune.microsoft.com/?pwa=1#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/logCollection/mdmDeviceId/$($deviceid)`" style=`"color: #8c8aeb`">collected<a/> (Last Updated: $($lastUpdatedDateTime.ToString('yyyy-MM-dd HH:mm:ss')))" + "<br>"
                }
            }
        }
        $count ++
    }
}
if ($alert){
#configure variables for email notifications
$from = "IntuneAlerts@domain.com"
$to = "cloudops@domain.com"
$cred = Get-AutomationPSCredential -Name 'Credential-Name'


$htmlbody = @"
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Email Template</title>
</head>
<body>
<table width="600" align="center" border="0" cellpadding="0" cellspacing="0">
<tr>
<td>
<p style="text-align: center; font-size: 16px; font-weight: 300; line-height: 1.5;">Device Diagnostics Available Report</p>
<p style="text-align: center; font-size: 14px; font-weight: 300; line-height: 1.5;">$alert</p>
</td>
</tr>
</body>
</html>
"@


if ($alert)
{
    try {
        Send-MailMessage -To $to -From $from -Credential $cred -Subject "Device Diagnostics Available Report" -Body $htmlbody -BodyAsHtml -SmtpServer "smtp.azurecomm.net" -Port 587 -UseSsl
    }
    catch {
        Throw "Failed to send email notification: $_"
    }
}
exit 0
}
