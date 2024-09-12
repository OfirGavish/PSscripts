<#
.SYNOPSIS
This script is designed to query Intune Proactive Remediation Policy Device Status and send an alert to Teams channel and Email if there are any errors.

.DESCRIPTION
This script is desinged to query Intune Proactive Remediation Policy Device Status and send an alert to Teams channel and Email if there are any errors.
The script will query the Intune API to get the status of the Proactive Remediation Policy and filter out the devices with status other than "Without Issues" and "Pending".
The script will then query the Intune API to get the device name and user principal name of the devices with errors.
The script will then build a JSON body for the Teams channel notification and send the notification to the Teams channel.
The script will then build an HTML body for the email message and send the email message to the recipient.

.EXAMPLE
Use this script on Azure Automation runbook to send alerts to Teams channel and Email if there are any errors in the Proactive Remediation Policy.
Assign the Automation account with Exchange.ManageAsApp Application permissions to send an email alert.
Assign the Automation account with the required permissions to query the Intune API to get the Proactive Remediation Policy Device Status.
Fill in the required values in the script and run the script on Azure Automation runbook.

.NOTES
Author: Ofir Gavish
Special Thanks to Hen Eshet
Date: 09/2024
Version: 2.0
#>


#Region Connect to Microsoft Graph API for Intune Remediations status
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
#Query Intune Proactive Remediation Policy Device Status (Fill in Policy ID)
$uri2 = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/<ProactiveRemediationPolicyID>/deviceRunStates"
try
{
  $response = (Invoke-RestMethod -Uri $uri2 -Headers $graphApiHeader -Method Get).Value
}
catch
{
  Throw $_.Exception.Message
}
#Filter out Status of Pending + success("Without Issues") and Query Intune API to resolve IntuneDeviceID to Device Name (Getting data for alert).
try
{
  #region Proccessing response to generate email alert and log analytics json
function merge ($target, $source) {
  $source.psobject.Properties | % {
      if ($_.TypeNameOfValue -eq 'System.Management.Automation.PSCustomObject' -and $target."$($_.Name)" ) {
          merge $target."$($_.Name)" $_.Value
      }
      else {
          $target | Add-Member -MemberType $_.MemberType -Name $_.Name -Value $_.Value -Force
      }
  }
}
$today = Get-Date
$dcejson = @()
#Filter out Status of Pending + success("Without Issues") include remediation failed devices and exclude devices with last status update more than 7 days ago, Query Intune API to resolve IntuneDeviceID to Device Name (Getting data for alert).
foreach ($resp in $response) 
{
  $timeSpan = New-TimeSpan $resp.lastStateUpdateDateTime $today
  if (($resp.detectionState -notmatch "success" -and $resp.detectionState -notmatch "pending") -or $resp.remediationState -match "remediationFailed" -and $timeSpan.Days -lt 7)
  {
      $intunedeviceid = $resp.id.Split(":")
      $intunedeviceid = $intunedeviceid[1]
      if (!($intunedeviceid))
         {Throw "error in response" }
      $uri3 = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$intunedeviceid"
      $deviceinfo = Invoke-RestMethod -Uri $uri3 -Headers $graphApiHeader -Method Get
      $alert +=  $deviceinfo.deviceName , $deviceinfo.userPrincipalName , $resp.detectionState

      merge $resp $deviceinfo
      $resp | Add-Member -Name "TimeGenerated" -Value $today -MemberType NoteProperty
      $resp.PSObject.Properties.Remove('@odata.context')
      $dcejson += $resp

  }
}
$newdcejson = $dcejson | ConvertTo-Json
}
catch
{
  Throw $_.Exception.Message
}
Write-Output $alert
if ($alert){
$text = "" 
for ($count = 0; $alert[$count]; $count+=3) 
{ $text += "`nDevice Name: $($alert[$count])" + "`nUser: $($alert[$count+1])" + "`nStatus: $($alert[$count+2])" + "<br>"}
#Build JSON body for Teams channel notification
$JSONBody = [PSCustomObject][Ordered]@{
    "@type" = "MessageCard"
    "@context" = "<http://schema.org/extensions>"
    "summary" = "Summary text"
    "themeColor" = '0078D7'
    "title" = "Title text"
    "text" = "$text" 

}
$body = convertto-json $JSONBody -Depth 50

#Send notification to Teams channel & Email if alert has content
if ($alert){Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $body -Uri "https://wermic.webhook.office.com/webhookb2/603f3321bd-b051-4dc8-836b-5e3f51235568@8a94ab74-f14378xxxx-4d02-89c7-ed0f224567644/IncomingWebhook/5xxxxxxx7659161/f8xxxxxx-4bfd-4a32-813e-xxxxxxd619"


##crafting html body for email (Fill in logo or remove line 90, Fill in specific Intune device status URL)
$htmlbody = @"
<html>
  <style>
    body {
      display: flex;
      align-items: center;
      justify-content: center;
      height: 100vh;
    }
    logo {
      width: 550px;
      height: 180px;
    }
    h1 {
      font-size: 16px;
      font-weight: 300;
    }
	h2 {
      font-size: 16px;
      font-weight: 300;
    }
    h3 {
      font-size: 16px;
      font-weight: 300;
    }
    .container {
      display: flex;
      align-items: center;
      justify-content: center;
      flex-direction: column;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="container">
    <img id="logo" src="https://domain.com/images/logo.png" alt="Company Logo">
    <h1>NAME Proactive Remediation script - Detection Status:</h1>
    <h2>$text</h2>
	<h3>See the status on Microsoft Intune Admin Center <a href="https://endpoint.microsoft.com/?ref=AdminCenter#view/Microsoft_Intune_Enrollment/UXAnalyticsScriptMenu/" style="color: #8c8aeb">click here to open in the browser<a/></h3>
  </div>
</body>
</html>

"@
}
}
#region craftig an e-mail (Fill in user to send as, recipient addresses(in the form of a hash table)
$MessageBody = @{
    content = $htmlbody
    contentType = 'HTML'
}
$MailSender = 'user@domain.com'
$recipient = @(@{emailAddress = @{address = 'address@domain.com'}})
#endregion

#region Sending the e-mail (Fill in Subject)
try
{
$NewMessage = New-MguserMessage -UserId $Mailsender -Body $MessageBody -ToRecipients $recipient -Subject 'NAME-Proactive-Remediation-Alerts'
Send-MgUserMessage -UserId $MailSender -Messageid $newmessage.id
#endregion
}
catch
{
  Throw $_.Exception.Message
}
<#
#region upload the data to Azure log analytics
Add-Type -AssemblyName System.Web
$Table = "S1RemediationStatus_CL"



$DceURI = "https://yourloganalyticsuploadendpoint-rm75.westeurope-1.ingest.monitor.azure.com"
$DcrId = "dcr-xxxxxxxxx9017877d2e31f3"

# Connect to Azure with system-assigned managed identity
Connect-AzAccount -Identity | Out-Null
# Retrieving bearer token for the system-assigned managed identity
$bearerToken = (Get-AzAccessToken -ResourceUrl "https://monitor.azure.com//.default").Token
    # Sending the data to Log Analytics via the DCR!
        if ($newdcejson)
        {
        Write-Output $body
        $headers = @{"Authorization" = "Bearer $bearerToken"; "Content-Type" = "application/json" };
        Write-Output $headers
        $uri = "$DceURI/dataCollectionRules/$DcrId/streams/Custom-$Table"+"?api-version=2023-01-01";
        Write-Output $uri
        $uploadResponse = Invoke-RestMethod -Uri $uri -Method "Post" -Body $newdcejson -Headers $headers;
        }
}
#>