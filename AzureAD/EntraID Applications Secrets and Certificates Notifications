<#
Written by:  Ofir Gavish
Date:        2024-16-10

.SYNOPSIS
This script checks Azure AD applications for expiring secrets and certificates. 
It sends email notifications for secrets and certificates that are expiring within 
30, 10, or 5 days.

.DESCRIPTION
The script retrieves all Azure AD applications and their associated secrets and 
certificates. It calculates the remaining days until expiration and sends an 
email notification if any secrets or certificates are set to expire within 
the specified thresholds. The script can be scheduled to run automatically.
The Script uses Azure Email Communication Service to send the email, Authenticating with the Automation Account Managed-Identity
Edit the "Sender Address" at line 132, and the "Replay to" address and name at line 148 and 149.

.PARAMETER DaysUntilExpiration
Defines the threshold for filtering applications by expiration days.

.PARAMETER IncludeAlreadyExpired
Specifies whether to include applications with already expired secrets and certificates.

.EXAMPLE
Run the script to receive notifications for any expiring secrets or certificates.

#>
Connect-MgGraph -ManagedIdentity

# Configuration
$DaysUntilExpiration = 30  # Change this as needed
$IncludeAlreadyExpired = 'No'  # Set to 'Yes' if you want to include already expired secrets
$emailSubject = "Important: Entra ID Secrets and Certificates Notification"
$EmailRecipient = "IT_Team@domain.com"
# Define the communication endpoint URL
$communicationendpointurl = "domain.region.communication.azure.com" # Update with your communication endpoint URL

$Now = Get-Date
$Applications = Get-MgApplication -all

foreach ($App in $Applications) {
    $AppName = $App.DisplayName
    $AppID   = $App.Id

    $AppCreds = Get-MgApplication -ApplicationId $AppID |
        Select-Object PasswordCredentials, KeyCredentials

    # Check password credentials
    foreach ($Secret in $AppCreds.PasswordCredentials) {
        $RemainingDaysCount = ($Secret.EndDateTime - $Now).Days

        if (($IncludeAlreadyExpired -eq 'No' -and $RemainingDaysCount -ge 0) -or 
            ($IncludeAlreadyExpired -eq 'Yes')) {
            if ($RemainingDaysCount -le $DaysUntilExpiration) {
                if ($RemainingDaysCount -eq 30 -or $RemainingDaysCount -eq 10 -or $RemainingDaysCount -eq 5) {
                    $alert += "Warning: The secret '$($Secret.DisplayName)' for application '$AppName' will expire in $RemainingDaysCount days.`n"
                }
            }
        }
    }

    # Check key credentials
    foreach ($Cert in $AppCreds.KeyCredentials) {
        $RemainingDaysCount = ($Cert.EndDateTime - $Now).Days

        if (($IncludeAlreadyExpired -eq 'No' -and $RemainingDaysCount -ge 0) -or 
            ($IncludeAlreadyExpired -eq 'Yes')) {
            if ($RemainingDaysCount -le $DaysUntilExpiration) {
                if ($RemainingDaysCount -eq 30 -or $RemainingDaysCount -eq 10 -or $RemainingDaysCount -eq 5) {
                    $alert += "Warning: The certificate '$($Cert.DisplayName)' for application '$AppName' will expire in $RemainingDaysCount days.`n"
                }
            }
        }
    }
}

#Construct the email body
$emailBody = @"
<html>
<body>
<p>Dear User,</p>
<p>This is to inform you that a <b><i>Entra ID Secrets and Certificates are close to expiration</i></b>.</p>
<p>Please take necessary action to renew the secrets and certificates to avoid any service disruption.</p>
<p>Below are the details of the secrets and certificates that are close to expiration:</p>
<p>$alert</p>
</body>
</html>
"@

# Send email if there are any warnings
if ($alert) {
    # Define the resource ID for Azure Communication Services
    $ResourceID = 'https://communication.azure.com'

    # Construct the URI for the identity endpoint
    $Uri = "$($env:IDENTITY_ENDPOINT)?api-version=2018-02-01&resource=$ResourceID"

    # Debug output
    # Print the constructed URI and headers
    Write-Output "URI: $Uri"
    Write-Output "Headers: @{ Metadata = 'true' }"

    # Try to get the access token
    try {
        # Invoke a GET request to the identity endpoint to get the access token
        $AzToken = Invoke-WebRequest -Uri $Uri -Method GET -Headers @{ Metadata = "true" } -UseBasicParsing | Select-Object -ExpandProperty Content | ConvertFrom-Json | Select-Object -ExpandProperty access_token
        # Print the obtained access token
        Write-Output "Access Token: $AzToken"
    }
    catch {
        # If there's an error, print the error message and response details
        Write-Error "Failed to get access token: $_"
        Write-Output "Response Status Code: $($_.Exception.Response.StatusCode.Value__)"
        Write-Output "Response Status Description: $($_.Exception.Response.StatusDescription)"
        Write-Output "Response Content: $($_.Exception.Response.GetResponseStream() | %{ $_.ReadToEnd() })"
    }

    # Construct the URI for the email sending endpoint
    $uri = "https://$communicationendpointurl/emails:send?api-version=2023-03-31"

    # Define the headers for the REST API call
    # Include the content type and the obtained access token in the Authorization header
    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $AzToken"
    }
    # Define the body for the REST API call
    $apiResponse = @{
        headers                        = @{
            id = (New-Guid).Guid
        }
        senderAddress                  = 'EntraAppNotifications@domain.com'
        content                        = @{
            subject = $emailSubject
            html    = $emailBody
        }
        recipients                     = @{
            to = @(
                @{
                    address     = $EmailRecipient
                    displayName = $EmailRecipient
                }
            )
        }

        replyTo                        = @(
            @{
                address     = "IT_Team@domain.com"
                displayName = "IT_Team"
            }
        )
        userEngagementTrackingDisabled = $true
    }
                
    # Convert the PowerShell object to JSON
    $body = $apiResponse | ConvertTo-Json -Depth 10
    # Send the email
    try {
        # Log the request details
        Write-Output "Sending email..."
        Write-Output "URI: $uri"
        Write-Output "Headers: $headers"
        Write-Output "Body: $body"
        # Make the request
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -UseBasicParsing
        # Log the response
        Write-Output "Response: $response"
        # Return the response
        $response
    }
    catch {
        # Log the error
        Write-Error "Failed to send email: $_"
        Write-Output "Exception Message: $($_.Exception.Message)"
        Write-Output "Exception StackTrace: $($_.Exception.StackTrace)"
    }
}
