<#
.SYNOPSIS
This script is located at https://github.com/OfirGavish/PSscripts/edit/main/Exchange%20Online/EXOMailboxSize.ps1.

.DESCRIPTION
This script is used for detection of Exchange Online Mailboxes with over 70 GB of Used space.
If the mailbox is over 70GB, the script will check if the mailbox is over 90GB and if it is, it will enable the Archive and AutoExpandingArchive for that mailbox.

.PARAMETER None
There are no parameters for this script.

.EXAMPLE
Example usage of the script:
    run in azure automation runbook, make sure to assign the correct permissions to the managed identity.
    email is sent using Azure Email Communication Service, feel free to change the email settings.
    You can export the information to a CSV file and send it as an attachment, uncomment line 65 and add the attachment flag to the Send-MailMessage command.

.NOTES
Author: Ofir Gavish
Special Thanks to Ram Apter @ https://github.com/Rmap91
Date: 09/05/2024
#>

#Connect to Exchange Online
Connect-ExchangeOnline -ManagedIdentity -Organization domain.onmicrosoft.com
#Get all mailboxes statistics
$mailbox = Get-EXOMailbox -ResultSize unlimited | Get-EXOMailboxStatistics
#initialize alert array
$alert = @()
#Loop through all mailboxes, check if the mailbox is over 70GB, if it is, check if it is over 90GB and enable Archive and AutoExpandingArchive
try {
    foreach ($ml in $mailbox) {
        if ($ml.TotalItemSize.Value -gt 70GB) {
            $AE = Get-ExoMailbox $ml -PropertySets Archive
            if ($ml.TotalItemSize.Value -gt 90GB -and !$AE.AutoExpandingArchiveEnabled) 
            {
                try 
                {
                    if ($AE.ArchiveStatus -eq "None")
                    {
                        Enable-Mailbox $ml -Archive
                    }
                    Enable-Mailbox $ml -AutoExpandingArchive
                    $AE = Get-ExoMailbox $ml -PropertySets Archive
                }
                catch {
                    Throw "Failed to Enable Archive or AutoExpandingArchive: $_"
                }
            }
            #Add the mailbox to the alert array with the relevant information
            $alert += [PSCustomObject]@{
                DisplayName = $ml.DisplayName
                TotalDeletedItemSize = $ml.TotalDeletedItemSize
                TotalItemSize = $ml.TotalItemSize
                AutoExpandingArchiveEnabled = $AE.AutoExpandingArchiveEnabled
            }
        }
    }
}
catch {
    Throw "Failed to run the foreach loop: $_"
}

#Export the alert array to a CSV file
#$alert | Export-Csv -Path "$env:Temp\test23324.csv" -NoTypeInformation

# Create HTML table from $alert
$htmlTable = $alert | ConvertTo-Html -Property DisplayName, TotalDeletedItemSize, TotalItemSize, AutoExpandingArchiveEnabled -As Table

# Create HTML body for the email
$htmlBody = @"
<html>
<head>
<style>
    table {
        border-collapse: collapse;
        width: 100%;
    }
    th, td {
        text-align: left;
        padding: 8px;
        border-bottom: 1px solid #ddd;
    }
    th {
        background-color: #f2f2f2;
    }
</style>
</head>
<body>
<h2>Alert Report</h2>
$htmlTable
</body>
</html>
"@

#Send email with the alert information
$from = "mailboxalerts@domain.com"
$to = "EXOTeam@domain.com"
$cred = Get-AutomationPSCredential -Name 'CredentialName'
try
{
    Send-MailMessage -To $to -From $from -Credential $cred -Subject "Mailboxes over 70GB Alerts" -Body $htmlbody -BodyAsHtml -SmtpServer "smtp.azurecomm.net" -Port 587 -UseSsl
}
catch
{
    throw $Error.Exception.Message
}
#Disconnect from Exchange Online
Disconnect-ExchangeOnline
