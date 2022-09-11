#Script to create 20 users and a Security group, join all users to the security group in Azure AD
$PSNativeCommandUseErrorActionPreference = $true
$ErrorActionPreference = 'Stop'
#connect to MgGraph
Connect-MgGraph -ClientID "0fc43f2b-697d-4224-2ffd-8c759a0841da" -TenantId "2a94ab54-f3c3-4d09-82c9-ed0f57d1c244" -CertificateThumbprint "854679B1BD618FF3A0372724DBA24D3074E20E80"
#build log folder + file
BEGIN { 
        
        $errorlogfile = "$home\Documents\PSlogs\Error_Log.txt"
        $errorlogfolder = "$home\Documents\PSlogs"
        
        if  ( !( Test-Path -Path $errorlogfolder -PathType "Container" ) ) {
            
            Write-Verbose "Create error log folder in: $errorlogfolder"
            New-Item -Path $errorlogfolder -ItemType "Container" -ErrorAction Stop
        
            if ( !( Test-Path -Path $errorlogfile -PathType "Leaf" ) ) {
                Write-Verbose "Create error log file in folder $errorlogfolder with name Error_Log.txt"
                New-Item -Path $errorlogfile -ItemType "File" -ErrorAction Stop
            }
        }
}
function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length,
        [int] $amountOfNonAlphanumeric = 1
    )
    Add-Type -AssemblyName 'System.Web'
    return [System.Web.Security.Membership]::GeneratePassword($length, $amountOfNonAlphanumeric)
}
#Create the security group
New-MgGroup -DisplayName 'Assignment Group' -MailEnabled:$False  -MailNickName 'AssignmentGroup' -SecurityEnabled
$Group = Get-MgGroup -Filter "DisplayName eq 'Assingment Group'"
#create a loop to count 20 nunmbers
[int] $UserNumber = 1
while ($UserNumber -lt 21) {
    #build passwords
    $password = Get-RandomPassword 12
    $PasswordProfile = @{
    Password = $password
    }
    #creates 20 users with number count in the names
    $newuser="user"+$usernumber
    $useremail=$newuser+"@contoso.com"
    New-MgUser -DisplayName $newuser -PasswordProfile $PasswordProfile -AccountEnabled -MailNickname $newuser -UserPrincipalName $useremail
    $createduser = Get-MgUser -Filter "DisplayName eq '$newuser'"
    try {
    $timestamp = Get-Date
    #adds the users to the group
    New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $createduser.Id
    "$createduser.UserPrincipalName, $timestamp, result:Success" | Out-File $errorlogfile -Append
    }
    #error checking
    catch ("Insufficient privileges") {
                    $timestamp = Get-Date

                    $error = Write-Error $_       
                    $body = "{'text':'Insufficient privileges error - 20 new users script'}"
                    Invoke-RestMethod -Method post -Body $body -Uri "https://wermic.webhook.office.com/webhookb2/603fd0bd-b051-4dc8-836b-5e3f51b15518@8a94ab74-f3c6-4d02-89c7-ed0f47d6c544/IncomingWebhook/593d507d317a4796baf1b53ec7659161/f81641ed-4bfd-4a32-813e-5ab9b1e7d619"
                    
                    #region craftig an e-mail
                    $SendingDateTime = "{0:G}" -f (Get-Date)
                    $MessageBody = @{
                        content = "$error"
                        contentType = 'HTML'
                    }
                    $MailSender = 'AutomationSP@contoso.com'
                    $recipient = @(@{emailAddress = @{address = 'devops@contoso.com'}},@{emailAddress = @{address = 'cloudops@contoso.com'}})
                    #endregion
                    
                    #region Sending the e-mail
                    $NewMessage = New-MguserMessage -UserId $Mailsender -Body $MessageBody -ToRecipients $recipient -Subject 'Insufficient privileges error - 20 new users script'
                    Send-MgUserMessage -UserId $MailSender -Messageid $newmessage.id
                    #endregion
                    
            "$createduser.UserPrincipalName, $timestamp, result:Failure" | Out-File $errorlogfile -Append        
                    
    }
    catch ("Cannot Update a mail-enabled security groups") {
        $timestamp = Get-Date

        Write-Warning "==> Error adding $newuser via Graph, failing back to ExO" 
        try {
            $timestamp = Get-Date

            Connect-ExchangeOnline -CertificateThumbPrint "012THISISADEMOTHUMBPRINT" -AppID "36ee4c6c-0812-40a2-b820-b22ebd02bce3" -Organization "contoso.onmicrosoft.com"
            Add-DistributionGroupMember -Identity "Assingment Group" -Member $createduser.mail -Confirm:$false -BypassSecurityGroupManagerCheck -ErrorAction Stop
            "$createduser.UserPrincipalName, $timestamp, result:Success" | Out-File $errorlogfile -Append
        } catch {
            $timestamp = Get-Date

            Write-Error "==> Error adding $newuser via ExO"
            "$createduser.UserPrincipalName, $timestamp, result:Failure" | Out-File $errorlogfile -Append
        }
    "$createduser.UserPrincipalName, $timestamp, result:Failure" | Out-File $errorlogfile -Append
    }

    
    catch (Exception e) {
                    $timestamp = Get-Date

                    "$createduser.UserPrincipalName, $timestamp, result:Failure" | Out-File $errorlogfile -Append
                    $error = Write-Error $_       
                    $body = "{'text':'unknown exeption - 20 new users script'}"
                    Invoke-RestMethod -Method post -Body $body -Uri "https://wermic.webhook.office.com/webhookb2/603fd0bd-b051-4dc8-836b-5e3f51b15518@8a94ab74-f3c6-4d02-89c7-ed0f47d6c544/IncomingWebhook/593d507d317a4796baf1b53ec7659161/f81641ed-4bfd-4a32-813e-5ab9b1e7d619"
                    
                    #region craftig an e-mail
                    $SendingDateTime = "{0:G}" -f (Get-Date)
                    $MessageBody = @{
                        content = "$error"
                        contentType = 'HTML'
                    }
                    $MailSender = 'AutomationSP@contoso.com'
                    $recipient = @(@{emailAddress = @{address = 'devops@contoso.com'}},@{emailAddress = @{address = 'cloudops@contoso.com'}})
                    #endregion
                    
                    #region Sending the e-mail
                    $NewMessage = New-MguserMessage -UserId $Mailsender -Body $MessageBody -ToRecipients $recipient -Subject 'Unknown execption - 20 new users script'
                    Send-MgUserMessage -UserId $MailSender -Messageid $newmessage.id
                    #endregion

    }
$UserNumber += 1 # Increase the value stored in the variable by 1.
}
