##Automatic Script to Delete Stale devices from Azure AD##
#Connect to MgGraph using certificate
$cert = Get-AutomationCertificate -Name 'Graph Powershell cert'
Connect-MgGraph -ClientID "ClientID\APPID" -TenantId "TenantID" -CertificateThumbprint $cert.Thumbprint
#Set the threshold to 90 days ago
$threshold = (Get-Date).AddDays(-90).ToString("MM/dd/yyyy hh:mm:ss")
#Get all Devices
$MGDevices = Get-MgDevice -All | select ApproximateLastSignInDateTime, Id, DisplayNAme, OnPremisesLastSyncDateTime
#delete disabled devices
Foreach ($device in $MGDevices)
        {
          If ($device.ApproximateLastSignInDateTime -lt $threshold -and $device.AccountEnabled -eq "False")
                    {
                        Write-Host $device.DisplayName Deleted
                        Remove-MgDevice -DeviceId $device.Id
                    }
        }
#disable devices with last activity before the threshold
Foreach ($device in $MGDevices)
        {
          If ($device.ApproximateLastSignInDateTime -lt $threshold)
                    {
                        Write-Host $device.DisplayName Deleted
                        Update-MgDevice -DeviceId $device.Id -AccountEnabled:$false
                    }
        }
