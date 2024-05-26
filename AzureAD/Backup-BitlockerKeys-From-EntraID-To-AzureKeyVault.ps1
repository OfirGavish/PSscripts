#region variables
$keyVaultName = "DeviceCredsBackupVault"


#region connect to services
Connect-AzAccount -CertificateThumbprint "THUMBPRINT" -ApplicationId "APP-ID" -Tenant "TenantID" -ServicePrincipal
Connect-MgGraph -ClientID "ClientID" -TenantId "TenantID" -CertificateThumbprint "THUMBPRINT"
Set-AzContext -SubscriptionName "SUB NAME"

$bitkeys = Get-MgInformationProtectionBitlockerRecoveryKey -All
foreach ($bitkey in $bitkeys)
{
    $deviceid = $bitkey.DeviceId
    $comp =  Get-MgDevice -Filter "DeviceId eq '$($deviceid)'"
    $compname = $comp.DisplayName + "-`BitLockerKeyID-" + $bitkey.id
    $btlckrkey = Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $bitkey.id -Property "key"
    $secretValue = $btlckrkey.Key.ToString()
    $secretValue = ConvertTo-SecureString $secretValue -AsPlainText -Force
    Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $compname -SecretValue $secretValue -ErrorAction Continue

}