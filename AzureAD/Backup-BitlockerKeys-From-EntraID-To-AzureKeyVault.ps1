#region variables
$keyVaultName = "BitLockerKeysBackupVault"

#region connect to services
Connect-AzAccount -Identity
Connect-MgGraph -ManagedIdentity
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
