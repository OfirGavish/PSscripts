#region variables
$secret = "GoodEnoughPassPhrase2024%@wait!itsnotlongenoughishouldaddafewmorewordsandnumbers38^#"
$keyVaultName = "DeviceCredsBackupVault"


#region connect to services
Connect-AzAccount -CertificateThumbprint "THUMBPRINT" -ApplicationId "APP-ID" -Tenant "TenantID" -ServicePrincipal
Connect-MgGraph -ClientID "ClientID" -TenantId "TenantID" -CertificateThumbprint "THUMBPRINT"
Set-AzContext -SubscriptionName "SUB NAME"

#region create a xor function
Add-Type @'
  public class MyXORString {
    public static System.String XORString(
      System.String source
      ,System.String key
    ) {
      System.Text.StringBuilder sb = new System.Text.StringBuilder();
      System.Int32 sourceLen = source.Length;
      System.Int32 keyLen = key.Length;
      System.Int32 sourceIdx = 0;
      System.Int32 keyIdx = 0;
      System.Char workingChar;
      while ( sourceIdx < sourceLen ) {
        workingChar = (System.Char)( source[sourceIdx] ^ key[keyIdx] );
        sb.Append( workingChar );
 
        sourceIdx += 1;
        keyIdx += 1;
        if ( keyIdx >= keyLen ) {
          keyIdx = 0;
        }
      }
 
      return sb.ToString();
    }
  }
'@
 
#region get all devices, and if laps password or bitlocker recovery key exist upload the xored password to keyvault
$devices = Get-MgDevice -All
foreach ($device in $devices)
{
    $lapspass = Get-LapsAADPassword -DeviceIds $device.DeviceId -IncludePasswords -AsPlainText | select "Password"
    if ($lapspass)
    {
        $comp = $device.DisplayName
        $xor = [MyXORString]::XORString( "$secret", "$lapspass" );
        #Upload xor and displayname to keyvault
        $secretValue = ConvertTo-SecureString -String $xor -AsPlainText -Force
        Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $comp -SecretValue $secretValue
    }
    else
    {
        "no LAPS password found for $comp" >> C:\Temp\devicecredsbackuplog.txt
    }
    $compid = $device.DeviceId.ToString()
    $btlck = Get-MgInformationProtectionBitlockerRecoveryKey -Filter "DeviceId eq '$($compid)'"
    if ($btlck)
    {
        $btlckrkey = Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $btlck.id -Property "key"
        $xor2 = [MyXORString]::XORString( "$secret", "$btlckrkey + $btlck.id" );
        $secretname = $device.DisplayName + "-`BitLockerKeyID-" + $btlck.id
        #Upload xor and displayname to keyvault
        $secretValue = ConvertTo-SecureString -String $xor2 -AsPlainText -Force
        Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretname -SecretValue $secretValue
    }
    else
    {
        "no bitlocker key found for $comp" >> C:\Temp\devicecredsbackuplog.txt
    }
}