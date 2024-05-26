$logpath = "$env:HomeDrive\ProgramData\Microsoft\IntuneManagementExtension\Logs\SystemRestoreRemediation.log"
if ((Get-ItemProperty -Path $logpath).length/1MB -gt 10)
{
    $a = get-date
    $b = $a.Day.ToString() + "." + $a.Month.ToString() + "." + $a.Year.ToString()
    Compress-Archive -Path $logpath -DestinationPath "$logpath.$b.zip"
}
Start-Transcript -append $logpath
try
{
    # Check System Restore status using registry keys
    $restoreKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"

    # Check for existing restore points
    $restorePoints = Get-ComputerRestorePoint

    # If System Restore is disabled or no restore points exist, enable it (not possible with registry)
    if ($restorekey.RPSessionInterval -eq 0 -or $restorePoints.Count -eq 0) 
    {
      Write-Host "System Restore is disabled or there are no restore points."
      Write-Host  "Enabling System Restore"
      Enable-ComputerRestore -Drive "C:\"
      if ($restorePoints.Count -eq 0)
      {
        Write-Host "Creating a Restore Point"
        Checkpoint-Computer -Description "RemediationSysRestorePoint" -RestorePointType MODIFY_SETTINGS
      }
    }
    else 
    {
      Write-Host "System Restore is already enabled with existing restore points."
      exit 0
    }
}
catch
{
    Exit 1
}
Exit 0