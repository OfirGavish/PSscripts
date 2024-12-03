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

    $today = Get-Date

    # Input string
    $string = $restorePoints.CreationTime

    # Extract the datetime part without milliseconds and timezone offset
    $datetimePart = $string.Substring(0, 14)

    # Convert the datetime part to a DateTime object
    $datetime = [datetime]::ParseExact($datetimePart[-1], "yyyyMMddHHmmss", $null)

    # Format the datetime object to the desired format
    $formattedDatetime = $datetime.ToString("MM/dd/yyyy h:mm:ss tt")

    $timespan = New-TimeSpan $formattedDatetime $today

    # If System Restore is disabled or no restore points exist, enable it (not possible with registry)
    if ($restorekey.RPSessionInterval -eq 0 -or $restorePoints.Count -eq 0 -or $timespan.Days -ge 30) 
    {
      Write-Host "System Restore is disabled or there are no restore points."
      Write-Host  "Enabling System Restore"
      Enable-ComputerRestore -Drive "C:\"
      if ($restorePoints.Count -eq 0 -or $timespan.Days -ge 30)
      {
        Write-Host "Creating a Restore Point"
        Checkpoint-Computer -Description "RemediationSysRestorePoint" -RestorePointType MODIFY_SETTINGS
      }
    }
    else 
    {
      Write-Host "restore points count: $(($restorePoints | Measure-Object).Count) , The latest restore point creation time: $($formattedDatetime)"
      Write-Host "System Restore is already enabled with existing restore points."
      Stop-Transcript
      Exit 0
    }
}
catch
{
    Stop-Transcript
    Exit 1
}
Stop-Transcript
Exit 0