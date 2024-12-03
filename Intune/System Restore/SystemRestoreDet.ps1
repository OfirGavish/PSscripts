$logpath = "$env:HomeDrive\ProgramData\Microsoft\IntuneManagementExtension\Logs\SystemRestoreDetection.log"
if ((Get-ItemProperty -Path $logpath).length/1MB -gt 10)
{
    $a = get-date
    $b = $a.Day.ToString() + "." + $a.Month.ToString() + "." + $a.Year.ToString()
    Compress-Archive -Path $logpath -DestinationPath "$logpath.$b.zip"
}
Start-Transcript -append $logpath
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
$datetime = [datetime]::ParseExact($datetimePart, "yyyyMMddHHmmss", $null)

# Format the datetime object to the desired format
$formattedDatetime = $datetime.ToString("MM/dd/yyyy h:mm:ss tt")

$timespan = New-TimeSpan $formattedDatetime $today

# If System Restore is disabled or no restore points exist, enable it (not possible with registry)
if ($restorekey.RPSessionInterval -eq 0 -or $restorePoints.Count -eq 0 -or $timespan.Days -ge 30)
{
    Exit 1
}
Exit 0