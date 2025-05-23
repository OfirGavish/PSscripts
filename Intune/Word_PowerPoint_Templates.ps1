#Variables - Word and PowerPoint Template Files:
$WordTemplateUri = "https://cloudninjastorage.blob.core.windows.net/$web/files/normal.dotm"
$PowerPointTemplateUri = "https://cloudninjastorage.blob.core.windows.net/$web/files/mscloudninjappt.potx"

if (Get-Module -Name Start-IntuneRemediationTranscript)
{
    try 
    {
        Install-Module -Name Start-IntuneRemediationTranscript
    } 
    catch 
    {
        Write-Host "Error installing module: $_"
    }
}
Start-IntuneRemediationTranscript -LogName WordPPTTemplateScript
# Ensure log file size check works correctly
if ((Get-Item $logpath).Length -gt 10MB) {
    $timestamp = (Get-Date -Format "dd.MM.yyyy")
    $zipPath = "$logpath.$timestamp.zip"
    Compress-Archive -Path $logpath -DestinationPath $zipPath
    Write-Host "Log file compressed to $zipPath"
}

# Start transcript after checking/compressing logs
Start-Transcript -Path $logpath -Append

try {
    # Remove template files if they exist
    Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal.dotm" -ErrorAction SilentlyContinue
    Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\blank.potx" -ErrorAction SilentlyContinue
} catch {
    Write-Host "Template files not found or could not be removed."
}

# Download new templates
try {
    $normalJob = Start-Job -ScriptBlock {
        Invoke-WebRequest -Uri $using:WordTemplateUri -OutFile "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Normal.dotm" -ErrorAction Stop
    }
    $blankJob = Start-Job -ScriptBlock {
        Invoke-WebRequest -Uri $using:PowerPointTemplateUri -OutFile "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\blank.potx" -ErrorAction Stop
    }

    # Wait for the jobs to complete and check their status
    $jobs = @($normalJob, $blankJob)
    $jobs | Wait-Job | ForEach-Object {
        if ($_.State -ne "Completed") {
            Write-Host "A download job failed: $($_.State)"
        }
    }
    # Clean up completed jobs
    $jobs | Remove-Job
} catch {
    Write-Host "Files could not be downloaded: $_"
    Start-IntuneRemediationTranscript -Stop
    Exit 1
}
Start-IntuneRemediationTranscript -Stop
Exit 0
