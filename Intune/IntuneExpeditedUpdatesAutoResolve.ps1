# IntuneExpeditedUpdatesAutoResolve.ps1
# Azure Automation Runbook - Queries Intune Expedited Update Status for all policies and deploys remediation script

# --- VARIABLES: Set these before running --- #
$TenantId       = "<YOUR_TENANT_ID>"
$ClientId       = "<YOUR_CLIENT_ID>"
$ClientSecret   = "<YOUR_CLIENT_SECRET>"
$RemediationGroupName = "Intune-ExpeditedUpdateRemediation-Devices"
$ScriptName     = "ExpeditedUpdateRemediation"
$MaxRetryCount  = 3
$RetryDelay     = 5

# --- END OF VARIABLES --- #

# Function for formatted logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "[$timestamp] [$Level] $Message"
}

Write-Log "Starting Intune Expedited Updates Auto-Resolve runbook"

try {
    # Authenticate with Microsoft Graph
    Write-Log "Authenticating to Microsoft Graph API..."
    $SecureClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $SecureClientSecret -Scopes "DeviceManagementManagedDevices.Read.All","DeviceManagementConfiguration.ReadWrite.All","Group.ReadWrite.All","Directory.Read.All"
    Write-Log "Successfully authenticated to Microsoft Graph API"

    # Get all Windows Quality Update (Expedited) policies
    Write-Log "Retrieving all Windows Quality Update (Expedited) policies..."
    $policies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles" -Method GET
    Write-Log "Found $($policies.value.Count) expedited update policies"

    $devicesWithErrors = @()

    foreach ($policy in $policies.value) {
        Write-Log "Processing policy: $($policy.displayName) (ID: $($policy.id))"

        # Start Export Job per policy
        Write-Log "Starting export job for policy: $($policy.displayName)"
        $exportJobBody = @{
            reportName = "QualityUpdateDeviceStatusByPolicy"
            filter     = "(PolicyId eq '$($policy.id)')"
            format     = "json"
        } | ConvertTo-Json

        $exportJob = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs" -Body $exportJobBody
        $jobId = $exportJob.id
        Write-Log "Export job created with ID: $jobId for policy: $($policy.displayName)"

        # Wait for job completion with timeout
        $statusUri = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$jobId')"
        $attempt = 0
        Write-Log "Waiting for export job to complete for policy: $($policy.displayName)..."
        do {
            Start-Sleep -Seconds 5
            $attempt++
            $jobStatus = Invoke-MgGraphRequest -Method GET -Uri $statusUri
            Write-Log "Policy $($policy.displayName) - job status check $attempt`: $($jobStatus.status)"
        } until ($jobStatus.status -eq "completed" -or $attempt -gt 60) # Add timeout after 5 minutes

        if ($jobStatus.status -ne "completed") {
            Write-Log "Export job for policy $($policy.displayName) did not complete in the expected time frame" -Level "WARNING"
            continue # Skip to next policy
        }

        # Download and parse report
        Write-Log "Downloading and extracting the report for policy: $($policy.displayName)..."
        $tempZip = "$env:TEMP\$($policy.displayName)-Report.zip"
        Invoke-WebRequest -Uri $jobStatus.url -OutFile $tempZip
        Expand-Archive -Path $tempZip -DestinationPath "$env:TEMP\$($policy.displayName)-Report" -Force

        # Find the report file using a wildcard pattern
        $reportFilePattern = "$env:TEMP\$($policy.displayName)-Report\QualityUpdateDeviceStatusByPolicy*.json"
        $reportFile = Get-ChildItem -Path $reportFilePattern -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($reportFile -and (Test-Path $reportFile.FullName)) {
            $reportData = Get-Content $reportFile.FullName | ConvertFrom-Json
            Write-Log "Report downloaded and parsed successfully for policy: $($policy.displayName)"

            # Fix: Update filter to match the actual error condition in the JSON data
            # Devices with errors have CurrentDeviceUpdateStatus=10 and LatestAlertMessage=50
            $policyDevicesWithErrors = $reportData.values | Where-Object { 
                $_.CurrentDeviceUpdateStatus -eq 10 -or 
                $_.AggregateState -eq "Error" 
            }
            Write-Log "Found $($policyDevicesWithErrors.Count) devices with errors for policy: $($policy.displayName)"
            
            foreach ($device in $policyDevicesWithErrors) {
                $device | Add-Member -NotePropertyName PolicyName -NotePropertyValue $policy.displayName
                $devicesWithErrors += $device
                Write-Log "Device with error: $($device.DeviceName) - $($device.LatestAlertMessage) - Policy: $($policy.displayName)"
            }
        } else {
            Write-Log "Report file not found for policy: $($policy.displayName)" -Level "WARNING"
        }
    }

    Write-Log "Total devices across all policies with errors: $($devicesWithErrors.Count)"

    if ($devicesWithErrors.Count -eq 0) {
        Write-Log "No devices need remediation. Script execution complete."
        exit
    }

    # Create/Update Remediation Group
    Write-Log "Checking for existing remediation group..."
    $existingGroup = Get-MgGroup -Filter "displayName eq '$RemediationGroupName'"
    if (!$existingGroup) {
        Write-Log "Creating new remediation group: $RemediationGroupName"
        $groupParams = @{
            displayName = $RemediationGroupName
            mailEnabled = $false
            mailNickname = "ExpeditedUpdateGroup"
            securityEnabled = $true
        }
        $group = New-MgGroup -BodyParameter $groupParams
        Write-Log "Created remediation group with ID: $($group.Id)"
    } else {
        $group = $existingGroup
        Write-Log "Using existing remediation group with ID: $($group.Id)"
    }

    # Add devices to remediation group
    Write-Log "Adding devices to remediation group..."
    $devicesAdded = 0
    foreach ($device in $devicesWithErrors) {
        try {
            $managedDevice = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $device.DeviceId
            $aadDeviceId = $managedDevice.AzureAdDeviceId
            if ($aadDeviceId) {
                # Get the Azure AD device object to obtain its Object ID
                $aadDevice = Get-MgDevice -Filter "DeviceId eq '$aadDeviceId'"
                if ($aadDevice) {
                    $deviceObjectId = $aadDevice.Id  # This is the Object ID we need
                    $params = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$deviceObjectId"
                    }
                    New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $params
                    Write-Log "Added device $($device.DeviceName) (ID: $deviceObjectId) to remediation group - Policy: $($device.PolicyName)"
                    $devicesAdded++
                } else {
                    Write-Log "Could not find Azure AD device object for device $($device.DeviceName)" -Level "WARNING"
                }
            } else {
                Write-Log "Could not find Azure AD device ID for device $($device.DeviceName)" -Level "WARNING"
            }
        } catch {
            Write-Log "Error adding device $($device.DeviceName) to group: $_" -Level "ERROR"
        }
    }

    Write-Log "Added $devicesAdded devices to remediation group"

    # Define your remediation script here
    Write-Log "Preparing remediation script..."
    $ScriptContent = @'
# Remediation script for Intune expedited update issues
$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\ExpeditedUpdateRemediation.log"

function Write-RemediationLog {
    param (
        [string]$Message
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$TimeStamp - $Message" | Out-File -FilePath $LogFile -Append
}

Write-RemediationLog "Starting expedited update remediation"

# Check Windows Update Health Service status
Write-RemediationLog "Checking Windows Update Health Service (uhssvc) status..."
$wus = Get-Service uhssvc -ErrorAction SilentlyContinue

if (-not $wus) {
    Write-RemediationLog "Windows Update Health Service not found - may need to be installed"
    $needsInstall = $true
} elseif ($wus.Status -ne "Running") {
    Write-RemediationLog "Windows Update Health Service exists but is not running (Status: $($wus.Status))"
    Write-RemediationLog "Attempting to enable and start the service..."
    Set-Service uhssvc -StartupType Automatic -ErrorAction SilentlyContinue
    Start-Service uhssvc -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    $wus = Get-Service uhssvc
    if ($wus.Status -ne "Running") {
        Write-RemediationLog "Failed to start service - need to reinstall"
        $needsInstall = $true
    } else {
        Write-RemediationLog "Successfully started Windows Update Health Service"
    }
} else {
    Write-RemediationLog "Windows Update Health Service is running correctly"
}

# Install/reinstall Windows Update Health Service if needed
if ($needsInstall) {
    Write-RemediationLog "Installing Windows Update Health Service..."
    try {
        # Create temp directory if it doesn't exist
        if (-not (Test-Path "C:\Temp")) {
            Write-RemediationLog "Creating C:\Temp directory"
            New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null
        }
        
        # Detect Windows version and download appropriate MSI
        $osVersion = (Get-WmiObject win32_operatingsystem).Version
        Write-RemediationLog "Detected Windows version: $osVersion"
        
        # Microsoft now provides a single MSI that works for both Windows 10 and Windows 11
        # Direct download link from Microsoft for Windows Update Health Tools
        $msiUrl = "https://download.microsoft.com/download/f/1/d/f1d7042a-c5e5-474e-a423-25ba31b35c7b/WindowsUpdateHealthTools.msi"
        $msiFile = "C:\Temp\WindowsUpdateHealthTools.msi"
        Write-RemediationLog "Using official Microsoft Windows Update Health Tools MSI"
        
        # Download and install the MSI
        Write-RemediationLog "Downloading MSI from $msiUrl"
        
        try {
            # Use TLS 1.2 for the download
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $msiUrl -OutFile $msiFile -UseBasicParsing
            
            if (Test-Path $msiFile) {
                Write-RemediationLog "MSI downloaded successfully"
                
                Write-RemediationLog "Installing MSI: $msiFile"
                $installResult = (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiFile`" /qn" -Wait -PassThru).ExitCode
                Write-RemediationLog "MSI installation completed with exit code: $installResult"
                
                # Check if service is running after install
                $count = 0
                $maxAttempts = 10
                Write-RemediationLog "Waiting for service to start (checking up to $maxAttempts times)..."
                
                do {
                    Start-Sleep -Seconds 60
                    $wus = Get-Service uhssvc -ErrorAction SilentlyContinue
                    $count++
                    Write-RemediationLog "Check $count/$maxAttempts - Service status: $(if($wus){"$($wus.Status)"}else{"Not found"})"
                } while (($wus -eq $null -or $wus.Status -ne "Running") -and ($count -lt $maxAttempts))
                
                if ($wus -and $wus.Status -eq "Running") {
                    Write-RemediationLog "Windows Update Health Service started successfully after installation"
                } else {
                    Write-RemediationLog "Failed to start Windows Update Health Service after installation"
                }
            } else {
                Write-RemediationLog "Failed to download MSI file" -Level "ERROR"
            }
        } catch {
            Write-RemediationLog "Error during download or installation: $($_.Exception.Message)"
        }
    }
    catch {
        Write-RemediationLog "Error during Windows Update Health Service installation: $($_.Exception.Message)"
    }
}

# Check for pending reboots
$PendingReboot = $false

# Method 1: Check Component-Based Servicing
if (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA SilentlyContinue) {
    Write-RemediationLog "Component-Based Servicing indicates a pending reboot"
    $PendingReboot = $true
}

# Method 2: Check Windows Update
if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA SilentlyContinue) {
    Write-RemediationLog "Windows Update indicates a pending reboot"
    $PendingReboot = $true
}

# Method 3: Check PendingFileRenameOperations
$PendingFileRename = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA SilentlyContinue
if ($PendingFileRename -and $PendingFileRename.PendingFileRenameOperations) {
    Write-RemediationLog "PendingFileRenameOperations indicates a pending reboot"
    $PendingReboot = $true
}

# If pending reboot detected, schedule it for off-hours
if ($PendingReboot) {
    Write-RemediationLog "Pending reboot detected, scheduling for tonight at 2 AM"
    
    # Schedule reboot for 2 AM
    $currentDate = Get-Date
    $rebootTime = Get-Date -Hour 2 -Minute 0 -Second 0
    
    # If it's already past 2 AM, schedule for tomorrow
    if ($currentDate -gt $rebootTime) {
        $rebootTime = $rebootTime.AddDays(1)
    }
    
    $timeSpan = $rebootTime - $currentDate
    $secondsUntilReboot = [math]::Ceiling($timeSpan.TotalSeconds)
    
    Write-RemediationLog "Scheduling reboot for $rebootTime (in $secondsUntilReboot seconds)"
    shutdown.exe /r /f /t $secondsUntilReboot /c "Scheduled reboot to complete Windows updates"
}
else {
    # Try to reset Windows Update components
    Write-RemediationLog "No pending reboot detected, resetting Windows Update components"
    
    # Stop Windows Update services
    Write-RemediationLog "Stopping Windows Update services"
    Stop-Service -Name BITS, wuauserv, appidsvc, cryptsvc -Force
    
    # Delete Windows Update cache
    Write-RemediationLog "Clearing Windows Update cache"
    Remove-Item "$env:SystemRoot\SoftwareDistribution\*" -Recurse -Force -EA SilentlyContinue
    
    # Reset Windows Update components
    Write-RemediationLog "Resetting Windows Update components"
    & "$env:SystemRoot\System32\wuauclt.exe" /resetauthorization /detectnow
    
    # Restart Windows Update services
    Write-RemediationLog "Starting Windows Update services"
    Start-Service -Name BITS, wuauserv, appidsvc, cryptsvc
    
    # Force Windows Update to check for updates
    Write-RemediationLog "Forcing Windows Update check"
    & "$env:SystemRoot\System32\UsoClient.exe" StartScan
}

Write-RemediationLog "Expedited update remediation completed"
'@

# Add custom functions for working with device management scripts
Function Get-DeviceManagementScripts(){
    <#
    .SYNOPSIS
    This function is used to get device management scripts from the Graph API REST interface
    .DESCRIPTION
    The function connects to the Graph API Interface and gets any device management scripts
    .EXAMPLE
    Get-DeviceManagementScripts
    Returns any device management scripts configured in Intune
    Get-DeviceManagementScripts -ScriptId $ScriptId
    Returns a device management script configured in Intune
    #>

    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$false)]
        $ScriptId
    )

    $graphApiVersion = "Beta"
    $Resource = "deviceManagement/deviceManagementScripts"
    
    try {
        if($ScriptId){
            $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$ScriptId"
            Invoke-MgGraphRequest -Uri $uri -Method Get
        }
        else {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$expand=groupAssignments"
            (Invoke-MgGraphRequest -Uri $uri -Method Get).Value
        }
    }
    catch {
        Write-Log "Error retrieving device management scripts: $_" -Level "ERROR"
        throw
    }
}

Function New-DeviceManagementScript(){
    <#
    .SYNOPSIS
    This function is used to create a new device management script in Intune
    .DESCRIPTION
    This function creates a new device management script in Intune
    #>

    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$true)]
        $ScriptParams
    )

    $graphApiVersion = "Beta"
    $Resource = "deviceManagement/deviceManagementScripts"
    
    try {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
        Invoke-MgGraphRequest -Uri $uri -Method Post -Body $ScriptParams
    }
    catch {
        Write-Log "Error creating device management script: $_" -Level "ERROR"
        throw
    }
}

Function Update-DeviceManagementScript(){
    <#
    .SYNOPSIS
    This function is used to update an existing device management script in Intune
    .DESCRIPTION
    This function updates an existing device management script in Intune
    #>

    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$true)]
        $ScriptId,
        
        [Parameter(Mandatory=$true)]
        $ScriptParams
    )

    $graphApiVersion = "Beta"
    $Resource = "deviceManagement/deviceManagementScripts"
    
    try {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$ScriptId"
        Invoke-MgGraphRequest -Uri $uri -Method Patch -Body $ScriptParams
        
        # Return the updated script
        Get-DeviceManagementScripts -ScriptId $ScriptId
    }
    catch {
        Write-Log "Error updating device management script: $_" -Level "ERROR"
        throw
    }
}

Function New-DeviceManagementScriptAssignment(){
    <#
    .SYNOPSIS
    This function is used to assign a device management script to a group in Intune
    .DESCRIPTION
    This function assigns a device management script to a group in Intune
    #>

    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$true)]
        $ScriptId,
        
        [Parameter(Mandatory=$true)]
        $AssignmentParams
    )

    $graphApiVersion = "Beta"
    $Resource = "deviceManagement/deviceManagementScripts/$ScriptId/assign"
    
    try {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
        
        # Format for the assignment differs from the original parameter
        $postParams = @{
            "deviceManagementScriptAssignments" = @(
                $AssignmentParams
            )
        } | ConvertTo-Json -Depth 10
        
        Invoke-MgGraphRequest -Uri $uri -Method Post -Body $postParams
    }
    catch {
        Write-Log "Error assigning device management script: $_" -Level "ERROR"
        throw
    }
}

    # Deploy remediation script
    Write-Log "Checking for existing remediation script..."
    $existingScript = Get-DeviceManagementScripts | Where-Object {$_.displayName -eq $ScriptName}
    
    $encodedScriptContent = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($ScriptContent))
    
    $scriptParams = @{
        "@odata.type" = "#microsoft.graph.deviceManagementScript"
        displayName = $ScriptName
        description = "Automated remediation script"
        scriptContent = $encodedScriptContent
        runAsAccount = "system"
        enforceSignatureCheck = $false
        fileName = "$ScriptName.ps1"
    } | ConvertTo-Json

    if (!$existingScript) {
        Write-Log "Creating new remediation script: $ScriptName"
        $script = New-DeviceManagementScript -ScriptParams $scriptParams
        Write-Log "Created script with ID: $($script.Id)"
    } else {
        Write-Log "Updating existing remediation script: $ScriptName"
        $script = Update-DeviceManagementScript -ScriptId $existingScript.Id -ScriptParams $scriptParams
        Write-Log "Updated script with ID: $($script.Id)"
    }

    # Assign script to remediation group
    Write-Log "Assigning remediation script to device group..."
    $assignment = @{
        target = @{
            "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
            groupId = $group.Id
        }
    }

    New-DeviceManagementScriptAssignment -ScriptId $script.Id -AssignmentParams $assignment
    Write-Log "Remediation script assigned successfully to group"

    # Get existing members of remediation group
    Write-Log "Checking for devices to remove from remediation group..."
    $currentGroupMembers = Get-MgGroupMember -GroupId $group.Id -All

    # Get Azure AD Device IDs of devices currently having issues
    $devicesWithErrorsAadIds = @()
    foreach ($device in $devicesWithErrors) {
        $managedDevice = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $device.DeviceId
        if ($managedDevice.AzureAdDeviceId) {
            # Get the Azure AD device object to obtain its Object ID
            $aadDevice = Get-MgDevice -Filter "DeviceId eq '$($managedDevice.AzureAdDeviceId)'"
            if ($aadDevice) {
                $devicesWithErrorsAadIds += $aadDevice.Id  # Add the Object ID to our list
            }
        }
    }

    # Identify devices to remove (devices currently in group but no longer have errors)
    $devicesToRemove = $currentGroupMembers | Where-Object { $_.Id -notin $devicesWithErrorsAadIds }
    Write-Log "Found $($devicesToRemove.Count) devices to remove from remediation group (no longer have errors)"

    # Remove healthy devices from remediation group
    $devicesRemoved = 0
    foreach ($device in $devicesToRemove) {
        try {
            # Remove member using the same URL pattern as when adding members
            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $device.Id
            Write-Log "Removed device $($device.Id) from remediation group (no longer has errors)"
            $devicesRemoved++
        } catch {
            Write-Log "Error removing device $($device.Id) from group: $_" -Level "ERROR"
        }
    }
    Write-Log "Removed $devicesRemoved devices from remediation group"

    Write-Log "Intune Expedited Updates Auto-Resolve runbook completed successfully"
} catch {
    Write-Log "An error occurred during script execution: $_" -Level "ERROR"
    Write-Log "Error details: $($_.Exception)" -Level "ERROR"
    throw
}