# Define global variables
$TenantId = "<YourTenantID>"
$TenantDomain = "<DefaultDomain>.onmicrosoft.com" # Tenant domain name for Exchange Online (not GUID)
$SiteURL = "https://defaultdomain.sharepoint.com/sites/<SiteName>"  
$LibraryName = "/sites/<SiteName>/Shared%20Documents/<FolderName>"
$ExcelFileName = "EventData_With_Organizer.xlsx"
$DefaultOrganizerEmail = "hr@domain.com" # Change this to your default organizer email

# App registration certificate authentication parameters
$cert = Get-AutomationCertificate -Name '<NAMEOFCERTIFICATE>'
$ClientId = "<APPCLIENTID>" # App registration client ID
$Thumb = $cert.Thumbprint # Certificate thumbprint

# Exchange Online connection tracking
$Global:ExchangeConnected = $false
$Global:LastExchangeAuthTime = $null

# Set up configuration
$IsRunningInAutomation = $null -ne $env:AUTOMATION_ASSET_ACCOUNTID -or $null -ne $env:AUTOMATION_WORKER_ID
$LogPath = if ($IsRunningInAutomation) { $null } else { "$env:TEMP\CalendarEventCreation_$((Get-Date -Format 'yyyyMMdd_HHmmss')).log" }
$MaxRetryCount = 3
$RetryDelay = 2 # seconds

# Configure debug stream for Azure Automation logging
if ($IsRunningInAutomation) {
    $GLOBAL:DebugPreference = "Continue"
}

###########################################################################
# Function for logging - Azure Automation compatible
###########################################################################
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"

    if ($IsRunningInAutomation) {
        # CRITICAL: Avoid Write-Output as it contaminates function return values
        switch ($Level) {
            "INFO" {
                # Use Write-Information for normal logs - not visible in Azure Automation job output
                Write-Information $LogEntry -InformationAction Continue
            }
            "WARNING" {
                # Use Write-Warning for warnings - visible by default in Azure Automation
                Write-Warning $Message
            }
            "ERROR" {
                # Use Write-Error for errors - visible by default in Azure Automation
                Write-Error $Message
            }
        }
    }
    else {
        # For local execution, use Write-Host with colors
        switch ($Level) {
            "INFO"    { Write-Host $LogEntry -ForegroundColor White }
            "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
            "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
        }
          # Append to log file if not in Azure Automation
        if ($LogPath) {
            Add-Content -Path $LogPath -Value $LogEntry
        }
    }
}

###########################################################################
# Function to implement retry logic
###########################################################################
function Invoke-WithRetry {
    param (
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory=$false)]
        [string]$OperationName = "Operation",

        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = $MaxRetryCount,

        [Parameter(Mandatory=$false)]
        [int]$DelayInSeconds = $RetryDelay
    )
    
    $RetryCount = 0
    $Completed = $false
    $Result = $null
    
    while (-not $Completed -and $RetryCount -le $MaxRetries) {
        try {
            if ($RetryCount -gt 0) {
                Write-Log "Retry attempt $RetryCount of $MaxRetries for $OperationName" -Level "WARNING"
                # Exponential backoff
                Start-Sleep -Seconds ($DelayInSeconds * [Math]::Pow(2, ($RetryCount - 1)))
            }
            
            $Result = & $ScriptBlock
            $Completed = $true
            
            if ($RetryCount -gt 0) {
                Write-Log "$OperationName succeeded after $RetryCount retries" -Level "INFO"
            }
        }
        catch {
            $RetryCount++
            $ErrorMessage = $_.Exception.Message
            if ($RetryCount -le $MaxRetries) {
                Write-Log "$OperationName failed with error: $ErrorMessage. Retrying..." -Level "WARNING"
            }
            else {
                Write-Log "$OperationName failed after $MaxRetries retries with error: $ErrorMessage" -Level "ERROR"
                throw $_
            }
        }
    }
    
    return $Result
}

###########################################################################
# Initialize logging
###########################################################################
Write-Log "Script execution started"
if ($IsRunningInAutomation) {
    Write-Log "Running in Azure Automation environment"
    Write-Log "Azure Automation Account: $env:COMPUTERNAME"
    Write-Log "Runbook Job ID: $env:AUTOMATION_JOB_ID"
    Write-Log "Temp directory: $env:TEMP"
} else {
    Write-Log "Running in local environment"
    Write-Log "Log file created at: $LogPath"
    Write-Log "Computer name: $env:COMPUTERNAME"
}

# Ensure required modules are available
$RequiredModules = @("Microsoft.Graph", "ImportExcel", "ExchangeOnlineManagement")
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Log "Required module not found: $Module" -Level "ERROR"
        throw "Required module not found: $Module. Please install it using: Install-Module $Module -Force"
    }
}

###########################################################################
# Time-based re-authentication for Microsoft Graph only
###########################################################################
$Global:LastGraphAuthTime = $null

function Assert-GraphAuth {
    param (
        [int]$MaxTokenMinutes = 50
    )
    if (-not $Global:LastGraphAuthTime) {
        # If not set yet, set it
        $Global:LastGraphAuthTime = Get-Date
    }
    $MinutesSinceAuth = (Get-Date) - $Global:LastGraphAuthTime
    if ($MinutesSinceAuth.TotalMinutes -ge $MaxTokenMinutes) {
        Write-Log "Token approaching expiry, reconnecting to Microsoft Graph"

        Write-Log "Connecting to Microsoft Graph using certificate authentication"
        
        # Connect to Graph API with certificate authentication
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumb
        
        $Global:LastGraphAuthTime = Get-Date
        Write-Log "Successfully reconnected to Microsoft Graph using certificate authentication"
    }
}

###########################################################################
# Exchange Online authentication with time-based re-authentication
###########################################################################
function Assert-ExchangeAuth {
    param (
        [int]$MaxTokenMinutes = 50
    )
    
    if (-not $Global:ExchangeConnected -or (-not $Global:LastExchangeAuthTime)) {
        Write-Log "Connecting to Exchange Online for the first time"
        Connect-ExchangeOnlineWithAppAuth
        return
    }
    
    $MinutesSinceAuth = (Get-Date) - $Global:LastExchangeAuthTime
    if ($MinutesSinceAuth.TotalMinutes -ge $MaxTokenMinutes) {
        Write-Log "Exchange token approaching expiry, reconnecting"
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log "Warning during Exchange disconnect: $_" -Level "WARNING"
        }
        Connect-ExchangeOnlineWithAppAuth
    }
}

function Connect-ExchangeOnlineWithAppAuth {
    try {
        Write-Log "Connecting to Exchange Online using certificate authentication"
        
        if (-not [string]::IsNullOrEmpty($ClientId) -and -not [string]::IsNullOrEmpty($Thumb) -and -not [string]::IsNullOrEmpty($TenantDomain)) {
            # Use certificate authentication with tenant domain (not GUID)
            Connect-ExchangeOnline -CertificateThumbPrint $Thumb -AppId $ClientId -Organization $TenantDomain -ShowProgress:$false -ErrorAction Stop
        }
        else {
            throw "Missing required parameters for Exchange Online authentication. ClientId, Thumb, and TenantDomain are required."
        }
        
        $Global:ExchangeConnected = $true
        $Global:LastExchangeAuthTime = Get-Date
        Write-Log "Successfully connected to Exchange Online using certificate authentication"
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $_" -Level "ERROR"
        $Global:ExchangeConnected = $false
        throw "Failed to connect to Exchange Online: $_"
    }
}

###########################################################################
# Function to retrieve file from SharePoint now uses Microsoft Graph API
###########################################################################
Write-Log "SharePoint access will be handled through Microsoft Graph API"

# Function to retrieve file using Microsoft Graph API
function Get-FileFromSharePointGraph {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$RelativeFilePath,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    try {
        # Extract hostname and site path from the URL
        $uri = [System.Uri]$SiteUrl
        $hostname = $uri.Host
        
        # Handle site path extraction
        $sitePath = $uri.AbsolutePath
        if ($sitePath.StartsWith("/")) {
            $sitePath = $sitePath.Substring(1)
        }
        
        Write-Log "Getting site ID for $hostname/$sitePath"
        
        # Get the site ID
        $siteResponse = Invoke-WithRetry -OperationName "Get SharePoint Site ID" -ScriptBlock {
            $siteApiUrl = "https://graph.microsoft.com/v1.0/sites/$hostname`:/sites/$($sitePath.Split('/')[1])"
            Invoke-MgGraphRequest -Method GET -Uri $siteApiUrl
        }
        
        $siteId = $siteResponse.id
        Write-Log "Found SharePoint site ID: $siteId"
        
        # Get drives in the site
        $drivesResponse = Invoke-WithRetry -OperationName "Get Site Drives" -ScriptBlock {
            $drivesApiUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
            Invoke-MgGraphRequest -Method GET -Uri $drivesApiUrl
        }
        
        # Find the Documents library or first available library
        $drive = $drivesResponse.value | Where-Object { $_.name -eq "Documents" -or $_.name -eq "Shared Documents" } | Select-Object -First 1
        if (-not $drive) {
            $drive = $drivesResponse.value | Select-Object -First 1
        }
        
        $driveId = $drive.id
        Write-Log "Using drive: $($drive.name) (ID: $driveId)"
        
        # Clean up file path for API request
        $filePathForApi = $RelativeFilePath.Replace("%20", " ")
        $filePathForApi = $filePathForApi -replace '^/sites/[^/]+/', ''  # Remove site path prefix if present
        $filePathForApi = $filePathForApi -replace 'Shared Documents/', '' # Remove "Shared Documents" if present
        $filePathForApi = $filePathForApi -replace 'Documents/', '' # Remove "Documents" if present
          # Download the file using Graph API
        $filePath = Join-Path -Path $DestinationPath -ChildPath $FileName
        
        Write-Log "Downloading file from SharePoint path: $filePathForApi/$FileName"
        Write-Log "Will save to local path: $filePath"
        
        # Make sure destination directory exists
        $destinationDir = Split-Path -Path $filePath -Parent
        if (-not (Test-Path -Path $destinationDir)) {
            Write-Log "Creating destination directory: $destinationDir"
            New-Item -Path $destinationDir -ItemType Directory -Force | Out-Null
        }
        
        Invoke-WithRetry -OperationName "Download SharePoint File" -ScriptBlock {
            $fileApiUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$filePathForApi/$FileName`:/content"
            Write-Log "Requesting file from: $fileApiUrl"
            Invoke-MgGraphRequest -Method GET -Uri $fileApiUrl -OutputFilePath $filePath
        }
          # Verify file was downloaded successfully
        if (Test-Path -Path $filePath) {
            $fileInfo = Get-Item -Path $filePath
            Write-Log "File downloaded successfully to $filePath (Size: $($fileInfo.Length) bytes)"
            # Just return the path string, not Write-Output which would capture log messages too
            $filePath
        } else {
            Write-Log "Failed to download file - file not found at destination path: $filePath" -Level "ERROR"
            throw "Downloaded file not found at expected location: $filePath"
        }
    }
    catch {
        Write-Log "Failed to download file from SharePoint: $_" -Level "ERROR"
        throw "Failed to download file from SharePoint: $_"
    }
}

###########################################################################
# Function to upload file to SharePoint using Microsoft Graph API
###########################################################################
function Set-FileToSharePointGraph {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$RelativeFilePath,
        
        [Parameter(Mandatory=$true)]
        [string]$LocalFilePath,
        
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    try {
        # Verify local file exists
        if (-not (Test-Path -Path $LocalFilePath)) {
            throw "Local file not found: $LocalFilePath"
        }
        
        $fileInfo = Get-Item -Path $LocalFilePath
        Write-Log "Uploading file to SharePoint: $FileName (Size: $($fileInfo.Length) bytes)"
        
        # Extract hostname and site path from the URL
        $uri = [System.Uri]$SiteUrl
        $hostname = $uri.Host
        
        # Handle site path extraction
        $sitePath = $uri.AbsolutePath
        if ($sitePath.StartsWith("/")) {
            $sitePath = $sitePath.Substring(1)
        }
        
        Write-Log "Getting site ID for $hostname/$sitePath"
        
        # Get the site ID
        $siteResponse = Invoke-WithRetry -OperationName "Get SharePoint Site ID" -ScriptBlock {
            $siteApiUrl = "https://graph.microsoft.com/v1.0/sites/$hostname`:/sites/$($sitePath.Split('/')[1])"
            Invoke-MgGraphRequest -Method GET -Uri $siteApiUrl
        }
        
        $siteId = $siteResponse.id
        Write-Log "Found SharePoint site ID: $siteId"
        
        # Get drives in the site
        $drivesResponse = Invoke-WithRetry -OperationName "Get Site Drives" -ScriptBlock {
            $drivesApiUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
            Invoke-MgGraphRequest -Method GET -Uri $drivesApiUrl
        }
        
        # Find the Documents library or first available library
        $drive = $drivesResponse.value | Where-Object { $_.name -eq "Documents" -or $_.name -eq "Shared Documents" } | Select-Object -First 1
        if (-not $drive) {
            $drive = $drivesResponse.value | Select-Object -First 1
        }
        
        $driveId = $drive.id
        Write-Log "Using drive: $($drive.name) (ID: $driveId)"
        
        # Clean up file path for API request
        $filePathForApi = $RelativeFilePath.Replace("%20", " ")
        $filePathForApi = $filePathForApi -replace '^/sites/[^/]+/', ''  # Remove site path prefix if present
        $filePathForApi = $filePathForApi -replace 'Shared Documents/', '' # Remove "Shared Documents" if present
        $filePathForApi = $filePathForApi -replace 'Documents/', '' # Remove "Documents" if present
        
        Write-Log "Uploading file to SharePoint path: $filePathForApi/$FileName"
        
        # Upload the file using Graph API
        Invoke-WithRetry -OperationName "Upload SharePoint File" -ScriptBlock {
            $uploadApiUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$filePathForApi/$FileName`:/content"
            Write-Log "Uploading to URL: $uploadApiUrl"
            
            # Read file content as bytes
            $fileBytes = [System.IO.File]::ReadAllBytes($LocalFilePath)
            
            # Upload file using PUT request
            Invoke-MgGraphRequest -Method PUT -Uri $uploadApiUrl -Body $fileBytes -ContentType "application/octet-stream"
        }
        
        Write-Log "File uploaded successfully to SharePoint: $filePathForApi/$FileName"
        return $true
    }
    catch {
        Write-Log "Failed to upload file to SharePoint: $_" -Level "ERROR"
        throw "Failed to upload file to SharePoint: $_"
    }
}

###########################################################################
# Authenticate with Microsoft Graph using certificate
###########################################################################
try {
    Write-Log "Connecting to Microsoft Graph API using certificate authentication"
    
    # Connect to Graph API with certificate authentication
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumb
    
    # Set LastGraphAuthTime after initial connect
    $Global:LastGraphAuthTime = Get-Date
    Write-Log "Connected to Microsoft Graph API using certificate authentication"
    
    # Verify permissions to ensure we have necessary access for calendar operations
    Write-Log "Checking Microsoft Graph API permissions" -Level "INFO"
    $context = Get-MgContext
    Write-Log "Connected as application ID: $($context.ClientId)" -Level "INFO"
    Write-Log "Scopes: $($context.Scopes -join ', ')" -Level "INFO"
    
    # Check for required scopes
    $requiredScopes = @(
        "Calendars.ReadWrite",
        "Directory.Read.All",
        "User.Read.All",
        "Group.Read.All",
        "Mail.Read",
        "Files.Read"
    )
    
    $missingScopes = $requiredScopes | Where-Object { 
        $scope = $_
        -not ($context.Scopes | Where-Object { $_ -like "*$scope*" }) 
    }
    
    if ($missingScopes.Count -gt 0) {
        Write-Log "WARNING: The following recommended permissions might be missing: $($missingScopes -join ', ')" -Level "WARNING"
        Write-Log "This might affect the script's ability to create calendar events properly" -Level "WARNING"
    } else {
        Write-Log "Required permissions appear to be available" -Level "INFO"
    }
}
catch {
    Write-Log "Failed to connect to Microsoft Graph API using certificate authentication: $_" -Level "ERROR"
    throw "Failed to connect to Microsoft Graph API using certificate authentication: $_"
}

###########################################################################
# Function to remove old events from Excel file
###########################################################################
function Remove-OldEventsFromExcel {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ExcelFilePath,
        
        [Parameter(Mandatory=$true)]
        [array]$EventData
    )
    
    try {
        Write-Log "Checking for events older than today to remove from Excel file" -Level "INFO"
        $Today = (Get-Date).Date
        $EventsToKeep = @()
        $RemovedEventsCount = 0
        
        foreach ($Event in $EventData) {
            try {
                # Parse the start date from the event
                $EventStartDate = $null
                
                if ($Event.StartTime) {
                    # Handle different date formats
                    if ($Event.StartTime -match '^\d{1,2}/\d{1,2}/\d{4}$' -or $Event.StartTime -match '^\d{4}-\d{1,2}-\d{1,2}$') {
                        # Date-only format
                        $EventStartDate = [DateTime]::Parse($Event.StartTime).Date
                    } else {
                        # Full datetime format
                        $EventStartDate = [DateTime]::Parse($Event.StartTime).Date
                    }
                    
                    # Check if event is older than today
                    if ($EventStartDate -lt $Today) {
                        Write-Log "Removing old event: '$($Event.Subject)' (Start: $($Event.StartTime))" -Level "INFO"
                        $RemovedEventsCount++
                    } else {
                        # Keep this event
                        $EventsToKeep += $Event
                    }
                } else {
                    Write-Log "Event '$($Event.Subject)' has no StartTime, keeping it for manual review" -Level "WARNING"
                    $EventsToKeep += $Event
                }
            }
            catch {
                Write-Log "Error parsing date for event '$($Event.Subject)': $_. Keeping event for manual review" -Level "WARNING"
                $EventsToKeep += $Event
            }
        }
          # If we removed any events, update the Excel file
        if ($RemovedEventsCount -gt 0) {
            Write-Log "Removed $RemovedEventsCount old events. Updating Excel file with $($EventsToKeep.Count) remaining events" -Level "INFO"
            
            try {
                # Create a backup of the original file before modifying
                $BackupPath = $ExcelFilePath -replace '\.xlsx$', "_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
                Copy-Item -Path $ExcelFilePath -Destination $BackupPath -Force
                Write-Log "Created backup of original Excel file: $BackupPath" -Level "INFO"
                  # Export the updated data back to Excel
                $EventsToKeep | Export-Excel -Path $ExcelFilePath -WorksheetName "Sheet1" -ClearSheet -AutoSize
                Write-Log "Excel file updated successfully" -Level "INFO"
                
                # Upload the updated Excel file back to SharePoint
                try {
                    Write-Log "Uploading updated Excel file back to SharePoint" -Level "INFO"
                    Set-FileToSharePointGraph -SiteUrl $SiteURL -RelativeFilePath $LibraryName -LocalFilePath $ExcelFilePath -FileName $ExcelFileName
                    Write-Log "Excel file uploaded to SharePoint successfully" -Level "INFO"
                }
                catch {
                    Write-Log "Failed to upload updated Excel file to SharePoint: $_" -Level "WARNING"
                    Write-Log "The local Excel file has been updated, but SharePoint version may be outdated" -Level "WARNING"
                }
            }
            catch {
                Write-Log "Failed to update Excel file: $_. Continuing with filtered events in memory only" -Level "WARNING"
            }
        } else {
            Write-Log "No old events found to remove" -Level "INFO"
        }
        
        Write-Log "Event filtering completed. Original count: $($EventData.Count), Remaining count: $($EventsToKeep.Count)" -Level "INFO"
        return $EventsToKeep
    }
    catch {
        Write-Log "Error during old event removal: $_. Returning original event data" -Level "ERROR"
        return $EventData
    }
}

###########################################################################
# Download and read Excel
###########################################################################
# In Azure Automation, we need to ensure we have a writable temp directory
$TempDir = if ($IsRunningInAutomation) { 
    # Azure Automation uses a specific temp directory
    $tempPath = Join-Path -Path $env:TEMP -ChildPath ([System.Guid]::NewGuid().ToString())
    # Create the directory if it doesn't exist
    if (-not (Test-Path -Path $tempPath)) {
        New-Item -Path $tempPath -ItemType Directory -Force | Out-Null
    }
    $tempPath
} else {
    $env:TEMP
}
$LocalExcelPath = Join-Path -Path $TempDir -ChildPath $ExcelFileName

try {
    Write-Log "Downloading Excel file from SharePoint: $LibraryName/$ExcelFileName"
    Write-Log "Target local path will be: $LocalExcelPath"
    
    # Make sure Graph is authenticated before downloading
    Assert-GraphAuth
    
    # Use Graph API to download the file
    $downloadedPath = Invoke-WithRetry -OperationName "Excel File Download" -ScriptBlock {
        Get-FileFromSharePointGraph -SiteUrl $SiteURL -RelativeFilePath $LibraryName -DestinationPath $TempDir -FileName $ExcelFileName
    }
    
    Write-Log "SharePoint download returned path: $downloadedPath"
    
    # Verify that the downloaded path and expected path match
    if ($downloadedPath -ne $LocalExcelPath) {
        Write-Log "Downloaded path ($downloadedPath) differs from expected path ($LocalExcelPath). Copying file to expected location." -Level "WARNING"
        Copy-Item -Path $downloadedPath -Destination $LocalExcelPath -Force
    }
}
catch {
    Write-Log "Failed to download Excel file from SharePoint: $_" -Level "ERROR"
    throw "Failed to download Excel file from SharePoint: $_"
}

# Final verification that the file exists
if (-not (Test-Path -Path $LocalExcelPath)) {
    Write-Log "Excel file not found at: $LocalExcelPath" -Level "ERROR"
    throw "Excel file not found at: $LocalExcelPath"
} else {
    $fileInfo = Get-Item -Path $LocalExcelPath
    Write-Log "Excel file verified at $LocalExcelPath (Size: $($fileInfo.Length) bytes)"
}

try {
    Write-Log "Importing data from Excel file"
    $EventData = Import-Excel -Path $LocalExcelPath
    
    if ($EventData.Count -eq 0) {
        Write-Log "No events found in the Excel file" -Level "ERROR"
        throw "No events found in the Excel file"
    }
    
    # Validate required columns
    $RequiredColumns = @('Subject','StartTime','EndTime','AttendeeEmails')
    $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $EventData[0].PSObject.Properties.Name }
    if ($MissingColumns.Count -gt 0) {
        Write-Log "Excel file is missing required columns: $($MissingColumns -join ', ')" -Level "ERROR"
        throw "Excel file is missing required columns: $($MissingColumns -join ', ')"
    }
      Write-Log "Excel data imported successfully: $($EventData.Count) events found"
    
    # Remove old events from the data and update Excel file if needed
    $EventData = Remove-OldEventsFromExcel -ExcelFilePath $LocalExcelPath -EventData $EventData
    
    if ($EventData.Count -eq 0) {
        Write-Log "No events remaining after removing old events" -Level "WARNING"
        Write-Log "Script execution completed - no events to process"
        exit 0
    }
}
catch {
    Write-Log "Failed to import Excel data: $_" -Level "ERROR"
    throw "Failed to import Excel data: $_"
}

###########################################################################
# Function to check if a user already has an event
###########################################################################
function Test-ExistingCalendarEvent {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId,
        
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$false)]
        [string]$TimeZone = "UTC"
    )

    # We'll look for events within a small time range around StartTime.
    $BufferMinutes = 1
    $SearchStart = $StartTime.AddMinutes(-$BufferMinutes).ToString("o") # 'o' = ISO 8601
    $SearchEnd   = $StartTime.AddMinutes($BufferMinutes).ToString("o")
    
    Write-Log "Checking for existing calendar event for user $UserId with subject '$Subject' around time $($StartTime.ToString("yyyy-MM-dd HH:mm:ss"))" -Level "INFO"

    # Re-auth for Graph if needed
    Assert-GraphAuth

    # Retrieve existing events in that time window
    try {
        $ExistingEvents = Invoke-WithRetry -OperationName "Get Calendar View" -ScriptBlock {
            Get-MgUserCalendarView -UserId $UserId `
                -StartDateTime $SearchStart `
                -EndDateTime   $SearchEnd `
                -All -ErrorAction Stop
        }
        
        Write-Log "Found $($ExistingEvents.Count) existing events in the time window" -Level "INFO"
        
        if ($ExistingEvents) {
            foreach ($Evt in $ExistingEvents) {
                if ($Evt.Subject -eq $Subject) {
                    Write-Log "Duplicate event found: Subject='$Subject', ID=$($Evt.Id), iCalUId=$($Evt.iCalUId)" -Level "INFO"
                    return $true
                }
            }
        }
        return $false
    }
    catch {
        Write-Log "Error checking for existing events: $($_.Exception.Message)" -Level "WARNING"
        # If we can't check for existing events, assume none exists
        return $false
    }
}

###########################################################################
# Function to find events to delete for a user
###########################################################################
function Find-EventsToDelete {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId,
        
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$false)]
        [string]$TimeZone = "UTC",
        
        [Parameter(Mandatory=$false)]
        [int]$BufferHours = 24
    )

    # Look for events within a broader time range to catch events that might have slight time differences
    $SearchStart = $StartTime.AddHours(-$BufferHours).ToString("o") # 'o' = ISO 8601
    $SearchEnd   = $StartTime.AddHours($BufferHours).ToString("o")
    
    Write-Log "Searching for events to delete for user $UserId with subject '$Subject' around time $($StartTime.ToString("yyyy-MM-dd HH:mm:ss"))" -Level "INFO"

    # Re-auth for Graph if needed
    Assert-GraphAuth

    # Retrieve existing events in that time window
    try {
        $ExistingEvents = Invoke-WithRetry -OperationName "Get Calendar View for Deletion" -ScriptBlock {
            Get-MgUserCalendarView -UserId $UserId `
                -StartDateTime $SearchStart `
                -EndDateTime   $SearchEnd `
                -All -ErrorAction Stop
        }
        
        Write-Log "Found $($ExistingEvents.Count) total events in the search window" -Level "INFO"
        
        $EventsToDelete = @()
        if ($ExistingEvents) {
            foreach ($Evt in $ExistingEvents) {
                if ($Evt.Subject -eq $Subject) {
                    Write-Log "Found matching event for deletion: Subject='$Subject', ID=$($Evt.Id), Start=$($Evt.Start.DateTime)" -Level "INFO"
                    $EventsToDelete += $Evt
                }
            }
        }
        
        Write-Log "Found $($EventsToDelete.Count) events matching deletion criteria" -Level "INFO"
        return $EventsToDelete
    }
    catch {
        Write-Log "Error searching for events to delete: $($_.Exception.Message)" -Level "WARNING"
        return @()
    }
}

###########################################################################
# Function to delete a calendar event for a single user
###########################################################################
function Remove-CalendarEventForUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserEmail,
        
        [Parameter(Mandatory=$true)]
        [string]$EventId
    )
    
    try {
        # Re-auth for Graph if needed
        Assert-GraphAuth

        # Lookup the user object in Graph
        Write-Log "Looking up user with email: $UserEmail for event deletion" -Level "INFO"
        $User = $null
        try {
            $User = Invoke-WithRetry -OperationName "Get User by Email for Deletion" -ScriptBlock {
                Get-MgUser -Filter "Mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" -ErrorAction Stop
            }
            
            if ($User) {
                Write-Log "Found user for deletion: $($User.DisplayName) (ID: $($User.Id))" -Level "INFO"
            }
            else {
                Write-Log "User not found with email: $UserEmail for event deletion" -Level "WARNING"
                return $false
            }
        }
        catch {
            Write-Log "Error looking up user '$UserEmail' for deletion - $($_.Exception.Message)" -Level "ERROR"
            return $false
        }

        # Delete the event
        try {
            Write-Log "Deleting event ID: $EventId for user: $UserEmail" -Level "INFO"
            
            Invoke-WithRetry -OperationName "Delete Calendar Event" -ScriptBlock {
                Remove-MgUserEvent -UserId $User.Id -EventId $EventId -ErrorAction Stop
            }
            
            Write-Log "Successfully deleted event ID: $EventId for user: $UserEmail" -Level "INFO"
            return $true
        }
        catch {
            Write-Log "Error deleting event ID: $EventId for user: $UserEmail - $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error in Remove-CalendarEventForUser for $UserEmail`: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

###########################################################################
# Function to delete events for group members or individuals
###########################################################################
function Remove-CalendarEventForAllGroupMembers {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$true)]
        [string[]]$AttendeeEmails,
        
        [Parameter(Mandatory=$false)]
        [string]$TimeZone = "UTC"
    )
    
    $TotalDeletions = 0
    $FailedDeletions = 0
    
    try {
        foreach ($GroupEmail in $AttendeeEmails) {
            Write-Log "Processing deletion for group or email: $GroupEmail" -Level "INFO"
            
            # Check if it's a distribution group first
            $IsDistributionGroup = $false
            try {
                Write-Log "Checking if '$GroupEmail' is a distribution group for deletion" -Level "INFO"
                $IsDistributionGroup = Test-DistributionGroup -EmailAddress $GroupEmail
            }
            catch {
                Write-Log "Error checking if '$GroupEmail' is a distribution group for deletion: $_" -Level "WARNING"
                $IsDistributionGroup = $false
            }
            
            if ($IsDistributionGroup) {
                # Handle distribution group deletion
                Write-Log "Processing deletion for distribution group: $GroupEmail" -Level "INFO"
                
                # Get members of the distribution group
                $DistGroupMembers = Get-DistributionGroupMembers -GroupEmail $GroupEmail
                
                if ($DistGroupMembers -and $DistGroupMembers.Count -gt 0) {
                    Write-Log "Deleting calendar events for distribution group members..." -Level "INFO"
                    
                    foreach ($MemberEmail in $DistGroupMembers) {
                        try {
                            if ([string]::IsNullOrWhiteSpace($MemberEmail) -or ($MemberEmail -notmatch "@")) {
                                Write-Log "✗ Skipping invalid email for deletion: $MemberEmail" -Level "WARNING"
                                $FailedDeletions++
                                continue
                            }
                            
                            Write-Log "Processing deletion for distribution group member: $MemberEmail" -Level "INFO"
                            $DeletedCount = Remove-EventsForUser -UserEmail $MemberEmail -Subject $Subject -StartTime $StartTime -TimeZone $TimeZone
                            
                            if ($DeletedCount -gt 0) {
                                Write-Log "✓ Successfully deleted $DeletedCount event(s) for: $MemberEmail" -Level "INFO"
                                $TotalDeletions += $DeletedCount
                            } else {
                                Write-Log "ℹ No matching events found to delete for: $MemberEmail" -Level "INFO"
                            }
                        }
                        catch {
                            Write-Log "✗ Error processing deletion for distribution group member $MemberEmail`: $($_.Exception.Message)" -Level "ERROR"
                            $FailedDeletions++
                        }
                    }
                }
                else {
                    Write-Log "No valid members found in distribution group '$GroupEmail' for deletion" -Level "WARNING"
                }
            }
            else {
                # Re-auth for Graph if needed
                Assert-GraphAuth

                # Try to find M365 group (same logic as creation, but for deletion)
                $Group = $null
                
                # Method 1: Try direct mail filter
                try {
                    Write-Log "Trying to find M365 group by mail for deletion: $GroupEmail" -Level "INFO"
                    $Group = Invoke-WithRetry -OperationName "Get Group by Mail for Deletion" -ScriptBlock {
                        Get-MgGroup -Filter "mail eq '$GroupEmail'" -ErrorAction Stop
                    }
                    
                    if ($Group) {
                        Write-Log "Found M365 group for deletion using mail filter: $($Group.DisplayName)" -Level "INFO"
                    }
                }
                catch {
                    Write-Log "Get Group by Mail for deletion failed with error: $_. Trying alternate methods." -Level "WARNING"
                }
                
                # Method 2: Try with mailNickname if Method 1 failed
                if (-not $Group) {
                    try {
                        $mailNickname = $GroupEmail.Split('@')[0]
                        Write-Log "Trying to find M365 group by mailNickname for deletion: $mailNickname" -Level "INFO"
                        
                        $Group = Invoke-WithRetry -OperationName "Get Group by MailNickname for Deletion" -ScriptBlock {
                            Get-MgGroup -Filter "mailNickname eq '$mailNickname'" -ErrorAction Stop
                        }
                        
                        if ($Group) {
                            Write-Log "Found M365 group for deletion using mailNickname filter: $($Group.DisplayName)" -Level "INFO"
                        }
                    }
                    catch {
                        Write-Log "Get Group by MailNickname for deletion failed with error: $_" -Level "WARNING"
                    }
                }
                
                if ($Group) {
                    # We have an M365 group
                    Write-Log "Processing deletion for M365 group: $($Group.DisplayName) (ID: $($Group.Id))" -Level "INFO"
                    
                    try {
                        $GroupMembers = Invoke-WithRetry -OperationName "Get Group Members for Deletion" -ScriptBlock {
                            Get-MgGroupMember -GroupId $Group.Id -All -ErrorAction Stop
                        }
                        
                        if ($GroupMembers) {
                            Write-Log "Found $($GroupMembers.Count) members in M365 group '$($Group.DisplayName)' for deletion" -Level "INFO"
                            
                            foreach ($Member in $GroupMembers) {
                                try {
                                    $MemberDetails = Invoke-WithRetry -OperationName "Get Member Details for Deletion" -ScriptBlock {
                                        Get-MgUser -UserId $Member.Id -Property "Id,DisplayName,Mail,UserPrincipalName,AccountEnabled" -ErrorAction Stop
                                    }
                                    
                                    if ($MemberDetails) {
                                        $EmailForEvent = if ($MemberDetails.Mail) { $MemberDetails.Mail } else { $MemberDetails.UserPrincipalName }
                                        $IsValidEmail = ![string]::IsNullOrWhiteSpace($EmailForEvent) -and ($EmailForEvent -match "@")
                                        
                                        if ($IsValidEmail) {
                                            Write-Log "Processing deletion for M365 group member: $EmailForEvent" -Level "INFO"
                                            $DeletedCount = Remove-EventsForUser -UserEmail $EmailForEvent -Subject $Subject -StartTime $StartTime -TimeZone $TimeZone
                                            
                                            if ($DeletedCount -gt 0) {
                                                Write-Log "✓ Successfully deleted $DeletedCount event(s) for: $EmailForEvent" -Level "INFO"
                                                $TotalDeletions += $DeletedCount
                                            } else {
                                                Write-Log "ℹ No matching events found to delete for: $EmailForEvent" -Level "INFO"
                                            }
                                        } else {
                                            Write-Log "✗ Invalid email for M365 group member deletion: $($MemberDetails.DisplayName)" -Level "WARNING"
                                            $FailedDeletions++
                                        }
                                    }
                                }
                                catch {
                                    Write-Log "✗ Error processing deletion for M365 group member ID $($Member.Id): $($_.Exception.Message)" -Level "ERROR"
                                    $FailedDeletions++
                                }
                            }
                        }
                        else {
                            Write-Log "No members found in M365 group '$($Group.DisplayName)' for deletion" -Level "WARNING"
                        }
                    }
                    catch {
                        Write-Log "Error retrieving group members for deletion: $($_.Exception.Message)" -Level "ERROR"
                        $FailedDeletions++
                    }
                }
                else {
                    # Not a group, treat as individual email
                    Write-Log "No M365 group found, treating '$GroupEmail' as an individual user for deletion" -Level "INFO"
                    
                    $DeletedCount = Remove-EventsForUser -UserEmail $GroupEmail -Subject $Subject -StartTime $StartTime -TimeZone $TimeZone
                    
                    if ($DeletedCount -gt 0) {
                        Write-Log "✓ Successfully deleted $DeletedCount event(s) for individual user: $GroupEmail" -Level "INFO"
                        $TotalDeletions += $DeletedCount
                    } else {
                        Write-Log "ℹ No matching events found to delete for individual user: $GroupEmail" -Level "INFO"
                    }
                }
            }
        }
        
        Write-Log "Deletion summary: $TotalDeletions events deleted successfully, $FailedDeletions failures" -Level "INFO"
        return @{
            TotalDeletions = $TotalDeletions
            FailedDeletions = $FailedDeletions
        }
    }
    catch {
        Write-Log "Error in Remove-CalendarEventForAllGroupMembers: $_" -Level "ERROR"
        return @{
            TotalDeletions = $TotalDeletions
            FailedDeletions = $FailedDeletions + 1
        }
    }
}

###########################################################################
# Helper function to delete events for a specific user
###########################################################################
function Remove-EventsForUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserEmail,
        
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$false)]
        [string]$TimeZone = "UTC"
    )
    
    try {
        # Re-auth for Graph if needed
        Assert-GraphAuth

        # Lookup the user object in Graph
        $User = $null
        try {
            $User = Invoke-WithRetry -OperationName "Get User by Email for Event Deletion" -ScriptBlock {
                Get-MgUser -Filter "Mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" -ErrorAction Stop
            }
            
            if (-not $User) {
                Write-Log "User not found with email: $UserEmail for event deletion" -Level "WARNING"
                return 0
            }
        }
        catch {
            Write-Log "Error looking up user '$UserEmail' for event deletion - $($_.Exception.Message)" -Level "ERROR"
            return 0
        }

        # Find events to delete
        $EventsToDelete = Find-EventsToDelete -UserId $User.Id -Subject $Subject -StartTime $StartTime -TimeZone $TimeZone
        
        if ($EventsToDelete.Count -eq 0) {
            Write-Log "No events found to delete for user: $UserEmail with subject: '$Subject'" -Level "INFO"
            return 0
        }

        # Delete each matching event
        $DeletedCount = 0
        foreach ($Event in $EventsToDelete) {
            try {
                $Success = Remove-CalendarEventForUser -UserEmail $UserEmail -EventId $Event.Id
                if ($Success) {
                    $DeletedCount++
                    Write-Log "Deleted event: '$($Event.Subject)' (ID: $($Event.Id)) for user: $UserEmail" -Level "INFO"
                } else {
                    Write-Log "Failed to delete event: '$($Event.Subject)' (ID: $($Event.Id)) for user: $UserEmail" -Level "WARNING"
                }
            }
            catch {
                Write-Log "Error deleting event '$($Event.Subject)' for user $UserEmail`: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        
        return $DeletedCount
    }
    catch {
        Write-Log "Error in Remove-EventsForUser for $UserEmail`: $($_.Exception.Message)" -Level "ERROR"
        return 0
    }
}

###########################################################################
# Function to create a calendar event for a single user
###########################################################################
function New-CalendarEventForUser {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$true)]
        [datetime]$EndTime,
        
        [Parameter(Mandatory=$true)]
        [string]$UserEmail,
        
        [Parameter(Mandatory=$true)]
        [string]$OrganizerEmail,
        
        [Parameter(Mandatory=$false)]
        [string]$TimeZone = "UTC",
        
        [Parameter(Mandatory=$false)]
        [string]$Location = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Body = "",
        
        [Parameter(Mandatory=$false)]
        [bool]$IsAllDay = $true,
        
        [Parameter(Mandatory=$false)]
        [string]$ShowAs = "Free"
    )
    
    try {
        # Re-auth for Graph if needed
        Assert-GraphAuth

        # Lookup the user object in Graph
        Write-Log "Looking up user with email: $UserEmail" -Level "INFO"
        $User = $null
        try {
            $User = Invoke-WithRetry -OperationName "Get User by Email" -ScriptBlock {
                Get-MgUser -Filter "Mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" -ErrorAction Stop
            }
            
            if ($User) {
                Write-Log "Found user: $($User.DisplayName) (ID: $($User.Id))" -Level "INFO"
            }
            else {
                Write-Log "User not found with email: $UserEmail" -Level "WARNING"
                return $false
            }
        }        catch {
            Write-Log "Error looking up user '$UserEmail' - $($_.Exception.Message)" -Level "ERROR"
            return $false
        }

        # Check for existing event
        $HasDuplicate = Test-ExistingCalendarEvent -UserId $User.Id -Subject $Subject -StartTime $StartTime -TimeZone $TimeZone
        if ($HasDuplicate) {
            Write-Log "Skipping duplicate event for $UserEmail - Subject '$Subject', Start '$StartTime'" -Level "INFO"
            return $true
        }
          # Build event parameters
        $EventParams = @{
            Subject           = $Subject
            IsAllDay          = $IsAllDay
            ShowAs            = $ShowAs
            ResponseRequested = $false
            
            # Always add attendees array with the user to ensure visibility
            Attendees         = @(
                @{
                    EmailAddress = @{ 
                        Address = $UserEmail
                        Name = $User.DisplayName 
                    }
                    Type = "Required" 
                }
            )
            
            # Set reminder to default of 15 minutes
            ReminderMinutesBeforeStart = 15
            
            # Ensure these fields are explicitly set
            IsReminderOn      = $true
            Importance        = "Normal"
            Sensitivity       = "Normal"
            
            # Explicitly set the creator and organizer
            Organizer         = @{ 
                EmailAddress = @{ 
                    Address = $OrganizerEmail
                    Name = $OrganizerEmail.Split('@')[0] 
                } 
            }
            
            # Allow forwarding
            AllowNewTimeProposals = $true
        }
        
        # Handle Start and End datetime formatting based on IsAllDay
        if ($IsAllDay) {
            # For all-day events, use date-only format (yyyy-MM-dd) without time component
            Write-Log "Formatting as all-day event - using date-only format" -Level "INFO"
            $EventParams.Start = @{ 
                DateTime = $StartTime.ToString("yyyy-MM-dd")
                TimeZone = $TimeZone 
            }
            $EventParams.End = @{ 
                DateTime = $EndTime.ToString("yyyy-MM-dd")
                TimeZone = $TimeZone 
            }        } else {
            # For timed events, use full datetime format
            Write-Log "Formatting as timed event - using full datetime format" -Level "INFO"
            $EventParams.Start = @{ 
                DateTime = $StartTime.ToString("yyyy-MM-ddTHH:mm:ss")
                TimeZone = $TimeZone 
            }
            $EventParams.End = @{ 
                DateTime = $EndTime.ToString("yyyy-MM-ddTHH:mm:ss")
                TimeZone = $TimeZone 
            }
        }

        # Optionally add location
        if (-not [string]::IsNullOrWhiteSpace($Location)) {
            $EventParams.Location = @{ DisplayName = $Location }
        }

        # Optionally add body (only if provided and not empty)
        if (-not [string]::IsNullOrWhiteSpace($Body)) {
            $EventParams.Body = @{
                ContentType = "HTML"
                Content     = $Body
            }
        }
        
        # Re-auth for Graph if needed
        Assert-GraphAuth
        
        Write-Log "Attempting to create calendar event for: $UserEmail (Subject: $Subject)" -Level "INFO"
          # Log the parameters being used for troubleshooting
        $ParamsJson = ($EventParams | ConvertTo-Json -Depth 3 -Compress).Replace('"', '\"')
        Write-Log "Calendar event parameters: $ParamsJson" -Level "INFO"
        Write-Log "Using Graph API to create event for User ID: $($User.Id)" -Level "INFO"
        
        # Special logging for date-only events
        if ($StartTime.Hour -eq 9 -and $StartTime.Minute -eq 0) {
            Write-Log "This appears to be a date-only event (9:00 AM start time)" -Level "INFO"
            Write-Log "Note: Date-only events are interpreted as starting at 9:00 AM" -Level "INFO"        }
        
        try {
            Write-Log "Executing New-MgUserEvent for user: $UserEmail (UserId: $($User.Id))" -Level "INFO"
            Write-Log "Event parameters: Subject='$Subject', StartTime=$($StartTime.ToString('yyyy-MM-dd HH:mm:ss')), IsAllDay=$IsAllDay" -Level "INFO"
                
                $NewEvent = Invoke-WithRetry -OperationName "Create Calendar Event" -ScriptBlock {
                    $result = New-MgUserEvent -UserId $User.Id -BodyParameter $EventParams -ErrorAction Stop
                    
                    if (-not $result) {
                        throw "New-MgUserEvent returned null without throwing an error."
                    }
                    
                    $result
                }
                
                if ($NewEvent) {
                    Write-Log "Event created successfully for: $UserEmail (Event ID: $($NewEvent.Id))" -Level "INFO"
                    Write-Log "Event ICalUID: $($NewEvent.ICalUId)" -Level "INFO"
                    Write-Log "Event WebLink: $($NewEvent.WebLink)" -Level "INFO"
                    Write-Log "Event Start Time: $($NewEvent.Start.DateTime) ($($NewEvent.Start.TimeZone))" -Level "INFO"
                    Write-Log "Event End Time: $($NewEvent.End.DateTime) ($($NewEvent.End.TimeZone))" -Level "INFO"
                    Write-Log "Event IsAllDay: $($NewEvent.IsAllDay)" -Level "INFO"
                    Write-Log "Event Organizer: $($NewEvent.Organizer.EmailAddress.Address)" -Level "INFO"
                    
                    # Verify the event was created by retrieving it again
                    try {
                        Write-Log "Verifying event creation by retrieving it directly..." -Level "INFO"
                        $VerifyEvent = Invoke-WithRetry -OperationName "Verify Calendar Event" -ScriptBlock {
                            Get-MgUserEvent -UserId $User.Id -EventId $NewEvent.Id -ErrorAction Stop
                        }
                        
                        if ($VerifyEvent) {
                            Write-Log "Event verification successful. Event exists with ID: $($VerifyEvent.Id)" -Level "INFO"
                            
                            # Additional verification by checking calendar view
                            # This helps ensure the event is actually visible in the calendar
                            Write-Log "Checking if event is visible in calendar view..." -Level "INFO"
                            $VisibilityCheck = Test-CalendarEventsVisibility -UserEmail $UserEmail -StartTime $StartTime -Subject $Subject
                            
                            if ($VisibilityCheck) {
                                Write-Log "Event visibility confirmed in user's calendar" -Level "INFO"
                            }
                            else {
                                Write-Log "Event created but may not be visible in calendar view. This could be due to:" -Level "WARNING"
                                Write-Log "  1. Time delay in propagation" -Level "WARNING"
                                Write-Log "  2. Permission issues between the app and user's calendar" -Level "WARNING"
                                Write-Log "  3. Special calendar settings on the user's account" -Level "WARNING"
                            }
                        }
                        else {
                            Write-Log "Event verification failed. Could not retrieve created event with ID: $($NewEvent.Id)" -Level "WARNING"
                        }
                    }
                    catch {
                        Write-Log "Error verifying created event: $($_.Exception.Message)" -Level "WARNING"
                    }
                
                return $true
            }
            else {
                Write-Log "Failed to create event for: $UserEmail - no event object returned" -Level "ERROR"
                return $false
            }
        }
        catch {
            Write-Log "Error creating event for $UserEmail in Create Calendar Event section: $($_.Exception.Message)" -Level "ERROR"
            
            if ($_.ErrorDetails) {
                Write-Log "Error details: $($_.ErrorDetails)" -Level "ERROR"
            }
            
            if ($_.Exception.Response) {
                Write-Log "Response status code: $($_.Exception.Response.StatusCode)" -Level "ERROR"
            }
            
            return $false
        }
    }
    catch {
        Write-Log "Error creating calendar event for $($UserEmail): $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

###########################################################################
# Function to create events for group members or individuals
###########################################################################
function New-GraphCalendarEventForAllGroupMembers {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$true)]
        [datetime]$EndTime,
        
        [Parameter(Mandatory=$true)]
        [string[]]$AttendeeEmails,
        
        [Parameter(Mandatory=$false)]
        [string]$TimeZone = "UTC",
        
        [Parameter(Mandatory=$false)]
        [string]$Location = "",
        
        [Parameter(Mandatory=$false)]
        [string]$Body = "",
        
        [Parameter(Mandatory=$false)]
        [bool]$IsAllDay = $true,
        
        [Parameter(Mandatory=$false)]
        [string]$ShowAs = "Free",
          [Parameter(Mandatory=$false)]
        [string]$OrganizerEmail = $DefaultOrganizerEmail
    )
    
    try {
        foreach ($GroupEmail in $AttendeeEmails) {
            Write-Log "Processing group or email: $GroupEmail" -Level "INFO"
            
            # Check if it's a distribution group first
            $IsDistributionGroup = $false
            try {
                Write-Log "Checking if '$GroupEmail' is a distribution group" -Level "INFO"
                $IsDistributionGroup = Test-DistributionGroup -EmailAddress $GroupEmail
            }            catch {
                Write-Log "Error checking if '$GroupEmail' is a distribution group: $_" -Level "WARNING"
                $IsDistributionGroup = $false
            }
            
            if ($IsDistributionGroup) {
                # Handle distribution group
                Write-Log "Processing distribution group: $GroupEmail" -Level "INFO"
                
                # Get members of the distribution group (enhanced logging is in the function)
                $DistGroupMembers = Get-DistributionGroupMembers -GroupEmail $GroupEmail
                
                if ($DistGroupMembers -and $DistGroupMembers.Count -gt 0) {
                    Write-Log "Creating calendar events for distribution group members..." -Level "INFO"
                    $SuccessfulCreations = 0
                    $FailedCreations = 0
                    
                    foreach ($MemberEmail in $DistGroupMembers) {
                        try {
                            if ([string]::IsNullOrWhiteSpace($MemberEmail) -or ($MemberEmail -notmatch "@")) {
                                Write-Log "✗ Skipping invalid email: $MemberEmail" -Level "WARNING"
                                $FailedCreations++
                                continue
                            }
                            
                            Write-Log "Creating event for distribution group member: $MemberEmail" -Level "INFO"
                            
                            # Re-auth for Graph if needed
                            Assert-GraphAuth

                            $Success = New-CalendarEventForUser `
                                -Subject $Subject `
                                -StartTime $StartTime `
                                -EndTime $EndTime `
                                -UserEmail $MemberEmail `
                                -OrganizerEmail $OrganizerEmail `
                                -TimeZone $TimeZone `
                                -Location $Location `
                                -Body $Body `
                                -IsAllDay $IsAllDay `
                                -ShowAs $ShowAs
                            
                            if ($Success) {
                                Write-Log "✓ Successfully created event for: $MemberEmail" -Level "INFO"
                                $SuccessfulCreations++
                            } else {
                                Write-Log "✗ Failed to create event for: $MemberEmail" -Level "WARNING"
                                $FailedCreations++
                            }
                        }
                        catch {
                            Write-Log "✗ Error processing distribution group member $MemberEmail`: $($_.Exception.Message)" -Level "ERROR"
                            $FailedCreations++
                        }
                    }
                    
                    # Summary of distribution group processing
                    Write-Log "Distribution group '$GroupEmail' processing summary:" -Level "INFO"
                    Write-Log "  - Total members processed: $($DistGroupMembers.Count)" -Level "INFO"
                    Write-Log "  - Successful event creations: $SuccessfulCreations" -Level "INFO"
                    Write-Log "  - Failed event creations: $FailedCreations" -Level "INFO"
                    
                    if ($FailedCreations -gt 0) {
                        Write-Log "Some events failed to be created for distribution group members" -Level "WARNING"
                    }
                }
                else {
                    Write-Log "No valid members found in distribution group '$GroupEmail' - no events will be created" -Level "WARNING"
                }            }
            else {
                # Only Graph used here to find group, so re-auth for Graph
                Assert-GraphAuth

                # Try multiple methods to find the group
                $Group = $null
                
                # Method 1: Try direct mail filter
                try {
                    Write-Log "Trying to find M365 group by mail: $GroupEmail" -Level "INFO"
                    $Group = Invoke-WithRetry -OperationName "Get Group by Mail" -ScriptBlock {
                        Get-MgGroup -Filter "mail eq '$GroupEmail'" -ErrorAction Stop
                    }
                    
                    if ($Group) {
                        Write-Log "Found M365 group using mail filter: $($Group.DisplayName)" -Level "INFO"
                    }
                }
                catch {
                    Write-Log "Get Group by Mail failed with error: $_. Trying alternate methods." -Level "WARNING"
                }
                
                # Method 2: Try with mailNickname if Method 1 failed
                if (-not $Group) {
                    try {
                        $mailNickname = $GroupEmail.Split('@')[0]
                        Write-Log "Trying to find M365 group by mailNickname: $mailNickname" -Level "INFO"
                        
                        $Group = Invoke-WithRetry -OperationName "Get Group by MailNickname" -ScriptBlock {
                            Get-MgGroup -Filter "mailNickname eq '$mailNickname'" -ErrorAction Stop
                        }
                        
                        if ($Group) {
                            Write-Log "Found M365 group using mailNickname filter: $($Group.DisplayName)" -Level "INFO"
                        }
                    }
                    catch {
                        Write-Log "Get Group by MailNickname failed with error: $_" -Level "WARNING"
                    }
                }
                
                # Method 3: Try with case-insensitive mail search
                if (-not $Group) {
                    try {
                        Write-Log "Trying case-insensitive mail search for: $GroupEmail" -Level "INFO"
                        $lowerEmail = $GroupEmail.ToLower()
                        
                        $Group = Invoke-WithRetry -OperationName "Get Group by Case-Insensitive Mail" -ScriptBlock {
                            Get-MgGroup -Filter "tolower(mail) eq '$lowerEmail'" -ErrorAction Stop
                        }
                        
                        if ($Group) {
                            Write-Log "Found M365 group using case-insensitive mail filter: $($Group.DisplayName)" -Level "INFO"
                        }
                    }
                    catch {
                        Write-Log "Get Group by case-insensitive mail failed with error: $_" -Level "WARNING"
                    }
                }
                
                # Method 4: Try with display name
                if (-not $Group) {
                    try {
                        $displayName = $GroupEmail.Split('@')[0]
                        Write-Log "Trying to find M365 group by displayName: $displayName" -Level "INFO"
                          $Group = Invoke-WithRetry -OperationName "Get Group by DisplayName" -ScriptBlock {
                            Get-MgGroup -Filter "displayName eq '$displayName'" -ErrorAction Stop
                        }
                        
                        if ($Group) {
                            Write-Log "Found M365 group using displayName filter: $($Group.DisplayName)" -Level "INFO"
                        }
                    }
                    catch {
                        Write-Log "Get Group by DisplayName failed with error: $_" -Level "WARNING"
                    }
                }
                
                # Method 5: List all groups and search manually as last resort
                if (-not $Group) {
                    try {
                        Write-Log "Trying to find M365 group by listing all groups" -Level "INFO"
                        
                        $allGroups = Invoke-WithRetry -OperationName "Get All Groups" -ScriptBlock {
                            Get-MgGroup -All -Property "Id,DisplayName,Mail,MailNickname" -ErrorAction Stop
                        }
                        
                        # Try to match by various properties
                        $matchNickname = $GroupEmail.Split('@')[0]
                        
                        foreach ($g in $allGroups) {
                            if (($g.Mail -and $g.Mail -eq $GroupEmail) -or 
                                ($g.Mail -and $g.Mail.ToLower() -eq $GroupEmail.ToLower()) -or
                                ($g.MailNickname -and $g.MailNickname -eq $matchNickname) -or
                                ($g.DisplayName -and $g.DisplayName -eq $matchNickname)) {
                                
                                $Group = $g
                                Write-Log "Found M365 group through manual comparison: $($Group.DisplayName) (ID: $($Group.Id))" -Level "INFO"
                                break
                            }
                        }
                    }
                    catch {
                        Write-Log "Manual group search failed with error: $_" -Level "WARNING"
                    }
                }
                
                if ($Group) {
                    # We have an M365 group
                    Write-Log "Found M365 group: $($Group.DisplayName) (ID: $($Group.Id))" -Level "INFO"
                    
                    try {
                        $GroupMembers = Invoke-WithRetry -OperationName "Get Group Members" -ScriptBlock {
                            Get-MgGroupMember -GroupId $Group.Id -All -ErrorAction Stop
                        }
                        
                        if ($GroupMembers) {
                            Write-Log "Found $($GroupMembers.Count) members in M365 group '$($Group.DisplayName)'" -Level "INFO"
                            
                            # Enhanced logging: Get details for all members and log them
                            Write-Log "Retrieving detailed information for all M365 group members..." -Level "INFO"
                            $ValidMembers = @()
                            $InvalidMembers = @()
                            
                            foreach ($Member in $GroupMembers) {
                                try {
                                    $MemberDetails = Invoke-WithRetry -OperationName "Get Member Details for Logging" -ScriptBlock {
                                        Get-MgUser -UserId $Member.Id -Property "Id,DisplayName,Mail,UserPrincipalName,AccountEnabled" -ErrorAction Stop
                                    }
                                    
                                    if ($MemberDetails) {
                                        $EmailForEvent = if ($MemberDetails.Mail) { $MemberDetails.Mail } else { $MemberDetails.UserPrincipalName }
                                        $IsValidEmail = ![string]::IsNullOrWhiteSpace($EmailForEvent) -and ($EmailForEvent -match "@")
                                        
                                        $MemberInfo = @{
                                            Id = $MemberDetails.Id
                                            DisplayName = $MemberDetails.DisplayName
                                            Mail = $MemberDetails.Mail
                                            UserPrincipalName = $MemberDetails.UserPrincipalName
                                            EmailForEvent = $EmailForEvent
                                            AccountEnabled = $MemberDetails.AccountEnabled
                                            IsValidEmail = $IsValidEmail
                                        }
                                        
                                        if ($IsValidEmail) {
                                            $ValidMembers += $MemberInfo
                                            Write-Log "✓ Valid member: $($MemberDetails.DisplayName) ($EmailForEvent) [Enabled: $($MemberDetails.AccountEnabled)]" -Level "INFO"
                                        } else {
                                            $InvalidMembers += $MemberInfo
                                            Write-Log "✗ Invalid member: $($MemberDetails.DisplayName) (No valid email) [Enabled: $($MemberDetails.AccountEnabled)]" -Level "WARNING"
                                        }
                                    } else {
                                        Write-Log "✗ Unable to retrieve details for member ID: $($Member.Id)" -Level "WARNING"
                                        $InvalidMembers += @{ Id = $Member.Id; DisplayName = "Unknown"; IsValidEmail = $false }
                                    }
                                } catch {
                                    Write-Log "✗ Error retrieving details for member ID $($Member.Id): $($_.Exception.Message)" -Level "WARNING"
                                    $InvalidMembers += @{ Id = $Member.Id; DisplayName = "Error"; IsValidEmail = $false }                                }
                            }
                            
                            # Summary of member discovery
                            Write-Log "M365 Group '$($Group.DisplayName)' member discovery summary:" -Level "INFO"
                            Write-Log "  - Total members found: $($GroupMembers.Count)" -Level "INFO"
                            Write-Log "  - Valid members (will receive events): $($ValidMembers.Count)" -Level "INFO"
                            Write-Log "  - Invalid/problematic members: $($InvalidMembers.Count)" -Level "INFO"
                            
                            if ($ValidMembers.Count -eq 0) {
                                Write-Log "No valid members found in M365 group '$($Group.DisplayName)' - no events will be created" -Level "WARNING"
                            } else {
                                # Create events for all valid M365 group members
                                Write-Log "Creating calendar events for M365 group members..." -Level "INFO"
                                $SuccessfulCreations = 0
                                $FailedCreations = 0
                                
                                foreach ($ValidMember in $ValidMembers) {
                                    try {
                                        Write-Log "Creating event for M365 group member: $($ValidMember.EmailForEvent)" -Level "INFO"
                                        
                                        # Re-auth for Graph if needed
                                        Assert-GraphAuth

                                        $Success = New-CalendarEventForUser `
                                            -Subject $Subject `
                                            -StartTime $StartTime `
                                            -EndTime $EndTime `
                                            -UserEmail $ValidMember.EmailForEvent `
                                            -OrganizerEmail $OrganizerEmail `
                                            -TimeZone $TimeZone `
                                            -Location $Location `
                                            -Body $Body `
                                            -IsAllDay $IsAllDay `
                                            -ShowAs $ShowAs
                                        
                                        if ($Success) {
                                            Write-Log "✓ Successfully created event for: $($ValidMember.EmailForEvent)" -Level "INFO"
                                            $SuccessfulCreations++
                                        } else {
                                            Write-Log "✗ Failed to create event for: $($ValidMember.EmailForEvent)" -Level "WARNING"
                                            $FailedCreations++
                                        }
                                    }
                                    catch {
                                        Write-Log "✗ Error creating event for M365 group member $($ValidMember.EmailForEvent): $($_.Exception.Message)" -Level "ERROR"
                                        $FailedCreations++
                                    }
                                }
                                
                                # Summary of M365 group processing
                                Write-Log "M365 group '$($Group.DisplayName)' processing summary:" -Level "INFO"
                                Write-Log "  - Valid members processed: $($ValidMembers.Count)" -Level "INFO"
                                Write-Log "  - Successful event creations: $SuccessfulCreations" -Level "INFO"
                                Write-Log "  - Failed event creations: $FailedCreations" -Level "INFO"
                                
                                if ($FailedCreations -gt 0) {
                                    Write-Log "Some events failed to be created for M365 group members" -Level "WARNING"
                                }
                            }
                        }
                        else {
                            Write-Log "No members found in M365 group '$($Group.DisplayName)'" -Level "WARNING"
                        }
                    }
                    catch {
                        Write-Log "Error retrieving group members: $($_.Exception.Message)" -Level "ERROR"
                        continue                    }
                }
                else {
                    # Not a group, treat as individual email
                    Write-Log "No M365 group found, treating '$GroupEmail' as an individual user" -Level "INFO"
                    Write-Log "Attempting to create calendar event for individual user: $GroupEmail" -Level "INFO"
                    
                    # Re-auth for Graph if needed
                    Assert-GraphAuth

                    $Success = New-CalendarEventForUser `
                        -Subject $Subject `
                        -StartTime $StartTime `
                        -EndTime $EndTime `
                        -UserEmail $GroupEmail `
                        -OrganizerEmail $OrganizerEmail `
                        -TimeZone $TimeZone `
                        -Location $Location `
                        -Body $Body `
                        -IsAllDay $IsAllDay `
                        -ShowAs $ShowAs
                    
                    if ($Success) {
                        Write-Log "✓ Successfully created event for individual user: $GroupEmail" -Level "INFO"
                    } else {
                        Write-Log "✗ Failed to create event for individual user: $GroupEmail" -Level "WARNING"
                    }
                }
            }
        }
        return $true
    }
    catch {
        Write-Log "Error creating calendar events: $_" -Level "ERROR"
        return $false
    }
}

###########################################################################
# Function to check if an email address is a distribution group
###########################################################################
function Test-DistributionGroup {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress
    )
    
    try {
        # Basic validation to avoid unnecessary API calls
        if ([string]::IsNullOrWhiteSpace($EmailAddress) -or ($EmailAddress -notmatch "@")) {
            Write-Log "Invalid email address format: $EmailAddress" -Level "WARNING"
            return $false
        }
        
        # First check if it's an M365 Group to avoid "GroupMailbox not supported" error
        Write-Log "Checking if '$EmailAddress' is an M365 Group first..." -Level "INFO"
        $IsM365Group = Test-M365Group -EmailAddress $EmailAddress
        
        if ($IsM365Group) {
            Write-Log "Email '$EmailAddress' is an M365 Group, not a distribution group" -Level "INFO"
            return $false
        }
        
        Write-Log "Not an M365 Group, checking if '$EmailAddress' is a distribution group..." -Level "INFO"
        
        # Ensure Exchange Online connection
        Assert-ExchangeAuth
        
        $DistGroup = Invoke-WithRetry -OperationName "Get Distribution Group" -ScriptBlock {
            Get-DistributionGroup -Identity $EmailAddress -ErrorAction Stop
        }
        
        if ($DistGroup) {
            Write-Log "✓ Found distribution group: $($DistGroup.DisplayName) (Email: $($DistGroup.PrimarySMTPAddress))" -Level "INFO"
            return $true
        }
        
        return $false
    }
    catch {
        # If the group is not found or other error occurs, it's not a distribution group
        if ($_.Exception.Message -like "*not supported on GroupMailbox*") {
            Write-Log "✗ Email '$EmailAddress' is likely an M365 Group (GroupMailbox error)" -Level "WARNING"
        } else {
            Write-Log "✗ Email '$EmailAddress' is not a distribution group: $($_.Exception.Message)" -Level "INFO"
        }
        return $false
    }
}

###########################################################################
# Function to get distribution group members
###########################################################################
function Get-DistributionGroupMembers {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GroupEmail    )
    
    try {
        # Ensure Exchange Online connection
        Assert-ExchangeAuth
        
        Write-Log "Getting members for distribution group: $GroupEmail" -Level "INFO"
        
        if ([string]::IsNullOrWhiteSpace($GroupEmail) -or ($GroupEmail -notmatch "@")) {
            Write-Log "Invalid distribution group email address: $GroupEmail" -Level "WARNING"
            return @()
        }
        
        $GroupMembers = Invoke-WithRetry -OperationName "Get Distribution Group Members" -ScriptBlock {
            Get-DistributionGroupMember -Identity $GroupEmail -ErrorAction Stop
        }
        
        if ($GroupMembers) {
            Write-Log "Found $($GroupMembers.Count) members in distribution group '$GroupEmail'" -Level "INFO"
            
            # Enhanced logging: Analyze and log all members with detailed information
            Write-Log "Analyzing distribution group members..." -Level "INFO"
            $ValidEmailMembers = @()
            $NestedGroups = @()
            $InvalidMembers = @()
            $TotalNestedMembers = 0
            
            foreach ($Member in $GroupMembers) {
                try {
                    $MemberType = $Member.RecipientType
                    $MemberDisplayName = if ($Member.DisplayName) { $Member.DisplayName } else { "Unknown" }
                    $MemberEmail = if ($Member.PrimarySMTPAddress) { $Member.PrimarySMTPAddress } else { "No Email" }
                    
                    # Check if member is a user (has PrimarySMTPAddress)
                    if ($Member.PrimarySMTPAddress -and ($Member.PrimarySMTPAddress -match "@")) {
                        $ValidEmailMembers += $Member.PrimarySMTPAddress
                        Write-Log "✓ Valid user member: $MemberDisplayName ($MemberEmail) [Type: $MemberType]" -Level "INFO"
                    }
                    # Handle nested groups - if member is another distribution group
                    elseif ($Member.RecipientType -eq "MailUniversalDistributionGroup" -or 
                            $Member.RecipientType -eq "MailUniversalSecurityGroup") {
                        Write-Log "⮑ Nested group member: $MemberDisplayName ($MemberEmail) [Type: $MemberType]" -Level "INFO"
                        Write-Log "  Getting nested members from: $($Member.PrimarySMTPAddress)..." -Level "INFO"
                        
                        $NestedMembers = Get-DistributionGroupMembers -GroupEmail $Member.PrimarySMTPAddress
                        if ($NestedMembers -and $NestedMembers.Count -gt 0) {
                            $ValidEmailMembers += $NestedMembers
                            $TotalNestedMembers += $NestedMembers.Count
                            Write-Log "  ✓ Retrieved $($NestedMembers.Count) members from nested group '$MemberDisplayName'" -Level "INFO"
                        } else {
                            Write-Log "  ✗ No members found in nested group '$MemberDisplayName'" -Level "WARNING"
                        }
                        $NestedGroups += @{
                            DisplayName = $MemberDisplayName
                            Email = $MemberEmail
                            Type = $MemberType
                            MemberCount = if ($NestedMembers) { $NestedMembers.Count } else { 0 }
                        }
                    }
                    else {
                        # Member with no valid email or unrecognized type
                        $InvalidMembers += @{
                            DisplayName = $MemberDisplayName
                            Email = $MemberEmail
                            Type = $MemberType
                        }
                        Write-Log "✗ Invalid/problematic member: $MemberDisplayName ($MemberEmail) [Type: $MemberType]" -Level "WARNING"
                    }
                }
                catch {
                    Write-Log "✗ Error processing distribution group member $($Member.DisplayName): $_" -Level "WARNING"
                    $InvalidMembers += @{
                        DisplayName = if ($Member.DisplayName) { $Member.DisplayName } else { "Error" }
                        Email = "Error"
                        Type = "Error"
                    }
                }
            }
            
            # Remove duplicates and provide comprehensive summary
            $UniqueEmailMembers = $ValidEmailMembers | Sort-Object -Unique
            
            Write-Log "Distribution Group '$GroupEmail' member analysis summary:" -Level "INFO"
            Write-Log "  - Total members found: $($GroupMembers.Count)" -Level "INFO"
            Write-Log "  - Direct user members: $($ValidEmailMembers.Count - $TotalNestedMembers)" -Level "INFO"
            Write-Log "  - Nested groups found: $($NestedGroups.Count)" -Level "INFO"
            Write-Log "  - Members from nested groups: $TotalNestedMembers" -Level "INFO"
            Write-Log "  - Total unique email recipients: $($UniqueEmailMembers.Count)" -Level "INFO"
            Write-Log "  - Invalid/problematic members: $($InvalidMembers.Count)" -Level "INFO"
            
            if ($NestedGroups.Count -gt 0) {
                Write-Log "Nested groups breakdown:" -Level "INFO"
                foreach ($NestedGroup in $NestedGroups) {
                    Write-Log "  - $($NestedGroup.DisplayName): $($NestedGroup.MemberCount) members" -Level "INFO"
                }
            }
            
            if ($UniqueEmailMembers.Count -eq 0) {
                Write-Log "No valid email recipients found in distribution group '$GroupEmail' - no events will be created" -Level "WARNING"
            }
            
            return $UniqueEmailMembers
        }
        else {
            Write-Log "No members found in distribution group '$GroupEmail'" -Level "WARNING"
            return @()
        }
    }
    catch {
        Write-Log "Failed to get distribution group members for '$GroupEmail': $_" -Level "ERROR"
        return @()
    }
}

###########################################################################
# Function to verify Graph API permissions
###########################################################################
function Test-GraphPermissions {
    try {
        Write-Log "Verifying Microsoft Graph API permissions..." -Level "INFO"
        
        # Try to get current permissions
        $Context = Get-MgContext
        if (-not $Context) {
            Write-Log "No active Graph connection found" -Level "ERROR"
            return $false
        }
        
        Write-Log "Connected with app: $($Context.AppName), scopes: $($Context.Scopes -join ', ')" -Level "INFO"
        
        # Check for required permissions
        $RequiredPermissions = @(
            'Calendars.ReadWrite', 
            'User.Read.All', 
            'Group.Read.All'
        )
        
        $MissingPermissions = @()
        foreach ($Permission in $RequiredPermissions) {
            $Found = $false
            foreach ($Scope in $Context.Scopes) {
                # Use wildcard matching to handle .All vs specific permissions
                if ($Scope -like "*$Permission*") {
                    $Found = $true
                    break
                }
            }
            
            if (-not $Found) {
                $MissingPermissions += $Permission
            }
        }
        
        if ($MissingPermissions.Count -gt 0) {
            Write-Log "Missing required permissions: $($MissingPermissions -join ', ')" -Level "WARNING"
            
            # Try to test calendar access directly
            $TestPermission = $false
            try {
                # Try to get current user's calendar to verify permissions
                $CurrentUser = Invoke-WithRetry -OperationName "Get Current User" -ScriptBlock {
                    Get-MgUser -UserId $Context.Account -Property "Id,DisplayName" -ErrorAction Stop
                }
                
                if ($CurrentUser) {
                    Write-Log "Authenticated as: $($CurrentUser.DisplayName)" -Level "INFO"
                    
                    # Test if we can access calendars
                    $Calendar = Invoke-WithRetry -OperationName "Get Calendar" -ScriptBlock {
                        Get-MgUserCalendar -UserId $CurrentUser.Id -Top 1 -ErrorAction Stop
                    }
                    
                    if ($Calendar) {
                        Write-Log "Successfully accessed calendar" -Level "INFO"
                        $TestPermission = $true
                      }
                }
            }
            catch {
                Write-Log "Permission test failed: $($_.Exception.Message)" -Level "ERROR"
            }
            
            if ($TestPermission) {
                Write-Log "Despite missing explicit permissions, calendar access appears to work" -Level "INFO"
                return $true
            }
            else {
                return $false
            }
        }
        
        return $true
    }
    catch {
        Write-Log "Error testing Graph permissions: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

###########################################################################
# Function to connect to Microsoft Graph
###########################################################################
function Connect-ToMicrosoftGraph {
    param (
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    
    try {
        # Check if we need to connect or reconnect
        $Context = Get-MgContext -ErrorAction SilentlyContinue
        
        if ((-not $Context) -or $Force) {
            Write-Log "Connecting to Microsoft Graph with certificate authentication..." -Level "INFO"

            # Disconnect any existing connections
            if ($Context) {
                Write-Log "Disconnecting existing Graph connection" -Level "INFO"
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
            
            # Connect with certificate
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumb -ErrorAction Stop
            
            # Get and display the connection context info
            $Context = Get-MgContext
            Write-Log "Connected to Microsoft Graph as application: $($Context.AppName)" -Level "INFO"
            
            # Test the permissions
            $PermissionsOk = Test-GraphPermissions
            
            if (-not $PermissionsOk) {
                Write-Log "Warning: Graph API permissions may not be sufficient for all operations" -Level "WARNING"
            }
            
            return $true
        }
        else {
            Write-Log "Already connected to Microsoft Graph" -Level "INFO"
            return $true
        }
    }
    catch {
        Write-Log "Error connecting to Microsoft Graph: $_" -Level "ERROR"
        return $false
    }
}

###########################################################################
# Function to verify calendar events are visible 
###########################################################################
function Test-CalendarEventsVisibility {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserEmail,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$false)]
        [string]$Subject = $null,
        
        [Parameter(Mandatory=$false)]
        [int]$WindowHours = 24
    )
    
    try {
        # Re-auth for Graph if needed
        Assert-GraphAuth
        
        Write-Log "Testing calendar visibility for user: $UserEmail around time: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
        Write-Log "Looking for events with subject: '$Subject'" -Level "INFO"
        
        # Lookup the user
        $User = Invoke-WithRetry -OperationName "Get User by Email" -ScriptBlock {
            Get-MgUser -Filter "Mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" -ErrorAction Stop
        }
        
        if (-not $User) {
            Write-Log "User not found with email: $UserEmail" -Level "WARNING"
            return $false
        }
        
        Write-Log "Found user: $($User.DisplayName) (ID: $($User.Id))" -Level "INFO"
        
        # Check broader time window to ensure we capture the event
        # For all-day events or events that might have different time interpretations, 
        # we need a wider window
        $SearchStart = $StartTime.AddHours(-$WindowHours).ToString("o")
        $SearchEnd = $StartTime.AddHours($WindowHours).ToString("o")
        
        Write-Log "Searching for events from $SearchStart to $SearchEnd" -Level "INFO"
        
        $Events = Invoke-WithRetry -OperationName "Get User Calendar View" -ScriptBlock {
            Get-MgUserCalendarView -UserId $User.Id `
                -StartDateTime $SearchStart `
                -EndDateTime $SearchEnd `
                -All -ErrorAction Stop
        }
          if ($null -eq $Events) {
            Write-Log "No events returned from calendar view (null response)" -Level "WARNING"
            return $false
        }
        
        Write-Log "Found $($Events.Count) total events in the time window" -Level "INFO"
        
        # Look for all-day events separately if we're dealing with possible all-day events
        if ($StartTime.Hour -eq 9 -and $StartTime.Minute -eq 0) {
            Write-Log "Checking specifically for all-day events since start time is 9:00 AM" -Level "INFO"
            $AllDayEvents = $Events | Where-Object { $_.IsAllDay -eq $true }
            Write-Log "Found $($AllDayEvents.Count) all-day events" -Level "INFO"
            
            if ($AllDayEvents.Count -gt 0) {
                foreach ($evt in $AllDayEvents) {
                    Write-Log "All-day event: Subject='$($evt.Subject)', Start=$($evt.Start.DateTime)" -Level "INFO"
                }
            }
        }
        
        if ($Events.Count -eq 0) {
            Write-Log "No events found for user in the specified time window" -Level "WARNING"
            return $false
        }
        
        if ($Subject) {
            $MatchingEvents = $Events | Where-Object { $_.Subject -eq $Subject }
            
            if ($MatchingEvents.Count -gt 0) {
                Write-Log "Found $($MatchingEvents.Count) events with subject '$Subject'" -Level "INFO"
                  foreach ($FoundEvent in $MatchingEvents) {
                    $EventStartTime = [DateTime]::Parse($FoundEvent.Start.DateTime)
                    $FormattedStartTime = $EventStartTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $TimeZone = $FoundEvent.Start.TimeZone
                      Write-Log "Event ID: $($FoundEvent.Id), Subject: '$($FoundEvent.Subject)', Start: $FormattedStartTime ($TimeZone)" -Level "INFO"
                    Write-Log "Event Organizer: $($FoundEvent.Organizer.EmailAddress.Address)" -Level "INFO"
                    Write-Log "Event Status: $($FoundEvent.ShowAs), IsCancelled: $($FoundEvent.IsCancelled)" -Level "INFO"
                }
                
                return $true
            }
            else {
                Write-Log "No events found with subject '$Subject'" -Level "WARNING"
            }
        }
        else {
            Write-Log "Events found in the time window, but no specific subject was provided for matching" -Level "INFO"
              # Log a few events for diagnostic purposes
            $MaxToLog = [Math]::Min($Events.Count, 3)
            for ($i = 0; $i -lt $MaxToLog; $i++) {
                $SampleEvent = $Events[$i]
                $EventStartTime = [DateTime]::Parse($SampleEvent.Start.DateTime)
                $FormattedStartTime = $EventStartTime.ToString("yyyy-MM-dd HH:mm:ss")
                
                Write-Log "Sample Event $($i+1): ID=$($SampleEvent.Id), Subject='$($SampleEvent.Subject)', Start=$FormattedStartTime ($($SampleEvent.Start.TimeZone))" -Level "INFO"
            }
            
            return $true
        }
        
        return $false
    }
    catch {
        Write-Log "Error testing calendar visibility: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}



###########################################################################
# Main loop: process events from Excel
###########################################################################
$SuccessCount = 0
$FailureCount = 0
$DeletionCount = 0
$DeletionFailureCount = 0

Write-Log "Starting to process $($EventData.Count) events from Excel"

foreach ($CalEvent in $EventData) {
    try {
        Write-Log "Processing event: $($CalEvent.Subject)"
        
        # Check if this event is marked for deletion
        $ShouldDelete = $false
        if ($CalEvent.PSObject.Properties.Name -contains "Delete") {
            $DeleteValue = $CalEvent.Delete
            if (-not [string]::IsNullOrWhiteSpace($DeleteValue)) {
                # Check for various ways to indicate "Yes" for deletion
                $DeleteIndicators = @("yes", "y", "true", "1", "delete")
                $ShouldDelete = $DeleteIndicators -contains $DeleteValue.ToString().ToLower().Trim()
            }
        }
        
        if ($ShouldDelete) {
            Write-Log "Event '$($CalEvent.Subject)' is marked for deletion" -Level "INFO"
            
            # Parse the dates for deletion (similar to creation logic)
            try {
                Write-Log "Processing date/time for deletion of event: $($CalEvent.Subject)" -Level "INFO"
                Write-Log "Raw StartTime value for deletion: '$($CalEvent.StartTime)'" -Level "INFO"
                
                # Handle dates without time components (same logic as creation)
                if ($CalEvent.StartTime -match '^\d{1,2}/\d{1,2}/\d{4}$' -or $CalEvent.StartTime -match '^\d{4}-\d{1,2}-\d{1,2}$') {
                    Write-Log "StartTime appears to be date-only format for deletion, setting to midnight" -Level "INFO"
                    $StartTime = [DateTime]::Parse("$($CalEvent.StartTime) 12:00 AM")
                } else {
                    # Parse the full datetime
                    $StartTime = Get-Date -Date $CalEvent.StartTime -ErrorAction Stop
                }
                
                Write-Log "Parsed StartTime for deletion: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
            } 
            catch {
                Write-Log "Invalid date format in event for deletion '$($CalEvent.Subject)': $_" -Level "ERROR"
                Write-Log "Failed date value for deletion - StartTime: '$($CalEvent.StartTime)'" -Level "ERROR"
                $DeletionFailureCount++
                continue
            }
            
            # Get attendees for deletion
            if ([string]::IsNullOrWhiteSpace($CalEvent.AttendeeEmails)) {
                Write-Log "No attendee emails specified for deletion of event: $($CalEvent.Subject)" -Level "WARNING"
                $DeletionFailureCount++
                continue
            }
            
            $AttendeeEmails = $CalEvent.AttendeeEmails -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            if ($AttendeeEmails.Count -eq 0) {
                Write-Log "No valid attendee emails found for deletion of event: $($CalEvent.Subject)" -Level "WARNING"
                $DeletionFailureCount++
                continue
            }
            
            # Delete the event from attendees' calendars
            Write-Log "Attempting to delete event '$($CalEvent.Subject)' from attendees' calendars" -Level "INFO"
            Write-Log "Attendees for deletion: $($AttendeeEmails -join ', ')" -Level "INFO"
            
            try {
                $DeletionResult = Remove-CalendarEventForAllGroupMembers `
                    -Subject $CalEvent.Subject `
                    -StartTime $StartTime `
                    -AttendeeEmails $AttendeeEmails `
                    -TimeZone "UTC"
                
                if ($DeletionResult.TotalDeletions -gt 0) {
                    Write-Log "Successfully deleted $($DeletionResult.TotalDeletions) instance(s) of event '$($CalEvent.Subject)'" -Level "INFO"
                    $DeletionCount += $DeletionResult.TotalDeletions
                } else {
                    Write-Log "No instances of event '$($CalEvent.Subject)' were found to delete" -Level "INFO"
                }
                
                if ($DeletionResult.FailedDeletions -gt 0) {
                    Write-Log "Failed to delete event from $($DeletionResult.FailedDeletions) attendee(s)" -Level "WARNING"
                    $DeletionFailureCount += $DeletionResult.FailedDeletions
                }
            }
            catch {
                Write-Log "Error during deletion of event '$($CalEvent.Subject)': $_" -Level "ERROR"
                $DeletionFailureCount++
            }
            
            # Skip to next event (don't create if we're deleting)
            continue
        }
        
        # If not marked for deletion, proceed with normal event creation
        Write-Log "Event '$($CalEvent.Subject)' is not marked for deletion, proceeding with creation/update" -Level "INFO"

        # First determine if this is an all-day event
        $IsAllDay = $true  # Default to true (all day event)
        if ($CalEvent.PSObject.Properties.Name -contains "IsAllDay") {
            # Handle empty cells, null values, or whitespace - default to $true
            if ([string]::IsNullOrWhiteSpace($CalEvent.IsAllDay)) {
                $IsAllDay = $true
            } elseif ($CalEvent.IsAllDay -is [string]) {
                # For string values, only set to false for explicit false indicators
                $IsAllDay = -not ($CalEvent.IsAllDay -eq "false" -or $CalEvent.IsAllDay -eq "0" -or $CalEvent.IsAllDay -eq "no")
            } else {
                # For non-string values, convert to boolean but handle null/empty gracefully
                try {
                    $IsAllDay = [bool]$CalEvent.IsAllDay
                } catch {
                    # If conversion fails, default to true
                    $IsAllDay = $true
                }
            }
        }
        
        # Validate date/time
        try {
            Write-Log "Processing date/time for event: $($CalEvent.Subject)" -Level "INFO"
            Write-Log "Raw StartTime value: '$($CalEvent.StartTime)'" -Level "INFO"
            Write-Log "Raw EndTime value: '$($CalEvent.EndTime)'" -Level "INFO"
            Write-Log "IsAllDay: $IsAllDay" -Level "INFO"            # Handle dates without time components (like "All Hands" events)
            if ($CalEvent.StartTime -match '^\d{1,2}/\d{1,2}/\d{4}$' -or $CalEvent.StartTime -match '^\d{4}-\d{1,2}-\d{1,2}$') {
                if ($IsAllDay) {
                    Write-Log "StartTime appears to be date-only format for all-day event, setting to midnight" -Level "INFO"
                    $StartTime = [DateTime]::Parse("$($CalEvent.StartTime) 12:00 AM")
                } else {
                    Write-Log "StartTime appears to be date-only format for timed event, adding 9:00 AM time component" -Level "INFO"
                    $StartTime = [DateTime]::Parse("$($CalEvent.StartTime) 9:00 AM")
                }            } else {
                # Parse the full datetime, but for all-day events we'll normalize to midnight below
                $StartTime = Get-Date -Date $CalEvent.StartTime -ErrorAction Stop
            }
            
            if ($IsAllDay) {
                # For all-day events, ALWAYS set start time to midnight and end time to next day midnight
                # This ensures Microsoft Graph requirement of exactly 24 hours duration
                Write-Log "Setting all-day event: start at midnight, end at next day midnight" -Level "INFO"
                $StartTime = $StartTime.Date # Set to midnight of the start date
                $EndTime = $StartTime.AddDays(1) # Next day at midnight
            } else {
                # For timed events, parse the end time normally
                if ($CalEvent.EndTime -match '^\d{1,2}/\d{1,2}/\d{4}$' -or $CalEvent.EndTime -match '^\d{4}-\d{1,2}-\d{1,2}$') {
                    Write-Log "EndTime appears to be date-only format for timed event, adding 5:00 PM time component" -Level "INFO"
                    $EndTime = [DateTime]::Parse("$($CalEvent.EndTime) 5:00 PM")
                } else {
                    $EndTime = Get-Date -Date $CalEvent.EndTime -ErrorAction Stop
                }
            }
            
            Write-Log "Parsed StartTime: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
            Write-Log "Parsed EndTime: $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
            Write-Log "All-day event duration: $($EndTime.Subtract($StartTime).TotalHours) hours" -Level "INFO"
            
            if ($EndTime -le $StartTime) {
                Write-Log "End time must be after start time for event: $($CalEvent.Subject)" -Level "ERROR"
                $FailureCount++
                continue
            }
            
            # Additional validation for all-day events
            if ($IsAllDay) {
                $duration = $EndTime.Subtract($StartTime)
                if ($duration.TotalHours -lt 24) {
                    Write-Log "All-day event must have duration of at least 24 hours. Current duration: $($duration.TotalHours) hours for event: $($CalEvent.Subject)" -Level "ERROR"
                    $FailureCount++
                    continue
                }                Write-Log "All-day event validation passed: Duration = $($duration.TotalHours) hours" -Level "INFO"
            }
        } 
        catch {
            Write-Log "Invalid date format in event '$($CalEvent.Subject)': $_" -Level "ERROR"
            Write-Log "Failed date values - StartTime: '$($CalEvent.StartTime)', EndTime: '$($CalEvent.EndTime)'" -Level "ERROR"
            $FailureCount++
            continue
        }
        
        # Attendees
        if ([string]::IsNullOrWhiteSpace($CalEvent.AttendeeEmails)) {
            Write-Log "No attendee emails specified for event: $($CalEvent.Subject)" -Level "WARNING"
            $FailureCount++
            continue
        }
        
        $AttendeeEmails = $CalEvent.AttendeeEmails -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        if ($AttendeeEmails.Count -eq 0) {
            Write-Log "No valid attendee emails found for event: $($CalEvent.Subject)" -Level "WARNING"            $FailureCount++
            continue
        }
        
        # Optional fields
        $Location = if ($CalEvent.PSObject.Properties.Name -contains "Location") { $CalEvent.Location } else { "" }
        $Body = if ($CalEvent.PSObject.Properties.Name -contains "Body") { $CalEvent.Body } else { "" }
        
        # Check for organizer email in Excel, otherwise use default
        $OrganizerEmail = $DefaultOrganizerEmail  # Default organizer
        if ($CalEvent.PSObject.Properties.Name -contains "OrganizerEmail" -and -not [string]::IsNullOrWhiteSpace($CalEvent.OrganizerEmail)) {
            $OrganizerEmail = $CalEvent.OrganizerEmail
            Write-Log "Using organizer email from Excel: $OrganizerEmail for event: $($CalEvent.Subject)"
        } else {
            Write-Log "Using default organizer email: $DefaultOrganizerEmail for event: $($CalEvent.Subject)"
        }
          # Handle ShowAs property if it exists (Free, Busy, Tentative, OutOfOffice, WorkingElsewhere)
        $ShowAs = "Free"  # Default to Free
        if ($CalEvent.PSObject.Properties.Name -contains "ShowAs") {
            $validShowAs = @("Free", "Busy", "Tentative", "OutOfOffice", "WorkingElsewhere")
            if ($validShowAs -contains $CalEvent.ShowAs) {
                $ShowAs = $CalEvent.ShowAs
            } else {
                Write-Log "Invalid ShowAs value '$($CalEvent.ShowAs)' for event: $($CalEvent.Subject), using default 'Free'" -Level "WARNING"
            }
        }
          # Determine appropriate timezone based on event type
        # All-day events should use UTC to avoid timezone issues
        # Timed events should use local timezone (UTC+3 for Israel)
        $EventTimeZone = if ($IsAllDay) {
            "UTC"  # All-day events use UTC
        } else {
            "Asia/Jerusalem"  # Timed events use local timezone (UTC+3)
        }
        
        # Log debug information before creating the event
        Write-Log "Event details:" -Level "INFO"
        Write-Log "  Subject: $($CalEvent.Subject)" -Level "INFO"
        Write-Log "  Start: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
        Write-Log "  End: $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level "INFO"
        Write-Log "  IsAllDay: $IsAllDay" -Level "INFO"
        Write-Log "  TimeZone: $EventTimeZone" -Level "INFO"
        Write-Log "  ShowAs: $ShowAs" -Level "INFO"
        Write-Log "  Organizer: $OrganizerEmail" -Level "INFO"
        Write-Log "  Attendees: $($AttendeeEmails -join ', ')" -Level "INFO"
        
        # Create event for group members or individuals
        $Result = New-GraphCalendarEventForAllGroupMembers `
            -Subject $CalEvent.Subject `
            -StartTime $StartTime `
            -EndTime $EndTime `
            -AttendeeEmails $AttendeeEmails `
            -TimeZone $EventTimeZone `
            -Location $Location `
            -Body $Body `
            -IsAllDay $IsAllDay `
            -ShowAs $ShowAs `
            -OrganizerEmail $OrganizerEmail
        
        if ($Result) {
            $SuccessCount++
            Write-Log "Successfully processed event: $($CalEvent.Subject)" -Level "INFO"        } else {
            $FailureCount++
            Write-Log "Failed to process event: $($CalEvent.Subject)" -Level "ERROR"
        }
    }
    catch {
        Write-Log "Failed to process event '$($CalEvent.Subject)': $_" -Level "ERROR"
        $FailureCount++
    }
}

###########################################################################
# Cleanup
###########################################################################
try {
    Write-Log "Disconnecting from Microsoft Graph"
    Disconnect-MgGraph
    Write-Log "Graph API disconnected successfully"
    
    # Disconnect from Exchange Online if connected
    if ($Global:ExchangeConnected) {
        Write-Log "Disconnecting from Exchange Online"
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            $Global:ExchangeConnected = $false
            Write-Log "Exchange Online disconnected successfully"
        }
        catch {
            Write-Log "Error disconnecting from Exchange Online: $_" -Level "WARNING"
        }
    }
}
catch {
    Write-Log "Error during cleanup: $_" -Level "WARNING"
}

###########################################################################
# Summary
###########################################################################
Write-Log "Script execution completed"
Write-Log "Summary: $SuccessCount events processed successfully, $FailureCount events failed"
Write-Log "Deletion Summary: $DeletionCount events deleted successfully, $DeletionFailureCount deletion failures"

$TotalOperations = $SuccessCount + $FailureCount + $DeletionCount + $DeletionFailureCount
Write-Log "Total Operations: $TotalOperations (Created: $SuccessCount, Creation Failures: $FailureCount, Deleted: $DeletionCount, Deletion Failures: $DeletionFailureCount)"

if ($FailureCount -gt 0 -or $DeletionFailureCount -gt 0) {
    Write-Log "Some operations failed to complete. Check the log file for details: $LogPath" -Level "WARNING"
    exit 1
} else {
    Write-Log "All operations completed successfully"
    exit 0
}
