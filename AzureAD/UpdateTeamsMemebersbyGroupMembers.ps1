# Connect to Microsoft Teams and Microsoft Graph using Managed-Identity
Connect-MicrosoftTeams -Identity
Connect-MgGraph -Identity

# Get members of the <GroupName> group using Microsoft Graph
try {
    $users = Get-MgGroupMember -GroupId 'xxxxxx-xxxx-xxxx-xx-xxxxxxxx'
    Write-Output "Successfully retrieved members of <GroupName> group."
} catch {
    Write-Error "Failed to get members of <GroupName> group: $_"
}

# Get members of the <GroupName> group using Microsoft Graph
try {
    $users2 = Get-MgGroupMember -GroupId 'xxxxxx-xxxx-xxxx-xx-xxxxxxxx'
    Write-Output "Successfully retrieved members of <GroupName> group."
} catch {
    Write-Error "Failed to get members of <GroupName> group: $_"
}

# Get members of the <TeamName> team using Microsoft Teams PowerShell module
try {
    $Teams = Get-TeamUser -GroupId 'xxxxxx-xxxx-xxxx-xx-xxxxxxxx' | Select-Object -ExpandProperty User
    Write-Output "Successfully retrieved members of <TeamName> team."
} catch {
    Write-Error "Failed to get members of <TeamName> team: $_"
}

# Compare members of AAD groups to Teams group
try {
    $members = Compare-Object -ReferenceObject $users.AdditionalProperties.userPrincipalName -DifferenceObject $Teams | Where-Object {$_.SideIndicator -eq "<="} | Select-Object -ExpandProperty InputObject
    Write-Output "Comparison of <GroupName> group to Teams group completed."
} catch {
    Write-Error "Failed to compare <GroupName> group members with Teams group: $_"
}

try {
    $members2 = Compare-Object -ReferenceObject $users2.AdditionalProperties.userPrincipalName -DifferenceObject $Teams | Where-Object {$_.SideIndicator -eq "<="} | Select-Object -ExpandProperty InputObject
    Write-Output "Comparison of <GroupName> group to Teams group completed."
} catch {
    Write-Error "Failed to compare <GroupName> group members with Teams group: $_"
}

# Add missing members from <GroupName> AAD group to Teams group
foreach ($user in $members) {
    try {
        Add-TeamUser -GroupId 'xxxxxx-xxxx-xxxx-xx-xxxxxxxx' -User $user
        Write-Output "User from <GroupName> was added: $user"
    } catch {
        Write-Error "Failed to add user from <GroupName> to Teams group: $user. Error: $_"
    }
}

# Add missing members from <GroupName> AAD group to Teams group
foreach ($user in $members2) {
    try {
        Add-TeamUser -GroupId 'xxxxxx-xxxx-xxxx-xx-xxxxxxxx' -User $user
        Write-Output "User from <GroupName> was added: $user"
    } catch {
        Write-Error "Failed to add user from <GroupName> to Teams group: $user. Error: $_"
    }
}
