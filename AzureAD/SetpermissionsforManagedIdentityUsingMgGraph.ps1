$tenantID = "12345678-1234-1234-1234-0123456789xx"
    $graphAppId = "00000003-0000-0000-c000-000000000000"
    $permissions = @("WindowsUpdates.ReadWrite.All", "Device.Read.All")
    $managedIdentities = @("mi-1", "mi-2", "mi-3") # Names of system-assigned or user-assigned managed service identity. (System-assigned use same name as resource).
    Connect-MgGraph -TenantId $tenantID -NoWelcome -Scopes "AppRoleAssignment.ReadWrite.All", "Directory.Read.All"
    $sp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
    $managedIdentities | ForEach-Object {
        $msi = Get-MgServicePrincipal -Filter "displayName eq '$_'"
        $appRoles = $sp.AppRoles | Where-Object {($_.Value -in $permissions) -and ($_.AllowedMemberTypes -contains "Application")}
        $appRoles | ForEach-Object {
            $appRoleAssignment = @{
                "PrincipalId" = $msi.Id
                "ResourceId" = $sp.Id
                "AppRoleId" = $_.Id
            }
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appRoleAssignment.PrincipalId -BodyParameter $appRoleAssignment -Verbose
        }
    }
    Disconnect-MgGraph
