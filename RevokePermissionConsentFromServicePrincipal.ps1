<#
.SYNOPSIS

Revokes admin consented and user consented permissions for a Service Principal

.DESCRIPTION

Revokes admin consented and user consented permissions for a Service Principal. Offers a choice between Admin Consent and/or User Consent, and allows to choose for which user(s) specifically you want to remove the consent.
Requires a connection with Connect-MgGraph that has scopes consented from:
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-list?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignments?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-delete-approleassignments?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-list-oauth2permissiongrants?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400
- https://docs.microsoft.com/en-us/graph/api/oauth2permissiongrant-delete?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400

#>

function Revoke-PermissionConsentFromServicePrincipal {
    [CmdletBinding()]
    param ()
    BEGIN {
        if (-not $PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
        }
        if (-not $PSBoundParameters.ContainsKey('Confirm')) {
            $ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference')
        }
        if (-not $PSBoundParameters.ContainsKey('WhatIf')) {
            $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference')
        }
        Write-Verbose ('[{0}] Confirm={1} ConfirmPreference={2} WhatIf={3} WhatIfPreference={4}' -f $MyInvocation.MyCommand, $Confirm, $ConfirmPreference, $WhatIf, $WhatIfPreference)

        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Throw "Microsoft Graph PowerShell module not installed. Please install and re-run script."
        }

        if (-not (Get-Module -ListAvailable -Name Microsoft.PowerShell.ConsoleGuiTools)) {
            Throw "Console Gui Tools PowerShell module not installed. Please install and re-run script."
        }

        $consentTypeOptions = @()
        $consentTypeOptions += [PSCustomObject]@{
            Name = "Admin consent";
            Type = "AllPrincipals"
        }
        $consentTypeOptions += [PSCustomObject]@{
            Name = "User consent";
            Type = "Principal"
        }

        $permissionTypeOptions = @()
        $permissionTypeOptions += [PSCustomObject]@{
            Name = "Application Permissions";
            Type = "application"
        }
        $permissionTypeOptions += [PSCustomObject]@{
            Name = "Delegated Permissions";
            Type = "delegated"
        }
    }
    PROCESS {
        function RevokeApplicationPermissions([Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal] $servicePrincipal) {
            Write-Host "`n`tRevoking Application Permissions"
            $permissionGrantsForApplication = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id | Sort-Object ResourceDisplayName
            if ($permissionGrantsForApplication) {
                $prevResourceName = ""
                foreach ($permissionGrantForApplication in $permissionGrantsForApplication) {
                    if ($permissionGrantForApplication.ResourceDisplayName -ne $prevResourceName) {
                        Write-Host "`t`t$($permissionGrantForApplication.ResourceDisplayName)"
                        $prevResourceName = $permissionGrantForApplication.ResourceDisplayName
                    }
                    Write-Host "`t`t- $($permissionGrantForApplication.AppRoleId)"
                }
                do {
                    $answer = Read-host -prompt "`nDo you wish to remove all these permissions (Y/N)?"
                } until (-not [string]::isnullorempty($answer))
                if ($answer -eq 'Y' -or $answer -eq 'y') {
                    foreach ($permissionGrantForApplication in $permissionGrantsForApplication) {
                        Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -AppRoleAssignmentId $permissionGrantForApplication.Id
                    }
                    Write-Host "`t`tRemoved"
                }
                else {
                    Write-Host "`t`tNo change made"
                }
            }
            else {
                Write-Host "`nNo Application Permissions permissions found for $($servicePrincipal.DisplayName)`n"
            }
        }

        function RevokeDelegatedPermissionsForUser($permissionGrantsForPrincipal) {
            foreach ($permissionGrantForPrincipal in $permissionGrantsForPrincipal) {
                $apiServicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $permissionGrantForPrincipal.ResourceId
                Write-Host "`t`t$($apiServicePrincipal.DisplayName)"
                $scopes = $permissionGrantForPrincipal.Scope -split " "
                foreach ($scope in $scopes) {
                    Write-Host "`t`t- $scope"
                }
            }
            do {
                $answer = Read-host -prompt "`nDo you wish to remove all these permissions (Y/N)?"
            } until (-not [string]::isnullorempty($answer))
            if ($answer -eq 'Y' -or $answer -eq 'y') {
                foreach ($permissionGrantForPrincipal in $permissionGrantsForPrincipal) {
                    Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $permissionGrantForPrincipal.Id
                }
                Write-Host "`t`tRemoved"
            }
            else {
                Write-Host "`t`tNo change made"
            }
        }

        function RevokeDelegatedPermissions([Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal] $servicePrincipal) {
            Write-Host "`n`tRevoking Delegated Permissions"
            $consentTypes = $consentTypeOptions | Out-ConsoleGridView -Title "Select consent type(s) to revoke from '$($servicePrincipal.DisplayName)'"
            foreach ($consentType in $consentTypes) {
                $permissionGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $servicePrincipal.Id -All | Where-Object { $_.ConsentType -eq $consentType.Type }
                
                if ($permissionGrants) {
                    switch ($consentType.Type) {
                        "AllPrincipals" {
                            $permissionGrantsForPrincipal = $permissionGrants | Where-Object { $null -eq $_.PrincipalId }
                            Write-Host "`tRevoking admin consented permissions"
                            RevokeDelegatedPermissionsForUser($permissionGrantsForPrincipal) 
                        }
                        "Principal" {
                            $consentedUsers = @()
                            foreach ($permissionGrant in $permissionGrants) {
                                $consentedUsers += Get-MgUser -UserId $permissionGrant.PrincipalId -Property "DisplayName", "UserPrincipalName", "Id", "Mail", "UserType"
                            }
                            $users = $consentedUsers | sort-object Displayname | Out-ConsoleGridView -title "Select user(s) from which to remove consent"

                            foreach ($user in $users) {
                                $permissionGrantsForPrincipal = $permissionGrants | Where-Object { $_.PrincipalId -eq $user.Id }
                                Write-Host "`tRevoking permissions consented by '$($user.DisplayName)'"
                                RevokeDelegatedPermissionsForUser($permissionGrantsForPrincipal)
                            }
                        }
                    }
                }
                else {
                    Write-Host "`nNo $($consentType.Name) permissions found for $($servicePrincipal.DisplayName)`n"
                }
            }
        }

        $servicePrincipals = Get-MgServicePrincipal -All | Sort-Object displayName | Out-ConsoleGridView -Title "Select Enterprise Application(s)"
        foreach ($servicePrincipal in $servicePrincipals) {
            Write-Host "Modifying consent for '$($servicePrincipal.DisplayName)'"

            $permissionTypes = $permissionTypeOptions | Out-ConsoleGridView -Title "Select permission type(s) to revoke from '$($servicePrincipal.DisplayName)'"
            foreach ($permissionType in $permissionTypes) {
                switch ($permissionType.Type) {
                    "application" { RevokeApplicationPermissions($servicePrincipal) }
                    "delegated" { RevokeDelegatedPermissions($servicePrincipal) }
                }
            }
        }
    }
    END {
        Write-Verbose ('[{0}] Confirm={1} ConfirmPreference={2} WhatIf={3} WhatIfPreference={4}' -f $MyInvocation.MyCommand, $Confirm, $ConfirmPreference, $WhatIf, $WhatIfPreference)
    }
}