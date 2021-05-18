<#
.SYNOPSIS

Adds application permissions defined by an API to a service principal

.DESCRIPTION

Adds application permissions defined by an API to a service principal and grants admin consent for these permissions scopes.
Requires a connection with Connect-MgGraph that has scopes consented from:
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-get?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignments?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400
- https://docs.microsoft.com/en-us/graph/api/serviceprincipal-post-approleassignments?view=graph-rest-1.0&tabs=http&WT.mc_id=M365-MVP-5003400

.PARAMETER ServicePrincipalAppId
Application Id of the Service Principal to which you are assigning the application permission(s)

.PARAMETER ServicePrincipalAppDisplayName
Display name of the Service Principal to which you are assigning the application permission(s)

.PARAMETER ApiServicePrincipalAppId
Application id of the Service Principal (the API) which has defined the application permission(s)

.PARAMETER ApiServicePrincipalAppDisplayName
Display name of the Service Principal (the API) which has defined the application permission(s)

.PARAMETER Api
Predefined Service Principals for which application permission(s) can be assigned. When selecting one of these, ApiServicePrincipalAppId and ApiServicePrincipalDisplayName aren't used.

.PARAMETER ApplicationPermissionScopes
The application permission scopes you want to assign to the Service Principal, and are defined by the Api Service Principal

#>

function Add-ApplicationPermissionsToServicePrincipal {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True, ParameterSetName = "ByAppIdWithApi")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByAppId")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByAppIdAndDisplayName")]
        [string]
        $ServicePrincipalAppId,
        [Parameter(Mandatory=$True, ParameterSetName = "ByDisplayNameWithApi")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByDisplayNameAndAppId")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByDisplayName")]
        [string]
        $ServicePrincipalAppDisplayName,
        [Parameter(Mandatory=$True, ParameterSetName = "ByAppIdWithApi")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByDisplayNameWithApi")]
        [ValidateSet("MicrosoftGraph", "SharePoint")]
        $Api,
        [Parameter(Mandatory=$True, ParameterSetName = "ByAppId")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByDisplayNameAndAppId")]
        [string]
        $ApiServicePrincipalAppId,
        [Parameter(Mandatory=$True, ParameterSetName = "ByAppIdAndDisplayName")]
        [Parameter(Mandatory=$True, ParameterSetName = "ByDisplayName")]
        [string]
        $ApiServicePrincipalAppDisplayName,
        [Parameter(Mandatory=$True)]
        [string[]]
        $ApplicationPermissionScopes
    )
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

        if ($PSCmdlet.ParameterSetName -eq "ByAppIdWithApi" -or $PSCmdlet.ParameterSetName -eq "ByDisplayNameWithApi") {
            switch ($Api) {
                "MicrosoftGraph" { $ApiServicePrincipalAppId = "00000003-0000-0000-c000-000000000000" }
                "SharePoint" { $ApiServicePrincipalAppId = "00000003-0000-0ff1-ce00-000000000000" }
            }
        }

        $ServicePrincipalFilter = ""
        if ($PSCmdlet.ParameterSetName -eq "ByAppIdWithApi" -or $PSCmdlet.ParameterSetName -eq "ByAppId" -or $PSCmdlet.ParameterSetName -eq "ByAppIdAndDisplayName") {
            $ServicePrincipalFilter = "appId eq '$ServicePrincipalAppId'"
        } elseif ($PSCmdlet.ParameterSetName -eq "ByDisplayNameWithApi" -or $PSCmdlet.ParameterSetName -eq "ByDisplayNameAndAppId" -or $PSCmdlet.ParameterSetName -eq "ByDisplayName") {
            $ServicePrincipalFilter = "displayName eq '$ServicePrincipalAppDisplayName'"
        }

        $ApiServicePrincipalFilter = ""
        if ($PSCmdlet.ParameterSetName -eq "ByAppIdWithApi" -or $PSCmdlet.ParameterSetName -eq "ByDisplayNameWithApi" -or $PSCmdlet.ParameterSetName -eq "ByAppId" -or $PSCmdlet.ParameterSetName -eq "ByDisplayNameAndAppId") {
            $ApiServicePrincipalFilter = "appId eq '$ApiServicePrincipalAppId'"
        } elseif ($PSCmdlet.ParameterSetName -eq "ByAppIdWithApi" -or $PSCmdlet.ParameterSetName -eq "ByDisplayNameWithApi" -or $PSCmdlet.ParameterSetName -eq "ByAppIdAndDisplayName" -or $PSCmdlet.ParameterSetName -eq "ByDisplayName") {
            $ApiServicePrincipalFilter = "displayName eq '$ApiServicePrincipalAppDisplayName'"
        }
    }
    PROCESS {
        $ServicePrincipal = Get-MgServicePrincipal -Filter $ServicePrincipalFilter
        if ($null -eq $ServicePrincipal) { Throw "Could not find the specified Service Principal" }
        if ($ServicePrincipal -is [array]) { Throw "Multiple Service Principals found that match the request, maybe try using -ServicePrincipalAppId instead of -ServicePrincipalDisplayName" }

        $ApiServicePrincipal = Get-MgServicePrincipal -Filter $ApiServicePrincipalFilter
        if ($null -eq $ApiServicePrincipal) { Throw "Could not find the specified Api Service Principal" }
        if ($ApiServicePrincipal -is [array]) { Throw "Multiple Api Service Principals found that match the request, maybe try using -ApiServicePrincipalAppId instead of -ApiServicePrincipalDisplayName" }

        Foreach ($Scope in $ApplicationPermissionScopes) {
            Write-Host "`nGetting App Role '$Scope'"
            $AppRole = $ApiServicePrincipal.AppRoles | Where-Object {$_.Value -eq $Scope -and $_.AllowedMemberTypes -contains "Application"}
            if ($null -eq $AppRole) { Write-Error "Could not find the specified App Role on the Api Service Principal"; continue; }
            if ($AppRole -is [array]) { Write-Error "Multiple App Roles found that match the request"; continue; }
            Write-Host "Found App Role, Id '$($AppRole.Id)'"

            $ExistingRoleAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id | Where-Object { $_.AppRoleId -eq $AppRole.Id }
            if ($null -eq $existingRoleAssignment) {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -PrincipalId $ServicePrincipal.Id -ResourceId $ApiServicePrincipal.Id -AppRoleId $AppRole.Id
            } else {
                Write-Host "App Role has already been assigned, skipping"
            }
        }
    }
    END {
        Write-Verbose ('[{0}] Confirm={1} ConfirmPreference={2} WhatIf={3} WhatIfPreference={4}' -f $MyInvocation.MyCommand, $Confirm, $ConfirmPreference, $WhatIf, $WhatIfPreference)
    }
}