# Microsoft Graph - Utility Scripts for PowerShell
This repository contains some PowerShell scripts that improve my daily development or administration work. These scripts are all using the [Microsoft Graph PowerShell module](https://docs.microsoft.com/en-us/graph/powershell/installation?WT.mc_id=M365-MVP-5003400.  
All these scripts work on PowerShell Core / PowerShell 7.1.

## Prerequisites
These scripts require the following PowerShell modules:
- [Microsoft.Graph](https://docs.microsoft.com/en-us/graph/powershell/installation?WT.mc_id=M365-MVP-5003400)
- [Microsoft.PowerShell.ConsoleGuiTools](https://github.com/powershell/GraphicalTools)

## Scripts
### Add-ApplicationPermissionsToServicePrincipal
#### Purpose
This scripts helps to assign Application Permissions to a Service Principal and grant admin consent. This is especially useful for assigning Application Permissions to Managed Identities, as there currently is no UI in the Azure Portal that can achieve the same result (and Managed Identities do not have an associated App Registration where this can be achieved).

#### Usage
Add permission scopes from Microsoft Graph to your Service Principal:
```
. ./AddApplicationPermissionsToServicePrincipal.ps1
Add-ApplicationPermissionsToServicePrincipal -ServicePrincipalAppId <YourServicePrincipalApplicationId> -Api MicrosoftGraph -ApplicationPermissionScopes "Team.Create", "Group.ReadWrite.All"
```

Add permission scopes from SharePoint to your Service Principal:
```
. ./AddApplicationPermissionsToServicePrincipal.ps1
Add-ApplicationPermissionsToServicePrincipal -ServicePrincipalAppId <YourServicePrincipalApplicationId> -Api SharePoint -ApplicationPermissionScopes "Sites.ReadWrite.All"
```

Add permission scopes from another API to your Service Principal
```
. ./AddApplicationPermissionsToServicePrincipal.ps1
Add-ApplicationPermissionsToServicePrincipal -ServicePrincipalAppId <YourServicePrincipalApplicationId> -ApiServicePrincipalAppId <ApiServicePrincipalApplicationId> -ApplicationPermissionScopes "Sites.ReadWrite.All"
```

### Revoke-PermissionConsentFromServicePrincipal
#### Purpose
Provide an intuitive way to revoke Application Permissions and Delegated Permissions that have been consented for one or more Service Principals. When selecting Delegated Permissions, an option is presented to select Admin consented and/or User consented permissions. With User consented permissions you can pick for which user specifically you'd like to revoke the consent.

#### Usage
```
. ./RevokePermissionConsentFromServicePrincipal.ps1
Revoke-PermissionConsentFromServicePrincipal
```

# Disclaimer
THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
