<#
  New-AppReg.ps1  -  six permissions + auto-consent
#>

param(
    [string]   $DisplayName = 'secto',
    [string[]] $RedirectUris = @('https://localhost'),
    [ValidateSet('SingleTenant','MultiTenant','MultiTenantAndPersonal')]
    [string]   $Audience  = 'SingleTenant',
    [string]   $TenantId,
    [switch]   $CreateClientSecret
)

# -- map friendly audience names -----------------------------------------------
$audienceMap = @{
    SingleTenant            = 'AzureADMyOrg'
    MultiTenant             = 'AzureADMultipleOrgs'
    MultiTenantAndPersonal  = 'AzureADandPersonalMicrosoftAccount'
}

# -- install only required Microsoft.Graph sub-modules ------------------------
$requiredModules = @(
    'Microsoft.Graph.Authentication',              # Always required for Connect-MgGraph
    'Microsoft.Graph.Applications',                # For app registration operations
    'Microsoft.Graph.Users',                       # For user creation and management
    'Microsoft.Graph.Identity.DirectoryManagement', # For role assignments
    'Microsoft.Graph.Identity.SignIns'               # For OAuth2PermissionGrant cmdlets
)

$modulesToInstall = @()
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable $module)) {
        $modulesToInstall += $module
    }
}

if ($modulesToInstall.Count -gt 0) {
    Write-Host "Installing required Microsoft.Graph modules..." -ForegroundColor Yellow
    foreach ($module in $modulesToInstall) {
        Write-Host "  Installing $module..." -ForegroundColor Cyan
        Install-Module $module -Scope CurrentUser -Force -ErrorAction Stop
    }
    Write-Host "[OK] Required modules installed successfully!" -ForegroundColor Green
    Write-Host ""
}

# -- sign in with the five admin scopes (no explicit import needed) ----------
$scopes = @(
    'Application.ReadWrite.All',
    'AppRoleAssignment.ReadWrite.All',
    'DelegatedPermissionGrant.ReadWrite.All',
    'User.ReadWrite.All',
    'RoleManagement.ReadWrite.Directory'
)
$connect  = @{ Scopes = $scopes }
if ($TenantId) { $connect.TenantId = $TenantId }

# Connect-MgGraph will auto-load Microsoft.Graph.Authentication module
Connect-MgGraph @connect
Write-Host "[INFO] Connected to Microsoft Graph for tenant '$TenantId'." -ForegroundColor Cyan

# -- Graph resource service-principal (only once) ------------------------------
$graphAppId = '00000003-0000-0000-c000-000000000000'
Write-Host "[INFO] Fetching Microsoft Graph service principal ..." -ForegroundColor Cyan
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Top 1
if (-not $graphSp) { throw "Unable to retrieve Microsoft Graph service principal" }

# -- permission IDs we need (Microsoft Graph only) --------------------
$permIds = @{
    # Microsoft Graph
    AuditLog_Read_All                 = 'b0afded3-3588-46d8-8b3d-9842eff778da'
    AuditLogsQuery_Read_All           = '5e1e9171-754d-478c-812c-f1755a9a4c2d'
    Directory_Read_All                = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'
    Domain_Read_All                   = 'dbb9058a-0e50-45d7-ae91-66909b5d4664'
    Organization_Read_All             = '498476ce-e0fe-48b0-b801-37ba7e2685c6'
    Policy_Read_All                   = '246dd0d5-5bd0-4def-940b-0421030a5b68'
    SharePointTenantSettings_Read_All = '83d4163d-a2d8-4d3b-9695-4ae3ca98f888'
    User_Read                         = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'
    User_RevokeSessions_All           = '77f3a031-c388-4f99-b373-dc68676a979e'

    # Office 365 Exchange Online
    Exchange_ManageAsApp              = 'dc50a0fb-09a3-484d-be87-e023b12c6440'
    ReportingWebService_Read_All      = 'b4d5a5c7-c085-487f-b922-ef0d6ebde6b1'

    # Office 365 Management APIs
    ActivityFeed_Read                 = '594c1fb6-4f81-4475-ae41-0c394909246c'
    ActivityFeed_ReadDlp              = '4807a72c-ad38-4250-94c9-4eabfe26cd55'
    ServiceHealth_Read                = 'e2cea78f-e743-4d8f-a16a-75b629a038ae'
}

# -- target service principals for required APIs ------------------------
$exchangeSp = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'"
if (-not $exchangeSp) {
    throw "Unable to locate the Office 365 Exchange Online service principal."
}

$managementSp = Get-MgServicePrincipal -Filter "appId eq 'c5393580-f805-4401-95e8-94b7a6ef2fc2'"
if (-not $managementSp) {
    throw "Unable to locate the Office 365 Management APIs service principal."
}

# -- #1  Build the resourceAccess array for Microsoft Graph ------
$graphResourceAccess = @()

# Application roles and delegated scopes
$graphResourceAccess += @{ id = $permIds.AuditLog_Read_All;                 type = 'Role' }
$graphResourceAccess += @{ id = $permIds.AuditLogsQuery_Read_All;           type = 'Role' }
$graphResourceAccess += @{ id = $permIds.Directory_Read_All;                type = 'Role' }
$graphResourceAccess += @{ id = $permIds.Domain_Read_All;                   type = 'Role' }
$graphResourceAccess += @{ id = $permIds.Organization_Read_All;             type = 'Role' }
$graphResourceAccess += @{ id = $permIds.Policy_Read_All;                   type = 'Role' }
$graphResourceAccess += @{ id = $permIds.SharePointTenantSettings_Read_All; type = 'Role' }
$graphResourceAccess += @{ id = $permIds.User_RevokeSessions_All;           type = 'Role' }
$graphResourceAccess += @{ id = $permIds.User_Read;                         type = 'Scope' }

# -- #2  Wrap in requiredResourceAccess ----------------
$requiredResourceAccess = @(
    @{
        resourceAppId  = $graphSp.AppId
        resourceAccess = $graphResourceAccess
    },
    @{
        resourceAppId  = $exchangeSp.AppId
        resourceAccess = @(
            @{ id = $permIds.Exchange_ManageAsApp;         type = 'Role' }
            @{ id = $permIds.ReportingWebService_Read_All; type = 'Role' }
        )
    },
    @{
        resourceAppId  = $managementSp.AppId
        resourceAccess = @(
            @{ id = $permIds.ActivityFeed_Read;    type = 'Role' }
            @{ id = $permIds.ActivityFeed_ReadDlp; type = 'Role' }
            @{ id = $permIds.ServiceHealth_Read;   type = 'Role' }
        )
    }
)

# -- helper: fetch app by display name -----------------------------------------
function Get-AppByName ($name) {
    $apps = Get-MgApplication -Filter "displayName eq '$name'" -ConsistencyLevel eventual -Count c -All
    if ($apps.Count -gt 1) {
        throw "Found $($apps.Count) applications named '$name'. Delete or rename duplicates to proceed."
    }
    return $apps | Select-Object -First 1
}

# -- create or update ----------------------------------------------------------
$app = Get-AppByName $DisplayName
if ($null -eq $app) {
    Write-Host "Creating application '$DisplayName' ..."
    $app = New-MgApplication `
              -DisplayName            $DisplayName `
              -SignInAudience         $audienceMap[$Audience] `
              -Web                    @{ RedirectUris = $RedirectUris } `
              -RequiredResourceAccess $requiredResourceAccess
    $sp  = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "[OK] Created app : $($app.AppId)"
} else {
    Write-Host "Updating application '$DisplayName' ..."
    $sp  = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"
    if (-not $sp) { $sp = New-MgServicePrincipal -AppId $app.AppId }

    # Merge redirect URIs (existing + new, unique)
    $mergedRedirectUris = @($app.Web.RedirectUris + $RedirectUris) | Select-Object -Unique

    # Merge required resource access to keep existing permissions intact
    $existingRRA = @($app.RequiredResourceAccess)
    $graphRRA    = $existingRRA | Where-Object { $_.resourceAppId -eq $graphSp.AppId }

    if ($graphRRA) {
        $existingIds  = $graphRRA.resourceAccess.id
        $missingItems = $graphResourceAccess | Where-Object { $existingIds -notcontains $_.id }
        $graphRRA.resourceAccess += $missingItems
    }
    else {
        $existingRRA += @{ resourceAppId = $graphSp.AppId; resourceAccess = $graphResourceAccess }
    }

    Update-MgApplication -ApplicationId $app.Id `
                         -Web @{ RedirectUris = $mergedRedirectUris } `
                         -RequiredResourceAccess $existingRRA
}

# -- optional: client secret ---------------------------------------------------
if ($CreateClientSecret) {
    $secret = Add-MgApplicationPassword -ApplicationId $app.Id `
                 -PasswordCredential @{
                     displayName = 'automation-secret'
                     endDateTime = (Get-Date).AddYears(1)
                 }
}

# -- admin consent --------------------------------------------
Write-Host "`nGranting admin consent ..."

# fetch current role assignments for this principal -> Graph
# Note: Filter by principalId is not supported, so we get all and filter client-side
$existingGraphRoles = Get-MgServicePrincipalAppRoleAssignment `
                        -ServicePrincipalId $graphSp.Id `
                        -All | Where-Object { $_.PrincipalId -eq $sp.Id }

# #1  application-role consent for Microsoft Graph
foreach ($roleId in @(
        $permIds.AuditLog_Read_All,
        $permIds.AuditLogsQuery_Read_All,
        $permIds.Directory_Read_All,
        $permIds.Domain_Read_All,
        $permIds.Organization_Read_All,
        $permIds.Policy_Read_All,
        $permIds.SharePointTenantSettings_Read_All,
        $permIds.User_RevokeSessions_All)
) {
    if ($existingGraphRoles.AppRoleId -contains $roleId) {
        Write-Verbose "Role $roleId already present - skipping"
        continue
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $graphSp.Id `
        -PrincipalId        $sp.Id       `
        -ResourceId         $graphSp.Id  `
        -AppRoleId          $roleId | Out-Null
}

# application-role consent for Office 365 Exchange Online
$existingExchangeRoles = Get-MgServicePrincipalAppRoleAssignment `
                            -ServicePrincipalId $exchangeSp.Id `
                            -All | Where-Object { $_.PrincipalId -eq $sp.Id }

foreach ($roleId in @(
        $permIds.Exchange_ManageAsApp,
        $permIds.ReportingWebService_Read_All)
) {
    if ($existingExchangeRoles.AppRoleId -contains $roleId) {
        Write-Verbose "Role $roleId already present - skipping"
        continue
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $exchangeSp.Id `
        -PrincipalId        $sp.Id          `
        -ResourceId         $exchangeSp.Id  `
        -AppRoleId          $roleId | Out-Null
}

# application-role consent for Office 365 Management APIs
$existingManagementRoles = Get-MgServicePrincipalAppRoleAssignment `
                              -ServicePrincipalId $managementSp.Id `
                              -All | Where-Object { $_.PrincipalId -eq $sp.Id }

foreach ($roleId in @(
        $permIds.ActivityFeed_Read,
        $permIds.ActivityFeed_ReadDlp,
        $permIds.ServiceHealth_Read)
) {
    if ($existingManagementRoles.AppRoleId -contains $roleId) {
        Write-Verbose "Role $roleId already present - skipping"
        continue
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $managementSp.Id `
        -PrincipalId        $sp.Id           `
        -ResourceId         $managementSp.Id `
        -AppRoleId          $roleId | Out-Null
}

# #2  delegated scope (User.Read) for Microsoft Graph
$grant = Get-MgOauth2PermissionGrant `
           -Filter "clientId eq '$($sp.Id)' and resourceId eq '$($graphSp.Id)'" | Select-Object -First 1

if (-not $grant -or ($grant.Scope -notmatch '\bUser\.Read\b')) {
    New-MgOauth2PermissionGrant -BodyParameter @{
        clientId    = $sp.Id
        consentType = 'AllPrincipals'
        resourceId  = $graphSp.Id
        scope       = 'User.Read'
    } | Out-Null
}

Write-Host "[OK] Admin consent granted.`n" -ForegroundColor Green

# -- display required information ----------------------------------------------
Write-Host "### Required Information for Secto Application ###" -ForegroundColor Cyan
Write-Host ""
$context = Get-MgContext
Write-Host " => Tenant ID.......................: '$($context.TenantId)' <---" -ForegroundColor Yellow
Write-Host " => Client ID (App ID)..............: '$($app.AppId)' <---" -ForegroundColor Yellow
if ($CreateClientSecret) {
    Write-Host " => Client Secret...................: '$($secret.SecretText)' <---" -ForegroundColor Yellow
    Write-Host " => Secret Valid Until..............: '$($secret.EndDateTime)'" -ForegroundColor Yellow
}
Write-Host ""

Disconnect-MgGraph
