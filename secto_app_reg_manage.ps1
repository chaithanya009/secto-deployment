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

# -- early validation of TenantId (fail fast) ----------------------------------
if ($TenantId) {
    # Starts with a letter, 3-27 chars total, letters/digits/hyphens, ends with letter or digit
    $tenantPattern = '^[A-Za-z][A-Za-z0-9-]{1,25}[A-Za-z0-9]\.onmicrosoft\.com$'
    if (-not ([regex]::IsMatch($TenantId, $tenantPattern))) {
        Write-Host "[ERROR] '$TenantId' is not a valid Entra tenant domain." -ForegroundColor Red
        Write-Host "        You can locate this domain in Microsoft Entra ID → Overview → Primary domain." -ForegroundColor Yellow
        exit 1   # graceful termination without verbose stack trace
    }
}

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

# -- permission IDs we need ----------------------------------------------------
$permIds = @{
    AuditLog_Read_All                 = 'b0afded3-3588-46d8-8b3d-9842eff778da'
    AuditLogsQuery_Read_All           = '5e1e9171-754d-478c-812c-f1755a9a4c2d'
    Directory_Read_All                = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'
    Domain_Read_All                   = 'dbb9058a-0e50-45d7-ae91-66909b5d4664'
    Organization_Read_All             = '498476ce-e0fe-48b0-b801-37ba7e2685c6'
    Policy_Read_All                   = '246dd0d5-5bd0-4def-940b-0421030a5b68'
    SharePointTenantSettings_Read_All = '83d4163d-a2d8-4d3b-9695-4ae3ca98f888'
    User_Read                         = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'
    User_RevokeSessions_All           = '77f3a031-c388-4f99-b373-dc68676a979e'
}

# -- #1  Build the resourceAccess array FIRST -------------
$resourceAccess = @()

# eight application roles
$resourceAccess += @{ id = $permIds.AuditLog_Read_All;                 type = 'Role' }
$resourceAccess += @{ id = $permIds.AuditLogsQuery_Read_All;           type = 'Role' }
$resourceAccess += @{ id = $permIds.Directory_Read_All;                type = 'Role' }
$resourceAccess += @{ id = $permIds.Domain_Read_All;                   type = 'Role' }
$resourceAccess += @{ id = $permIds.Organization_Read_All;             type = 'Role' }
$resourceAccess += @{ id = $permIds.Policy_Read_All;                   type = 'Role' }
$resourceAccess += @{ id = $permIds.SharePointTenantSettings_Read_All; type = 'Role' }
$resourceAccess += @{ id = $permIds.User_RevokeSessions_All;           type = 'Role' }

# one delegated scope
$resourceAccess += @{ id = $permIds.User_Read; type = 'Scope' }

# -- #2  Wrap in requiredResourceAccess  (single-element array) ---------------
$requiredResourceAccess = @(
    @{
        resourceAppId  = $graphSp.AppId
        resourceAccess = $resourceAccess
    }
)

# -- helper: fetch app by display name -----------------------------------------
function Get-AppByName ($name) {
    Get-MgApplication -Filter "displayName eq '$name'" -ConsistencyLevel eventual `
                      -Count c -All | Select-Object -First 1
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
        $missingItems = $resourceAccess | Where-Object { $existingIds -notcontains $_.id }
        $graphRRA.resourceAccess += $missingItems
    }
    else {
        $existingRRA += @{ resourceAppId = $graphSp.AppId; resourceAccess = $resourceAccess }
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

# fetch current role assignments for this application against Microsoft Graph
$existingRoles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All |
                 Where-Object { $_.ResourceId -eq $graphSp.Id }

# #1  application-role consent
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
    if ($existingRoles.AppRoleId -contains $roleId) {
        Write-Verbose "Role $roleId already present - skipping"
        continue
    }

    try {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $graphSp.Id `
            -PrincipalId        $sp.Id       `
            -ResourceId         $graphSp.Id  `
            -AppRoleId          $roleId | Out-Null
    } catch {
        if ($_.ErrorDetails.Message -match 'already exists') {
            Write-Verbose "Role $roleId already assigned (caught 400) - skipping"
        } else { throw }
    }
}

# #2  delegated scope (User.Read)
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

# -- create user with Global Reader role --------------------------------------
Write-Host "Creating service user with Global Reader permissions ..." -ForegroundColor Cyan

# Validate required permissions
$context = Get-MgContext
$requiredScopes = @('User.ReadWrite.All', 'RoleManagement.ReadWrite.Directory')
$missingScopes = @()

foreach ($scope in $requiredScopes) {
    if ($context.Scopes -notcontains $scope) {
        $missingScopes += $scope
    }
}

if ($missingScopes.Count -gt 0) {
    Write-Host "[WARN] Missing required permissions: $($missingScopes -join ', ')" -ForegroundColor Red
    Write-Host "   -> Please ensure these scopes are consented to in your app registration." -ForegroundColor Yellow
    Write-Host "   -> Continuing anyway, but operations may fail..." -ForegroundColor Yellow
}

# Generate secure random password
$passwordChars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#$%&*"
$password = -join ((1..16) | ForEach-Object { $passwordChars[(Get-Random -Maximum $passwordChars.Length)] })

# Use TenantId parameter for predictable domain creation
$context = Get-MgContext
$userName = "secto-service-reader"

# Use the TenantId parameter directly for better predictability
if ($TenantId) {
    $userDomain = $TenantId
    Write-Host "Using provided TenantId as domain: $userDomain"
} else {
    # If no TenantId parameter provided, use the authenticated tenant's ID
    $userDomain = $context.TenantId
    Write-Host "Using authenticated tenant ID as domain: $userDomain"
}

$userPrincipalName = "$userName@$userDomain"

# Create password profile (no forced password change on first login)
$passwordProfile = @{
    Password = $password
    ForceChangePasswordNextSignIn = $false
}

# -- user creation or retrieval -----------------------------------------------
$existingUser = Get-MgUser -Filter "userPrincipalName eq '$userPrincipalName'" -ConsistencyLevel eventual -Count c | Select-Object -First 1
if ($existingUser) {
    Write-Host "Service user '$userPrincipalName' already exists - skipping creation." -ForegroundColor Green
    $newUser = $existingUser
}
else {
    try {
        $newUser = New-MgUser -DisplayName "Secto Service Reader" `
            -PasswordProfile $passwordProfile `
            -AccountEnabled `
            -MailNickName $userName `
            -UserPrincipalName $userPrincipalName `
            -UsageLocation "US"
        Write-Host "[OK] Created user: $userPrincipalName" -ForegroundColor Green
        # Wait for replication
        Write-Host "Waiting for user replication (10 seconds)..." -ForegroundColor Yellow
        Start-Sleep -Seconds 10
    } catch {
        Write-Host "[WARN] Error creating user: $($_.Exception.Message)" -ForegroundColor Red
        $newUser = $null
    }
}

# -- ensure Global Reader role assignment -------------------------------------
if ($newUser) {
    $roleName = "Global Reader"
    $role = Get-MgDirectoryRole | Where-Object { $_.displayName -eq $roleName }
    if (-not $role) {
        Write-Host "Activating Global Reader role in tenant ..."
        $roleTemplate = Get-MgDirectoryRoleTemplate | Where-Object { $_.displayName -eq $roleName }
        if ($roleTemplate) { $role = New-MgDirectoryRole -DisplayName $roleName -RoleTemplateId $roleTemplate.Id }
    }

    $isMember = $false
    if ($role) {
        try {
            $isMember = (Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All | Where-Object { $_.Id -eq $newUser.Id }) -ne $null
        } catch { $isMember = $false }
    }

    if (-not $isMember) {
        Write-Host "Assigning Global Reader role to user ..."
        $newRoleMember = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)" }
        try {
            New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $newRoleMember
            Write-Host "[OK] Assigned Global Reader role to user" -ForegroundColor Green
        } catch {
            Write-Host "[WARN] Failed to assign Global Reader role: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "User already has Global Reader role – skipping assignment." -ForegroundColor Green
    }
}

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
if ($newUser) {
    Write-Host "### Service User Information ###" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " => Service User Principal Name.....: '$($newUser.UserPrincipalName)' <--- Global Reader Account" -ForegroundColor Yellow
    Write-Host " => Service User Password...........: '$password' <--- Ready to use immediately" -ForegroundColor Yellow
    Write-Host " => Assigned Role...................: 'Global Reader'" -ForegroundColor Yellow
    Write-Host ""
}
Write-Host ""

Disconnect-MgGraph
