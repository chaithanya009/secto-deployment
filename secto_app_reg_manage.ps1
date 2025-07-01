<#
  New-AppReg.ps1  ─  six permissions + auto-consent
#>

param(
    [string]   $DisplayName = 'secto',
    [string[]] $RedirectUris = @('https://localhost'),
    [ValidateSet('SingleTenant','MultiTenant','MultiTenantAndPersonal')]
    [string]   $Audience  = 'SingleTenant',
    [string]   $TenantId,
    [switch]   $CreateClientSecret
)

# ── map friendly audience names ───────────────────────────────────────────────
$audienceMap = @{
    SingleTenant            = 'AzureADMyOrg'
    MultiTenant             = 'AzureADMultipleOrgs'
    MultiTenantAndPersonal  = 'AzureADandPersonalMicrosoftAccount'
}

# ── install only required Microsoft.Graph sub-modules ────────────────────────
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
    Write-Host "✔ Required modules installed successfully!" -ForegroundColor Green
    Write-Host ""
}

# ── sign in with the five admin scopes (no explicit import needed) ──────────
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

# ── Graph resource service-principal (only once) ──────────────────────────────
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# ── permission IDs we need ────────────────────────────────────────────────────
$permIds = @{
    AuditLog_Read_All                 = 'b0afded3-3588-46d8-8b3d-9842eff778da'
    AuditLogsQuery_Read_All           = '5e1e9171-754d-478c-812c-f1755a9a4c2d'
    Directory_Read_All                = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'
    Domain_Read_All                   = 'dbb9058a-0e50-45d7-ae91-66909b5d4664'
    Organization_Read_All             = '498476ce-e0fe-48b0-b801-37ba7e2685c6'
    Policy_Read_All                   = '246dd0d5-5bd0-4def-940b-0421030a5b68'
    SharePointTenantSettings_Read_All = '83d4163d-a2d8-4d3b-9695-4ae3ca98f888'
    User_Read                         = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'
}

# ── 1️⃣  Build the resourceAccess array FIRST ─────────────
$resourceAccess = @()

# seven application roles
$resourceAccess += @{ id = $permIds.AuditLog_Read_All;                 type = 'Role' }
$resourceAccess += @{ id = $permIds.AuditLogsQuery_Read_All;           type = 'Role' }
$resourceAccess += @{ id = $permIds.Directory_Read_All;                type = 'Role' }
$resourceAccess += @{ id = $permIds.Domain_Read_All;                   type = 'Role' }
$resourceAccess += @{ id = $permIds.Organization_Read_All;             type = 'Role' }
$resourceAccess += @{ id = $permIds.Policy_Read_All;                   type = 'Role' }
$resourceAccess += @{ id = $permIds.SharePointTenantSettings_Read_All; type = 'Role' }

# one delegated scope
$resourceAccess += @{ id = $permIds.User_Read; type = 'Scope' }

# ── 2️⃣  Wrap in requiredResourceAccess  (single-element array) ───────────────
$requiredResourceAccess = @(
    @{
        resourceAppId  = $graphSp.AppId
        resourceAccess = $resourceAccess
    }
)

# ── helper: fetch app by display name ─────────────────────────────────────────
function Get-AppByName ($name) {
    Get-MgApplication -Filter "displayName eq '$name'" -ConsistencyLevel eventual `
                      -Count c -All | Select-Object -First 1
}

# ── create or update ──────────────────────────────────────────────────────────
$app = Get-AppByName $DisplayName
if ($null -eq $app) {
    Write-Host "Creating application '$DisplayName' …"
    $app = New-MgApplication `
              -DisplayName            $DisplayName `
              -SignInAudience         $audienceMap[$Audience] `
              -Web                    @{ RedirectUris = $RedirectUris } `
              -RequiredResourceAccess $requiredResourceAccess
    $sp  = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "✔ Created app : $($app.AppId)"
} else {
    Write-Host "Updating application '$DisplayName' …"
    $sp  = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"
    Update-MgApplication -ApplicationId $app.Id `
                         -Web                    @{ RedirectUris = $RedirectUris } `
                         -RequiredResourceAccess $requiredResourceAccess
}

# ── optional: client secret ───────────────────────────────────────────────────
if ($CreateClientSecret) {
    $secret = Add-MgApplicationPassword -ApplicationId $app.Id `
                 -PasswordCredential @{
                     displayName = 'automation-secret'
                     endDateTime = (Get-Date).AddYears(1)
                 }
}

# ── admin consent ────────────────────────────────────────────
Write-Host "`nGranting admin consent …"

# fetch current role assignments for this principal → Graph
# Note: Filter by principalId is not supported, so we get all and filter client-side
$existingRoles = Get-MgServicePrincipalAppRoleAssignment `
                   -ServicePrincipalId $graphSp.Id `
                   -All | Where-Object { $_.PrincipalId -eq $sp.Id }

# 1️⃣  application-role consent
foreach ($roleId in @(
        $permIds.AuditLog_Read_All,
        $permIds.AuditLogsQuery_Read_All,
        $permIds.Directory_Read_All,
        $permIds.Domain_Read_All,
        $permIds.Organization_Read_All,
        $permIds.Policy_Read_All,
        $permIds.SharePointTenantSettings_Read_All)
) {
    if ($existingRoles.AppRoleId -contains $roleId) {
        Write-Verbose "Role $roleId already present – skipping"
        continue
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $graphSp.Id `
        -PrincipalId        $sp.Id       `
        -ResourceId         $graphSp.Id  `
        -AppRoleId          $roleId | Out-Null
}

# 2️⃣  delegated scope (User.Read)
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
Write-Host "✔ Admin consent granted.`n" -ForegroundColor Green

# ── create user with Global Reader role ──────────────────────────────────────
Write-Host "Creating service user with Global Reader permissions …" -ForegroundColor Cyan

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
    Write-Host "⚠ Missing required permissions: $($missingScopes -join ', ')" -ForegroundColor Red
    Write-Host "   → Please ensure these scopes are consented to in your app registration." -ForegroundColor Yellow
    Write-Host "   → Continuing anyway, but operations may fail..." -ForegroundColor Yellow
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

# Create the user
try {
    $newUser = New-MgUser -DisplayName "Secto Service Reader" `
        -PasswordProfile $passwordProfile `
        -AccountEnabled `
        -MailNickName $userName `
        -UserPrincipalName $userPrincipalName `
        -UsageLocation "US"
    
    Write-Host "✔ Created user: $userPrincipalName" -ForegroundColor Green
    
    # Wait for user replication across Microsoft services
    Write-Host "Waiting for user replication (30 seconds)..." -ForegroundColor Yellow
    Start-Sleep -Seconds 30
    
    # Verify user exists before proceeding
    try {
        $verifyUser = Get-MgUser -UserId $newUser.Id -ErrorAction Stop
        Write-Host "✔ User verified in directory" -ForegroundColor Green
    } catch {
        Write-Host "⚠ User not yet replicated, waiting additional 15 seconds..." -ForegroundColor Yellow
        Start-Sleep -Seconds 15
        $verifyUser = Get-MgUser -UserId $newUser.Id
    }
    
    # Get Global Reader role (activate if needed)
    $roleName = "Global Reader"
    $role = Get-MgDirectoryRole | Where-Object {$_.displayName -eq $roleName}
    
    if ($null -eq $role) {
        Write-Host "Activating Global Reader role in tenant …"
        $roleTemplate = Get-MgDirectoryRoleTemplate | Where-Object {$_.displayName -eq $roleName}
        if ($roleTemplate) {
            $role = New-MgDirectoryRole -DisplayName $roleName -RoleTemplateId $roleTemplate.Id
        } else {
            throw "Global Reader role template not found"
        }
    }
    
    # Assign Global Reader role to the user using correct URI format
    try {
        $newRoleMember = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)"
        }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $newRoleMember
        Write-Host "✔ Assigned Global Reader role to user" -ForegroundColor Green
    } catch {
        Write-Host "⚠ Role assignment error: $($_.Exception.Message)" -ForegroundColor Yellow
        # Try alternative approach using direct REST API call
        try {
            $uri = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.Id)/members/`$ref"
            $body = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)"
            } | ConvertTo-Json
            
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json"
            Write-Host "✔ Assigned Global Reader role using REST API" -ForegroundColor Green
        } catch {
            Write-Host "⚠ Failed to assign role via REST API: $($_.Exception.Message)" -ForegroundColor Red
            throw "Could not assign Global Reader role to user"
        }
    }
}
catch {
    Write-Host "⚠ Error creating user or assigning role: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Message -like "*domain*") {
        Write-Host "   → Domain validation issue. Please ensure your tenant has verified domains." -ForegroundColor Yellow
        Write-Host "   → Run 'Get-MgDomain' to check available verified domains." -ForegroundColor Yellow
    } elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*forbidden*") {
        Write-Host "   → Permission issue. Ensure you have User.ReadWrite.All and RoleManagement.ReadWrite.Directory permissions." -ForegroundColor Yellow
    } elseif ($_.Exception.Message -like "*role*") {
        Write-Host "   → Role assignment issue. The user was created but role assignment failed." -ForegroundColor Yellow
        Write-Host "   → You can manually assign the Global Reader role in the Azure portal." -ForegroundColor Yellow
    }
    Write-Host "   → For troubleshooting, check: https://learn.microsoft.com/graph/errors" -ForegroundColor Yellow
    $newUser = $null
}

# ── display required information ──────────────────────────────────────────────
Write-Host "### Required Information for Secto Application ###" -ForegroundColor Cyan
Write-Host ""
$context = Get-MgContext
Write-Host " => Tenant ID.......................: '$($context.TenantId)' <---" -ForegroundColor Yellow
Write-Host " => Application Name................: '$DisplayName'" -ForegroundColor Yellow
Write-Host " => Client ID (App ID)..............: '$($app.AppId)' <---" -ForegroundColor Yellow
if ($CreateClientSecret) {
    Write-Host " => Client Secret...................: '$($secret.SecretText)' <---" -ForegroundColor Yellow
    Write-Host " => Secret Valid Until..............: '$($secret.EndDateTime)'" -ForegroundColor Yellow
}
Write-Host ""
if ($newUser) {
    Write-Host "### Service User Information ###" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " => Service User Display Name.......: '$($newUser.DisplayName)'" -ForegroundColor Yellow
    Write-Host " => Service User Principal Name.....: '$($newUser.UserPrincipalName)' <--- Global Reader Account" -ForegroundColor Yellow
    Write-Host " => Service User Password...........: '$password' <--- Ready to use immediately" -ForegroundColor Yellow
    Write-Host " => Service User ID.................: '$($newUser.Id)'" -ForegroundColor Yellow
    Write-Host " => Assigned Role...................: 'Global Reader'" -ForegroundColor Yellow
    Write-Host ""
}
Write-Host ""

Disconnect-MgGraph