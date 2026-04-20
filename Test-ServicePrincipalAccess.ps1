<#
.SYNOPSIS
    Test Service Principal access to Microsoft Graph Teams API
    
.AUTHOR
    Bjørnar Aassveen
    Blog: https://aassveen.com
    Version: 1.0 (2026-04-20)
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Service Principal Access Diagnostic" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Load configuration
Write-Host "Loading configuration..."
$config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
Write-Host "✓ Configuration loaded`n"

# Step 1: Authenticate
Write-Host "STEP 1: Authenticating to Microsoft Graph..." -ForegroundColor Yellow

$TokenEndpoint = "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token"
$Scope = "https://graph.microsoft.com/.default"

$Body = @{
    grant_type    = "client_credentials"
    client_id     = $config.ClientId
    client_secret = $config.ClientSecret
    scope         = $Scope
}

try {
    $response = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -Body $Body -ContentType "application/x-www-form-urlencoded"
    $token = $response.access_token
    Write-Host "✓ Authentication successful`n" -ForegroundColor Green
}
catch {
    Write-Host "✗ Authentication failed: $_`n" -ForegroundColor Red
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

# Step 2: Test /teams endpoint
Write-Host "STEP 2: Testing /teams endpoint..." -ForegroundColor Yellow

try {
    $uri = "https://graph.microsoft.com/v1.0/teams?`$top=10"
    Write-Host "  URL: $uri"
    
    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
    
    Write-Host "  ✓ Endpoint accessible" -ForegroundColor Green
    Write-Host "  Teams found: $($response.value.Count)" -ForegroundColor Green
    
    if ($response.value.Count -gt 0) {
        Write-Host "`n  First few teams:"
        $response.value | Select-Object -First 3 | ForEach-Object {
            Write-Host "    - $($_.displayName)"
        }
    }
    else {
        Write-Host "  ⚠ No teams returned (tenant may not have any Teams created)" -ForegroundColor Yellow
    }
    Write-Host ""
}
catch {
    Write-Host "  ✗ Failed to access /teams: $_`n" -ForegroundColor Red
    Write-Host "  This suggests Team.ReadBasic.All permission is missing or not granted`n" -ForegroundColor Yellow
}

# Step 3: Test /groups endpoint
Write-Host "STEP 3: Testing /groups endpoint (alternative)..." -ForegroundColor Yellow

try {
    $uri = "https://graph.microsoft.com/v1.0/groups?`$top=10&`$filter=resourceProvisioningOptions/any(x:x eq 'Team')"
    Write-Host "  URL: $uri"
    
    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
    
    Write-Host "  ✓ Endpoint accessible" -ForegroundColor Green
    Write-Host "  Team groups found: $($response.value.Count)" -ForegroundColor Green
    
    if ($response.value.Count -gt 0) {
        Write-Host "`n  First few team-backed groups:"
        $response.value | Select-Object -First 3 | ForEach-Object {
            Write-Host "    - $($_.displayName)"
        }
    }
    else {
        Write-Host "  ⚠ No team-backed groups returned" -ForegroundColor Yellow
    }
    Write-Host ""
}
catch {
    Write-Host "  ✗ Failed to access /groups: $_`n" -ForegroundColor Red
}

# Step 4: Test service principal info
Write-Host "STEP 4: Checking Service Principal details..." -ForegroundColor Yellow

try {
    $uri = "https://graph.microsoft.com/v1.0/serviceprincipals?`$filter=appId eq '$($config.ClientId)'&`$select=id,displayName,servicePrincipalNames"
    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
    
    if ($response.value.Count -gt 0) {
        $sp = $response.value[0]
        Write-Host "  ✓ Service Principal found" -ForegroundColor Green
        Write-Host "  Display Name: $($sp.displayName)"
        Write-Host "  Service Principal Names: $($sp.servicePrincipalNames -join ', ')"
    }
    else {
        Write-Host "  ✗ Service Principal not found" -ForegroundColor Red
    }
    Write-Host ""
}
catch {
    Write-Host "  ⚠ Could not verify Service Principal: $_`n" -ForegroundColor Yellow
}

# Step 5: Test app registration permissions
Write-Host "STEP 5: Testing Microsoft Graph access..." -ForegroundColor Yellow

$testEndpoints = @(
    @{ Name = "User Directory"; Uri = "https://graph.microsoft.com/v1.0/me"; Permission = "User.Read" },
    @{ Name = "Teams"; Uri = "https://graph.microsoft.com/v1.0/teams?`$top=1"; Permission = "Team.ReadBasic.All" },
    @{ Name = "Groups"; Uri = "https://graph.microsoft.com/v1.0/groups?`$top=1"; Permission = "Group.Read.All" },
    @{ Name = "Users"; Uri = "https://graph.microsoft.com/v1.0/users?`$top=1"; Permission = "User.Read.All" },
    @{ Name = "Sites"; Uri = "https://graph.microsoft.com/v1.0/sites?`$top=1"; Permission = "Sites.Read.All" }
)

foreach ($endpoint in $testEndpoints) {
    try {
        $response = Invoke-RestMethod -Method Get -Uri $endpoint.Uri -Headers $headers -ErrorAction Stop
        Write-Host "  ✓ $($endpoint.Name): Accessible" -ForegroundColor Green
    }
    catch {
        $errorCode = $_.Exception.Response.StatusCode
        if ($errorCode -eq "Forbidden") {
            Write-Host "  ✗ $($endpoint.Name): Permission Denied (needs: $($endpoint.Permission))" -ForegroundColor Red
        }
        else {
            Write-Host "  ✗ $($endpoint.Name): $errorCode" -ForegroundColor Red
        }
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Diagnostic Complete" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "RECOMMENDATIONS:" -ForegroundColor Yellow
Write-Host "1. If /teams returned 0: Your tenant may not have any Teams created"
Write-Host "2. If /teams failed: Ensure Team.ReadBasic.All permission is granted and admin consent given"
Write-Host "3. If permissions show as denied: Wait 10-15 minutes and try again"
Write-Host "4. If everything passes but no teams: Create a test Team and run again`n"
