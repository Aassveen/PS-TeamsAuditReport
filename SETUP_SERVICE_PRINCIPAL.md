# Setup Guide: Creating Service Principal for Teams Audit Script

**Author**: Bjørnar Aassveen | **Blog**: https://aassveen.com | **Version**: 1.0 (2026-04-20)

---

This guide walks you through creating the Azure AD Service Principal required to run the Teams Audit Report script.

## Prerequisites

- Global Administrator or Application Administrator role in Azure AD
- Either:
  - Azure CLI installed (`az login` ready)
  - OR access to Azure Portal
  - OR PowerShell with Azure AD module

## Option 1: Using Azure CLI (Recommended - Fastest)

### Step 1: Authenticate to Azure

```powershell
az login
az account set --subscription "YOUR_SUBSCRIPTION_ID"  # if multiple subscriptions
```

### Step 2: Create App Registration

```powershell
$appName = "Teams Audit Reporter"
$appId = az ad app create --display-name $appName --query appId -o tsv
Write-Host "App ID: $appId"
```

### Step 3: Create Service Principal

```powershell
az ad sp create --id $appId
Write-Host "Service Principal created"
```

### Step 4: Create Client Secret

```powershell
$secret = az ad app credential create --id $appId `
  --display-name "TeamsAudit-Secret" `
  --years 2 `
  --query password -o tsv

Write-Host "Client Secret: $secret"
Write-Host "SAVE THIS SECRET SECURELY - IT CANNOT BE RETRIEVED LATER"
```

### Step 5: Get Tenant ID

```powershell
$tenantId = az account show --query tenantId -o tsv
Write-Host "Tenant ID: $tenantId"
```

### Step 6: Grant API Permissions

```powershell
$graphApiId = "00000003-0000-0000-c000-000000000000"

# Permission IDs for Microsoft Graph
$permissions = @(
    "06da0dbc-49e7-4d38-b3e8-fe7b444e486f",  # Directory.Read.All
    "df021288-bdef-4463-88db-98f22db89214",  # Team.ReadBasic.All
    "48638d0a-70d7-4b6e-8c28-c146acad650e",  # TeamSettings.Read.All
    "e1fe6dd8-ba31-4d61-89e7-88639da4683d",  # User.Read.All
    "5f8c59db-677d-491f-a6b3-d4ab5a87edda",  # Group.Read.All
    "a3371d5f-7653-4997-9266-a89cb23adc3b",  # TeamMember.Read.All (REQUIRED for member counts!)
    "205e0cee-aea8-4d0f-a8f5-8cda821ae187",  # Sites.Read.All
    "810c84a8-4a9e-49e6-bf7d-12d183f40d01"   # Mail.Read
)

foreach ($permId in $permissions) {
    az ad app permission add --id $appId `
      --api $graphApiId `
      --api-permissions "${permId}=Role"
}

Write-Host "API Permissions added (still need admin consent)"
```

### Step 7: Grant Admin Consent

```powershell
# Grant admin consent (requires Global Admin)
az ad app permission admin-consent --id $appId
Write-Host "Admin consent granted"
```

## Option 2: Using Azure Portal (More Manual)

### Step 1: Create App Registration

1. Sign in to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **+ New registration**
4. Enter:
   - Name: `Teams Audit Reporter`
   - Supported account types: `Accounts in this organizational directory only`
5. Click **Register**

### Step 2: Note Down Key Values

From the Overview page, copy and save:
- **Application (client) ID** → This is your `ClientId`
- **Directory (tenant) ID** → This is your `TenantId`

### Step 3: Create Client Secret

1. In the app registration, go to **Certificates & secrets**
2. Under **Client secrets**, click **+ New client secret**
3. Enter:
   - Description: `TeamsAudit-Secret`
   - Expires: `24 months`
4. Click **Add**
5. **Immediately copy the Value** (shown once, cannot be retrieved later) → This is your `ClientSecret`

### Step 4: Grant API Permissions

1. In the app registration, go to **API permissions**
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Click **Application permissions**
5. Search for and select each of these permissions:
   - `Directory.Read.All`
   - `Group.Read.All`
   - `Mail.Read`
   - `Sites.Read.All`
   - `Team.ReadBasic.All`
   - `TeamMember.Read.All` ← **REQUIRED for member/guest counts**
   - `TeamSettings.Read.All`
   - `User.Read.All`
6. Click **Add permissions**

### Step 5: Grant Admin Consent

1. Back in **API permissions**
2. Click **Grant admin consent for [Your Tenant]**
3. Click **Yes** to confirm

## Option 3: Using PowerShell with Azure AD Module

```powershell
# Install module if not already installed
Install-Module AzureAD -Force

# Connect to Azure AD
Connect-AzureAD

# Create App Registration
$app = New-AzureADApplication -DisplayName "Teams Audit Reporter"

# Create Service Principal
New-AzureADServicePrincipal -AppId $app.AppId

# Create Client Secret
$passwordCredential = New-AzureADApplicationPasswordCredential `
  -ObjectId $app.ObjectId -EndDate (Get-Date).AddYears(2)

# Output values
Write-Host "Client ID: $($app.AppId)"
Write-Host "Client Secret: $($passwordCredential.Value)"
Write-Host "Tenant ID: $(Get-AzureADTenantDetail | Select-Object -ExpandProperty ObjectId)"

# Note: For API permissions, you still need to use Portal or az cli
# as PowerShell doesn't have a straightforward way to grant permissions
```

## Verification: Test the Service Principal

Once you've created the Service Principal, test it with this PowerShell script:

```powershell
param(
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$TenantId
)

$tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

$body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
}

try {
    $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body `
      -ContentType "application/x-www-form-urlencoded"
    
    $token = $response.access_token
    
    # Test by listing teams
    $headers = @{
        "Authorization" = "Bearer $token"
    }
    
    $teams = Invoke-RestMethod -Method Get `
      -Uri "https://graph.microsoft.com/v1.0/teams?`$top=1" `
      -Headers $headers
    
    Write-Host "✓ Authentication successful!"
    Write-Host "✓ Can access Teams API!"
    Write-Host "Found $($teams.value.Count) teams (showing top 1 of total)"
    return $true
}
catch {
    Write-Host "✗ Authentication or API access failed:"
    Write-Host $_.Exception.Message
    return $false
}
```

Save this as `Test-ServicePrincipal.ps1` and run:

```powershell
.\Test-ServicePrincipal.ps1 -ClientId "YOUR_CLIENT_ID" `
  -ClientSecret "YOUR_CLIENT_SECRET" -TenantId "YOUR_TENANT_ID"
```

## Securely Store Credentials

### Option A: Use config.json (NOT recommended for production)

Edit `config.json` with your credentials:

```json
{
  "ClientId": "12345678-1234-1234-1234-123456789012",
  "ClientSecret": "Your_Secret_Here",
  "TenantId": "12345678-1234-1234-1234-123456789012",
  ...
}
```

**Security Note**: This file contains sensitive credentials. Ensure it's:
- Not checked into version control (add to `.gitignore`)
- Protected with NTFS permissions (Administrators only)
- Not shared or displayed

### Option B: Use Environment Variables (Better)

Store credentials in environment variables and modify config.json:

```powershell
# Set environment variables (Windows)
[System.Environment]::SetEnvironmentVariable("TEAMS_AUDIT_CLIENT_ID", "YOUR_CLIENT_ID", "User")
[System.Environment]::SetEnvironmentVariable("TEAMS_AUDIT_CLIENT_SECRET", "YOUR_CLIENT_SECRET", "User")
[System.Environment]::SetEnvironmentVariable("TEAMS_AUDIT_TENANT_ID", "YOUR_TENANT_ID", "User")
```

Then modify `config.json`:

```json
{
  "ClientId": "{ENV:TEAMS_AUDIT_CLIENT_ID}",
  "ClientSecret": "{ENV:TEAMS_AUDIT_CLIENT_SECRET}",
  "TenantId": "{ENV:TEAMS_AUDIT_TENANT_ID}",
  ...
}
```

And update the script to read from environment:

```powershell
if ($config.ClientId.StartsWith("{ENV:")) {
    $envVar = $config.ClientId.Replace("{ENV:", "").Replace("}", "")
    $config.ClientId = [System.Environment]::GetEnvironmentVariable($envVar, "User")
}
# ... repeat for ClientSecret and TenantId
```

### Option C: Use Azure Key Vault (Best for Production)

Store secrets in Azure Key Vault:

```powershell
# Create Key Vault (if not exists)
$kvName = "teamsauditkv"
az keyvault create --name $kvName --resource-group "YOUR_RG"

# Store secrets
az keyvault secret set --vault-name $kvName `
  --name "teams-audit-client-id" --value "YOUR_CLIENT_ID"
az keyvault secret set --vault-name $kvName `
  --name "teams-audit-client-secret" --value "YOUR_CLIENT_SECRET"
az keyvault secret set --vault-name $kvName `
  --name "teams-audit-tenant-id" --value "YOUR_TENANT_ID"
```

Then retrieve in script:

```powershell
$kvName = "teamsauditkv"
$clientId = az keyvault secret show --vault-name $kvName `
  --name "teams-audit-client-id" --query value -o tsv
# ... retrieve others similarly
```

## Troubleshooting

### "Permission denied" when granting consent

**Solution**: You need Global Administrator or Application Administrator role. Contact your Azure AD administrator.

### "Invalid client secret"

**Solution**: Client secrets expire and cannot be retrieved after creation:
1. Generate a new secret in the app registration
2. Update config.json
3. Delete the old secret

### "Insufficient privileges to complete the operation"

**Solution**: The Service Principal doesn't have required permissions:
1. Verify all 7 permissions are added in the app registration
2. Verify admin consent has been granted (green checkmark visible)
3. Wait 5-10 minutes for permissions to propagate
4. Sign out and back in to portal to refresh token

### "Teams API returns empty list"

**Solution**: Typically means either:
1. No teams exist in the tenant (expected)
2. The Service Principal doesn't have `Team.ReadBasic.All` permission
3. The permission hasn't fully propagated (wait 10+ minutes after granting)

## Next Steps

1. Update your `config.json` with the credentials you created:
   ```json
   {
     "ClientId": "YOUR_APP_ID",
     "ClientSecret": "YOUR_CLIENT_SECRET",
     "TenantId": "YOUR_TENANT_ID",
     ...
   }
   ```

2. Run the test script to verify: `.\Test-ServicePrincipal.ps1`

3. Execute the audit report: `.\Get-TeamsAuditReport.ps1`

## Cleanup (Optional)

If you need to delete the app registration later:

```powershell
# Using Azure CLI
az ad app delete --id "YOUR_CLIENT_ID"

# Using Portal
# Azure AD → App registrations → Find app → Delete
```

---

**Author**: Bjørnar Aassveen  
**Blog**: https://aassveen.com  
**Version**: 1.0 (2026-04-20)

**Need help?** Review the API permissions reference in the main README.md or check Microsoft's [Graph API documentation](https://docs.microsoft.com/graph).
