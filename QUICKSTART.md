# Quick Start Guide

**Author**: Bjørnar Aassveen | **Blog**: https://aassveen.com | **Version**: 1.0 (2026-04-20)

---

Get the Teams Audit Report running in 5 minutes.

## 1. Create Service Principal (5 minutes)

### Using Azure CLI (Easiest)

```powershell
# Login to Azure
az login

# Create app and get ID
$appId = az ad app create --display-name "Teams Audit Reporter" --query appId -o tsv

# Create service principal
az ad sp create --id $appId

# Create secret (SAVE THIS!)
$secret = az ad app credential create --id $appId --display-name "TeamsAudit" --query password -o tsv

# Get tenant ID
$tenantId = az account show --query tenantId -o tsv

# Grant permissions
$graphApiId = "00000003-0000-0000-c000-000000000000"
$permissions = @("06da0dbc-49e7-4d38-b3e8-fe7b444e486f", "df021288-bdef-4463-88db-98f22db89214", "48638d0a-70d7-4b6e-8c28-c146acad650e", "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "5f8c59db-677d-491f-a6b3-d4ab5a87edda", "205e0cee-aea8-4d0f-a8f5-8cda821ae187", "810c84a8-4a9e-49e6-bf7d-12d183f40d01")
$permissions | ForEach-Object { az ad app permission add --id $appId --api $graphApiId --api-permissions "${_}=Role" }

# Grant admin consent
az ad app permission admin-consent --id $appId

echo "App ID: $appId"
echo "Secret: $secret"
echo "Tenant ID: $tenantId"
```

**Full instructions**: See `SETUP_SERVICE_PRINCIPAL.md`

## 2. Configure Script (2 minutes)

```powershell
# Copy template to config
Copy-Item config-template.json config.json

# Edit config.json with your values
notepad config.json
```

Replace these values:
```json
{
  "ClientId": "PASTE_YOUR_APP_ID",
  "ClientSecret": "PASTE_YOUR_SECRET",
  "TenantId": "PASTE_YOUR_TENANT_ID",
  ...
}
```

## 3. Run the Script (5-30 minutes depending on tenant size)

```powershell
.\Get-TeamsAuditReport.ps1
```

## 4. View Reports

Reports are generated in the `Reports` folder:
- **TeamsAuditReport_*.html** - Open in browser (prettiest!)
- **TeamsAuditReport_*.csv** - Open in Excel
- **TeamsAuditReport_*.json** - For automation/analysis

---

## Common Questions

**Q: The script is taking forever**

A: This is normal for large tenants (500+ teams). Typical time:
- 50 teams: 5-10 min
- 500 teams: 45-90 min  
- 2000 teams: 3-6 hours

Run it overnight or during off-peak hours.

**Q: "Authentication failed"**

A: Make sure `ClientId`, `ClientSecret`, and `TenantId` in config.json are correct. Verify the Service Principal has the required permissions.

**Q: What do the statuses mean?**

- **ACTIVE** (green): Team has activity in last 60 days
- **CONSIDER ARCHIVAL** (yellow): No activity for 60-90 days
- **RECOMMEND ARCHIVAL** (red): No activity for 90+ days

**Q: Can I change the thresholds?**

A: Yes! In `config.json`:
```json
{
  "InactiveThresholdDays": 90,    // Change to 180 for 6 months
  "ReviewThresholdDays": 60,      // Change to 120 for 4 months
  ...
}
```

**Q: How do I analyze the results?**

A: See `EXAMPLES.md` for PowerShell scripts to analyze the data.

---

**Author**: Bjørnar Aassveen  
**Blog**: https://aassveen.com  
**Version**: 1.0 (2026-04-20)

**Need detailed help?** Read the full `README.md`
