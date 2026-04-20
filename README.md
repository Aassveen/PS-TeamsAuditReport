# Teams Audit Report Generator

**Author**: Bjørnar Aassveen | **Blog**: https://aassveen.com | **Version**: 1.0 (2026-04-20)

---

PowerShell script that generates comprehensive audit reports of all Teams in your tenant with activity analysis and archival recommendations.

## Features

- **Teams Metadata Collection**: Lists all teams with owner, creator, member/guest counts, description, privacy settings, and archived status
- **Activity Detection**: Analyzes activity from:
  - Channel messages (last message timestamp)
  - SharePoint sites (last file modification)
  - Group mailbox (last email received/sent)
- **Intelligent Classification**: Recommends archival based on configurable thresholds:
  - **RECOMMEND ARCHIVAL**: No activity in last 90 days
  - **CONSIDER ARCHIVAL**: No activity in last 60-90 days
  - **ACTIVE**: Recent activity (< 60 days)
- **Multiple Output Formats**: Generates reports in CSV (Excel), JSON, and formatted HTML

## Prerequisites

1. **PowerShell 5.1+** (Windows PowerShell or PowerShell Core)
2. **Microsoft Graph API Access** via Service Principal
3. **Appropriate Azure AD Permissions** (see Setup section)

## Setup

### Step 1: Create Service Principal

You need to create an Azure AD App Registration with appropriate permissions to read Teams data.

```powershell
# Using Azure CLI
az ad app create --display-name "Teams Audit Reporter" `
  --query appId -o tsv

# Then create a service principal
az ad sp create --id <APP_ID>

# Create a client secret
az ad app credential create --id <APP_ID> `
  --display-name "TeamsAudit" `
  --query password -o tsv
```

Or manually via Azure Portal:
1. Go to **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Enter name: "Teams Audit Reporter"
4. Register the app
5. Go to **API permissions** and add:
   - `Directory.Read.All`
   - `Team.ReadBasic.All`
   - `TeamSettings.Read.All`
   - `TeamMember.Read.All` ← **REQUIRED** for member/guest counts
   - `User.Read.All`
   - `Group.Read.All`
   - `Sites.Read.All`
   - `Mail.Read`
6. Grant admin consent
7. Create a client secret under **Certificates & secrets**

### Step 2: Configure the Script

1. Copy `config-template.json` to `config.json`:
   ```powershell
   Copy-Item config-template.json config.json
   ```

2. Edit `config.json` with your values:
   ```json
   {
     "ClientId": "YOUR_APP_ID_HERE",
     "ClientSecret": "YOUR_CLIENT_SECRET_HERE",
     "TenantId": "YOUR_TENANT_ID_HERE",
     "OutputPath": "C:\\entra-terraform\\Powershell Script\\Reports",
     "InactiveThresholdDays": 90,
     "ReviewThresholdDays": 60,
     "ActivityLookbackDays": 90,
     "BatchSize": 50,
     "MaxRetries": 3,
     "RetryDelaySeconds": 2,
     "UseExchangeOnlineFallback": false,
     "Verbose": true
   }
   ```

## Usage

### Basic Execution

```powershell
# Using default config.json location
.\Get-TeamsAuditReport.ps1

# Using custom config path
.\Get-TeamsAuditReport.ps1 -ConfigPath "C:\path\to\config.json"
```

### Configuration Options

| Parameter | Description | Default |
|-----------|-------------|---------|
| `InactiveThresholdDays` | Days for "RECOMMEND ARCHIVAL" classification | 90 |
| `ReviewThresholdDays` | Days for "CONSIDER ARCHIVAL" classification | 60 |
| `ActivityLookbackDays` | How far back to check for activity | 90 |
| `BatchSize` | Graph API page size (higher = faster but more memory) | 50 |
| `MaxRetries` | Max API retry attempts on failure | 3 |
| `RetryDelaySeconds` | Initial delay between retries (exponential backoff) | 2 |
| `Verbose` | Enable verbose logging | true |

## Output Files

The script generates three reports in the configured output path:

1. **TeamsAuditReport_YYYYMMDD_HHMMSS.csv**
   - Tab-separated values
   - Opens easily in Excel
   - Columns: TeamName, Owner, CreatedBy, CreatedDateTime, MemberCount, GuestCount, LastActivityDates, RecommendedAction, etc.

2. **TeamsAuditReport_YYYYMMDD_HHMMSS.json**
   - Structured JSON array
   - Includes summary statistics
   - Suitable for automation and further processing

3. **TeamsAuditReport_YYYYMMDD_HHMMSS.html**
   - Formatted HTML report
   - Color-coded status (Green=Active, Yellow=Consider, Red=Recommend)
   - Summary statistics dashboard
   - Sortable tables

## Performance Considerations

### Execution Time

Expected execution times for different tenant sizes (estimated):
- **50 Teams**: 5-10 minutes
- **500 Teams**: 45-90 minutes
- **1000 Teams**: 90-180 minutes
- **2000 Teams**: 180-360 minutes

Times depend on:
- Activity level (more channels = longer)
- Network latency
- Microsoft Graph API throttling
- Retry rate due to throttling

### Optimization Tips

1. **Batch Size**: Increase `BatchSize` to 100 for faster pagination (uses more memory)
2. **Retry Delays**: Reduce `RetryDelaySeconds` if you know your tenant has good API quota
3. **Activity Lookback**: Reduce `ActivityLookbackDays` if you only care about recent activity
4. **Schedule on Off-Hours**: Run during tenant off-hours to reduce API contention

### API Throttling

Microsoft Graph API has rate limits (~1000 requests per minute per tenant). The script implements:
- Exponential backoff retry logic
- Batch pagination to minimize requests
- Error handling for 429 (Too Many Requests) responses

If you hit throttling:
1. Check the console output for "API call failed" messages
2. Reduce `BatchSize` to 25-30
3. Increase `RetryDelaySeconds` to 5-10
4. Consider running during tenant off-hours

## Troubleshooting

### "Missing role permissions" or "Forbidden" errors
- **IMPORTANT**: Verify **all 8 API permissions** are added (especially `TeamMember.Read.All`)
- Grant **admin consent** for all permissions
- Wait 10-15 minutes for permissions to propagate
- Try using the updated version: `Get-TeamsAuditReportv01.ps1`

### "API is not supported" errors
- This usually means the endpoint format is wrong
- Use `Get-TeamsAuditReportv01.ps1` which corrects API endpoint issues

### "Authentication failed"
- Verify `ClientId`, `ClientSecret`, and `TenantId` are correct
- Ensure the Service Principal has required API permissions
- Check if the client secret has expired

### "API call failed after X retries"
- Your tenant is hitting Graph API throttling
- Reduce `BatchSize` or increase `RetryDelaySeconds`
- Try running during off-hours
- Check Microsoft's current API health status

### Missing Activity Data
- Some teams may not have channels with messages
- Some teams may not have associated SharePoint sites
- Missing data is intentional and reflects actual inactivity
- The "LastActivityDate" uses MAX of all sources, so one inactive source doesn't disqualify the team

### Empty SharePoint/Mailbox Activity
- Teams without file changes or email traffic will show null
- This is expected and factored into classification
- Teams are still classified correctly based on ANY available activity

## Data Privacy

This script:
- Only reads data (no modification or deletion)
- Does not store authentication tokens after execution
- Generates reports in your specified output path only
- Does not transmit data to external services

## API Permissions Reference

The Service Principal requires these Microsoft Graph API permissions:

| Permission | Type | Purpose |
|-----------|------|---------|
| `Directory.Read.All` | Application | Read directory structure |
| `Team.ReadBasic.All` | Application | List teams and basic info |
| `TeamSettings.Read.All` | Application | Read team settings |
| `User.Read.All` | Application | Read user information |
| `Group.Read.All` | Application | Read group properties |
| `TeamMember.Read.All` | Application | **Read team/group members (REQUIRED)** |
| `Sites.Read.All` | Application | Access associated SharePoint sites |
| `Mail.Read` | Application | Read group mailbox conversations |

## Examples

### Example 1: Generate Report for Specific Teams Only

Edit config.json to change `ReviewThresholdDays` and `InactiveThresholdDays`:

```json
{
  "InactiveThresholdDays": 180,
  "ReviewThresholdDays": 120,
  "ActivityLookbackDays": 180
}
```

This will recommend archival for teams inactive > 180 days.

### Example 2: Schedule as Recurring Report

Create a Windows Task Scheduler task:

```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass `
  -File 'C:\entra-terraform\Powershell Script\Get-TeamsAuditReport.ps1'"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 3:00AM
Register-ScheduledTask -Action $action -Trigger $trigger `
  -TaskName "Teams Audit Report" -Description "Weekly Teams audit report"
```

### Example 3: Filter Results Using JSON Output

```powershell
# Get all teams recommended for archival
$report = Get-Content 'Reports\TeamsAuditReport_*.json' | ConvertFrom-Json
$report.Teams | Where-Object { $_.RecommendedAction -eq 'RECOMMEND ARCHIVAL' } | 
  Select-Object TeamName, LastActivityDate, DaysSinceActivity |
  Format-Table
```

## Known Limitations

1. **Private Channel Messages**: Private channel message queries have stricter permissions and may require additional scopes
2. **Teams with No Channels**: Some teams may be created but empty - they'll show creation date as last activity
3. **Activity Granularity**: SharePoint and mailbox activity checked only within the lookback period (max 90 days by default)
4. **File-only Teams**: Teams with files but no messages will still be correctly classified based on file modification dates

## Support & Contributing

For issues or improvements:
1. Check the troubleshooting section above
2. Review the configuration options
3. Check Microsoft Graph API documentation: https://docs.microsoft.com/graph

## License

This script is provided as-is for tenant administration purposes.

## Version History

- **1.0-FIXED** (2026-04-20): Fixed version with corrected API endpoints and all required permissions documented
  - Fixed `/teams/{id}/members` → `/groups/{id}/members`
  - Fixed mailbox and SharePoint activity detection
  - Added `TeamMember.Read.All` permission requirement
  - Better error handling for unsupported endpoints
  - Recommended: Use this version for most reliable operation

- **1.0** (2026-04-20): Initial release with CSV, JSON, and HTML export

---

**Author**: Bjørnar Aassveen  
**Blog**: https://aassveen.com  
**Version**: 1.0 (2026-04-20)

**Questions?** Review the configuration section and troubleshooting guide above.
