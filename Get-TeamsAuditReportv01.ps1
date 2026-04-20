<#
.SYNOPSIS
    Generates a comprehensive audit report of all Teams in a tenant with activity analysis and archival recommendations.

.DESCRIPTION
    This script:
    - Lists all Teams with metadata (owner, creator, member/guest counts, description, privacy, archived status)
    - Detects activity from: Channel messages, SharePoint files, and Group mailbox
    - Classifies Teams based on inactivity thresholds (90 days = recommend archival, 60 days = consider archival, < 60 = active)
    - Generates reports in CSV, JSON, and HTML formats

.PARAMETER ConfigPath
    Path to the configuration JSON file. Default: config.json in the script directory.

.EXAMPLE
    .\Get-TeamsAuditReport.ps1 -ConfigPath "C:\path\to\config.json"

.NOTES
    Requires Microsoft Graph API access via Service Principal (Client Credentials flow)
    Required API Permissions:
    - Directory.Read.All
    - Team.ReadBasic.All
    - TeamSettings.Read.All
    - TeamMember.Read.All
    - User.Read.All
    - Group.Read.All
    - Sites.Read.All
    - Mail.Read

.AUTHOR
    Bjørnar Aassveen
    Blog: https://aassveen.com
    Version: 1.0-FIXED (2026-04-20)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

#region Classes & Types
class TeamActivity {
    [string]$TeamId
    [string]$TeamName
    [string]$Owner
    [string]$CreatedBy
    [datetime]$CreatedDateTime
    [int]$MemberCount
    [int]$GuestCount
    [string]$Description
    [string]$Privacy
    [bool]$IsArchived
    [string]$SensitivityLabel
    [datetime]$LastChannelMessageDate = [DateTime]::MinValue
    [datetime]$LastSharePointActivityDate = [DateTime]::MinValue
    [datetime]$LastMailboxActivityDate = [DateTime]::MinValue
    [datetime]$LastActivityDate = [DateTime]::MinValue
    [int]$DaysSinceActivity
    [string]$RecommendedAction
    [string]$Details
}
#endregion

#region Script Configuration
$ErrorActionPreference = "Continue"
$VerbosePreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

# Load configuration
function Load-Configuration {
    param([string]$Path)
    
    Write-Host "Loading configuration from: $Path"
    
    if (-not (Test-Path $Path)) {
        Write-Error "Configuration file not found: $Path"
        return $null
    }
    
    try {
        $config = Get-Content -Path $Path -Raw | ConvertFrom-Json
        Write-Host "✓ Configuration loaded successfully" -ForegroundColor Green
        return $config
    }
    catch {
        Write-Error "Failed to load configuration: $_"
        return $null
    }
}

$script:Config = Load-Configuration -Path $ConfigPath
if (-not $script:Config) {
    exit 1
}

# Validate required config values
$requiredFields = @("ClientId", "ClientSecret", "TenantId", "OutputPath")
foreach ($field in $requiredFields) {
    if (-not $script:Config.$field) {
        Write-Error "Missing required configuration field: $field"
        exit 1
    }
}

# Ensure output path exists
if (-not (Test-Path $script:Config.OutputPath)) {
    New-Item -ItemType Directory -Path $script:Config.OutputPath -Force | Out-Null
}

$script:StartTime = Get-Date
$script:ReportDateTime = $script:StartTime.ToString("yyyyMMdd_HHmmss")
#endregion

#region Authentication & Graph API Functions
function Connect-ToGraph {
    param()
    
    Write-Host "Authenticating to Microsoft Graph as Service Principal..."
    
    $TokenEndpoint = "https://login.microsoftonline.com/$($script:Config.TenantId)/oauth2/v2.0/token"
    $Scope = "https://graph.microsoft.com/.default"
    
    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $script:Config.ClientId
        client_secret = $script:Config.ClientSecret
        scope         = $Scope
    }
    
    try {
        $response = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -Body $Body -ContentType "application/x-www-form-urlencoded"
        $script:GraphToken = $response.access_token
        $script:TokenExpiry = (Get-Date).AddSeconds($response.expires_in - 300)
        Write-Host "✓ Successfully authenticated to Microsoft Graph" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Authentication failed: $_"
        return $false
    }
}

function Invoke-GraphAPI {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Body,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 0
    )
    
    # Refresh token if expiring
    if ((Get-Date) -gt $script:TokenExpiry) {
        Connect-ToGraph | Out-Null
    }
    
    $headers = @{
        "Authorization" = "Bearer $($script:GraphToken)"
        "Content-Type"  = "application/json"
    }
    
    try {
        $params = @{
            Method      = $Method
            Uri         = $Uri
            Headers     = $headers
            ErrorAction = "Stop"
        }
        
        if ($Body) {
            $params["Body"] = $Body | ConvertTo-Json -Depth 10
        }
        
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        if ($RetryCount -lt $script:Config.MaxRetries) {
            $delay = [Math]::Min($script:Config.RetryDelaySeconds * [Math]::Pow(2, $RetryCount), 60)
            Write-Verbose "API call failed, retrying in ${delay}s... (Attempt $($RetryCount + 1)/$($script:Config.MaxRetries)): $($_.Exception.Message)"
            Start-Sleep -Seconds $delay
            return Invoke-GraphAPI -Method $Method -Uri $Uri -Body $Body -RetryCount ($RetryCount + 1)
        }
        else {
            Write-Verbose "API call failed after $($script:Config.MaxRetries) retries: $_"
            return $null
        }
    }
}

function Get-GraphAPIPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 50
    )
    
    $results = @()
    
    # Build URI with $top if not already present and if the endpoint supports it
    if ($Uri -notmatch '\$top=') {
        if ($Uri -match '\?') {
            $pageUri = "$Uri&`$top=$PageSize"
        } else {
            $pageUri = "$Uri`?`$top=$PageSize"
        }
    }
    else {
        $pageUri = $Uri
    }
    
    while ($pageUri) {
        $response = Invoke-GraphAPI -Method "GET" -Uri $pageUri
        
        if ($response) {
            if ($response.value) {
                $results += $response.value
            }
            elseif ($response -is [array]) {
                $results += $response
            }
            else {
                $results += $response
            }
            
            $pageUri = $response.'@odata.nextLink'
        }
        else {
            break
        }
    }
    
    return $results
}
#endregion

#region Teams Metadata Collection (Phase 2)
function Get-AllTeams {
    param()
    
    Write-Host "Fetching all Teams from tenant..."
    
    # Note: /teams endpoint doesn't support $top parameter
    # It returns all teams (up to 1000) or filtered results
    $uri = "https://graph.microsoft.com/v1.0/teams"
    
    try {
        $response = Invoke-GraphAPI -Method "GET" -Uri $uri
        
        if ($response -and $response.value) {
            $teams = $response.value
        }
        elseif ($response -is [array]) {
            $teams = $response
        }
        else {
            $teams = @()
        }
        
        Write-Host "✓ Found $($teams.Count) teams" -ForegroundColor Green
        return $teams
    }
    catch {
        Write-Host "✗ Failed to fetch teams: $_" -ForegroundColor Red
        return @()
    }
}

function Get-TeamOwners {
    param([string]$GroupId)
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/owners"
        $response = Invoke-GraphAPI -Method "GET" -Uri $uri
        
        if ($response.value) {
            return ($response.value | ForEach-Object { $_.displayName }) -join ", "
        }
        return "N/A"
    }
    catch {
        Write-Verbose "Failed to get owners for group $GroupId : $_"
        return "N/A"
    }
}

function Get-GroupInfo {
    param([string]$GroupId)
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId"
        $group = Invoke-GraphAPI -Method "GET" -Uri $uri
        
        return @{
            CreatedBy    = $group.createdDateTime
            Description  = $group.description
            DisplayName  = $group.displayName
            CreatedByUPN = if ($group.createdOnBehalfOf) { $group.createdOnBehalfOf.userPrincipalName } else { "System" }
        }
    }
    catch {
        Write-Verbose "Failed to get group info for $GroupId : $_"
        return $null
    }
}

function Get-TeamMembers {
    param([string]$GroupId)
    
    try {
        # Use /groups/{id}/members endpoint which properly supports paging
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members"
        $members = Get-GraphAPIPage -Uri $uri -PageSize 999
        
        $memberCount = 0
        $guestCount = 0
        
        foreach ($member in $members) {
            if ($member.userType -eq "Guest") {
                $guestCount++
            }
            else {
                $memberCount++
            }
        }
        
        return @{
            MemberCount = $memberCount
            GuestCount  = $guestCount
        }
    }
    catch {
        Write-Verbose "Failed to get members for group $GroupId : $_"
        return @{
            MemberCount = 0
            GuestCount  = 0
        }
    }
}

function Get-TeamChannels {
    param([string]$TeamId)
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels"
        $channels = Get-GraphAPIPage -Uri $uri -PageSize 50
        return $channels
    }
    catch {
        Write-Verbose "Failed to get channels for team $TeamId : $_"
        return @()
    }
}
#endregion

#region Activity Detection (Phases 3-5)
function Get-LastChannelMessageDate {
    param([string]$TeamId)
    
    try {
        $channels = Get-TeamChannels -TeamId $TeamId
        $lastMessageDate = $null
        
        foreach ($channel in $channels) {
            try {
                # Get messages directly without $top first, then parse
                $uri = "https://graph.microsoft.com/v1.0/teams/$TeamId/channels/$($channel.id)/messages?`$orderby=createdDateTime desc"
                $response = Invoke-GraphAPI -Method "GET" -Uri $uri
                
                if ($response.value -and $response.value.Count -gt 0) {
                    $messageDate = [datetime]$response.value[0].createdDateTime
                    if ($null -eq $lastMessageDate -or $messageDate -gt $lastMessageDate) {
                        $lastMessageDate = $messageDate
                    }
                }
            }
            catch {
                # Channel may not have messages or may be private - continue
                Write-Verbose "Could not access messages in channel $($channel.displayName) : $_"
            }
        }
        
        return $lastMessageDate ?? [DateTime]::MinValue
    }
    catch {
        Write-Verbose "Failed to get channel activity for team $TeamId : $_"
        return [DateTime]::MinValue
    }
}

function Get-LastSharePointActivityDate {
    param([string]$GroupId)
    
    try {
        # Get SharePoint site associated with the group
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/sites/root"
        $site = Invoke-GraphAPI -Method "GET" -Uri $uri
        
        if ($site -and $site.id) {
            # Get recent files on the SharePoint site
            $lookbackDate = (Get-Date).AddDays(-$script:Config.ActivityLookbackDays).ToString("o")
            $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive/root/children?`$filter=lastModifiedDateTime ge $lookbackDate&`$orderby=lastModifiedDateTime desc"
            $response = Invoke-GraphAPI -Method "GET" -Uri $uri
            
            if ($response.value -and $response.value.Count -gt 0 -and $response.value[0].lastModifiedDateTime) {
                return [datetime]$response.value[0].lastModifiedDateTime
            }
        }
        
        return [DateTime]::MinValue
    }
    catch {
        Write-Verbose "Failed to get SharePoint activity for group $GroupId : $_"
        return [DateTime]::MinValue
    }
}

function Get-LastMailboxActivityDate {
    param([string]$GroupId)
    
    try {
        # Get conversations from group mailbox
        $lookbackDate = (Get-Date).AddDays(-$script:Config.ActivityLookbackDays).ToString("o")
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/conversations?`$filter=lastDeliveryDateTime ge $lookbackDate&`$orderby=lastDeliveryDateTime desc"
        $response = Invoke-GraphAPI -Method "GET" -Uri $uri
        
        if ($response.value -and $response.value.Count -gt 0) {
            $latestConvo = $response.value | Where-Object { $_.lastDeliveryDateTime } | Select-Object -First 1
            if ($latestConvo -and $latestConvo.lastDeliveryDateTime) {
                return [datetime]$latestConvo.lastDeliveryDateTime
            }
        }
        
        return [DateTime]::MinValue
    }
    catch {
        Write-Verbose "Failed to get mailbox activity for group $GroupId : $_"
        return [DateTime]::MinValue
    }
}
#endregion

#region Classification & Analysis (Phase 6)
function Classify-TeamActivity {
    param([TeamActivity]$TeamData)
    
    # Determine last activity date as MAX of all sources
    $dates = @()
    if ($TeamData.LastChannelMessageDate -ne [DateTime]::MinValue) { $dates += $TeamData.LastChannelMessageDate }
    if ($TeamData.LastSharePointActivityDate -ne [DateTime]::MinValue) { $dates += $TeamData.LastSharePointActivityDate }
    if ($TeamData.LastMailboxActivityDate -ne [DateTime]::MinValue) { $dates += $TeamData.LastMailboxActivityDate }
    
    if ($dates.Count -gt 0) {
        $TeamData.LastActivityDate = ($dates | Measure-Object -Maximum).Maximum
    }
    else {
        $TeamData.LastActivityDate = $TeamData.CreatedDateTime
    }
    
    # Calculate days since activity
    $TeamData.DaysSinceActivity = [int]((Get-Date) - $TeamData.LastActivityDate).TotalDays
    
    # Classify based on thresholds
    if ($TeamData.DaysSinceActivity -gt $script:Config.InactiveThresholdDays) {
        $TeamData.RecommendedAction = "RECOMMEND ARCHIVAL"
        $TeamData.Details = "No activity detected in the last $($script:Config.InactiveThresholdDays) days"
    }
    elseif ($TeamData.DaysSinceActivity -gt $script:Config.ReviewThresholdDays) {
        $TeamData.RecommendedAction = "CONSIDER ARCHIVAL"
        $TeamData.Details = "Minimal activity in the last $($script:Config.ReviewThresholdDays)-$($script:Config.InactiveThresholdDays) days"
    }
    else {
        $TeamData.RecommendedAction = "ACTIVE"
        $TeamData.Details = "Recent activity detected ($($TeamData.DaysSinceActivity) days ago)"
    }
    
    return $TeamData
}
#endregion

#region Output Generation (Phase 7)
function Export-ToCSV {
    param([array]$TeamDataArray, [string]$OutputPath)
    
    $csvPath = Join-Path $OutputPath "TeamsAuditReport_$($script:ReportDateTime).csv"
    
    Write-Host "Exporting to CSV: $csvPath"
    
    $TeamDataArray | Select-Object `
        TeamName, Owner, CreatedBy, CreatedDateTime, MemberCount, GuestCount, `
        LastChannelMessageDate, LastSharePointActivityDate, LastMailboxActivityDate, `
        LastActivityDate, DaysSinceActivity, RecommendedAction, Description, `
        Privacy, IsArchived, SensitivityLabel, Details | `
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter "`t"
    
    Write-Host "✓ CSV report generated: $(Split-Path $csvPath -Leaf)" -ForegroundColor Green
    return $csvPath
}

function Export-ToJSON {
    param([array]$TeamDataArray, [string]$OutputPath)
    
    $jsonPath = Join-Path $OutputPath "TeamsAuditReport_$($script:ReportDateTime).json"
    
    Write-Host "Exporting to JSON: $jsonPath"
    
    $jsonData = @{
        metadata = @{
            author         = "Bjørnar Aassveen"
            blog           = "https://aassveen.com"
            version        = "1.0"
            generatedDate  = $script:StartTime.ToString("yyyy-MM-dd HH:mm:ss")
            reportCount    = $TeamDataArray.Count
        }
        summary = @{
            TotalTeams       = $TeamDataArray.Count
            RecommendArchival = ($TeamDataArray | Where-Object { $_.RecommendedAction -eq "RECOMMEND ARCHIVAL" }).Count
            ConsiderArchival  = ($TeamDataArray | Where-Object { $_.RecommendedAction -eq "CONSIDER ARCHIVAL" }).Count
            ActiveTeams      = ($TeamDataArray | Where-Object { $_.RecommendedAction -eq "ACTIVE" }).Count
        }
        teams = $TeamDataArray
    }
    
    $jsonData | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding UTF8
    
    Write-Host "✓ JSON report generated: $(Split-Path $jsonPath -Leaf)" -ForegroundColor Green
    return $jsonPath
}

function Export-ToHTML {
    param([array]$TeamDataArray, [string]$OutputPath)
    
    $htmlPath = Join-Path $OutputPath "TeamsAuditReport_$($script:ReportDateTime).html"
    
    Write-Host "Exporting to HTML: $htmlPath"
    
    # Calculate statistics
    $totalTeams = $TeamDataArray.Count
    $recommendArchival = ($TeamDataArray | Where-Object { $_.RecommendedAction -eq "RECOMMEND ARCHIVAL" }).Count
    $considerArchival = ($TeamDataArray | Where-Object { $_.RecommendedAction -eq "CONSIDER ARCHIVAL" }).Count
    $activeTeams = ($TeamDataArray | Where-Object { $_.RecommendedAction -eq "ACTIVE" }).Count
    $totalMembers = ($TeamDataArray | Measure-Object -Property MemberCount -Sum).Sum
    $totalGuests = ($TeamDataArray | Measure-Object -Property GuestCount -Sum).Sum
    
    # Build HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Teams Audit Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5; color: #333; }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
        header { background-color: #0078d4; color: white; padding: 20px; border-radius: 8px; margin-bottom: 30px; }
        header h1 { font-size: 28px; margin-bottom: 5px; }
        header p { font-size: 14px; opacity: 0.9; }
        .metadata { background-color: #e7f3ff; padding: 15px; border-left: 4px solid #0078d4; margin-bottom: 20px; font-size: 12px; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 30px; }
        .summary-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }
        .summary-card .number { font-size: 32px; font-weight: bold; margin: 10px 0; }
        .summary-card .label { font-size: 12px; color: #666; text-transform: uppercase; }
        .status-recommend { color: #d13438; }
        .status-consider { color: #f7630c; }
        .status-active { color: #107c10; }
        .table-container { background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
        table { width: 100%; border-collapse: collapse; }
        th { background-color: #f3f2f1; padding: 12px; text-align: left; font-weight: 600; font-size: 13px; border-bottom: 2px solid #e1dfdd; }
        td { padding: 12px; border-bottom: 1px solid #e1dfdd; font-size: 13px; }
        tr:hover { background-color: #f9f8f7; }
        .status-badge { padding: 4px 12px; border-radius: 12px; font-weight: 600; font-size: 12px; }
        .badge-recommend { background-color: #fed9cc; color: #d13438; }
        .badge-consider { background-color: #fff4ce; color: #f7630c; }
        .badge-active { background-color: #dffcf0; color: #107c10; }
        .date-cell { font-size: 12px; color: #666; }
        footer { margin-top: 40px; padding: 20px; text-align: center; font-size: 11px; color: #999; border-top: 1px solid #e1dfdd; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>Teams Audit Report</h1>
            <p>Generated: $($script:StartTime.ToString("yyyy-MM-dd HH:mm:ss"))</p>
            <div class="metadata">
                <strong>Author:</strong> Bjørnar Aassveen | <strong>Blog:</strong> https://aassveen.com | <strong>Version:</strong> 1.0 (2026-04-20)
            </div>
        </header>
        
        <div class="summary">
            <div class="summary-card">
                <div class="label">Total Teams</div>
                <div class="number">$totalTeams</div>
            </div>
            <div class="summary-card">
                <div class="label">Recommend Archival</div>
                <div class="number status-recommend">$recommendArchival</div>
            </div>
            <div class="summary-card">
                <div class="label">Consider Archival</div>
                <div class="number status-consider">$considerArchival</div>
            </div>
            <div class="summary-card">
                <div class="label">Active Teams</div>
                <div class="number status-active">$activeTeams</div>
            </div>
            <div class="summary-card">
                <div class="label">Total Members</div>
                <div class="number">$totalMembers</div>
            </div>
            <div class="summary-card">
                <div class="label">Total Guests</div>
                <div class="number">$totalGuests</div>
            </div>
        </div>
        
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Team Name</th>
                        <th>Owner</th>
                        <th>Members</th>
                        <th>Guests</th>
                        <th>Channel Activity</th>
                        <th>SharePoint Activity</th>
                        <th>Mailbox Activity</th>
                        <th>Last Activity</th>
                        <th>Days Since Activity</th>
                        <th>Status</th>
                        <th>Privacy</th>
                        <th>Archived</th>
                    </tr>
                </thead>
                <tbody>
"@
    
    foreach ($team in $TeamDataArray | Sort-Object DaysSinceActivity -Descending) {
        $statusClass = switch ($team.RecommendedAction) {
            "RECOMMEND ARCHIVAL" { "badge-recommend" }
            "CONSIDER ARCHIVAL" { "badge-consider" }
            "ACTIVE" { "badge-active" }
            default { "badge-active" }
        }
        
        $channelDate = if ($team.LastChannelMessageDate -ne [DateTime]::MinValue) { $team.LastChannelMessageDate.ToString("yyyy-MM-dd") } else { "N/A" }
        $sharePointDate = if ($team.LastSharePointActivityDate -ne [DateTime]::MinValue) { $team.LastSharePointActivityDate.ToString("yyyy-MM-dd") } else { "N/A" }
        $mailboxDate = if ($team.LastMailboxActivityDate -ne [DateTime]::MinValue) { $team.LastMailboxActivityDate.ToString("yyyy-MM-dd") } else { "N/A" }
        $lastActivityDate = if ($team.LastActivityDate -ne [DateTime]::MinValue) { $team.LastActivityDate.ToString("yyyy-MM-dd") } else { "N/A" }
        
        $html += @"
                    <tr>
                        <td>$($team.TeamName)</td>
                        <td>$($team.Owner)</td>
                        <td>$($team.MemberCount)</td>
                        <td>$($team.GuestCount)</td>
                        <td class="date-cell">$channelDate</td>
                        <td class="date-cell">$sharePointDate</td>
                        <td class="date-cell">$mailboxDate</td>
                        <td class="date-cell">$lastActivityDate</td>
                        <td>$($team.DaysSinceActivity)</td>
                        <td><span class="status-badge $statusClass">$($team.RecommendedAction)</span></td>
                        <td>$($team.Privacy)</td>
                        <td>$(if ($team.IsArchived) { "Yes" } else { "No" })</td>
                    </tr>
"@
    }
    
    $html += @"
                </tbody>
            </table>
        </div>
        
        <footer>
            <p>Generated by Bjørnar Aassveen (https://aassveen.com) | Version 1.0 (2026-04-20)</p>
            <p>This report was automatically generated by the Teams Audit Report script.</p>
            <p>Inactive threshold: $($script:Config.InactiveThresholdDays) days | Review threshold: $($script:Config.ReviewThresholdDays) days</p>
        </footer>
    </div>
</body>
</html>
"@
    
    $html | Set-Content -Path $htmlPath -Encoding UTF8
    
    Write-Host "✓ HTML report generated: $(Split-Path $htmlPath -Leaf)" -ForegroundColor Green
    return $htmlPath
}
#endregion

#region Main Execution
function Main {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Teams Audit Report Generator" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Phase 1: Authentication
    if (-not (Connect-ToGraph)) {
        exit 1
    }
    
    # Phase 2: Get all teams
    $allTeams = Get-AllTeams
    if (-not $allTeams) {
        Write-Error "Failed to fetch teams"
        exit 1
    }
    
    # Process teams and collect activity data
    $teamActivities = @()
    $processedCount = 0
    $totalCount = $allTeams.Count
    
    Write-Host "`nProcessing teams and collecting activity data...`n"
    
    foreach ($team in $allTeams) {
        $processedCount++
        [int]$percentComplete = ($processedCount / $totalCount) * 100
        Write-Progress -Activity "Processing Teams" -Status "Team $processedCount of $totalCount" -PercentComplete $percentComplete
        
        try {
            # Create team activity object
            $teamActivity = [TeamActivity]::new()
            $teamActivity.TeamId = $team.id
            $teamActivity.TeamName = $team.displayName
            
            # Get owner information (using group ID since teams are backed by groups)
            $teamActivity.Owner = Get-TeamOwners -GroupId $team.id
            
            # Get group metadata (created by, description, etc.)
            $groupInfo = Get-GroupInfo -GroupId $team.id
            if ($groupInfo) {
                $teamActivity.CreatedBy = $groupInfo.CreatedByUPN
                $teamActivity.CreatedDateTime = $groupInfo.CreatedBy
                $teamActivity.Description = $groupInfo.Description
            }
            
            # Get members and guests (using group ID)
            $memberInfo = Get-TeamMembers -GroupId $team.id
            $teamActivity.MemberCount = $memberInfo.MemberCount
            $teamActivity.GuestCount = $memberInfo.GuestCount
            
            # Get properties from team object
            $teamActivity.Privacy = if ($team.visibility -eq "Private") { "Private" } else { "Public" }
            $teamActivity.IsArchived = $team.isArchived
            $teamActivity.SensitivityLabel = if ($team.classification) { $team.classification } else { "None" }
            
            # Phase 3-5: Get activity dates (run sequentially to avoid throttling)
            $teamActivity.LastChannelMessageDate = Get-LastChannelMessageDate -TeamId $team.id
            $teamActivity.LastSharePointActivityDate = Get-LastSharePointActivityDate -GroupId $team.id
            $teamActivity.LastMailboxActivityDate = Get-LastMailboxActivityDate -GroupId $team.id
            
            # Phase 6: Classify based on activity
            $teamActivity = Classify-TeamActivity -TeamData $teamActivity
            
            $teamActivities += $teamActivity
        }
        catch {
            Write-Verbose "Error processing team $($team.displayName): $_"
        }
    }
    
    Write-Progress -Activity "Processing Teams" -Completed
    
    Write-Host "`n✓ Processed $($teamActivities.Count) teams successfully`n" -ForegroundColor Green
    
    # Phase 7: Generate reports
    Write-Host "Generating reports...`n"
    
    $csvPath = Export-ToCSV -TeamDataArray $teamActivities -OutputPath $script:Config.OutputPath
    $jsonPath = Export-ToJSON -TeamDataArray $teamActivities -OutputPath $script:Config.OutputPath
    $htmlPath = Export-ToHTML -TeamDataArray $teamActivities -OutputPath $script:Config.OutputPath
    
    # Summary
    $endTime = Get-Date
    $duration = $endTime - $script:StartTime
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Report Generation Complete" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Total Teams: $($teamActivities.Count)"
    Write-Host "  Recommend Archival: $(($teamActivities | Where-Object { $_.RecommendedAction -eq 'RECOMMEND ARCHIVAL' }).Count)" -ForegroundColor Red
    Write-Host "  Consider Archival: $(($teamActivities | Where-Object { $_.RecommendedAction -eq 'CONSIDER ARCHIVAL' }).Count)" -ForegroundColor Yellow
    Write-Host "  Active Teams: $(($teamActivities | Where-Object { $_.RecommendedAction -eq 'ACTIVE' }).Count)" -ForegroundColor Green
    Write-Host "  Duration: $($duration.TotalMinutes.ToString('0.0')) minutes`n"
    
    Write-Host "Reports generated at: $($script:Config.OutputPath)" -ForegroundColor Green
    Write-Host "  - CSV: $(Split-Path $csvPath -Leaf)"
    Write-Host "  - JSON: $(Split-Path $jsonPath -Leaf)"
    Write-Host "  - HTML: $(Split-Path $htmlPath -Leaf)`n"
    
    return $true
}

# Execute main
try {
    Main
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
