# Examples: Using the Teams Audit Reports

**Author**: Bjørnar Aassveen | **Blog**: https://aassveen.com | **Version**: 1.0 (2026-04-20)

---

This document provides practical examples for using the output from the Teams Audit Report script.

## 1. Analyzing Reports in PowerShell

### Load JSON Report

```powershell
# Load the latest JSON report
$latestReport = Get-Item "Reports\TeamsAuditReport_*.json" | 
  Sort-Object LastWriteTime -Descending | 
  Select-Object -First 1
$report = Get-Content $latestReport.FullName | ConvertFrom-Json
```

### Get High-Level Statistics

```powershell
Write-Host "Total Teams: $($report.TotalTeams)"
Write-Host "Recommend Archival: $($report.RecommendArchival)"
Write-Host "Consider Archival: $($report.ConsiderArchival)"
Write-Host "Active Teams: $($report.ActiveTeams)"

# Calculate percentages
$recommendPercent = [math]::Round(($report.RecommendArchival / $report.TotalTeams) * 100, 1)
$considerPercent = [math]::Round(($report.ConsiderArchival / $report.TotalTeams) * 100, 1)
$activePercent = [math]::Round(($report.ActiveTeams / $report.TotalTeams) * 100, 1)

Write-Host "`nPercentages:"
Write-Host "  Recommend Archival: $recommendPercent%"
Write-Host "  Consider Archival: $considerPercent%"
Write-Host "  Active: $activePercent%"
```

### Find Teams by Recommendation

```powershell
# Teams recommended for archival
$archivalCandidates = $report.Teams | 
  Where-Object { $_.RecommendedAction -eq 'RECOMMEND ARCHIVAL' } |
  Sort-Object DaysSinceActivity -Descending

Write-Host "Teams Recommended for Archival ($($archivalCandidates.Count)):"
$archivalCandidates | ForEach-Object {
    Write-Host "  $($_.TeamName) - Inactive: $($_.DaysSinceActivity) days"
}

# Teams to consider
$reviewCandidates = $report.Teams | 
  Where-Object { $_.RecommendedAction -eq 'CONSIDER ARCHIVAL' } |
  Sort-Object DaysSinceActivity -Descending

Write-Host "`nTeams to Consider Archiving ($($reviewCandidates.Count)):"
$reviewCandidates | ForEach-Object {
    Write-Host "  $($_.TeamName) - Inactive: $($_.DaysSinceActivity) days"
}
```

### Find Most Inactive Teams

```powershell
$report.Teams | 
  Sort-Object DaysSinceActivity -Descending | 
  Select-Object -First 20 TeamName, LastActivityDate, DaysSinceActivity, RecommendedAction |
  Format-Table -AutoSize
```

### Find Most Active Teams

```powershell
$report.Teams | 
  Sort-Object DaysSinceActivity | 
  Select-Object -First 20 TeamName, LastActivityDate, DaysSinceActivity, MemberCount |
  Format-Table -AutoSize
```

### Find Largest Teams

```powershell
$report.Teams | 
  Sort-Object { $_.MemberCount + $_.GuestCount } -Descending |
  Select-Object -First 10 TeamName, MemberCount, GuestCount, @{Name="Total";Expression={$_.MemberCount + $_.GuestCount}} |
  Format-Table -AutoSize
```

### Find Teams with Many Guests

```powershell
$report.Teams | 
  Where-Object { $_.GuestCount -gt 0 } |
  Sort-Object GuestCount -Descending |
  Select-Object TeamName, MemberCount, GuestCount, @{Name="GuestPercent";Expression={[math]::Round(($_.GuestCount / ($_.MemberCount + $_.GuestCount)) * 100, 1)}} |
  Format-Table -AutoSize
```

### Find All Archived Teams

```powershell
$archivedTeams = $report.Teams | 
  Where-Object { $_.IsArchived -eq $true }

Write-Host "Archived Teams: $($archivedTeams.Count)"
$archivedTeams | 
  Select-Object TeamName, Owner, CreatedDateTime |
  Format-Table -AutoSize
```

### Export Specific Teams to CSV

```powershell
# Export archival candidates to CSV for stakeholder review
$report.Teams | 
  Where-Object { $_.RecommendedAction -eq 'RECOMMEND ARCHIVAL' } |
  Select-Object TeamName, Owner, MemberCount, GuestCount, LastActivityDate, DaysSinceActivity, Description |
  Export-Csv -Path "archival-candidates.csv" -NoTypeInformation -Encoding UTF8
```

## 2. Analysis with Excel

### Import CSV Report

1. Open Excel
2. **File** → **Open** → Select `TeamsAuditReport_*.csv`
3. Data is automatically formatted in columns

### Create Pivot Table

1. Select all data (Ctrl+A)
2. **Insert** → **Pivot Table** → **New Worksheet**
3. Drag fields:
   - **Rows**: `RecommendedAction`
   - **Values**: Count of `TeamName`
   - This shows breakdown: Active, Consider, Recommend

### Create Charts

#### Pie Chart: Status Distribution
1. Select columns: `RecommendedAction` and `TeamName` (with header)
2. **Insert** → **Pie Chart** → Choose style
3. Right-click → **Move Chart** → **New Sheet**

#### Bar Chart: Inactivity by Owner
1. Select columns: `Owner` and `DaysSinceActivity`
2. **Insert** → **Column Chart**
3. Filter to only "RECOMMEND ARCHIVAL" teams for clarity

#### Scatter Plot: Members vs Activity
1. Insert columns for Members (`MemberCount + GuestCount`) and Days Since Activity (`DaysSinceActivity`)
2. **Insert** → **Scatter Chart**
3. Reveals if size correlates with inactivity

### Add Filters

1. Select header row
2. **Data** → **AutoFilter**
3. Use dropdown filters on any column:
   - Filter by `RecommendedAction` = "RECOMMEND ARCHIVAL"
   - Filter by `Privacy` = "Private"
   - Filter by Owner = "Specific User"

### Create Conditional Formatting

1. Select `DaysSinceActivity` column
2. **Home** → **Conditional Formatting** → **Color Scales**
3. Green (low) to Red (high) shows inactivity visually

## 3. Automated Analysis Script

Here's a comprehensive analysis script:

```powershell
param(
    [string]$ReportPath = "Reports"
)

function Analyze-TeamsReport {
    # Get latest report
    $latestReport = Get-Item "$ReportPath\TeamsAuditReport_*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    
    if (-not $latestReport) {
        Write-Error "No report found in $ReportPath"
        return
    }
    
    $report = Get-Content $latestReport.FullName | ConvertFrom-Json
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Teams Audit Analysis Report" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Summary
    Write-Host "SUMMARY STATISTICS" -ForegroundColor Yellow
    Write-Host "  Total Teams: $($report.TotalTeams)"
    Write-Host "  Recommend Archival: $($report.RecommendArchival) ($([math]::Round(($report.RecommendArchival/$report.TotalTeams)*100,1))%)"
    Write-Host "  Consider Archival: $($report.ConsiderArchival) ($([math]::Round(($report.ConsiderArchival/$report.TotalTeams)*100,1))%)"
    Write-Host "  Active Teams: $($report.ActiveTeams) ($([math]::Round(($report.ActiveTeams/$report.TotalTeams)*100,1))%)`n"
    
    # Top archival candidates
    Write-Host "TOP ARCHIVAL CANDIDATES (Most Inactive)" -ForegroundColor Yellow
    $archival = $report.Teams | 
      Where-Object { $_.RecommendedAction -eq 'RECOMMEND ARCHIVAL' } |
      Sort-Object DaysSinceActivity -Descending |
      Select-Object -First 10
    
    $archival | ForEach-Object {
        Write-Host "  $($_.TeamName)" -ForegroundColor Red
        Write-Host "    Owner: $($_.Owner)"
        Write-Host "    Inactive: $($_.DaysSinceActivity) days"
        Write-Host "    Members: $($_.MemberCount), Guests: $($_.GuestCount)"
    }
    
    # Largest teams
    Write-Host "`nLARGEST TEAMS (by total people)" -ForegroundColor Yellow
    $largest = $report.Teams |
      Sort-Object { $_.MemberCount + $_.GuestCount } -Descending |
      Select-Object -First 5
    
    $largest | ForEach-Object {
        $total = $_.MemberCount + $_.GuestCount
        Write-Host "  $($_.TeamName): $total people ($($_.MemberCount) members, $($_.GuestCount) guests)"
    }
    
    # Activity breakdown
    Write-Host "`nACTIVITY BREAKDOWN" -ForegroundColor Yellow
    $channelOnly = $report.Teams | Where-Object { $_.LastChannelMessageDate -and -not $_.LastSharePointActivityDate -and -not $_.LastMailboxActivityDate } | Measure-Object | Select-Object -ExpandProperty Count
    $sharepointOnly = $report.Teams | Where-Object { -not $_.LastChannelMessageDate -and $_.LastSharePointActivityDate -and -not $_.LastMailboxActivityDate } | Measure-Object | Select-Object -ExpandProperty Count
    $mailboxOnly = $report.Teams | Where-Object { -not $_.LastChannelMessageDate -and -not $_.LastSharePointActivityDate -and $_.LastMailboxActivityDate } | Measure-Object | Select-Object -ExpandProperty Count
    $mixed = $report.Teams | Where-Object { ($_.LastChannelMessageDate, $_.LastSharePointActivityDate, $_.LastMailboxActivityDate | Where-Object { $_ }) -gt 1 } | Measure-Object | Select-Object -ExpandProperty Count
    $noActivity = $report.Teams | Where-Object { -not $_.LastChannelMessageDate -and -not $_.LastSharePointActivityDate -and -not $_.LastMailboxActivityDate } | Measure-Object | Select-Object -ExpandProperty Count
    
    Write-Host "  Activity Type:"
    Write-Host "    Channel messages only: $channelOnly teams"
    Write-Host "    SharePoint only: $sharepointOnly teams"
    Write-Host "    Mailbox only: $mailboxOnly teams"
    Write-Host "    Multiple sources: $mixed teams"
    Write-Host "    No activity: $noActivity teams`n"
    
    # Privacy breakdown
    Write-Host "TEAM PRIVACY" -ForegroundColor Yellow
    $public = $report.Teams | Where-Object { $_.Privacy -eq 'Public' } | Measure-Object | Select-Object -ExpandProperty Count
    $private = $report.Teams | Where-Object { $_.Privacy -eq 'Private' } | Measure-Object | Select-Object -ExpandProperty Count
    
    Write-Host "  Public Teams: $public"
    Write-Host "  Private Teams: $private`n"
}

Analyze-TeamsReport -ReportPath $ReportPath
```

Save as `Analyze-Report.ps1` and run:

```powershell
.\Analyze-Report.ps1 -ReportPath "Reports"
```

## 4. Send Report via Email

```powershell
param(
    [string]$EmailTo,
    [string]$EmailFrom,
    [string]$SMTPServer,
    [int]$SMTPPort = 587
)

function Send-AuditReportEmail {
    # Get latest reports
    $htmlReport = Get-Item "Reports\TeamsAuditReport_*.html" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    
    $jsonReport = Get-Item "Reports\TeamsAuditReport_*.json" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    
    $csvReport = Get-Item "Reports\TeamsAuditReport_*.csv" |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    
    $mailParams = @{
        To         = $EmailTo
        From       = $EmailFrom
        Subject    = "Teams Audit Report - $(Get-Date -Format 'yyyy-MM-dd')"
        Body       = "Please see attached Teams Audit Reports"
        Attachments = @(
            $htmlReport.FullName,
            $jsonReport.FullName,
            $csvReport.FullName
        )
        SmtpServer = $SMTPServer
        Port       = $SMTPPort
        UseSsl     = $true
    }
    
    Send-MailMessage @mailParams
    Write-Host "Email sent to $EmailTo"
}

Send-AuditReportEmail -EmailTo "admin@contoso.com" `
  -EmailFrom "reports@contoso.com" `
  -SMTPServer "smtp.outlook.com"
```

## 5. Integration Examples

### Export to Compliance Tool

```powershell
# Convert to format for governance/compliance tool
$report = Get-Content "Reports\TeamsAuditReport_*.json" | ConvertFrom-Json

$complianceFormat = $report.Teams | ForEach-Object {
    [PSCustomObject]@{
        TeamId              = $_.TeamId
        TeamName            = $_.TeamName
        Owner               = $_.Owner
        Status              = $_.RecommendedAction
        LastActivityDays    = $_.DaysSinceActivity
        ComplianceCheckDate = Get-Date
        ReviewedBy          = "Auto-Report"
    }
}

$complianceFormat | Export-Csv "compliance-review.csv" -NoTypeInformation
```

### Bulk Archival Preparation

```powershell
$report = Get-Content "Reports\TeamsAuditReport_*.json" | ConvertFrom-Json

# Export teams to archive
$toArchive = $report.Teams | 
  Where-Object { $_.RecommendedAction -eq 'RECOMMEND ARCHIVAL' } |
  Select-Object TeamId, TeamName, Owner

# Save for review before archiving
$toArchive | Export-Csv "teams-to-archive.csv" -NoTypeInformation

Write-Host "Teams to archive: $($toArchive.Count)"
Write-Host "Review and approve in teams-to-archive.csv before archiving"
```

---

**Author**: Bjørnar Aassveen  
**Blog**: https://aassveen.com  
**Version**: 1.0 (2026-04-20)

**Need more examples?** Create a custom analysis based on the JSON structure or modify these scripts for your specific needs.
