#Requires -Version 5.1
<#
.SYNOPSIS
    Advanced HTML Generator Module for Conditional Access Policy Documenter

.DESCRIPTION
    Generates a self-contained, interactive HTML report with multiple analysis views:
    - Policy Cards: Filterable, expandable policy details
    - Coverage Matrix: 2D grid of users/groups vs applications
    - Overlap Analyzer: Conflict and subset detection
#>

$script:HtmlTheme = @{
    PrimaryColor    = "#0078D4"
    EnabledColor    = "#107C10"
    DisabledColor   = "#A80000"
    ReportOnlyColor = "#FFB900"
}

function New-HtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Policies,

        [Parameter(Mandatory = $false)]
        [string]$Title = "Conditional Access Policies Report",

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $outputDir = Split-Path -Parent $Path
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $stats = @{
        Total      = $Policies.Count
        Enabled    = ($Policies | Where-Object { $_.StateRaw -eq "enabled" }).Count
        Disabled   = ($Policies | Where-Object { $_.StateRaw -eq "disabled" }).Count
        ReportOnly = ($Policies | Where-Object { $_.StateRaw -eq "enabledForReportingButNotEnforced" }).Count
    }

    $policyCards = $Policies | ForEach-Object { ConvertTo-PolicyHtmlCard -Policy $_ }
    $policyCardsHtml = $policyCards -join "`n"
    $policiesJson = @($Policies) | ConvertTo-Json -Depth 10 -Compress

    $html = Get-HtmlTemplate -Title $Title -Stats $stats -PoliciesHtml $policyCardsHtml -PoliciesJson $policiesJson
    $html | Out-File -FilePath $Path -Encoding UTF8

    Write-Host "HTML report saved to: $Path" -ForegroundColor Green
    return $true
}

function ConvertTo-PolicyHtmlCard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Policy
    )

    $stateClass = switch ($Policy.StateRaw) {
        "enabled" { "state-enabled" }
        "disabled" { "state-disabled" }
        "enabledForReportingButNotEnforced" { "state-reportonly" }
        default { "state-unknown" }
    }

    $policyName = [System.Web.HttpUtility]::HtmlEncode($Policy.DisplayName)
    $policyId = [System.Web.HttpUtility]::HtmlEncode($Policy.Id)

    $usersHtml = Get-UsersSectionHtml -Users $Policy.Conditions.Users
    $appsHtml = Get-ApplicationsSectionHtml -Applications $Policy.Conditions.Applications
    $conditionsHtml = Get-ConditionsSectionHtml -Conditions $Policy.Conditions
    $controlsHtml = Get-ControlsSectionHtml -GrantControls $Policy.GrantControls -SessionControls $Policy.SessionControls

    return @"
    <div class="policy-card" data-state="$($Policy.StateRaw)" data-name="$policyName" data-id="$policyId">
      <div class="policy-header" onclick="toggleCard(this)">
        <div class="policy-title">
          <h3>$policyName</h3>
          <span class="policy-id">$policyId</span>
        </div>
        <div class="policy-badges">
          <span class="state-badge $stateClass">$($Policy.State)</span>
          <span class="expand-icon">+</span>
        </div>
      </div>
      <div class="policy-content">
        <div class="policy-grid">
          <div class="policy-section">
            <h4>Users & Groups</h4>
            $usersHtml
          </div>
          <div class="policy-section">
            <h4>Applications</h4>
            $appsHtml
          </div>
          <div class="policy-section">
            <h4>Conditions</h4>
            $conditionsHtml
          </div>
          <div class="policy-section">
            <h4>Access Controls</h4>
            $controlsHtml
          </div>
        </div>
        <div class="policy-footer">
          <span class="modified-date">Modified: $($Policy.ModifiedDateTime)</span>
          <button class="copy-btn" onclick="event.stopPropagation(); copyPolicyId('$policyId')">Copy ID</button>
        </div>
      </div>
    </div>
"@
}

function Get-UsersSectionHtml {
    param([hashtable]$Users)
    $html = ""

    if ($Users.IncludeUsers.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Users:</strong><ul>"
        foreach ($user in $Users.IncludeUsers) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($user)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    if ($Users.IncludeGroups.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Groups:</strong><ul>"
        foreach ($group in $Users.IncludeGroups) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($group)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    if ($Users.IncludeRoles.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Roles:</strong><ul>"
        foreach ($role in $Users.IncludeRoles) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($role)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    $hasExclusions = ($Users.ExcludeUsers.Count -gt 0) -or ($Users.ExcludeGroups.Count -gt 0) -or ($Users.ExcludeRoles.Count -gt 0)
    if ($hasExclusions) {
        $html += "<div class='subsection exclusions'><strong>Exclusions:</strong><ul>"
        foreach ($user in $Users.ExcludeUsers) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($user)
            $html += "<li>$escaped (User)</li>"
        }
        foreach ($group in $Users.ExcludeGroups) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($group)
            $html += "<li>$escaped (Group)</li>"
        }
        foreach ($role in $Users.ExcludeRoles) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($role)
            $html += "<li>$escaped (Role)</li>"
        }
        $html += "</ul></div>"
    }

    if (-not $html) { $html = "<p class='empty'>No users configured</p>" }
    return $html
}

function Get-ApplicationsSectionHtml {
    param([hashtable]$Applications)
    $html = ""

    if ($Applications.IncludeApplications.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Apps:</strong><ul>"
        foreach ($app in $Applications.IncludeApplications) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($app)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    if ($Applications.IncludeUserActions.Count -gt 0) {
        $html += "<div class='subsection'><strong>User Actions:</strong><ul>"
        foreach ($action in $Applications.IncludeUserActions) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($action)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    if ($Applications.ExcludeApplications.Count -gt 0) {
        $html += "<div class='subsection exclusions'><strong>Exclude Apps:</strong><ul>"
        foreach ($app in $Applications.ExcludeApplications) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($app)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    if (-not $html) { $html = "<p class='empty'>No applications configured</p>" }
    return $html
}

function Get-ConditionsSectionHtml {
    param([hashtable]$Conditions)
    $html = "<ul class='conditions-list'>"

    if ($Conditions.Platforms.IncludePlatforms.Count -gt 0) {
        $platforms = ($Conditions.Platforms.IncludePlatforms | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Platforms:</strong> $platforms</li>"
    }

    if ($Conditions.Locations.IncludeLocations.Count -gt 0) {
        $locations = ($Conditions.Locations.IncludeLocations | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Locations:</strong> $locations</li>"
    }

    if ($Conditions.ClientAppTypes.Count -gt 0) {
        $clientApps = ($Conditions.ClientAppTypes | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Client Apps:</strong> $clientApps</li>"
    }

    if ($Conditions.SignInRiskLevels.Count -gt 0) {
        $risk = ($Conditions.SignInRiskLevels | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Sign-in Risk:</strong> $risk</li>"
    }

    if ($Conditions.UserRiskLevels.Count -gt 0) {
        $userRisk = ($Conditions.UserRiskLevels | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>User Risk:</strong> $userRisk</li>"
    }

    if ($Conditions.Devices.DeviceFilter) {
        $mode = [System.Web.HttpUtility]::HtmlEncode($Conditions.Devices.DeviceFilter.Mode)
        $rule = [System.Web.HttpUtility]::HtmlEncode($Conditions.Devices.DeviceFilter.Rule)
        $html += "<li><strong>Device Filter ($mode):</strong><br><code>$rule</code></li>"
    }

    $html += "</ul>"
    if ($html -eq "<ul class='conditions-list'></ul>") {
        $html = "<p class='empty'>No additional conditions</p>"
    }
    return $html
}

function Get-ControlsSectionHtml {
    param([hashtable]$GrantControls, [hashtable]$SessionControls)
    $html = ""

    if ($GrantControls.BuiltInControls.Count -gt 0) {
        $operator = if ($GrantControls.Operator -eq "AND") { "Require ALL of:" } else { "Require ONE of:" }
        $html += "<div class='subsection'><strong>Grant Controls ($operator)</strong><ul>"
        foreach ($control in $GrantControls.BuiltInControls) {
            $escaped = [System.Web.HttpUtility]::HtmlEncode($control)
            $html += "<li>$escaped</li>"
        }
        $html += "</ul></div>"
    }

    if ($GrantControls.AuthenticationStrength) {
        $authStrength = [System.Web.HttpUtility]::HtmlEncode($GrantControls.AuthenticationStrength.DisplayName)
        $html += "<div class='subsection'><strong>Authentication Strength:</strong> $authStrength</div>"
    }

    $sessionItems = @()
    if ($SessionControls.SignInFrequency) {
        $sif = [System.Web.HttpUtility]::HtmlEncode($SessionControls.SignInFrequency)
        $sessionItems += "<li>Sign-in frequency: $sif</li>"
    }
    if ($SessionControls.PersistentBrowser) {
        $pb = [System.Web.HttpUtility]::HtmlEncode($SessionControls.PersistentBrowser)
        $sessionItems += "<li>Persistent browser: $pb</li>"
    }
    if ($SessionControls.CloudAppSecurity) {
        $cas = [System.Web.HttpUtility]::HtmlEncode($SessionControls.CloudAppSecurity)
        $sessionItems += "<li>Cloud App Security: $cas</li>"
    }
    if ($SessionControls.ApplicationEnforcedRestrictions) {
        $sessionItems += "<li>App enforced restrictions: Enabled</li>"
    }

    if ($sessionItems.Count -gt 0) {
        $html += "<div class='subsection'><strong>Session Controls:</strong><ul>"
        $html += $sessionItems -join ""
        $html += "</ul></div>"
    }

    if (-not $html) { $html = "<p class='empty'>No access controls configured</p>" }
    return $html
}

function Get-HtmlTemplate {
    param(
        [string]$Title,
        [hashtable]$Stats,
        [string]$PoliciesHtml,
        [string]$PoliciesJson
    )

    $date = Get-Date -Format "yyyy-MM-dd HH:mm"

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$Title</title>
  <style>
    :root {
      --bg-primary: #0a0f1e;
      --bg-secondary: #111827;
      --bg-card: #1a2332;
      --bg-hover: #243044;
      --accent: $($script:HtmlTheme.PrimaryColor);
      --accent-dim: #005a9e;
      --text-primary: #e5e7eb;
      --text-secondary: #9ca3af;
      --text-muted: #6b7280;
      --border: #2d3748;
      --enabled: $($script:HtmlTheme.EnabledColor);
      --disabled: $($script:HtmlTheme.DisabledColor);
      --reportonly: $($script:HtmlTheme.ReportOnlyColor);
      --gap: #dc2626;
      --overlap: #f59e0b;
      --shadow: 0 4px 6px -1px rgba(0,0,0,0.3);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: var(--bg-primary);
      color: var(--text-primary);
      line-height: 1.5;
      min-height: 100vh;
    }

    /* Top Navigation */
    .top-nav {
      position: sticky;
      top: 0;
      z-index: 1000;
      background: var(--bg-secondary);
      border-bottom: 1px solid var(--border);
      padding: 0 1.5rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      height: 60px;
    }

    .nav-brand {
      font-size: 1.1rem;
      font-weight: 600;
      color: var(--accent);
    }

    .nav-tabs {
      display: flex;
      gap: 0.25rem;
      height: 100%;
    }

    .nav-tab {
      padding: 0 1.25rem;
      height: 100%;
      display: flex;
      align-items: center;
      color: var(--text-secondary);
      text-decoration: none;
      font-size: 0.9rem;
      font-weight: 500;
      border-bottom: 2px solid transparent;
      cursor: pointer;
      transition: all 0.15s;
    }

    .nav-tab:hover { color: var(--text-primary); background: var(--bg-hover); }
    .nav-tab.active { color: var(--accent); border-bottom-color: var(--accent); }

    .nav-stats {
      display: flex;
      gap: 1.5rem;
      font-size: 0.8rem;
    }

    .stat-item {
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }

    .stat-dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
    }

    .stat-dot.total { background: var(--accent); }
    .stat-dot.enabled { background: var(--enabled); }
    .stat-dot.reportonly { background: var(--reportonly); }
    .stat-dot.disabled { background: var(--disabled); }

    .stat-count { font-weight: 600; color: var(--text-primary); }
    .stat-label { color: var(--text-muted); }

    /* Main Content */
    main {
      max-width: 1600px;
      margin: 0 auto;
      padding: 1.5rem;
    }

    .view { display: none; }
    .view.active { display: block; }

    /* View Header */
    .view-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.5rem;
      flex-wrap: wrap;
      gap: 1rem;
    }

    .view-title {
      font-size: 1.5rem;
      font-weight: 600;
    }

    .view-controls {
      display: flex;
      gap: 0.75rem;
      align-items: center;
    }

    .search-input {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 0.5rem 1rem;
      color: var(--text-primary);
      font-size: 0.9rem;
      width: 250px;
    }

    .search-input:focus {
      outline: none;
      border-color: var(--accent);
    }

    .filter-select {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 0.5rem 0.75rem;
      color: var(--text-primary);
      font-size: 0.9rem;
      cursor: pointer;
    }

    .btn {
      background: var(--accent);
      color: white;
      border: none;
      border-radius: 6px;
      padding: 0.5rem 1rem;
      font-size: 0.85rem;
      font-weight: 500;
      cursor: pointer;
      transition: background 0.15s;
    }

    .btn:hover { background: var(--accent-dim); }
    .btn-secondary {
      background: var(--bg-card);
      border: 1px solid var(--border);
    }
    .btn-secondary:hover { background: var(--bg-hover); }

    /* Policy Cards */
    .cards-container {
      display: flex;
      flex-direction: column;
      gap: 0.75rem;
    }

    .policy-card {
      background: var(--bg-card);
      border-radius: 8px;
      border-left: 4px solid var(--accent);
      overflow: hidden;
      transition: box-shadow 0.15s;
    }

    .policy-card:hover { box-shadow: var(--shadow); }
    .policy-card[data-state="enabled"] { border-left-color: var(--enabled); }
    .policy-card[data-state="disabled"] { border-left-color: var(--disabled); }
    .policy-card[data-state="enabledForReportingButNotEnforced"] { border-left-color: var(--reportonly); }

    .policy-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 1rem 1.25rem;
      cursor: pointer;
      transition: background 0.15s;
    }

    .policy-header:hover { background: var(--bg-hover); }

    .policy-title h3 {
      font-size: 1rem;
      font-weight: 600;
      margin-bottom: 0.2rem;
    }

    .policy-id {
      font-family: 'SF Mono', Monaco, 'Cascadia Code', monospace;
      font-size: 0.7rem;
      color: var(--text-muted);
    }

    .policy-badges {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .state-badge {
      padding: 0.25rem 0.6rem;
      border-radius: 4px;
      font-size: 0.75rem;
      font-weight: 600;
      text-transform: uppercase;
    }

    .state-enabled { background: var(--enabled); color: white; }
    .state-disabled { background: var(--disabled); color: white; }
    .state-reportonly { background: var(--reportonly); color: #1a1a1a; }

    .expand-icon {
      font-size: 1.25rem;
      color: var(--text-muted);
      transition: transform 0.2s;
      font-weight: 300;
    }

    .policy-card.expanded .expand-icon { transform: rotate(45deg); }

    .policy-content {
      display: none;
      padding: 0 1.25rem 1.25rem;
      border-top: 1px solid var(--border);
    }

    .policy-card.expanded .policy-content { display: block; }

    .policy-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 1rem;
      margin-top: 1rem;
    }

    @media (max-width: 768px) {
      .policy-grid { grid-template-columns: 1fr; }
    }

    .policy-section {
      background: var(--bg-secondary);
      border-radius: 6px;
      padding: 1rem;
    }

    .policy-section h4 {
      font-size: 0.8rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: var(--accent);
      margin-bottom: 0.75rem;
      padding-bottom: 0.5rem;
      border-bottom: 1px solid var(--border);
    }

    .subsection {
      margin-bottom: 0.75rem;
    }

    .subsection strong {
      font-size: 0.8rem;
      color: var(--text-secondary);
      display: block;
      margin-bottom: 0.25rem;
    }

    .subsection ul {
      list-style: none;
      padding-left: 0;
    }

    .subsection li {
      font-size: 0.85rem;
      padding: 0.2rem 0;
      padding-left: 1rem;
      position: relative;
    }

    .subsection li::before {
      content: '';
      position: absolute;
      left: 0;
      top: 0.6rem;
      width: 4px;
      height: 4px;
      background: var(--accent);
      border-radius: 50%;
    }

    .exclusions {
      border-left: 2px solid var(--disabled);
      padding-left: 0.75rem;
      margin-top: 0.5rem;
    }

    .exclusions strong { color: var(--disabled); }

    .conditions-list {
      list-style: none;
    }

    .conditions-list li {
      font-size: 0.85rem;
      padding: 0.25rem 0;
    }

    .conditions-list code {
      display: block;
      background: var(--bg-primary);
      padding: 0.5rem;
      border-radius: 4px;
      font-family: 'SF Mono', Monaco, monospace;
      font-size: 0.75rem;
      margin-top: 0.25rem;
      word-break: break-all;
      color: var(--reportonly);
    }

    .empty {
      font-style: italic;
      color: var(--text-muted);
      font-size: 0.85rem;
    }

    .policy-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 1rem;
      padding-top: 1rem;
      border-top: 1px solid var(--border);
    }

    .modified-date {
      font-size: 0.75rem;
      color: var(--text-muted);
    }

    .copy-btn {
      padding: 0.3rem 0.6rem;
      background: var(--bg-secondary);
      border: 1px solid var(--border);
      border-radius: 4px;
      color: var(--text-secondary);
      font-size: 0.75rem;
      cursor: pointer;
      transition: all 0.15s;
    }

    .copy-btn:hover {
      background: var(--accent);
      color: white;
      border-color: var(--accent);
    }

    /* Coverage Matrix */
    .matrix-container {
      overflow-x: auto;
      background: var(--bg-card);
      border-radius: 8px;
      padding: 1rem;
    }

    .matrix-table {
      border-collapse: collapse;
      min-width: 100%;
      font-size: 0.8rem;
    }

    .matrix-table th,
    .matrix-table td {
      border: 1px solid var(--border);
      padding: 0.5rem;
      text-align: left;
      vertical-align: top;
      min-width: 120px;
    }

    .matrix-table th {
      background: var(--bg-secondary);
      font-weight: 600;
      position: sticky;
      top: 0;
      z-index: 10;
    }

    .matrix-table th:first-child {
      position: sticky;
      left: 0;
      z-index: 20;
    }

    .matrix-table td:first-child {
      background: var(--bg-secondary);
      font-weight: 500;
      position: sticky;
      left: 0;
      z-index: 5;
    }

    .matrix-cell {
      min-height: 40px;
    }

    .matrix-cell.gap {
      background: rgba(220, 38, 38, 0.15);
    }

    .matrix-cell.covered {
      background: rgba(16, 124, 16, 0.1);
    }

    .policy-chip {
      display: inline-block;
      background: var(--accent);
      color: white;
      padding: 0.15rem 0.4rem;
      border-radius: 3px;
      font-size: 0.7rem;
      margin: 0.1rem;
      cursor: pointer;
      max-width: 100px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .policy-chip:hover {
      background: var(--accent-dim);
    }

    .policy-chip.enabled { background: var(--enabled); }
    .policy-chip.disabled { background: var(--disabled); opacity: 0.6; }
    .policy-chip.reportonly { background: var(--reportonly); color: #1a1a1a; }

    /* Matrix Popover */
    .matrix-popover {
      display: none;
      position: fixed;
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 1rem;
      max-width: 400px;
      max-height: 300px;
      overflow-y: auto;
      z-index: 1000;
      box-shadow: 0 10px 25px rgba(0,0,0,0.5);
    }

    .matrix-popover.visible { display: block; }

    .matrix-popover h4 {
      font-size: 0.9rem;
      margin-bottom: 0.75rem;
      color: var(--accent);
    }

    .matrix-popover-close {
      position: absolute;
      top: 0.5rem;
      right: 0.5rem;
      background: none;
      border: none;
      color: var(--text-muted);
      cursor: pointer;
      font-size: 1.25rem;
    }

    /* Overlap Analyzer */
    .analysis-container {
      display: flex;
      flex-direction: column;
      gap: 1rem;
    }

    .analysis-summary {
      display: flex;
      gap: 1rem;
      flex-wrap: wrap;
    }

    .analysis-stat {
      background: var(--bg-card);
      border-radius: 8px;
      padding: 1rem 1.5rem;
      text-align: center;
      min-width: 150px;
    }

    .analysis-stat .count {
      font-size: 2rem;
      font-weight: 700;
      color: var(--accent);
    }

    .analysis-stat .label {
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-top: 0.25rem;
    }

    .analysis-stat.conflicts .count { color: var(--disabled); }
    .analysis-stat.overlaps .count { color: var(--reportonly); }

    .finding-card {
      background: var(--bg-card);
      border-radius: 8px;
      padding: 1.25rem;
      border-left: 4px solid var(--overlap);
    }

    .finding-card.conflict { border-left-color: var(--disabled); }

    .finding-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 1rem;
    }

    .finding-type {
      font-size: 0.75rem;
      font-weight: 600;
      text-transform: uppercase;
      padding: 0.2rem 0.5rem;
      border-radius: 3px;
      background: var(--overlap);
      color: #1a1a1a;
    }

    .finding-card.conflict .finding-type {
      background: var(--disabled);
      color: white;
    }

    .finding-explanation {
      font-size: 0.9rem;
      color: var(--text-secondary);
      margin-bottom: 1rem;
      padding: 0.75rem;
      background: var(--bg-secondary);
      border-radius: 6px;
    }

    .finding-policies {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 1rem;
    }

    @media (max-width: 768px) {
      .finding-policies { grid-template-columns: 1fr; }
    }

    .finding-policy {
      background: var(--bg-secondary);
      border-radius: 6px;
      padding: 1rem;
    }

    .finding-policy h5 {
      font-size: 0.9rem;
      margin-bottom: 0.5rem;
      color: var(--text-primary);
    }

    .finding-policy .detail {
      font-size: 0.8rem;
      color: var(--text-secondary);
      margin: 0.25rem 0;
    }

    .finding-policy .detail strong {
      color: var(--text-primary);
    }

    .no-findings {
      text-align: center;
      padding: 3rem;
      color: var(--text-muted);
    }

    .no-findings svg {
      width: 64px;
      height: 64px;
      margin-bottom: 1rem;
      opacity: 0.5;
    }

    /* Footer */
    footer {
      text-align: center;
      padding: 2rem;
      color: var(--text-muted);
      font-size: 0.8rem;
      border-top: 1px solid var(--border);
      margin-top: 2rem;
    }

    /* Scrollbar */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: var(--bg-secondary); }
    ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--text-muted); }
  </style>
</head>
<body>
  <nav class="top-nav">
    <span class="nav-brand">CA Policy Analyzer</span>
    <div class="nav-tabs">
      <div class="nav-tab active" data-view="cards">Policy Cards</div>
      <div class="nav-tab" data-view="matrix">Coverage Matrix</div>
      <div class="nav-tab" data-view="analyzer">Overlap Analyzer</div>
    </div>
    <div class="nav-stats">
      <div class="stat-item">
        <span class="stat-dot total"></span>
        <span class="stat-count">$($Stats.Total)</span>
        <span class="stat-label">Total</span>
      </div>
      <div class="stat-item">
        <span class="stat-dot enabled"></span>
        <span class="stat-count">$($Stats.Enabled)</span>
        <span class="stat-label">Enabled</span>
      </div>
      <div class="stat-item">
        <span class="stat-dot reportonly"></span>
        <span class="stat-count">$($Stats.ReportOnly)</span>
        <span class="stat-label">Report-Only</span>
      </div>
      <div class="stat-item">
        <span class="stat-dot disabled"></span>
        <span class="stat-count">$($Stats.Disabled)</span>
        <span class="stat-label">Disabled</span>
      </div>
    </div>
  </nav>

  <main>
    <!-- View 1: Policy Cards -->
    <div id="view-cards" class="view active">
      <div class="view-header">
        <h2 class="view-title">Policy Cards</h2>
        <div class="view-controls">
          <input type="search" class="search-input" id="card-search" placeholder="Search policies...">
          <select class="filter-select" id="state-filter">
            <option value="all">All States</option>
            <option value="enabled">Enabled</option>
            <option value="enabledForReportingButNotEnforced">Report-Only</option>
            <option value="disabled">Disabled</option>
          </select>
          <button class="btn btn-secondary" onclick="expandAllCards()">Expand All</button>
          <button class="btn btn-secondary" onclick="collapseAllCards()">Collapse All</button>
        </div>
      </div>
      <div class="cards-container" id="cards-container">
$PoliciesHtml
      </div>
    </div>

    <!-- View 2: Coverage Matrix -->
    <div id="view-matrix" class="view">
      <div class="view-header">
        <h2 class="view-title">Coverage Matrix</h2>
        <div class="view-controls">
          <label style="font-size: 0.85rem; color: var(--text-secondary);">
            <input type="checkbox" id="hide-disabled-matrix" checked> Hide disabled policies
          </label>
        </div>
      </div>
      <div class="matrix-container" id="matrix-container">
        <p style="color: var(--text-muted);">Loading matrix...</p>
      </div>
    </div>

    <!-- View 3: Overlap Analyzer -->
    <div id="view-analyzer" class="view">
      <div class="view-header">
        <h2 class="view-title">Overlap & Conflict Analyzer</h2>
        <div class="view-controls">
          <button class="btn" onclick="runAnalysis()">Re-analyze</button>
        </div>
      </div>
      <div class="analysis-container" id="analysis-container">
        <p style="color: var(--text-muted);">Analyzing policies...</p>
      </div>
    </div>
  </main>

  <div class="matrix-popover" id="matrix-popover">
    <button class="matrix-popover-close" onclick="hidePopover()">&times;</button>
    <h4 id="popover-title">Policies</h4>
    <div id="popover-content"></div>
  </div>

  <footer>
    Generated on $date | Conditional Access Policy Analyzer
  </footer>

  <script>
    const policiesData = [].concat($PoliciesJson);

    // Navigation
    document.querySelectorAll('.nav-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
        tab.classList.add('active');
        document.getElementById('view-' + tab.dataset.view).classList.add('active');

        if (tab.dataset.view === 'matrix') buildMatrix();
        if (tab.dataset.view === 'analyzer') runAnalysis();
      });
    });

    // Card functions
    function toggleCard(header) {
      header.parentElement.classList.toggle('expanded');
    }

    function expandAllCards() {
      document.querySelectorAll('.policy-card').forEach(c => {
        if (c.style.display !== 'none') c.classList.add('expanded');
      });
    }

    function collapseAllCards() {
      document.querySelectorAll('.policy-card').forEach(c => c.classList.remove('expanded'));
    }

    function copyPolicyId(id) {
      navigator.clipboard.writeText(id);
      const btn = event.target;
      const orig = btn.textContent;
      btn.textContent = 'Copied!';
      setTimeout(() => btn.textContent = orig, 1500);
    }

    // Card filtering
    document.getElementById('card-search').addEventListener('input', filterCards);
    document.getElementById('state-filter').addEventListener('change', filterCards);

    function filterCards() {
      const search = document.getElementById('card-search').value.toLowerCase();
      const state = document.getElementById('state-filter').value;

      document.querySelectorAll('.policy-card').forEach(card => {
        const name = card.dataset.name.toLowerCase();
        const cardState = card.dataset.state;
        const content = card.textContent.toLowerCase();

        const matchSearch = !search || name.includes(search) || content.includes(search);
        const matchState = state === 'all' || cardState === state;

        card.style.display = matchSearch && matchState ? 'block' : 'none';
      });
    }

    // Safe array helper — handles PS ConvertTo-Json unwrapping single items to scalars
    function arr(x) { return x == null ? [] : [].concat(x); }

    // Coverage Matrix
    function buildMatrix() {
      const hideDisabled = document.getElementById('hide-disabled-matrix').checked;
      const policies = hideDisabled
        ? policiesData.filter(p => p.StateRaw !== 'disabled')
        : policiesData;

      // Extract unique user populations
      const userPops = new Set(['All Users']);
      const apps = new Set(['All cloud apps']);

      policies.forEach(p => {
        const c = p.Conditions || {};
        const u = c.Users || {};
        const a = c.Applications || {};

        arr(u.IncludeUsers).forEach(x => { if (x !== 'All') userPops.add(x); });
        arr(u.IncludeGroups).forEach(x => userPops.add(x));
        arr(u.IncludeRoles).forEach(x => userPops.add(x));
        arr(a.IncludeApplications).forEach(x => { if (x !== 'All') apps.add(x); });
      });

      const userList = Array.from(userPops).sort();
      const appList = Array.from(apps).sort();

      // Build coverage map
      const coverage = {};
      userList.forEach(u => {
        coverage[u] = {};
        appList.forEach(a => coverage[u][a] = []);
      });

      policies.forEach(p => {
        const c = p.Conditions || {};
        const u = c.Users || {};
        const a = c.Applications || {};

        const targetUsers = new Set();
        if (arr(u.IncludeUsers).includes('All')) {
          userList.forEach(x => targetUsers.add(x));
        } else {
          arr(u.IncludeUsers).forEach(x => targetUsers.add(x));
          arr(u.IncludeGroups).forEach(x => targetUsers.add(x));
          arr(u.IncludeRoles).forEach(x => targetUsers.add(x));
        }

        const targetApps = new Set();
        if (arr(a.IncludeApplications).includes('All')) {
          appList.forEach(x => targetApps.add(x));
        } else {
          arr(a.IncludeApplications).forEach(x => targetApps.add(x));
        }

        targetUsers.forEach(user => {
          targetApps.forEach(app => {
            if (coverage[user] && coverage[user][app] !== undefined) {
              coverage[user][app].push(p);
            }
          });
        });
      });

      // Render table
      let html = '<table class="matrix-table"><thead><tr><th>User / App</th>';
      appList.forEach(app => {
        const shortApp = app.length > 20 ? app.substring(0, 18) + '...' : app;
        html += '<th title="' + escapeHtml(app) + '">' + escapeHtml(shortApp) + '</th>';
      });
      html += '</tr></thead><tbody>';

      userList.forEach(user => {
        const shortUser = user.length > 25 ? user.substring(0, 23) + '...' : user;
        html += '<tr><td title="' + escapeHtml(user) + '">' + escapeHtml(shortUser) + '</td>';
        appList.forEach(app => {
          const pols = coverage[user][app];
          const cellClass = pols.length === 0 ? 'gap' : 'covered';
          html += '<td class="matrix-cell ' + cellClass + '" onclick="showPopover(event, \'' +
            escapeHtml(user) + '\', \'' + escapeHtml(app) + '\')">';
          pols.slice(0, 3).forEach(pol => {
            const stateClass = pol.StateRaw === 'enabled' ? 'enabled' :
              (pol.StateRaw === 'disabled' ? 'disabled' : 'reportonly');
            const shortName = pol.DisplayName.length > 12 ?
              pol.DisplayName.substring(0, 10) + '...' : pol.DisplayName;
            html += '<span class="policy-chip ' + stateClass + '" title="' +
              escapeHtml(pol.DisplayName) + '">' + escapeHtml(shortName) + '</span>';
          });
          if (pols.length > 3) {
            html += '<span class="policy-chip">+' + (pols.length - 3) + '</span>';
          }
          html += '</td>';
        });
        html += '</tr>';
      });

      html += '</tbody></table>';
      document.getElementById('matrix-container').innerHTML = html;
    }

    document.getElementById('hide-disabled-matrix').addEventListener('change', buildMatrix);

    // Matrix Popover
    let currentPopoverData = null;

    function showPopover(event, user, app) {
      const hideDisabled = document.getElementById('hide-disabled-matrix').checked;
      const policies = hideDisabled
        ? policiesData.filter(p => p.StateRaw !== 'disabled')
        : policiesData;

      const matching = policies.filter(p => {
        const c = p.Conditions || {};
        const u = c.Users || {};
        const a = c.Applications || {};

        let matchUser = arr(u.IncludeUsers).includes('All') ||
          arr(u.IncludeUsers).includes(user) ||
          arr(u.IncludeGroups).includes(user) ||
          arr(u.IncludeRoles).includes(user);

        let matchApp = arr(a.IncludeApplications).includes('All') ||
          arr(a.IncludeApplications).includes(app);

        return matchUser && matchApp;
      });

      const popover = document.getElementById('matrix-popover');
      document.getElementById('popover-title').textContent = user + ' + ' + app;

      let content = '';
      if (matching.length === 0) {
        content = '<p style="color: var(--disabled);">No coverage - potential gap!</p>';
      } else {
        matching.forEach(p => {
          const gc = p.GrantControls || {};
          const controls = arr(gc.BuiltInControls).join(', ') || 'None';
          content += '<div style="margin-bottom: 0.75rem; padding: 0.5rem; background: var(--bg-secondary); border-radius: 4px;">';
          content += '<strong>' + escapeHtml(p.DisplayName) + '</strong>';
          content += '<div style="font-size: 0.8rem; color: var(--text-secondary);">';
          content += 'State: ' + p.State + '<br>';
          content += 'Controls: ' + escapeHtml(controls);
          content += '</div></div>';
        });
      }

      document.getElementById('popover-content').innerHTML = content;

      const rect = event.target.getBoundingClientRect();
      popover.style.left = Math.min(rect.left, window.innerWidth - 420) + 'px';
      popover.style.top = Math.min(rect.bottom + 5, window.innerHeight - 320) + 'px';
      popover.classList.add('visible');
    }

    function hidePopover() {
      document.getElementById('matrix-popover').classList.remove('visible');
    }

    document.addEventListener('click', (e) => {
      if (!e.target.closest('.matrix-cell') && !e.target.closest('.matrix-popover')) {
        hidePopover();
      }
    });

    // Overlap Analyzer
    function runAnalysis() {
      const activePolicies = policiesData.filter(p => p.StateRaw !== 'disabled');
      const findings = [];

      // Compare each pair of policies
      for (let i = 0; i < activePolicies.length; i++) {
        for (let j = i + 1; j < activePolicies.length; j++) {
          const p1 = activePolicies[i];
          const p2 = activePolicies[j];

          const overlap = checkOverlap(p1, p2);
          if (overlap.hasOverlap) {
            findings.push({
              type: overlap.isConflict ? 'conflict' : 'overlap',
              policy1: p1,
              policy2: p2,
              explanation: overlap.explanation
            });
          }
        }
      }

      const conflicts = findings.filter(f => f.type === 'conflict').length;
      const overlaps = findings.filter(f => f.type === 'overlap').length;

      let html = '<div class="analysis-summary">';
      html += '<div class="analysis-stat"><div class="count">' + activePolicies.length + '</div><div class="label">Active Policies</div></div>';
      html += '<div class="analysis-stat conflicts"><div class="count">' + conflicts + '</div><div class="label">Potential Conflicts</div></div>';
      html += '<div class="analysis-stat overlaps"><div class="count">' + overlaps + '</div><div class="label">Overlapping Scopes</div></div>';
      html += '</div>';

      if (findings.length === 0) {
        html += '<div class="no-findings">';
        html += '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>';
        html += '<h3>No conflicts or overlaps detected</h3>';
        html += '<p>All active policies have distinct scopes or consistent controls.</p>';
        html += '</div>';
      } else {
        findings.forEach(f => {
          const cardClass = f.type === 'conflict' ? 'conflict' : '';
          html += '<div class="finding-card ' + cardClass + '">';
          html += '<div class="finding-header">';
          html += '<span class="finding-type">' + (f.type === 'conflict' ? 'Potential Conflict' : 'Scope Overlap') + '</span>';
          html += '</div>';
          html += '<div class="finding-explanation">' + escapeHtml(f.explanation) + '</div>';
          html += '<div class="finding-policies">';

          [f.policy1, f.policy2].forEach(p => {
            const gc = p.GrantControls || {};
            const controls = arr(gc.BuiltInControls).join(', ') || 'None';
            const users = summarizeUsers(p);
            const apps = summarizeApps(p);

            html += '<div class="finding-policy">';
            html += '<h5>' + escapeHtml(p.DisplayName) + '</h5>';
            html += '<div class="detail"><strong>State:</strong> ' + p.State + '</div>';
            html += '<div class="detail"><strong>Users:</strong> ' + escapeHtml(users) + '</div>';
            html += '<div class="detail"><strong>Apps:</strong> ' + escapeHtml(apps) + '</div>';
            html += '<div class="detail"><strong>Controls:</strong> ' + escapeHtml(controls) + '</div>';
            html += '</div>';
          });

          html += '</div></div>';
        });
      }

      document.getElementById('analysis-container').innerHTML = html;
    }

    function checkOverlap(p1, p2) {
      const u1 = p1.Conditions?.Users || {};
      const u2 = p2.Conditions?.Users || {};
      const a1 = p1.Conditions?.Applications || {};
      const a2 = p2.Conditions?.Applications || {};

      // Check user overlap
      const users1 = new Set([
        ...arr(u1.IncludeUsers),
        ...arr(u1.IncludeGroups),
        ...arr(u1.IncludeRoles)
      ]);
      const users2 = new Set([
        ...arr(u2.IncludeUsers),
        ...arr(u2.IncludeGroups),
        ...arr(u2.IncludeRoles)
      ]);

      const userOverlap = users1.has('All') || users2.has('All') ||
        [...users1].some(u => users2.has(u));

      // Check app overlap
      const apps1 = new Set(arr(a1.IncludeApplications));
      const apps2 = new Set(arr(a2.IncludeApplications));

      const appOverlap = apps1.has('All') || apps2.has('All') ||
        [...apps1].some(a => apps2.has(a));

      if (!userOverlap || !appOverlap) {
        return { hasOverlap: false };
      }

      // Check for conflict (different controls)
      const gc1 = p1.GrantControls || {};
      const gc2 = p2.GrantControls || {};
      const controls1 = arr(gc1.BuiltInControls).sort().join(',');
      const controls2 = arr(gc2.BuiltInControls).sort().join(',');

      const isConflict = controls1 !== controls2 && controls1 && controls2;

      let explanation = '';
      if (isConflict) {
        explanation = 'These policies target overlapping users and applications but require different grant controls. ';
        explanation += 'Policy "' + p1.DisplayName + '" requires [' + arr(gc1.BuiltInControls).join(', ') + '] ';
        explanation += 'while "' + p2.DisplayName + '" requires [' + arr(gc2.BuiltInControls).join(', ') + '].';
      } else {
        // Check if one is subset of other
        const isSubset = (users1.has('All') || isSubsetOf(users2, users1)) &&
                        (apps1.has('All') || isSubsetOf(apps2, apps1));
        const isSuperset = (users2.has('All') || isSubsetOf(users1, users2)) &&
                          (apps2.has('All') || isSubsetOf(apps1, apps2));

        if (isSubset && !isSuperset) {
          explanation = '"' + p2.DisplayName + '" targets a subset of users/apps covered by "' + p1.DisplayName + '". Consider if the more specific policy is necessary.';
        } else if (isSuperset && !isSubset) {
          explanation = '"' + p1.DisplayName + '" targets a subset of users/apps covered by "' + p2.DisplayName + '". Consider if the more specific policy is necessary.';
        } else {
          explanation = 'These policies have overlapping scope (some users and apps are targeted by both). They apply the same controls, so this may be intentional redundancy.';
        }
      }

      return { hasOverlap: true, isConflict, explanation };
    }

    function isSubsetOf(setA, setB) {
      return [...setA].every(item => setB.has(item));
    }

    function summarizeUsers(p) {
      const u = (p.Conditions || {}).Users || {};
      const parts = [];
      if (arr(u.IncludeUsers).includes('All')) return 'All Users';
      const iu = arr(u.IncludeUsers); if (iu.length) parts.push(iu.length + ' users');
      const ig = arr(u.IncludeGroups); if (ig.length) parts.push(ig.length + ' groups');
      const ir = arr(u.IncludeRoles);  if (ir.length) parts.push(ir.length + ' roles');
      return parts.join(', ') || 'None specified';
    }

    function summarizeApps(p) {
      const a = (p.Conditions || {}).Applications || {};
      const ia = arr(a.IncludeApplications);
      if (ia.includes('All')) return 'All cloud apps';
      return ia.length ? ia.length + ' apps' : 'None specified';
    }

    function escapeHtml(str) {
      if (!str) return '';
      const div = document.createElement('div');
      div.textContent = str;
      return div.innerHTML;
    }
  </script>
</body>
</html>
"@
}

function Set-HtmlTheme {
    [CmdletBinding()]
    param(
        [string]$PrimaryColor,
        [string]$EnabledColor,
        [string]$DisabledColor,
        [string]$ReportOnlyColor
    )

    if ($PrimaryColor) { $script:HtmlTheme.PrimaryColor = $PrimaryColor }
    if ($EnabledColor) { $script:HtmlTheme.EnabledColor = $EnabledColor }
    if ($DisabledColor) { $script:HtmlTheme.DisabledColor = $DisabledColor }
    if ($ReportOnlyColor) { $script:HtmlTheme.ReportOnlyColor = $ReportOnlyColor }
}

Export-ModuleMember -Function @(
    'New-HtmlReport',
    'Set-HtmlTheme'
)
