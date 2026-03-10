#Requires -Version 5.1
<#
.SYNOPSIS
    HTML Generator Module for Conditional Access Policy Documenter

.DESCRIPTION
    Generates a self-contained, interactive HTML report with embedded CSS and JavaScript.
    No external dependencies required.
#>

# Script-level theme configuration
$script:HtmlTheme = @{
    PrimaryColor    = "#0078D4"
    EnabledColor    = "#107C10"
    DisabledColor   = "#A80000"
    ReportOnlyColor = "#FFB900"
}

function New-HtmlReport {
    <#
    .SYNOPSIS
        Generates a complete HTML report for Conditional Access Policies

    .PARAMETER Policies
        Array of parsed policy objects

    .PARAMETER Title
        Report title

    .PARAMETER Path
        Output file path
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Policies,

        [Parameter(Mandatory = $false)]
        [string]$Title = "Conditional Access Policies Report",

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Ensure output directory exists
    $outputDir = Split-Path -Parent $Path
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    # Calculate statistics
    $stats = @{
        Total      = $Policies.Count
        Enabled    = ($Policies | Where-Object { $_.StateRaw -eq "enabled" }).Count
        Disabled   = ($Policies | Where-Object { $_.StateRaw -eq "disabled" }).Count
        ReportOnly = ($Policies | Where-Object { $_.StateRaw -eq "enabledForReportingButNotEnforced" }).Count
    }

    # Generate policy cards HTML
    $policyCards = $Policies | ForEach-Object {
        ConvertTo-PolicyHtmlCard -Policy $_
    }
    $policyCardsHtml = $policyCards -join "`n"

    # Convert policies to JSON for JavaScript
    $policiesJson = $Policies | ConvertTo-Json -Depth 10 -Compress

    # Generate the complete HTML document
    $html = Get-HtmlTemplate -Title $Title -Stats $stats -PoliciesHtml $policyCardsHtml -PoliciesJson $policiesJson

    # Write to file
    $html | Out-File -FilePath $Path -Encoding UTF8

    Write-Host "HTML report saved to: $Path" -ForegroundColor Green
    return $true
}

function ConvertTo-PolicyHtmlCard {
    <#
    .SYNOPSIS
        Converts a policy object to an HTML card element
    #>
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

    # Build users section
    $usersHtml = Get-UsersSectionHtml -Users $Policy.Conditions.Users

    # Build applications section
    $appsHtml = Get-ApplicationsSectionHtml -Applications $Policy.Conditions.Applications

    # Build conditions section
    $conditionsHtml = Get-ConditionsSectionHtml -Conditions $Policy.Conditions

    # Build controls section
    $controlsHtml = Get-ControlsSectionHtml -GrantControls $Policy.GrantControls -SessionControls $Policy.SessionControls

    return @"
    <div class="policy-card" data-state="$($Policy.StateRaw)" data-name="$policyName">
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
          <span>Modified: $($Policy.ModifiedDateTime)</span>
          <button class="copy-btn" onclick="copyPolicyId('$policyId')">Copy ID</button>
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
            $escapedUser = [System.Web.HttpUtility]::HtmlEncode($user)
            $html += "<li>$escapedUser</li>"
        }
        $html += "</ul></div>"
    }

    if ($Users.IncludeGroups.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Groups:</strong><ul>"
        foreach ($group in $Users.IncludeGroups) {
            $escapedGroup = [System.Web.HttpUtility]::HtmlEncode($group)
            $html += "<li>$escapedGroup</li>"
        }
        $html += "</ul></div>"
    }

    if ($Users.IncludeRoles.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Roles:</strong><ul>"
        foreach ($role in $Users.IncludeRoles) {
            $escapedRole = [System.Web.HttpUtility]::HtmlEncode($role)
            $html += "<li>$escapedRole</li>"
        }
        $html += "</ul></div>"
    }

    # Exclusions
    $hasExclusions = ($Users.ExcludeUsers.Count -gt 0) -or ($Users.ExcludeGroups.Count -gt 0) -or ($Users.ExcludeRoles.Count -gt 0)
    if ($hasExclusions) {
        $html += "<div class='subsection exclusions'><strong>Exclusions:</strong><ul>"
        foreach ($user in $Users.ExcludeUsers) {
            $escapedUser = [System.Web.HttpUtility]::HtmlEncode($user)
            $html += "<li>$escapedUser (User)</li>"
        }
        foreach ($group in $Users.ExcludeGroups) {
            $escapedGroup = [System.Web.HttpUtility]::HtmlEncode($group)
            $html += "<li>$escapedGroup (Group)</li>"
        }
        foreach ($role in $Users.ExcludeRoles) {
            $escapedRole = [System.Web.HttpUtility]::HtmlEncode($role)
            $html += "<li>$escapedRole (Role)</li>"
        }
        $html += "</ul></div>"
    }

    if (-not $html) {
        $html = "<p class='empty'>No users configured</p>"
    }

    return $html
}

function Get-ApplicationsSectionHtml {
    param([hashtable]$Applications)

    $html = ""

    if ($Applications.IncludeApplications.Count -gt 0) {
        $html += "<div class='subsection'><strong>Include Apps:</strong><ul>"
        foreach ($app in $Applications.IncludeApplications) {
            $escapedApp = [System.Web.HttpUtility]::HtmlEncode($app)
            $html += "<li>$escapedApp</li>"
        }
        $html += "</ul></div>"
    }

    if ($Applications.IncludeUserActions.Count -gt 0) {
        $html += "<div class='subsection'><strong>User Actions:</strong><ul>"
        foreach ($action in $Applications.IncludeUserActions) {
            $escapedAction = [System.Web.HttpUtility]::HtmlEncode($action)
            $html += "<li>$escapedAction</li>"
        }
        $html += "</ul></div>"
    }

    if ($Applications.ExcludeApplications.Count -gt 0) {
        $html += "<div class='subsection exclusions'><strong>Exclude Apps:</strong><ul>"
        foreach ($app in $Applications.ExcludeApplications) {
            $escapedApp = [System.Web.HttpUtility]::HtmlEncode($app)
            $html += "<li>$escapedApp</li>"
        }
        $html += "</ul></div>"
    }

    if (-not $html) {
        $html = "<p class='empty'>No applications configured</p>"
    }

    return $html
}

function Get-ConditionsSectionHtml {
    param([hashtable]$Conditions)

    $html = "<ul class='conditions-list'>"

    # Platforms
    if ($Conditions.Platforms.IncludePlatforms.Count -gt 0) {
        $platforms = ($Conditions.Platforms.IncludePlatforms | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Platforms:</strong> $platforms</li>"
    }

    # Locations
    if ($Conditions.Locations.IncludeLocations.Count -gt 0) {
        $locations = ($Conditions.Locations.IncludeLocations | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Locations:</strong> $locations</li>"
    }

    # Client App Types
    if ($Conditions.ClientAppTypes.Count -gt 0) {
        $clientApps = ($Conditions.ClientAppTypes | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Client Apps:</strong> $clientApps</li>"
    }

    # Sign-in Risk
    if ($Conditions.SignInRiskLevels.Count -gt 0) {
        $riskLevels = ($Conditions.SignInRiskLevels | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>Sign-in Risk:</strong> $riskLevels</li>"
    }

    # User Risk
    if ($Conditions.UserRiskLevels.Count -gt 0) {
        $userRiskLevels = ($Conditions.UserRiskLevels | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ", "
        $html += "<li><strong>User Risk:</strong> $userRiskLevels</li>"
    }

    # Device Filter
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

    # Grant Controls
    if ($GrantControls.BuiltInControls.Count -gt 0) {
        $operator = if ($GrantControls.Operator -eq "AND") { "Require ALL of:" } else { "Require ONE of:" }
        $html += "<div class='subsection'><strong>Grant Controls ($operator)</strong><ul>"
        foreach ($control in $GrantControls.BuiltInControls) {
            $escapedControl = [System.Web.HttpUtility]::HtmlEncode($control)
            $html += "<li>$escapedControl</li>"
        }
        $html += "</ul></div>"
    }

    # Authentication Strength
    if ($GrantControls.AuthenticationStrength) {
        $authStrength = [System.Web.HttpUtility]::HtmlEncode($GrantControls.AuthenticationStrength.DisplayName)
        $html += "<div class='subsection'><strong>Authentication Strength:</strong> $authStrength</div>"
    }

    # Session Controls
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

    if ($SessionControls.ContinuousAccessEvaluation) {
        $cae = [System.Web.HttpUtility]::HtmlEncode($SessionControls.ContinuousAccessEvaluation)
        $sessionItems += "<li>Continuous Access Evaluation: $cae</li>"
    }

    if ($sessionItems.Count -gt 0) {
        $html += "<div class='subsection'><strong>Session Controls:</strong><ul>"
        $html += $sessionItems -join ""
        $html += "</ul></div>"
    }

    if (-not $html) {
        $html = "<p class='empty'>No access controls configured</p>"
    }

    return $html
}

function Get-HtmlTemplate {
    param(
        [string]$Title,
        [hashtable]$Stats,
        [string]$PoliciesHtml,
        [string]$PoliciesJson
    )

    $date = Get-Date -Format "MMMM d, yyyy 'at' h:mm tt"

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$Title</title>
  <style>
    :root {
      --primary-color: $($script:HtmlTheme.PrimaryColor);
      --enabled-color: $($script:HtmlTheme.EnabledColor);
      --disabled-color: $($script:HtmlTheme.DisabledColor);
      --reportonly-color: $($script:HtmlTheme.ReportOnlyColor);
      --bg-color: #f5f5f5;
      --card-bg: #ffffff;
      --text-color: #333333;
      --text-secondary: #666666;
      --border-color: #e0e0e0;
    }

    [data-theme="dark"] {
      --bg-color: #1a1a2e;
      --card-bg: #16213e;
      --text-color: #eaeaea;
      --text-secondary: #b0b0b0;
      --border-color: #2a2a4a;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: var(--bg-color);
      color: var(--text-color);
      line-height: 1.6;
    }

    header {
      background: linear-gradient(135deg, var(--primary-color), #005a9e);
      color: white;
      padding: 2rem;
      position: sticky;
      top: 0;
      z-index: 100;
      box-shadow: 0 2px 10px rgba(0,0,0,0.2);
    }

    header h1 {
      font-size: 1.8rem;
      font-weight: 300;
      margin-bottom: 1rem;
    }

    .controls {
      display: flex;
      flex-wrap: wrap;
      gap: 1rem;
      align-items: center;
      margin-bottom: 1rem;
    }

    .controls input[type="search"] {
      flex: 1;
      min-width: 200px;
      padding: 0.75rem 1rem;
      border: none;
      border-radius: 4px;
      font-size: 1rem;
      background: rgba(255,255,255,0.9);
    }

    .controls select {
      padding: 0.75rem 1rem;
      border: none;
      border-radius: 4px;
      font-size: 1rem;
      background: rgba(255,255,255,0.9);
      cursor: pointer;
    }

    .controls button {
      padding: 0.75rem 1.5rem;
      border: 2px solid white;
      border-radius: 4px;
      background: transparent;
      color: white;
      font-size: 1rem;
      cursor: pointer;
      transition: all 0.2s;
    }

    .controls button:hover {
      background: white;
      color: var(--primary-color);
    }

    .summary {
      display: flex;
      flex-wrap: wrap;
      gap: 1.5rem;
    }

    .summary-item {
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }

    .summary-count {
      font-size: 1.5rem;
      font-weight: bold;
    }

    .summary-label {
      opacity: 0.9;
    }

    .dot {
      width: 12px;
      height: 12px;
      border-radius: 50%;
      display: inline-block;
    }

    .dot-total { background: white; }
    .dot-enabled { background: var(--enabled-color); }
    .dot-disabled { background: var(--disabled-color); }
    .dot-reportonly { background: var(--reportonly-color); }

    main {
      max-width: 1400px;
      margin: 0 auto;
      padding: 2rem;
    }

    .policy-card {
      background: var(--card-bg);
      border-radius: 8px;
      margin-bottom: 1rem;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      overflow: hidden;
      transition: box-shadow 0.2s;
    }

    .policy-card:hover {
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }

    .policy-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 1rem 1.5rem;
      cursor: pointer;
      border-left: 4px solid var(--primary-color);
    }

    .policy-header:hover {
      background: rgba(0,0,0,0.02);
    }

    .policy-title h3 {
      font-size: 1.1rem;
      font-weight: 600;
      margin-bottom: 0.25rem;
    }

    .policy-id {
      font-size: 0.75rem;
      color: var(--text-secondary);
      font-family: monospace;
    }

    .policy-badges {
      display: flex;
      align-items: center;
      gap: 1rem;
    }

    .state-badge {
      padding: 0.25rem 0.75rem;
      border-radius: 20px;
      font-size: 0.8rem;
      font-weight: 600;
      color: white;
    }

    .state-enabled { background: var(--enabled-color); }
    .state-disabled { background: var(--disabled-color); }
    .state-reportonly { background: var(--reportonly-color); }

    .expand-icon {
      font-size: 1.5rem;
      color: var(--text-secondary);
      transition: transform 0.2s;
    }

    .policy-card.expanded .expand-icon {
      transform: rotate(45deg);
    }

    .policy-content {
      display: none;
      padding: 0 1.5rem 1.5rem;
      border-top: 1px solid var(--border-color);
    }

    .policy-card.expanded .policy-content {
      display: block;
    }

    .policy-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 1.5rem;
      margin-top: 1.5rem;
    }

    .policy-section {
      background: var(--bg-color);
      padding: 1rem;
      border-radius: 6px;
    }

    .policy-section h4 {
      color: var(--primary-color);
      font-size: 0.9rem;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 0.75rem;
      padding-bottom: 0.5rem;
      border-bottom: 2px solid var(--primary-color);
    }

    .subsection {
      margin-bottom: 0.75rem;
    }

    .subsection strong {
      font-size: 0.85rem;
      display: block;
      margin-bottom: 0.25rem;
    }

    .subsection ul {
      list-style: disc;
      padding-left: 1.2rem;
      margin: 0.25rem 0;
    }

    .subsection li {
      font-size: 0.85rem;
      padding: 0.15rem 0;
    }

    .exclusions {
      border-left: 3px solid var(--disabled-color);
      padding-left: 0.75rem;
      margin-top: 0.5rem;
    }

    .exclusions strong {
      color: var(--disabled-color);
    }

    .conditions-list {
      list-style: none;
    }

    .conditions-list li {
      font-size: 0.85rem;
      padding: 0.3rem 0;
    }

    .conditions-list code {
      display: block;
      background: var(--card-bg);
      padding: 0.5rem;
      border-radius: 4px;
      font-size: 0.8rem;
      margin-top: 0.25rem;
      word-break: break-all;
    }

    .empty {
      font-style: italic;
      color: var(--text-secondary);
      font-size: 0.85rem;
    }

    .policy-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 1rem;
      padding-top: 1rem;
      border-top: 1px solid var(--border-color);
      font-size: 0.8rem;
      color: var(--text-secondary);
    }

    .copy-btn {
      padding: 0.4rem 0.8rem;
      background: var(--primary-color);
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 0.8rem;
      transition: background 0.2s;
    }

    .copy-btn:hover {
      background: #005a9e;
    }

    .no-results {
      text-align: center;
      padding: 3rem;
      color: var(--text-secondary);
    }

    .no-results h3 {
      margin-bottom: 0.5rem;
    }

    footer {
      text-align: center;
      padding: 2rem;
      color: var(--text-secondary);
      font-size: 0.85rem;
    }

    @media print {
      header {
        position: relative;
        background: var(--primary-color) !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }

      .controls {
        display: none;
      }

      .policy-card {
        break-inside: avoid;
        page-break-inside: avoid;
      }

      .policy-content {
        display: block !important;
      }

      .expand-icon, .copy-btn {
        display: none;
      }
    }

    @media (max-width: 768px) {
      header {
        padding: 1rem;
      }

      header h1 {
        font-size: 1.4rem;
      }

      .controls {
        flex-direction: column;
      }

      .controls input[type="search"],
      .controls select {
        width: 100%;
      }

      main {
        padding: 1rem;
      }

      .policy-grid {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>$Title</h1>
    <div class="controls">
      <input type="search" id="search" placeholder="Search policies..." oninput="filterPolicies()">
      <select id="stateFilter" onchange="filterPolicies()">
        <option value="all">All States</option>
        <option value="enabled">Enabled</option>
        <option value="disabled">Disabled</option>
        <option value="enabledForReportingButNotEnforced">Report-Only</option>
      </select>
      <select id="sortBy" onchange="sortPolicies()">
        <option value="name">Sort by Name</option>
        <option value="state">Sort by State</option>
      </select>
      <button onclick="toggleTheme()">Toggle Theme</button>
      <button onclick="expandAll()">Expand All</button>
      <button onclick="collapseAll()">Collapse All</button>
      <button onclick="exportJson()">Export JSON</button>
    </div>
    <div class="summary">
      <div class="summary-item">
        <span class="dot dot-total"></span>
        <span class="summary-count">$($Stats.Total)</span>
        <span class="summary-label">Total</span>
      </div>
      <div class="summary-item">
        <span class="dot dot-enabled"></span>
        <span class="summary-count">$($Stats.Enabled)</span>
        <span class="summary-label">Enabled</span>
      </div>
      <div class="summary-item">
        <span class="dot dot-reportonly"></span>
        <span class="summary-count">$($Stats.ReportOnly)</span>
        <span class="summary-label">Report-Only</span>
      </div>
      <div class="summary-item">
        <span class="dot dot-disabled"></span>
        <span class="summary-count">$($Stats.Disabled)</span>
        <span class="summary-label">Disabled</span>
      </div>
    </div>
  </header>

  <main id="policies-container">
$PoliciesHtml
  </main>

  <footer>
    <p>Generated on $date</p>
    <p>Conditional Access Policy Documenter</p>
  </footer>

  <script>
    // Store policies data for filtering/export
    const policiesData = $PoliciesJson;

    function toggleCard(header) {
      const card = header.parentElement;
      card.classList.toggle('expanded');
    }

    function filterPolicies() {
      const searchTerm = document.getElementById('search').value.toLowerCase();
      const stateFilter = document.getElementById('stateFilter').value;
      const cards = document.querySelectorAll('.policy-card');
      let visibleCount = 0;

      cards.forEach(card => {
        const name = card.getAttribute('data-name').toLowerCase();
        const state = card.getAttribute('data-state');
        const content = card.textContent.toLowerCase();

        const matchesSearch = name.includes(searchTerm) || content.includes(searchTerm);
        const matchesState = stateFilter === 'all' || state === stateFilter;

        if (matchesSearch && matchesState) {
          card.style.display = 'block';
          visibleCount++;
        } else {
          card.style.display = 'none';
        }
      });

      // Show/hide no results message
      let noResults = document.getElementById('no-results');
      if (visibleCount === 0) {
        if (!noResults) {
          noResults = document.createElement('div');
          noResults.id = 'no-results';
          noResults.className = 'no-results';
          noResults.innerHTML = '<h3>No policies found</h3><p>Try adjusting your search or filter criteria.</p>';
          document.getElementById('policies-container').appendChild(noResults);
        }
        noResults.style.display = 'block';
      } else if (noResults) {
        noResults.style.display = 'none';
      }
    }

    function sortPolicies() {
      const sortBy = document.getElementById('sortBy').value;
      const container = document.getElementById('policies-container');
      const cards = Array.from(container.querySelectorAll('.policy-card'));

      cards.sort((a, b) => {
        if (sortBy === 'name') {
          return a.getAttribute('data-name').localeCompare(b.getAttribute('data-name'));
        } else if (sortBy === 'state') {
          const stateOrder = { 'enabled': 0, 'enabledForReportingButNotEnforced': 1, 'disabled': 2 };
          return (stateOrder[a.getAttribute('data-state')] || 3) - (stateOrder[b.getAttribute('data-state')] || 3);
        }
        return 0;
      });

      cards.forEach(card => container.appendChild(card));
    }

    function toggleTheme() {
      const body = document.body;
      const currentTheme = body.getAttribute('data-theme');
      body.setAttribute('data-theme', currentTheme === 'dark' ? 'light' : 'dark');
      localStorage.setItem('theme', body.getAttribute('data-theme'));
    }

    function expandAll() {
      document.querySelectorAll('.policy-card').forEach(card => {
        if (card.style.display !== 'none') {
          card.classList.add('expanded');
        }
      });
    }

    function collapseAll() {
      document.querySelectorAll('.policy-card').forEach(card => {
        card.classList.remove('expanded');
      });
    }

    function copyPolicyId(id) {
      navigator.clipboard.writeText(id).then(() => {
        const btn = event.target;
        const originalText = btn.textContent;
        btn.textContent = 'Copied!';
        setTimeout(() => { btn.textContent = originalText; }, 1500);
      });
    }

    function exportJson() {
      const stateFilter = document.getElementById('stateFilter').value;
      const searchTerm = document.getElementById('search').value.toLowerCase();

      let filteredData = policiesData;

      if (stateFilter !== 'all') {
        filteredData = filteredData.filter(p => p.StateRaw === stateFilter);
      }

      if (searchTerm) {
        filteredData = filteredData.filter(p =>
          p.DisplayName.toLowerCase().includes(searchTerm) ||
          JSON.stringify(p).toLowerCase().includes(searchTerm)
        );
      }

      const blob = new Blob([JSON.stringify(filteredData, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'conditional-access-policies.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    // Initialize theme from localStorage
    (function() {
      const savedTheme = localStorage.getItem('theme');
      if (savedTheme) {
        document.body.setAttribute('data-theme', savedTheme);
      }
    })();
  </script>
</body>
</html>
"@
}

function Set-HtmlTheme {
    <#
    .SYNOPSIS
        Sets the color theme for generated HTML reports
    #>
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

# Export functions
Export-ModuleMember -Function @(
    'New-HtmlReport',
    'Set-HtmlTheme'
)
