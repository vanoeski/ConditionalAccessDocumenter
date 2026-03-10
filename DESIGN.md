# Conditional Access Policy Documenter - Design Document

## Overview

A PowerShell-based solution that queries Microsoft Entra ID (Azure AD) Conditional Access Policies via Microsoft Graph API and generates both a PowerPoint presentation and an interactive HTML website for documentation purposes.

## Requirements

### Functional Requirements
- Query all Conditional Access Policies from Entra ID using Microsoft Graph API
- Translate Application IDs (GUIDs) to human-readable names
- Translate User/Group IDs to display names
- Translate Location IDs to named locations
- Generate a PowerPoint (.pptx) slide deck with one policy per slide
- Generate an interactive HTML site with filtering/search capabilities
- No third-party library dependencies (native PowerShell only)

### Non-Functional Requirements
- Support for Microsoft Graph PowerShell SDK authentication OR direct REST API with device code flow
- Error handling for API failures
- Progress indication during execution
- Output files saved to configurable location

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Main Script (Orchestrator)                    │
│                  Get-ConditionalAccessReport.ps1                 │
└─────────────────────────┬───────────────────────────────────────┘
                          │
        ┌─────────────────┼─────────────────┐
        ▼                 ▼                 ▼
┌───────────────┐ ┌───────────────┐ ┌───────────────┐
│  Graph API    │ │   PowerPoint  │ │     HTML      │
│   Module      │ │   Generator   │ │   Generator   │
└───────────────┘ └───────────────┘ └───────────────┘
```

---

## Files to Create

### 1. Main Orchestrator Script
**File:** `Get-ConditionalAccessReport.ps1`

**Purpose:** Entry point that coordinates all modules and handles authentication

**Functions:**
- `Connect-ToGraph` - Authenticate to Microsoft Graph (device code or existing session)
- `Get-AllConditionalAccessPolicies` - Retrieve all CA policies
- `Start-ReportGeneration` - Orchestrate the full report generation

---

### 2. Graph API Helper Module
**File:** `Modules\GraphApiHelper.psm1`

**Purpose:** Handle all Microsoft Graph API interactions

**Functions:**
| Function | Description |
|----------|-------------|
| `Invoke-GraphRequest` | Generic Graph API request wrapper with pagination support |
| `Get-ConditionalAccessPolicies` | Retrieve all CA policies from `/identity/conditionalAccess/policies` |
| `Get-ApplicationDisplayName` | Translate App ID to friendly name via `/servicePrincipals` or `/applications` |
| `Get-UserDisplayName` | Get user display name from user ID |
| `Get-GroupDisplayName` | Get group display name from group ID |
| `Get-NamedLocationName` | Get named location display name |
| `Get-RoleDisplayName` | Get directory role display name |
| `Get-WellKnownApplications` | Return hashtable of well-known Microsoft app IDs |

**Well-Known Application ID Mapping (built-in):**
```powershell
$WellKnownApps = @{
    "00000002-0000-0000-c000-000000000000" = "Azure Active Directory Graph"
    "00000003-0000-0000-c000-000000000000" = "Microsoft Graph"
    "00000002-0000-0ff1-ce00-000000000000" = "Office 365 Exchange Online"
    "00000003-0000-0ff1-ce00-000000000000" = "Office 365 SharePoint Online"
    "00000004-0000-0ff1-ce00-000000000000" = "Office 365 Skype for Business"
    "797f4846-ba00-4fd7-ba43-dac1f8f63013" = "Windows Azure Service Management API"
    "c5393580-f805-4401-95e8-94b7a6ef2fc2" = "Office 365 Management APIs"
    "fc780465-2017-40d4-a0c5-307022471b92" = "My Apps"
    "d3590ed6-52b3-4102-aeff-aad2292ab01c" = "Microsoft Office"
    "de8bc8b5-d9f9-48b1-a8ad-b748da725064" = "Microsoft Graph Command Line Tools"
    "1fec8e78-bce4-4aaf-ab1b-5451cc387264" = "Microsoft Teams"
    "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" = "Microsoft Teams Web Client"
    "5e3ce6c0-2b1f-4285-8d4b-75ee78787346" = "Microsoft Teams - Device Admin Agent"
    # ... additional mappings
}
```

---

### 3. Policy Parser Module
**File:** `Modules\PolicyParser.psm1`

**Purpose:** Parse and transform raw CA policy JSON into structured objects

**Functions:**
| Function | Description |
|----------|-------------|
| `ConvertTo-PolicyObject` | Transform raw API response to structured object |
| `Get-PolicyConditionsSummary` | Extract and format conditions (users, apps, locations, platforms, etc.) |
| `Get-PolicyGrantControls` | Parse grant controls (MFA, compliant device, etc.) |
| `Get-PolicySessionControls` | Parse session controls (sign-in frequency, persistent browser, etc.) |
| `Format-IncludeExcludeList` | Format include/exclude lists for display |

**Policy Object Structure:**
```powershell
[PSCustomObject]@{
    Id                  = [string]
    DisplayName         = [string]
    State               = [string]  # enabled, disabled, enabledForReportingButNotEnforced
    CreatedDateTime     = [datetime]
    ModifiedDateTime    = [datetime]
    Conditions          = @{
        Users           = @{
            IncludeUsers     = @()  # Resolved display names
            ExcludeUsers     = @()
            IncludeGroups    = @()
            ExcludeGroups    = @()
            IncludeRoles     = @()
            ExcludeRoles     = @()
        }
        Applications    = @{
            IncludeApps      = @()  # Resolved friendly names
            ExcludeApps      = @()
            IncludeUserActions = @()
        }
        Platforms       = @{
            IncludePlatforms = @()
            ExcludePlatforms = @()
        }
        Locations       = @{
            IncludeLocations = @()
            ExcludeLocations = @()
        }
        ClientAppTypes  = @()
        SignInRiskLevels = @()
        UserRiskLevels  = @()
        DeviceStates    = @{}
    }
    GrantControls       = @{
        Operator        = [string]  # AND/OR
        BuiltInControls = @()       # mfa, compliantDevice, etc.
        CustomControls  = @()
        TermsOfUse      = @()
    }
    SessionControls     = @{
        SignInFrequency           = [string]
        PersistentBrowser         = [string]
        CloudAppSecurity          = [string]
        ApplicationEnforcedRestrictions = [bool]
        DisableResilienceDefaults = [bool]
    }
}
```

---

### 4. PowerPoint Generator Module
**File:** `Modules\PowerPointGenerator.psm1`

**Purpose:** Generate PowerPoint (.pptx) without third-party libraries using Open XML SDK via COM-free approach (Office Open XML direct manipulation)

**Technical Approach:**
Since we cannot use third-party libraries, we'll generate the .pptx file by:
1. Creating the Open XML structure manually (a .pptx is a ZIP file containing XML)
2. Using `System.IO.Compression` (built into .NET/PowerShell) to create the ZIP
3. Writing the required XML files for the presentation

**Functions:**
| Function | Description |
|----------|-------------|
| `New-PptxDocument` | Initialize a new PPTX file structure |
| `Add-TitleSlide` | Add the title/cover slide |
| `Add-PolicySlide` | Add a slide for a single CA policy |
| `Add-SummarySlide` | Add summary/statistics slide |
| `Save-PptxDocument` | Finalize and save the PPTX file |
| `New-SlideXml` | Generate XML for a slide |
| `ConvertTo-OpenXmlColor` | Convert color codes to Open XML format |

**Slide Layout Per Policy:**
```
┌────────────────────────────────────────────────────────────┐
│  [Policy Name]                              [State Badge]  │
├────────────────────────────────────────────────────────────┤
│                                                            │
│  USERS & GROUPS              │  APPLICATIONS               │
│  ─────────────────           │  ─────────────              │
│  Include:                    │  Include:                   │
│  • All Users                 │  • Office 365               │
│  Exclude:                    │  • Microsoft Teams          │
│  • Break Glass Accounts      │  Exclude:                   │
│  • Service Accounts          │  • None                     │
│                              │                             │
├──────────────────────────────┼─────────────────────────────┤
│  CONDITIONS                  │  GRANT CONTROLS             │
│  ─────────────────           │  ─────────────              │
│  • Platforms: All            │  Require ALL of:            │
│  • Locations: Any            │  • Multi-factor auth        │
│  • Client Apps: Browser,     │  • Compliant device         │
│    Mobile apps               │                             │
│  • Sign-in Risk: Medium+     │  SESSION CONTROLS           │
│                              │  ─────────────              │
│                              │  • Sign-in freq: 1 hour     │
└────────────────────────────────────────────────────────────┘
```

---

### 5. HTML Generator Module
**File:** `Modules\HtmlGenerator.psm1`

**Purpose:** Generate a self-contained, interactive HTML file with embedded CSS and JavaScript

**Functions:**
| Function | Description |
|----------|-------------|
| `New-HtmlReport` | Generate the complete HTML document |
| `Get-HtmlTemplate` | Return the base HTML template with embedded styles/scripts |
| `ConvertTo-PolicyHtmlCard` | Convert a policy object to an HTML card element |
| `Get-CssStyles` | Return embedded CSS (no external dependencies) |
| `Get-JavaScriptCode` | Return embedded JavaScript for interactivity |

**HTML Features:**
- **Search/Filter:** Real-time search across policy names and contents
- **State Filter:** Filter by Enabled/Disabled/Report-only
- **Sort Options:** Sort by name, state, modified date
- **Expand/Collapse:** Expandable policy cards for detailed view
- **Export:** Button to export filtered view to JSON
- **Dark/Light Mode:** Theme toggle
- **Print Friendly:** Print stylesheet for hard copies

**HTML Structure:**
```html
<!DOCTYPE html>
<html>
<head>
    <title>Conditional Access Policies Report</title>
    <style>/* Embedded CSS */</style>
</head>
<body>
    <header>
        <h1>Conditional Access Policies</h1>
        <div class="controls">
            <input type="search" id="search" placeholder="Search policies...">
            <select id="stateFilter">...</select>
            <select id="sortBy">...</select>
            <button id="themeToggle">Toggle Theme</button>
        </div>
        <div class="summary">
            <span>Total: X</span>
            <span>Enabled: X</span>
            <span>Disabled: X</span>
            <span>Report-Only: X</span>
        </div>
    </header>
    <main id="policies-container">
        <!-- Policy cards rendered here -->
    </main>
    <script>/* Embedded JavaScript */</script>
</body>
</html>
```

---

### 6. Configuration File
**File:** `config.json`

**Purpose:** Store configurable settings

```json
{
    "outputDirectory": "./Output",
    "pptxFileName": "ConditionalAccessPolicies.pptx",
    "htmlFileName": "ConditionalAccessPolicies.html",
    "includeDisabledPolicies": true,
    "theme": {
        "primaryColor": "#0078D4",
        "enabledColor": "#107C10",
        "disabledColor": "#A80000",
        "reportOnlyColor": "#FFB900"
    },
    "graphApiVersion": "v1.0",
    "maxConcurrentApiCalls": 5
}
```

---

## File Structure

```
ConditionalAccessDocumenter/
│
├── Get-ConditionalAccessReport.ps1    # Main entry point
├── config.json                         # Configuration file
├── DESIGN.md                          # This document
├── README.md                          # User documentation
│
├── Modules/
│   ├── GraphApiHelper.psm1            # MS Graph API interactions
│   ├── PolicyParser.psm1              # Policy data transformation
│   ├── PowerPointGenerator.psm1       # PPTX generation
│   └── HtmlGenerator.psm1             # HTML report generation
│
├── Templates/
│   ├── slide-template.xml             # Base slide XML template
│   └── html-template.html             # Base HTML template (optional)
│
└── Output/                            # Generated reports (created at runtime)
    ├── ConditionalAccessPolicies.pptx
    └── ConditionalAccessPolicies.html
```

---

## Authentication Flow

### Option 1: Interactive Device Code Flow (No dependencies)
```powershell
# Uses direct REST API with device code authentication
$authUrl = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
$tokenUrl = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
$scope = "https://graph.microsoft.com/.default"
```

### Option 2: Microsoft Graph PowerShell SDK (If available)
```powershell
# Checks if Microsoft.Graph module is available
Connect-MgGraph -Scopes "Policy.Read.All", "Application.Read.All", "User.Read.All", "Group.Read.All"
```

### Required Permissions
| Permission | Type | Purpose |
|------------|------|---------|
| `Policy.Read.All` | Delegated/Application | Read CA policies |
| `Application.Read.All` | Delegated/Application | Resolve app IDs to names |
| `Directory.Read.All` | Delegated/Application | Read users, groups, roles |
| `User.Read.All` | Delegated | Read user display names |
| `Group.Read.All` | Delegated | Read group display names |

---

## Execution Flow

```
1. Start Script
       │
2. Load Configuration
       │
3. Authenticate to MS Graph
       │
4. Fetch All CA Policies ──────────────┐
       │                               │
5. For Each Policy:                    │ (Parallel where possible)
   ├── Resolve User/Group IDs          │
   ├── Resolve Application IDs         │
   ├── Resolve Location IDs            │
   └── Resolve Role IDs                │
       │                               │
6. Build Structured Policy Objects ◄───┘
       │
       ├──────────────────┐
       │                  │
7. Generate PPTX     Generate HTML
       │                  │
       └──────────────────┤
                          │
8. Save Output Files
       │
9. Display Summary & Exit
```

---

## API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /identity/conditionalAccess/policies` | Get all CA policies |
| `GET /servicePrincipals?$filter=appId eq '{id}'` | Get app display name |
| `GET /applications?$filter=appId eq '{id}'` | Fallback for app name |
| `GET /users/{id}?$select=displayName` | Get user display name |
| `GET /groups/{id}?$select=displayName` | Get group display name |
| `GET /directoryRoles?$filter=roleTemplateId eq '{id}'` | Get role display name |
| `GET /identity/conditionalAccess/namedLocations/{id}` | Get named location |

---

## Error Handling Strategy

| Scenario | Handling |
|----------|----------|
| Authentication failure | Display error, exit with code 1 |
| API rate limiting (429) | Exponential backoff retry (3 attempts) |
| Missing permissions | Display required permissions, exit |
| Failed ID resolution | Use original ID with "[Unresolved]" prefix |
| Network timeout | Retry up to 3 times with increasing delay |
| File write failure | Display error, attempt alternate location |

---

## Output Examples

### PowerPoint Slide Example
Each slide contains:
- Policy name as title
- State indicator (color-coded badge)
- Four quadrant layout: Users, Apps, Conditions, Controls
- Footer with last modified date

### HTML Card Example
Interactive cards with:
- Click to expand/collapse
- Color-coded state badges
- Hover tooltips for additional info
- Copy policy ID button

---

## Implementation Phases

### Phase 1: Core Infrastructure
- [ ] Main script skeleton
- [ ] Configuration loading
- [ ] Graph API authentication (device code flow)
- [ ] Basic API request wrapper

### Phase 2: Data Retrieval
- [ ] CA policy retrieval with pagination
- [ ] ID to name resolution for apps, users, groups
- [ ] Caching layer for resolved names
- [ ] Policy object transformation

### Phase 3: PowerPoint Generation
- [ ] PPTX structure creation (Open XML)
- [ ] Title slide generation
- [ ] Policy slide generation
- [ ] Summary slide generation

### Phase 4: HTML Generation
- [ ] HTML template with embedded CSS
- [ ] Policy card generation
- [ ] JavaScript interactivity
- [ ] Search/filter functionality

### Phase 5: Polish
- [ ] Error handling refinement
- [ ] Progress indicators
- [ ] Logging
- [ ] Documentation

---

## Files to Create Summary

| # | File | Type | Purpose |
|---|------|------|---------|
| 1 | `Get-ConditionalAccessReport.ps1` | Script | Main entry point |
| 2 | `Modules/GraphApiHelper.psm1` | Module | Graph API functions |
| 3 | `Modules/PolicyParser.psm1` | Module | Policy data parsing |
| 4 | `Modules/PowerPointGenerator.psm1` | Module | PPTX file generation |
| 5 | `Modules/HtmlGenerator.psm1` | Module | HTML report generation |
| 6 | `config.json` | Config | Settings file |
| 7 | `README.md` | Doc | User documentation |

---

## Usage

```powershell
# Basic usage - interactive authentication
.\Get-ConditionalAccessReport.ps1

# Specify output directory
.\Get-ConditionalAccessReport.ps1 -OutputPath "C:\Reports"

# Generate only HTML (skip PowerPoint)
.\Get-ConditionalAccessReport.ps1 -HtmlOnly

# Generate only PowerPoint (skip HTML)
.\Get-ConditionalAccessReport.ps1 -PptxOnly

# Specify tenant ID for multi-tenant scenarios
.\Get-ConditionalAccessReport.ps1 -TenantId "contoso.onmicrosoft.com"
```

---

## Dependencies

**Required (Built-in to Windows/PowerShell):**
- PowerShell 5.1+ or PowerShell 7+
- `System.IO.Compression` (.NET)
- `System.Xml` (.NET)
- `Invoke-RestMethod` (PowerShell)

**Optional (Enhances functionality if present):**
- Microsoft.Graph PowerShell SDK (for simpler auth)

**Explicitly NOT Required:**
- ImportExcel module
- Any Office COM automation
- Any third-party REST clients
- Any external JavaScript libraries
