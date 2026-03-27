# Conditional Access Policy Documenter

A PowerShell-based tool that documents Microsoft Entra ID Conditional Access Policies by generating both a PowerPoint slide deck and a multi-view interactive HTML report — with no third-party dependencies.

## Features

- **No Third-Party Dependencies**: Uses only native PowerShell and .NET capabilities
- **Online Mode**: Connects to Microsoft Graph API — supports device code (interactive) and service principal (unattended) authentication
- **Offline Mode**: Generate reports from an exported JSON file — no internet connection or credentials required
- **ID Resolution**: Translates application, user, group, and role IDs to friendly display names; custom names supported via a mapping file
- **PowerPoint Generation**: Creates professional slide decks using Open XML format
- **Multi-View HTML Reports**: Self-contained HTML with four analysis views (see below)

## Requirements

- PowerShell 5.1 or PowerShell 7+
- **Online mode**: Internet connectivity and appropriate permissions in Entra ID (see Permissions section)
- **Offline mode**: No internet or credentials required — just an exported policies JSON file

## Quick Start

### Online (connect to your tenant)

1. Clone or download this repository
2. Open PowerShell and navigate to the project directory
3. Run the script:

```powershell
.\Get-ConditionalAccessReport.ps1
```

4. Follow the device code authentication prompt
5. Find your reports in the `Output` folder

### Offline (from exported JSON)

1. Export your policies from the Azure Portal (Conditional Access blade → Export) or via Graph Explorer
2. Run:

```powershell
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json"
```

3. Optionally provide a name mapping file to resolve custom app/group names:

```powershell
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json" -NameMappingPath ".\NameMapping.json"
```

## Usage

### Interactive Authentication (Device Code)

```powershell
# Run with default settings — prompts for login in browser
.\Get-ConditionalAccessReport.ps1

# Specify tenant
.\Get-ConditionalAccessReport.ps1 -TenantId "contoso.onmicrosoft.com"
```

### Service Principal Authentication (Unattended)

```powershell
# Client credentials — no interactive prompt
.\Get-ConditionalAccessReport.ps1 `
    -TenantId "contoso.onmicrosoft.com" `
    -ClientId "your-app-client-id" `
    -ClientSecret "your-client-secret"

# Force a specific auth method explicitly
.\Get-ConditionalAccessReport.ps1 -AuthMethod DeviceCode
.\Get-ConditionalAccessReport.ps1 -AuthMethod ClientCredentials -ClientId "..." -ClientSecret "..."
```

### Offline Mode

```powershell
# Basic — uses built-in well-known app/role names only
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json"

# With custom name mapping for apps, groups, users, and locations not in the built-in tables
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json" -NameMappingPath ".\NameMapping.json"

# Offline, HTML report only
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json" -HtmlOnly
```

The exported JSON can come from:
- The Azure Portal → Conditional Access → Export
- Graph Explorer: `GET https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies`
- PowerShell: `Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Headers @{Authorization="Bearer $token"}`

Both `{"value":[...]}` wrapper format and bare array format are supported.

### Output Options

```powershell
# Generate only HTML report
.\Get-ConditionalAccessReport.ps1 -HtmlOnly

# Generate only PowerPoint
.\Get-ConditionalAccessReport.ps1 -PptxOnly

# Custom output directory
.\Get-ConditionalAccessReport.ps1 -OutputPath "C:\Reports"

# Exclude disabled policies
.\Get-ConditionalAccessReport.ps1 -IncludeDisabled $false
```

## Offline Name Mapping

When using offline mode, GUIDs for custom (non-Microsoft) applications, groups, users, and named locations cannot be resolved without Graph API. Provide a mapping file to show friendly names instead of `[Unknown Application]` / `[Unknown Group]` placeholders.

### Step 1 — Find which GUIDs need mapping

Run offline without a mapping file first:

```powershell
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json" -HtmlOnly
```

Open the HTML report and look for `[Unknown Application]`, `[Unknown Group]`, etc. These are the entries you need to map. The raw GUIDs for those entries are in your `policies.json` — search the file for the relevant section (e.g. `includeApplications`, `includeGroups`) to find them.

### Step 2 — Look up the friendly names

Find the GUID → name mapping in the Azure Portal:

| Entity | Where to find the Object ID / App ID |
|--------|--------------------------------------|
| Applications | **Azure Portal** → Enterprise Applications → select app → **Application ID** (under Properties) |
| Groups | **Azure Portal** → Groups → select group → **Object ID** (under Overview) |
| Users | **Azure Portal** → Users → select user → **Object ID** (under Overview) |
| Named Locations | **Azure Portal** → Conditional Access → Named Locations → select location → the GUID is in the browser URL |

### Step 3 — Create your mapping file

Copy `NameMapping.example.json` to `NameMapping.json` and fill in your GUIDs:

```json
{
    "applications": {
        "12345678-1234-1234-1234-123456789abc": "My Custom Enterprise App"
    },
    "groups": {
        "aaaabbbb-cccc-dddd-eeee-ffffaaaabbbb": "Finance Department"
    },
    "users": {
        "abcdef01-2345-6789-abcd-ef0123456789": "Break Glass Account"
    },
    "namedLocations": {
        "a1b2c3d4-e5f6-7890-a1b2-c3d4e5f67890": "Corporate Headquarters"
    }
}
```

Then re-run with `-NameMappingPath`:

```powershell
.\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath ".\policies.json" -NameMappingPath ".\NameMapping.json"
```

`NameMapping.json` is excluded from git (see `.gitignore`) since it contains tenant-specific identifiers. The example file is tracked as a template.

**Name resolution priority in offline mode:**
1. Built-in well-known Microsoft apps and roles (~120 entries, always available)
2. Entries in your `NameMapping.json`
3. `[Unknown Application]` / `[Unknown Group]` / etc. — no raw GUIDs are ever shown

## Required Permissions (Online Mode)

### Delegated (Device Code / Interactive)

| Permission | Purpose |
|------------|---------|
| `Policy.Read.All` | Read Conditional Access policies |
| `Application.Read.All` | Resolve application IDs to names |
| `Directory.Read.All` | Read users, groups, and roles |

### Application (Service Principal / Client Credentials)

| Permission | Type | Purpose |
|------------|------|---------|
| `Policy.Read.All` | Application | Read Conditional Access policies |
| `Application.Read.All` | Application | Resolve application IDs to names |
| `Directory.Read.All` | Application | Read users, groups, and roles |

> **Note:** Application permissions require admin consent in your Entra ID tenant.

## Output Files

### PowerPoint (.pptx)

Each presentation includes:
- **Title Slide**: Presentation cover
- **Policy Slides**: One slide per policy with:
  - Policy name and state badge (color-coded)
  - Users & Groups section
  - Applications section
  - Conditions section
  - Access Controls section
- **Summary Slide**: Statistics overview

### HTML Report (Multi-View)

The self-contained HTML report includes four analysis views:

#### View 1 — Policy Cards
Searchable, filterable list of all policies. Each card shows full policy details including users, apps, conditions, grant controls, and session controls. Supports real-time search, state filtering, and dark mode.

#### View 2 — Coverage Matrix
A grid showing which user populations are covered by which applications. Red cells highlight gaps (combinations with no policy). Click any cell to see which policies cover that intersection.

#### View 3 — Overlap & Conflict Analyzer
Pairwise comparison of all policies. Identifies:
- **Conflicts**: Policies that apply to the same users and apps but enforce different controls
- **Redundancies**: Policies where one is a subset of another (may be unnecessary)

#### View 4 — Application Lookup
Search for any application to see every policy that covers it, along with coverage statistics.

## Configuration

Customize the tool by editing `config.json`:

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
    }
}
```

## Project Structure

```
ConditionalAccessDocumenter/
├── Get-ConditionalAccessReport.ps1   # Main entry point
├── config.json                        # Configuration file
├── NameMapping.example.json           # Template for offline name resolution
├── README.md                          # This file
├── DESIGN.md                          # Technical design document
├── .github/
│   └── workflows/
│       └── security.yml               # Secret detection + code scanning CI
├── Modules/
│   ├── GraphApiHelper.psm1           # Microsoft Graph API functions + offline mode
│   ├── PolicyParser.psm1             # Policy data transformation
│   ├── PowerPointGenerator.psm1      # PPTX generation (Open XML)
│   └── HtmlGenerator.psm1            # HTML report generation (multi-view SPA)
└── Output/                            # Generated reports (created at runtime)
```

## How It Works

### Online Mode
1. **Authentication**: OAuth 2.0 via device code (interactive) or client credentials (service principal)
2. **Data Retrieval**: Fetches all Conditional Access policies via Microsoft Graph API with pagination
3. **ID Resolution**: Translates GUIDs to display names using an in-memory cache; falls back to well-known tables then API queries
4. **Report Generation**: PowerPoint (Open XML) and HTML (self-contained, client-side analysis)

### Offline Mode
1. **Load Policies**: Reads a previously exported JSON file (no authentication)
2. **Load Name Mappings**: If a mapping file is provided, pre-populates the name resolution cache
3. **ID Resolution**: Checks well-known tables → mapping file → friendly placeholder (no GUIDs shown)
4. **Report Generation**: Same pipeline as online mode — identical output format

## Security

On every push and pull request to `main`, GitHub Actions runs two automated security checks:

| Check | Tool | What it catches |
|-------|------|----------------|
| Secret Detection | [Gitleaks](https://github.com/gitleaks/gitleaks) | Hardcoded credentials, tokens, and API keys in code or history |
| Code Analysis | [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) | PowerShell security anti-patterns and code quality issues |

Results from code analysis are uploaded to GitHub Code Scanning and appear inline in pull requests.

## Troubleshooting

### Authentication Issues

- **Device code**: Ensure you have the required delegated permissions and are logging in as a user with access to CA policies
- **Service principal**: Verify application permissions are granted and admin-consented; check that the client secret has not expired
- Try running PowerShell as Administrator if you encounter module loading issues

### Missing Permissions Error

Ask your Global Administrator to grant and consent to:
- `Policy.Read.All`
- `Application.Read.All`
- `Directory.Read.All`

### Rate Limiting

The tool includes automatic retry logic with exponential backoff. Large tenants with many policies may take a few minutes to process.

### Offline Mode — Names Not Resolving

If you see `[Unknown Application]` or similar placeholders:
1. Check if the app is a Microsoft first-party service — it should be in the built-in tables already
2. Add the GUID → name mapping to `NameMapping.json` and pass it via `-NameMappingPath`

## License

This project is provided as-is for documentation purposes.

## Acknowledgments

- Built using Microsoft Graph API
- PowerPoint generation uses Office Open XML format
- No third-party libraries required
