# Conditional Access Policy Documenter

A PowerShell-based tool that documents Microsoft Entra ID Conditional Access Policies by generating both a PowerPoint slide deck and a multi-view interactive HTML report — with no third-party dependencies.

## Features

- **No Third-Party Dependencies**: Uses only native PowerShell and .NET capabilities
- **Microsoft Graph Integration**: Supports both interactive (device code) and unattended (service principal) authentication
- **ID Resolution**: Translates application, user, group, and role IDs to friendly display names
- **PowerPoint Generation**: Creates professional slide decks using Open XML format
- **Multi-View HTML Reports**: Self-contained HTML with three analysis views (see below)

## Requirements

- PowerShell 5.1 or PowerShell 7+
- Internet connectivity for Microsoft Graph API
- Appropriate permissions in Entra ID (see Permissions section)

## Quick Start

1. Clone or download this repository
2. Open PowerShell and navigate to the project directory
3. Run the script:

```powershell
.\Get-ConditionalAccessReport.ps1
```

4. Follow the device code authentication prompt
5. Find your reports in the `Output` folder

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

## Required Permissions

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

The self-contained HTML report includes three analysis views:

#### View 1 — Policy Cards
Searchable, filterable list of all policies. Each card shows full policy details including users, apps, conditions, grant controls, and session controls. Supports real-time search, state filtering, and dark mode.

#### View 2 — Coverage Matrix
A grid showing which user populations are covered by which applications. Red cells highlight gaps (combinations with no policy). Click any cell to see which policies cover that intersection.

#### View 3 — Overlap & Conflict Analyzer
Pairwise comparison of all policies. Identifies:
- **Conflicts**: Policies that apply to the same users and apps but enforce different controls
- **Redundancies**: Policies where one is a subset of another (may be unnecessary)

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
├── README.md                          # This file
├── DESIGN.md                          # Technical design document
├── Modules/
│   ├── GraphApiHelper.psm1           # Microsoft Graph API functions
│   ├── PolicyParser.psm1             # Policy data transformation
│   ├── PowerPointGenerator.psm1      # PPTX generation (Open XML)
│   └── HtmlGenerator.psm1            # HTML report generation (multi-view SPA)
└── Output/                            # Generated reports (created at runtime)
```

## How It Works

1. **Authentication**: OAuth 2.0 via device code (interactive) or client credentials (service principal)
2. **Data Retrieval**: Fetches all Conditional Access policies via Microsoft Graph API with pagination
3. **ID Resolution**: Translates GUIDs to display names using an in-memory cache for performance
4. **Report Generation**:
   - PowerPoint: Constructs Open XML structure directly (PPTX is a ZIP of XML files)
   - HTML: Generates a self-contained file — all analysis runs client-side against embedded JSON

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

## License

This project is provided as-is for documentation purposes.

## Acknowledgments

- Built using Microsoft Graph API
- PowerPoint generation uses Office Open XML format
- No third-party libraries required
