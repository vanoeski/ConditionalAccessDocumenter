# Conditional Access Policy Documenter

A PowerShell-based tool that documents Microsoft Entra ID (Azure AD) Conditional Access Policies by generating both PowerPoint presentations and interactive HTML reports.

## Features

- **No Third-Party Dependencies**: Uses only native PowerShell and .NET capabilities
- **Microsoft Graph Integration**: Authenticates via device code flow
- **ID Resolution**: Translates application, user, group, and role IDs to friendly names
- **PowerPoint Generation**: Creates professional slide decks using Open XML format
- **Interactive HTML Reports**: Self-contained HTML with search, filter, dark mode, and JSON export
- **Comprehensive Coverage**: Documents users, groups, applications, conditions, and access controls

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

### Basic Usage

```powershell
# Run with default settings
.\Get-ConditionalAccessReport.ps1

# Specify tenant
.\Get-ConditionalAccessReport.ps1 -TenantId "contoso.onmicrosoft.com"

# Custom output directory
.\Get-ConditionalAccessReport.ps1 -OutputPath "C:\Reports"
```

### Output Options

```powershell
# Generate only HTML report
.\Get-ConditionalAccessReport.ps1 -HtmlOnly

# Generate only PowerPoint
.\Get-ConditionalAccessReport.ps1 -PptxOnly

# Exclude disabled policies
.\Get-ConditionalAccessReport.ps1 -IncludeDisabled $false
```

## Required Permissions

The following Microsoft Graph permissions are required:

| Permission | Type | Purpose |
|------------|------|---------|
| `Policy.Read.All` | Delegated | Read Conditional Access policies |
| `Application.Read.All` | Delegated | Resolve application IDs to names |
| `Directory.Read.All` | Delegated | Read users, groups, and roles |

When you authenticate, you'll be prompted to consent to these permissions.

## Output Files

### PowerPoint (.pptx)

Each presentation includes:
- **Title Slide**: Presentation cover
- **Policy Slides**: One slide per policy with:
  - Policy name and state badge
  - Users & Groups section
  - Applications section
  - Conditions section
  - Access Controls section
- **Summary Slide**: Statistics overview

### HTML Report

The interactive HTML report features:
- **Search**: Real-time filtering across all policy content
- **State Filter**: Filter by Enabled/Disabled/Report-only
- **Sort Options**: Sort by name or state
- **Expand/Collapse**: Click to expand policy details
- **Dark Mode**: Theme toggle for comfortable viewing
- **Export JSON**: Download filtered policies as JSON
- **Print Friendly**: Optimized print stylesheet

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
│   └── HtmlGenerator.psm1            # HTML report generation
└── Output/                            # Generated reports (created at runtime)
```

## How It Works

1. **Authentication**: Uses OAuth 2.0 device code flow to authenticate to Microsoft Graph
2. **Data Retrieval**: Fetches all Conditional Access policies via Graph API
3. **ID Resolution**: Translates GUIDs to display names using caching for performance
4. **Report Generation**:
   - PowerPoint: Creates Open XML structure directly (PPTX is a ZIP of XML files)
   - HTML: Generates self-contained file with embedded CSS/JavaScript

## Troubleshooting

### Authentication Issues

- Ensure you have the required permissions in your Entra ID tenant
- Try running PowerShell as Administrator
- Check your network connectivity

### Missing Permissions Error

If you receive permission errors, ask your Global Administrator to grant consent for:
- `Policy.Read.All`
- `Application.Read.All`
- `Directory.Read.All`

### Rate Limiting

The tool includes automatic retry logic with exponential backoff for API rate limiting. Large tenants with many policies may take longer to process.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

This project is provided as-is for documentation purposes.

## Acknowledgments

- Built using Microsoft Graph API
- PowerPoint generation uses Office Open XML format
- No third-party libraries required
