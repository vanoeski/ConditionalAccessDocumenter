# Conditional Access Policy Documenter — Design Document

## Overview

A PowerShell-based solution that queries Microsoft Entra ID Conditional Access Policies via Microsoft Graph API and generates:
- A PowerPoint slide deck (one policy per slide, Open XML format, no Office dependency)
- A self-contained multi-view HTML report (all analysis client-side, no external JS libraries)

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Main Script (Orchestrator)                    │
│                  Get-ConditionalAccessReport.ps1                 │
└─────────────────────────┬───────────────────────────────────────┘
                          │
        ┌─────────────────┼──────────────────┐
        ▼                 ▼                  ▼
┌───────────────┐ ┌───────────────┐ ┌───────────────┐
│  GraphApi     │ │  PowerPoint   │ │     HTML      │
│  Helper       │ │  Generator    │ │   Generator   │
├───────────────┤ └───────────────┘ └───────────────┘
│  PolicyParser │
└───────────────┘
```

---

## Modules

### `Get-ConditionalAccessReport.ps1` — Main Orchestrator

Entry point. Loads modules, authenticates, fetches policies, processes them, and calls both generators.

**Parameters:**

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-TenantId` | `"common"` | Entra ID tenant ID or domain |
| `-ClientId` | *(built-in public app ID)* | App registration client ID |
| `-ClientSecret` | *(none)* | Client secret — triggers service principal auth |
| `-AuthMethod` | *(auto-detect)* | `DeviceCode` or `ClientCredentials` |
| `-OutputPath` | `".\Output"` | Directory for generated reports |
| `-HtmlOnly` | `$false` | Skip PowerPoint generation |
| `-PptxOnly` | `$false` | Skip HTML generation |
| `-IncludeDisabled` | `$true` | Include disabled policies |
| `-ConfigPath` | `".\config.json"` | Path to configuration file |

---

### `Modules/GraphApiHelper.psm1`

Handles all Microsoft Graph API interaction: authentication, paginated requests, ID resolution, and caching.

**Authentication:**

Two flows are supported. Auto-detection: if `-ClientSecret` is provided, `ClientCredentials` is used; otherwise `DeviceCode`.

| Flow | Grant Type | Use Case |
|------|-----------|---------|
| Device Code | `urn:ietf:params:oauth:grant-type:device_code` | Interactive, delegated permissions |
| Client Credentials | `client_credentials` | Unattended, application permissions |

Token endpoint: `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`

**Exported Functions:**

| Function | Description |
|----------|-------------|
| `Connect-Graph` | Authenticate and store access token in module scope |
| `Disconnect-Graph` | Clear stored token |
| `Invoke-GraphRequest` | Generic REST wrapper with pagination (`@odata.nextLink`) and retry |
| `Get-ConditionalAccessPolicies` | Fetch all policies from `/identity/conditionalAccess/policies` |
| `Get-ApplicationDisplayName` | Resolve app ID → name via `/servicePrincipals` then `/applications` |
| `Get-UserDisplayName` | Resolve user ID → display name |
| `Get-GroupDisplayName` | Resolve group ID → display name |
| `Get-RoleDisplayName` | Resolve role template ID → display name |
| `Get-NamedLocationName` | Resolve named location ID → display name |

**Caching:** All resolved names are stored in module-scoped hashtables (`$script:AppCache`, `$script:UserCache`, etc.) to avoid redundant API calls.

**Well-Known Application IDs:** A built-in hashtable maps Microsoft first-party app GUIDs to friendly names without API calls.

---

### `Modules/PolicyParser.psm1`

Transforms raw Graph API policy JSON into structured PowerShell objects with all IDs resolved to display names.

**Exported Functions:**

| Function | Description |
|----------|-------------|
| `ConvertTo-PolicyObject` | Main transform: raw policy → structured PSCustomObject |
| `Get-PolicyConditionsSummary` | Extract conditions (users, apps, platforms, locations, risks) |
| `Get-PolicyGrantControls` | Parse grant controls (MFA, compliant device, etc.) |
| `Get-PolicySessionControls` | Parse session controls (sign-in freq, persistent browser, etc.) |

**Output Object Shape:**

```powershell
[PSCustomObject]@{
    Id          = [string]
    DisplayName = [string]
    State       = [string]   # "Enabled", "Disabled", "Report-Only"
    StateRaw    = [string]   # Raw API value
    Conditions  = @{
        Users         = @{
            IncludeUsers  = [string[]]  # Resolved display names
            ExcludeUsers  = [string[]]
            IncludeGroups = [string[]]
            ExcludeGroups = [string[]]
            IncludeRoles  = [string[]]
            ExcludeRoles  = [string[]]
        }
        Applications  = @{
            IncludeApps       = [string[]]
            ExcludeApps       = [string[]]
            IncludeUserActions = [string[]]
        }
        Platforms     = @{ IncludePlatforms = [string[]]; ExcludePlatforms = [string[]] }
        Locations     = @{ IncludeLocations = [string[]]; ExcludeLocations = [string[]] }
        ClientAppTypes   = [string[]]
        SignInRiskLevels = [string[]]
        UserRiskLevels   = [string[]]
        Devices          = [hashtable]
    }
    GrantControls  = @{
        Operator        = [string]   # "AND" or "OR"
        BuiltInControls = [string[]] # "mfa", "compliantDevice", etc.
        CustomControls  = [string[]]
        TermsOfUse      = [string[]]
    }
    SessionControls = @{
        SignInFrequency                  = [string]
        PersistentBrowser               = [string]
        CloudAppSecurity                = [string]
        ApplicationEnforcedRestrictions = [bool]
        DisableResilienceDefaults       = [bool]
    }
}
```

---

### `Modules/PowerPointGenerator.psm1`

Generates `.pptx` files using Open XML (a PPTX is a ZIP archive of XML files). No Office installation required.

**Approach:** Uses `System.IO.Compression.ZipArchive` to build the ZIP structure in memory, writes XML parts directly. Key quirk: `[Content_Types].xml` must be written with `-LiteralPath` because PowerShell's `-FilePath` treats `[` as a glob wildcard.

**Exported Functions:**

| Function | Description |
|----------|-------------|
| `New-PptxDocument` | Initialize document structure in memory |
| `Add-TitleSlide` | Add cover slide |
| `Add-PolicySlide` | Add one slide per policy |
| `Add-SummarySlide` | Add statistics summary slide |
| `Save-PptxDocument` | Write the ZIP/XML structure to disk |
| `Set-PptxTheme` | Configure color theme |

**Slide Layout:**

```
┌─────────────────────────────────────────────────────────┐
│  [Policy Name]                           [State Badge]  │
├──────────────────────────┬──────────────────────────────┤
│  USERS & GROUPS          │  APPLICATIONS                │
│  Include: All Users      │  Include: Office 365         │
│  Exclude: Break Glass    │  Exclude: (none)             │
├──────────────────────────┼──────────────────────────────┤
│  CONDITIONS              │  GRANT CONTROLS              │
│  Platforms: All          │  Require ALL of:             │
│  Locations: Any          │  • MFA                       │
│  Client Apps: Browser    │  • Compliant device          │
│  Sign-in Risk: Medium+   │                              │
│                          │  SESSION CONTROLS            │
│                          │  Sign-in freq: 1 hour        │
└──────────────────────────┴──────────────────────────────┘
```

---

### `Modules/HtmlGenerator.psm1`

Generates a self-contained HTML file. All policy data is embedded as JSON; all analysis runs client-side in JavaScript. No external CSS or JS libraries.

**Exported Functions:**

| Function | Description |
|----------|-------------|
| `New-HtmlReport` | Generate and save the complete HTML file |
| `Set-HtmlTheme` | Configure color theme before generation |

**Important implementation detail:** `ConvertTo-Json` in PowerShell silently unwraps single-element arrays to bare scalars at every nesting level. Mitigated with:
- `@($Policies) | ConvertTo-Json` — force top-level array
- `const policiesData = [].concat($PoliciesJson)` — JS-side array guard
- `function arr(x) { return x == null ? [] : [].concat(x); }` — used on every nested field access

#### HTML Views

**View 1 — Policy Cards**

Searchable, filterable list of all policies. Features:
- Real-time search across all policy content
- Filter by state (Enabled / Disabled / Report-Only)
- Sort by name or state
- Expand/collapse cards
- Dark mode toggle
- JSON export

**View 2 — Coverage Matrix**

Grid of user population rows × application columns. Each cell shows whether a policy covers that combination. Red cells highlight gaps. Clicking a cell shows a popover listing which policies apply.

Building the matrix:
1. Collect all unique user populations (`IncludeUsers`, `IncludeGroups`, `IncludeRoles`) across all policies
2. Collect all unique applications (`IncludeApps`) across all policies
3. For each cell `[userPop][app]`, find policies whose scope intersects both

**View 3 — Overlap & Conflict Analyzer**

Pairwise comparison of all enabled policies. For each pair:
1. Compute user scope intersection (sets of users/groups/roles or "All")
2. Compute app scope intersection
3. If both overlap — flag as **Conflict** (different controls) or **Redundant** (same controls, one is a subset)

---

## Authentication Flows

### Device Code (Interactive)

```
1. POST /oauth2/v2.0/devicecode  →  device_code, user_code, verification_uri
2. Display user_code + verification_uri to user
3. Poll /oauth2/v2.0/token until authorized or expired
4. Store access_token + expiry in module scope
```

### Client Credentials (Service Principal)

```
1. POST /oauth2/v2.0/token
   grant_type=client_credentials
   client_id={id}
   client_secret={secret}
   scope=https://graph.microsoft.com/.default
2. Store access_token + expiry in module scope
```

Required application permissions for service principal:
- `Policy.Read.All`
- `Application.Read.All`
- `Directory.Read.All`

---

## API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /identity/conditionalAccess/policies` | All CA policies (paginated) |
| `GET /servicePrincipals?$filter=appId eq '{id}'` | App display name (primary) |
| `GET /applications?$filter=appId eq '{id}'` | App display name (fallback) |
| `GET /users/{id}?$select=displayName` | User display name |
| `GET /groups/{id}?$select=displayName` | Group display name |
| `GET /directoryRoles?$filter=roleTemplateId eq '{id}'` | Role display name |
| `GET /identity/conditionalAccess/namedLocations/{id}` | Named location display name |

---

## Error Handling

| Scenario | Handling |
|----------|----------|
| Auth failure | Display error, exit with code 1 |
| API 429 (rate limit) | Exponential backoff, up to 3 retries |
| Failed ID resolution | Use original GUID with `[Unresolved]` prefix |
| Network timeout | Retry up to 3 times with increasing delay |
| Single-element array serialization | `@()` operator + `arr()` JS helper |
| `[Content_Types].xml` wildcard expansion | `-LiteralPath` in `Out-File` |

---

## File Structure

```
ConditionalAccessDocumenter/
├── Get-ConditionalAccessReport.ps1    # Main entry point
├── config.json                         # Configuration
├── README.md                           # User documentation
├── DESIGN.md                           # This document
├── Modules/
│   ├── GraphApiHelper.psm1             # Graph API + auth
│   ├── PolicyParser.psm1               # Data transformation
│   ├── PowerPointGenerator.psm1        # PPTX generation (Open XML)
│   └── HtmlGenerator.psm1             # HTML multi-view report
└── Output/                             # Generated reports (runtime)
    ├── ConditionalAccessPolicies.pptx
    └── ConditionalAccessPolicies.html
```

---

## Dependencies

**Built into Windows / PowerShell — no installation required:**
- `System.IO.Compression` (.NET) — PPTX ZIP construction
- `System.Web` (.NET) — HTML encoding
- `Invoke-RestMethod` (PowerShell) — Graph API calls

**Explicitly not used:**
- Microsoft.Graph PowerShell SDK
- ImportExcel or any Office COM automation
- Any external JavaScript or CSS libraries
