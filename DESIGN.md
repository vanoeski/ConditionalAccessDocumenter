# Conditional Access Policy Documenter — Design Document

## Overview

A PowerShell-based solution that generates CA policy reports in two modes:

- **Online**: Queries Microsoft Entra ID via Microsoft Graph API, resolves GUIDs to display names in real time
- **Offline**: Reads a previously exported JSON file, resolves names from built-in tables and an optional user-supplied mapping file — no credentials or internet required

Output: a PowerPoint slide deck (Open XML, no Office dependency) and a self-contained multi-view HTML report (all analysis client-side, no external JS libraries).

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                   Main Script (Orchestrator)                  │
│                 Get-ConditionalAccessReport.ps1               │
│                                                               │
│   Online path: Connect-Graph → Get-ConditionalAccessPolicies  │
│   Offline path: Load JSON → Initialize-OfflineCacheFromMapping│
└──────────────────────┬───────────────────────────────────────┘
                       │
       ┌───────────────┼──────────────────┐
       ▼               ▼                  ▼
┌─────────────┐ ┌─────────────┐ ┌──────────────┐
│  GraphApi   │ │  PowerPoint │ │     HTML     │
│  Helper     │ │  Generator  │ │   Generator  │
├─────────────┤ └─────────────┘ └──────────────┘
│ PolicyParser│
└─────────────┘
```

---

## Modules

### `Get-ConditionalAccessReport.ps1` — Main Orchestrator

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
| `-OfflineMode` | `$false` | Skip Graph API; load policies from JSON file |
| `-PoliciesJsonPath` | *(none)* | Path to exported CA policies JSON (required with `-OfflineMode`) |
| `-NameMappingPath` | *(none)* | Path to GUID→name mapping JSON for offline resolution |

**Flow — Online:**
1. Load config → import modules
2. `Connect-Graph` (device code or client credentials)
3. `Get-ConditionalAccessPolicies` (paginated Graph API)
4. `ConvertTo-PolicyObject -ResolveNames` for each policy (API-backed resolution)
5. Generate PowerPoint and/or HTML → `Disconnect-Graph`

**Flow — Offline:**
1. Load config → import modules
2. `Initialize-OfflineCacheFromMapping` (if `-NameMappingPath` supplied)
3. `Enable-OfflineMode` (blocks any Graph API calls)
4. Parse JSON from `-PoliciesJsonPath` (handles `{"value":[...]}` wrapper or bare array)
5. `ConvertTo-PolicyObject -ResolveNames` for each policy (cache/well-known-only resolution)
6. Generate PowerPoint and/or HTML

---

### `Modules/GraphApiHelper.psm1`

Handles Graph API authentication, paginated requests, ID resolution with caching, and offline mode.

**Authentication:**

| Flow | Grant Type | Use Case |
|------|-----------|---------|
| Device Code | `urn:ietf:params:oauth:grant-type:device_code` | Interactive, delegated permissions |
| Client Credentials | `client_credentials` | Unattended, application permissions |

Auto-detection: if `-ClientSecret` is provided → `ClientCredentials`; otherwise → `DeviceCode`.

Token endpoint: `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`

**Offline Mode:**

Controlled by `$script:OfflineMode`. When `$true`, all five ID resolution functions skip the Graph API block and fall through to the cache or a friendly placeholder.

Name resolution order (both modes):
1. Special values (`"All"`, `"Office365"`, `"AllTrusted"`, etc.)
2. In-memory cache (`$script:NameCache`)
3. Well-known tables (`$script:WellKnownApps`, `$script:WellKnownRoles` — 120+ entries)
4. **Online only:** Graph API query → cache result
5. **Offline fallback:** `[Unknown Application]` / `[Unknown Group]` / `[Unknown User]` / `[Unknown Role]` / `[Unknown Location]` — no raw GUIDs exposed

**Exported Functions:**

| Function | Description |
|----------|-------------|
| `Connect-Graph` | Authenticate and store access token |
| `Disconnect-Graph` | Clear stored token and cache |
| `Invoke-GraphRequest` | REST wrapper with pagination and retry |
| `Get-ConditionalAccessPolicies` | Fetch all policies from Graph |
| `Get-ApplicationDisplayName` | Resolve app ID → name |
| `Get-UserDisplayName` | Resolve user ID → name |
| `Get-GroupDisplayName` | Resolve group ID → name |
| `Get-RoleDisplayName` | Resolve role template ID → name |
| `Get-NamedLocationName` | Resolve named location ID → name |
| `Get-WellKnownApplications` | Return built-in app GUID table |
| `Get-WellKnownRoles` | Return built-in role GUID table |
| `Clear-NameCache` | Flush all cached resolutions |
| `Enable-OfflineMode` | Set `$script:OfflineMode = $true` |
| `Initialize-OfflineCacheFromMapping` | Pre-populate cache from a JSON mapping file |

**Caching:** All resolved names stored in `$script:NameCache` (hashtable per entity type). Well-known entries are also written to cache on first hit to avoid repeated table lookups.

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
        Users = @{
            IncludeUsers  = [string[]]  # Resolved display names
            ExcludeUsers  = [string[]]
            IncludeGroups = [string[]]
            ExcludeGroups = [string[]]
            IncludeRoles  = [string[]]
            ExcludeRoles  = [string[]]
        }
        Applications = @{
            IncludeApplications = [string[]]
            ExcludeApplications = [string[]]
            IncludeUserActions  = [string[]]
        }
        Platforms        = @{ IncludePlatforms = [string[]]; ExcludePlatforms = [string[]] }
        Locations        = @{ IncludeLocations = [string[]]; ExcludeLocations = [string[]] }
        ClientAppTypes   = [string[]]
        SignInRiskLevels = [string[]]
        UserRiskLevels   = [string[]]
        Devices          = [hashtable]
    }
    GrantControls = @{
        Operator        = [string]   # "AND" or "OR"
        BuiltInControls = [string[]]
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

Uses `System.IO.Compression.ZipArchive` to build the ZIP structure in memory. Key quirk: `[Content_Types].xml` must be written with `-LiteralPath` because PowerShell's `-FilePath` treats `[` as a glob wildcard.

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
- `function arr(x) { return x == null ? [] : [].concat(x); }` — used on every nested field access in JavaScript

**HTML Views:**

| View | Description |
|------|-------------|
| Policy Cards | Searchable, filterable list; real-time search, state filter, dark mode, JSON export |
| Coverage Matrix | User population × application grid; red = gap; click cell to see covering policies |
| Overlap & Conflict Analyzer | Pairwise policy comparison; flags conflicts and redundancies |
| Application Lookup | Search any app to see all policies covering it and coverage stats |

---

### `Build-NameMapping.ps1` — Name Mapping Utility

Standalone utility script. Parses Entra ID CSV/JSON exports and writes a `NameMapping.json` file for use with `-OfflineMode -NameMappingPath`. Does not depend on any project module.

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-UsersCSV` | Path to users CSV (Entra Portal → Users → Download users) |
| `-GroupsCSV` | Path to groups CSV (Entra Portal → Groups → Download groups) |
| `-AppsCSV` | Path to enterprise apps CSV (Entra Portal → Enterprise Applications → Download) |
| `-NamedLocationsJson` | Path to named locations JSON (`GET /identity/conditionalAccess/namedLocations`) |
| `-OutputPath` | Output path for `NameMapping.json` (default: `.\NameMapping.json`) |
| `-Merge` | Merge new entries into an existing file rather than overwriting |

All inputs are optional — pass only what's available. The script detects column name variations across Entra export formats (e.g. `id` vs `Object ID`; `displayName` vs `Display name`).

**Key implementation note — Apps CSV column:** CA policies reference the `appId` (Application ID), not the `objectId` (service principal Object ID). The script explicitly looks for `appId` / `Application ID` columns and ignores `objectId` to avoid silent mismatches.

**Named locations:** There is no CSV export for named locations in the Entra Portal. The script accepts a JSON export from the Graph API endpoint `GET /identity/conditionalAccess/namedLocations` (supports `{"value":[...]}` wrapper or bare array).

---

## CI/CD — Security Scanning

`.github/workflows/security.yml` runs on every push and pull request to `main`.

| Job | Tool | Purpose |
|-----|------|---------|
| `secret-scan` | [Gitleaks](https://github.com/gitleaks/gitleaks) | Detect hardcoded credentials, tokens, and keys across the full git history |
| `powershell-scan` | [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) | Identify PowerShell security anti-patterns; results uploaded as SARIF to GitHub Code Scanning |

The workflow is intentionally set to trigger on PRs **targeting `main`** so security checks gate every merge from `dev` → `main`. No code reaches `main` without passing both checks.

---

## Branch Strategy

| Branch | Purpose |
|--------|---------|
| `main` | Stable, production-ready code; protected by security scanning CI |
| `dev` | Active development; PRs from here into `main` trigger the full security scan |

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
   grant_type=client_credentials, client_id, client_secret
   scope=https://graph.microsoft.com/.default
2. Store access_token + expiry in module scope
```

Required application permissions: `Policy.Read.All`, `Application.Read.All`, `Directory.Read.All`

---

## API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /identity/conditionalAccess/policies` | All CA policies (paginated) |
| `GET /servicePrincipals?$filter=appId eq '{id}'` | App display name (primary) |
| `GET /applications?$filter=appId eq '{id}'` | App display name (fallback) |
| `GET /users/{id}?$select=displayName` | User display name |
| `GET /groups/{id}?$select=displayName` | Group display name |
| `GET /directoryRoleTemplates/{id}` | Role display name |
| `GET /identity/conditionalAccess/namedLocations/{id}` | Named location display name |

All endpoints are skipped in offline mode.

---

## Error Handling

| Scenario | Handling |
|----------|----------|
| Auth failure | Display error, exit with code 1 |
| API 429 (rate limit) | Exponential backoff, up to 3 retries |
| Failed ID resolution (online) | `[Unknown: {guid}]` |
| Failed ID resolution (offline) | `[Unknown Application]` / `[Unknown Group]` / etc. — no GUID exposed |
| Network timeout | Retry up to 3 times with increasing delay |
| Single-element array serialization | `@()` operator + `arr()` JS helper |
| `[Content_Types].xml` wildcard expansion | `-LiteralPath` in `Out-File` |

---

## File Structure

```
ConditionalAccessDocumenter/
├── Get-ConditionalAccessReport.ps1    # Main entry point
├── Build-NameMapping.ps1              # Utility: build NameMapping.json from Entra exports
├── config.json                         # Configuration
├── NameMapping.example.json            # Template for offline name resolution
├── README.md                           # User documentation
├── DESIGN.md                           # This document
├── .github/
│   └── workflows/
│       └── security.yml                # Gitleaks + PSScriptAnalyzer CI
├── Modules/
│   ├── GraphApiHelper.psm1             # Graph API + auth + offline mode
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
