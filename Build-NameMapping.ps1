#Requires -Version 5.1
<#
.SYNOPSIS
    Builds a NameMapping.json file from Entra ID CSV/JSON exports for use with offline mode.

.DESCRIPTION
    Parses CSV exports from the Entra ID portal and/or a named locations JSON export,
    extracts GUID-to-display-name mappings, and writes a NameMapping.json file ready
    for use with Get-ConditionalAccessReport.ps1 -OfflineMode -NameMappingPath.

    All parameters are optional — pass only the exports you have available.
    If the output file already exists, use -Merge to combine with existing entries
    rather than overwriting.

.PARAMETER UsersCSV
    Path to the users CSV exported from:
    Entra Portal → Users → Download users

.PARAMETER GroupsCSV
    Path to the groups CSV exported from:
    Entra Portal → Groups → Download groups

.PARAMETER AppsCSV
    Path to the enterprise applications CSV exported from:
    Entra Portal → Enterprise Applications → Download

.PARAMETER NamedLocationsJson
    Path to a JSON file containing named locations exported from:
    Graph Explorer → GET /identity/conditionalAccess/namedLocations
    (supports both {"value":[...]} wrapper and bare array)

.PARAMETER OutputPath
    Path to write the NameMapping.json file. Defaults to .\NameMapping.json

.PARAMETER Merge
    If the output file already exists, merge new entries into it rather than overwriting.
    Existing entries are preserved; new entries are added; conflicts favour the new value.

.EXAMPLE
    .\Build-NameMapping.ps1 -UsersCSV ".\users.csv" -GroupsCSV ".\groups.csv" -AppsCSV ".\apps.csv"

    Builds NameMapping.json from three CSV exports.

.EXAMPLE
    .\Build-NameMapping.ps1 -AppsCSV ".\apps.csv" -NamedLocationsJson ".\namedLocations.json" -Merge

    Adds app and named location entries to an existing NameMapping.json.

.NOTES
    Author: Conditional Access Documenter
    Requires: PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UsersCSV,

    [Parameter(Mandatory = $false)]
    [string]$GroupsCSV,

    [Parameter(Mandatory = $false)]
    [string]$AppsCSV,

    [Parameter(Mandatory = $false)]
    [string]$NamedLocationsJson,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\NameMapping.json",

    [Parameter(Mandatory = $false)]
    [switch]$Merge
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  Build-NameMapping" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Validate at least one input was provided
if (-not $UsersCSV -and -not $GroupsCSV -and -not $AppsCSV -and -not $NamedLocationsJson) {
    Write-Error "No input files specified. Provide at least one of: -UsersCSV, -GroupsCSV, -AppsCSV, -NamedLocationsJson"
    exit 1
}

#region Helpers

function Resolve-InputPath {
    param([string]$Path, [string]$Label)
    if (-not $Path) { return $null }
    $resolved = $Path
    if (-not [System.IO.Path]::IsPathRooted($resolved)) {
        $resolved = Join-Path (Get-Location) $resolved
    }
    if (-not (Test-Path $resolved)) {
        Write-Error "$Label file not found: $resolved"
        exit 1
    }
    return $resolved
}

function Find-Column {
    <#
        Returns the first matching column name from a list of candidates,
        or $null if none found. Case-insensitive.
    #>
    param([string[]]$Headers, [string[]]$Candidates)
    foreach ($candidate in $Candidates) {
        $match = $Headers | Where-Object { $_ -ieq $candidate } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

function Read-CsvMapping {
    param(
        [string]$Path,
        [string]$Label,
        [string[]]$IdCandidates,
        [string[]]$NameCandidates
    )

    $rows = Import-Csv -Path $Path
    if (-not $rows) {
        Write-Warning "  $Label — file is empty, skipping"
        return @{}
    }

    $headers = $rows[0].PSObject.Properties.Name
    $idCol   = Find-Column -Headers $headers -Candidates $IdCandidates
    $nameCol = Find-Column -Headers $headers -Candidates $NameCandidates

    if (-not $idCol) {
        Write-Warning "  $Label — could not find ID column. Expected one of: $($IdCandidates -join ', '). Columns found: $($headers -join ', ')"
        return @{}
    }
    if (-not $nameCol) {
        Write-Warning "  $Label — could not find name column. Expected one of: $($NameCandidates -join ', '). Columns found: $($headers -join ', ')"
        return @{}
    }

    $mapping = @{}
    $skipped = 0

    foreach ($row in $rows) {
        $id   = $row.$idCol
        $name = $row.$nameCol

        # Skip rows with missing or obviously invalid GUIDs
        if (-not $id -or $id -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
            $skipped++
            continue
        }
        if (-not $name) {
            $skipped++
            continue
        }

        $mapping[$id.ToLower()] = $name
    }

    $loaded = $mapping.Count
    Write-Host "  $Label — $loaded entries loaded" -ForegroundColor Green
    if ($skipped -gt 0) {
        Write-Host "            $skipped rows skipped (missing or invalid GUID/name)" -ForegroundColor DarkGray
    }

    return $mapping
}

#endregion

#region Parse inputs

$applications   = @{}
$groups         = @{}
$users          = @{}
$namedLocations = @{}

# Users CSV
if ($UsersCSV) {
    $path = Resolve-InputPath -Path $UsersCSV -Label "UsersCSV"
    Write-Host "Reading users:            $path" -ForegroundColor Gray
    $users = Read-CsvMapping -Path $path -Label "Users" `
        -IdCandidates   @("id", "objectId", "Object ID", "Object Id") `
        -NameCandidates @("displayName", "Display name", "Display Name", "name")
}

# Groups CSV
if ($GroupsCSV) {
    $path = Resolve-InputPath -Path $GroupsCSV -Label "GroupsCSV"
    Write-Host "Reading groups:           $path" -ForegroundColor Gray
    $groups = Read-CsvMapping -Path $path -Label "Groups" `
        -IdCandidates   @("id", "objectId", "Object ID", "Object Id") `
        -NameCandidates @("displayName", "Display name", "Display Name", "name")
}

# Enterprise Apps CSV
# The column we need is appId (the Application ID), not objectId (the service principal ID)
if ($AppsCSV) {
    $path = Resolve-InputPath -Path $AppsCSV -Label "AppsCSV"
    Write-Host "Reading applications:     $path" -ForegroundColor Gray
    $applications = Read-CsvMapping -Path $path -Label "Applications" `
        -IdCandidates   @("appId", "Application ID", "Application Id", "applicationId", "App ID", "App Id") `
        -NameCandidates @("displayName", "Display name", "Display Name", "name", "Name")
}

# Named Locations JSON
if ($NamedLocationsJson) {
    $path = Resolve-InputPath -Path $NamedLocationsJson -Label "NamedLocationsJson"
    Write-Host "Reading named locations:  $path" -ForegroundColor Gray

    $json = Get-Content $path -Raw | ConvertFrom-Json
    $locations = if ($json.value) { @($json.value) } else { @($json) }

    $skipped = 0
    foreach ($loc in $locations) {
        $id   = $loc.id
        $name = $loc.displayName
        if ($id -and $name) {
            $namedLocations[$id.ToLower()] = $name
        } else {
            $skipped++
        }
    }
    Write-Host "  Named Locations — $($namedLocations.Count) entries loaded" -ForegroundColor Green
    if ($skipped -gt 0) {
        Write-Host "                   $skipped entries skipped (missing id or displayName)" -ForegroundColor DarkGray
    }
}

#endregion

#region Merge or build output

$resolvedOutput = $OutputPath
if (-not [System.IO.Path]::IsPathRooted($resolvedOutput)) {
    $resolvedOutput = Join-Path (Get-Location) $resolvedOutput
}

$existing = @{ applications = @{}; groups = @{}; users = @{}; namedLocations = @{} }

if ($Merge -and (Test-Path $resolvedOutput)) {
    Write-Host ""
    Write-Host "Merging with existing: $resolvedOutput" -ForegroundColor Gray
    $existingJson = Get-Content $resolvedOutput -Raw | ConvertFrom-Json

    foreach ($section in @("applications", "groups", "users", "namedLocations")) {
        if ($existingJson.$section) {
            $existingJson.$section.PSObject.Properties | ForEach-Object {
                $existing[$section][$_.Name] = $_.Value
            }
        }
    }
}

# Merge: existing entries first, new entries override on conflict
foreach ($key in $applications.Keys)   { $existing.applications[$key]   = $applications[$key] }
foreach ($key in $groups.Keys)         { $existing.groups[$key]          = $groups[$key] }
foreach ($key in $users.Keys)          { $existing.users[$key]           = $users[$key] }
foreach ($key in $namedLocations.Keys) { $existing.namedLocations[$key]  = $namedLocations[$key] }

# Build ordered output object
$output = [ordered]@{
    applications   = $existing.applications
    groups         = $existing.groups
    users          = $existing.users
    namedLocations = $existing.namedLocations
}

#endregion

#region Write output

$output | ConvertTo-Json -Depth 5 | Set-Content -Path $resolvedOutput -Encoding UTF8

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  Done" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output:        $resolvedOutput" -ForegroundColor White
Write-Host "Applications:  $($existing.applications.Count)" -ForegroundColor Gray
Write-Host "Groups:        $($existing.groups.Count)" -ForegroundColor Gray
Write-Host "Users:         $($existing.users.Count)" -ForegroundColor Gray
Write-Host "Named Locations: $($existing.namedLocations.Count)" -ForegroundColor Gray
Write-Host ""
Write-Host "Next step:" -ForegroundColor White
Write-Host "  .\Get-ConditionalAccessReport.ps1 -OfflineMode -PoliciesJsonPath '.\policies.json' -NameMappingPath '$resolvedOutput'" -ForegroundColor DarkGray
Write-Host ""

#endregion
