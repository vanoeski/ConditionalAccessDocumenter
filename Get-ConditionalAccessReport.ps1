#Requires -Version 5.1
<#
.SYNOPSIS
    Generates PowerPoint and HTML documentation of Entra ID Conditional Access Policies

.DESCRIPTION
    This script connects to Microsoft Graph API, retrieves all Conditional Access policies
    from your Entra ID tenant, and generates both a PowerPoint presentation and an interactive
    HTML report documenting each policy.

    Features:
    - Translates application IDs to friendly names
    - Resolves user, group, and role IDs to display names
    - Generates professional PowerPoint slides (one per policy)
    - Creates interactive HTML with search, filter, and dark mode

.PARAMETER TenantId
    The tenant ID or domain name (e.g., contoso.onmicrosoft.com).
    Defaults to "common" for multi-tenant authentication.

.PARAMETER OutputPath
    Directory where reports will be saved. Defaults to ./Output

.PARAMETER HtmlOnly
    Generate only the HTML report, skip PowerPoint

.PARAMETER PptxOnly
    Generate only the PowerPoint report, skip HTML

.PARAMETER IncludeDisabled
    Include disabled policies in the report. Defaults to $true

.EXAMPLE
    .\Get-ConditionalAccessReport.ps1

    Runs with default settings, prompting for authentication.

.EXAMPLE
    .\Get-ConditionalAccessReport.ps1 -TenantId "contoso.onmicrosoft.com" -OutputPath "C:\Reports"

    Specifies tenant and output directory.

.EXAMPLE
    .\Get-ConditionalAccessReport.ps1 -HtmlOnly

    Generates only the HTML report.

.NOTES
    Author: Conditional Access Documenter
    Requires: PowerShell 5.1 or later
    No third-party dependencies required
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId = "common",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\Output",

    [Parameter(Mandatory = $false)]
    [switch]$HtmlOnly,

    [Parameter(Mandatory = $false)]
    [switch]$PptxOnly,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeDisabled = $true,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = ".\config.json"
)

#region Initialization

$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date

# Get script directory for module loading
$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $scriptDir) {
    $scriptDir = Get-Location
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Conditional Access Policy Documenter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Load configuration
$config = $null
$configFile = Join-Path $scriptDir $ConfigPath
if (Test-Path $configFile) {
    Write-Host "Loading configuration from: $configFile" -ForegroundColor Gray
    $config = Get-Content $configFile | ConvertFrom-Json
} else {
    Write-Host "Using default configuration" -ForegroundColor Gray
    $config = [PSCustomObject]@{
        outputDirectory = "./Output"
        pptxFileName = "ConditionalAccessPolicies.pptx"
        htmlFileName = "ConditionalAccessPolicies.html"
        includeDisabledPolicies = $true
        theme = @{
            primaryColor = "#0078D4"
            enabledColor = "#107C10"
            disabledColor = "#A80000"
            reportOnlyColor = "#FFB900"
        }
    }
}

# Override config with parameters if specified
if ($PSBoundParameters.ContainsKey('OutputPath')) {
    $config.outputDirectory = $OutputPath
}
if ($PSBoundParameters.ContainsKey('IncludeDisabled')) {
    $config.includeDisabledPolicies = $IncludeDisabled
}

# Resolve output path
$outputDir = $config.outputDirectory
if (-not [System.IO.Path]::IsPathRooted($outputDir)) {
    $outputDir = Join-Path $scriptDir $outputDir
}

# Load required modules
Write-Host "Loading modules..." -ForegroundColor Gray

$modulesPath = Join-Path $scriptDir "Modules"
$modules = @(
    "GraphApiHelper",
    "PolicyParser",
    "PowerPointGenerator",
    "HtmlGenerator"
)

foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath "$module.psm1"
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force -DisableNameChecking
        Write-Host "  Loaded: $module" -ForegroundColor DarkGray
    } else {
        Write-Error "Required module not found: $modulePath"
        exit 1
    }
}

# Add System.Web for HTML encoding
Add-Type -AssemblyName System.Web

#endregion

#region Authentication

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host ""

$connected = Connect-Graph -TenantId $TenantId

if (-not $connected) {
    Write-Error "Failed to authenticate to Microsoft Graph. Please try again."
    exit 1
}

Write-Host ""

#endregion

#region Fetch Policies

$rawPolicies = @(Get-ConditionalAccessPolicies)

if (-not $rawPolicies -or $rawPolicies.Count -eq 0) {
    Write-Warning "No Conditional Access policies found in this tenant."
    Disconnect-Graph
    exit 0
}

Write-Host "Found $($rawPolicies.Count) policies" -ForegroundColor Green
Write-Host ""

#endregion

#region Process Policies

Write-Host "Processing policies and resolving names..." -ForegroundColor Cyan
Write-Host "This may take a few minutes depending on the number of policies." -ForegroundColor Gray
Write-Host ""

$policies = @()
$totalPolicies = @($rawPolicies).Count
$currentPolicy = 0

foreach ($rawPolicy in $rawPolicies) {
    $currentPolicy++
    $percentComplete = if ($totalPolicies -gt 0) { [math]::Round(($currentPolicy / $totalPolicies) * 100) } else { 0 }

    Write-Progress -Activity "Processing Policies" -Status "$currentPolicy of $totalPolicies - $($rawPolicy.displayName)" -PercentComplete $percentComplete

    # Skip disabled policies if configured
    if (-not $config.includeDisabledPolicies -and $rawPolicy.state -eq "disabled") {
        Write-Host "  Skipping disabled: $($rawPolicy.displayName)" -ForegroundColor DarkGray
        continue
    }

    try {
        $parsedPolicy = $rawPolicy | ConvertTo-PolicyObject -ResolveNames
        $policies += $parsedPolicy
        Write-Host "  Processed: $($rawPolicy.displayName)" -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "  Failed to process: $($rawPolicy.displayName) - $_"
    }
}

Write-Progress -Activity "Processing Policies" -Completed
Write-Host ""
Write-Host "Successfully processed $($policies.Count) policies" -ForegroundColor Green
Write-Host ""

#endregion

#region Generate Reports

# Ensure output directory exists
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Gray
}

# Apply theme settings if available
if ($config.theme) {
    $themeParams = @{}
    if ($config.theme.primaryColor) { $themeParams.PrimaryColor = $config.theme.primaryColor }
    if ($config.theme.enabledColor) { $themeParams.EnabledColor = $config.theme.enabledColor }
    if ($config.theme.disabledColor) { $themeParams.DisabledColor = $config.theme.disabledColor }
    if ($config.theme.reportOnlyColor) { $themeParams.ReportOnlyColor = $config.theme.reportOnlyColor }

    if (-not $HtmlOnly) {
        Set-PptxTheme @themeParams
    }
    if (-not $PptxOnly) {
        Set-HtmlTheme @themeParams
    }
}

# Generate PowerPoint
if (-not $HtmlOnly) {
    Write-Host "Generating PowerPoint presentation..." -ForegroundColor Cyan

    $pptxPath = Join-Path $outputDir $config.pptxFileName

    # Create presentation
    $pptx = New-PptxDocument -Title "Conditional Access Policies" -Author "Conditional Access Documenter"

    # Add title slide
    $pptx = Add-TitleSlide -Document $pptx -Title "Conditional Access Policies" -Subtitle "Entra ID Security Documentation"

    # Add policy slides
    foreach ($policy in ($policies | Sort-Object DisplayName)) {
        $pptx = Add-PolicySlide -Document $pptx -Policy $policy
    }

    # Add summary slide
    $pptx = Add-SummarySlide -Document $pptx -Policies $policies

    # Save
    $saved = Save-PptxDocument -Document $pptx -Path $pptxPath

    if ($saved) {
        Write-Host "  PowerPoint: $pptxPath" -ForegroundColor Green
    }

    Write-Host ""
}

# Generate HTML
if (-not $PptxOnly) {
    Write-Host "Generating HTML report..." -ForegroundColor Cyan

    $htmlPath = Join-Path $outputDir $config.htmlFileName

    $htmlSaved = New-HtmlReport -Policies $policies -Title "Conditional Access Policies Report" -Path $htmlPath

    if ($htmlSaved) {
        Write-Host "  HTML: $htmlPath" -ForegroundColor Green
    }

    Write-Host ""
}

#endregion

#region Summary

$endTime = Get-Date
$duration = $endTime - $script:StartTime

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Generation Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor White
Write-Host "  Total Policies: $($policies.Count)" -ForegroundColor Gray
Write-Host "  Enabled: $(($policies | Where-Object { $_.StateRaw -eq 'enabled' }).Count)" -ForegroundColor Green
Write-Host "  Report-Only: $(($policies | Where-Object { $_.StateRaw -eq 'enabledForReportingButNotEnforced' }).Count)" -ForegroundColor Yellow
Write-Host "  Disabled: $(($policies | Where-Object { $_.StateRaw -eq 'disabled' }).Count)" -ForegroundColor Red
Write-Host ""
Write-Host "Output Directory: $outputDir" -ForegroundColor Gray
Write-Host "Duration: $($duration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Gray
Write-Host ""

# Disconnect
Disconnect-Graph

#endregion
