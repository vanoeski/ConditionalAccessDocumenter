#Requires -Version 5.1
<#
.SYNOPSIS
    Policy Parser Module for Conditional Access Policy Documenter

.DESCRIPTION
    Provides functions to parse and transform raw Conditional Access Policy
    JSON data into structured objects with resolved display names.
#>

function ConvertTo-PolicyObject {
    <#
    .SYNOPSIS
        Transforms a raw CA policy API response into a structured object with resolved names

    .PARAMETER Policy
        The raw policy object from Graph API

    .PARAMETER ResolveNames
        If specified, resolves IDs to display names (requires GraphApiHelper module)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$Policy,

        [Parameter(Mandatory = $false)]
        [switch]$ResolveNames
    )

    process {
        $policyObject = [PSCustomObject]@{
            Id               = $Policy.id
            DisplayName      = $Policy.displayName
            State            = Get-PolicyStateDisplayName -State $Policy.state
            StateRaw         = $Policy.state
            CreatedDateTime  = $Policy.createdDateTime
            ModifiedDateTime = $Policy.modifiedDateTime
            Conditions       = Get-PolicyConditionsSummary -Conditions $Policy.conditions -ResolveNames:$ResolveNames
            GrantControls    = Get-PolicyGrantControls -GrantControls $Policy.grantControls
            SessionControls  = Get-PolicySessionControls -SessionControls $Policy.sessionControls
        }

        return $policyObject
    }
}

function Get-PolicyStateDisplayName {
    <#
    .SYNOPSIS
        Converts policy state to a human-readable display name
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$State
    )

    switch ($State) {
        "enabled" { return "Enabled" }
        "disabled" { return "Disabled" }
        "enabledForReportingButNotEnforced" { return "Report-only" }
        default { return $State }
    }
}

function Get-PolicyConditionsSummary {
    <#
    .SYNOPSIS
        Extracts and formats policy conditions

    .PARAMETER Conditions
        The conditions object from the policy

    .PARAMETER ResolveNames
        If specified, resolves IDs to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Conditions,

        [Parameter(Mandatory = $false)]
        [switch]$ResolveNames
    )

    if (-not $Conditions) {
        return @{}
    }

    $summary = @{
        Users = @{
            IncludeUsers  = @()
            ExcludeUsers  = @()
            IncludeGroups = @()
            ExcludeGroups = @()
            IncludeRoles  = @()
            ExcludeRoles  = @()
            IncludeGuestsOrExternalUsers = $null
            ExcludeGuestsOrExternalUsers = $null
        }
        Applications = @{
            IncludeApplications   = @()
            ExcludeApplications   = @()
            IncludeUserActions    = @()
            IncludeAuthenticationContextClassReferences = @()
        }
        Platforms = @{
            IncludePlatforms = @()
            ExcludePlatforms = @()
        }
        Locations = @{
            IncludeLocations = @()
            ExcludeLocations = @()
        }
        ClientAppTypes     = @()
        SignInRiskLevels   = @()
        UserRiskLevels     = @()
        ServicePrincipalRiskLevels = @()
        Devices = @{
            IncludeDevices = @()
            ExcludeDevices = @()
            DeviceFilter   = $null
        }
    }

    # Parse Users
    if ($Conditions.users) {
        $users = $Conditions.users

        # Include Users
        if ($users.includeUsers) {
            $summary.Users.IncludeUsers = $users.includeUsers | ForEach-Object {
                if ($ResolveNames -and $_ -notin @("All", "GuestsOrExternalUsers", "None")) {
                    Get-UserDisplayName -UserId $_
                } else {
                    Format-SpecialValue -Value $_
                }
            }
        }

        # Exclude Users
        if ($users.excludeUsers) {
            $summary.Users.ExcludeUsers = $users.excludeUsers | ForEach-Object {
                if ($ResolveNames -and $_ -notin @("All", "GuestsOrExternalUsers", "None")) {
                    Get-UserDisplayName -UserId $_
                } else {
                    Format-SpecialValue -Value $_
                }
            }
        }

        # Include Groups
        if ($users.includeGroups) {
            $summary.Users.IncludeGroups = $users.includeGroups | ForEach-Object {
                if ($ResolveNames) {
                    Get-GroupDisplayName -GroupId $_
                } else {
                    $_
                }
            }
        }

        # Exclude Groups
        if ($users.excludeGroups) {
            $summary.Users.ExcludeGroups = $users.excludeGroups | ForEach-Object {
                if ($ResolveNames) {
                    Get-GroupDisplayName -GroupId $_
                } else {
                    $_
                }
            }
        }

        # Include Roles
        if ($users.includeRoles) {
            $summary.Users.IncludeRoles = $users.includeRoles | ForEach-Object {
                if ($ResolveNames) {
                    Get-RoleDisplayName -RoleId $_
                } else {
                    $_
                }
            }
        }

        # Exclude Roles
        if ($users.excludeRoles) {
            $summary.Users.ExcludeRoles = $users.excludeRoles | ForEach-Object {
                if ($ResolveNames) {
                    Get-RoleDisplayName -RoleId $_
                } else {
                    $_
                }
            }
        }

        # Guest/External Users
        if ($users.includeGuestsOrExternalUsers) {
            $summary.Users.IncludeGuestsOrExternalUsers = Format-GuestOrExternalUsers -Config $users.includeGuestsOrExternalUsers
        }
        if ($users.excludeGuestsOrExternalUsers) {
            $summary.Users.ExcludeGuestsOrExternalUsers = Format-GuestOrExternalUsers -Config $users.excludeGuestsOrExternalUsers
        }
    }

    # Parse Applications
    if ($Conditions.applications) {
        $apps = $Conditions.applications

        # Include Applications
        if ($apps.includeApplications) {
            $summary.Applications.IncludeApplications = $apps.includeApplications | ForEach-Object {
                if ($ResolveNames) {
                    Get-ApplicationDisplayName -AppId $_
                } else {
                    $_
                }
            }
        }

        # Exclude Applications
        if ($apps.excludeApplications) {
            $summary.Applications.ExcludeApplications = $apps.excludeApplications | ForEach-Object {
                if ($ResolveNames) {
                    Get-ApplicationDisplayName -AppId $_
                } else {
                    $_
                }
            }
        }

        # User Actions
        if ($apps.includeUserActions) {
            $summary.Applications.IncludeUserActions = $apps.includeUserActions | ForEach-Object {
                Format-UserAction -Action $_
            }
        }

        # Authentication Context
        if ($apps.includeAuthenticationContextClassReferences) {
            $summary.Applications.IncludeAuthenticationContextClassReferences = $apps.includeAuthenticationContextClassReferences
        }
    }

    # Parse Platforms
    if ($Conditions.platforms) {
        $platforms = $Conditions.platforms

        if ($platforms.includePlatforms) {
            $summary.Platforms.IncludePlatforms = $platforms.includePlatforms | ForEach-Object {
                Format-Platform -Platform $_
            }
        }

        if ($platforms.excludePlatforms) {
            $summary.Platforms.ExcludePlatforms = $platforms.excludePlatforms | ForEach-Object {
                Format-Platform -Platform $_
            }
        }
    }

    # Parse Locations
    if ($Conditions.locations) {
        $locations = $Conditions.locations

        if ($locations.includeLocations) {
            $summary.Locations.IncludeLocations = $locations.includeLocations | ForEach-Object {
                if ($ResolveNames -and $_ -notin @("All", "AllTrusted")) {
                    Get-NamedLocationName -LocationId $_
                } else {
                    Format-SpecialValue -Value $_
                }
            }
        }

        if ($locations.excludeLocations) {
            $summary.Locations.ExcludeLocations = $locations.excludeLocations | ForEach-Object {
                if ($ResolveNames -and $_ -notin @("All", "AllTrusted")) {
                    Get-NamedLocationName -LocationId $_
                } else {
                    Format-SpecialValue -Value $_
                }
            }
        }
    }

    # Parse Client App Types
    if ($Conditions.clientAppTypes) {
        $summary.ClientAppTypes = $Conditions.clientAppTypes | ForEach-Object {
            Format-ClientAppType -AppType $_
        }
    }

    # Parse Risk Levels
    if ($Conditions.signInRiskLevels) {
        $summary.SignInRiskLevels = $Conditions.signInRiskLevels | ForEach-Object {
            Format-RiskLevel -Level $_
        }
    }

    if ($Conditions.userRiskLevels) {
        $summary.UserRiskLevels = $Conditions.userRiskLevels | ForEach-Object {
            Format-RiskLevel -Level $_
        }
    }

    if ($Conditions.servicePrincipalRiskLevels) {
        $summary.ServicePrincipalRiskLevels = $Conditions.servicePrincipalRiskLevels | ForEach-Object {
            Format-RiskLevel -Level $_
        }
    }

    # Parse Devices
    if ($Conditions.devices) {
        $devices = $Conditions.devices

        if ($devices.includeDevices) {
            $summary.Devices.IncludeDevices = $devices.includeDevices
        }

        if ($devices.excludeDevices) {
            $summary.Devices.ExcludeDevices = $devices.excludeDevices
        }

        if ($devices.deviceFilter) {
            $summary.Devices.DeviceFilter = @{
                Mode = $devices.deviceFilter.mode
                Rule = $devices.deviceFilter.rule
            }
        }
    }

    return $summary
}

function Get-PolicyGrantControls {
    <#
    .SYNOPSIS
        Parses grant controls from a policy

    .PARAMETER GrantControls
        The grantControls object from the policy
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$GrantControls
    )

    if (-not $GrantControls) {
        return @{
            Operator        = $null
            BuiltInControls = @()
            CustomControls  = @()
            TermsOfUse      = @()
            AuthenticationStrength = $null
        }
    }

    $controls = @{
        Operator        = $GrantControls.operator
        BuiltInControls = @()
        CustomControls  = @()
        TermsOfUse      = @()
        AuthenticationStrength = $null
    }

    if ($GrantControls.builtInControls) {
        $controls.BuiltInControls = $GrantControls.builtInControls | ForEach-Object {
            Format-GrantControl -Control $_
        }
    }

    if ($GrantControls.customAuthenticationFactors) {
        $controls.CustomControls = $GrantControls.customAuthenticationFactors
    }

    if ($GrantControls.termsOfUse) {
        $controls.TermsOfUse = $GrantControls.termsOfUse
    }

    if ($GrantControls.authenticationStrength) {
        $controls.AuthenticationStrength = @{
            Id          = $GrantControls.authenticationStrength.id
            DisplayName = $GrantControls.authenticationStrength.displayName
        }
    }

    return $controls
}

function Get-PolicySessionControls {
    <#
    .SYNOPSIS
        Parses session controls from a policy

    .PARAMETER SessionControls
        The sessionControls object from the policy
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$SessionControls
    )

    if (-not $SessionControls) {
        return @{
            SignInFrequency           = $null
            PersistentBrowser         = $null
            CloudAppSecurity          = $null
            ApplicationEnforcedRestrictions = $false
            DisableResilienceDefaults = $false
            ContinuousAccessEvaluation = $null
        }
    }

    $controls = @{
        SignInFrequency           = $null
        PersistentBrowser         = $null
        CloudAppSecurity          = $null
        ApplicationEnforcedRestrictions = $false
        DisableResilienceDefaults = $false
        ContinuousAccessEvaluation = $null
    }

    if ($SessionControls.signInFrequency) {
        $sif = $SessionControls.signInFrequency
        if ($sif.isEnabled) {
            if ($sif.frequencyInterval -eq "everyTime") {
                $controls.SignInFrequency = "Every time"
            } else {
                $controls.SignInFrequency = "$($sif.value) $($sif.type)"
            }
            if ($sif.authenticationType) {
                $controls.SignInFrequency += " ($($sif.authenticationType))"
            }
        }
    }

    if ($SessionControls.persistentBrowser) {
        $pb = $SessionControls.persistentBrowser
        if ($pb.isEnabled) {
            $controls.PersistentBrowser = $pb.mode
        }
    }

    if ($SessionControls.cloudAppSecurity) {
        $cas = $SessionControls.cloudAppSecurity
        if ($cas.isEnabled) {
            $controls.CloudAppSecurity = $cas.cloudAppSecurityType
        }
    }

    if ($SessionControls.applicationEnforcedRestrictions) {
        $controls.ApplicationEnforcedRestrictions = $SessionControls.applicationEnforcedRestrictions.isEnabled -eq $true
    }

    if ($SessionControls.disableResilienceDefaults -eq $true) {
        $controls.DisableResilienceDefaults = $true
    }

    if ($SessionControls.continuousAccessEvaluation) {
        $cae = $SessionControls.continuousAccessEvaluation
        $controls.ContinuousAccessEvaluation = $cae.mode
    }

    return $controls
}

function Format-SpecialValue {
    <#
    .SYNOPSIS
        Formats special values like "All", "None", etc. to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    switch ($Value) {
        "All" { return "All" }
        "None" { return "None" }
        "GuestsOrExternalUsers" { return "Guests or external users" }
        "AllTrusted" { return "All trusted locations" }
        default { return $Value }
    }
}

function Format-Platform {
    <#
    .SYNOPSIS
        Formats platform values to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Platform
    )

    switch ($Platform) {
        "all" { return "All platforms" }
        "android" { return "Android" }
        "iOS" { return "iOS" }
        "windows" { return "Windows" }
        "windowsPhone" { return "Windows Phone" }
        "macOS" { return "macOS" }
        "linux" { return "Linux" }
        "unknownFutureValue" { return "Unknown" }
        default { return $Platform }
    }
}

function Format-ClientAppType {
    <#
    .SYNOPSIS
        Formats client app type values to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppType
    )

    switch ($AppType) {
        "all" { return "All client apps" }
        "browser" { return "Browser" }
        "mobileAppsAndDesktopClients" { return "Mobile apps and desktop clients" }
        "exchangeActiveSync" { return "Exchange ActiveSync clients" }
        "easSupported" { return "EAS supported clients" }
        "other" { return "Other clients" }
        default { return $AppType }
    }
}

function Format-RiskLevel {
    <#
    .SYNOPSIS
        Formats risk level values to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Level
    )

    switch ($Level) {
        "low" { return "Low" }
        "medium" { return "Medium" }
        "high" { return "High" }
        "hidden" { return "Hidden" }
        "none" { return "None" }
        "unknownFutureValue" { return "Unknown" }
        default { return $Level }
    }
}

function Format-GrantControl {
    <#
    .SYNOPSIS
        Formats grant control values to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Control
    )

    switch ($Control) {
        "mfa" { return "Require multi-factor authentication" }
        "compliantDevice" { return "Require device to be marked as compliant" }
        "domainJoinedDevice" { return "Require Hybrid Azure AD joined device" }
        "approvedApplication" { return "Require approved client app" }
        "compliantApplication" { return "Require app protection policy" }
        "passwordChange" { return "Require password change" }
        "block" { return "Block access" }
        default { return $Control }
    }
}

function Format-UserAction {
    <#
    .SYNOPSIS
        Formats user action values to display names
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    switch ($Action) {
        "urn:user:registersecurityinfo" { return "Register security info" }
        "urn:user:registerdevice" { return "Register or join devices" }
        default { return $Action }
    }
}

function Format-GuestOrExternalUsers {
    <#
    .SYNOPSIS
        Formats guest/external user configuration to readable text
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $types = @()

    if ($Config.guestOrExternalUserTypes) {
        $typeFlags = $Config.guestOrExternalUserTypes
        if ($typeFlags -match "internalGuest") { $types += "Internal guests" }
        if ($typeFlags -match "b2bCollaborationGuest") { $types += "B2B collaboration guests" }
        if ($typeFlags -match "b2bCollaborationMember") { $types += "B2B collaboration members" }
        if ($typeFlags -match "b2bDirectConnectUser") { $types += "B2B direct connect users" }
        if ($typeFlags -match "otherExternalUser") { $types += "Other external users" }
        if ($typeFlags -match "serviceProvider") { $types += "Service providers" }
    }

    $tenantInfo = ""
    if ($Config.externalTenants) {
        $tenants = $Config.externalTenants
        if ($tenants.membershipKind -eq "all") {
            $tenantInfo = " from all organizations"
        } elseif ($tenants.membershipKind -eq "enumerated" -and $tenants.members) {
            $tenantInfo = " from specific organizations"
        }
    }

    return ($types -join ", ") + $tenantInfo
}

function Get-PolicySummaryText {
    <#
    .SYNOPSIS
        Generates a brief text summary of a policy

    .PARAMETER Policy
        The parsed policy object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Policy
    )

    $parts = @()

    # Users summary
    $userPart = ""
    if ($Policy.Conditions.Users.IncludeUsers -contains "All") {
        $userPart = "All users"
    } elseif ($Policy.Conditions.Users.IncludeUsers.Count -gt 0) {
        $userPart = "$($Policy.Conditions.Users.IncludeUsers.Count) user(s)"
    }
    if ($Policy.Conditions.Users.IncludeGroups.Count -gt 0) {
        $userPart += $(if ($userPart) { ", " }) + "$($Policy.Conditions.Users.IncludeGroups.Count) group(s)"
    }
    if ($Policy.Conditions.Users.IncludeRoles.Count -gt 0) {
        $userPart += $(if ($userPart) { ", " }) + "$($Policy.Conditions.Users.IncludeRoles.Count) role(s)"
    }
    if ($userPart) { $parts += "Users: $userPart" }

    # Apps summary
    $appPart = ""
    if ($Policy.Conditions.Applications.IncludeApplications -contains "All") {
        $appPart = "All cloud apps"
    } elseif ($Policy.Conditions.Applications.IncludeApplications.Count -gt 0) {
        $appPart = "$($Policy.Conditions.Applications.IncludeApplications.Count) app(s)"
    }
    if ($Policy.Conditions.Applications.IncludeUserActions.Count -gt 0) {
        $appPart = "$($Policy.Conditions.Applications.IncludeUserActions.Count) user action(s)"
    }
    if ($appPart) { $parts += "Apps: $appPart" }

    # Grant controls summary
    if ($Policy.GrantControls.BuiltInControls.Count -gt 0) {
        $controls = $Policy.GrantControls.BuiltInControls -join ", "
        $parts += "Controls: $controls"
    }

    return $parts -join " | "
}

# Export functions
Export-ModuleMember -Function @(
    'ConvertTo-PolicyObject',
    'Get-PolicyStateDisplayName',
    'Get-PolicyConditionsSummary',
    'Get-PolicyGrantControls',
    'Get-PolicySessionControls',
    'Get-PolicySummaryText',
    'Format-SpecialValue',
    'Format-Platform',
    'Format-ClientAppType',
    'Format-RiskLevel',
    'Format-GrantControl',
    'Format-UserAction',
    'Format-GuestOrExternalUsers'
)
