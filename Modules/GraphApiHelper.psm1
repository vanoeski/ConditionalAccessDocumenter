#Requires -Version 5.1
<#
.SYNOPSIS
    Microsoft Graph API Helper Module for Conditional Access Policy Documenter

.DESCRIPTION
    Provides functions to authenticate and interact with Microsoft Graph API
    for retrieving Conditional Access Policies and resolving IDs to display names.
#>

# Script-level variables for caching
$script:GraphToken = $null
$script:TokenExpiry = $null
$script:TenantId = $null
$script:NameCache = @{
    Applications = @{}
    Users = @{}
    Groups = @{}
    Roles = @{}
    NamedLocations = @{}
}

# Well-known Microsoft Application IDs
$script:WellKnownApps = @{
    "00000001-0000-0000-c000-000000000000" = "Azure ESTS Service"
    "00000002-0000-0000-c000-000000000000" = "Azure Active Directory Graph"
    "00000003-0000-0000-c000-000000000000" = "Microsoft Graph"
    "00000002-0000-0ff1-ce00-000000000000" = "Office 365 Exchange Online"
    "00000003-0000-0ff1-ce00-000000000000" = "Office 365 SharePoint Online"
    "00000004-0000-0ff1-ce00-000000000000" = "Office 365 Skype for Business"
    "00000006-0000-0ff1-ce00-000000000000" = "Microsoft Office 365 Portal"
    "00000007-0000-0ff1-ce00-000000000000" = "Office 365 Outlook Online"
    "00000009-0000-0000-c000-000000000000" = "Power BI Service"
    "0000000c-0000-0000-c000-000000000000" = "Microsoft App Access Panel"
    "00000012-0000-0000-c000-000000000000" = "Microsoft Rights Management Services"
    "00000015-0000-0000-c000-000000000000" = "Microsoft Dynamics ERP"
    "00b41c95-dab0-4487-9791-b9d2c32c80f2" = "Office 365 Management"
    "04b07795-8ddb-461a-bbee-02f9e1bf7b46" = "Microsoft Azure CLI"
    "0cb7b9ec-5336-483b-bc31-b15b5788de71" = "ASM Campaign Servicing"
    "0cd196ee-71bf-4fd6-a57c-b491ffd4fb1e" = "Microsoft Intune Enrollment"
    "0f698dd4-f011-4d23-a33e-b36416dcb1e6" = "Microsoft Intune API"
    "1195a167-45d4-4ed0-9f16-4f7b74e19ac8" = "Azure Multi-Factor Auth Connector"
    "14d82eec-204b-4c2f-b7e8-296a70dab67e" = "Microsoft Graph PowerShell"
    "18fbca16-2224-45f6-85b0-f7bf2b39b3f3" = "Microsoft Docs"
    "1950a258-227b-4e31-a9cf-717495945fc2" = "Microsoft Azure PowerShell"
    "1b730954-1685-4b74-9bfd-dac224a7b894" = "Azure Active Directory PowerShell"
    "1fec8e78-bce4-4aaf-ab1b-5451cc387264" = "Microsoft Teams"
    "23523755-3a2b-41ca-9315-f81f3f566a95" = "ACOM Azure Website"
    "268761a2-03f3-40df-8a8b-c3db24145b6b" = "OneDrive SyncEngine"
    "26a7ee05-5602-4d76-a7ba-eae8b7b67941" = "Windows Search"
    "26abc9a8-24f0-4b11-8234-e86ede698878" = "Office 365 Information Protection"
    "27922004-5251-4030-b22d-91ecd9a37ea4" = "Outlook Mobile"
    "28b567f6-162c-4f54-99a0-6887f387bbcc" = "SharePoint Home"
    "29d9ed98-a469-4536-ade2-f981bc1d605e" = "Microsoft Authentication Broker"
    "2d4d3d8e-2be3-4bef-9f87-7875a61c29de" = "OneNote"
    "2d7f3606-b07d-41d1-b9d2-0d0c9296a6e8" = "Microsoft Bing Search for Microsoft Edge"
    "4345a7b9-9a63-4910-a426-35363201d503" = "O365 Suite UX"
    "45a330b1-b1ec-4cc1-9161-9f03992aa49f" = "Windows Update for Business Deployment Service"
    "4765445b-32c6-49b0-83e6-1d93765276ca" = "Office 365 Information Protection"
    "497effe9-df71-4043-a8bb-14cf78c4b63b" = "Windows Virtual Desktop"
    "4990cffe-04e8-4e8b-808a-1175604b879f" = "Microsoft Authenticator App"
    "51be292c-a17e-4f17-9a7e-4b661fb16dd2" = "Skype for Business"
    "57336123-6e14-4acc-8dcf-287b6088aa28" = "Microsoft Whiteboard"
    "5e3ce6c0-2b1f-4285-8d4b-75ee78787346" = "Microsoft Teams - Device Admin Agent"
    "60c8bde5-3167-4f92-8fdb-059f6176dc0f" = "Enterprise Roaming and Backup"
    "66375f6b-983f-4c2c-9701-d680650f588f" = "Microsoft Planner"
    "67e3df25-268a-4324-a550-0de1c7f97287" = "Microsoft Stream Portal"
    "6a462b07-e56f-4817-94c2-ed83cd1f037d" = "Azure SQL Database and Data Warehouse"
    "7557eb47-c689-4224-abcf-aef9bd7573df" = "Microsoft Kaizala"
    "797f4846-ba00-4fd7-ba43-dac1f8f63013" = "Windows Azure Service Management API"
    "7ab7862c-4c57-491e-8a45-d52a7e023983" = "Windows Store for Business"
    "7ae974c5-1af7-4923-af3a-fb1fd14dcb7e" = "Microsoft Approval Management"
    "835b2a73-6e10-4aa5-a979-21dfda45231c" = "Microsoft Bing"
    "871c010f-5e61-4fb1-83ac-98610a7e9110" = "Microsoft Flow"
    "89bee1f7-5e6e-4d8a-9f3d-ecd601259da7" = "Office 365 Metrics"
    "8edd93e1-2103-40b4-bd70-6e34e586362d" = "Windows Azure Security Resource Provider"
    "91ca2ca5-3b3e-41dd-ab65-809fa3dffffa" = "Skype Teams Firehose"
    "9ba1a5c7-f17a-4de9-a1f1-6178c8d51223" = "Microsoft Intune API"
    "9ea1ad79-fdb6-4f9a-8bc3-2b70f96e34c7" = "Bing"
    "a0c73c16-a7e3-4564-9a95-2bdf47383716" = "Microsoft Exchange Online Protection"
    "a3475900-ccec-4a69-98f5-a65cd5dc5306" = "Partner Customer Delegated Admin"
    "a57aca87-cbc0-4f3c-8b9e-dc095fdc8978" = "Microsoft Flow Portal"
    "a672d62c-fc7b-4e81-a576-e60dc46e951d" = "Microsoft Managed Desktop"
    "a970bac6-63fe-4ec5-8884-8536862c42d4" = "Minecraft Education Edition"
    "aa580612-c342-4ace-9055-8edee43ccb89" = "Microsoft Staff Hub"
    "ab9b8c07-8f02-4f72-87fa-80105867a763" = "OneDrive Web"
    "ae8e128e-080f-4086-b0e3-4c19301ada69" = "Security Events Service"
    "b23dd4db-9142-4734-867f-3577f640ad0c" = "Windows Configuration Designer"
    "b669c6ea-1adf-453f-b8bc-6571571f8e6f" = "Microsoft Azure AD Identity Governance Insights"
    "b73f62d0-210b-4f1f-87bf-dd4742d59d09" = "Microsoft Bing Default Search Engine"
    "b779f6cc-7900-4a8a-b0fd-e0e2dc24aba4" = "Office 365 YammerOnOls"
    "c1c74fed-04c9-4704-80dc-9f79a2e515cb" = "Skype Presence Service"
    "c26550d6-bc82-4484-82ca-ac1c75308ca3" = "Microsoft Forms"
    "c44b4083-3bb0-49c1-b47d-974e53cbdf3c" = "Azure Portal"
    "c5393580-f805-4401-95e8-94b7a6ef2fc2" = "Office 365 Management APIs"
    "c9a559d2-7aab-4f13-a6ed-e7e9c52aec87" = "Microsoft Forms"
    "ca9319e0-74eb-4c9c-b7e9-4c1abb1d5dbc" = "Microsoft Power BI Information Service"
    "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" = "Microsoft Teams Web Client"
    "cf36b471-5b44-428c-9ce7-313bf84528de" = "Microsoft Bing Search"
    "cf53fce8-def6-4aeb-8d30-b158e7b1cf83" = "Microsoft Stream Service"
    "d176f6e7-38e5-40c9-8a78-3998aab820e7" = "My Apps"
    "d3590ed6-52b3-4102-aeff-aad2292ab01c" = "Microsoft Office"
    "d73f4b35-55c9-48c7-8b10-651f6f2acb2e" = "Windows ADFS"
    "d924a533-3729-4708-b3e8-1d2445af35e3" = "Microsoft Teams Tasks Service"
    "de8bc8b5-d9f9-48b1-a8ad-b748da725064" = "Microsoft Graph Command Line Tools"
    "e1ef36fd-b883-4dbf-97f0-9ece4b576fc6" = "Microsoft Edge Sync Service"
    "e9c51622-460d-4d3d-952d-966a5b1da34c" = "Microsoft Edge DevTools"
    "e9f49c6b-5ce5-44c8-925d-015017e9f7ad" = "Azure Data Lake"
    "eaf8a961-f56e-47eb-9ffd-936e22a554ef" = "Microsoft Bookings"
    "f44b1140-bc5e-48c6-8dc0-5cf5a53c0e34" = "Microsoft Edge Enterprise New Tab Page"
    "f5aeb603-2a64-4f37-b9a8-b544f3542865" = "Microsoft Pay"
    "fa163d49-dcc1-4686-87b3-f49e05c0d6bb" = "Microsoft To-Do"
    "fc0f3af4-6835-4174-b806-f7db311fd2f3" = "Microsoft Intune Portal"
    "fc780465-2017-40d4-a0c5-307022471b92" = "Azure Portal"
    "fdf9885b-dd37-42bf-82e5-c3129ef5a302" = "Microsoft Intune Company Portal"
    "All" = "All cloud apps"
    "Office365" = "Office 365"
    "MicrosoftAdminPortals" = "Microsoft Admin Portals"
    "none" = "None"
}

# Well-known Directory Role Template IDs
$script:WellKnownRoles = @{
    "62e90394-69f5-4237-9190-012177145e10" = "Global Administrator"
    "f28a1f50-f6e7-4571-818b-6a12f2af6b6c" = "SharePoint Administrator"
    "29232cdf-9323-42fd-ade2-1d097af3e4de" = "Exchange Administrator"
    "b1be1c3e-b65d-4f19-8427-f6fa0d97feb9" = "Conditional Access Administrator"
    "194ae4cb-b126-40b2-bd5b-6091b380977d" = "Security Administrator"
    "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3" = "Application Administrator"
    "158c047a-c907-4556-b7ef-446551a6b5f7" = "Cloud Application Administrator"
    "e8611ab8-c189-46e8-94e1-60213ab1f814" = "Privileged Role Administrator"
    "7be44c8a-adaf-4e2a-84d6-ab2649e08a13" = "Privileged Authentication Administrator"
    "c4e39bd9-1100-46d3-8c65-fb160da0071f" = "Authentication Administrator"
    "966707d0-3269-4727-9be2-8c3a10f19b9d" = "Password Administrator"
    "fdd7a751-b60b-444a-984c-02652fe8fa1c" = "Groups Administrator"
    "fe930be7-5e62-47db-91af-98c3a49a38b1" = "User Administrator"
    "729827e3-9c14-49f7-bb1b-9608f156bbb8" = "Helpdesk Administrator"
    "f023fd81-a637-4b56-95fd-791ac0226033" = "Service Support Administrator"
    "b0f54661-2d74-4c50-afa3-1ec803f12efe" = "Billing Administrator"
    "a9ea8996-122f-4c74-9520-8edcd192826c" = "Dynamics 365 Administrator"
    "44367163-eba1-44c3-98af-f5787879f96a" = "Dynamics 365 Service Administrator"
    "11648597-926c-4cf3-9c36-bcebb0ba8dcc" = "Power Platform Administrator"
    "3a2c62db-5318-420d-8d74-23affee5d9d5" = "Intune Administrator"
    "e00e864a-17c5-4a4b-9c06-f5b95a8d5bd8" = "Identity Governance Administrator"
    "8ac3fc64-6eca-42ea-9e69-59f4c7b60eb2" = "Hybrid Identity Administrator"
    "5d6b6bb7-de71-4623-b4af-96380a352509" = "Security Reader"
    "17315797-102d-40b4-93e0-432062caca18" = "Compliance Administrator"
    "d29b2b05-8046-44ba-8758-1e26182fcf32" = "Directory Synchronization Accounts"
    "9360feb5-f418-4baa-8175-e2a00bac4301" = "Directory Writers"
    "88d8e3e3-8f55-4a1e-953a-9b9898b8876b" = "Directory Readers"
    "e6d1a23a-da11-4be4-9570-befc86d067a7" = "Compliance Data Administrator"
    "3edaf663-341e-4475-9f94-5c398ef6c070" = "Customer LockBox Access Approver"
    "38a96431-2bdf-4b4c-8b6e-5d3d8abac1a4" = "Desktop Analytics Administrator"
    "4a5d8f65-41da-4de4-8968-e035b65339cf" = "Reports Reader"
    "790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b" = "Message Center Reader"
    "4d6ac14f-3453-41d0-bef9-a3e0c569773a" = "License Administrator"
    "2b745bdf-0803-4d80-aa65-822c4493daac" = "Office Apps Administrator"
    "8835291a-918c-4fd7-a9ce-faa49f0cf7d9" = "Teams Communications Administrator"
    "baf37b3a-610e-45da-9e62-d9d1e5e8914b" = "Teams Communications Support Engineer"
    "f70938a0-fc10-4177-9e90-2178f8765737" = "Teams Communications Support Specialist"
    "69091246-20e8-4a56-aa4d-066075b2a7a8" = "Teams Administrator"
    "eb1f4a8d-243a-41f0-9fbd-c7cdf6c5ef7c" = "Insights Administrator"
    "25a516ed-2fa0-40ea-a2d0-12923a21473a" = "Kaizala Administrator"
    "31392ffb-586c-42d1-9346-e59415a2cc4e" = "Exchange Recipient Administrator"
    "ac16e43d-7b2d-40e0-ac05-243ff356ab5b" = "Message Center Privacy Reader"
    "75941009-915a-4869-abe7-691bff18279e" = "Network Administrator"
    "644ef478-e28f-4e28-b9dc-3fdde9aa0b1f" = "Printer Administrator"
    "e8cef6f1-e4bd-4ea8-bc07-4b8d950f4477" = "Printer Technician"
    "7495fdc4-34c4-4d15-a289-98788ce399fd" = "Azure DevOps Administrator"
    "0526716b-113d-4c15-b2c8-68e3c22b9f80" = "Authentication Policy Administrator"
    "be2f45a1-457d-42af-a067-6ec1fa63bc45" = "External Identity Provider Administrator"
    "cf1c38e5-3621-4004-a7cb-879624dced7c" = "Azure Information Protection Administrator"
    "5f2222b1-57c3-48ba-8ad5-d4759f1fde6f" = "Security Operator"
    "74ef975b-6605-40af-a5d2-b9539d836353" = "Knowledge Administrator"
    "b5a8dcf3-09d5-43a9-a639-8e29ef291470" = "Knowledge Manager"
    "92b086b3-e367-4ef2-b869-1de128fb986e" = "Attribute Assignment Administrator"
    "58a13ea3-c632-46ae-9ee0-9c0d43cd7f3d" = "Attribute Assignment Reader"
    "ffd52fa5-98dc-465c-991d-fc073eb59f8f" = "Attribute Definition Administrator"
    "1d336d2c-4ae8-42ef-9711-b3604ce3fc2c" = "Attribute Definition Reader"
    "All" = "All Roles"
}

function Get-WellKnownApplications {
    <#
    .SYNOPSIS
        Returns the hashtable of well-known Microsoft application IDs
    #>
    return $script:WellKnownApps
}

function Get-WellKnownRoles {
    <#
    .SYNOPSIS
        Returns the hashtable of well-known directory role IDs
    #>
    return $script:WellKnownRoles
}

function Connect-Graph {
    <#
    .SYNOPSIS
        Authenticates to Microsoft Graph using device code flow

    .PARAMETER TenantId
        The tenant ID or domain name (e.g., contoso.onmicrosoft.com)

    .PARAMETER ClientId
        Optional client ID for a custom app registration. Defaults to Microsoft Graph PowerShell
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId = "common",

        [Parameter(Mandatory = $false)]
        [string]$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # Microsoft Graph PowerShell
    )

    $script:TenantId = $TenantId
    $scope = "https://graph.microsoft.com/.default offline_access"

    # Device code flow
    $deviceCodeUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode"
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    try {
        # Request device code
        $deviceCodeBody = @{
            client_id = $ClientId
            scope     = $scope
        }

        $deviceCodeResponse = Invoke-RestMethod -Uri $deviceCodeUrl -Method POST -Body $deviceCodeBody -ContentType "application/x-www-form-urlencoded"

        Write-Host "`n$($deviceCodeResponse.message)" -ForegroundColor Yellow
        Write-Host "`nWaiting for authentication..." -ForegroundColor Cyan

        # Poll for token
        $tokenBody = @{
            grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
            client_id   = $ClientId
            device_code = $deviceCodeResponse.device_code
        }

        $timeout = [DateTime]::Now.AddSeconds($deviceCodeResponse.expires_in)
        $interval = $deviceCodeResponse.interval

        while ([DateTime]::Now -lt $timeout) {
            Start-Sleep -Seconds $interval

            try {
                $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded"

                $script:GraphToken = $tokenResponse.access_token
                $script:TokenExpiry = [DateTime]::Now.AddSeconds($tokenResponse.expires_in - 300) # 5 min buffer

                Write-Host "Successfully authenticated to Microsoft Graph!" -ForegroundColor Green
                return $true
            }
            catch {
                $errorResponse = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($errorResponse.error -eq "authorization_pending") {
                    continue
                }
                elseif ($errorResponse.error -eq "authorization_declined") {
                    throw "Authentication was declined by the user."
                }
                elseif ($errorResponse.error -eq "expired_token") {
                    throw "Device code expired. Please try again."
                }
                else {
                    throw $_
                }
            }
        }

        throw "Authentication timed out."
    }
    catch {
        Write-Error "Failed to authenticate: $_"
        return $false
    }
}

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
        Makes a request to Microsoft Graph API with automatic pagination and retry logic

    .PARAMETER Uri
        The Graph API endpoint (can be relative or absolute)

    .PARAMETER Method
        HTTP method (GET, POST, etc.)

    .PARAMETER Body
        Request body for POST/PATCH requests

    .PARAMETER All
        If specified, follows pagination to retrieve all results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [ValidateSet("GET", "POST", "PATCH", "DELETE")]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [object]$Body,

        [Parameter(Mandatory = $false)]
        [switch]$All,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 2
    )

    if (-not $script:GraphToken) {
        throw "Not authenticated. Please run Connect-Graph first."
    }

    # Build full URI if relative
    if ($Uri -notmatch "^https://") {
        $Uri = "https://graph.microsoft.com/v1.0$Uri"
    }

    $headers = @{
        "Authorization" = "Bearer $($script:GraphToken)"
        "Content-Type"  = "application/json"
    }

    $allResults = @()
    $currentUri = $Uri
    $retryCount = 0

    do {
        try {
            $params = @{
                Uri     = $currentUri
                Method  = $Method
                Headers = $headers
            }

            if ($Body) {
                $params.Body = $Body | ConvertTo-Json -Depth 10
            }

            $response = Invoke-RestMethod @params
            $retryCount = 0

            if ($response.value) {
                $allResults += $response.value
            }
            else {
                $allResults += $response
            }

            # Check for pagination
            if ($All -and $response.'@odata.nextLink') {
                $currentUri = $response.'@odata.nextLink'
            }
            else {
                $currentUri = $null
            }
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__

            # Handle rate limiting
            if ($statusCode -eq 429) {
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                if (-not $retryAfter) { $retryAfter = $RetryDelaySeconds * [Math]::Pow(2, $retryCount) }

                Write-Warning "Rate limited. Waiting $retryAfter seconds..."
                Start-Sleep -Seconds $retryAfter
                $retryCount++

                if ($retryCount -ge $MaxRetries) {
                    throw "Max retries exceeded for rate limiting."
                }
                continue
            }

            # Handle transient errors
            if ($statusCode -in @(500, 502, 503, 504) -and $retryCount -lt $MaxRetries) {
                $retryCount++
                $delay = $RetryDelaySeconds * [Math]::Pow(2, $retryCount)
                Write-Warning "Transient error ($statusCode). Retrying in $delay seconds..."
                Start-Sleep -Seconds $delay
                continue
            }

            throw $_
        }
    } while ($currentUri)

    return $allResults
}

function Get-ConditionalAccessPolicies {
    <#
    .SYNOPSIS
        Retrieves all Conditional Access policies from the tenant
    #>
    [CmdletBinding()]
    param()

    Write-Host "Fetching Conditional Access Policies..." -ForegroundColor Cyan

    $policies = Invoke-GraphRequest -Uri "/identity/conditionalAccess/policies" -All

    Write-Host "Found $($policies.Count) Conditional Access Policies" -ForegroundColor Green

    return $policies
}

function Get-ApplicationDisplayName {
    <#
    .SYNOPSIS
        Resolves an application ID to its display name

    .PARAMETER AppId
        The application ID (GUID) to resolve
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId
    )

    # Check for special values
    if ($AppId -eq "All") { return "All cloud apps" }
    if ($AppId -eq "Office365") { return "Office 365" }
    if ($AppId -eq "MicrosoftAdminPortals") { return "Microsoft Admin Portals" }
    if ($AppId -eq "none") { return "None" }

    # Check cache
    if ($script:NameCache.Applications.ContainsKey($AppId)) {
        return $script:NameCache.Applications[$AppId]
    }

    # Check well-known apps
    if ($script:WellKnownApps.ContainsKey($AppId)) {
        $name = $script:WellKnownApps[$AppId]
        $script:NameCache.Applications[$AppId] = $name
        return $name
    }

    # Query Graph API
    try {
        # Try service principals first (most common)
        $sp = Invoke-GraphRequest -Uri "/servicePrincipals?`$filter=appId eq '$AppId'&`$select=displayName"
        if ($sp -and $sp.displayName) {
            $script:NameCache.Applications[$AppId] = $sp.displayName
            return $sp.displayName
        }

        # Fallback to applications
        $app = Invoke-GraphRequest -Uri "/applications?`$filter=appId eq '$AppId'&`$select=displayName"
        if ($app -and $app.displayName) {
            $script:NameCache.Applications[$AppId] = $app.displayName
            return $app.displayName
        }
    }
    catch {
        Write-Verbose "Could not resolve application ID: $AppId - $_"
    }

    # Return ID if unable to resolve
    $script:NameCache.Applications[$AppId] = "[Unknown: $AppId]"
    return "[Unknown: $AppId]"
}

function Get-UserDisplayName {
    <#
    .SYNOPSIS
        Resolves a user ID to their display name

    .PARAMETER UserId
        The user ID (GUID) to resolve
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    # Check for special values
    if ($UserId -eq "All") { return "All users" }
    if ($UserId -eq "GuestsOrExternalUsers") { return "Guests or external users" }
    if ($UserId -eq "None") { return "None" }

    # Check cache
    if ($script:NameCache.Users.ContainsKey($UserId)) {
        return $script:NameCache.Users[$UserId]
    }

    # Query Graph API
    try {
        $user = Invoke-GraphRequest -Uri "/users/$UserId`?`$select=displayName,userPrincipalName"
        if ($user) {
            $displayName = if ($user.displayName) { $user.displayName } else { $user.userPrincipalName }
            $script:NameCache.Users[$UserId] = $displayName
            return $displayName
        }
    }
    catch {
        Write-Verbose "Could not resolve user ID: $UserId - $_"
    }

    $script:NameCache.Users[$UserId] = "[Unknown User: $UserId]"
    return "[Unknown User: $UserId]"
}

function Get-GroupDisplayName {
    <#
    .SYNOPSIS
        Resolves a group ID to its display name

    .PARAMETER GroupId
        The group ID (GUID) to resolve
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    # Check for special values
    if ($GroupId -eq "All") { return "All groups" }

    # Check cache
    if ($script:NameCache.Groups.ContainsKey($GroupId)) {
        return $script:NameCache.Groups[$GroupId]
    }

    # Query Graph API
    try {
        $group = Invoke-GraphRequest -Uri "/groups/$GroupId`?`$select=displayName"
        if ($group -and $group.displayName) {
            $script:NameCache.Groups[$GroupId] = $group.displayName
            return $group.displayName
        }
    }
    catch {
        Write-Verbose "Could not resolve group ID: $GroupId - $_"
    }

    $script:NameCache.Groups[$GroupId] = "[Unknown Group: $GroupId]"
    return "[Unknown Group: $GroupId]"
}

function Get-RoleDisplayName {
    <#
    .SYNOPSIS
        Resolves a directory role template ID to its display name

    .PARAMETER RoleId
        The role template ID (GUID) to resolve
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RoleId
    )

    # Check for special values
    if ($RoleId -eq "All") { return "All roles" }

    # Check cache
    if ($script:NameCache.Roles.ContainsKey($RoleId)) {
        return $script:NameCache.Roles[$RoleId]
    }

    # Check well-known roles
    if ($script:WellKnownRoles.ContainsKey($RoleId)) {
        $name = $script:WellKnownRoles[$RoleId]
        $script:NameCache.Roles[$RoleId] = $name
        return $name
    }

    # Query Graph API
    try {
        $role = Invoke-GraphRequest -Uri "/directoryRoleTemplates/$RoleId"
        if ($role -and $role.displayName) {
            $script:NameCache.Roles[$RoleId] = $role.displayName
            return $role.displayName
        }
    }
    catch {
        Write-Verbose "Could not resolve role ID: $RoleId - $_"
    }

    $script:NameCache.Roles[$RoleId] = "[Unknown Role: $RoleId]"
    return "[Unknown Role: $RoleId]"
}

function Get-NamedLocationName {
    <#
    .SYNOPSIS
        Resolves a named location ID to its display name

    .PARAMETER LocationId
        The named location ID (GUID) to resolve
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LocationId
    )

    # Check for special values
    if ($LocationId -eq "All") { return "Any location" }
    if ($LocationId -eq "AllTrusted") { return "All trusted locations" }
    if ($LocationId -eq "00000000-0000-0000-0000-000000000000") { return "All Compliant Network locations" }

    # Check cache
    if ($script:NameCache.NamedLocations.ContainsKey($LocationId)) {
        return $script:NameCache.NamedLocations[$LocationId]
    }

    # Query Graph API
    try {
        $location = Invoke-GraphRequest -Uri "/identity/conditionalAccess/namedLocations/$LocationId"
        if ($location -and $location.displayName) {
            $script:NameCache.NamedLocations[$LocationId] = $location.displayName
            return $location.displayName
        }
    }
    catch {
        Write-Verbose "Could not resolve location ID: $LocationId - $_"
    }

    $script:NameCache.NamedLocations[$LocationId] = "[Unknown Location: $LocationId]"
    return "[Unknown Location: $LocationId]"
}

function Clear-NameCache {
    <#
    .SYNOPSIS
        Clears the name resolution cache
    #>
    $script:NameCache = @{
        Applications = @{}
        Users = @{}
        Groups = @{}
        Roles = @{}
        NamedLocations = @{}
    }
}

function Disconnect-Graph {
    <#
    .SYNOPSIS
        Clears the authentication token and cache
    #>
    $script:GraphToken = $null
    $script:TokenExpiry = $null
    $script:TenantId = $null
    Clear-NameCache
    Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Yellow
}

# Export functions
Export-ModuleMember -Function @(
    'Connect-Graph',
    'Disconnect-Graph',
    'Invoke-GraphRequest',
    'Get-ConditionalAccessPolicies',
    'Get-ApplicationDisplayName',
    'Get-UserDisplayName',
    'Get-GroupDisplayName',
    'Get-RoleDisplayName',
    'Get-NamedLocationName',
    'Get-WellKnownApplications',
    'Get-WellKnownRoles',
    'Clear-NameCache'
)
