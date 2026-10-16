# NOTE: Deliberately no "#Requires -Modules" directive here. #Requires aborts the ENTIRE script before
# any code runs if a listed module is missing, which would make the auto-install logic in the
# "Pre-flight validation" region below unreachable. Module presence/version is validated and
# auto-installed at runtime instead, with actionable error messages.
<#
.SYNOPSIS
    Inventories Azure AI Foundry / Azure OpenAI / Cognitive Services model deployments and Foundry
    hub-project workspaces across every tenant and subscription you have access to, and cross-references
    each deployed model against the Azure model catalog to report its lifecycle / retirement status.

.DESCRIPTION
    For every accessible tenant and subscription (or the ones you specify), the script:
      1. Uses Azure Resource Graph to discover:
           - Microsoft.CognitiveServices/accounts            (Azure OpenAI / AI Foundry / Cognitive Services resources)
           - Microsoft.CognitiveServices/accounts/projects    (AI Foundry projects nested under an account)
           - Microsoft.MachineLearningServices/workspaces     (classic AI Foundry Hub / Project workspaces)
      2. Lists every model deployment on each Cognitive Services account.
      3. Looks up the regional model catalog (Get-AzCognitiveServicesModel) to get each model's
         LifecycleStatus and Deprecation (retirement) dates for inference and fine-tuning.
      4. Pulls the "TotalTokens" Azure Monitor metric (prompt + completion tokens) for each deployment,
         split into the last 4 individual weeks, so you can gauge the real-world impact of a retirement
         (high-token-volume deployments on a retiring model are higher priority to migrate).
      5. Produces an Excel workbook with three sheets:
           - "Model Deployments"       : one row per deployment with full retirement detail + weekly token usage
           - "Hub-Project Workspaces"  : one row per Foundry Hub/Project workspace found
           - "Summary"                 : counts by retirement flag / tenant / subscription

.PARAMETER TenantId
    One or more Entra ID tenant IDs to scan. Default: every tenant returned by Get-AzTenant
    (i.e. every tenant your signed-in identity has at least Reader-level visibility into).

.PARAMETER SubscriptionId
    One or more subscription IDs to restrict the scan to. Default: every enabled subscription
    in the scanned tenant(s).

.PARAMETER OutputPath
    Path to the Excel report that will be created. Default: a time-stamped .xlsx file next to this script.

.PARAMETER RetirementWarningDays
    Number of days before a model's inference retirement date to flag it as "RETIRING SOON". Default: 90.

.PARAMETER SkipWorkspaces
    Skip discovery of Microsoft.MachineLearningServices/workspaces (Hub/Project workspaces). Use this to
    speed up the scan if you only care about Azure OpenAI / Cognitive Services deployments.

.PARAMETER SkipTokenUsage
    Skip the per-deployment weekly token-consumption lookup (Azure Monitor "TotalTokens" metric). Use this
    to speed up the scan if you only need retirement status, not usage/impact data.

.EXAMPLE
    .\Get-AIFoundryInventory.ps1
    Scans every tenant/subscription you can access and writes a report next to the script.

.EXAMPLE
    .\Get-AIFoundryInventory.ps1 -TenantId '00000000-0000-0000-0000-000000000000' -OutputPath C:\Reports\inventory.xlsx

.EXAMPLE
    .\Get-AIFoundryInventory.ps1 -SubscriptionId '11111111-1111-1111-1111-111111111111' -RetirementWarningDays 60

.NOTES
    Requires: Az.Accounts, Az.CognitiveServices, Az.ResourceGraph, Az.Resources, Az.Monitor, ImportExcel
             (auto-installed to CurrentUser scope if missing), PowerShell 5.1 or later (7+ recommended).
    Auth: reuses your existing Connect-AzAccount / az login session. If a tenant can't be queried with the
          current token, the script attempts Connect-AzAccount -Tenant <id> and otherwise skips it with a warning.
    Token usage requires "Monitoring Reader" (or higher) on each Cognitive Services account; accounts/kinds
    that don't emit the "TotalTokens" metric (e.g. ContentSafety) simply report 0 for all weeks.
    Pre-flight checks performed before any scanning begins: PowerShell version, required module
    presence/install/import, Azure sign-in, and that the signed-in account can enumerate at least one tenant.
#>
[CmdletBinding()]
param(
    [string[]] $TenantId,
    [string[]] $SubscriptionId,
    [string] $OutputPath = (Join-Path -Path $PSScriptRoot -ChildPath ("AI-Foundry-Inventory-{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))),
    [int] $RetirementWarningDays = 90,
    [switch] $SkipWorkspaces,
    [switch] $SkipTokenUsage
)

$ErrorActionPreference = 'Stop'

#region Pre-flight validation
Write-Host "Running pre-flight checks..." -ForegroundColor Cyan

# 1. PowerShell version - Az modules and the syntax used in this script require 5.1+ (7+ recommended).
$minPSVersion = [version]'5.1'
if ($PSVersionTable.PSVersion -lt $minPSVersion) {
    throw "Unsupported PowerShell version: $($PSVersionTable.PSVersion). This script requires PowerShell $minPSVersion or later (PowerShell 7+ recommended). Install PowerShell 7 from https://aka.ms/powershell-release-current and re-run this script with 'pwsh'."
}
Write-Host "  [OK] PowerShell version $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" -ForegroundColor Green

# 2. Required modules - install to CurrentUser scope if missing, then import. Collect failures so we can
#    report every missing/broken module at once instead of stopping at the first one.
$requiredModules = 'Az.Accounts', 'Az.CognitiveServices', 'Az.ResourceGraph', 'Az.Resources', 'Az.Monitor', 'ImportExcel'
$moduleFailures = [System.Collections.Generic.List[string]]::new()
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "  Module '$m' not found. Attempting installation to CurrentUser scope..." -ForegroundColor Yellow
        try {
            Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
        } catch {
            $moduleFailures.Add("$m (install failed: $($_.Exception.Message))")
            continue
        }
    }
    try {
        Import-Module -Name $m -ErrorAction Stop
        Write-Host "  [OK] Module '$m' v$((Get-Module -Name $m).Version)" -ForegroundColor Green
    } catch {
        $moduleFailures.Add("$m (import failed: $($_.Exception.Message))")
    }
}
if ($moduleFailures.Count -gt 0) {
    throw "The following required module(s) could not be installed/imported:`n  - $($moduleFailures -join "`n  - ")`n`nTry installing manually with: Install-Module -Name <ModuleName> -Scope CurrentUser -Force"
}

# 3. Azure authentication - reuse an existing session if present, otherwise prompt interactively.
Write-Host "  Checking Azure sign-in..." -ForegroundColor Cyan
$azContext = Get-AzContext
if (-not $azContext -or -not $azContext.Account) {
    Write-Host "  No active Azure session found. Launching interactive sign-in (Connect-AzAccount)..." -ForegroundColor Yellow
    try {
        Connect-AzAccount -ErrorAction Stop | Out-Null
        $azContext = Get-AzContext
    } catch {
        throw "Azure sign-in failed: $($_.Exception.Message). Run 'Connect-AzAccount' manually and re-run this script."
    }
}
if (-not $azContext -or -not $azContext.Account) {
    throw "Unable to establish an authenticated Azure session. Run 'Connect-AzAccount' manually and re-run this script."
}
Write-Host "  [OK] Signed in as $($azContext.Account.Id) (home tenant: $($azContext.Tenant.Id))" -ForegroundColor Green

# 4. Minimal access check - confirm the signed-in identity can actually enumerate at least one tenant.
#    This catches cases like a disabled account, a token with no directory read access, or no network
#    route to Azure AD before we spend time looping through tenants/subscriptions later in the script.
try {
    $preflightTenants = @(Get-AzTenant -ErrorAction Stop)
} catch {
    throw "Could not enumerate Entra ID tenants for the signed-in account ($($azContext.Account.Id)): $($_.Exception.Message). Confirm you are signed in with the correct account, have network access to Azure, and have at least Reader access to an Azure subscription."
}
if ($preflightTenants.Count -eq 0) {
    throw "The signed-in account ($($azContext.Account.Id)) has no accessible Entra ID tenants. Confirm you are signed in with the correct account and have at least Reader access to an Azure subscription."
}
Write-Host "  [OK] $($preflightTenants.Count) accessible tenant(s) detected for this account" -ForegroundColor Green
Write-Host "  Note: per-subscription access (Reader) and per-account Monitoring Reader (for token usage) are validated as each resource is scanned; any gaps are reported as warnings, not fatal errors." -ForegroundColor DarkGray
#endregion

#region Helpers
function Invoke-FullGraphQuery {
    <# Pages through Search-AzGraph results so inventories larger than one page are not truncated. #>
    param(
        [Parameter(Mandatory)] [string] $Query,
        [Parameter(Mandatory)] [string[]] $Subscription
    )
    $all = [System.Collections.Generic.List[object]]::new()
    $skipToken = $null
    do {
        $params = @{
            Query        = $Query
            Subscription = $Subscription
            First        = 1000
        }
        if ($skipToken) { $params['SkipToken'] = $skipToken }
        $page = Search-AzGraph @params
        if ($page) { $all.AddRange(@($page)) }
        $skipToken = $page.SkipToken
    } while ($skipToken)
    return $all
}

$script:ModelCatalogCache = @{}
function Get-ModelCatalogForLocation {
    param([Parameter(Mandatory)] [string] $Location)
    if (-not $script:ModelCatalogCache.ContainsKey($Location)) {
        try {
            Write-Host "    Fetching model catalog for location '$Location' (cached per run)..." -ForegroundColor DarkGray
            $script:ModelCatalogCache[$Location] = @(Get-AzCognitiveServicesModel -Location $Location -ErrorAction Stop)
        } catch {
            Write-Warning "Could not retrieve model catalog for location '$Location': $($_.Exception.Message)"
            $script:ModelCatalogCache[$Location] = @()
        }
    }
    return $script:ModelCatalogCache[$Location]
}

function Get-RetirementFlag {
    param(
        [string] $LifecycleStatus,
        [Nullable[datetime]] $InferenceRetireDate,
        [int] $WarningDays,
        [bool] $FoundInCatalog
    )
    if (-not $FoundInCatalog) { return 'UNKNOWN (not in catalog)' }
    if ($LifecycleStatus -eq 'Retired') { return 'RETIRED' }
    if ($InferenceRetireDate) {
        $daysUntil = ($InferenceRetireDate - (Get-Date).ToUniversalTime()).TotalDays
        if ($daysUntil -lt 0) { return 'RETIRED' }
        if ($daysUntil -le $WarningDays) { return 'RETIRING SOON' }
        return 'OK'
    }
    switch ($LifecycleStatus) {
        'Deprecating' { return 'DEPRECATING' }
        'Legacy'      { return 'LEGACY' }
        'Preview'     { return 'PREVIEW' }
        default       { return 'OK' }
    }
}

# Reference point for the "past 4 weeks" token-usage window, computed once so every row in this run
# uses identical week boundaries. Week 1 = most recent 7 days, Week 4 = oldest of the 4 weeks.
$script:TokenUsageNow = (Get-Date).ToUniversalTime()
$script:TokenWeekRanges = for ($w = 1; $w -le 4; $w++) {
    $weekEnd   = $script:TokenUsageNow.AddDays(-7 * ($w - 1))
    $weekStart = $script:TokenUsageNow.AddDays(-7 * $w)
    [pscustomobject]@{
        Index = $w - 1
        Start = $weekStart
        End   = $weekEnd
        Label = 'Wk{0} Tokens ({1:MMM dd} - {2:MMM dd})' -f $w, $weekStart, $weekEnd
    }
}

function Get-AccountWeeklyTokenUsage {
    <#
    Queries the Azure Monitor "TotalTokens" metric (prompt + completion tokens) for a Cognitive Services
    account over the last 28 days, split by the ModelDeploymentName dimension, and buckets the daily
    values into the 4 weekly windows defined in $script:TokenWeekRanges.

    Uses the Microsoft.Insights/metrics REST API directly via Invoke-AzRestMethod rather than
    Get-AzMetric, because Get-AzMetric 6.0.1 throws a null-reference error for several Cognitive
    Services metric/dimension combinations.

    Returns a hashtable: DeploymentName -> double[4] (index 0 = Week1/most recent ... index 3 = Week4/oldest).
    Accounts/kinds that don't support the TotalTokens metric (e.g. ContentSafety) return an empty hashtable.
    #>
    param([Parameter(Mandatory)] [string] $ResourceId)

    $usage = @{}
    try {
        $start = $script:TokenWeekRanges[3].Start
        $end   = $script:TokenUsageNow
        $timespan = '{0:yyyy-MM-ddTHH:mm:ssZ}/{1:yyyy-MM-ddTHH:mm:ssZ}' -f $start, $end
        $filterEnc = [uri]::EscapeDataString("ModelDeploymentName eq '*'")
        $path = "$ResourceId/providers/Microsoft.Insights/metrics?api-version=2018-01-01&metricnames=TotalTokens&timespan=$timespan&interval=P1D&aggregation=Total&`$filter=$filterEnc&top=500"

        $resp = Invoke-AzRestMethod -Path $path -Method GET -ErrorAction Stop
        if ($resp.StatusCode -ne 200) {
            Write-Verbose "TotalTokens metric not available for $ResourceId (HTTP $($resp.StatusCode)): $($resp.Content)"
            return $usage
        }

        $json = $resp.Content | ConvertFrom-Json
        $series = $json.value[0].timeseries
        foreach ($ts in $series) {
            $depName = ($ts.metadatavalues | Where-Object { $_.name.value -eq 'ModelDeploymentName' } | Select-Object -First 1).value
            if (-not $depName) { continue }

            $buckets = [double[]]@(0, 0, 0, 0)
            foreach ($point in $ts.data) {
                if ($null -eq $point.total) { continue }
                $pointTime = [datetime]::Parse($point.timeStamp, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
                foreach ($wk in $script:TokenWeekRanges) {
                    if ($pointTime -ge $wk.Start -and $pointTime -lt $wk.End) {
                        $buckets[$wk.Index] += [double]$point.total
                        break
                    }
                }
            }
            $usage[$depName] = $buckets
        }
    } catch {
        Write-Verbose "Could not retrieve token usage for $ResourceId : $($_.Exception.Message)"
    }
    return $usage
}
#endregion

#region Determine tenants to scan
# Reuses the tenant list already fetched during pre-flight validation above - no need to call Get-AzTenant again.
$allTenants = $preflightTenants | Select-Object @{n = 'Id'; e = { $_.Id } }, @{n = 'Name'; e = { $_.Name } }
if ($TenantId) {
    $tenantsToScan = $TenantId | ForEach-Object {
        $tid = $_
        $known = $allTenants | Where-Object { $_.Id -eq $tid }
        if ($known) { $known } else { [pscustomobject]@{ Id = $tid; Name = $tid } }
    }
} else {
    $tenantsToScan = $allTenants
}
Write-Host "Tenants to scan: $($tenantsToScan.Count)" -ForegroundColor Cyan
#endregion

$deploymentResults = [System.Collections.Generic.List[object]]::new()
$workspaceResults  = [System.Collections.Generic.List[object]]::new()

$graphQuery = @'
resources
| where type in~ ('microsoft.cognitiveservices/accounts','microsoft.cognitiveservices/accounts/projects','microsoft.machinelearningservices/workspaces')
| project id, name, type, kind, resourceGroup, subscriptionId, location, properties, systemData, tags
'@

foreach ($tenant in $tenantsToScan) {
    $tId = $tenant.Id
    Write-Host "`n=== Tenant: $($tenant.Name) ($tId) ===" -ForegroundColor Cyan

    try {
        $subs = Get-AzSubscription -TenantId $tId -ErrorAction Stop
    } catch {
        Write-Warning "Could not list subscriptions for tenant $tId with the current session. Trying interactive login for this tenant..."
        try {
            Connect-AzAccount -Tenant $tId -ErrorAction Stop | Out-Null
            $subs = Get-AzSubscription -TenantId $tId -ErrorAction Stop
        } catch {
            Write-Warning "Skipping tenant $tId - unable to authenticate or list subscriptions: $($_.Exception.Message). This usually means the signed-in account has no role assignment (e.g. Reader) in this tenant."
            continue
        }
    }

    $subs = $subs | Where-Object { $_.State -eq 'Enabled' }
    if ($SubscriptionId) {
        $subs = $subs | Where-Object { $_.Id -in $SubscriptionId }
    }
    if (-not $subs) {
        Write-Host "  No matching enabled subscriptions in this tenant." -ForegroundColor DarkGray
        continue
    }

    try {
        Set-AzContext -Tenant $tId -Subscription $subs[0].Id -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Could not set context in tenant $tId - $($_.Exception.Message)"
        continue
    }

    Write-Host "  Querying Azure Resource Graph across $($subs.Count) subscription(s)..." -ForegroundColor DarkCyan
    try {
        $aiResources = Invoke-FullGraphQuery -Query $graphQuery -Subscription $subs.Id
    } catch {
        Write-Warning "Resource Graph query failed in tenant $tId - $($_.Exception.Message). Confirm the signed-in account has at least Reader access on these subscription(s)."
        continue
    }

    $accountsByType   = $aiResources | Where-Object { $_.type -eq 'microsoft.cognitiveservices/accounts' }
    $projectsByType   = $aiResources | Where-Object { $_.type -eq 'microsoft.cognitiveservices/accounts/projects' }
    $workspacesByType = $aiResources | Where-Object { $_.type -eq 'microsoft.machinelearningservices/workspaces' }

    Write-Host "  Found $($accountsByType.Count) Cognitive Services/AI Foundry account(s), $($projectsByType.Count) project(s), $($workspacesByType.Count) ML workspace(s)." -ForegroundColor DarkCyan

    # Index nested Foundry projects by their parent account name
    $projectsByAccount = @{}
    foreach ($p in $projectsByType) {
        $parentName = ($p.name -split '/')[0]
        if (-not $projectsByAccount.ContainsKey($parentName)) { $projectsByAccount[$parentName] = [System.Collections.Generic.List[object]]::new() }
        $projectsByAccount[$parentName].Add($p)
    }
    $subsById = @{}
    foreach ($s in $subs) { $subsById[$s.Id] = $s }

    foreach ($acct in $accountsByType) {
        $sub = $subsById[$acct.subscriptionId]
        if (-not $sub) { continue }

        try {
            Set-AzContext -Tenant $tId -Subscription $sub.Id -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "Could not switch context to subscription $($sub.Id): $($_.Exception.Message)"
            continue
        }

        $acctProjects = if ($projectsByAccount.ContainsKey($acct.name)) { $projectsByAccount[$acct.name] } else { @() }
        $projectNames = ($acctProjects | ForEach-Object { (($_.name -split '/')[1]) }) -join '; '

        Write-Host "    Account: $($acct.name) [$($acct.kind)] - $($acct.resourceGroup)" -ForegroundColor Gray

        try {
            # NOTE: Get-AzCognitiveServicesAccountDeployment returns its results as a single
            # List<Deployment> object (NoEnumerate). Do NOT wrap the call in @(...) - that would
            # add an extra array layer and break per-item enumeration below for multi-deployment accounts.
            $deployments = Get-AzCognitiveServicesAccountDeployment -ResourceGroupName $acct.resourceGroup -AccountName $acct.name -ErrorAction Stop
        } catch {
            Write-Warning "    Could not list deployments for account $($acct.name): $($_.Exception.Message). Confirm the signed-in account has at least Cognitive Services Reader on this account."
            $deployments = [System.Collections.Generic.List[object]]::new()
        }

        if ($deployments.Count -eq 0) {
            $noUsageRow = [ordered]@{
                Tenant               = $tenant.Name
                TenantId             = $tId
                SubscriptionName     = $sub.Name
                SubscriptionId       = $sub.Id
                ResourceGroup        = $acct.resourceGroup
                AccountName          = $acct.name
                AccountKind          = $acct.kind
                FoundryProject       = $projectNames
                Location             = $acct.location
                AccountCreatedOn     = $acct.systemData.createdAt
                DeploymentName       = $null
                ModelName            = $null
                ModelVersion         = $null
                ModelFormat          = $null
                SkuName              = $null
                Capacity             = $null
                ProvisioningState    = $null
                DeploymentCreatedOn  = $null
                DeploymentModifiedOn = $null
                LifecycleStatus      = $null
                InferenceRetirement  = $null
                FineTuneRetirement   = $null
                DaysUntilRetirement  = $null
                RetirementFlag       = 'No deployments'
            }
            foreach ($wk in $script:TokenWeekRanges) { $noUsageRow[$wk.Label] = $null }
            $noUsageRow['Total4WeekTokens'] = $null
            $noUsageRow['Note'] = $null
            $deploymentResults.Add([pscustomobject]$noUsageRow)
            continue
        }

        $catalog = Get-ModelCatalogForLocation -Location $acct.location

        $tokenUsageByDeployment = @{}
        if (-not $SkipTokenUsage) {
            $tokenUsageByDeployment = Get-AccountWeeklyTokenUsage -ResourceId $acct.id
        }

        foreach ($dep in $deployments) {
            $modelName    = $dep.Properties.Model.Name
            $modelVersion = $dep.Properties.Model.Version
            $modelFormat  = $dep.Properties.Model.Format

            $catalogMatch = $catalog | Where-Object {
                $_.ModelProperty.Name -ieq $modelName -and $_.ModelProperty.Version -ieq $modelVersion
            } | Select-Object -First 1

            $catalogNote = $null
            if (-not $catalogMatch) {
                $catalogMatch = $catalog | Where-Object { $_.ModelProperty.Name -ieq $modelName } |
                    Sort-Object { $_.ModelProperty.SystemData.CreatedAt } -Descending | Select-Object -First 1
                if ($catalogMatch) { $catalogNote = "Exact version '$modelVersion' not found in '$($acct.location)' catalog; showing closest match (version '$($catalogMatch.ModelProperty.Version)')." }
            }

            $lifecycle       = $catalogMatch.ModelProperty.LifecycleStatus
            $inferenceRetire = $catalogMatch.ModelProperty.Deprecation.Inference
            $fineTuneRetire  = $catalogMatch.ModelProperty.Deprecation.FineTune
            $inferenceRetireDate = if ($inferenceRetire) { [datetime]$inferenceRetire } else { $null }

            $daysUntil = if ($inferenceRetireDate) { [math]::Round(($inferenceRetireDate - (Get-Date).ToUniversalTime()).TotalDays) } else { $null }
            $flag = Get-RetirementFlag -LifecycleStatus $lifecycle -InferenceRetireDate $inferenceRetireDate -WarningDays $RetirementWarningDays -FoundInCatalog ([bool]$catalogMatch)

            $weeklyTokens = if ($tokenUsageByDeployment.ContainsKey($dep.Name)) { $tokenUsageByDeployment[$dep.Name] } else { [double[]]@(0, 0, 0, 0) }

            $row = [ordered]@{
                Tenant               = $tenant.Name
                TenantId             = $tId
                SubscriptionName     = $sub.Name
                SubscriptionId       = $sub.Id
                ResourceGroup        = $acct.resourceGroup
                AccountName          = $acct.name
                AccountKind          = $acct.kind
                FoundryProject       = $projectNames
                Location             = $acct.location
                AccountCreatedOn     = $acct.systemData.createdAt
                DeploymentName       = $dep.Name
                ModelName            = $modelName
                ModelVersion         = $modelVersion
                ModelFormat          = $modelFormat
                SkuName              = $dep.Sku.Name
                Capacity             = $dep.Sku.Capacity
                ProvisioningState    = $dep.Properties.ProvisioningState
                DeploymentCreatedOn  = $dep.SystemData.CreatedAt
                DeploymentModifiedOn = $dep.SystemData.LastModifiedAt
                LifecycleStatus      = $lifecycle
                InferenceRetirement  = $inferenceRetire
                FineTuneRetirement   = $fineTuneRetire
                DaysUntilRetirement  = $daysUntil
                RetirementFlag       = $flag
            }
            for ($i = 0; $i -lt $script:TokenWeekRanges.Count; $i++) { $row[$script:TokenWeekRanges[$i].Label] = $weeklyTokens[$i] }
            $row['Total4WeekTokens'] = ($weeklyTokens | Measure-Object -Sum).Sum
            $row['Note'] = $catalogNote

            $deploymentResults.Add([pscustomobject]$row)
        }
    }

    if (-not $SkipWorkspaces) {
        foreach ($ws in $workspacesByType) {
            $sub = $subsById[$ws.subscriptionId]
            if (-not $sub) { continue }
            $wsKind = if ($ws.kind) { $ws.kind } elseif ($ws.properties.kind) { $ws.properties.kind } else { 'Default' }
            $workspaceResults.Add([pscustomobject]@{
                Tenant           = $tenant.Name
                TenantId         = $tId
                SubscriptionName = $sub.Name
                SubscriptionId   = $sub.Id
                ResourceGroup    = $ws.resourceGroup
                WorkspaceName    = ($ws.name -split '/')[0]
                WorkspaceKind    = $wsKind
                Location         = $ws.location
                CreatedOn        = $ws.systemData.createdAt
                LastModifiedOn   = $ws.systemData.lastModifiedAt
            })
        }
    }
}

#region Export to Excel
Write-Host "`nWriting report to: $OutputPath" -ForegroundColor Cyan

if (Test-Path -Path $OutputPath) { Remove-Item -Path $OutputPath -Force }

$ctRetired      = New-ConditionalText -Text 'RETIRED'        -BackgroundColor '#FFC7CE' -ConditionalTextColor '#9C0006'
$ctSoon         = New-ConditionalText -Text 'RETIRING SOON'  -BackgroundColor '#FFEB9C' -ConditionalTextColor '#9C6500'
$ctDeprecating  = New-ConditionalText -Text 'DEPRECATING'    -BackgroundColor '#FFEB9C' -ConditionalTextColor '#9C6500'
$ctLegacy       = New-ConditionalText -Text 'LEGACY'         -BackgroundColor '#FFEB9C' -ConditionalTextColor '#9C6500'
$ctUnknown      = New-ConditionalText -Text 'UNKNOWN'        -BackgroundColor '#D9D9D9' -ConditionalTextColor '#404040'
$ctOk           = New-ConditionalText -Text 'OK'             -BackgroundColor '#C6EFCE' -ConditionalTextColor '#006100'

if ($deploymentResults.Count -gt 0) {
    $excelPkg = $deploymentResults | Sort-Object Tenant, SubscriptionName, AccountName, DeploymentName |
        Export-Excel -Path $OutputPath -WorksheetName 'Model Deployments' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -TableStyle Medium2 `
            -ConditionalText @($ctRetired, $ctSoon, $ctDeprecating, $ctLegacy, $ctUnknown, $ctOk) -PassThru

    # Use fixed token bands so colors are consistent across weekly/total columns and reports.
    $ws = $excelPkg.Workbook.Worksheets['Model Deployments']
    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column
    $tokenColumnNames = @($script:TokenWeekRanges.Label) + 'Total4WeekTokens'
    if ($lastRow -gt 1) {
        foreach ($colName in $tokenColumnNames) {
            for ($c = 1; $c -le $lastCol; $c++) {
                if ($ws.Cells[1, $c].Text -eq $colName) {
                    $colLetter = ([OfficeOpenXml.ExcelCellAddress]::new(1, $c).Address) -replace '\d', ''
                    $address = "$($colLetter)2:$($colLetter)$lastRow"
                    $cell = "$($colLetter)2"
                    Add-ConditionalFormatting -Worksheet $ws -Address $address -RuleType Expression `
                        -ConditionValue "=AND(ISNUMBER($cell),$cell>=0,$cell<=50000)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml('#C6EFCE')) -ForegroundColor ([System.Drawing.ColorTranslator]::FromHtml('#006100'))
                    Add-ConditionalFormatting -Worksheet $ws -Address $address -RuleType Expression `
                        -ConditionValue "=AND(ISNUMBER($cell),$cell>50000,$cell<=250000)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml('#FFEB9C')) -ForegroundColor ([System.Drawing.ColorTranslator]::FromHtml('#9C6500'))
                    Add-ConditionalFormatting -Worksheet $ws -Address $address -RuleType Expression `
                        -ConditionValue "=AND(ISNUMBER($cell),$cell>250000,$cell<5000000)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml('#FCE4D6')) -ForegroundColor ([System.Drawing.ColorTranslator]::FromHtml('#C65911'))
                    Add-ConditionalFormatting -Worksheet $ws -Address $address -RuleType Expression `
                        -ConditionValue "=AND(ISNUMBER($cell),$cell>=5000000)" -BackgroundColor ([System.Drawing.ColorTranslator]::FromHtml('#FFC7CE')) -ForegroundColor ([System.Drawing.ColorTranslator]::FromHtml('#9C0006'))
                    break
                }
            }
        }
    }
    Close-ExcelPackage -ExcelPackage $excelPkg
} else {
    [pscustomobject]@{ Message = 'No Cognitive Services / AI Foundry accounts or deployments were found in the scanned scope.' } |
        Export-Excel -Path $OutputPath -WorksheetName 'Model Deployments' -AutoSize -BoldTopRow
}

if (-not $SkipWorkspaces) {
    if ($workspaceResults.Count -gt 0) {
        $workspaceResults | Sort-Object Tenant, SubscriptionName, WorkspaceName |
            Export-Excel -Path $OutputPath -WorksheetName 'Hub-Project Workspaces' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -TableStyle Medium3
    } else {
        [pscustomobject]@{ Message = 'No Microsoft.MachineLearningServices/workspaces (Hub/Project) were found in the scanned scope.' } |
            Export-Excel -Path $OutputPath -WorksheetName 'Hub-Project Workspaces' -AutoSize -BoldTopRow
    }
}

$summary = $deploymentResults | Group-Object RetirementFlag | Sort-Object Count -Descending |
    Select-Object @{n = 'RetirementFlag'; e = { $_.Name } }, @{n = 'DeploymentCount'; e = { $_.Count } }
$summaryByTenant = $deploymentResults | Group-Object Tenant, SubscriptionName |
    Select-Object @{n = 'Tenant / Subscription'; e = { $_.Name } }, @{n = 'DeploymentCount'; e = { $_.Count } }
$impactFlags = @('RETIRED', 'RETIRING SOON', 'DEPRECATING', 'LEGACY')
$impactRanking = $deploymentResults | Where-Object { $_.RetirementFlag -in $impactFlags -and $null -ne $_.Total4WeekTokens } |
    Sort-Object Total4WeekTokens -Descending | Select-Object -First 25 |
    Select-Object Tenant, SubscriptionName, AccountName, DeploymentName, ModelName, ModelVersion, RetirementFlag, DaysUntilRetirement, Total4WeekTokens

if ($summary.Count -gt 0) {
    $summary | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow -Title 'Retirement status summary' -TitleBold -StartRow 1
    $summaryByTenant | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow -Title 'Deployments by tenant/subscription' -TitleBold -StartRow ($summary.Count + 4)
    if ($impactRanking.Count -gt 0) {
        $impactStartRow = $summary.Count + $summaryByTenant.Count + 8
        $impactRanking | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -BoldTopRow `
            -Title 'Highest-impact retiring/deprecating/legacy deployments (by 4-week token volume)' -TitleBold -StartRow $impactStartRow
    }
}
#endregion

Write-Host "`nDone." -ForegroundColor Green
Write-Host "  Accounts/deployments scanned : $($deploymentResults.Count) deployment row(s)" -ForegroundColor Green
Write-Host "  Hub/Project workspaces found  : $($workspaceResults.Count)" -ForegroundColor Green
Write-Host "  Report                        : $OutputPath" -ForegroundColor Green

$retiredOrSoon = $deploymentResults | Where-Object { $_.RetirementFlag -in @('RETIRED', 'RETIRING SOON') }
if ($retiredOrSoon.Count -gt 0) {
    Write-Host "`n⚠ $($retiredOrSoon.Count) deployment(s) are RETIRED or RETIRING SOON - see the report for details." -ForegroundColor Yellow
}

if ($impactRanking.Count -gt 0) {
    Write-Host "`nTop retirement-impact deployments (highest 4-week token volume) - see 'Summary' sheet for the full list:" -ForegroundColor Yellow
    $impactRanking | Select-Object -First 5 | ForEach-Object {
        Write-Host ("  [{0}] {1} / {2} -> {3}  v{4}  {5} tokens (last 4 wks)" -f $_.RetirementFlag, $_.AccountName, $_.DeploymentName, $_.ModelName, $_.ModelVersion, ('{0:N0}' -f $_.Total4WeekTokens)) -ForegroundColor Yellow
    }
}

return $OutputPath
