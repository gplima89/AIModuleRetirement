# AIModuleRetirement

**`Get-AIFoundryInventory.ps1`** — a PowerShell script that inventories every Azure AI Foundry / Azure OpenAI / Cognitive Services model deployment across all your tenants and subscriptions, cross-references each one against the Azure model catalog's lifecycle/retirement data, and pulls the last 4 weeks of token consumption so you can gauge the real-world **impact** of each upcoming retirement.

It produces a single Excel workbook you can share with stakeholders: which models are retiring, when, and how much production traffic each one is actually carrying.

## What it does

For every accessible tenant and subscription (or the ones you specify), the script:

1. Uses **Azure Resource Graph** to discover:
   - `Microsoft.CognitiveServices/accounts` — Azure OpenAI / AI Foundry / Cognitive Services resources
   - `Microsoft.CognitiveServices/accounts/projects` — AI Foundry projects nested under an account
   - `Microsoft.MachineLearningServices/workspaces` — classic AI Foundry Hub / Project workspaces
2. Lists every **model deployment** on each Cognitive Services account.
3. Looks up the regional **model catalog** (`Get-AzCognitiveServicesModel`) to get each model's `LifecycleStatus` and `Deprecation` (retirement) dates for inference and fine-tuning.
4. Pulls the Azure Monitor **`TotalTokens`** metric (prompt + completion tokens) for each deployment, split into the last 4 individual weeks, so high-volume deployments on a retiring model can be prioritized for migration.
5. Produces an **Excel workbook** with three sheets:
   - **Model Deployments** — one row per deployment: tenant, subscription, resource group, account, Foundry project, location, creation dates, model name/version/format, SKU/capacity, lifecycle status, retirement dates, days-until-retirement, retirement flag, 4 weekly token columns, 4-week total, and notes. Conditionally formatted — red/yellow/green by retirement flag, and a color scale on the token columns.
   - **Hub-Project Workspaces** — one row per classic AI Foundry Hub/Project workspace found.
   - **Summary** — counts by retirement flag, counts by tenant/subscription, and a ranked list of the highest-impact retiring/deprecating/legacy deployments by 4-week token volume.

## Prerequisites

- **PowerShell 7+** (Windows PowerShell 5.1 also works).
- **PowerShell modules** (auto-installed to your user scope on first run if missing):
  `Az.Accounts`, `Az.CognitiveServices`, `Az.ResourceGraph`, `Az.Resources`, `Az.Monitor`, `ImportExcel`
- An active **Azure sign-in** (`Connect-AzAccount` or `az login`) — the script reuses your existing session and credentials. No secrets are stored or required.
- **Reader** access (or higher) on the subscriptions/tenants you want scanned, and **Monitoring Reader** (or higher) on each Cognitive Services account if you want token-usage data (otherwise those columns report `0`).

## Getting started

```powershell
# Clone the repo
git clone https://github.com/gplima89/AIModuleRetirement.git
cd AIModuleRetirement

# Sign in to Azure (if you aren't already)
Connect-AzAccount

# Run the inventory - scans every tenant/subscription you can access
.\Get-AIFoundryInventory.ps1
```

The script will open (or tell you the path to) a time-stamped `.xlsx` report next to the script, e.g. `AI-Foundry-Inventory-20261015-143455.xlsx`.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-TenantId` | `string[]` | all accessible tenants | One or more Entra ID tenant IDs to scan. Default: every tenant returned by `Get-AzTenant` (i.e. every tenant your signed-in identity has at least Reader-level visibility into). |
| `-SubscriptionId` | `string[]` | all enabled subscriptions | One or more subscription IDs to restrict the scan to. Default: every enabled subscription in the scanned tenant(s). |
| `-OutputPath` | `string` | time-stamped `.xlsx` next to the script | Path to the Excel report that will be created. |
| `-RetirementWarningDays` | `int` | `90` | Number of days before a model's inference retirement date to flag it as `RETIRING SOON`. |
| `-SkipWorkspaces` | `switch` | off | Skip discovery of `Microsoft.MachineLearningServices/workspaces` (Hub/Project workspaces). Speeds up the scan if you only care about Azure OpenAI / Cognitive Services deployments. |
| `-SkipTokenUsage` | `switch` | off | Skip the per-deployment weekly token-consumption lookup. Speeds up the scan if you only need retirement status, not usage/impact data. |

Full parameter documentation is also available via:

```powershell
Get-Help .\Get-AIFoundryInventory.ps1 -Full
```

## Examples

```powershell
# Scan everything you have access to (default)
.\Get-AIFoundryInventory.ps1

# Restrict scope to a specific tenant
.\Get-AIFoundryInventory.ps1 -TenantId '00000000-0000-0000-0000-000000000000' -OutputPath C:\Reports\inventory.xlsx

# Restrict scope to a specific subscription and tighten the "retiring soon" warning window
.\Get-AIFoundryInventory.ps1 -SubscriptionId '11111111-1111-1111-1111-111111111111' -RetirementWarningDays 60

# Faster scan: skip Hub/Project workspace discovery and token-usage lookups
.\Get-AIFoundryInventory.ps1 -SkipWorkspaces -SkipTokenUsage
```

## Output columns (Model Deployments sheet)

`Tenant`, `TenantId`, `SubscriptionName`, `SubscriptionId`, `ResourceGroup`, `AccountName`, `AccountKind`, `FoundryProject`, `Location`, `AccountCreatedOn`, `DeploymentName`, `ModelName`, `ModelVersion`, `ModelFormat`, `SkuName`, `Capacity`, `ProvisioningState`, `DeploymentCreatedOn`, `DeploymentModifiedOn`, `LifecycleStatus`, `InferenceRetirement`, `FineTuneRetirement`, `DaysUntilRetirement`, `RetirementFlag`, **`Wk1 Tokens (…)` … `Wk4 Tokens (…)`**, `Total4WeekTokens`, `Note`.

`RetirementFlag` values: `OK`, `PREVIEW`, `LEGACY`, `DEPRECATING`, `RETIRING SOON`, `RETIRED`, `UNKNOWN (not in catalog)`, `No deployments`.

## Notes & limitations

- Auth reuses your existing `Connect-AzAccount` / `az login` session. If a tenant can't be queried with the current token (e.g. MFA-restricted guest access), the script attempts `Connect-AzAccount -Tenant <id>` once and otherwise skips that tenant with a warning — it does not fail the whole run.
- Token usage is read directly from the `Microsoft.Insights/metrics` REST API (not `Get-AzMetric`, which has a known bug for this metric/dimension combination in some `Az.Monitor` versions).
- Accounts/kinds that don't emit the `TotalTokens` metric (e.g. `ContentSafety`) simply report `0` for all weeks instead of erroring.
- The Azure Monitor dimension-split query returns up to 500 series per account (`&top=500`), which comfortably covers typical deployment counts per account.

## License

[MIT](./LICENSE)
