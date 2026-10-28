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
5. Produces an **Excel workbook** with up to three sheets:
   - **Model Deployments** — one row per deployment: tenant, subscription, resource group, account, Foundry project, location, creation dates, model name/version/format, SKU/capacity, lifecycle status, retirement dates, days-until-retirement, retirement flag, 4 weekly token columns, 4-week total, notes, and the account's Azure **tags** (a combined `Tags` column plus one `Tag: <key>` column per key — useful for cost center, business unit, owner, etc.). Conditionally formatted — red/yellow/green by retirement flag, and fixed color bands on all weekly token columns and the 4-week total: **green** for 0–50,000, **yellow** for >50,000–250,000, **orange** for >250,000–<5,000,000, and **red** for 5,000,000+. Blank and nonnumeric token cells are not colored.
   - **Hub-Project Workspaces** — one row per classic AI Foundry Hub/Project workspace found, including its tags.
   - **Summary** — counts by retirement flag, counts by tenant/subscription, and a ranked list of the highest-impact retiring/deprecating/legacy deployments by 4-week token volume.

## Prerequisites

- **PowerShell 5.1+** (PowerShell 7+ recommended). Checked automatically on every run.
- **PowerShell modules** (checked and auto-installed to your user scope on first run if missing):
  `Az.Accounts`, `Az.CognitiveServices`, `Az.ResourceGraph`, `Az.Resources`, `Az.Monitor`, `ImportExcel`
- An active **Azure sign-in** (`Connect-AzAccount` or `az login`) — the script reuses your existing session and credentials, or prompts you to sign in interactively if none is found. No secrets are stored or required.
- **Reader** access (or higher) on the subscriptions/tenants you want scanned, and **Monitoring Reader** (or higher) on each Cognitive Services account if you want token-usage data (otherwise those columns report `0`).

### Pre-flight checks

Every run starts with a pre-flight validation pass that fails fast with an actionable message instead of partway through a long scan:

1. **PowerShell version** — must be 5.1 or later.
2. **Required modules** — presence, auto-install (CurrentUser scope), and import; every missing/broken module is reported together, not just the first one.
3. **Azure sign-in** — reuses an existing session or launches `Connect-AzAccount` interactively; fails clearly if sign-in doesn't succeed.
4. **Minimal access check** — confirms the signed-in account can enumerate at least one Entra ID tenant before the scan begins.

Per-subscription (Reader) and per-account (Monitoring Reader) access gaps discovered *during* the scan are reported as warnings and skipped, not fatal errors — see [Notes & limitations](#notes--limitations).

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

The script saves a time-stamped `.xlsx` report next to the script, e.g. `AI-Foundry-Inventory-20261015-143455.xlsx`, and prints and returns its path. Open that file in Excel to explore the results. If you supply `-OutputPath`, an existing file at that path is replaced.

## Understanding the report

The workbook is a **point-in-time inventory and migration-triage report**. It helps answer:

- Which deployed model versions are retired or approaching retirement?
- Where are those deployments located, and which tenant/subscription should investigate them?
- Which affected deployments have the most recent token consumption?
- Is usage increasing, decreasing, or still present as a retirement date approaches?

The script reads Azure inventory, catalog, and metric data; it does **not** upgrade models, change deployments, or migrate application traffic. The workbook contains exported values, not a live connection to Azure. Rerun the script to refresh it.

### What the workbook looks like

| Worksheet | Layout | How to use it |
|---|---|---|
| **Model Deployments** | A wide, filterable table with bold headers, a frozen top row, and one row per model deployment. Rows are initially sorted by tenant, subscription, account, and deployment name. The last columns hold the account's Azure tags (`Tags` plus one `Tag: <key>` column per tag key). | Your detailed investigation view: filter by retirement flag, locate the deployment, and compare deadlines with usage. Filter on `Tag: …` columns (e.g. `Tag: CostCenter`, `Tag: BusinessUnit`) to route work to the right owner. Accounts without listed deployments have a placeholder row instead. |
| **Hub-Project Workspaces** | A separate inventory table with tenant/subscription, resource group, workspace name/kind, location, creation/modification dates, and the workspace's Azure tags. | Locate classic Hub/Project workspaces. This sheet does not contain model-retirement or token-usage assessments. It is omitted with `-SkipWorkspaces`. |
| **Summary** | Stacked sections for retirement-status counts, tenant/subscription counts, and up to 25 affected deployments ranked by 4-week tokens (including their `Tags`). | Start here for a quick overview and a shortlist for migration planning, then find the corresponding rows in **Model Deployments**. |

If no accounts or deployments are found in the scanned scope, **Model Deployments** contains an explanatory message and **Summary** is not created. If workspace discovery runs but finds nothing, its sheet also contains a message.

> **Counts are inventory-row counts.** The summary's `DeploymentCount` includes the `No deployments` placeholder row for each account without listed deployments. It is not always a count of actual model deployments. A deployment-listing failure can also produce such a placeholder; check the run's warnings before concluding that an account is empty.

### Reading a deployment row

| Column group | Important fields | What to look for |
|---|---|---|
| Resource identity | `Tenant`, `SubscriptionName`, `ResourceGroup`, `AccountName`, `DeploymentName`, `Location` | Use these together to locate the resource; deployment names alone may repeat across accounts. |
| Account/project context | `AccountKind`, `FoundryProject` | `FoundryProject` lists discovered projects under the account, separated by semicolons. It is account context, not a definitive deployment-to-project ownership mapping. |
| Deployed model | `ModelName`, `ModelVersion`, `ModelFormat`, `SkuName`, `Capacity`, `ProvisioningState` | Identify the exact deployed version and provisioning state. `Capacity` is the deployment SKU's capacity value, not measured token consumption. |
| Lifecycle and deadlines | `LifecycleStatus`, `InferenceRetirement`, `FineTuneRetirement`, `DaysUntilRetirement`, `RetirementFlag` | Distinguish the catalog lifecycle from the script's calculated urgency flag. Inference retirement drives the deadline assessment; fine-tuning retirement is shown separately for reference. |
| Usage | Four `Wk… Tokens (…)` columns and `Total4WeekTokens` | Assess recent activity and relative migration impact, using the time windows and caveats below. |
| Data qualification | `Note` | Read this before acting on retirement dates. It identifies cases where the exact deployed version was absent and another catalog version was used. |
| Ownership / chargeback | `Tags`, `Tag: <key>` columns | Identify cost center, business unit, owner, environment, etc. See [Resource tags](#resource-tags) below. |

`DaysUntilRetirement` is the rounded number of days between the run time and the inference-retirement date: positive means ahead, negative means past, and blank means no inference date was available. A value of `0` can mean slightly before or after the deadline because of rounding; the flag uses the unrounded interval.

### Resource tags

Azure tags (for example `CostCenter`, `BusinessUnit`, `Owner`, `Environment`) are exported so each row can be attributed to an owner or chargeback unit:

- **`Tags`** lists every tag on the resource as `key=value; key=value`, sorted by key.
- **`Tag: <key>`** columns contain one tag each, ready for filtering/pivoting. By default a column is created for every tag key found in the scanned scope; use `-TagKeys CostCenter, BusinessUnit` to create columns only for the keys you care about (the `Tags` column still lists all tags).
- On **Model Deployments**, tags come from the parent **Cognitive Services / AI Foundry account**, so all deployments of the same account show the same tags. On **Hub-Project Workspaces**, they are the workspace's own tags.
- Only tags set directly on the resource are shown. Resource-group or subscription tags are **not** inherited into the report (unless your Azure Policy copies them onto the resource). A blank tag cell means the resource doesn't have that tag.
- Tag keys are matched case-insensitively (as Azure does), so `CostCenter` and `costcenter` share one column, using the casing first seen.
- Tag values are written as literal text: leading zeros are preserved (e.g. cost center `00123`), URLs are not turned into hyperlinks, and a value starting with `=` is prefixed with `'` so Excel doesn't evaluate it as a formula. Values longer than Excel's 32,767-character cell limit are truncated with a `...[TRUNCATED]` marker.

### Interpreting retirement flags

| `RetirementFlag` | Meaning in this report | Suggested next step |
|---|---|---|
| `RETIRED` | The catalog lifecycle is `Retired`, or the inference-retirement date has passed. | Investigate immediately, confirm the exact version's status, and check application dependencies and service behavior. |
| `RETIRING SOON` | The inference-retirement date is within `-RetirementWarningDays` of the run time; the default is 90 days. | Schedule migration and testing before the deadline, starting with high-usage dependencies. |
| `DEPRECATING` | No inference-retirement date is available, and the catalog lifecycle is `Deprecating`. | Confirm the published timeline and plan a replacement. |
| `LEGACY` | No inference-retirement date is available, and the catalog lifecycle is `Legacy`. | Review upgrade options and support expectations. |
| `PREVIEW` | No inference-retirement date is available, and the catalog lifecycle is `Preview`. | Review suitability and support requirements for the workload. |
| `OK` | An available inference-retirement date is beyond the warning window, or no other flag rule matched. | Continue monitoring. This is not a guarantee of long-term support or application health. |
| `UNKNOWN (not in catalog)` | No model match was found in the regional catalog, including when the catalog could not be retrieved. | Validate the exact model/version and catalog access manually. Do not treat this as safe or as proof of retirement. |
| `No deployments` | No deployments were returned for the account, either because none were listed or because listing failed. | Check console warnings and account access before treating the account as unused. |

**Flag precedence matters:** a catalog lifecycle of `Retired` takes priority. Otherwise, an available inference-retirement date determines `RETIRED`, `RETIRING SOON`, or `OK` before the `Deprecating`/`Legacy`/`Preview` lifecycle checks. For example, a `Legacy` model with a retirement date beyond the warning window appears as `OK`; inspect both lifecycle and flag columns.

Retirement text highlighting is separate from token-volume coloring. Always read the actual `RetirementFlag` value rather than relying on color alone.

### Reading weekly usage and token colors

The usage columns represent **prompt plus completion tokens**, not request counts, unique users, or monetary cost.

- **Wk1** is the most recent seven days before the run.
- **Wk2** is the preceding seven days, followed by **Wk3** and **Wk4**; **Wk4 is the oldest**.
- Headers include each window's dates. Boundaries are calculated in **UTC**, relative to the run time, rather than calendar weeks.
- `Total4WeekTokens` is the sum of the four weekly values, covering the last 28 days.
- The script buckets daily Azure Monitor aggregates by their timestamps. These are not exact request-level measurements at arbitrary week boundaries.

All four weekly columns and the total use the same fixed bands on **Model Deployments**:

| Color | Token count in that cell |
|---|---|
| Green | 0 through 50,000 |
| Yellow | More than 50,000 through 250,000 |
| Orange | More than 250,000 but less than 5,000,000 |
| Red | 5,000,000 or more |

These are absolute bands, not a gradient relative to other rows. A red token cell means **high consumption**, not that the model is retired. A green token cell means **low reported consumption**, not that the deployment is safe to ignore. A 4-week total can have a different color from every individual week because it covers a longer period. Blank and nonnumeric token cells receive no token-band color.

> **Zero does not necessarily mean unused.** Deployments receive zero values when usage is skipped with `-SkipTokenUsage`, when the metric is unsupported or unavailable, when retrieval fails, or when no matching deployment series is returned. Accounts with no listed deployments have blank usage cells instead. Use `-Verbose` to see metric-retrieval diagnostics, and validate telemetry before making decommissioning or migration decisions.

### Example: turning the report into an action list

The following rows are **illustrative only**, with shortened columns and fictional deployment names:

| DeploymentName | RetirementFlag | DaysUntilRetirement | Wk1 tokens | Total4WeekTokens | Interpretation |
|---|---|---|---|---|---|
| `support-chat` | `RETIRING SOON` | 21 | 1,500,000 | 6,200,000 | Active workload with a near-term deadline. The latest week is orange and the total is red; prioritize migration planning and testing. |
| `batch-extract` | `LEGACY` | blank | 40,000 | 160,000 | Lower reported usage, but the missing deadline is not evidence of indefinite support. Confirm the lifecycle timeline. |
| `test-assistant` | `RETIRED` | -3 | 0 | 0 | Investigate the retirement status now, but verify telemetry before assuming nothing depends on it. |

A practical review workflow:

1. **Check run coverage.** Review console warnings for skipped tenants/subscriptions, deployment-listing failures, or catalog-access problems. A successfully saved workbook can still be a partial inventory.
2. **Start with Summary.** Review status counts and the highest-impact list. That list includes only `RETIRED`, `RETIRING SOON`, `DEPRECATING`, and `LEGACY`, sorted by total tokens descending; it is not a deadline-first risk score. Zero-valued entries can still appear.
3. **Filter Model Deployments.** Investigate retired and soon-to-retire rows first. Sort `DaysUntilRetirement` ascending for urgency and compare `Total4WeekTokens` for traffic impact. Also review unknown statuses rather than relying solely on the ranked list.
4. **Read Note and confirm the exact version.** A fallback catalog match can supply lifecycle/deadline information for a different version. Blank retirement dates do not prove that no retirement is planned.
5. **Check the usage trend.** Compare Wk4 toward Wk1 to see whether traffic is rising, falling, or persisting. Cross-check zeros and unexpected changes with application telemetry.
6. **Assign follow-up work outside the report.** Identify application owners, confirm dependencies, select a supported replacement, and test it. The workbook does not identify application owners or recommend replacement models.
7. **Rerun after changes.** Compare dated reports to track deployment/version changes and remaining usage. Because the usage window rolls forward, totals from separate runs are not necessarily like-for-like periods.

The console finishes with row/workspace counts and the workbook path, warns if retired or soon-to-retire rows exist, and prints up to five entries from the impact ranking. The full ranked list (up to 25) is on **Summary**; all discovered deployment rows remain on **Model Deployments**.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-TenantId` | `string[]` | all accessible tenants | One or more Entra ID tenant IDs to scan. Default: every tenant returned by `Get-AzTenant` (i.e. every tenant your signed-in identity has at least Reader-level visibility into). |
| `-SubscriptionId` | `string[]` | all enabled subscriptions | One or more subscription IDs to restrict the scan to. Default: every enabled subscription in the scanned tenant(s). |
| `-OutputPath` | `string` | time-stamped `.xlsx` next to the script | Path to the Excel report that will be created. |
| `-RetirementWarningDays` | `int` | `90` | Number of days before a model's inference retirement date to flag it as `RETIRING SOON`. |
| `-SkipWorkspaces` | `switch` | off | Skip discovery of `Microsoft.MachineLearningServices/workspaces` (Hub/Project workspaces). Speeds up the scan if you only care about Azure OpenAI / Cognitive Services deployments. |
| `-SkipTokenUsage` | `switch` | off | Skip the per-deployment weekly token-consumption lookup. Speeds up the scan if you only need retirement status, not usage/impact data. |
| `-TagKeys` | `string[]` | all tag keys found | Tag keys (case-insensitive) to expand into dedicated `Tag: <key>` columns, e.g. `CostCenter, BusinessUnit, Owner`. The `Tags` column always lists all tags. |

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

# Only create dedicated columns for the chargeback tags (all tags still appear in the "Tags" column)
.\Get-AIFoundryInventory.ps1 -TagKeys CostCenter, BusinessUnit, Owner
```

## Output columns (Model Deployments sheet)

`Tenant`, `TenantId`, `SubscriptionName`, `SubscriptionId`, `ResourceGroup`, `AccountName`, `AccountKind`, `FoundryProject`, `Location`, `AccountCreatedOn`, `DeploymentName`, `ModelName`, `ModelVersion`, `ModelFormat`, `SkuName`, `Capacity`, `ProvisioningState`, `DeploymentCreatedOn`, `DeploymentModifiedOn`, `LifecycleStatus`, `InferenceRetirement`, `FineTuneRetirement`, `DaysUntilRetirement`, `RetirementFlag`, **`Wk1 Tokens (…)` … `Wk4 Tokens (…)`**, `Total4WeekTokens`, `Note`, `Tags`, **`Tag: <key>` …** (one per tag key).

`RetirementFlag` values: `OK`, `PREVIEW`, `LEGACY`, `DEPRECATING`, `RETIRING SOON`, `RETIRED`, `UNKNOWN (not in catalog)`, `No deployments`.

## Notes & limitations

- Auth reuses your existing `Connect-AzAccount` / `az login` session. If a tenant can't be queried with the current token (e.g. MFA-restricted guest access), the script attempts `Connect-AzAccount -Tenant <id>` once and otherwise skips that tenant with a warning — it does not fail the whole run.
- Token usage is read directly from the `Microsoft.Insights/metrics` REST API (not `Get-AzMetric`, which has a known bug for this metric/dimension combination in some `Az.Monitor` versions).
- Accounts/kinds that don't emit the `TotalTokens` metric (e.g. `ContentSafety`) simply report `0` for all weeks instead of erroring.
- The Azure Monitor dimension-split query returns up to 500 series per account (`&top=500`), which comfortably covers typical deployment counts per account.
- The script does not inspect your Azure RBAC role assignments up front; access gaps (missing Reader/Monitoring Reader) are instead discovered organically per tenant/subscription/account and reported as warnings with the specific failure, so the scan keeps going instead of aborting.

## Troubleshooting

### Fatal errors (stop the script)

These occur during the pre-flight checks, before any scanning begins. Each one prints an actionable message; the fixes below summarize them.

| Error | Likely cause | Fix |
|---|---|---|
| `Unsupported PowerShell version: ...` | Running on Windows PowerShell 5.0 or older. | Install PowerShell 7+ from https://aka.ms/powershell-release-current and re-run the script with `pwsh .\Get-AIFoundryInventory.ps1` (PowerShell 5.1 also works, but 7+ is recommended). |
| `The following required module(s) could not be installed/imported: ...` | `Az.Accounts`, `Az.CognitiveServices`, `Az.ResourceGraph`, `Az.Resources`, `Az.Monitor`, or `ImportExcel` failed to auto-install or import — often due to no internet access to the PowerShell Gallery, a corporate proxy, lack of permission to write to the CurrentUser module path, or a conflicting module version already loaded in the session. | Run the suggested command manually for each listed module, e.g. `Install-Module -Name Az.CognitiveServices -Scope CurrentUser -Force`, then start a fresh PowerShell session and re-run the script. If you're behind a proxy, configure `$env:HTTPS_PROXY` or use `Register-PSRepository`/`Set-PSResourceRepository` as required by your network policy. |
| `Azure sign-in failed: ...` | `Connect-AzAccount` was cancelled, timed out, or failed (e.g. MFA prompt not completed, no browser available in a headless/remote session). | Run `Connect-AzAccount` by itself first and confirm it completes successfully, then re-run the script. In a headless/remote session, use `Connect-AzAccount -UseDeviceAuthentication` and follow the device-code prompt. |
| `Unable to establish an authenticated Azure session.` | `Connect-AzAccount` returned without error, but no usable account/context resulted (rare; usually an interrupted or partial sign-in). | Run `Disconnect-AzAccount` (if applicable) followed by `Connect-AzAccount`, confirm `Get-AzContext` shows a signed-in account, then re-run the script. |
| `Could not enumerate Entra ID tenants for the signed-in account ...` | The signed-in identity has no network route to Azure AD, the account is disabled/locked, or it has zero role assignments anywhere in Azure. | Confirm you signed in with the intended account (`Get-AzContext`), check connectivity to `login.microsoftonline.com`/`management.azure.com`, and confirm the account has at least Reader on some subscription. |
| `The signed-in account (...) has no accessible Entra ID tenants.` | The account authenticated successfully but isn't associated with any Entra ID tenant the Az SDK can see (e.g. a brand-new guest invite not yet accepted, or an account with no Azure role assignments at all). | Verify the account has at least Reader access to one subscription, or sign in with a different account via `Connect-AzAccount -Tenant <tenantId>`. |

### Warnings (scan continues, data may be incomplete)

These are printed to the console during the scan but do not stop the run. The report is still produced, but the affected tenant/subscription/account may be missing or incomplete — always check the console output for warnings before trusting a "clean" report.

| Warning | Likely cause | Fix |
|---|---|---|
| `Could not retrieve model catalog for location '<region>': ...` | The regional model catalog lookup (`Get-AzCognitiveServicesModel`) failed — often a transient Azure API error or a region with no catalog data. | Re-run the script; if it persists for a specific region, confirm that region is valid and that your account has Cognitive Services Reader there. Affected deployments will show `UNKNOWN (not in catalog)`. |
| `Could not list subscriptions for tenant <tenantId> with the current session. Trying interactive login for this tenant...` followed by `Skipping tenant <tenantId> - unable to authenticate or list subscriptions: ...` | Your current token doesn't have visibility into that tenant (common for guest/B2B accounts, especially with Conditional Access or MFA requirements specific to that tenant), and the automatic `Connect-AzAccount -Tenant <id>` retry also failed or was cancelled. | Run `Connect-AzAccount -Tenant <tenantId>` manually to confirm you can sign in to that tenant at all, and check with the tenant admin that your account has at least Reader access to a subscription there. If you don't need that tenant, scope the scan with `-TenantId` to exclude it. |
| `Could not set context in tenant <tenantId> - ...` or `Could not switch context to subscription <subId>: ...` | `Set-AzContext` failed — the subscription may be disabled, moved, or the account lost access to it mid-run. | Confirm the subscription is still `Enabled` (`Get-AzSubscription`) and that your account retains Reader access; re-run the script. |
| `Resource Graph query failed in tenant <tenantId> - ...` | Azure Resource Graph query failed for every subscription in that tenant, usually because the account has no Reader (or higher) role on any of them. | Grant the signed-in account Reader on at least one subscription in that tenant, or exclude it with `-TenantId`/`-SubscriptionId` if it's not in scope. |
| `Could not list deployments for account <accountName>: ...` | The account lacks `Cognitive Services Reader` (or higher) on that specific Cognitive Services/Azure OpenAI account, even though the parent subscription is readable. | Grant `Cognitive Services Reader` (or higher) on the account, or expect that account's row(s) to show `No deployments` until access is fixed. |
| Token usage columns are unexpectedly `0` for an account/deployment you know is active (no warning text; check with `-Verbose`) | The signed-in account lacks `Monitoring Reader` (or higher) on that Cognitive Services account, the account's kind doesn't emit the `TotalTokens` metric (e.g. `ContentSafety`), or the deployment had no traffic in the last 28 days. | Grant `Monitoring Reader` on the account and re-run with `-Verbose` to see the specific `Could not retrieve token usage for <resourceId> : ...` message (shown only at Verbose level), which distinguishes a permissions/API error from genuinely zero usage. |

### Other common issues

| Symptom | Likely cause | Fix |
|---|---|---|
| The script runs but the `.xlsx` is empty or shows "No Cognitive Services / AI Foundry accounts or deployments were found in the scanned scope." | No `Microsoft.CognitiveServices/accounts` were discovered in the scanned tenants/subscriptions, or the account has Reader access to none of them. | Confirm the expected resources exist in the scope you're targeting (check the Azure Portal), and scope the run explicitly with `-TenantId`/`-SubscriptionId` to rule out an overly broad/narrow default scan. |
| `Export-Excel`/`Close-ExcelPackage` fails with a file-in-use or access-denied error. | The target `.xlsx` (default or `-OutputPath`) is already open in Excel, or the folder is read-only/OneDrive-locked. | Close the file in Excel before re-running, or pass a different `-OutputPath`. |
| The run takes a very long time. | Token-usage lookups and model-catalog calls add an Azure API round-trip per account/region, which adds up across large tenants with many accounts. | Use `-SkipTokenUsage` and/or `-SkipWorkspaces` for a faster pass when you don't need usage/workspace data, or scope the run with `-TenantId`/`-SubscriptionId`. |
| PowerShell prompts to trust the PSGallery repository or shows SSL/TLS errors during module install. | PSGallery isn't yet configured as a trusted repository, or the machine's TLS settings are outdated. | Run `Set-PSRepository -Name PSGallery -InstallationPolicy Trusted` once, and ensure `[Net.ServicePointManager]::SecurityProtocol` includes `Tls12` (default on PowerShell 7+). |

## License

[MIT](./LICENSE)
