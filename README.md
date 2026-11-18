# Export Conditional Access Policies PowerShell Scripts

Export and analyze Microsoft Entra ID (Azure AD) Conditional Access (CA) policies for documentation, security review, and remediation planning.

Two scripts are provided:

| Script | Purpose |
|--------|---------|
| `Export-CAPolicy.ps1` | Lightweight export – single HTML table of policies. |
| `Export-CAPolicyWithRecs.ps1` | Full featured export – enrichment, recommendations, duplicate detection, multiple output formats, interactive HTML UI. |

---
## Key Features (Full Export Script)

* Automatic Microsoft Graph connection (Policy.Read.All, Directory.Read.All, RoleManagement.Read.All), with post-connect verification that all required scopes were granted.
* Policy enrichment: resolves Users, Groups, Roles (active + template/definition), Service Principals, Named Locations, Terms of Use.
* Duplicate policy detection (normalized JSON → SHA256 hash) with cross‑reference listing.
* Offline re‑hydration mode (`-RawInputFile <*_raw.json>`): regenerate HTML / CSV / recommendations from a prior raw snapshot without calling Microsoft Graph.

* Interactive HTML report with tabbed layout:
	* Summary (deployment metric cards, categorized inventory table with aligned values, percentage bars and explicit denominators, assessment snapshot)
	* Policy Details (combined search/state filters, visible row count, contained scrolling with sticky names and two-row headers)
	* Recommendations (32 assessments, four purpose-based views, a separate assessment dropdown, prerequisites, background, Microsoft links, and policy evidence)
	* Named Locations (all tenant definitions, IPv4/IPv6 ranges, trusted flags, countries/regions, unknown-area handling, lookup method, IDs and timestamps)
* Consistent spacing, readable typography, responsive cards/navigation, and automatic light/dark styling across all tabs (`?scoutTheme=light` or `?scoutTheme=dark` overrides the system theme)
* Overlap legend chip & row highlighting: quickly filter policies where include/exclude sets intersect (potentially nullifying scope)
* Authentication-strength assessment uses known built-in IDs or verified method combinations, not display-name guesses
* Enhanced admin role detection: resolves active roles, template IDs, and role definition IDs (GUID → name) before MFA checks
* Legend / utility bar (status badges, recommendation icons, boolean toggle, ID column toggle)
* Layout toggle (dynamic natural column widths vs fixed legacy widths)
* Boolean display toggle (icons ✔/— or text True/False)
* Description clamp (expand/collapse with accessible aria attributes)
* Per‑policy original vs mutated JSON toggle + Copy JSON button (clipboard)
* Preference persistence (remembers active tab, boolean mode, ID column visibility, recommendation view and assessment filter)
* Accessibility: keyboard‑navigable tabs, aria-live announcements (copy feedback), sticky headers, focusable controls
* CSV (long), CSV Pivot (wide) and JSON export options
* Duplicate detection, role/name resolution, risk & device metrics in Summary
* Optional fast mode: `-NoRecommendations` skips recommendation analysis (faster, omits Recommendations tab & related JavaScript/state)
* Inline recommendation catalog (no external data file dependency) with escaped policy evidence

Screenshots:
![Recommendations Tab](./images/CaExport-rec.png)
![Results Table](./images/CaExport-result.png)
![Duplicate Highlighting](./images/CAExport-Select.png)
![Summary Table](./images/CaExport-summary.png)
---
## Installation / Prereqs

PowerShell 7.5 + is recommended. Ensure Microsoft Graph modules are installed. You can install all of the modules using:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

To target only the required Microsoft Graph modules, you can specify them individually:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.Governance -Scope CurrentUser
Install-Module Microsoft.Graph.DeviceManagement.Enrollment -Scope CurrentUser
```

When the script runs it will request admin consent to allow it to use the Microsoft Graph command line tools, if this not already granted. If you do not have permission to grant admin consent, you will need to work with your Entra ID administrator to obtain the necessary permissions.

After sign-in or reconnection, the script reads the new Graph context and stops if the authenticated account or any required scope is missing. This verifies granted scopes, not the account's Entra role or access to every individual resource; API failures still surface separately.

![Requesting Admin Consent for Graph Command Line Tools ](./images/MicrosoftGraphConsent.png)

---
## Quick Start

```powershell
# Basic full export (now defaults to HTML + JSON + CSV)
./Export-CAPolicyWithRecs.ps1

# All formats
./Export-CAPolicyWithRecs.ps1 -Html -Json -Csv -CsvPivot

# Identify designated emergency accounts for exclusion review (object IDs, not account names)
./Export-CAPolicyWithRecs.ps1 -Html -EmergencyAccessAccountId `
  '11111111-2222-3333-4444-555555555555','66666666-7777-8888-9999-000000000000'

# Single policy by Id (also export JSON)
./Export-CAPolicyWithRecs.ps1 -PolicyID 11111111-2222-3333-4444-555555555555 -Json

# Custom CSV columns
./Export-CAPolicyWithRecs.ps1 -Csv -CsvColumns Name,Status,'Require MFA','Block'

# Fast export (skip recommendations & omit tab)
./Export-CAPolicyWithRecs.ps1 -NoRecommendations

# Offline (re‑hydrate from prior raw snapshot produced by an earlier run)
#   - Skips Graph connection & enrichment queries
#   - Still evaluates recommendations on provided raw policy array
./Export-CAPolicyWithRecs.ps1 -RawInputFile .\CAExportRecs_Tenant_20250101_120000_raw.json -Html -Json
```

**Output files are timestamped:**
- Full script: `CAExportRecs_<TenantName>_YYYYMMDD_HHMMSS.*`
- Lightweight script: `CAExport_<TenantName>_YYYYMMDD_HHMMSS.*`

---
## Parameters (Full Script)

| Parameter | Type | Description |
|-----------|------|-------------|
| `PolicyID` | String (GUID) | Limit export to a single CA policy. |
| `Html` | Switch | Emit interactive HTML (default if no other export is chosen). |
| `NoBrowser` | Switch | Generate HTML without automatically launching browser. |
| `OutputPath` | String | Custom output directory path (defaults to current directory). |
| `Json` | Switch | Emit enriched JSON (includes duplicate markers). |
| `Csv` | Switch | Emit flattened CSV of policies. |
| `CsvPivot` | Switch | Emit pivot‑friendly wide CSV. |
| `CsvColumns` | String[] | Custom column subset/order for `-Csv`. |
| `NoRecommendations` | Switch | Skip recommendation analysis & omit Recommendations tab (improves runtime in large tenants). |
| `RawInputFile` | String (path) | Use a previously exported `*_raw.json` file (offline mode). Skips Graph auth + enrichment; recommendations derived from provided data. |
| `EmergencyAccessAccountId` | String[] (GUIDs) | Explicitly designated emergency-account object IDs for exclusion review. No account names or membership are guessed. |

If no export switch is specified, HTML, JSON, and CSV are produced automatically (pivot remains opt-in).

## Named Locations

Connected HTML exports retrieve the complete tenant inventory with Graph pagination, including locations not referenced by the exported policies. The same inventory resolves policy location IDs without extra per-location queries. JSON/CSV-only exports retain their existing referenced-location enrichment.

The Named Locations tab distinguishes an empty tenant inventory from retrieval failure. Offline policy snapshots do not contain location definitions, so the tab explicitly reports that the inventory is unavailable instead of claiming no locations exist. Location definitions do not change recommendation assessments or prove effective network access.

Background and permissions: [Microsoft Entra network assignments](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-assignment-network) and [Microsoft Graph list namedLocations](https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-namedlocations?view=graph-rest-1.0).

---
## Recommendation Checks

Each recommendation includes background, verification/remediation steps, requirements, Microsoft documentation links, and matching policy evidence with deployment state and scope limitations.

| Checks | Assessment area |
|--------|-----------------|
| CA-00 to CA-05 | Legacy authentication; all-user MFA; mobile compliance/app protection; Windows/macOS compliance; administrator MFA and phishing-resistant MFA |
| CA-06 to CA-08 | Exact include/exclude overlaps, empty identity assignments, and direct-user assignment review |
| CA-09 | Medium/high sign-in risk, MFA/authentication strength, and every-time reauthentication |
| CA-10 to CA-12 | Device-code flow, MFA for device registration/join, and unsupported-platform blocking |
| CA-13 | MFA/authentication strength for security-information registration (distinct from device registration) |
| CA-14 | High-user-risk remediation; legacy password-change policies flagged for separate review |
| CA-15 to CA-18 | Managed administrator devices; guest/external-user MFA; Microsoft admin portal and Azure management MFA |
| CA-19 | Migration from the legacy approved-client-app grant to app protection |
| CA-20 to CA-22 | Unmanaged-device restrictions, browser/session controls, and designated emergency-access exclusions |
| CA-23 to CA-27 | Token protection, workload identities, Purview insider risk, autonomous/high-risk agents, and delegated agent access |
| CA-28 | Optional Microsoft 365 unmanaged-device blocking template: compliance grants or recognized compliant-device block filters |
| CA-29 | Optional device restrictions for explicitly selected resource applications, with actual application IDs shown in evidence |
| CA-30 | Optional SharePoint/OneDrive limited browser access template; application-side restrictions and companion client policies require verification |
| CA-31 | Optional Microsoft 365 Android/iOS app-protection template, distinct from full device enrollment/blocking |

**Assessment states:** `Enforced`, `Report-only`, `Disabled`, `Missing`, `Needs review`, and `No findings` (configuration hygiene only). Only an enabled policy meeting the checked grant and direct-scope criteria receives `Enforced`. An MFA **OR** compliant-device policy does not prove mandatory MFA. Broad all-resource MFA can satisfy guest, administrator, portal, and Azure management checks without duplicate policies. Empty Graph SDK containers are ignored; explicit all-platform/all-location assignments without exclusions are not restrictive conditions.

**Attention** includes protection/hygiene findings and gaps for explicitly designated emergency accounts, not absent optional scenarios. Optional templates and optional/operational reviews have their own filters and are counted separately from baseline states in Summary. Cards show observed enabled/report-only/disabled policy counts even when exclusions or scope still need review; `Needs review` does not mean a policy is disabled or ineffective.

The Recommendations toolbar has four views: **Attention** (the default for new preferences), **All recommendations**, **Optional templates**, and **Operational reviews**. A separate **Assessment** dropdown optionally narrows the selected view; deployment-state filters remain in Policy Details. The visible count and empty-state message explain combined filters. Saved view/assessment choices persist, and previous status-button preferences migrate to the dropdown.

This is a **configuration assessment, not an effective-access simulation or compliance certification**. Group/PIM membership, cross-tenant trust, combined coverage across multiple scoped policies, licensing, supporting application configuration, and actual sign-in outcomes are not verified. Exclusions other than supplied emergency-account IDs, restrictive/unknown conditions, and incomplete strength metadata yield `Needs review`. A `-PolicyID` export marks all recommendations `Needs review` because it is not a tenant-wide snapshot. Offline results describe only the supplied snapshot.

Optional/business-specific controls (device-code dependencies, unmanaged-device restrictions, sessions, token protection, workloads, insider risk, and agents) stay `Needs review`, with observed evidence where available, rather than becoming universal failures. Agent policy coverage cannot be certified from the v1.0 export; delegated-agent use cannot be inferred from ordinary user MFA policies.

### Optional application templates

Use the **Optional templates** filter to inspect intended users, resources, conditions, grants/session controls, rollout guidance, requirements, and Microsoft references. Each template separately reports `Matching policy enabled`, `Matching policy report-only`, `Matching policy disabled`, `Partial match`, or `Not observed`. These are policy-pattern measurements, not applicability decisions or effective-access certifications. No policies are created or changed.

The Microsoft 365 template recognizes the `Office365` resource grouping or a stronger all-resource policy; SharePoint-only protection is not suite-wide protection. The selected-app template inventories explicit assignments rather than guessing which business apps should be covered. Resource exclusions, dynamic application filters, scoped users/conditions, and grant bypasses remain visible as partial matches.

Device restrictions recognize mandatory compliance (optionally hybrid join), or Block with a simple `exclude: device.isCompliant -eq True` / `include: device.isCompliant -ne True` filter. Complex filters require review. `isManaged` or device registration alone is not compliance; `isCompliant -eq False` does not cover null attributes on unregistered devices. A blanket application block without a device condition is not identified as a device-restriction template.

SharePoint limited-access measurements detect the Conditional Access session setting only: verify tenant/site settings, download/print/sync behavior, and companion non-browser/legacy-client restrictions separately. Mobile app protection requires assigned Intune policies and supported apps/brokers; it can allow approved personal devices rather than denying every unmanaged device.

### Prerequisites and reference guidance

| Area | Requirements and background |
|------|-----------------------------|
| General Conditional Access | [Entra ID P1 or an eligible suite, appropriate roles, and staged deployment](https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview). Start changes in [report-only mode](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-report-only). Export permissions are not deployment permissions. |
| MFA and registration | Register supported methods and plan bootstrap/recovery. Review [authentication strengths](https://learn.microsoft.com/en-us/entra/identity/authentication/concept-authentication-strengths), [security-information registration](https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-security-info-registration), and [Temporary Access Pass](https://learn.microsoft.com/en-us/entra/identity/authentication/howto-authentication-temporary-access-pass). |
| Risk | Entra ID P2 and ID Protection signals. See [sign-in risk](https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-sign-in) and [user-risk remediation](https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-user). Modern `riskRemediation` requires authentication strength with AND and every-time reauthentication. Legacy password-change remediation also needs applicable SSPR/password-writeback setup. SDK/API versions can omit newer controls; unknown values need review. |
| Devices and mobile apps | Appropriate Intune or supported partner licensing plus deployed [compliance policies](https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started) or [app-protection policies](https://learn.microsoft.com/en-us/intune/intune-service/apps/app-protection-policy). Hybrid join does not apply to macOS. [Approved-client-app migration](https://learn.microsoft.com/en-us/entra/identity/conditional-access/migrate-approved-client-app): the control became read-only June 30, 2026; existing enabled policies still enforce. |
| Application-scoped device restrictions | Choose [resource applications or the Microsoft 365 suite](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-cloud-apps#office-365), verify [service dependencies](https://learn.microsoft.com/en-us/entra/identity/conditional-access/service-dependencies), and understand [device-filter null/unregistered behavior](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-condition-filters-for-devices). Pilot enrollment, guests and emergency recovery before blocking. |
| SharePoint limited browser access | [Configure SharePoint tenant/site unmanaged-device controls](https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices), with appropriate SharePoint/Entra licensing, SharePoint Administrator, supported clients and companion non-browser restrictions. A session-control flag does not verify application-side restrictions. |
| Guests | Review applicable External ID licensing/billing and [cross-tenant MFA/device trust](https://learn.microsoft.com/en-us/entra/external-id/cross-tenant-access-overview). |
| Emergency access | Supply designated IDs, then independently verify at least two accounts, resilient authentication, monitoring, and recovery tests. Follow [Microsoft emergency-access guidance](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access). Missing direct exclusions require targeting/membership review, not automatic changes. |
| Token protection | P1 and supported devices, clients and resources; review the [support matrix and platform deployment guides](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-token-protection), including Apple SSO/MDM and web-preview limitations. |
| Workload identities | [Workload ID Premium](https://learn.microsoft.com/en-us/entra/identity/conditional-access/workload-identity) to create/modify policies. Supported tenant-owned service principals must be targeted directly; group membership, managed identities and third-party/multitenant apps do not establish coverage. |
| Insider risk | Appropriate Purview licensing and enabled [Adaptive Protection](https://learn.microsoft.com/en-us/purview/insider-risk-management-adaptive-protection) before using [insider-risk conditions](https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-insider-block). |
| Agents | [Agent 365 with the documented Entra license combination, supported identities/flows, and feature availability](https://learn.microsoft.com/en-us/entra/identity/conditional-access/agent-id); P2 for agent risk. Distinguish autonomous agent identities, delegated users, and agent user accounts. Confirm current preview limitations and sign-in evidence. |

Microsoft guidance and prerequisite links were reviewed November 18, 2026. Licensing, preview availability, and supported-client matrices can change; consult the linked pages before deployment. This script remains read-only against the tenant.

---
## Duplicate Detection Logic

Policies are normalized (sorted properties, non‑semantic fields removed) → serialized to JSON → hashed (SHA256). Matching hashes mark `IsDuplicate` and list `DuplicateMatches`. This protects against silent drift or redundant policies.

---
## Output Artifacts

| File | Contents |
|------|----------|
| `*.html` | Interactive report (Summary, Details, Recommendations – last omitted when `-NoRecommendations` used). |
| `*.json` | Enriched policy objects + duplicate metadata. |
| `*.csv` | Flat list (one row per policy). |
| `*-pivot.csv` | Wide format for Excel / BI pivoting. |
| `*_raw.json` | Raw policy array snapshot (input for `-RawInputFile` offline regeneration). |

---
## Troubleshooting

| Issue | Resolution |
|-------|-----------|
| Empty policy list | Confirm permissions: Policy.Read.All + Directory.Read.All. Run `Disconnect-MgGraph` then rerun. |
| Unresolved role IDs | Role not active and no definition found: activate role or ensure sufficient RoleManagement.Read.All consent. |
| Sanitized recommendation policy markup shows literal `&lt;` tags | You may be viewing an old report generated before sanitizer simplification; regenerate with latest script. |
| Missing images in HTML | Ensure PNGs are in same directory as report or adjust paths. |
| Duplicate false positives | Open JSON, verify normalized fields; adjust logic if environment-specific noise appears. |
| Offline mode shows fewer enriched names | `-RawInputFile` skips live resolution; rerun without offline mode for fresh enrichment. |

---
## Contributing

PRs welcome. Please:
1. Run PSScriptAnalyzer locally
2. Keep alignment style (hash property columns) intact
3. Add comment‑based help for any new function/parameter
4. Update this README when adding features

---
## Version History

| Version | Highlights |
|---------|-----------|
| 3.5 | Fixed false restrictions from empty Graph SDK models and separated optional applicability reviews from Attention; four application-scoped templates with match/deployment measurements and resource/filter evidence; refreshed all tabs with responsive light/dark layouts, contained sticky scrolling and combined filters. |
| 3.4 | 28 Microsoft-backed recommendation assessments with state/scope/grant-aware evaluation, prerequisites/background links, escaped evidence, emergency-account IDs, and optional advanced-scenario reviews. |
| 3.3 | Offline mode (`-RawInputFile`), recommendations array inlined (removed external data file), improved Recommendations tab styling, simplified HTML sanitizer (prevents escaped structural tags), removed obsolete quiet behavior, minor accessibility & layout refinements. |
| 3.2 | Overlap filter chip + row highlighting; Phish‑resistant MFA evaluation using AllowedCombinations; improved admin role GUID resolution; authentication strength caching; CA‑05 logic hardened; null overlap guard fix. |
| 3.1| Pivot CSV dataset restored (wide format); improved role resolution (template & definition IDs), GUID parameter validation, duplicate detection hashing improvements, accessibility & JSON copy utilities. |

| 3.0 | Recommendations refactor, duplicate detection, summary tab, accessibility & help improvements. |
| 2.x | Styling overhaul, column selection improvements. |
| 1.x | Initial export functionality. |

---
## License

It's cool to share stuff to make other people's lives easier, so let's keep doing that.

Shield: [![CC BY-NC-SA 4.0][cc-by-nc-sa-shield]][cc-by-nc-sa]

This work is licensed under a [Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License][cc-by-nc-sa].

[![CC BY-NC-SA 4.0][cc-by-nc-sa-image]][cc-by-nc-sa]

[cc-by-nc-sa]: http://creativecommons.org/licenses/by-nc-sa/4.0/
[cc-by-nc-sa-image]: https://licensebuttons.net/l/by-nc-sa/4.0/88x31.png
[cc-by-nc-sa-shield]: https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg
