<#
.SYNOPSIS
    Export existing Conditional Access Policies along with recommendations.

.DESCRIPTION
    Supported version 4.0 exporter for Microsoft Entra Conditional Access policies.
    Generates HTML, enriched JSON, CSV and pivot CSV, with naming hints and 32 configuration assessments.
    Connected runs save a portable raw snapshot; RawInputFile regenerates reports without Graph.
    HTML supports ID-based assignment highlighting and exact-field comparisons, not effective-access simulation.
    Does not create, rename or modify tenant policies. Replaces the deprecated Export-CaPolicy.ps1.

.PARAMETER PolicyID
  (Optional) A specific Conditional Access policy Id (GUID). When supplied the export/report is limited
  to this single policy. When omitted all policies are processed.

.PARAMETER TenantID
  Optional tenant GUID for connected exports. Reconnects when the existing Graph context targets
  a different tenant and verifies the tenant after sign-in. Cannot be combined with RawInputFile.

.PARAMETER Html
  Switch. Generate the interactive HTML report. With no export switches, HTML, JSON and CSV are produced.

.PARAMETER NoBrowser
  Switch. Save the HTML report without opening the default browser. Has no effect when HTML is not selected.

.PARAMETER OutputPath
  Existing directory for output files. Defaults to the current working directory, not the script directory.

.PARAMETER Json
  Switch. Generate a JSON file containing the enriched policy objects.

.PARAMETER RawInputFile
  Path to a portable *_raw.json snapshot from a connected run, or a legacy raw Graph policy array.
  Portable snapshots restore named locations, resolved names, authentication strengths, and capture metadata
  without Microsoft Graph modules, a connection, or network access.

.PARAMETER Csv
  Switch. Generate a flattened CSV export of the policy objects suitable for spreadsheet review.

.PARAMETER CsvPivot
  Switch. Generate a pivot‑friendly CSV (wide format) for ad‑hoc aggregation in Excel / BI tools.

.PARAMETER NoRecommendations
  Switch. Skip the 32 assessments and omit the Recommendations tab. Naming hints, assignment comparison,
  enrichment, other report tabs and connected-run snapshot collection remain available.

.PARAMETER CsvColumns
  Optional string array specifying a custom subset / order of columns for the Csv export. When omitted
  the default full set is used.

.PARAMETER EmergencyAccessAccountId
  Object IDs of designated emergency-access accounts. Used to review exclusions without guessing account names.
  Group membership, account readiness, licensing, and sign-in outcomes are not inferred from policy JSON.

.OUTPUTS
  Files written to the working directory unless OutputPath is supplied. Filenames are timestamped
  and prefixed with CAExportRecs_<TenantName>_. Connected runs also save a portable *_raw.json snapshot.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1
  Exports all Conditional Access policies to HTML, JSON and CSV, plus a portable raw snapshot.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -PolicyID 11111111-2222-3333-4444-555555555555 -Json
  Exports only the specified policy and produces a JSON file with enriched data.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -Csv -CsvColumns Name,Status,'Require MFA','Block'
  Produces a CSV limited to the selected columns in the specified order.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -Html -Json -Csv -CsvPivot
  Produces all supported output formats in a single invocation.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -NoRecommendations -Csv
  Exports policies to CSV plus a connected-run raw snapshot, while skipping recommendation analysis.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -PolicyID (Get-Clipboard) -Html
  Uses a policy Id copied to the clipboard and generates only the HTML report.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -TenantID 11111111-2222-3333-4444-555555555555 -Html -NoBrowser
  Selects and verifies the requested tenant, then saves HTML and a portable snapshot without opening a browser.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -RawInputFile .\snapshot_raw.json -Html -NoBrowser
  Regenerates HTML from a portable snapshot without Graph modules or sign-in. Do not combine with TenantID.

.NOTES
  Author:  Douglas Baker @dougsbaker
  Version: 4.0
  Requires PowerShell 7.5 or later (pwsh). Windows PowerShell 5.1 (powershell.exe) is not supported.
  Run: pwsh -NoProfile -File .\Export-CAPolicyWithRecs.ps1

############################################################################
This sample script is not supported under any standard support program or service.
This sample script is provided AS IS without warranty of any kind.
This work is licensed under a Creative Commons Attribution 4.0 International License
https://creativecommons.org/licenses/by-nc-sa/4.0/
############################################################################

#>

#Requires -Version 7.5

[CmdletBinding()]
# Suppress PSAvoidLongLines for unavoidable embedded HTML/URLs
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidLongLines', '')]
param (
  [Parameter()]
  [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
  [String]$PolicyID,
  [Parameter(HelpMessage = 'Path to a portable *_raw.json snapshot or legacy raw policy array (skips Graph calls).')]
  [ValidateScript({ Test-Path $_ -PathType Leaf })]
  [string]$RawInputFile,
  [switch]$Html,
  [switch]$NoBrowser,
  [Parameter()][ValidateScript({ Test-Path $_ -PathType Container })][string]$OutputPath = '.',
  [switch]$Json,
  [switch]$Csv,
  [switch]$CsvPivot,
  [switch]$NoRecommendations,
  [string[]]$CsvColumns,
  [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
  [string[]]$EmergencyAccessAccountId,
  [Parameter()]
  [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
  [string]$TenantID
)

function Write-Info {
  <#
.SYNOPSIS
  Write an informational message to the information stream.
.DESCRIPTION
  Wrapper around Write-Information so that callers can use -InformationAction / -InformationPreference
  and optionally suppress or capture messages. Replaces earlier Write-Host usage for lint compliance.
.PARAMETER Message
  The text to emit.
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Message)
  Write-Information -MessageData "[INFO] $Message" -InformationAction Continue
}


function Write-Warn {
  <#
.SYNOPSIS
  Write a warning message.
.DESCRIPTION
  Thin wrapper kept for symmetry with Write-Info.
.EXAMPLE
  Write-Warn 'Policy collection returned zero results.'
  Emits a formatted warning to the warning stream.
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Message)
  Write-Warning $Message
}


function Write-Err {
  <#
.SYNOPSIS
  Write an error message and optionally terminate the script.
.DESCRIPTION
  Wrapper around Write-Error for consistent error handling throughout the script.
.PARAMETER Message
  The error message to display.
.EXAMPLE
  Write-Err 'Critical failure occurred. Exiting.'
  Emits a formatted error message to the error stream.
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Message)
  Write-Error -Message $Message -ErrorAction Continue
}

function Format-PolicyStatus {
  <#
  .SYNOPSIS
    Formats policy status for better readability in exports
  .PARAMETER Status
    The raw policy status from Microsoft Graph
  .OUTPUTS
    String - Formatted status for display
  #>
  param([string]$Status)
  switch ($Status) {
    'enabledForReportingButNotEnforced' {
      return 'report only' 
    }
    default {
      return $Status 
    }
  }
}

function Get-RecommendedPolicyName {
  <#
  .SYNOPSIS
    Suggest a descriptive policy name from observed configuration.
  .DESCRIPTION
    Uses CA<random numeric prefix>-<scope>-<conditions>-<controls>.
    This heuristic display suggestion is not a stable policy identifier or an effective-coverage assessment.
  .PARAMETER Policy
    The Conditional Access policy object from Microsoft Graph
  .OUTPUTS
    String - Suggested policy name; no tenant policy is renamed.
  .EXAMPLE
    Get-RecommendedPolicyName -Policy $caPolicy
    Returns: "CA001-AllUsers-HighRisk-RequireMFA"
  #>
  param([Parameter(Mandatory)]$Policy)

  # Prefix - Standard CA prefix
  $prefix = 'CA'

  # Scope - Determine target user scope
  $scope = 'Unknown'
  $users = Get-CaProperty $Policy 'conditions.users'
  $includeUsers = @(Get-CaProperty $users 'includeUsers')
  $includeGroups = @(Get-CaProperty $users 'includeGroups')
  $includeRoles = @(Get-CaProperty $users 'includeRoles')
  if ($includeUsers -contains 'All') {
    if (@(Get-CaProperty $users 'excludeUsers').Count -or
        @(Get-CaProperty $users 'excludeGroups').Count -or
        @(Get-CaProperty $users 'excludeRoles').Count -or
        (Test-CaConfiguredValue (Get-CaProperty $users 'excludeGuestsOrExternalUsers'))) {
      $scope = 'AllUsers-Exceptions'
    }
    else {
      $scope = 'AllUsers'
    }
  }
  elseif ($includeUsers -contains 'GuestsOrExternalUsers') {
    $scope = 'Guests'
  }
  elseif ($includeRoles.Count -gt 0) {
    # Check for privileged admin roles
    $adminRoles = @('Global Administrator', 'Privileged Role Administrator', 'Security Administrator', 'Conditional Access Administrator')
    $hasAdminRole = $false
    foreach ($roleId in $includeRoles) {
      if ($RoleMap.ContainsKey($roleId) -and $RoleMap[$roleId] -in $adminRoles) {
        $hasAdminRole = $true
        break
      }
    }
    $scope = if ($hasAdminRole) {
      'PrivAdmins' 
    }
    else {
      'Roles' 
    }
  }
  elseif ($includeGroups.Count -gt 0) {
    $scope = 'Groups'
  }
  elseif ($includeUsers.Count -gt 0) {
    $scope = 'Users'
  }

  # Condition - Determine primary triggering condition
  $condition = @()
  $conditions = $Policy.conditions

  # Risk-based conditions (highest priority)
  if ($conditions.userRiskLevels -and $conditions.userRiskLevels.Count -gt 0) {
    if ($conditions.userRiskLevels -contains 'high') {
      $condition += 'HighUserRisk'
    }
    else {
      $condition += 'UserRisk'
    }
  }
  if ($conditions.signInRiskLevels -and $conditions.signInRiskLevels.Count -gt 0) {
    if ($conditions.signInRiskLevels -contains 'high') {
      $condition += 'HighSignInRisk'
    }
    else {
      $condition += 'SignInRisk'
    }
  }

  # Location-based conditions
  if ($conditions.locations.excludeLocations -and $conditions.locations.excludeLocations.Count -gt 0) {
    $condition += 'UntrustedLocation'
  }
  if ($conditions.locations.includeLocations -and $conditions.locations.includeLocations.Count -gt 0) {
    $condition += 'SpecificLocation'
  }

  # Client app conditions
  if ($conditions.clientAppTypes -contains 'exchangeActiveSync' -or
    $conditions.clientAppTypes -contains 'other') {
    $condition += 'LegacyAuth'
  }

  # Authentication flows conditions
  if ($conditions.authenticationFlows -and $conditions.authenticationFlows.transferMethods -and $conditions.authenticationFlows.transferMethods.Count -gt 0) {
    $condition += 'AuthFlows'
  }

  # Platform conditions
  if ($conditions.platforms.includePlatforms -and $conditions.platforms.includePlatforms.Count -gt 0) {
    $platforms = $conditions.platforms.includePlatforms
    if ($platforms -contains 'android' -and $platforms -contains 'iOS') {
      $condition += 'MobileDevices'
    }
    elseif ($platforms.Count -eq 1) {
      $condition += $platforms[0].Substring(0, 1).ToUpper() + $platforms[0].Substring(1)
    }
    else {
      $condition += 'Platforms'
    }
  }

  # Device conditions
  if ($conditions.devices.includeDevices -and $conditions.devices.includeDevices.Count -gt 0) {
    if ($conditions.devices.includeDevices -contains 'All') {
      $condition += 'AllDevices'
    }
    else {
      $condition += 'Devices'
    }
  }

  # Device filter conditions
  if ($conditions.devices.deviceFilter -and $conditions.devices.deviceFilter.rule) {
    $condition += 'DeviceFilter'
  }

  # Application conditions
  if ($conditions.applications.includeApplications -and $conditions.applications.includeApplications.Count -gt 0) {
    $apps = $conditions.applications.includeApplications
    if ($apps -contains 'All') {
      $condition += 'AllApps'
    }
    elseif ($apps -contains 'Office365') {
      $condition += 'Office365'
    }
    else {
      $condition += 'CloudApps'
    }
  }

  # Default condition if none detected
  if ($condition.Count -eq 0) {
    $condition += 'General'
  }

  # Control - Determine primary enforcement action
  $control = 'Unknown'
  $grantControls = Get-CaProperty $Policy 'grantControls'
  $builtIn = @(Get-CaProperty $grantControls 'builtInControls')
  $grantOperator = Get-CaProperty $grantControls 'operator'
  $strength = Get-CaProperty $grantControls 'authenticationStrength'
  $termsOfUse = @(Get-CaProperty $grantControls 'termsOfUse')
  $customFactors = @(Get-CaProperty $grantControls 'customAuthenticationFactors')
  $sessionControls = $Policy.sessionControls
  $grantOptions = $builtIn.Count + $termsOfUse.Count + $customFactors.Count
  if (Test-CaConfiguredValue $strength) { $grantOptions++ }

  if ($builtIn -contains 'Block') {
    $control = 'Block'
  }
  elseif ($grantOptions -gt 1 -and $grantOperator -ne 'AND') {
    $control = if ($grantOperator -eq 'OR') { 'AlternativeControls' } else { 'ReviewGrantControls' }
  }
  elseif ($builtIn -contains 'Mfa') {
    if (Test-CaConfiguredValue $strength) {
      $control = 'RequireAuthStrength'
    }
    elseif ($grantOptions -eq 1) {
      $control = 'RequireMFA'
    }
    else {
      $control = 'RequireMFA-Plus'
    }
  }
  elseif ($builtIn -contains 'CompliantDevice') {
    $control = 'RequireCompliantDevice'
  }
  elseif ($builtIn -contains 'DomainJoinedDevice') {
    $control = 'RequireDomainJoined'
  }
  elseif ($builtIn -contains 'ApprovedApplication') {
    $control = 'RequireApprovedApp'
  }
  elseif ($builtIn -contains 'CompliantApplication') {
    $control = 'RequireCompliantApp'
  }
  elseif ($builtIn -contains 'PasswordChange') {
    $control = 'RequirePasswordChange'
  }
  elseif ($builtIn -contains 'riskRemediation') {
    $control = 'RequireRiskRemediation'
  }
  elseif (Test-CaConfiguredValue $strength) {
    $control = 'RequireAuthStrength'
  }
  elseif ($termsOfUse.Count -gt 0) {
    $control = 'RequireToU'
  }
  elseif ($customFactors.Count -gt 0) {
    $control = 'RequireCustomAuth'
  }
  elseif ($grantOptions -gt 0) {
    $control = 'ReviewGrantControls'
  }
  else {
    $control = 'NonBlocking'
  }

  # Add session control modifiers
  $sessionModifiers = @()
  if ($sessionControls.signInFrequency -and $sessionControls.signInFrequency.isEnabled) {
    $sessionModifiers += 'SignInFreq'
  }
  if ($sessionControls.continuousAccessEvaluation -and $sessionControls.continuousAccessEvaluation.mode -eq 'strictEnforcement') {
    $sessionModifiers += 'CAE'
  }
  if ($sessionControls.persistentBrowser -and $sessionControls.persistentBrowser.isEnabled) {
    $sessionModifiers += 'PersistentBrowser'
  }
  if ($sessionControls.applicationEnforcedRestrictions -and $sessionControls.applicationEnforcedRestrictions.isEnabled) {
    $sessionModifiers += 'AppRestrictions'
  }
  if ($sessionControls.cloudAppSecurity -and $sessionControls.cloudAppSecurity.isEnabled) {
    $sessionModifiers += 'CloudAppSec'
  }

  # Append session modifiers to control
  if ($sessionModifiers.Count -gt 0) {
    $control += '-' + ($sessionModifiers -join '-')
  }

  # Add status modifier if reporting only
  if ($Policy.state -eq 'enabledForReportingButNotEnforced') {
    $control += '-ReportOnly'
  }

  # Random display prefix, not a policy identifier.
  $id = '{0:D3}' -f (Get-Random -Maximum 999)

  # Combine components - limit condition to first 3 for readability with enhanced naming
  $conditionStr = ($condition | Select-Object -First 3) -join '-'
  $recommendedName = "$prefix$id-$scope-$conditionStr-$control"

  # Ensure name doesn't exceed reasonable length (max 100 characters for enhanced naming)
  if ($recommendedName.Length -gt 100) {
    # Truncate condition part if too long
    $maxConditionLength = 100 - $prefix.Length - $id.Length - $scope.Length - $control.Length - 3 # 3 hyphens
    if ($conditionStr.Length -gt $maxConditionLength -and $maxConditionLength -gt 0) {
      $conditionStr = $conditionStr.Substring(0, $maxConditionLength)
    }
    $recommendedName = "$prefix$id-$scope-$conditionStr-$control"
  }

  return $recommendedName
}

function Initialize-GraphModule {
  <#
.SYNOPSIS
  Ensure required PowerShell modules are installed and loadable for the current user.
.DESCRIPTION
  Loads core Graph modules or explicitly requested lookup modules, installing missing modules in CurrentUser scope.
  Stops on required-module failures. Optional failures warn once per module and return false so callers skip lookups.
  Does not contact the package repository when all requested modules are already installed.
.PARAMETER RequiredModules
  Module names to validate/install. Defaults to Authentication and Identity.SignIns for connected collection.
.PARAMETER Optional
  Return false instead of throwing if any requested module cannot be installed or imported.
.OUTPUTS
  Boolean indicating whether every requested module loaded successfully.
.EXAMPLE
  Initialize-GraphModule
#>
  [CmdletBinding()]
  param(
    [string[]]$RequiredModules = @(
      'Microsoft.Graph.Authentication',
      'Microsoft.Graph.Identity.SignIns'
    ),
    [switch]$Optional
  )

  $allLoaded = $true
  foreach ($m in $RequiredModules) {
    try {
      $installed = Get-Module -ListAvailable -Name $m -ErrorAction Stop
      if (-not $installed) {
        Write-Info ('Installing module: {0} (CurrentUser)' -f $m)
        Install-Module -Name $m -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
      }
      Import-Module -Name $m -Global -ErrorAction Stop
    }
    catch {
      $message = "Graph module '$m' could not be loaded: $($_.Exception.Message). Preinstall/import this module in a clean PowerShell session; for download/XML/document errors check PSGallery access, proxy authentication and TLS inspection. Do not disable certificate validation."
      if (-not $Optional) { throw "$message Required Graph collection cannot continue." }
      Write-Warn "$message Optional enrichment using this module is skipped; affected IDs remain unresolved."
      $allLoaded = $false
    }
  }
  return $allLoaded
}

function Connect-GraphContext {
  <#
.SYNOPSIS
  Establish a Microsoft Graph connection if one does not already exist.
.DESCRIPTION
  Checks for an existing MgGraph context and, if missing, connects with the required scopes
  (Policy.Read.All, Directory.Read.All, RoleManagement.Read.All).
.PARAMETER TenantID
  Optional tenant GUID. Reconnects if the cached context targets another tenant and verifies the replacement.
.NOTES
  Replaces Ensure-GraphConnection (deprecated) to satisfy approved verb list (Connect-*).
.EXAMPLE
  Connect-GraphContext
  Ensures the current session is connected with required scopes; returns immediately if already connected.
#>
  [CmdletBinding()]
  param([string]$TenantID)
  $ctx = $null
  try {
    $ctx = Get-MgContext -ErrorAction Stop 
  }
  catch {
    Write-Verbose 'No existing Graph context found (Get-MgContext failed).' 
  }
  $requiredScopes = 'Policy.Read.All', 'Directory.Read.All', 'RoleManagement.Read.All'
  $needsConnect = $false
  if (-not $ctx -or -not $ctx.Account) {
    $needsConnect = $true 
  }
  else {
    $granted = @($ctx.Scopes)
    foreach ($r in $requiredScopes) {
      if ($granted -notcontains $r) {
        $needsConnect = $true; break 
      }
    }
  }
  if ($TenantID -and $ctx.TenantId -ne $TenantID) { $needsConnect = $true }
  if ($needsConnect) {
    Write-Info 'Connecting to Microsoft Graph (Policy.Read.All, Directory.Read.All, RoleManagement.Read.All)'
    $connectionParameters = @{ Scopes = $requiredScopes; ErrorAction = 'Stop' }
    if ($TenantID) { $connectionParameters.TenantId = $TenantID }
    Connect-MgGraph @connectionParameters | Out-Null
    $ctx = Get-MgContext -ErrorAction Stop
    if (-not $ctx -or -not $ctx.Account) {
      throw 'Microsoft Graph did not return an authenticated account after connecting. Sign in and retry the export.'
    }
    $missingScopes = @($requiredScopes | Where-Object { $ctx.Scopes -notcontains $_ })
    if ($missingScopes.Count) {
      throw ("Required Microsoft Graph scopes were not granted: {0}. Grant admin consent and reconnect before exporting." -f ($missingScopes -join ', '))
    }
  }
  if ($TenantID -and $ctx.TenantId -ne $TenantID) {
    throw "Microsoft Graph connected to tenant '$($ctx.TenantId)' instead of requested tenant '$TenantID'."
  }
}

function Get-CaNamedLocationInventory {
  [CmdletBinding()]
  param([switch]$Offline)

  if ($Offline) {
    return [pscustomobject]@{
      Status = 'Offline'; Locations = @()
      Message = 'Named location definitions are not included in raw policy snapshots from legacy policy-only exports. Use a portable *_raw.json snapshot from a connected run to include IP ranges and countries/regions offline.'
    }
  }
  try {
    $locations = @(Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop)
    return [pscustomobject]@{ Status = 'Available'; Locations = $locations; Message = '' }
  }
  catch {
    $message = 'Could not retrieve named locations: {0}' -f $_.Exception.Message
    Write-Warn $message
    return [pscustomobject]@{ Status = 'Failed'; Locations = @(); Message = $message }
  }
}

function Import-CaSnapshot {
  <#
  .SYNOPSIS
    Read a portable snapshot or a legacy raw Graph policy array without calling Graph.
  #>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Path)

  try {
    $data = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop |
      ConvertFrom-Json -Depth 100 -NoEnumerate -ErrorAction Stop
    if ($data.PSObject.Properties['SnapshotFormat']) {
      if ($data.SnapshotFormat -ne 'CAExportSnapshot' -or $data.SchemaVersion -ne 1) {
        throw 'Unsupported snapshot format or schema version. Use a compatible exporter version.'
      }
      foreach ($field in @('TenantName', 'CapturedAt', 'PartialExport', 'Policies', 'NamedLocations', 'Lookups', 'AuthenticationStrengths', 'EmergencyAccessAccountIds')) {
        if (-not $data.PSObject.Properties[$field]) { throw "Snapshot is missing required field '$field'." }
      }
      $capturedAt = [datetimeoffset]::MinValue
      if ($data.TenantName -isnot [string] -or [string]::IsNullOrWhiteSpace($data.TenantName) -or
          -not [datetimeoffset]::TryParse([string]$data.CapturedAt, [ref]$capturedAt) -or
          $data.PartialExport -isnot [bool] -or $data.Policies -isnot [array] -or
          $data.AuthenticationStrengths -isnot [array] -or $data.EmergencyAccessAccountIds -isnot [array]) {
        throw 'Snapshot metadata or collections are invalid.'
      }
      if ($data.NamedLocations.Status -notin @('Available', 'Failed', 'Offline') -or
          $data.NamedLocations.Locations -isnot [array]) {
        throw 'Snapshot named-location inventory is invalid.'
      }
      foreach ($name in @('Users', 'Groups', 'Roles', 'Applications', 'Locations', 'TermsOfUse')) {
        if ($data.Lookups.$name -isnot [pscustomobject]) { throw "Snapshot lookup map '$name' is invalid." }
      }
      $policies = @($data.Policies)
    }
    else {
      $policies = @($data)
      $data = [pscustomobject]@{
        Policies = $policies; TenantName = 'OfflineTenant'; CapturedAt = (Get-Date).ToUniversalTime().ToString('u')
        PartialExport = $false
      }
    }
    if (-not $policies.Count) { throw 'The input contains no policies.' }
    foreach ($policy in $policies) {
      if ($policy -isnot [pscustomobject] -or
          $policy.id -isnot [string] -or [string]::IsNullOrWhiteSpace($policy.id) -or
          $policy.state -isnot [string] -or [string]::IsNullOrWhiteSpace($policy.state) -or
          $policy.conditions -isnot [pscustomobject]) {
        throw 'Expected raw Graph policies: each entry must be an object with nonempty string id/state and an object conditions field, not enriched policy JSON.'
      }
    }
    return $data
  }
  catch {
    throw ("Failed to load offline input '{0}': {1}" -f $Path, $_.Exception.Message)
  }
}

function Invoke-SafeGet {
  <#
.SYNOPSIS
  Safely invoke a script block, returning $null on failure.
.DESCRIPTION
  Executes the provided script block with error trapping. Failures are logged to Verbose, or as warnings
  when FailureMessage is supplied, and $null is returned for best-effort enrichment.
.PARAMETER ScriptBlock
  The code to execute.
.PARAMETER FailureMessage
  Optional warning context for enrichment failures that should be visible without verbose logging.
.EXAMPLE
  Invoke-SafeGet { Get-MgUser -UserId $id -Property Id,DisplayName }
#>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][ScriptBlock]$ScriptBlock,
    [string]$FailureMessage
  )
  try {
    & $ScriptBlock 
  }
  catch {
    if ($FailureMessage) {
      Write-Warn ('{0}: {1}' -f $FailureMessage, $_.Exception.Message)
    }
    else {
      Write-Verbose ('Invoke-SafeGet suppressed error: {0}' -f $_.Exception.Message)
    }
    return $null
  }
}


function Convert-IdListToName {
  <#
.SYNOPSIS
  Convert a list of IDs to their friendly names when present in a lookup map.
.DESCRIPTION
  For each element in List, if the Map contains that key the mapped value is output; otherwise the original value.
  Null / empty input yields an empty array.
.PARAMETER List
  Collection of IDs or values.
.PARAMETER Map
  Hashtable keyed by ID with friendly values.
.EXAMPLE
  Convert-IdListToNames -List $policy.conditions.users.includeUsers -Map $UserMap
#>
  [CmdletBinding()]
  param([string[]]$List, [hashtable]$Map)
  if (-not $List) {
    return @() 
  }
  return $List | ForEach-Object { if ($Map.ContainsKey($_)) {
      $Map[$_] 
    }
    else {
      $_ 
    } }
}

function Test-IsGuid {
  <#
.SYNOPSIS
  Test whether a string is a GUID.
.DESCRIPTION
  Wraps [guid]::TryParse for readability and reuse.
  Filters out sentinel tokens (e.g. 'All') earlier in the pipeline.
.EXAMPLE
  Test-IsGuid 'd2719d52-3f4e-4f7c-9d0d-4f5c2a8ab123'
  Returns True.
.EXAMPLE
  Test-IsGuid 'All'
  Returns False.
  .PARAMETER Value
  The string value to test.

#>
  [CmdletBinding()]
  param([string]$Value)
  if (-not $Value) {
    return $false 
  }
  return [bool]([guid]::TryParse($Value, [ref]([guid]::Empty)))
}

# Unified resolver: given a list of IDs (users/groups/roles/apps), return friendly names when available.
function Resolve-EntityNameList {
  <#
  .SYNOPSIS
  Resolve a list of entity IDs to their friendly names.
  .DESCRIPTION
  For each element in Ids, if the Map contains that key the mapped value is output; otherwise the original value.
  Null / empty input yields an empty array.
.PARAMETER Ids
  Collection of IDs to resolve.
.PARAMETER UserMap
  Hashtable mapping user IDs to friendly names.
.PARAMETER GroupMap
  Hashtable mapping group IDs to friendly names.
.PARAMETER RoleMap
  Hashtable mapping role IDs to friendly names.
.PARAMETER AppMap
  Hashtable mapping application IDs to friendly names.
.EXAMPLE
  Convert-IdListToNames -List $policy.conditions.users.includeUsers -Map $UserMap
#>
  [CmdletBinding()]
  param(
    [string[]]$Ids,
    [hashtable]$UserMap,
    [hashtable]$GroupMap,
    [hashtable]$RoleMap,
    [hashtable]$AppMap
  )
  if (-not $Ids) {
    return @() 
  }
  return $Ids | ForEach-Object {
    $id = $_
    if ([string]::IsNullOrWhiteSpace($id)) {
      return $id 
    }
    if ($UserMap -and $UserMap.ContainsKey($id)) {
      return $UserMap[$id] 
    }
    if ($GroupMap -and $GroupMap.ContainsKey($id)) {
      return $GroupMap[$id] 
    }
    if ($RoleMap -and $RoleMap.ContainsKey($id)) {
      return $RoleMap[$id] 
    }
    if ($AppMap -and $AppMap.ContainsKey($id)) {
      return $AppMap[$id] 
    }
    return $id
  }
}

function Resolve-EntityGuidsInText {
  <#
  .SYNOPSIS
  Replace GUIDs in text with friendly names from provided maps.
  .DESCRIPTION
  Scans the input text for GUIDs and replaces them with their corresponding friendly names
  .PARAMETER Text
  The input text containing potential GUIDs to resolve.
  .PARAMETER UserMap
  Hashtable mapping user IDs to friendly names.
  .PARAMETER GroupMap
  Hashtable mapping group IDs to friendly names.
  .PARAMETER RoleMap
  Hashtable mapping role IDs to friendly names.
  .PARAMETER AppMap
  Hashtable mapping application IDs to friendly names.

  #>
  [CmdletBinding()]
  param(
    [string]$Text,
    [hashtable]$UserMap,
    [hashtable]$GroupMap,
    [hashtable]$RoleMap,
    [hashtable]$AppMap
  )
  # Touch parameters to satisfy static analyzers; they are also used within the regex scriptblock below
  $null = $UserMap; $null = $GroupMap; $null = $RoleMap; $null = $AppMap
  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $Text 
  }
  # Standard GUID pattern
  $pattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
  return ([regex]::Replace($Text, $pattern, {
        param($m)
        $g = $m.Value
        if ($UserMap -and $UserMap.ContainsKey($g)) {
          return $UserMap[$g] 
        }
        if ($GroupMap -and $GroupMap.ContainsKey($g)) {
          return $GroupMap[$g] 
        }
        if ($RoleMap -and $RoleMap.ContainsKey($g)) {
          return $RoleMap[$g] 
        }
        if ($AppMap -and $AppMap.ContainsKey($g)) {
          return $AppMap[$g] 
        }
        return $g
      }))
}

#region Helper: Admin role & authentication strength evaluation (added in 3.3)
function Test-PolicyTargetsAdminRoles {
  <#
    .SYNOPSIS
      Determines whether a Conditional Access policy targets any privileged/admin roles.
    .DESCRIPTION
      Resolves included role GUIDs to display names (when cached lookups are available) and checks
      against a curated allowlist of administrative role names (case-insensitive).
    .PARAMETER Policy
      A raw Conditional Access policy object (as returned by Microsoft Graph) with Conditions.Users.IncludeRoles.
    .OUTPUTS
      [bool] True if policy targets at least one administrative role, otherwise False.
  #>
  param($Policy)
  if (-not $Policy -or -not $Policy.Conditions.Users.IncludeRoles) { return $false }
  $rawRoleIds = @($Policy.Conditions.Users.IncludeRoles)
  $resolved = @()
  foreach ($rid in $rawRoleIds) {
    if ([string]::IsNullOrWhiteSpace($rid)) { continue }
    if ($script:RoleMap -and $RoleMap.ContainsKey($rid)) { $resolved += $RoleMap[$rid] }
    elseif ($rid -match '^[0-9a-fA-F-]{36}$' -and $script:roleLookup -and $roleLookup.ContainsKey($rid)) { $resolved += $roleLookup[$rid] }
    else { $resolved += $rid }
  }
  $resolvedLower = $resolved | ForEach-Object { ($_ | Out-String).Trim().ToLowerInvariant() }
  $adminTargets = @(
    'privileged role administrator', 'global administrator', 'privileged authentication administrator', 'security administrator',
    'sharepoint administrator', 'exchange administrator', 'conditional access administrator', 'helpdesk administrator',
    'billing administrator', 'user administrator', 'authentication administrator', 'application administrator',
    'cloud application administrator', 'password administrator'
  )
  return (($resolvedLower | Where-Object { $adminTargets -contains $_ }).Count -gt 0)
}

function Initialize-AuthStrengthCache {
  if ($script:AuthStrengthCachePopulated) { return }
  $script:AuthStrengthCache = @{}
  if ($script:OfflineMode) {
    # Offline: cannot query strengths; leave empty cache so phish-resistant checks default to False.
    $script:AuthStrengthCachePopulated = $true
    return
  }
  $strengths = Invoke-SafeGet -FailureMessage 'Could not retrieve authentication strengths; custom-strength assessments may require review' {
    Get-MgPolicyAuthenticationStrengthPolicy -All -Property Id, DisplayName, AllowedCombinations
  }
  if ($strengths) {
    foreach ($s in $strengths) {
      if ($s.Id) { $script:AuthStrengthCache[$s.Id] = $s }
      if ($s.DisplayName) { $script:AuthStrengthCache[$s.DisplayName] = $s }
    }
  }
  $script:AuthStrengthCachePopulated = $true
}

function Get-AssignedAuthStrengthObject {
  param($Policy)
  $as = Get-CaProperty $Policy 'grantControls.authenticationStrength'
  if (-not $as) { return $null }
  Initialize-AuthStrengthCache
  if (Get-CaProperty $as 'allowedCombinations') { return $as }
  $id = [string](Get-CaProperty $as 'id')
  $name = [string](Get-CaProperty $as 'displayName')
  $cached = if ($id -and $script:AuthStrengthCache.ContainsKey($id)) { $script:AuthStrengthCache[$id] }
  elseif ($name -and $script:AuthStrengthCache.ContainsKey($name)) { $script:AuthStrengthCache[$name] }
  if ($cached -and ((Get-CaProperty $cached 'allowedCombinations') -or -not (Get-CaProperty $as 'requirementsSatisfied'))) {
    return $cached
  }
  return $as
}

function Test-IsPhishResistantStrength {
  <#
    .SYNOPSIS
      Determines if a policy's assigned Authentication Strength is phishing resistant.
    .DESCRIPTION
      Uses the same verified strength IDs and allowed combinations as the recommendation assessments.
    .PARAMETER Policy
      Raw Conditional Access policy object with GrantControls.AuthenticationStrength populated.
    .OUTPUTS
      [bool] True when strength is considered phishing resistant; otherwise False.
  #>
  param($Policy)
  return (Test-CaMfaStrength -Policy $Policy -PhishingResistant)
}

function Test-PolicyRequiresMfaForAdmins {
  param($Policy)
  if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }
  return (@(Get-CaProperty $Policy 'grantControls.builtInControls') -contains 'mfa' -or (Test-CaMfaStrength -Policy $Policy))
}

function Test-PolicyRequiresPhishResistantMfaForAdmins {
  param($Policy)
  if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }
  return (Test-IsPhishResistantStrength -Policy $Policy) 
}

function Test-OverlapIncludeExclude {
  param($Include, $Exclude)
  if (-not $Include -or -not $Exclude) { return $false }
  $inc = @($Include) | Where-Object { $_ -ne $null -and $_ -ne '' }
  $exc = @($Exclude) | Where-Object { $_ -ne $null -and $_ -ne '' }
  if ($inc.Count -eq 0 -or $exc.Count -eq 0) { return $false }
  foreach ($i in $inc) { if ($exc -contains $i) { return $true } }
  return $false 
}

function New-TokenSet {
  param($Value)
  $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  if ($null -ne $Value) { foreach ($tok in (@($Value) -join "`n") -split '[,\n]') { $t = ($tok).Trim(); if ($t) { $null = $set.Add($t) } } }
  return $set 
}

function Protect-RecNote {
  param([string]$Html)
  if ([string]::IsNullOrWhiteSpace($Html)) { return $Html }
  # The note markup we inject is fully script-controlled (policy detail cards); heavy encoding caused visible artifacts.
  # Strategy: strip executable/script/style content + dangerous inline handlers + javascript: URIs, otherwise pass through.
  $clean = $Html -replace '(?is)<script[^>]*>.*?</script>', '' -replace '(?is)<style[^>]*>.*?</style>', ''
  $clean = $clean -replace '(?i) on[a-z]+\s*=\s*"[^"]*"', '' -replace "(?i) on[a-z]+\s*=\s*'[^']*'", ''
  $clean = $clean -replace '(?i)href\s*=\s*"javascript:[^"]*"', 'href="#"' -replace "(?i)href\s*=\s*'javascript:[^']*'", "href='#'"
  return $clean 
}
#endregion Helper additions

#region CA Check Functions (extracted from inline hashtable for testability)
function Test-CA00 {
  param($PolicyCheck)
  $PolicyCheck.GrantControls.BuiltInControls -contains 'Block' -and
  $PolicyCheck.Conditions.ClientAppTypes -contains 'exchangeActiveSync' -and
  $PolicyCheck.Conditions.ClientAppTypes -contains 'other'
}
function Test-CA01 {
  param($PolicyCheck)
  (Test-CaRequiredGrant -Policy $PolicyCheck -Controls @('mfa') -AllowMfaStrength) -and
  $PolicyCheck.Conditions.Users.IncludeUsers -eq 'all' -and
  $PolicyCheck.Conditions.Applications.IncludeApplications -eq 'all'
}
function Test-CA02 {
  param($PolicyCheck)
  ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'android' -or
  $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'iOS' -or
  $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'windowsPhone') -and
  (Test-CaRequiredGrant -Policy $PolicyCheck -Controls @('compliantApplication', 'compliantDevice'))
}
function Test-CA03 {
  param($PolicyCheck)
  ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'windows' -or
  $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'macOS') -and
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'compliantDevice' -or
  $PolicyCheck.GrantControls.BuiltInControls -contains 'domainJoinedDevice')
}
function Test-CA06 {
  param($PolicyCheck)
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Users.IncludeUsers $PolicyCheck.Conditions.Users.ExcludeUsers) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Users.IncludeGroups $PolicyCheck.Conditions.Users.ExcludeGroups) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Users.IncludeRoles $PolicyCheck.Conditions.Users.ExcludeRoles) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Platforms.IncludePlatforms $PolicyCheck.Conditions.Platforms.ExcludePlatforms) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Locations.IncludeLocations $PolicyCheck.Conditions.Locations.ExcludeLocations) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Applications.IncludeApplications $PolicyCheck.Conditions.Applications.ExcludeApplications)
}

function Test-CA07 {
  param($PolicyCheck)
  ([string]::IsNullOrWhiteSpace($PolicyCheck.Conditions.Users.IncludeUsers) -or $PolicyCheck.Conditions.Users.IncludeUsers.Count -eq 0 -or $PolicyCheck.Conditions.Users.IncludeUsers -eq 'None') -and
  (([string]::IsNullOrWhiteSpace($PolicyCheck.Conditions.Users.IncludeGroups)) -or $PolicyCheck.Conditions.Users.IncludeGroups.Count -eq 0) -and
  (([string]::IsNullOrWhiteSpace($PolicyCheck.Conditions.Users.IncludeRoles)) -or $PolicyCheck.Conditions.Users.IncludeRoles.Count -eq 0) -and
  ($null -eq $PolicyCheck.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes)
}

function Test-CA08 {
  param($PolicyCheck)
  $PolicyCheck.Conditions.Users.IncludeUsers -ne 'None' -and
  $null -ne $PolicyCheck.Conditions.Users.IncludeUsers -and
  $PolicyCheck.Conditions.Users.IncludeUsers -ne 'All' -and
  $PolicyCheck.Conditions.Users.IncludeUsers -ne 'GuestsOrExternalUsers'
}
function Test-CA09 {
  param($PolicyCheck)
  (@($PolicyCheck.Conditions.SignInRiskLevels | Where-Object { $_ -in @('medium', 'high') }).Count -gt 0) -and
  (Test-CaRequiredGrant -Policy $PolicyCheck -Controls @('mfa') -AllowMfaStrength)
}
function Test-CA10 {
  param($PolicyCheck)
  $PolicyCheck.Conditions.AdditionalProperties.authenticationFlows.Values -split ',' -contains 'deviceCodeFlow' -and
  $PolicyCheck.grantcontrols.BuiltInControls -contains 'Block'
}
function Test-CA11 {
  param($PolicyCheck)
  ($PolicyCheck.Conditions.Applications.IncludeUserActions -contains 'urn:user:registerdevice') -and
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'Mfa')
}
function Test-CA12 {
  param($PolicyCheck)
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'Block') -and
  ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'all') -and
  ($PolicyCheck.Conditions.Platforms.ExcludePlatforms.Count -gt 0)
}
#endregion CA Check Functions
#endregion Helper additions

function Get-CaProperty {
  <#
  .SYNOPSIS
    Read a dotted property path from Graph SDK objects, dictionaries, or offline JSON.
  #>
  param($InputObject, [string]$Path)
  $value = $InputObject
  foreach ($name in $Path.Split('.')) {
    if ($null -eq $value) { return }
    $next = $null
    foreach ($source in @($value, $value.AdditionalProperties)) {
      if ($null -eq $source) { continue }
      if ($source -is [System.Collections.IDictionary]) {
        $candidate = $source[$name]
        if ($null -eq $candidate) {
          foreach ($key in $source.Keys) {
            if ($key -ieq $name -and $null -ne $source[$key]) { $candidate = $source[$key]; break }
          }
        }
      }
      else { $candidate = $source.PSObject.Properties[$name].Value }
      if ($null -eq $next) { $next = $candidate }
      if (Test-CaConfiguredValue $candidate) { $next = $candidate; break }
    }
    $value = $next
  }
  if ($null -ne $value) { return $value }
}

function Test-CaConfiguredValue {
  <#
  .SYNOPSIS
    Distinguish populated configuration from Graph SDK containers containing only nulls or empty collections.
  #>
  param($Value)
  if ($null -eq $Value) { return $false }
  if ($Value -is [string]) { return -not [string]::IsNullOrWhiteSpace($Value) }
  if ($Value -is [ValueType]) { return $true }
  if ($Value -is [System.Collections.IDictionary]) {
    foreach ($key in $Value.Keys) {
      if ($key -notlike '@odata.*' -and (Test-CaConfiguredValue $Value[$key])) { return $true }
    }
  }
  elseif ($Value -is [System.Collections.IEnumerable]) {
    foreach ($item in $Value) { if (Test-CaConfiguredValue $item) { return $true } }
  }
  else {
    foreach ($property in $Value.PSObject.Properties) {
      if ($property.Name -notlike '@odata.*' -and $property.Name -ne 'OdataType' -and (Test-CaConfiguredValue $property.Value)) { return $true }
    }
  }
  return $false
}

function Test-CaMfaStrength {
  <#
  .SYNOPSIS
    Verify MFA or phishing resistance using strength IDs and defined combinations, never display names.
  #>
  param($Policy, [switch]$PhishingResistant)
  $strength = Get-CaProperty $Policy 'grantControls.authenticationStrength'
  if (-not (Test-CaConfiguredValue $strength)) { return $false }
  $id = [string](Get-CaProperty $strength 'id')
  if ($id -eq '00000000-0000-0000-0000-000000000004') { return $true }
  if (-not $PhishingResistant -and $id -in @('00000000-0000-0000-0000-000000000002', '00000000-0000-0000-0000-000000000003')) { return $true }
  if (-not $PhishingResistant -and (Get-CaProperty $strength 'requirementsSatisfied') -eq 'mfa') { return $true }
  $assigned = Get-AssignedAuthStrengthObject -Policy $Policy
  if ($assigned) { $strength = $assigned }
  if (-not $PhishingResistant -and (Get-CaProperty $strength 'requirementsSatisfied') -eq 'mfa') { return $true }
  $combinations = @(Get-CaProperty $strength 'allowedCombinations')
  if (-not $combinations.Count) { return $false }
  $safe = @('fido2', 'windowsHelloForBusiness', 'x509CertificateMultiFactor')
  if (-not $PhishingResistant) {
    $safe += @('deviceBasedPush', 'temporaryAccessPassOneTime', 'temporaryAccessPassMultiUse', 'federatedMultiFactor')
  }
  foreach ($combination in $combinations) {
    if ($combination -in $safe) { continue }
    if (-not $PhishingResistant -and $combination -match '^password,(microsoftAuthenticatorPush|sms|voice|softwareOath|hardwareOath)$') { continue }
    return $false
  }
  return $true
}

function Test-CaRequiredGrant {
  <#
  .SYNOPSIS
    Check whether every permitted grant path requires an acceptable control.
  #>
  param($Policy, [string[]]$Controls, [switch]$AllowMfaStrength, [switch]$PhishingResistant)
  $grant = Get-CaProperty $Policy 'grantControls'
  $builtIn = @(Get-CaProperty $grant 'builtInControls' | Where-Object { $_ -and $_ -ne 'unknownFutureValue' })
  if ($builtIn -contains 'block') { return $false }
  $strength = Get-CaProperty $grant 'authenticationStrength'
  $total = $builtIn.Count + @(Get-CaProperty $grant 'termsOfUse').Count + @(Get-CaProperty $grant 'customAuthenticationFactors').Count
  $accepted = @($builtIn | Where-Object { $_ -in $Controls }).Count
  if (Test-CaConfiguredValue $strength) {
    $total++
    if (($AllowMfaStrength -or $PhishingResistant) -and (Test-CaMfaStrength -Policy $Policy -PhishingResistant:$PhishingResistant)) { $accepted++ }
  }
  if (@(Get-CaProperty $grant 'builtInControls') -contains 'unknownFutureValue') { return $false }
  if (-not $accepted) { return $false }
  $operator = Get-CaProperty $grant 'operator'
  if ($operator -eq 'AND') { return $true }
  if ($operator -eq 'OR') { return $accepted -eq $total }
  return $total -eq 1
}

function Get-CaNamingHint {
  <#
  .SYNOPSIS
    Flag potential name/configuration mismatches without inferring effective access or renaming policies.
  #>
  param($Policy)
  $name = ([string](Get-CaProperty $Policy 'displayName') -creplace '([a-z])([A-Z])', '$1 $2') -replace '[_-]+', ' '
  $controls = @(Get-CaProperty $Policy 'grantControls.builtInControls')
  $blocks = $controls -contains 'block'
  # "Block unmanaged devices" can legitimately use a mandatory compliance grant.
  $blockClaim = @([regex]::Matches($name, '(?i)\bblock(?:s|ed|ing)?\b') | Where-Object {
    $name.Substring(0, $_.Index) -notmatch '\b(?:non|not|never|don''t)\s+$'
  }).Count -gt 0
  $deviceRestriction = $name -match '\b(unmanaged|non[- ]?compliant)\b' -and
    (Test-CaRequiredGrant -Policy $Policy -Controls @('compliantDevice', 'domainJoinedDevice'))
  if ($blockClaim -and -not $blocks -and -not $deviceRestriction) {
    'Name says Block, but no Block grant is configured. Review whether the configured controls express the intended restriction.'
  }
  if ($blockClaim -and ($blocks -or $deviceRestriction) -and (Get-CaProperty $Policy 'state') -ne 'enabled') {
    'Name describes blocking, but this policy is not enabled for enforcement.'
  }
  if ($name -match '\b(?:require|enforce)(?:s|d)?\s+(?:\S+\s+)?MFA\b' -and
      -not (Test-CaRequiredGrant -Policy $Policy -Controls @('mfa') -AllowMfaStrength)) {
    'Name says Require MFA, but mandatory MFA is not verified for every grant path (check OR alternatives and authentication strength).'
  }
  if ($name -match '\ball[- ]users\b') {
    if (@(Get-CaProperty $Policy 'conditions.users.includeUsers') -notcontains 'All') {
      'Name says All users, but the explicit assignment does not include All users.'
    }
    elseif (@(Get-CaProperty $Policy 'conditions.users.excludeUsers').Count -or
            @(Get-CaProperty $Policy 'conditions.users.excludeGroups').Count -or
            @(Get-CaProperty $Policy 'conditions.users.excludeRoles').Count -or
            (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.users.excludeGuestsOrExternalUsers'))) {
      'Name says All users, but identity exclusions are configured; review whether the name should describe exceptions.'
    }
  }
}

function ConvertTo-CaGuestSelectorValue {
  <#
  .SYNOPSIS
    Canonicalize guest-selector properties and known sets without dropping additional restrictions.
  #>
  param($Value, [string]$Path = '')
  if ($null -eq $Value) { return $null }
  if ($Path -in @('guestorexternalusertypes', 'externaltenants.members')) {
    $tokens = @((New-TokenSet $Value) | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object)
    if ($Path -eq 'guestorexternalusertypes') { return $tokens -join ',' }
    return ,$tokens
  }
  if ($Path -eq 'externaltenants.membershipkind') { return ([string]$Value).Trim().ToLowerInvariant() }
  if ($Value -is [string] -or $Value -is [ValueType]) { return $Value }
  if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [System.Collections.IDictionary]) {
    return ,@(foreach ($item in $Value) { ConvertTo-CaGuestSelectorValue -Value $item -Path "$Path[]" })
  }
  $properties = @{}
  foreach ($source in @($Value.AdditionalProperties, $Value)) {
    if ($null -eq $source) { continue }
    $names = if ($source -is [System.Collections.IDictionary]) { @($source.Keys) } else { @($source.PSObject.Properties.Name) }
    foreach ($name in $names) {
      if ($name -eq 'AdditionalProperties') { continue }
      if ($source -is [System.Collections.IDictionary]) { $candidate = $source[$name] }
      else { $candidate = $source.PSObject.Properties[$name].Value }
      if (-not $properties.ContainsKey($name) -or (Test-CaConfiguredValue $candidate)) {
        $properties[$name.ToLowerInvariant()] = $candidate
      }
    }
  }
  $canonical = [ordered]@{}
  foreach ($name in ($properties.Keys | Sort-Object)) {
    $childPath = if ($Path) { "$Path.$name" } else { $name }
    $canonical[$name] = ConvertTo-CaGuestSelectorValue -Value $properties[$name] -Path $childPath
  }
  return $canonical
}

function Get-CaAssignmentToken {
  <#
  .SYNOPSIS
    Preserve assignment type and original identifier for exact cross-policy comparison, with resolved labels.
  #>
  param($Policy, [string]$Column, [hashtable]$Lookups)
  $direction = if ($Column -like 'Included *') { 'include' } else { 'exclude' }
  $fields = switch -Wildcard ($Column) {
    '* Users' { @('users', 'groups', 'roles') }
    '* Applications' { @('applications') }
    '* Locations' { @('locations') }
    default { @() }
  }
  foreach ($type in $fields) {
    $path = switch ($type) {
      'applications' { "conditions.applications.${direction}Applications" }
      'locations' { "conditions.locations.${direction}Locations" }
      default { "conditions.users.$direction$type" }
    }
    foreach ($id in @(Get-CaProperty $Policy $path)) {
      if ([string]::IsNullOrWhiteSpace([string]$id)) { continue }
      $label = [string]$id
      if ($Lookups[$type] -and $Lookups[$type].ContainsKey([string]$id)) { $label = [string]$Lookups[$type][$id] }
      [pscustomobject]@{ Type = $type; Id = [string]$id; Label = $label; Direction = $direction }
    }
  }
  if ($Column -like '* Users') {
    $guests = Get-CaProperty $Policy "conditions.users.${direction}GuestsOrExternalUsers"
    if (Test-CaConfiguredValue $guests) {
      [pscustomobject]@{
        Type = 'guest-selector'; Id = (ConvertTo-Json -InputObject (ConvertTo-CaGuestSelectorValue $guests) -Depth 100 -Compress -WarningAction Stop)
        Label = "Guest selector: $(Get-CaProperty $guests 'guestOrExternalUserTypes')"; Direction = $direction
      }
    }
  }
}

function Get-CaRecommendationCatalog {
  <#
  .SYNOPSIS
    Define Microsoft-backed checks, prerequisites, verification criteria, and applicability limits.
  #>
  $base = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/'
  $p1 = 'Microsoft Entra ID P1 (or a suite including it) for in-scope users; Conditional Access Administrator to configure policies.'
  $p2 = 'Microsoft Entra ID P2 for in-scope users; ID Protection risk signals; registered authentication methods and a tested recovery process.'
  $definitions = @(
    @{ Id='CA-00'; Name='Block legacy authentication'; Users='All'; Resources='All'; Path='policy-block-legacy-authentication'; Background='Legacy clients cannot satisfy modern MFA controls.'; Verify='Block Exchange ActiveSync and other legacy clients for all users and resources.'; Requirements=@($p1, 'Inventory legacy clients and service dependencies before blocking.'); Allow=@('clientAppTypes') },
    @{ Id='CA-01'; Name='Require MFA for all users and resources'; Users='All'; Resources='All'; Path='policy-all-users-mfa-strength'; Background='Broad MFA coverage reduces gaps left by resource-specific policies.'; Verify='Require MFA or an MFA-capable authentication strength on every allowed grant path.'; Requirements=@($p1, 'Register usable authentication methods; plan onboarding and emergency access.') },
    @{ Id='CA-02'; Name='Require mobile device compliance or app protection'; Users='All'; Resources='All'; Path='policy-all-users-approved-app-or-app-protection'; Background='Device compliance and app protection protect organizational data on mobile devices; approved-client-app-only is not a modern substitute.'; Verify='Require compliantDevice or compliantApplication for both Android and iOS, without an unprotected OR alternative.'; Requirements=@($p1, 'Microsoft Intune licensing and assigned compliance or app-protection policies; supported mobile apps.'); Platforms=@('android','iOS'); Allow=@('platforms') },
    @{ Id='CA-03'; Name='Require managed Windows and macOS devices'; Users='All'; Resources='All'; Path='policy-all-users-device-compliance'; Background='Compliant devices provide device-health assurance. Microsoft Entra hybrid join applies to Windows, not macOS.'; Verify='Require compliance for Windows and macOS; Windows-only hybrid join is evidence for review, not proof of macOS protection.'; Requirements=@($p1, 'Intune compliance policies or a supported compliance partner; device registration and enrollment.'); Platforms=@('windows','macOS'); Allow=@('platforms') },
    @{ Id='CA-04'; Name='Require MFA for administrators'; Users='Admins'; Resources='All'; Path='policy-old-require-mfa-admin'; Background='Privileged identities need broad MFA protection, not just protection on one application.'; Verify='Require MFA for all users or the documented administrator role set, across all resources.'; Requirements=@($p1, 'Inventory privileged assignments, including PIM-eligible roles.') },
    @{ Id='CA-05'; Name='Require phishing-resistant MFA for administrators'; Users='Admins'; Resources='All'; Path='policy-admin-phish-resistant-mfa'; Background='Phishing-resistant authentication reduces credential-phishing exposure; ordinary MFA alone does not establish phishing resistance.'; Verify='Require a phishing-resistant strength without an OR bypass.'; Requirements=@($p1, 'Provision FIDO2/passkeys, Windows Hello for Business, or multifactor certificate-based authentication; test bootstrap and recovery.') },
    @{ Id='CA-06'; Name='Review overlapping inclusions and exclusions'; Kind='Hygiene'; Path='concept-conditional-access-users-groups'; Background='Exclusions take precedence, but a partial overlap does not necessarily make the entire policy ineffective.'; Verify='Review exact include/exclude overlaps for users, groups, roles, resources, locations, and platforms.'; Requirements=@($p1, 'Directory membership review is required to detect indirect overlaps.') },
    @{ Id='CA-07'; Name='Review policies without a target identity'; Kind='Hygiene'; Path='concept-conditional-access-users-groups'; Background='Policies need a user, workload, or supported agent assignment to affect access.'; Verify='Distinguish empty user assignments from legitimate workload or agent policies.'; Requirements=@($p1, 'SDK/API schema support for workload and agent assignments.') },
    @{ Id='CA-08'; Name='Review direct user assignments'; Kind='Hygiene'; Path='plan-conditional-access'; Background='Group-based administration can improve maintainability. Direct assignment is not inherently insecure and can be appropriate for emergency accounts.'; Verify='Review direct inclusions as a maintainability finding, not a universal Microsoft security failure.'; Requirements=@($p1, 'Account owners and an approved group-management process.') },
    @{ Id='CA-09'; Name='Require MFA and reauthentication for risky sign-ins'; Users='All'; Resources='All'; Path='policy-risk-based-sign-in'; Background='Sign-in risk and user risk are separate signals and need separate assessments.'; Verify='Cover medium and high sign-in risk, require MFA/authentication strength, and use every-time sign-in frequency.'; Requirements=@($p2, 'Pilot with registered users; risky sessions cannot be used to bootstrap MFA registration.'); Allow=@('signInRiskLevels') },
    @{ Id='CA-10'; Name='Block device code flow where not required'; Kind='Review'; Users='All'; Resources='All'; Path='policy-block-authentication-flows'; Background='Device-code phishing can be reduced by blocking the flow where business dependencies do not require it.'; Verify='Identify deviceCodeFlow blocking policies and validate approved device/client exceptions.'; Requirements=@($p1, 'Inventory device-code dependencies and pilot in report-only mode.'); Allow=@('authenticationFlows') },
    @{ Id='CA-11'; Name='Require MFA for device registration or join'; Users='All'; Resources='DeviceRegistration'; Path='policy-all-users-device-registration'; Background='Register or join devices is a distinct action from registering security information and from Intune enrollment configuration.'; Verify='Protect urn:user:registerdevice with MFA or an MFA-capable strength.'; Requirements=@($p1, 'Coordinate Entra device registration MFA settings and enrollment/bootstrap workflows.') },
    @{ Id='CA-12'; Name='Block unknown or unsupported device platforms'; Users='All'; Resources='All'; Path='policy-all-users-device-unknown-unsupported'; Background='Platform detection is based on client claims; platform blocking is not equivalent to device compliance.'; Verify='Block all platforms except an explicitly reviewed supported-platform list.'; Requirements=@($p1, 'Define supported platforms and test clients that do not send recognizable platform information.'); Allow=@('platforms') },
    @{ Id='CA-13'; Name='Secure security-information registration'; Users='All'; Resources='Registration'; Path='policy-all-users-security-info-registration'; Background='Protect MFA and self-service password-reset registration. This is separate from the device-registration action.'; Verify='Require MFA/authentication strength for urn:user:registersecurityinfo; validate bootstrap and recovery.'; Requirements=@($p1, 'Combined security-information registration; registered methods or Temporary Access Pass; test Windows Hello and macOS Platform SSO registration.') },
    @{ Id='CA-14'; Name='Require remediation for high-risk users'; Users='All'; Resources='All'; Path='policy-risk-based-user'; Background='Current Microsoft guidance supports password-based and passwordless users through Require risk remediation.'; Verify='For high user risk require riskRemediation AND an MFA-capable authentication strength, with every-time sign-in frequency. Review legacy password-change-only remediation separately.'; Requirements=@($p2, 'Risk-remediation feature/API support. Legacy password-change remediation also needs MFA, SSPR and password writeback for applicable hybrid users.'); Allow=@('userRiskLevels') },
    @{ Id='CA-15'; Name='Require managed devices for administrators'; Users='Admins'; Resources='All'; Path='policy-alt-admin-device-compliand-hybrid'; Background='Privileged access should originate from appropriately managed devices, independent of the MFA requirement.'; Verify='Require compliant or hybrid-joined devices for administrator coverage; verify OS support and privileged workstation design.'; Requirements=@($p1, 'Intune/device-compliance licensing where used; registered devices and a tested privileged-device rollout.') },
    @{ Id='CA-16'; Name='Require MFA for guests and external users'; Users='Guests'; Resources='All'; Path='policy-old-require-mfa-guest'; Background='Guest MFA can be satisfied by trusted home-tenant MFA claims when cross-tenant access trust permits.'; Verify='Protect all guest/external populations or all users; review external tenant restrictions and cross-tenant MFA trust.'; Requirements=@($p1, 'Applicable Microsoft Entra External ID billing/licensing; cross-tenant trust and guest onboarding review.') },
    @{ Id='CA-17'; Name='Protect Microsoft admin portals with MFA'; Users='All'; Resources='AdminPortals'; Path='policy-old-require-mfa-admin-portals'; Background='A broad all-resource MFA policy can satisfy this check; a duplicate portal-only policy is not required.'; Verify='Require MFA for MicrosoftAdminPortals or all resources.'; Requirements=@($p1, 'Verify administrative identities, emergency access, and portal sign-in flows.') },
    @{ Id='CA-18'; Name='Protect Azure management with MFA'; Users='All'; Resources='AzureManagement'; Path='policy-old-require-mfa-azure-mgmt'; Background='Azure management includes the Windows Azure Service Management API resource, including dependent management tools.'; Verify='Require MFA for Azure management resource 797f4846-ba00-4fd7-ba43-dac1f8f63013 or all resources.'; Requirements=@($p1, 'Inventory Azure portal, CLI, PowerShell and automation; move unattended user-based automation to workload identities.') },
    @{ Id='CA-19'; Name='Migrate legacy approved-client-app controls'; Kind='Hygiene'; Path='migrate-approved-client-app'; Background='Since June 30, 2026, approved-client-app policies are read-only. Existing enabled policies still enforce; new policies should use app protection.'; Verify='Identify policies containing approvedApplication and plan supported app-protection replacements.'; Requirements=@($p1, 'Intune app-protection licensing, deployed app-protection policies, and client compatibility testing.') },
    @{ Id='CA-20'; Name='Restrict unmanaged-device access'; Kind='Review'; Users='All'; Resources='Any'; Path='policy-all-users-app-enforced-restrictions'; Background='Application-enforced restrictions or compliant-device grants can reduce unmanaged-device exposure; application-side settings determine actual restrictions.'; Verify='Review compliance grants or enabled application-enforced restrictions and verify the corresponding SharePoint/Exchange configuration.'; Requirements=@($p1, 'Supported applications and application-side unmanaged-device restrictions; Intune licensing where compliance is required.') },
    @{ Id='CA-21'; Name='Control browser persistence and sensitive sessions'; Kind='Review'; Users='All'; Resources='Any'; Path='policy-all-users-persistent-browser'; Background='Session controls should match resource sensitivity. A universal short sign-in timeout is not recommended by this report.'; Verify='Review never-persistent browser sessions and enabled sign-in frequency, including scope and user impact.'; Requirements=@($p1, 'Define sensitive/admin scenarios and test browser persistence, reauthentication, and application compatibility.'); Allow=@('clientAppTypes') },
    @{ Id='CA-22'; Name='Validate designated emergency-access exclusions'; Kind='Emergency'; Path='../role-based-access-control/security-emergency-access'; Background='Policy exports cannot identify or test emergency accounts automatically. Avoid dependence on the same authentication and device infrastructure as everyday administration.'; Verify='Supply -EmergencyAccessAccountId with object IDs; review direct exclusions on potentially applicable user policies. Group-based exclusions need membership verification.'; Requirements=@($p1, 'Designate at least two emergency-access accounts; protect, monitor, and regularly test them. IDs alone do not prove account readiness.') },
    @{ Id='CA-23'; Name='Review token protection'; Kind='Review'; Users='All'; Resources='Any'; Path='concept-token-protection'; Background='Device-bound sign-in session tokens reduce token replay risk for supported resources, devices, and clients.'; Verify='Review secureSignInSession.isEnabled, then validate the current platform/resource/client support matrix and sign-in results.'; Requirements=@($p1, 'Supported device registration, operating systems, resources and current clients; Apple scenarios require the documented SSO/MDM setup; pilot first.'); Extra=@{ 'Windows prerequisites and compatibility'="$($base)deployment-guide-token-protection-windows"; 'Apple deployment'="$($base)deployment-guide-token-protection-apple"; 'Web-app limitations'="$($base)deployment-guide-token-protection-web-apps" } },
    @{ Id='CA-24'; Name='Review workload-identity protection'; Kind='Review'; Users='Workloads'; Resources='All'; Path='workload-identity'; Background='User-scoped MFA policies do not protect service principals. Workload policies support blocking based on location or service-principal risk.'; Verify='Find directly targeted service-principal block policies with risk or location conditions; inventory uncovered workloads.'; Requirements=@('Microsoft Entra Workload ID Premium to create or modify these policies; Conditional Access Administrator.', 'Only supported tenant-owned single-tenant service principals; managed identities and third-party/multitenant apps are not covered. Group assignment does not enforce policy for a service principal.'); Allow=@('servicePrincipalRiskLevels','locations') },
    @{ Id='CA-25'; Name='Review insider-risk-based access'; Kind='Review'; Users='All'; Resources='All'; Path='policy-risk-based-insider-block'; Background='Microsoft Purview Adaptive Protection can supply insider-risk signals to Conditional Access.'; Verify='Review elevated insider-risk block policies and confirm that the Purview signal is configured and populated.'; Requirements=@($p1, 'Appropriate Microsoft Purview Insider Risk Management/Adaptive Protection licensing and enabled Adaptive Protection.'); Allow=@('insiderRiskLevels'); Extra=@{ 'Adaptive Protection prerequisites'='https://learn.microsoft.com/en-us/purview/insider-risk-management-adaptive-protection' } },
    @{ Id='CA-26'; Name='Review autonomous agent access and high-risk agents'; Kind='Review'; Users='Agents'; Resources='All'; Path='policy-autonomous-agents'; Background='Autonomous agents authenticate as agent identities, not signed-in users. Risk-based agent controls have separate availability constraints.'; Verify='Review agent-targeted allow-list/block and high-risk protections in Entra; this exporter cannot certify agent coverage from the v1.0 policy snapshot.'; Requirements=@('Microsoft Agent 365 plus the documented Entra Conditional Access license combination; P2 for agent risk. Consult current product licensing guidance.', 'Registered agent identities using supported autonomous flows; Conditional Access Administrator; verify preview availability and agent inventory.'); Extra=@{ 'Agent boundaries and limitations'="$($base)agent-id"; 'Agent assignment configuration'="$($base)howto-target-agent-identities" } },
    @{ Id='CA-27'; Name='Review on-behalf-of agent access'; Kind='Review'; Users='Agents'; Resources='All'; Path='policy-on-behalf-of-agents'; Background='Delegated agent access and autonomous agent access have different identity and policy evaluation boundaries.'; Verify='Review delegated user and agent policies, supported agent flows, resource scope and sign-in evidence; do not infer coverage from a user MFA policy alone.'; Requirements=@('Microsoft Agent 365 and the documented Entra licensing and supported agent configuration.', 'Inventory agents acting for users and validate supported service boundaries; this v1.0 snapshot cannot certify coverage.'); Extra=@{ 'Conditional Access for agents'="$($base)agent-id" } },
    @{ Id='CA-28'; Name='Microsoft 365: block unmanaged-device access'; Kind='Template'; Users='All'; Resources='Microsoft365'; Path='policy-all-users-device-compliance'; Background='An optional resource-scoped alternative to an all-resource device baseline. The Office 365 resource grouping covers Microsoft 365 service dependencies; one service alone does not establish suite-wide coverage. Compliance is stronger than management or registration alone.'; Verify='Review the matching deployment state, user/resource exclusions, device identity in supported browsers, and real sign-ins. Recognized patterns require compliance (optionally hybrid join), or block with a simple compliant-device exclusion/negative filter. Complex filters need manual review.'; Requirements=@($p1, 'Intune or supported partner licensing and assigned compliance policies; registered/enrolled devices; supported browser/client device authentication.', 'Pilot enrollment, guests, emergency access and service dependencies before blocking. Blocking everyone from an app is not evidence of device-specific protection.'); Allow=@('devices'); Template=[ordered]@{ 'Users'='All users, with designated emergency accounts excluded; review any approved population changes.'; 'Resources'='Office 365 (Graph token Office365), or a stronger all-resource policy; no unreviewed exclusions.'; 'Device / grant'='Require compliantDevice (optionally OR domainJoinedDevice for hybrid Windows); alternatively Block with device filter exclude: device.isCompliant -eq True, or include: device.isCompliant -ne True.'; 'Rollout'='Start report-only; verify unmanaged and unregistered devices are denied, and approved devices can sign in.' } },
    @{ Id='CA-29'; Name='Selected applications: block unmanaged-device access'; Kind='Template'; Users='All'; Resources='SelectedApps'; Path='policy-all-users-device-compliance'; Background='Sensitive enterprise applications can have stricter device requirements without applying the same restriction to every resource. This template inventories explicit resource assignments; it cannot discover which business applications ought to be protected.'; Verify='Confirm the application IDs in evidence match your sensitive-app inventory. Target resource/service applications, not public client apps such as Outlook. Validate dependencies and actual users; a match is not complete coverage of your application estate.'; Requirements=@($p1, 'Defined sensitive-resource inventory and owners; Intune/supported compliance partner, assigned compliance policies, device registration and supported clients.', 'Test application dependencies, guest access, emergency recovery and report-only results. A device marked managed or registered is not necessarily compliant.'); Allow=@('devices'); Template=[ordered]@{ 'Users'='All users of the selected resources; review scoped groups and exclusions separately.'; 'Resources'='Explicitly selected resource application IDs. This measurement reports observed assignments, not coverage of an externally defined inventory.'; 'Device / grant'='Require compliantDevice (optionally OR domainJoinedDevice); or Block using the simple compliant-device filter patterns documented in CA-28.'; 'Rollout'='Pilot each application and its service dependencies; verify both registered non-compliant and unregistered devices.' } },
    @{ Id='CA-30'; Name='SharePoint / OneDrive: limited browser access'; Kind='Template'; Users='All'; Resources='SharePoint'; Path='policy-all-users-app-enforced-restrictions'; Background='Limited web-only access is an alternative to total blocking: supported SharePoint/OneDrive configurations restrict downloading, printing and syncing on unmanaged devices. A Conditional Access session control alone does not prove that these restrictions are configured.'; Verify='Verify SharePoint tenant/site ConditionalAccessPolicy settings and the companion policy blocking unsupported desktop/mobile access. Confirm download, print, sync and editing behavior. This export cannot read those SharePoint settings.'; Requirements=@($p1, 'SharePoint Online licensing and SharePoint Administrator to configure unmanaged-device controls; supported OS/browser combinations.', 'Configure SharePoint tenant/site limited-access settings and protect legacy/non-browser clients. SharePoint-only restrictions do not protect Exchange or every Microsoft 365 service.'); Template=[ordered]@{ 'Users'='All users or an explicitly approved population; validate guests and exclusions.'; 'Resources'='SharePoint Online (00000003-0000-0ff1-ce00-000000000000), Office 365, or All resources.'; 'Session'='Enable application-enforced restrictions for browser access; configure SharePoint unmanaged-device limited web-only access.'; 'Companion controls'='Review the companion non-browser block/device policy and legacy authentication block. These and SharePoint-side settings are not inferred from the session control.' } },
    @{ Id='CA-31'; Name='Microsoft 365 mobile: require app protection'; Kind='Template'; Users='All'; Resources='Microsoft365'; Path='policy-all-users-approved-app-or-app-protection'; Background='App protection can protect organizational data on supported personal iOS/Android devices without requiring full device enrollment. This is a different business choice from blocking all unmanaged devices.'; Verify='Require compliantApplication on every allowed grant path for Android and iOS. Validate assigned Intune app-protection policies, broker registration and supported Microsoft 365 apps. Approved-client-app alone or compliance OR app protection does not establish mandatory app protection.'; Requirements=@($p1, 'Microsoft Intune licenses and assigned app-protection policies for the intended users and supported apps.', 'Supported Android/iOS applications, required authentication broker/device registration and current client versions; pilot personal and corporate devices.'); Platforms=@('android','iOS'); Allow=@('platforms'); Template=[ordered]@{ 'Users'='All intended Microsoft 365 mobile users; review assignment groups and exceptions.'; 'Resources'='Office 365 or All resources; inspect dependencies and resource exclusions.'; 'Conditions'='Android and iOS; Mobile apps and desktop clients (or all clients). Browser-only scope is not a match.'; 'Grant'='Require app protection policy (compliantApplication), with no OR alternative that bypasses it. Deploy the supporting Intune app-protection policy.' } }
  )
  foreach ($definition in $definitions) {
    $links = [ordered]@{
      'Microsoft guidance' = "$base$($definition.Path)"
      'Conditional Access deployment and licensing' = "$($base)overview"
      'Report-only testing' = "$($base)concept-conditional-access-report-only"
    }
    if ($definition.Id -eq 'CA-22') { $links['Microsoft guidance'] = 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access' }
    if ($definition.Id -eq 'CA-14') { $links['Grant-control prerequisites and API limitations'] = 'https://learn.microsoft.com/en-us/graph/api/resources/conditionalaccessgrantcontrols?view=graph-rest-beta' }
    switch ($definition.Id) {
      { $_ -in @('CA-01','CA-04','CA-05','CA-09','CA-13','CA-14') } {
        $links['Authentication strengths and supported methods'] = 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-authentication-strengths'
      }
      { $_ -in @('CA-02','CA-19','CA-31') } {
        $links['Intune app-protection prerequisites'] = 'https://learn.microsoft.com/en-us/intune/intune-service/apps/app-protection-policy'
        $links['Approved-client-app migration'] = "$($base)migrate-approved-client-app"
      }
      { $_ -in @('CA-03','CA-15','CA-20','CA-28','CA-29') } {
        $links['Device compliance prerequisites'] = 'https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started'
      }
      'CA-13' { $links['Temporary Access Pass bootstrap'] = 'https://learn.microsoft.com/en-us/entra/identity/authentication/howto-authentication-temporary-access-pass' }
      'CA-14' { $links['Hybrid password writeback prerequisites'] = 'https://learn.microsoft.com/en-us/entra/identity/authentication/tutorial-enable-sspr-writeback' }
      'CA-16' { $links['Cross-tenant MFA and device trust'] = 'https://learn.microsoft.com/en-us/entra/external-id/cross-tenant-access-overview' }
      'CA-27' {
        $links['Microsoft guidance'] = "$($base)agent-id#agents-that-act-on-behalf-of-a-user"
        $links['On-behalf-of OAuth flow'] = 'https://learn.microsoft.com/en-us/entra/agent-id/agent-on-behalf-of-oauth-flow'
        $links['Agents using their own user accounts'] = "$($base)policy-agent-user"
      }
      { $_ -in @('CA-28','CA-29','CA-30','CA-31') } {
        $links['Resource targeting and Microsoft 365 suite'] = "$($base)concept-conditional-access-cloud-apps#office-365"
        $links['Service dependencies'] = "$($base)service-dependencies"
      }
      { $_ -in @('CA-28','CA-29') } {
        $links['Device filters and unregistered-device behavior'] = "$($base)concept-condition-filters-for-devices"
      }
      'CA-30' {
        $links['SharePoint unmanaged-device setup and prerequisites'] = 'https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices'
        $links['SharePoint application-enforced restrictions'] = 'https://learn.microsoft.com/en-us/sharepoint/app-enforced-restrictions'
      }
    }
    if ($definition.Extra) { foreach ($key in $definition.Extra.Keys) { $links[$key] = $definition.Extra[$key] } }
    [pscustomobject]@{
      Control = $definition.Id; Name = $definition.Name
      Kind = $(if ($definition.Kind) { $definition.Kind } else { 'Protection' })
      Users = $definition.Users; Resources = $definition.Resources
      Requirements = @($definition.Requirements); Background = $definition.Background
      Verification = $definition.Verify; Links = $links
      TemplateSettings = $definition.Template
      AllowedConditions = @($definition.Allow | Where-Object { $_ }); Platforms = @($definition.Platforms | Where-Object { $_ })
    }
  }
}

function Get-CaDeviceRestrictionPattern {
  <#
  .SYNOPSIS
    Recognize conservative device-restriction patterns, including filters that cover unregistered devices.
  #>
  param($Policy)
  $filter = Get-CaProperty $Policy 'conditions.devices.deviceFilter'
  if (-not (Test-CaConfiguredValue $filter)) {
    if (Test-CaRequiredGrant $Policy @('compliantDevice')) { return 'Required device compliance' }
    if (Test-CaRequiredGrant $Policy @('compliantDevice','domainJoinedDevice')) { return 'Required compliant or hybrid-joined device (hybrid join is not compliance)' }
    return ''
  }
  $controls = @(Get-CaProperty $Policy 'grantControls.builtInControls')
  if ($controls.Count -ne 1 -or $controls[0] -ne 'block') { return '' }
  $rule = ([string]$filter.rule).Trim()
  while ($rule.StartsWith('(') -and $rule.EndsWith(')')) { $rule = $rule.Substring(1, $rule.Length - 2).Trim() }
  if (($filter.mode -eq 'exclude' -and $rule -match '^device\.isCompliant\s+-eq\s+(true|"true"|''true'')$') -or
      ($filter.mode -eq 'include' -and $rule -match '^device\.isCompliant\s+-ne\s+(true|"true"|''true'')$')) {
    return 'Block non-compliant and unregistered devices using a compliant-device filter'
  }
  return ''
}

function Test-CaRecommendationSignal {
  <#
  .SYNOPSIS
    Find relevant policy configuration; scope and enforcement are assessed separately.
  #>
  param($Policy, [string]$Control)
  $builtIn = @(Get-CaProperty $Policy 'grantControls.builtInControls')
  $hasMfa = $builtIn -contains 'mfa' -or (Test-CaConfiguredValue (Get-CaProperty $Policy 'grantControls.authenticationStrength'))
  switch ($Control) {
    'CA-00' { return ($builtIn -contains 'block' -and @('exchangeActiveSync','other' | Where-Object { $_ -in @(Get-CaProperty $Policy 'conditions.clientAppTypes') }).Count -gt 0) }
    { $_ -in @('CA-01','CA-04','CA-05','CA-16') } { return $hasMfa }
    'CA-11' { return ($hasMfa -and @(Get-CaProperty $Policy 'conditions.applications.includeUserActions') -contains 'urn:user:registerdevice') }
    'CA-13' { return ($hasMfa -and @(Get-CaProperty $Policy 'conditions.applications.includeUserActions') -contains 'urn:user:registersecurityinfo') }
    'CA-17' { return ($hasMfa -and @(Get-CaProperty $Policy 'conditions.applications.includeApplications' | Where-Object { $_ -in @('All','MicrosoftAdminPortals') }).Count -gt 0) }
    'CA-18' { return ($hasMfa -and @(Get-CaProperty $Policy 'conditions.applications.includeApplications' | Where-Object { $_ -in @('All','797f4846-ba00-4fd7-ba43-dac1f8f63013') }).Count -gt 0) }
    { $_ -in @('CA-02','CA-03','CA-15') } { return @($builtIn | Where-Object { $_ -in @('compliantApplication','compliantDevice','domainJoinedDevice') }).Count -gt 0 }
    'CA-06' { return [bool](Test-CA06 $Policy) }
    'CA-07' {
      $users = Get-CaProperty $Policy 'conditions.users'
      $targets = @(Get-CaProperty $users 'includeUsers' | Where-Object { $_ -and $_ -ne 'None' }) + @(Get-CaProperty $users 'includeGroups') + @(Get-CaProperty $users 'includeRoles')
      $clientTargets = @(Get-CaProperty $Policy 'conditions.clientApplications.includeServicePrincipals') + @(Get-CaProperty $Policy 'conditions.clientApplications.includeAgentIdServicePrincipals')
      $clientFilter = (Get-CaProperty $Policy 'conditions.clientApplications.servicePrincipalFilter.rule') -or (Get-CaProperty $Policy 'conditions.clientApplications.agentIdServicePrincipalFilter.rule')
      return (-not $targets.Count -and -not (Test-CaConfiguredValue (Get-CaProperty $users 'includeGuestsOrExternalUsers')) -and -not $clientTargets.Count -and -not $clientFilter -and -not (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.agents')))
    }
    'CA-08' { return @((Get-CaProperty $Policy 'conditions.users.includeUsers') | Where-Object { Test-IsGuid $_ }).Count -gt 0 }
    'CA-09' { return (@(Get-CaProperty $Policy 'conditions.signInRiskLevels' | Where-Object { $_ -in @('medium','high') }).Count -gt 0 -and $hasMfa) }
    'CA-10' {
      $flows = [string](Get-CaProperty $Policy 'conditions.authenticationFlows.transferMethods')
      return ($builtIn -contains 'block' -and 'deviceCodeFlow' -in ($flows -split ',\s*'))
    }
    'CA-12' { return [bool](Test-CA12 $Policy) }
    'CA-14' { return ($builtIn -contains 'riskRemediation' -or $builtIn -contains 'passwordChange' -or ($builtIn -contains 'unknownFutureValue' -and (Get-CaProperty $Policy 'conditions.userRiskLevels'))) }
    'CA-19' { return $builtIn -contains 'approvedApplication' }
    'CA-20' { return ($builtIn -contains 'compliantDevice' -or ($builtIn -contains 'block' -and (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.devices.deviceFilter'))) -or (Get-CaProperty $Policy 'sessionControls.applicationEnforcedRestrictions.isEnabled') -eq $true) }
    'CA-21' {
      return (((Get-CaProperty $Policy 'sessionControls.persistentBrowser.isEnabled') -eq $true -and (Get-CaProperty $Policy 'sessionControls.persistentBrowser.mode') -eq 'never') -or (Get-CaProperty $Policy 'sessionControls.signInFrequency.isEnabled') -eq $true)
    }
    'CA-23' { return (Get-CaProperty $Policy 'sessionControls.secureSignInSession.isEnabled') -eq $true }
    'CA-24' { return ($builtIn -contains 'block' -and (Get-CaProperty $Policy 'conditions.clientApplications.includeServicePrincipals') -and ((Get-CaProperty $Policy 'conditions.servicePrincipalRiskLevels') -or (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.locations')))) }
    'CA-25' { return ($builtIn -contains 'block' -and 'elevated' -in (([string](Get-CaProperty $Policy 'conditions.insiderRiskLevels')) -split ',\s*')) }
    'CA-26' { return [bool]((Get-CaProperty $Policy 'conditions.clientApplications.includeAgentIdServicePrincipals') -or (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.clientApplications.agentIdServicePrincipalFilter')) -or (Get-CaProperty $Policy 'conditions.agentIdRiskLevels')) }
    'CA-27' { return $false }
    { $_ -in @('CA-28','CA-29','CA-30','CA-31') } {
      $resources = @(Get-CaProperty $Policy 'conditions.applications.includeApplications')
      $suite = $resources -contains 'All' -or $resources -contains 'Office365'
      switch ($Control) {
        'CA-28' { return ($suite -and ($builtIn -contains 'compliantDevice' -or $builtIn -contains 'domainJoinedDevice' -or ($builtIn -contains 'block' -and (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.devices.deviceFilter'))))) }
        'CA-29' { return (@($resources | Where-Object { $_ -and $_ -notin @('All','None','Office365') }).Count -gt 0 -and $resources -notcontains 'All' -and ($builtIn -contains 'compliantDevice' -or $builtIn -contains 'domainJoinedDevice' -or ($builtIn -contains 'block' -and (Test-CaConfiguredValue (Get-CaProperty $Policy 'conditions.devices.deviceFilter'))))) }
        'CA-30' { return (($suite -or $resources -contains '00000003-0000-0ff1-ce00-000000000000') -and (Get-CaProperty $Policy 'sessionControls.applicationEnforcedRestrictions.isEnabled') -eq $true) }
        'CA-31' { return ($suite -and ($builtIn -contains 'compliantApplication' -or $builtIn -contains 'approvedApplication')) }
      }
    }
  }
  return $false
}

function Get-CaPolicyReviewReason {
  <#
  .SYNOPSIS
    Explain scope gaps and grant bypasses instead of treating a matching setting as full coverage.
  #>
  param($Policy, $Recommendation, [string[]]$EmergencyAccessAccountId)
  $id = $Recommendation.Control
  $conditions = Get-CaProperty $Policy 'conditions'
  $users = Get-CaProperty $conditions 'users'
  $apps = Get-CaProperty $conditions 'applications'
  $includedUsers = @(Get-CaProperty $users 'includeUsers')
  $includedRoles = @(Get-CaProperty $users 'includeRoles')
  $includedApps = @(Get-CaProperty $apps 'includeApplications')
  $includedActions = @(Get-CaProperty $apps 'includeUserActions')
  $allUsers = $includedUsers -contains 'All'
  $adminRoles = @('62e90394-69f5-4237-9190-012177145e10','e8611ab8-c189-46e8-94e1-60213ab1f814','b1be1c3e-b65d-4f19-8427-f6fa0d97feb9','194ae4cb-b126-40b2-bd5b-6091b380977d','9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3','158c047a-c907-4556-b7ef-446551a6b5f7','c4e39bd9-1100-46d3-8c65-fb160da0071f','7be44c8a-adaf-4e2a-84d6-ab2649e08a13','729827e3-9c14-49f7-bb1b-9608f156bbb8','966707d0-3269-4727-9be2-8c3a10f19b9d','fe930be7-5e62-47db-91af-98c3a49a38b1','b0f54661-2d74-4c50-afa3-1ec803f12efe','29232cdf-9323-42fd-ade2-1d097af3e4de','f28a1f50-f6e7-4571-818b-6a12f2af6b6c')
  switch ($Recommendation.Users) {
    'All' { if (-not $allUsers) { 'All-user coverage is not established; group membership and combinations of targeted policies are not expanded.' } }
    'Admins' { if (-not $allUsers -and @($adminRoles | Where-Object { $_ -notin $includedRoles }).Count) { 'The full administrator role set is not directly covered. Group/PIM membership and combined policy coverage need review.' } }
    'Guests' {
      $guestAssignment = Get-CaProperty $users 'includeGuestsOrExternalUsers'
      $guestTypes = @((Get-CaProperty $guestAssignment 'guestOrExternalUserTypes') -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
      $requiredGuestTypes = @('internalGuest', 'b2bCollaborationGuest', 'b2bCollaborationMember', 'b2bDirectConnectUser', 'otherExternalUser', 'serviceProvider')
      $allGuestTypes = -not @($requiredGuestTypes | Where-Object { $_ -notin $guestTypes }).Count -and
        -not @($guestTypes | Where-Object { $_ -notin $requiredGuestTypes }).Count
      $allExternalTenants = (Get-CaProperty $guestAssignment 'externalTenants.membershipKind') -eq 'all'
      if (-not $allUsers -and $includedUsers -notcontains 'GuestsOrExternalUsers' -and -not ($allGuestTypes -and $allExternalTenants)) {
        'Full guest/external-user coverage is not established. Review guest types, external tenants and cross-tenant MFA trust.'
      }
    }
  }
  if (@(Get-CaProperty $users 'excludeGroups').Count -or @(Get-CaProperty $users 'excludeRoles').Count -or (Test-CaConfiguredValue (Get-CaProperty $users 'excludeGuestsOrExternalUsers'))) { 'Group, role or guest exclusions require membership and population review.' }
  $extraUsers = @(Get-CaProperty $users 'excludeUsers' | Where-Object { $_ -and $_ -notin $EmergencyAccessAccountId })
  if ($extraUsers.Count) { 'User exclusions are not identified as designated emergency accounts; validate the coverage gap.' }
  if ($Recommendation.Users -in @('All','Admins','Guests') -and (Test-CaConfiguredValue (Get-CaProperty $conditions 'clientApplications'))) { 'A workload assignment is not proof of user protection.' }
  $resourceCovered = switch ($Recommendation.Resources) {
    'All' { $includedApps -contains 'All' }
    'AdminPortals' { $includedApps -contains 'All' -or $includedApps -contains 'MicrosoftAdminPortals' }
    'AzureManagement' { $includedApps -contains 'All' -or $includedApps -contains '797f4846-ba00-4fd7-ba43-dac1f8f63013' }
    'Registration' { $includedActions -contains 'urn:user:registersecurityinfo' }
    'DeviceRegistration' { $includedActions -contains 'urn:user:registerdevice' }
    'Microsoft365' { $includedApps -contains 'All' -or $includedApps -contains 'Office365' }
    'SharePoint' { $includedApps -contains 'All' -or $includedApps -contains 'Office365' -or $includedApps -contains '00000003-0000-0ff1-ce00-000000000000' }
    'SelectedApps' { @($includedApps | Where-Object { $_ -and $_ -notin @('All','None','Office365') }).Count -gt 0 }
    default { $true }
  }
  if (-not $resourceCovered) { 'The required target resources or user action are not covered by this policy.' }
  if (@(Get-CaProperty $apps 'excludeApplications').Count -or (Test-CaConfiguredValue (Get-CaProperty $apps 'applicationFilter'))) { 'Resource exclusions or application filters require coverage review.' }
  if ($Recommendation.Resources -notin @('Registration','DeviceRegistration') -and (@(Get-CaProperty $apps 'includeUserActions').Count -or @(Get-CaProperty $apps 'includeAuthenticationContextClassReferences').Count)) { 'A user-action or authentication-context policy does not prove broad resource coverage.' }
  $allowed = @('users','applications','clientAppTypes') + @($Recommendation.AllowedConditions)
  $conditionNames = if ($conditions -is [System.Collections.IDictionary]) { @($conditions.Keys) } else { @($conditions.PSObject.Properties.Name) }
  $additional = $conditions.AdditionalProperties
  if ($additional) {
    $conditionNames += if ($additional -is [System.Collections.IDictionary]) { @($additional.Keys) } else { @($additional.PSObject.Properties.Name) }
  }
  foreach ($name in ($conditionNames | Sort-Object -Unique)) {
    if ($name -in $allowed -or $name -in @('AdditionalProperties','@odata.type')) { continue }
    $value = Get-CaProperty $conditions $name
    if (Test-CaConfiguredValue $value) {
      if ($name -eq 'platforms' -and @(Get-CaProperty $value 'includePlatforms').Count -eq 1 -and @(Get-CaProperty $value 'includePlatforms') -contains 'all' -and -not @(Get-CaProperty $value 'excludePlatforms').Count) { continue }
      if ($name -eq 'locations' -and @(Get-CaProperty $value 'includeLocations').Count -eq 1 -and @(Get-CaProperty $value 'includeLocations') -contains 'All' -and -not @(Get-CaProperty $value 'excludeLocations').Count) { continue }
      "Additional condition '$name' limits coverage and needs review."
    }
  }
  $clients = @(Get-CaProperty $conditions 'clientAppTypes')
  if ($id -eq 'CA-00') {
    if ('exchangeActiveSync' -notin $clients -or 'other' -notin $clients) { 'Both Exchange ActiveSync and other legacy clients must be blocked.' }
  }
  elseif ($id -eq 'CA-30') {
    if ($clients.Count -and $clients -notcontains 'all' -and $clients -notcontains 'browser') { 'Browser access is not targeted by the application-enforced restrictions policy.' }
  }
  elseif ($id -eq 'CA-31') {
    if ($clients.Count -and $clients -notcontains 'all' -and $clients -notcontains 'mobileAppsAndDesktopClients') { 'Mobile apps and desktop clients are not targeted by this app-protection policy.' }
  }
  elseif ($id -ne 'CA-21' -and $clients.Count -and $clients -notcontains 'all') { 'Client-app targeting does not cover all client types.' }
  if ($Recommendation.Platforms.Count) {
    $platforms = @(Get-CaProperty $conditions 'platforms.includePlatforms')
    if ($platforms.Count -and $platforms -notcontains 'all' -and @($Recommendation.Platforms | Where-Object { $_ -notin $platforms }).Count) { 'The expected platform set is not fully covered by this policy.' }
    if (@(Get-CaProperty $conditions 'platforms.excludePlatforms').Count) { 'Platform exclusions require coverage review.' }
  }
  $mfaChecks = @('CA-01','CA-04','CA-09','CA-11','CA-13','CA-16','CA-17','CA-18')
  if ($id -in $mfaChecks -and -not (Test-CaRequiredGrant $Policy @('mfa') -AllowMfaStrength)) { 'MFA is not required on every grant path, or authentication strength metadata is unavailable.' }
  if ($id -eq 'CA-05' -and -not (Test-CaRequiredGrant $Policy @() -PhishingResistant)) { 'Phishing-resistant authentication is not required on every grant path, or strength metadata is unavailable.' }
  if ($id -eq 'CA-02' -and -not (Test-CaRequiredGrant $Policy @('compliantDevice','compliantApplication'))) { 'A modern compliant-device or app-protection control is not required on every grant path.' }
  if ($id -eq 'CA-03' -and -not (Test-CaRequiredGrant $Policy @('compliantDevice'))) { 'Hybrid join alone does not protect macOS; verify compliance coverage for both operating systems.' }
  if ($id -eq 'CA-15' -and -not (Test-CaRequiredGrant $Policy @('compliantDevice','domainJoinedDevice'))) { 'Managed-device assurance can be bypassed through another grant option.' }
  if ($id -in @('CA-28','CA-29')) {
    if (-not (Get-CaDeviceRestrictionPattern $Policy)) { 'No recognized mandatory device restriction: review grant alternatives and device-filter mode/rule. Managed/registered alone is not compliant; isCompliant -eq False omits null/unregistered devices. Complex filters are not evaluated.' }
    if (@(Get-CaProperty $conditions 'devices.includeDevices').Count -or @(Get-CaProperty $conditions 'devices.excludeDevices').Count) { 'Additional legacy device assignments can narrow this restriction; review effective device scope.' }
  }
  if ($id -eq 'CA-31' -and -not (Test-CaRequiredGrant $Policy @('compliantApplication'))) { 'App protection is not required on every grant path; legacy approved applications or another OR alternative can bypass this template.' }
  if ($id -eq 'CA-31') {
    $mobilePlatforms = @(Get-CaProperty $conditions 'platforms.includePlatforms')
    if (-not $mobilePlatforms.Count -or @($mobilePlatforms | Where-Object { $_ -notin @('android','iOS') }).Count) { 'Restrict this mobile app-protection template to Android and iOS; broader platform scope can block unsupported clients.' }
  }
  if ($id -in @('CA-30','CA-31') -and @(Get-CaProperty $Policy 'grantControls.builtInControls') -contains 'block') { 'A block grant prevents the intended limited-access/app-protection experience.' }
  if ($id -eq 'CA-12' -and @(Get-CaProperty $conditions 'platforms.excludePlatforms') -contains 'all') { 'Excluding every platform makes the platform-blocking policy ineffective.' }
  if ($id -eq 'CA-12' -and @(Get-CaProperty $conditions 'platforms.excludePlatforms' | Where-Object { $_ -notin @('windows','macOS','android','iOS','linux','windowsPhone') }).Count) { 'The supported-platform exclusion list contains an unknown or overly broad value.' }
  if (@(Get-CaProperty $Policy 'grantControls.builtInControls') -contains 'unknownFutureValue') { 'An unknown grant control was returned by the API/SDK; update the schema or review the original policy.' }
  if ($id -in @('CA-09','CA-14')) {
    $frequency = Get-CaProperty $Policy 'sessionControls.signInFrequency'
    if ($frequency.isEnabled -ne $true -or $frequency.frequencyInterval -ne 'everyTime') { 'Enabled every-time sign-in frequency is not configured.' }
  }
  if ($id -eq 'CA-09') {
    $risk = @(Get-CaProperty $conditions 'signInRiskLevels')
    if ('medium' -notin $risk -or 'high' -notin $risk) { 'Both medium and high sign-in risk must be covered by the recommended baseline.' }
  }
  if ($id -eq 'CA-14') {
    $grant = Get-CaProperty $Policy 'grantControls'
    if (@($grant.builtInControls) -notcontains 'riskRemediation' -or @($grant.builtInControls) -contains 'passwordChange' -or $grant.operator -ne 'AND' -or -not (Test-CaMfaStrength $Policy)) { 'Modern remediation requires riskRemediation AND an MFA-capable strength; a legacy password-change policy needs passwordless/SSPR/writeback review.' }
    if (@(Get-CaProperty $conditions 'userRiskLevels') -notcontains 'high') { 'High user risk is not covered.' }
  }
}

function Get-CaRecommendationResult {
  <#
  .SYNOPSIS
    Assess recommendation configuration without confusing policy presence, enforcement, and verified coverage.
  #>
  param([AllowEmptyCollection()][object[]]$Policies, [string[]]$EmergencyAccessAccountId, [switch]$PartialExport)
  foreach ($definition in (Get-CaRecommendationCatalog)) {
    $evidence = @()
    foreach ($policy in $Policies) {
      if ($definition.Kind -eq 'Emergency') {
        if (-not $EmergencyAccessAccountId.Count) { continue }
        $users = Get-CaProperty $policy 'conditions.users'
        if (-not $users) { continue }
        $includedUsers = @(Get-CaProperty $users 'includeUsers')
        $broadAssignment = 'All' -in $includedUsers -or @(Get-CaProperty $users 'includeGroups').Count -or @(Get-CaProperty $users 'includeRoles').Count
        $targetedAccounts = @($EmergencyAccessAccountId | Where-Object { $broadAssignment -or $_ -in $includedUsers })
        $excludedUsers = @(Get-CaProperty $users 'excludeUsers')
        $missing = @($targetedAccounts | Where-Object { $_ -notin $excludedUsers })
        if (-not $missing.Count) { continue }
        $issues = @('Potentially targeted designated emergency accounts lack direct exclusions. Verify actual targeting and group-based exclusions before changing policy.')
      }
      else {
        if ($definition.Resources -and $definition.Resources -notin @('Registration','DeviceRegistration') -and
            (@(Get-CaProperty $policy 'conditions.applications.includeUserActions').Count -or @(Get-CaProperty $policy 'conditions.applications.includeAuthenticationContextClassReferences').Count)) { continue }
        if (-not (Test-CaRecommendationSignal $policy $definition.Control)) { continue }
        $issues = if ($definition.Kind -eq 'Hygiene') { @('Review this configuration in its business context; it is not automatically an exploitable gap.') } else { @(Get-CaPolicyReviewReason $policy $definition $EmergencyAccessAccountId) }
      }
      $evidence += [pscustomobject]@{
        PolicyId = [string](Get-CaProperty $policy 'id'); Name = [string](Get-CaProperty $policy 'displayName')
        State = [string](Get-CaProperty $policy 'state'); Issues = @($issues)
        Resources = @(Get-CaProperty $policy 'conditions.applications.includeApplications')
        ExcludedResources = @(Get-CaProperty $policy 'conditions.applications.excludeApplications')
        DeviceFilter = Get-CaProperty $policy 'conditions.devices.deviceFilter'
        DevicePattern = $(if ($definition.Control -in @('CA-28','CA-29')) { Get-CaDeviceRestrictionPattern $policy } else { '' })
      }
    }
    $qualified = @($evidence | Where-Object { $_.Issues.Count -eq 0 })
    $assessment = 'Needs review'
    $templateAssessment = ''
    $message = 'Configuration alone cannot establish applicability, licensing, supporting service setup or actual sign-in outcomes. Review the prerequisites and evidence.'
    if ($definition.Kind -eq 'Protection') {
      if (@($qualified | Where-Object State -eq 'enabled').Count) { $assessment = 'Enforced' }
      elseif (@($qualified | Where-Object State -eq 'enabledForReportingButNotEnforced').Count) { $assessment = 'Report-only' }
      elseif (@($qualified | Where-Object State -eq 'disabled').Count) { $assessment = 'Disabled' }
      elseif (-not $evidence.Count) { $assessment = 'Missing' }
      $message = switch ($assessment) {
        'Enforced' { 'An enabled policy meets the checked configuration and direct-scope criteria. This is not a sign-in simulation or a certification of tenant-wide effective access.' }
        'Report-only' { 'A qualifying policy is present only in report-only mode; it does not enforce this protection.' }
        'Disabled' { 'A qualifying policy is disabled; it does not enforce this protection.' }
        'Missing' { 'No matching protection was found in the exported policies. Confirm applicability and prerequisites before deployment.' }
        default { 'Relevant settings exist, but scope, exclusions, grant alternatives or missing metadata prevent confirming the recommended configuration.' }
      }
    }
    elseif ($definition.Kind -eq 'Hygiene' -and -not $evidence.Count) { $assessment = 'No findings'; $message = 'No matching configuration findings in this snapshot. Indirect membership and effective access are not evaluated.' }
    elseif ($definition.Kind -eq 'Emergency') {
      $message = if (-not $EmergencyAccessAccountId.Count) { 'Supply -EmergencyAccessAccountId to assess designated accounts. Account identities, readiness and exclusions are not guessed.' }
      elseif ($evidence.Count) { 'Some potentially targeted designated accounts lack direct exclusions. Check the evidence, membership, and deployment state; do not blindly add exclusions.' }
      else { 'No direct-exclusion gaps were found for the designated accounts on potentially applicable policies. Independently validate account count, readiness, monitoring, indirect targeting and recovery tests.' }
    }
    elseif ($definition.Kind -eq 'Template') {
      $templateAssessment = if (@($qualified | Where-Object State -eq 'enabled').Count) { 'Matching policy enabled' }
      elseif (@($qualified | Where-Object State -eq 'enabledForReportingButNotEnforced').Count) { 'Matching policy report-only' }
      elseif (@($qualified | Where-Object State -eq 'disabled').Count) { 'Matching policy disabled' }
      elseif ($evidence.Count) { 'Partial match' }
      else { 'Not observed' }
      $message = "Optional template: $templateAssessment. This measures the documented policy pattern, not business applicability or effective access. Review actual resource assignments, prerequisites and supporting service configuration; absence is not a universal failure."
    }
    if ($PartialExport) { $assessment = 'Needs review'; $message = 'This is a single-policy export, not a tenant-wide assessment. Missing policies and combined coverage cannot be determined. ' + $message }
    [pscustomobject]@{
      Control = $definition.Control; Name = $definition.Name; Assessment = $assessment
      Message = $message; Requirements = $definition.Requirements; Background = $definition.Background
      Verification = $definition.Verification; Links = $definition.Links; Evidence = @($evidence)
      Kind = $definition.Kind; TemplateSettings = $definition.TemplateSettings; TemplateAssessment = $templateAssessment
      NeedsAttention = ($definition.Kind -eq 'Protection' -and $assessment -ne 'Enforced') -or ($definition.Kind -eq 'Hygiene' -and $evidence.Count -gt 0) -or ($definition.Kind -eq 'Emergency' -and $EmergencyAccessAccountId.Count -gt 0 -and $evidence.Count -gt 0)
    }
  }
}

# Backward compatibility aliases (deprecated names). Retained temporarily so external callers
# referencing prior function names do not break. Marked for removal in a future major version.
Set-Alias -Name Ensure-GraphConnection -Value Connect-GraphContext -ErrorAction SilentlyContinue
Set-Alias -Name Safe-Get -Value Invoke-SafeGet -ErrorAction SilentlyContinue
Set-Alias -Name Translate-List -Value Convert-IdListToName -ErrorAction SilentlyContinue

# Determine offline mode early so we can skip any Graph interaction
$script:OfflineMode = [bool]$RawInputFile
$script:AuthStrengthCache = @{}
$script:AuthStrengthCachePopulated = $false
$OfflineSnapshot = $null
$IsPartialSnapshot = [bool]$PolicyID
if ($RawInputFile -and $TenantID) { throw 'TenantID cannot be combined with RawInputFile; offline reports use the captured tenant.' }
if (-not $script:OfflineMode) {
  # Ensure required modules are present before attempting to connect
  $null = Initialize-GraphModule
  Connect-GraphContext -TenantID $TenantID
}
else {
  Write-Info 'Offline mode detected: skipping module initialization and Graph connection.'
}

# Script metadata / version stamp (bump when feature changes)
$Script:CAExportVersion = '4.0'

# Map external switch parameters to internal export control flags for backward compatibility / clarity
if ($PSBoundParameters.ContainsKey('Html')) {
  $HTMLExport = [bool]$Html
}
if ($PSBoundParameters.ContainsKey('Json')) {
  $JsonExport = [bool]$Json
}
if ($PSBoundParameters.ContainsKey('Csv')) {
  $CsvExport = [bool]$Csv
}
if ($PSBoundParameters.ContainsKey('CsvPivot')) {
  $CsvPivotExport = [bool]$CsvPivot
}
# Default behavior: if no explicit export switches supplied, emit HTML + JSON + CSV (pivot remains opt-in)
if (-not ($HTMLExport -or $JsonExport -or $CsvExport -or $CsvPivotExport)) {
  $HTMLExport = $true; $JsonExport = $true; $CsvExport = $true
}

if ($OutputPath) {
  try {
    if (-not (Test-Path -LiteralPath $OutputPath)) {
      New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null 
    }
    $ExportLocation = (Resolve-Path -LiteralPath $OutputPath).Path
    Write-Info "Using custom output path: $ExportLocation"
  }
  catch {
    Write-Warn ("Failed to set custom OutputPath '{0}': {1}" -f $OutputPath, $_.Exception.Message)
  }
}

# If no -OutputPath was supplied or resolution failed, default to current working directory
if (-not $ExportLocation) {
  $ExportLocation = (Get-Location).Path
  Write-Info "No OutputPath specified – defaulting to current directory: $ExportLocation"
}

# ---------------- Retrieve Tenant & Policies ----------------
if ($RawInputFile) {
  Write-Info "Offline mode: importing snapshot or raw policy JSON from $RawInputFile"
  $OfflineSnapshot = Import-CaSnapshot -Path $RawInputFile
  $allPolicies = @($OfflineSnapshot.Policies)
  $TenantName = $OfflineSnapshot.TenantName
  $Date = $OfflineSnapshot.CapturedAt
  $IsPartialSnapshot = $IsPartialSnapshot -or $OfflineSnapshot.PartialExport
  if (-not $PSBoundParameters.ContainsKey('EmergencyAccessAccountId')) {
    $EmergencyAccessAccountId = @($OfflineSnapshot.EmergencyAccessAccountIds | Where-Object { $_ })
  }
  foreach ($strength in $OfflineSnapshot.AuthenticationStrengths) {
    if ($strength.Id) { $script:AuthStrengthCache[$strength.Id] = $strength }
    if ($strength.DisplayName) { $script:AuthStrengthCache[$strength.DisplayName] = $strength }
  }
  $script:AuthStrengthCachePopulated = $true
  if ($PolicyID) {
    $CAPolicy = @($allPolicies | Where-Object { $_.id -eq $PolicyID })
  }
  else {
    $CAPolicy = @($allPolicies)
  }
  if (-not $CAPolicy.Count) { throw "Policy with ID '$PolicyID' was not found in the offline input." }
  Write-Info "Loaded $($CAPolicy.Count) policies from raw file"
}
else {
  Write-Info 'Retrieving tenant information'
  $DirectoryModuleAvailable = Initialize-GraphModule -RequiredModules 'Microsoft.Graph.Identity.DirectoryManagement' -Optional
  $TenantName = if ($DirectoryModuleAvailable) {
    Invoke-SafeGet -FailureMessage 'Could not retrieve tenant display name; using UnknownTenant' {
      Get-MgOrganization -ErrorAction Stop | Select-Object -First 1 -ExpandProperty DisplayName
    }
  }
  if (-not $TenantName) { $TenantName = 'UnknownTenant' }
  $Date = (Get-Date).ToUniversalTime().ToString('u')
  Write-Info 'Retrieving Conditional Access policies'
  try {
    $allPolicies = Get-MgIdentityConditionalAccessPolicy -All -Headers @{ Prefer = 'include-unknown-enum-members' } -ErrorAction Stop
  }
  catch {
    Write-Err ('Failed to retrieve policies: {0}' -f $_.Exception.Message)
    Write-Err 'Cannot continue without policies. Exiting.'
    exit 1
  }
  if ($PolicyID) {
    $CAPolicy = $allPolicies | Where-Object { $_.id -eq $PolicyID }
    if (-not $CAPolicy) {
      Write-Err "Policy with ID '$PolicyID' not found. Exiting."
      exit 1
    }
  }
  else {
    $CAPolicy = $allPolicies
  }
  if (-not $CAPolicy -or $CAPolicy.Count -eq 0) {
    Write-Err 'No Conditional Access policies found in tenant. Exiting.'
    exit 1
  }
  Write-Info "Successfully retrieved $($CAPolicy.Count) Conditional Access $(if($CAPolicy.Count -eq 1){'policy'}else{'policies'})"
}

# Raw snapshot & index
$RawPolicyIndex = @{}
foreach ($rp in $CAPolicy) {
  if ($rp.id) {
    $RawPolicyIndex[$rp.id] = $rp 
  } 
}

# ---------------- Collect IDs for enrichment ----------------
$userIds = [System.Collections.Generic.HashSet[string]]::new()
$groupIds = [System.Collections.Generic.HashSet[string]]::new()
$roleIds = [System.Collections.Generic.HashSet[string]]::new()
$appIds = [System.Collections.Generic.HashSet[string]]::new()
$locIds = [System.Collections.Generic.HashSet[string]]::new()
$touIds = [System.Collections.Generic.HashSet[string]]::new()

# Collect IDs for enrichment (consolidated for clarity)
foreach ($p in $CAPolicy) {
  $c = $p.conditions
  if ($c.users) {
    # Users (excluding sentinels)
    foreach ($i in @($c.users.includeUsers)) {
      if ($i -and $i -notin @('All', 'None', 'GuestsOrExternalUsers')) { [void]$userIds.Add($i) }
    }
    foreach ($i in @($c.users.excludeUsers)) { if ($i) { [void]$userIds.Add($i) } }
    # Groups
    foreach ($i in @($c.users.includeGroups)) { if ($i) { [void]$groupIds.Add($i) } }
    foreach ($i in @($c.users.excludeGroups)) { if ($i) { [void]$groupIds.Add($i) } }
    # Roles
    foreach ($i in @($c.users.includeRoles)) { if ($i) { [void]$roleIds.Add($i) } }
    foreach ($i in @($c.users.excludeRoles)) { if ($i) { [void]$roleIds.Add($i) } }
  }
  if ($c.applications) {
    foreach ($i in @($c.applications.includeApplications)) { if ($i -and (Test-IsGuid $i)) { [void]$appIds.Add($i) } }
    foreach ($i in @($c.applications.excludeApplications)) { if ($i -and (Test-IsGuid $i)) { [void]$appIds.Add($i) } }
  }
  if ($c.locations) {
    foreach ($i in @($c.locations.includeLocations)) { if ($i -and (Test-IsGuid $i)) { [void]$locIds.Add($i) } }
    foreach ($i in @($c.locations.excludeLocations)) { if ($i -and (Test-IsGuid $i)) { [void]$locIds.Add($i) } }
  }
  if ($p.grantControls -and $p.grantControls.termsOfUse) {
    foreach ($i in @($p.grantControls.termsOfUse)) { if ($i -and (Test-IsGuid $i)) { [void]$touIds.Add($i) } }
  }
}

# ---------------- Build lookup maps (best-effort) ----------------
$UserMap = @{}; $GroupMap = @{}; $RoleMap = @{}; $AppMap = @{}; $LocMap = @{}; $TouMap = @{}
$NamedLocationInventory = if ($script:OfflineMode -and $OfflineSnapshot.NamedLocations) {
  $OfflineSnapshot.NamedLocations
}
else { Get-CaNamedLocationInventory -Offline:$script:OfflineMode }

if ($script:OfflineMode) {
  Write-Info 'Offline mode: restoring saved lookup maps without Graph calls; unresolved IDs remain unchanged.'
  foreach ($entry in @{
    Users = $UserMap; Groups = $GroupMap; Roles = $RoleMap
    Applications = $AppMap; Locations = $LocMap; TermsOfUse = $TouMap
  }.GetEnumerator()) {
    $savedMap = $OfflineSnapshot.Lookups.($entry.Key)
    if ($savedMap) {
      foreach ($property in $savedMap.PSObject.Properties) { $entry.Value[$property.Name] = $property.Value }
    }
  }
}
else {
  $usersAvailable = if ($userIds.Count) { Initialize-GraphModule -RequiredModules 'Microsoft.Graph.Users' -Optional } else { $false }
  $groupsAvailable = if ($groupIds.Count) { Initialize-GraphModule -RequiredModules 'Microsoft.Graph.Groups' -Optional } else { $false }
  $appsAvailable = if ($appIds.Count) { Initialize-GraphModule -RequiredModules 'Microsoft.Graph.Applications' -Optional } else { $false }
  $governanceAvailable = if ($roleIds.Count -or $touIds.Count) { Initialize-GraphModule -RequiredModules 'Microsoft.Graph.Identity.Governance' -Optional } else { $false }
  foreach ($id in $(if ($usersAvailable) { $userIds })) {
    if ($id -and ($id -match '^[0-9a-fA-F\-]{36}$')) {
      $obj = Invoke-SafeGet -FailureMessage "Could not resolve user '$id'; keeping its ID" {
        $user = Get-MgUser -UserId $id -Property Id, DisplayName -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($user.DisplayName)) { throw 'Graph returned no display name. Check Directory.Read.All consent and whether the user still exists.' }
        $user
      }
      if ($obj) { $UserMap[$id] = $obj.DisplayName }
    }
  }
  foreach ($id in $(if ($groupsAvailable) { $groupIds })) {
    $obj = Invoke-SafeGet -FailureMessage "Could not resolve group '$id'; keeping its ID" {
      $group = Get-MgGroup -GroupId $id -Property Id, DisplayName -ErrorAction Stop
      if ([string]::IsNullOrWhiteSpace($group.DisplayName)) { throw 'Graph returned no display name. Check Directory.Read.All consent and whether the group still exists.' }
      $group
    }
    if ($obj) { $GroupMap[$id] = $obj.DisplayName }
  }

  $roleLookup = @{}
  $commonRoleTemplates = @{
    '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' = 'Application Administrator'
    'c4e39bd9-1100-46d3-8c65-fb160da0071f' = 'Authentication Administrator'
    'e3973bdf-4987-49ae-837a-ba8e231c7286' = 'Azure DevOps Administrator'
    'b0f54661-2d74-4c50-afa3-1ec803f12efe' = 'Billing Administrator'
    '158c047a-c907-4556-b7ef-446551a6b5f7' = 'Cloud Application Administrator'
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' = 'Conditional Access Administrator'
    '29232cdf-9323-42fd-ade2-1d097af3e4de' = 'Exchange Administrator'
    '62e90394-69f5-4237-9190-012177145e10' = 'Global Administrator'
    '729827e3-9c14-49f7-bb1b-9608f156bbb8' = 'Helpdesk Administrator'
    '966707d0-3269-4727-9be2-8c3a10f19b9d' = 'Password Administrator'
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' = 'Privileged Authentication Administrator'
    'e8611ab8-c189-46e8-94e1-60213ab1f814' = 'Privileged Role Administrator'
    '194ae4cb-b126-40b2-bd5b-6091b380977d' = 'Security Administrator'
    '5d6b6bb7-de71-4623-b4af-96380a352509' = 'Security Reader'
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' = 'SharePoint Administrator'
    'fe930be7-5e62-47db-91af-98c3a49a38b1' = 'User Administrator'
  }

  $dirRoles = if ($roleIds.Count -and $DirectoryModuleAvailable) {
    Invoke-SafeGet { Get-MgDirectoryRole -All -ErrorAction Stop }
  }
  if ($dirRoles) {
    foreach ($r in $dirRoles) {
      if ($r.id -and -not $roleLookup.ContainsKey($r.id)) { $roleLookup[$r.id] = $r.displayName }
      if ($r.roleTemplateId -and -not $roleLookup.ContainsKey($r.roleTemplateId)) { $roleLookup[$r.roleTemplateId] = $r.displayName }
    }
  }
  else {
    foreach ($kv in $commonRoleTemplates.GetEnumerator()) { $roleLookup[$kv.Key] = $kv.Value }
  }

  $roleDefs = if ($roleIds.Count -and $governanceAvailable) {
    Invoke-SafeGet { Get-MgRoleManagementDirectoryRoleDefinition -All -Property Id, DisplayName, TemplateId -ErrorAction Stop }
  }
  if ($roleDefs) {
    foreach ($rd in $roleDefs) {
      if ($rd.Id -and -not $roleLookup.ContainsKey($rd.Id)) { $roleLookup[$rd.Id] = $rd.DisplayName }
      if ($rd.TemplateId -and -not $roleLookup.ContainsKey($rd.TemplateId)) { $roleLookup[$rd.TemplateId] = $rd.DisplayName }
    }
  }
  else {
    foreach ($kv in $commonRoleTemplates.GetEnumerator()) { if (-not $roleLookup.ContainsKey($kv.Key)) { $roleLookup[$kv.Key] = $kv.Value } }
  }

  foreach ($id in $roleIds) {
    if ($roleLookup.ContainsKey($id)) {
      $RoleMap[$id] = $roleLookup[$id]
    }
    else {
      $obj = if ($DirectoryModuleAvailable) { Invoke-SafeGet { Get-MgDirectoryRole -DirectoryRoleId $id -Property Id, DisplayName -ErrorAction Stop } }
      if ($obj) { $RoleMap[$id] = $obj.DisplayName } elseif ($commonRoleTemplates.ContainsKey($id)) { $RoleMap[$id] = $commonRoleTemplates[$id] }
    }
  }
  foreach ($id in $(if ($appsAvailable) { $appIds })) {
    if (Test-IsGuid $id) {
      $obj = Invoke-SafeGet { Get-MgServicePrincipal -ServicePrincipalId $id -Property Id, DisplayName, AppId }
      if ($obj) { $AppMap[$id] = $obj.DisplayName }
    }
  }
  if ($NamedLocationInventory -and $NamedLocationInventory.Status -eq 'Available') {
    foreach ($location in $NamedLocationInventory.Locations) {
      if ($location.Id -and -not [string]::IsNullOrWhiteSpace($location.DisplayName)) {
        $LocMap[$location.Id] = $location.DisplayName
      }
    }
  }
  else {
    foreach ($id in $locIds) {
      if (Test-IsGuid $id) {
        $obj = Invoke-SafeGet -FailureMessage "Could not resolve named location '$id'; keeping its ID" {
          Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $id -Property Id, DisplayName -ErrorAction Stop
        }
        if ($obj -and -not [string]::IsNullOrWhiteSpace($obj.DisplayName)) { $LocMap[$id] = $obj.DisplayName }
      }
    }
  }
  foreach ($id in $(if ($governanceAvailable) { $touIds })) {
    if (Test-IsGuid $id) {
      $obj = Invoke-SafeGet { Get-MgAgreement -AgreementId $id -Property Id, DisplayName }
      if ($obj) { $TouMap[$id] = $obj.DisplayName }
    }
  }
}

# ---------------- Save portable snapshot ----------------
# Save collection data before rendering so it remains reusable even if a report export fails.
$tenantFileName = ($TenantName -replace '[<>:"/\\|?*\x00-\x1f]', '_').Trim().TrimEnd('.')
if ([string]::IsNullOrWhiteSpace($tenantFileName)) { $tenantFileName = 'UnknownTenant' }
$baseName = "CAExportRecs_${tenantFileName}_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
if (-not $script:OfflineMode) {
  Initialize-AuthStrengthCache
  $snapshot = [ordered]@{
    SnapshotFormat = 'CAExportSnapshot'; SchemaVersion = 1; ExporterVersion = $CAExportVersion
    TenantName = $TenantName; CapturedAt = $Date; PartialExport = $IsPartialSnapshot
    Policies = @($CAPolicy); NamedLocations = $NamedLocationInventory
    Lookups = [ordered]@{
      Users = $UserMap; Groups = $GroupMap; Roles = $RoleMap
      Applications = $AppMap; Locations = $LocMap; TermsOfUse = $TouMap
    }
    AuthenticationStrengths = @($script:AuthStrengthCache.Values | Sort-Object Id -Unique)
    EmergencyAccessAccountIds = @($EmergencyAccessAccountId | Where-Object { $_ })
  }
  $snapshotFile = Join-Path -Path $ExportLocation -ChildPath "${baseName}_raw.json"
  ConvertTo-Json -InputObject $snapshot -Depth 100 -WarningAction Stop |
    Set-Content -LiteralPath $snapshotFile -Encoding utf8 -ErrorAction Stop
  Write-Info "Portable offline snapshot saved: $snapshotFile"
}

# ---------------- Construct CAExport ----------------
$AssignmentLookups = @{ users = $UserMap; groups = $GroupMap; roles = $RoleMap; applications = $AppMap; locations = $LocMap }
$NamingHintIndex = @{}
$AssignmentTokenIndex = @{}
$CAExport = @()
foreach ($Policy in $CAPolicy) {
  $NamingHintIndex[$Policy.id] = @(Get-CaNamingHint -Policy $Policy)
  $AssignmentTokenIndex[$Policy.id] = @{}
  foreach ($column in @('Included Users', 'Excluded Users', 'Included Applications', 'Excluded Applications', 'Included Locations', 'Excluded Locations')) {
    $AssignmentTokenIndex[$Policy.id][$column] = @(Get-CaAssignmentToken -Policy $Policy -Column $column -Lookups $AssignmentLookups)
  }
  $DateModified = if ($Policy.modifiedDateTime) {
    $Policy.modifiedDateTime 
  }
  else {
    $Policy.createdDateTime 
  }
  $InclPlat = $Policy.conditions.platforms.includePlatforms
  $ExclPlat = $Policy.conditions.platforms.excludePlatforms
  $InclDev = $Policy.conditions.devices.includeDevices
  $ExclDev = $Policy.conditions.devices.excludeDevices
  $devFilters = $Policy.conditions.devices.deviceFilter.rule
  $authenticationFlowsString = @(Get-CaProperty $Policy 'conditions.authenticationFlows.transferMethods') -join ', '
  $InclLocation = $Policy.conditions.locations.includeLocations | ForEach-Object { if ($_ -and (Test-IsGuid $_) -and $LocMap.ContainsKey($_)) {
      $LocMap[$_] 
    }
    else {
      $_ 
    } }
  $ExclLocation = $Policy.conditions.locations.excludeLocations | ForEach-Object { if ($_ -and (Test-IsGuid $_) -and $LocMap.ContainsKey($_)) {
      $LocMap[$_] 
    }
    else {
      $_ 
    } }
  $IncludeUG = @()
  $IncludeUG += (Convert-IdListToName $Policy.conditions.users.includeUsers $UserMap)
  $IncludeUG += (Convert-IdListToName $Policy.conditions.users.includeGroups $GroupMap)
  $IncludeUG += (Convert-IdListToName $Policy.conditions.users.includeRoles $RoleMap)
  if ($Policy.conditions.users.includeGuestsOrExternalUsers.guestOrExternalUserTypes) {
    $IncludeUG += $Policy.conditions.users.includeGuestsOrExternalUsers.guestOrExternalUserTypes 
  }
  $ExcludeUG = @()
  $ExcludeUG += (Convert-IdListToName $Policy.conditions.users.excludeUsers $UserMap)
  $ExcludeUG += (Convert-IdListToName $Policy.conditions.users.excludeGroups $GroupMap)
  $ExcludeUG += (Convert-IdListToName $Policy.conditions.users.excludeRoles $RoleMap)
  if ($Policy.conditions.users.excludeGuestsOrExternalUsers.guestOrExternalUserTypes) {
    $ExcludeUG += $Policy.conditions.users.excludeGuestsOrExternalUsers.guestOrExternalUserTypes 
  }

  $CAExport += [PSCustomObject][ordered]@{
    Name                                = $Policy.displayName
    'Recommended Name'                  = Get-RecommendedPolicyName -Policy $Policy
    'Naming Hints'                      = ($NamingHintIndex[$Policy.id] -join ' | ')
    PolicyId                            = $Policy.id
    Status                              = Format-PolicyStatus -Status $Policy.state
    Modified                            = $DateModified
    Created                             = $Policy.createdDateTime
    Description                         = $Policy.description
    'Included Users'                    = ($IncludeUG -join ", `r`n")
    'Excluded Users'                    = ($ExcludeUG -join ", `r`n")
    'Included Applications'             = ($Policy.conditions.applications.includeApplications -join ", `r`n")
    'Excluded Applications'             = ($Policy.conditions.applications.excludeApplications -join ", `r`n")
    'User Actions'                      = (($Policy.conditions.applications.includeUserActions | ForEach-Object { 
                                              if ($_ -eq 'urn:user:registersecurityinfo') { 'Register Security Info' }
                                              elseif ($_ -eq 'urn:user:registerdevice') { 'Register Device' }
                                              else { $_ }
                                            }) -join ", `r`n")
    'Auth Context'                      = ($Policy.conditions.applications.includeAuthenticationContextClassReferences -join ", `r`n")
    'User Risk'                         = ($Policy.conditions.userRiskLevels -join ", `r`n")
    'SignIn Risk'                       = ($Policy.conditions.signInRiskLevels -join ", `r`n")
    'Platforms Included'                = ($InclPlat -join ", `r`n")
    'Platforms Excluded'                = ($ExclPlat -join ", `r`n")
    'Included Locations'                = ($InclLocation -join ", `r`n")
    'Excluded Locations'                = ($ExclLocation -join ", `r`n")
    'Client Apps'                       = ($Policy.conditions.clientAppTypes -join ", `r`n")
    'Included Devices'                  = ($InclDev -join ", `r`n")
    'Excluded Devices'                  = ($ExclDev -join ", `r`n")
    'Device Filters'                    = ($devFilters -join ", `r`n")
    'Authentication Flows'              = $authenticationFlowsString
    'Block'                             = if (@(Get-CaProperty $Policy 'grantControls.builtInControls') -contains 'Block') {
      'True' 
    }
    else {
      '' 
    }
    'Require MFA'                       = if ($Policy.grantControls.builtInControls -contains 'Mfa') {
      'True' 
    }
    else {
      '' 
    }
    'Authentication Strength MFA'       = $Policy.grantControls.authenticationStrength.displayName
    'Compliant Device'                  = if ($Policy.grantControls.builtInControls -contains 'CompliantDevice') {
      'True' 
    }
    else {
      '' 
    }
    'Domain Joined Device'              = if ($Policy.grantControls.builtInControls -contains 'DomainJoinedDevice') {
      'True' 
    }
    else {
      '' 
    }
    'Compliant Application'             = if ($Policy.grantControls.builtInControls -contains 'CompliantApplication') {
      'True' 
    }
    else {
      '' 
    }
    'Approved Application'              = if ($Policy.grantControls.builtInControls -contains 'ApprovedApplication') {
      'True' 
    }
    else {
      '' 
    }
    'Password Change'                   = if ($Policy.grantControls.builtInControls -contains 'PasswordChange') {
      'True' 
    }
    else {
      '' 
    }
    'Terms Of Use'                      = ((Convert-IdListToName $Policy.grantControls.termsOfUse $TouMap) -join ", `r`n")
    'Custom Controls'                   = ($Policy.grantControls.customAuthenticationFactors -join ", `r`n")
    'Grant Operator'                    = $Policy.grantControls.operator
    'Application Enforced Restrictions' = $Policy.sessionControls.applicationEnforcedRestrictions.isEnabled
    'Cloud App Security'                = $Policy.sessionControls.cloudAppSecurity.isEnabled
    'Sign In Frequency'                 = if ($Policy.sessionControls.signInFrequency.value -and $Policy.sessionControls.signInFrequency.type) {
      "$( $Policy.sessionControls.signInFrequency.value ) $( $Policy.sessionControls.signInFrequency.type )" 
    }
    'Persistent Browser'                = $Policy.sessionControls.persistentBrowser.mode
    'Continuous Access Evaluation'      = $Policy.sessionControls.continuousAccessEvaluation.mode
    'Resilient Defaults'                = $Policy.sessionControls.disableResilienceDefaults
    'Secure Sign In Session'            = Get-CaProperty $Policy 'sessionControls.secureSignInSession.isEnabled'
    RawJson                             = ''
  }
}

$LinkURL = 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/'
$FileName = "$baseName.html"
$JsonFileName = "$baseName.json"
$CsvFileName = "$baseName.csv"
$CsvPivotFileName = "$baseName-pivot.csv"
$HtmlParts = @()
if (-not $NoRecommendations) {
  Write-Info 'Analyzing: getting recommendations'
  $recommendations = @(Get-CaRecommendationResult -Policies @($CAPolicy) -EmergencyAccessAccountId $EmergencyAccessAccountId -PartialExport:$IsPartialSnapshot)
}

## (Removed obsolete Set Row Order block)

# ---------------- Build Pivot Dataset (wide) ----------------
# The pivot format rotates policy-centric data so that each policy becomes a column and each attribute a row.
# This allows quick scanning in Excel / BI of where controls/conditions differ across policies.
# Implementation notes:
#  - Use a stable ordered list of pivot fields ($pivotFields)
#  - Each row object has a 'CA Item' property (row label) plus one property per policy (policy display name)
#  - Boolean-like values preserved (✔ for True style fields) while multi-line fields condensed to single line for readability

$pivot = @()
$pivotFields = @(
  #  'Recommended Name', 'Modified', 'Created', 'Description',
  #   'Included Users', 'Excluded Users',
  #   'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations',
  #   'User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps','Device Filters', 'Authentication Flows',
  #   'Block', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'Grant Operator',
  #   'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session'
  'Recommended Name', 'Naming Hints', 'Modified', 'Created', 'Description',
  'Included Users', 'Excluded Users',
  'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations', 'User Risk', 'SignIn Risk', <#'Insider Risk',#> 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows', 'Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session', 'RawJson'
)

if ($CAExport.Count -gt 0) {
  foreach ($field in $pivotFields) {
    $row = [ordered]@{ 'CA Item' = $field }
    foreach ($pol in $CAExport) {
      $val = $null
      if ($pol.PSObject.Properties.Match($field)) {
        $val = $pol.$field 
      }
      # Normalize multiline -> semi-colon separated single line for sheet friendliness
      if ($val -is [string]) {
        $val = ($val -split "`r?`n") -join '; ' 
      }
      $row[$pol.Name] = $val
    }
    $pivot += [pscustomobject]$row
  }
}
function Get-RecommendationsHtmlFragment {
  <#
  .SYNOPSIS
    Render escaped recommendation evidence, assessment states, prerequisites and Microsoft references.
  #>
  param (
    [Parameter(Mandatory = $true)]
    [object[]]$Recommendations
  )
  $htmlFragment = @'
<div class='recommendations' id='ca-security-checks' style=''>
<div class='page-heading'><p class='eyebrow'>Configuration review</p><h1>Recommendations &amp; optional templates</h1><p>Prioritize baseline gaps, then evaluate scenarios that match your business needs.</p></div>
<p class='rec-assessment-note'>Configuration assessment, not an effective-access simulation. Enforced means an enabled policy meets the checked direct-scope criteria; it does not verify licenses, group/PIM membership, combined policies, application setup or sign-in outcomes. Attention shows protection/hygiene findings and designated emergency-account gaps, not absent optional scenarios. Optional reviews and templates are separate; Needs review does not itself mean a policy is disabled or ineffective. Review exclusions and pilot changes in report-only mode. Guidance reviewed November 18, 2026.</p>
<div class='rec-toolbar'>
  <div class='rec-filter-group' role='group' aria-label='Recommendation view'>
    <button type='button' id='rec-filter-fail' class='util-button active' data-filter='fail' aria-pressed='true' aria-controls='rec-results'>Attention</button>
    <button type='button' id='rec-filter-all' class='util-button' data-filter='all' aria-pressed='false' aria-controls='rec-results'>All recommendations</button>
    <button type='button' id='rec-filter-templates' class='util-button' data-filter='templates' aria-pressed='false' aria-controls='rec-results'>Optional templates</button>
    <button type='button' id='rec-filter-optional' class='util-button' data-filter='optional' aria-pressed='false' aria-controls='rec-results'>Operational reviews</button>
  </div>
  <div class='rec-assessment-control'><label for='rec-assessment-filter'>Assessment</label><select id='rec-assessment-filter' aria-controls='rec-results'>
    <option value='all'>All assessments</option><option value='enforced'>Enforced</option><option value='report-only'>Report-only</option><option value='disabled'>Disabled</option><option value='missing'>Missing</option><option value='needs-review'>Needs review</option><option value='no-findings'>No findings</option>
  </select></div>
  <p id='rec-result-count' class='rec-filter-summary' role='status' aria-live='polite'></p>
</div>
<p id='rec-filter-empty' class='rec-filter-empty' hidden>No recommendations match this view and assessment. Choose All recommendations or All assessments to broaden the results.</p>
<div id='rec-results'>
'@
  foreach ($rec in $Recommendations) {
    $links = ''
    foreach ($key in $rec.Links.Keys) {
      $url = [System.Web.HttpUtility]::HtmlEncode([string]$rec.Links[$key])
      $label = [System.Web.HttpUtility]::HtmlEncode($key)
      $links += "<a class='rec-link' href='$url' target='_blank' rel='noopener noreferrer'>$label</a>"
    }
    $RecStatus = if ($rec.Assessment -eq 'Enforced') { 'pass' } elseif ($rec.Assessment -eq 'No findings') { 'clear' } else { 'fail' }
    $detailOpen = if ($RecStatus -eq 'pass' -or $RecStatus -eq 'clear') { '' } else { ' open' }
    $state = [System.Web.HttpUtility]::HtmlEncode($rec.Assessment)
    $stateKey = $rec.Assessment.ToLowerInvariant() -replace ' ', '-'
    $encName = [System.Web.HttpUtility]::HtmlEncode($rec.Name)
    $encCtrl = [System.Web.HttpUtility]::HtmlEncode($rec.Control)
    $background = [System.Web.HttpUtility]::HtmlEncode($rec.Background)
    $verification = [System.Web.HttpUtility]::HtmlEncode($rec.Verification)
    $message = [System.Web.HttpUtility]::HtmlEncode($rec.Message)
    $requirements = ($rec.Requirements | ForEach-Object { '<li>' + [System.Web.HttpUtility]::HtmlEncode($_) + '</li>' }) -join ''
    $templateHtml = ''
    if ($rec.TemplateSettings) {
      $settings = foreach ($key in $rec.TemplateSettings.Keys) {
        '<dt>' + [System.Web.HttpUtility]::HtmlEncode($key) + '</dt><dd>' + [System.Web.HttpUtility]::HtmlEncode($rec.TemplateSettings[$key]) + '</dd>'
      }
      $measurement = [System.Web.HttpUtility]::HtmlEncode($rec.TemplateAssessment)
      $templateHtml = "<section class='rec-template'><h3>Optional template <span class='template-measurement'>$measurement</span></h3><dl>$($settings -join '')</dl></section>"
    }
    $kind = [System.Web.HttpUtility]::HtmlEncode($rec.Kind.ToLowerInvariant())
    $category = switch ($rec.Kind) { 'Template' { 'Optional template' }; 'Review' { 'Optional review' }; 'Hygiene' { 'Configuration hygiene' }; 'Emergency' { 'Operational review' }; default { 'Protection check' } }
    $deploymentSummary = @(
      foreach ($deploymentState in @('enabled','enabledForReportingButNotEnforced','disabled')) {
        $count = @($rec.Evidence | Where-Object State -eq $deploymentState).Count
        if ($count) { "$count $(Format-PolicyStatus $deploymentState)" }
      }
    ) -join ', '
    if ($deploymentSummary) { $category += ' | Observed: ' + [System.Web.HttpUtility]::HtmlEncode($deploymentSummary) }
    $attention = $rec.NeedsAttention.ToString().ToLowerInvariant()
    $evidenceHtml = ''
    foreach ($item in $rec.Evidence) {
      $policyName = [System.Web.HttpUtility]::HtmlEncode($item.Name)
      $deployment = [System.Web.HttpUtility]::HtmlEncode((Format-PolicyStatus $item.State))
      if (Test-IsGuid $item.PolicyId) {
        $policyName = "<a href='https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/$($item.PolicyId)' target='_blank' rel='noopener noreferrer'>$policyName</a>"
      }
      $issues = ($item.Issues | ForEach-Object { '<li>' + [System.Web.HttpUtility]::HtmlEncode($_) + '</li>' }) -join ''
      if (-not $issues) { $issues = '<li>Meets the checked configuration and direct-scope criteria; validate effective access separately.</li>' }
      $scope = ''
      if ($item.Resources.Count) {
        $resourceNames = foreach ($resource in $item.Resources) {
          switch ($resource) {
            'All' { 'All resources' }
            'Office365' { 'Microsoft 365 (Office365)' }
            '00000003-0000-0ff1-ce00-000000000000' { 'SharePoint / OneDrive (00000003-0000-0ff1-ce00-000000000000)' }
            default { [string]$resource }
          }
        }
        $scope += '<p><strong>Included resources:</strong> ' + [System.Web.HttpUtility]::HtmlEncode(($resourceNames -join '; ')) + '</p>'
      }
      if ($item.ExcludedResources.Count) { $scope += '<p><strong>Excluded resources:</strong> ' + [System.Web.HttpUtility]::HtmlEncode(($item.ExcludedResources -join '; ')) + '</p>' }
      if ($item.DevicePattern) { $scope += '<p><strong>Device pattern:</strong> ' + [System.Web.HttpUtility]::HtmlEncode($item.DevicePattern) + '</p>' }
      if (Test-CaConfiguredValue $item.DeviceFilter) { $scope += '<p><strong>Device filter:</strong> <code>' + [System.Web.HttpUtility]::HtmlEncode("$($item.DeviceFilter.mode): $($item.DeviceFilter.rule)") + '</code></p>' }
      $evidenceHtml += "<div class='rec-policy-evidence'><strong>$policyName</strong> <span class='rec-deployment'>($deployment)</span>$scope<ul>$issues</ul></div>"
    }
    if (-not $evidenceHtml) { $evidenceHtml = '<p>No matching policy evidence in this snapshot. This does not establish that an optional scenario applies.</p>' }
    $htmlFragment += @"
<details class='recommendation-card $RecStatus' data-assessment='$stateKey' data-kind='$kind' data-attention='$attention' id='rec-$encCtrl'$detailOpen>
  <summary><span class='rec-code'>$encCtrl</span><span class='rec-title'>$encName<small>$category</small></span><span class='rec-assessment'>$state</span></summary>
  <div class='rec-body'>
    <p class='rec-status-text'>$message</p>
    $templateHtml
    <div class='rec-guidance-grid'><section><h3>Background</h3><p>$background</p><h3>What to verify / next steps</h3><p>$verification</p></section>
    <section class='rec-prerequisites'><h3>Requirements and prerequisites:</h3><ul>$requirements</ul></section></div>
    <div class='rec-links' aria-label='Microsoft documentation'>$links</div>
    <details class='rec-evidence'><summary>Policy evidence ($($rec.Evidence.Count))</summary>$evidenceHtml</details>
  </div>
</details>
"@
  }
  $htmlFragment += "</div><div class='timestamp-note'>Report generated: $([System.Web.HttpUtility]::HtmlEncode((Get-Date).ToString('u')))</div></div>"
  return $htmlFragment
}

function Get-NamedLocationHtmlFragment {
  [CmdletBinding()]
  param([Parameter(Mandatory)]$Inventory)

  $fragment = @"
<div class='named-locations' id='panel-locations' role='tabpanel' aria-labelledby='tab-locations' aria-hidden='true' style='display:none'>
<div class='page-heading'><p class='eyebrow'>Network inventory</p><h1>Named Locations</h1><p>Tenant network locations available to Conditional Access, including locations not referenced by the exported policies. These definitions are inventory, not an effective-access assessment.</p></div>
"@
  if ($Inventory.Status -ne 'Available') {
    return $fragment + "<p class='summary-note' role='note'>$([System.Web.HttpUtility]::HtmlEncode($Inventory.Message))</p></div>"
  }
  $locations = @($Inventory.Locations)
  if (-not $locations.Count) {
    return $fragment + "<p class='summary-note'>No named locations are defined in this tenant.</p></div>"
  }
  $fragment += "<p class='summary-note'>$($locations.Count) named locations</p><div class='named-location-grid'>"
  foreach ($location in ($locations | Sort-Object DisplayName)) {
    $type = [string](Get-CaProperty $location 'OdataType')
    if (-not $type) { $type = [string]$location.'@odata.type' }
    if (-not $type) {
      $additional = Get-CaProperty $location 'AdditionalProperties'
      $type = if ($additional -is [System.Collections.IDictionary]) { [string]$additional['@odata.type'] } else { [string]$additional.'@odata.type' }
    }
    $name = [string](Get-CaProperty $location 'displayName')
    if ([string]::IsNullOrWhiteSpace($name)) { $name = '(Unnamed location)' }
    $safeName = [System.Web.HttpUtility]::HtmlEncode($name)
    $kind = 'other'
    $label = 'Unknown type'
    $detail = ''
    switch ($type -replace '^#?microsoft\.graph\.', '') {
      'ipNamedLocation' {
        $kind = 'ip'; $label = 'IP ranges'
        $trusted = Get-CaProperty $location 'isTrusted'
        $trustLabel = if ($trusted -is [bool]) { if ($trusted) { 'Trusted' } else { 'Not trusted' } } else { 'Trust not returned' }
        $ranges = @(Get-CaProperty $location 'ipRanges' | Where-Object { $null -ne $_ })
        $detail = "<p class='named-location-flag'>$trustLabel</p><h3>IP ranges ($($ranges.Count))</h3>"
        if ($ranges.Count) {
          $detail += "<ul class='named-location-values'>"
          foreach ($range in $ranges) {
            $cidr = [string](Get-CaProperty $range 'cidrAddress')
            if ([string]::IsNullOrWhiteSpace($cidr)) { $cidr = 'Address not returned' }
            $detail += "<li><code>$([System.Web.HttpUtility]::HtmlEncode($cidr))</code></li>"
          }
          $detail += '</ul>'
        }
        else { $detail += "<p class='summary-note'>No IP ranges returned.</p>" }
      }
      'countryNamedLocation' {
        $kind = 'country'; $label = 'Countries / regions'
        $unknown = Get-CaProperty $location 'includeUnknownCountriesAndRegions'
        $unknownLabel = if ($unknown -is [bool]) { if ($unknown) { 'Included' } else { 'Excluded' } } else { 'Not returned' }
        $lookup = [string](Get-CaProperty $location 'countryLookupMethod')
        if ([string]::IsNullOrWhiteSpace($lookup)) { $lookup = 'Not returned' }
        $countries = @(Get-CaProperty $location 'countriesAndRegions' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)
        $detail = "<p><strong>Unknown countries/regions:</strong> $unknownLabel</p><p><strong>Country lookup:</strong> $([System.Web.HttpUtility]::HtmlEncode($lookup))</p><h3>Countries / regions ($($countries.Count))</h3>"
        if ($countries.Count) {
          $detail += "<ul class='named-location-values'>"
          foreach ($country in $countries) { $detail += "<li>$([System.Web.HttpUtility]::HtmlEncode([string]$country))</li>" }
          $detail += '</ul>'
        }
        else { $detail += "<p class='summary-note'>No countries or regions returned.</p>" }
      }
      default {
        if ($type) { $label = $type }
        $json = [System.Web.HttpUtility]::HtmlEncode(($location | ConvertTo-Json -Depth 20))
        $detail = "<p class='summary-note'>Unrecognized location type; review the returned properties.</p><details><summary>Location properties</summary><pre>$json</pre></details>"
      }
    }
    $metadata = ''
    foreach ($field in @('id', 'createdDateTime', 'modifiedDateTime')) {
      $value = [string](Get-CaProperty $location $field)
      if ([string]::IsNullOrWhiteSpace($value)) { $value = 'Not returned' }
      $fieldLabel = switch ($field) { 'id' { 'ID' } 'createdDateTime' { 'Created' } 'modifiedDateTime' { 'Modified' } }
      $metadata += "<dt>$fieldLabel</dt><dd>$([System.Web.HttpUtility]::HtmlEncode($value))</dd>"
    }
    $fragment += "<article class='summary-card named-location-card' data-location-type='$kind'><header class='named-location-heading'><h2>$safeName</h2><span class='named-location-type'>$([System.Web.HttpUtility]::HtmlEncode($label))</span></header>$detail<dl class='named-location-meta'>$metadata</dl></article>"
  }
  return $fragment + '</div></div>'
}

# If HTML export requested, generate full HTML report with embedded CSS/JS and write to file
if ($HTMLExport) {
  if (-not $NoRecommendations) {
    # Replace the legacy recommendations root div id while preserving original class and avoid overriding base CSS padding.
    # Inline padding previously here caused mismatch with stylesheet; rely on unified .recommendations styling instead.
    $SecurityCheck = (Get-RecommendationsHtmlFragment -Recommendations $recommendations) -replace "id='ca-security-checks' style=''", "id='panel-recommendations' role='tabpanel' aria-labelledby='tab-recommendations' aria-hidden='true' style='display:none'"
  }
  else {
    $SecurityCheck = ''
  }
  $reportTabs = "<button type='button' id='tab-summary' class='btn-toggle' role='tab' aria-selected='false' tabindex='-1' aria-controls='panel-summary'>Summary</button><button type='button' id='tab-policies' class='btn-toggle active' role='tab' aria-selected='true' tabindex='0' aria-controls='panel-policies'>Policy Details</button>"
  if (-not $NoRecommendations) {
    $reportTabs += "<button type='button' id='tab-recommendations' class='btn-toggle' role='tab' aria-selected='false' tabindex='-1' aria-controls='panel-recommendations'>Recommendations</button>"
  }
  $reportTabs += "<button type='button' id='tab-locations' class='btn-toggle' role='tab' aria-selected='false' tabindex='-1' aria-controls='panel-locations'>Named Locations</button>"
  $OmissionBannerHtml = if ($NoRecommendations) {
    "<div class='no-recs-banner' role='note'>Recommendations omitted (-NoRecommendations)</div>" 
  }
  else {
    '' 
  }
  Write-Info 'Saving to File: HTML'
  # Self-contained CSS (no external dependencies)
  $style = @'
  :root {
    color-scheme: light;
    --cp-bg: #f7f4ef;
    --cp-bg-elevated: #fcfbf8;
    --cp-surface: #ffffff;
    --cp-surface-soft: #f5f5f5;
    --cp-border: #dedede;
    --cp-border-strong: #919191;
    --cp-text: #242424;
    --cp-text-muted: #5c5c5c;
    --cp-text-soft: #6f6f6f;
    --cp-accent: #b11f4b;
    --cp-accent-hover: #9a1a41;
    --cp-accent-soft: rgba(177, 31, 75, 0.08);
    --cp-accent-fg: #ffffff;
    --cp-success: #16a34a;
    --cp-danger: #dc2626;
    --cp-warning: #f59e0b;
    --cp-link: #0078d4;
    --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.12);
    --cp-overlay: rgba(255, 255, 255, 0.8);
    --cp-panel: rgba(255, 255, 255, 0.86);
    --cp-panel-strong: rgba(255, 255, 255, 0.96);
    --cp-sheen: rgba(255, 255, 255, 0.55);
    --cp-highlight: rgba(177, 31, 75, 0.12);
  }
  html[data-theme="dark"] {
    color-scheme: dark;
    --cp-bg: #3d3b3a;
    --cp-bg-elevated: #343231;
    --cp-surface: #292929;
    --cp-surface-soft: #2e2e2e;
    --cp-border: #474747;
    --cp-border-strong: #5f5f5f;
    --cp-text: #dedede;
    --cp-text-muted: #919191;
    --cp-text-soft: #b0b0b0;
    --cp-accent: #fd8ea1;
    --cp-accent-hover: #fb7b91;
    --cp-accent-soft: rgba(253, 142, 161, 0.14);
    --cp-accent-fg: #1a1a1a;
    --cp-success: #4ade80;
    --cp-danger: #f87171;
    --cp-warning: #fbbf24;
    --cp-link: #4da6ff;
    --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.32);
    --cp-overlay: rgba(41, 41, 41, 0.88);
    --cp-panel: rgba(41, 41, 41, 0.72);
    --cp-panel-strong: rgba(41, 41, 41, 0.96);
    --cp-sheen: rgba(255, 255, 255, 0.04);
    --cp-highlight: rgba(253, 142, 161, 0.12);
  }
  /* Use the existing secondary text color for accessible dark-mode contrast. */
  html[data-theme="dark"] { --cp-text-muted: var(--cp-text-soft); }
  * { box-sizing:border-box; }
  html, body { margin:0; padding:0; }
  body { font-family:"Segoe UI",Aptos,Calibri,-apple-system,BlinkMacSystemFont,sans-serif; font-size:0.9375rem; line-height:1.5; background:var(--cp-bg); color:var(--cp-text); }
  a { color:var(--cp-link); text-underline-offset:0.15em; }
  button, input, select { font:inherit; }
  button, select, summary { cursor:pointer; }
  button, input, select { border:1px solid var(--cp-border-strong); border-radius:0.625rem; background:var(--cp-surface); color:var(--cp-text); padding:0.5rem 0.75rem; }
  button:hover { background:var(--cp-accent-soft); }
  :focus-visible { outline:2px solid var(--cp-accent); outline-offset:3px; }
  h1, h2, h3, p { margin:0 0 0.75rem; }
  h1 { font-size:1.75rem; line-height:1.2; letter-spacing:-0.025em; }
  h2 { font-size:1.125rem; }
  h3 { font-size:0.9375rem; }
  ul { padding-left:1.25rem; margin:0.5rem 0; }
  li + li { margin-top:0.5rem; }
  code, pre, .rec-code { font-family:Consolas,"Courier New",Courier,monospace; }
  code { overflow-wrap:anywhere; }
  .navbar-custom { position:sticky; top:0; z-index:40; padding:1rem 1.5rem 0; background:var(--cp-surface); border-bottom:1px solid var(--cp-border); }
  .nav-heading { display:flex; flex-wrap:wrap; align-items:baseline; gap:0.5rem 1.5rem; margin-bottom:0.75rem; }
  .brand { font-weight:700; color:var(--cp-accent); white-space:nowrap; }
  .nav-center { flex:1; min-width:0; overflow-wrap:anywhere; font-weight:600; }
  .nav-right { color:var(--cp-text-muted); font-size:0.8125rem; }
  .view-toggle-group { display:flex; flex-wrap:wrap; gap:0.25rem; }
  .btn-toggle { display:inline-block; padding:0.625rem 1rem; border:0; border-bottom:3px solid var(--cp-surface); border-radius:0.625rem 0.625rem 0 0; color:var(--cp-text-muted); font-weight:600; }
  .btn-toggle.active { color:var(--cp-accent); border-bottom-color:var(--cp-accent); background:var(--cp-accent-soft); }
  [role=tabpanel] { padding:1.5rem; }
  .page-heading { margin-bottom:1.5rem; }
  .page-heading p:not(.eyebrow) { color:var(--cp-text-muted); max-width:70rem; }
  .eyebrow { color:var(--cp-accent); font-weight:700; font-size:0.75rem; letter-spacing:0.08em; text-transform:uppercase; margin-bottom:0.5rem; }
  .no-recs-banner, .rec-assessment-note { padding:1rem; border:1px solid var(--cp-border); border-left:4px solid var(--cp-accent); background:var(--cp-surface); border-radius:0.625rem; }
  .no-recs-banner { margin:1.5rem 1.5rem 0; }
  .summary-wrapper, .recommendations, .named-locations { max-width:100rem; margin:0 auto; }
  .named-location-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:1rem; }
  .named-location-heading { display:flex; flex-wrap:wrap; align-items:baseline; justify-content:space-between; gap:0.5rem 1rem; margin-bottom:0.75rem; }
  .named-location-heading h2 { overflow-wrap:anywhere; margin:0; }
  .named-location-type, .named-location-flag { display:inline-block; max-width:100%; overflow-wrap:anywhere; padding:0.25rem 0.5rem; border:1px solid var(--cp-border); border-radius:0.625rem; font-size:0.8125rem; }
  .named-location-values { display:flex; flex-wrap:wrap; gap:0.5rem; list-style:none; padding:0; margin-bottom:1rem; }
  .named-location-values li { margin:0; padding:0.375rem 0.625rem; max-width:100%; overflow-wrap:anywhere; border:1px solid var(--cp-border); border-radius:0.625rem; background:var(--cp-surface-soft); }
  .named-location-meta { display:grid; grid-template-columns:auto minmax(0,1fr); gap:0.5rem 1rem; font-size:0.8125rem; color:var(--cp-text-muted); margin-bottom:0; }
  .named-location-meta dd { margin:0; overflow-wrap:anywhere; }
  .named-location-card pre { white-space:pre-wrap; overflow-wrap:anywhere; }
  .metric-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:1rem; margin-bottom:1.5rem; }
  .metric-card, .summary-card { border:1px solid var(--cp-border); border-radius:16px; background:var(--cp-surface); padding:1.25rem; min-width:0; }
  .metric-value { display:block; font-size:2rem; font-weight:700; line-height:1.2; margin-bottom:0.25rem; }
  .metric-label, .summary-note, .timestamp-note { color:var(--cp-text-muted); font-size:0.875rem; }
  .summary-grid { display:grid; grid-template-columns:minmax(0,2fr) minmax(0,1fr); align-items:start; gap:1.5rem; }
  table { border-collapse:separate; border-spacing:0; width:100%; }
  th, td { padding:0.75rem; border-bottom:1px solid var(--cp-border); vertical-align:top; text-align:left; }
  th { font-weight:600; background:var(--cp-surface-soft); }
  .summary-card.summary-inventory { padding:0; overflow:hidden; }
  .summary-inventory > h2 { margin:0; padding:1.25rem 1.25rem 0.5rem; }
  .summary-inventory > .summary-note { padding:0 1.25rem 1.25rem; margin:0; }
  .summary-table-wrap { overflow-x:auto; max-width:100%; }
  .summary-table { font-size:0.9375rem; min-width:28rem; }
  .summary-table .summary-metric-col { width:56%; }
  .summary-table .summary-value-col { width:16%; }
  .summary-table .summary-share-col { width:28%; }
  .summary-table th, .summary-table td { padding:0.75rem 1rem; vertical-align:middle; }
  .summary-table thead th { font-size:0.75rem; color:var(--cp-text-muted); letter-spacing:0.04em; text-transform:uppercase; }
  .summary-table thead th:not(:first-child), .summary-value, .summary-share-cell { text-align:right; font-variant-numeric:tabular-nums; }
  .summary-table tbody th[scope=row] { font-weight:500; background:var(--cp-surface); line-height:1.4; }
  .summary-table .summary-section-row th { padding:0.625rem 1rem; background:var(--cp-accent-soft); color:var(--cp-accent); font-size:0.75rem; letter-spacing:0.06em; text-transform:uppercase; }
  .summary-table tbody + tbody .summary-section-row th { border-top:0.5rem solid var(--cp-surface); }
  .summary-table .summary-metric-row:hover th, .summary-table .summary-metric-row:hover td { background:var(--cp-surface-soft); }
  .summary-table tbody:last-child tr:last-child th, .summary-table tbody:last-child tr:last-child td { border-bottom:0; }
  .summary-number { font-weight:650; font-size:1.0625rem; white-space:nowrap; }
  .summary-value-unit { display:block; font-size:0.75rem; color:var(--cp-text-muted); }
  .summary-share { display:grid; gap:0.375rem; min-width:5rem; max-width:10rem; margin-left:auto; }
  .summary-share-label { display:flex; flex-wrap:wrap; justify-content:flex-end; align-items:baseline; gap:0.125rem 0.375rem; }
  .summary-share-value { font-weight:600; white-space:nowrap; }
  .summary-share-base, .summary-na { color:var(--cp-text-muted); font-size:0.75rem; white-space:nowrap; }
  .summary-share-track { display:block; height:0.25rem; border-radius:0.25rem; overflow:hidden; background:var(--cp-border); }
  .summary-share-fill { display:block; height:100%; background:var(--cp-accent); border-radius:0.25rem; }
  .assessment-list { display:grid; grid-template-columns:minmax(0,1fr) auto; gap:0.625rem 1rem; margin:1rem 0 1.5rem; }
  .assessment-list dt { color:var(--cp-text-muted); }
  .assessment-list dd { margin:0; font-weight:700; font-variant-numeric:tabular-nums; }
  .policy-toolbar { display:flex; flex-wrap:wrap; align-items:end; gap:1rem; margin-bottom:1rem; }
  .search-box { position:relative; flex:1 1 18rem; min-width:0; }
  .search-box label, .status-filter label { display:block; font-weight:600; font-size:0.8125rem; margin-bottom:0.375rem; }
  .search-box input { width:100%; padding-right:3rem; }
  .search-clear { position:absolute; right:0.25rem; bottom:0.125rem; display:none; padding:0.375rem 0.75rem; border:0; }
  .search-box.has-value .search-clear { display:block; }
  .status-filter select { max-width:100%; }
  .policy-count { font-size:0.875rem; color:var(--cp-text-muted); padding-bottom:0.625rem; }
  .legend-bar { display:flex; flex-wrap:wrap; gap:0.75rem 1rem; align-items:center; padding:1rem; border:1px solid var(--cp-border); border-radius:0.625rem; margin-bottom:1rem; background:var(--cp-surface); font-size:0.8125rem; }
  .legend-title { font-weight:700; }
  .legend-item { display:inline-flex; align-items:center; gap:0.375rem; }
  .legend-swatch { display:inline-flex; align-items:center; justify-content:center; min-width:1.75rem; padding:0.125rem 0.25rem; border:1px solid var(--cp-border); border-radius:0.25rem; font-size:0.75rem; font-weight:600; }
  .legend-divider { width:1px; height:1.5rem; background:var(--cp-border); }
  .bool-mode-toggle { margin-left:auto; }
  .util-button, .layout-toggle, .legend-chip, .json-toggle-bar button { font-size:0.8125rem; }
  .util-button.active, .layout-toggle.active, .legend-chip.active, .json-toggle-bar button.active { color:var(--cp-accent-fg); background:var(--cp-accent); border-color:var(--cp-accent); }
  .table-scroll { overflow:auto; max-width:100%; max-height:72vh; border:1px solid var(--cp-border-strong); border-radius:0.625rem; background:var(--cp-surface); isolation:isolate; }
  .sticky-header { width:max-content; min-width:100%; font-size:0.875rem; }
  .sticky-header th, .sticky-header td { min-width:9rem; max-width:22rem; border-right:1px solid var(--cp-border); overflow-wrap:anywhere; }
  .sticky-header td { white-space:pre-line; }
  .sticky-header .name-col { min-width:min(17rem,45vw); width:min(19rem,45vw); max-width:min(24rem,45vw); }
  .sticky-header .bool-col { min-width:6rem; width:6rem; text-align:center; }
  .sticky-header thead th { position:sticky; z-index:5; }
  .sticky-header .group-header th { top:0; background:var(--cp-accent); color:var(--cp-accent-fg); font-size:0.75rem; letter-spacing:0.035em; text-transform:uppercase; white-space:nowrap; }
  .sticky-header .column-header-row th { top:var(--group-header-height,2.625rem); background:var(--cp-surface-soft); }
  .sticky-header .sticky-name { position:sticky; left:0; z-index:3; background:var(--cp-surface); border-right:2px solid var(--cp-border-strong); }
  .sticky-header thead .sticky-name { z-index:6; }
  .sticky-header tbody tr:nth-child(even), .sticky-header tbody tr:nth-child(even) .sticky-name { background:var(--cp-surface-soft); }
  .sticky-header tbody tr:hover td, .sticky-header tbody tr.selected td, .sticky-header td.selected, .sticky-header th.selected { background:var(--cp-highlight); }
  .sticky-header tbody tr:hover .sticky-name, .sticky-header tbody tr.selected .sticky-name { background:var(--cp-bg-elevated); }
  td.sticky-name { padding:0; }
  td.sticky-name .name-content { display:block; padding:0.75rem 0.75rem 0.75rem 2.25rem; min-height:5.5rem; white-space:normal; font-weight:600; }
  td.sticky-name .status-label { position:absolute; left:0; top:0; bottom:0; width:1.5rem; height:100%; }
  .status-label { writing-mode:vertical-rl; transform:rotate(180deg); font-size:0.75rem; font-weight:700; padding:0.5rem 0; white-space:nowrap; display:flex; align-items:center; justify-content:center; text-transform:uppercase; letter-spacing:0.5px; background:var(--cp-surface-soft); border-left:3px solid; }
  .status-enabled, .legend-swatch.enabled, .legend-swatch.pass, .bool-yes { color:var(--cp-success); }
  .status-report, .legend-swatch.report, .legend-swatch.fail { color:var(--cp-text); border-color:var(--cp-warning); }
  .status-disabled, .legend-swatch.disabled { color:var(--cp-danger); }
  .status-label.status-enabled, .legend-swatch.enabled, .legend-swatch.pass { color:var(--cp-text); border-color:var(--cp-success); }
  .policy-export:has(#bool-mode-toggle[data-mode=text]) .bool-yes { color:var(--cp-text); font-weight:600; }
  .bool-no, .placeholder { color:var(--cp-text-soft); }
  .status-badge, .pill-block, .pill-grant, .pill-recommended { display:inline-block; padding:0.25rem 0.5rem; border:1px solid var(--cp-border); border-radius:0.625rem; font-size:0.75rem; font-weight:600; }
  .pill-block { color:var(--cp-danger); }
  .pill-grant { color:var(--cp-text); border-color:var(--cp-border-strong); }
  .pill-recommended { color:var(--cp-accent); margin-left:0.375rem; }
  .naming-guidance { margin-top:0.5rem; font-size:0.75rem; font-weight:400; }
  .naming-guidance summary { color:var(--cp-warning); cursor:pointer; }
  .naming-guidance ul { margin:0.5rem 0; padding-left:1rem; }
  .suggested-name { margin-top:0.5rem; color:var(--cp-text-muted); font-weight:400; font-size:0.75rem; }
  .comparison-bar { padding:0.75rem; margin-bottom:1rem; border:1px solid var(--cp-border); border-radius:0.625rem; background:var(--cp-surface); }
  .comparison-bar p { margin:0.5rem 0; font-size:0.8125rem; color:var(--cp-text-muted); }
  .comparison-controls { display:flex; flex-wrap:wrap; gap:0.75rem; align-items:center; }
  .entity-token { display:block; max-width:100%; margin:0.25rem 0; padding:0.25rem 0.375rem; text-align:left; font-size:0.8125rem; white-space:normal; overflow-wrap:anywhere; }
  .entity-token.match-include { border:2px solid var(--cp-accent); background:var(--cp-accent-soft); }
  .entity-token.match-exclude { border:2px dashed var(--cp-warning); background:var(--cp-surface-soft); }
  .entity-token.match-include::before { content:"Included match: "; font-weight:700; }
  .entity-token.match-exclude::before { content:"Excluded match: "; font-weight:700; }
  .sticky-header td.cell-match { outline:2px solid var(--cp-accent); outline-offset:-2px; background:var(--cp-accent-soft); }
  .compare-cell { display:block; margin-top:0.5rem; padding:0.125rem 0.375rem; font-size:0.6875rem; color:var(--cp-text-muted); }
  .policy-link { display:inline-flex; vertical-align:middle; align-items:center; justify-content:center; width:1.75rem; height:1.75rem; margin-left:0.25rem; border-radius:0.25rem; }
  .policy-link:hover { background:var(--cp-accent-soft); }
  svg.icon-ext { width:1rem; height:1rem; stroke:currentColor; stroke-width:1.9; fill:none; }
  .desc-cell, .assignment-cell { position:relative; padding-bottom:2.5rem; }
  .desc-text, .assignment-text { display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; white-space:normal; }
  .desc-cell.expanded .desc-text, .assignment-cell.expanded .assignment-text { -webkit-line-clamp:unset; }
  .desc-expand, .assignment-expand { position:absolute; bottom:0.25rem; right:0.25rem; padding:0.125rem 0.5rem; font-size:0.75rem; }
  details.raw-json { width:24rem; max-width:100%; white-space:normal; }
  details.raw-json summary, .rec-evidence > summary { font-weight:600; color:var(--cp-link); }
  details.raw-json pre { max-height:22rem; overflow:auto; background:var(--cp-surface-soft); padding:0.75rem; border-radius:0.625rem; font-size:0.8125rem; white-space:pre; }
  .json-toggle-bar { display:flex; flex-wrap:wrap; gap:0.5rem; margin:0.75rem 0; }
  .rec-assessment-note { color:var(--cp-text-muted); font-size:0.875rem; margin-bottom:1rem; }
  .rec-toolbar { display:flex; flex-wrap:wrap; align-items:flex-end; gap:0.75rem 1rem; padding:1rem; border:1px solid var(--cp-border); border-radius:0.625rem; background:var(--cp-surface); margin-bottom:1.5rem; }
  .rec-filter-group { display:flex; flex:1 1 32rem; flex-wrap:wrap; gap:0.5rem; min-width:0; }
  .rec-filter-group .util-button { font-size:0.875rem; }
  .rec-assessment-control { display:flex; flex-direction:column; gap:0.375rem; min-width:0; max-width:100%; }
  .rec-assessment-control label { font-size:0.8125rem; font-weight:600; color:var(--cp-text-muted); }
  .rec-assessment-control select { font-size:0.875rem; max-width:100%; }
  .rec-filter-summary { flex-basis:100%; margin:0; color:var(--cp-text-muted); font-size:0.8125rem; }
  .rec-filter-empty { padding:1rem; border:1px solid var(--cp-border); border-radius:0.625rem; background:var(--cp-surface); color:var(--cp-text-muted); }
  details.recommendation-card { border:1px solid var(--cp-border); border-left:4px solid var(--cp-border-strong); border-radius:16px; background:var(--cp-surface); margin-bottom:1rem; }
  details.recommendation-card.pass { border-left-color:var(--cp-success); }
  details.recommendation-card.fail { border-left-color:var(--cp-warning); }
  details.recommendation-card[data-kind=review], details.recommendation-card[data-kind=emergency] { border-left-color:var(--cp-border-strong); }
  details.recommendation-card[data-kind=template] { border-left-color:var(--cp-accent); }
  .recommendation-card > summary { display:flex; flex-wrap:wrap; align-items:center; gap:0.75rem; padding:1rem 1.25rem; list-style:none; font-weight:600; }
  .recommendation-card > summary::-webkit-details-marker { display:none; }
  .recommendation-card > summary::before { content:"+"; color:var(--cp-accent); font-size:1.25rem; width:0.75rem; }
  .recommendation-card[open] > summary::before { content:"-"; }
  .rec-code { background:var(--cp-surface-soft); color:var(--cp-text-muted); border:1px solid var(--cp-border); padding:0.25rem 0.5rem; border-radius:0.25rem; font-size:0.8125rem; }
  .rec-title { flex:1 1 14rem; min-width:0; overflow-wrap:anywhere; }
  .rec-title small { display:block; font-size:0.75rem; color:var(--cp-text-muted); font-weight:400; margin-top:0.125rem; }
  .rec-assessment, .template-measurement { padding:0.25rem 0.625rem; background:var(--cp-surface-soft); border:1px solid var(--cp-border); border-radius:0.625rem; font-size:0.8125rem; font-weight:600; }
  .rec-body { padding:0 1.25rem 1.25rem; }
  .rec-body p, .rec-body li { line-height:1.6; }
  .rec-status-text { padding-top:1rem; border-top:1px solid var(--cp-border); color:var(--cp-text-muted); }
  .rec-guidance-grid { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:1.5rem; margin:1.25rem 0; }
  .rec-guidance-grid h3 { margin-bottom:0.5rem; }
  .rec-links { display:flex; flex-wrap:wrap; gap:0.5rem; padding-top:0.75rem; border-top:1px solid var(--cp-border); }
  .rec-link { font-size:0.8125rem; padding:0.375rem 0.625rem; border:1px solid var(--cp-border); border-radius:0.625rem; overflow-wrap:anywhere; }
  .rec-link:hover { background:var(--cp-accent-soft); }
  .rec-evidence { margin-top:1rem; }
  .rec-policy-evidence { padding:1rem; margin-top:0.75rem; border:1px solid var(--cp-border); border-radius:0.625rem; background:var(--cp-surface-soft); overflow-wrap:anywhere; }
  .rec-policy-evidence p { margin:0.5rem 0; }
  .rec-deployment { font-size:0.8125rem; color:var(--cp-text-muted); }
  .rec-template { padding:1rem; margin:1rem 0; border:1px solid var(--cp-border); border-radius:0.625rem; background:var(--cp-accent-soft); }
  .rec-template h3 { display:flex; flex-wrap:wrap; gap:0.5rem 1rem; align-items:center; }
  .rec-template dl { display:grid; grid-template-columns:minmax(6rem,1fr) minmax(0,4fr); gap:0.75rem 1rem; margin-bottom:0; }
  .rec-template dt { font-weight:600; }
  .rec-template dd { margin:0; overflow-wrap:anywhere; }
  .timestamp-note { margin:1.5rem 0 0; }
  #back-to-top { display:none; position:fixed; right:1rem; bottom:1rem; z-index:30; background:var(--cp-accent); color:var(--cp-accent-fg); }
  .visually-hidden { position:absolute !important; width:1px !important; height:1px !important; padding:0 !important; margin:-1px !important; overflow:hidden !important; clip:rect(0 0 0 0) !important; white-space:nowrap !important; border:0 !important; }
  @media (max-width:900px) {
    .summary-grid, .rec-guidance-grid, .named-location-grid { grid-template-columns:minmax(0,1fr); }
    .metric-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
  }
  @media (max-width:600px) {
    .navbar-custom { position:relative; padding:0.75rem 0.75rem 0; }
    .nav-center { flex-basis:100%; order:1; }
    .nav-right { margin-left:auto; }
    .btn-toggle { padding:0.5rem; font-size:0.875rem; }
    [role=tabpanel] { padding:1rem 0.75rem; }
    .metric-card, .summary-card { padding:0.875rem; }
    .metric-grid, .summary-grid { gap:0.75rem; }
    .recommendation-card > summary, .rec-body { padding-left:0.875rem; padding-right:0.875rem; }
    .rec-template dl { grid-template-columns:minmax(0,1fr); gap:0.25rem; }
    .rec-template dd { margin-bottom:0.5rem; }
    .sticky-header .name-col { min-width:min(12rem,45vw); width:min(12rem,45vw); max-width:min(12rem,45vw); }
    .summary-table { min-width:22rem; }
    .summary-table th, .summary-table td { padding:0.625rem 0.75rem; }
    .summary-table .summary-section-row th { padding-left:0.75rem; padding-right:0.75rem; }
  }
'@

  $safeTenantName = [System.Web.HttpUtility]::HtmlEncode($TenantName)
  $reportDateText = if ($OfflineSnapshot.SnapshotFormat) { "Snapshot captured: $Date" } else { $Date }
  $safeReportDate = [System.Web.HttpUtility]::HtmlEncode($reportDateText)
  $htmlDoc = @"
<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<script>
(() => {
  const param = new URLSearchParams(window.location.search).get("scoutTheme");
  const theme = param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  document.documentElement.setAttribute("data-theme", theme);
})();
</script>
<style>$style</style><title>CA Export v$CAExportVersion - $safeTenantName</title></head><body>
<nav class='navbar-custom' aria-label='Conditional Access report'>
  <div class='nav-heading'><div class='brand'>CA Export $CAExportVersion</div><div class='nav-center'>$safeTenantName</div><div class='nav-right'>$safeReportDate</div></div>
  <div class='view-toggle-group' role='tablist' aria-label='Report view'>$reportTabs</div>
</nav>$OmissionBannerHtml
<div id='live-region' class='visually-hidden' aria-live='polite' aria-atomic='true'></div><button type='button' id='back-to-top' aria-label='Back to top' title='Back to top'>&uarr;</button>
"@

  Write-Info 'Launching: Web Browser'
  # Ensure export path ends with separator
  if ($ExportLocation -and $ExportLocation[-1] -notin @('\', '/')) {
    $ExportLocation = $ExportLocation + [IO.Path]::DirectorySeparatorChar 
  }
  $Launch = Join-Path -Path $ExportLocation -ChildPath $FileName
  # Build the summary fragment first, then add the real summary panel (no placeholder/replacement needed)

  # ----- Summary Table -----
  $policyTotal = $CAPolicy.Count
  $enabledCount = ($CAPolicy | Where-Object { $_.State -eq 'enabled' }).Count
  $reportOnly = ($CAPolicy | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' }).Count
  $disabledCount = ($CAPolicy | Where-Object { $_.State -eq 'disabled' }).Count
  $withMfa = @($CAPolicy | Where-Object { @(Get-CaProperty $_ 'grantControls.builtInControls') -contains 'mfa' -or (Test-CaMfaStrength -Policy $_) }).Count
  $withStrength = @($CAPolicy | Where-Object { Test-CaConfiguredValue (Get-CaProperty $_ 'grantControls.authenticationStrength') }).Count
  $withBlock = ($CAExport | Where-Object { $_.Block -eq 'True' }).Count
  $riskPolicies = ($CAPolicy | Where-Object { $_.Conditions.SignInRiskLevels -or $_.Conditions.UserRiskLevels }).Count
  $devicePolicies = ($CAPolicy | Where-Object { $_.Conditions.Devices.IncludeDevices -or $_.Conditions.Devices.ExcludeDevices -or $_.GrantControls.BuiltInControls -contains 'compliantDevice' -or $_.GrantControls.BuiltInControls -contains 'domainJoinedDevice' }).Count
  $termsPolicies = ($CAPolicy | Where-Object { $_.GrantControls.TermsOfUse }).Count
  $phishResistant = @($CAPolicy | Where-Object { Test-IsPhishResistantStrength -Policy $_ }).Count
  $avgModifiedDays = 0
  $now = Get-Date
  $dateVals = @()
  foreach ( $p in $CAExport ) {
    if ( $p.DateModified ) {
      $d = [datetime] $p.DateModified
      $dateVals += ( ( $now - $d ).TotalDays )
    }
  }
  if ( $dateVals.Count -gt 0 ) {
    $avgModifiedDays = [math]::Round( ( $dateVals | Measure-Object -Average | Select-Object -ExpandProperty Average ), 1 )
  }

  $summaryTable = @()
  $summaryTable += '<div class="summary-section summary-wrapper" id="summary-root">'
  $summaryTable += '<div class="page-heading"><p class="eyebrow">Tenant snapshot</p><h1>Policy summary</h1><p>Deployment counts and configuration inventory. Counts include enabled, report-only and disabled policies; they do not measure effective user coverage.</p></div>'
  $summaryTable += "<div class='metric-grid'><div class='metric-card'><span class='metric-value'>$policyTotal</span><span class='metric-label'>Total policies</span></div><div class='metric-card'><span class='metric-value'>$enabledCount</span><span class='metric-label'>Enabled</span></div><div class='metric-card'><span class='metric-value'>$reportOnly</span><span class='metric-label'>Report-only</span></div><div class='metric-card'><span class='metric-value'>$disabledCount</span><span class='metric-label'>Disabled</span></div></div>"
  $summaryTable += '<div class="summary-grid"><section class="summary-card summary-inventory"><h2>Configuration inventory</h2><p class="summary-note" id="summary-scope-note">Each share shows its policy-count denominator. Administrator MFA and strength metrics reference administrator-targeted policies; other shares reference all policies. A dash means not applicable.</p>'
  $summaryTable += '<div class="summary-table-wrap" role="region" aria-label="Policy summary metrics" tabindex="0"><table class="summary-table" aria-describedby="summary-scope-note"><caption class="visually-hidden">Policy configuration inventory by category</caption><colgroup><col class="summary-metric-col"><col class="summary-value-col"><col class="summary-share-col"></colgroup>'
  $summaryTable += '<thead><tr><th scope="col">Metric</th><th scope="col">Count / value</th><th scope="col">Share of scope</th></tr></thead><tbody aria-labelledby="summary-deployment"><tr class="summary-section-row"><th scope="rowgroup" colspan="3" id="summary-deployment">Deployment</th></tr>'
  function Add-SummaryRow {
    param(
      [string]$Name,
      $Value,
      [double]$TotalRef,
      [string]$Unit
    )
    $pct = ''
    $numericVal = 0.0
    $canParse = $false
    if ($null -ne $Value) {
      if (
        $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        $numericVal = [double]$Value; $canParse = $true
      }
      elseif ($Value -isnot [string]) {
        try {
          $numericVal = [double]$Value; $canParse = $true 
        }
        catch {
          Write-Verbose 'Add-SummaryRow: numeric conversion failed.' 
        }
      }
      elseif ([double]::TryParse([string]$Value, [ref]$numericVal)) {
        $canParse = $true
      }
    }
    $shareHtml = '<span class="summary-na" aria-label="Not applicable">&mdash;</span>'
    if ($canParse -and $TotalRef -gt 0) {
      $pct = '{0:P1}' -f ($numericVal / $TotalRef)
      $barPercent = [math]::Min(100.0, [math]::Max(0.0, 100 * $numericVal / $TotalRef))
      $barWidth = $barPercent.ToString('0.###', [System.Globalization.CultureInfo]::InvariantCulture)
      $scopeCount = [System.Web.HttpUtility]::HtmlEncode([string]$TotalRef)
      $encodedPct = [System.Web.HttpUtility]::HtmlEncode($pct)
      $shareHtml = "<div class='summary-share'><div class='summary-share-label'><span class='summary-share-value'>$encodedPct</span><span class='summary-share-base'>of $scopeCount</span></div><span class='summary-share-track' aria-hidden='true'><span class='summary-share-fill' style='width:$barWidth%'></span></span></div>"
    }
    $encName = [System.Web.HttpUtility]::HtmlEncode($Name)
    $displayValue = if ($Value -is [string]) {
      [System.Web.HttpUtility]::HtmlEncode($Value) 
    }
    else {
      $Value 
    }
    $unitHtml = if ($Unit) { '<span class="summary-value-unit">' + [System.Web.HttpUtility]::HtmlEncode($Unit) + '</span>' } else { '' }
    return "<tr class='summary-metric-row'><th scope='row'>$encName</th><td class='summary-value'><span class='summary-number'>$displayValue</span>$unitHtml</td><td class='summary-share-cell'>$shareHtml</td></tr>"
  }
  # Total and maintenance values have no percentage denominator.
  $summaryTable += (Add-SummaryRow -Name 'Total Policies' -Value $policyTotal -TotalRef '')
  $summaryTable += (Add-SummaryRow -Name 'Enabled' -Value $enabledCount -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Report-only' -Value $reportOnly -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Disabled' -Value $disabledCount -TotalRef $policyTotal)
  $summaryTable += '</tbody><tbody aria-labelledby="summary-authentication"><tr class="summary-section-row"><th scope="rowgroup" colspan="3" id="summary-authentication">Authentication</th></tr>'
  $summaryTable += (Add-SummaryRow -Name 'Policies Requiring MFA (any)' -Value $withMfa -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Policies With Auth Strength' -Value $withStrength -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Phishing-resistant Strength' -Value $phishResistant -TotalRef $policyTotal)
  $summaryTable += '</tbody><tbody aria-labelledby="summary-access"><tr class="summary-section-row"><th scope="rowgroup" colspan="3" id="summary-access">Access controls and conditions</th></tr>'
  $summaryTable += (Add-SummaryRow -Name 'Policies With Block Control' -Value $withBlock -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Risk-Based Policies' -Value $riskPolicies -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Device Condition / Control Policies' -Value $devicePolicies -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Terms of Use Policies' -Value $termsPolicies -TotalRef $policyTotal)
  # Admin-focused metrics using raw policy objects with helper functions
  if ($RawPolicyIndex.Count -gt 0) {
    $rawPolicies = $RawPolicyIndex.Values
    $adminTargets = 0
    $adminMfa = 0
    $adminPhish = 0
    foreach ($rp in $rawPolicies) {
      if (Test-PolicyTargetsAdminRoles -Policy $rp) {
        $adminTargets++
        if (Test-PolicyRequiresMfaForAdmins -Policy $rp) { $adminMfa++ }
        if (Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $rp) { $adminPhish++ }
      }
    }
    $summaryTable += '</tbody><tbody aria-labelledby="summary-administrators"><tr class="summary-section-row"><th scope="rowgroup" colspan="3" id="summary-administrators">Administrator policies</th></tr>'
    $summaryTable += (Add-SummaryRow -Name 'Policies Targeting Admin Roles' -Value $adminTargets -TotalRef $policyTotal)
    $summaryTable += (Add-SummaryRow -Name 'Admin Policies Requiring MFA' -Value $adminMfa -TotalRef $adminTargets)
    $summaryTable += (Add-SummaryRow -Name 'Admin Policies Phish-Resistant' -Value $adminPhish -TotalRef $adminTargets)
  }
  $summaryTable += '</tbody><tbody aria-labelledby="summary-maintenance"><tr class="summary-section-row"><th scope="rowgroup" colspan="3" id="summary-maintenance">Policy maintenance</th></tr>'
  $summaryTable += (Add-SummaryRow -Name 'Average Age Since Modified' -Value $avgModifiedDays -Unit 'days')
  $summaryTable += '</tbody></table></div>'
  $summaryTable += '</section><aside class="summary-card"><h2>Reading this report</h2>'
  if (-not $NoRecommendations) {
    $summaryTable += '<p class="summary-note">Protection and configuration-hygiene assessment states (optional/operational reviews are counted separately):</p><dl class="assessment-list">'
    foreach ($assessmentState in @('Enforced','Report-only','Disabled','Missing','Needs review','No findings')) {
      $stateCount = @($recommendations | Where-Object { $_.Kind -in @('Protection','Hygiene') -and $_.Assessment -eq $assessmentState }).Count
      $summaryTable += "<dt>$assessmentState</dt><dd>$stateCount</dd>"
    }
    $templateCount = @($recommendations | Where-Object Kind -eq 'Template').Count
    $reviewCount = @($recommendations | Where-Object { $_.Kind -in @('Review','Emergency') }).Count
    $summaryTable += "</dl><p><strong>$templateCount optional templates</strong> and <strong>$reviewCount optional/operational reviews</strong> are separate from the counts above. Their policy evidence and deployment state do not imply applicability; absence is not a baseline failure.</p>"
  }
  $summaryTable += '<p>Use <strong>Policy Details</strong> to inspect assignments and controls. Search and deployment filters work together; scroll within the table to reach every field.</p><p class="summary-note">A configuration match is not a sign-in simulation. Validate group/PIM membership, exclusions, licensing, application setup and sign-in results before changing policies.</p></aside></div></div>'
  $SummaryHtmlFragment = $summaryTable -join "`n"
  # Add populated summary panel, then open policies panel
  $HtmlParts += "<div class='policy-summary' id='panel-summary' role='tabpanel' aria-labelledby='tab-summary' aria-hidden='true' style='display:none;'>$SummaryHtmlFragment</div>"
  $HtmlParts += "<div class='policy-export' id='panel-policies' role='tabpanel' aria-labelledby='tab-policies' aria-hidden='false'>"
  $HtmlParts += @'
  <div class='page-heading'><p class='eyebrow'>Assignments &amp; controls</p><h1>Policy details</h1><p>Search policy values and combine with a deployment-state filter. The policy name and both header rows stay visible as you scroll within the table.</p></div>
  <div class='policy-toolbar'>
    <div class='search-box'><label for='policy-search'>Search policies</label><input id='policy-search' type='search' placeholder='Name, user, group, application or setting...'/><button type='button' class='search-clear' id='policy-search-clear' title='Clear search' aria-label='Clear search'>&times;</button></div>
    <div class='status-filter'><label for='status-filter'>Deployment state</label><select id='status-filter'><option value='all'>All policies</option><option value='enabled'>Enabled only</option><option value='report'>Report-only</option><option value='disabled'>Disabled only</option></select></div>
    <label><input id='show-suggested-names' type='checkbox'/> Show suggested names</label>
    <span class='policy-count' id='policy-count' role='status'></span>
  </div>
  <div class='comparison-bar'>
    <div class='comparison-controls'><strong>Pattern comparison</strong><label><input id='match-only' type='checkbox' disabled/> Show matching policies only</label><button type='button' id='clear-comparison' disabled>Clear comparison</button></div>
    <p id='comparison-status' role='status'>Click an assignment to match its type and original ID across includes/excludes, or Compare to match an exact value in the same field.</p>
    <p>Included matches use a solid border; excluded matches use a dashed border and explicit labels. Counts describe direct assignments in this export, not group membership or effective access. All users and group assignments are not expanded. Search and deployment filters still apply.</p>
  </div>
'@
  # Legend / utilities bar
  $legendHtml = @'
  <div class='legend-bar' id='legend-bar'>
    <span class='legend-title'>Legend:</span>
    <span class='legend-item'><span class='legend-swatch'>—</span> Empty / None</span>
    <span class='legend-item'><span class='legend-swatch enabled'>EN</span> Enabled</span>
    <span class='legend-item'><span class='legend-swatch report'>RP</span> Report-only</span>
    <span class='legend-item'><span class='legend-swatch disabled'>DIS</span> Disabled</span>
    <span class='legend-item'><span class='legend-swatch pass'>✔</span> Rec Enforced configuration</span>
    <span class='legend-item'><span class='legend-swatch fail'>⚠</span> Rec Attention</span>
    <span class='legend-item'><span class='legend-swatch'>✔</span> Boolean True</span>
    <span class='legend-item'><span class='legend-swatch'>—</span> Boolean False</span>
    <span class='legend-divider' aria-hidden='true'></span>
  <button type='button' id='bool-mode-toggle' class='util-button bool-mode-toggle' data-mode='icon' aria-pressed='false' title='Toggle boolean display mode'>Boolean: Icons</button>
  </div>
'@
  $HtmlParts += $legendHtml
  # Define columns (reuse CSV default ordering for consistency, but can trim for HTML readability)
  $htmlColumns = @(
    'Name', 'Modified', 'Created',
    'Included Users', 'Excluded Users',
    'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations', 'User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows', 'Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA',
    'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls',
    'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session', 'RawJson'
  )
  # Build table rows per policy
  $table = @()
  $table += '<div class="table-scroll" role="region" aria-label="Conditional Access policy configuration, horizontally and vertically scrollable" tabindex="0"><table class="sticky-header">'
  # Specify boolean columns for icon rendering
  $boolColumns = @('Block', 'Require MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Application Enforced Restrictions', 'Cloud App Security', 'Resilient Defaults')
  # Build multi-row header: first row (spanning categories) + second row (individual columns)
  $groupDefinitions = @(
    @{ Name = 'General'; Class = 'group-general'; Columns = @('Name', 'Modified', 'Created') },
    @{ Name = 'Assignments - Users'; Class = 'group-assign-users'; Columns = @('Included Users', 'Excluded Users') },
    @{ Name = 'Assignments - Target Resources'; Class = 'group-assign-target'; Columns = @('Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context') },
    @{ Name = 'Assignments - Network'; Class = 'group-network'; Columns = @('Included Locations', 'Excluded Locations') },
    @{ Name = 'Assignments - Conditions'; Class = 'group-conditions'; Columns = @('User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows') },
    @{ Name = 'Access controls - Grant'; Class = 'group-grant'; Columns = @('Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls') },
    @{ Name = 'Access controls - Session'; Class = 'group-session'; Columns = @('Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session') },
    @{ Name = 'Output'; Class = 'group-output'; Columns = @('RawJson') }
  )
  # Sanity: ensure ordering in htmlColumns matches group concatenation; if mismatch, groups may mis-align
  # Compose group header row
  $groupRow = '<tr class="group-header">' + (
    $groupDefinitions | ForEach-Object {
      $totalCols = $_.Columns.Count
      # Count how many ID columns are in this group (these will be hidden)
      $hiddenCount = ($_.Columns | Where-Object { $idColumns -contains $_ }).Count
      $visibleColspan = $totalCols - $hiddenCount
      $gName = [System.Web.HttpUtility]::HtmlEncode($_.Name)
      $gClass = $_.Class
      "<th class='group-span $gClass' colspan='$visibleColspan' data-total='$totalCols' data-visible='$visibleColspan'>$gName</th>"
    }
  ) -join '' + '</tr>'
  # Map column -> group class
  $columnGroupClassMap = @{}
  foreach ($gd in $groupDefinitions) { foreach ($col in $gd.Columns) { $columnGroupClassMap[$col] = $gd.Class } }
  # Compose second (field) header row with group color classes
  $fieldHeaderRow = '<tr class="column-header-row">' + ($htmlColumns | ForEach-Object {
      $gCls = $columnGroupClassMap[$_]
      if ($_ -eq 'Name') {
        "<th id='th-name' class='sticky-name name-col $gCls'>Name</th>" 
      }
      elseif ($boolColumns -contains $_) {
        "<th class='bool-col $gCls'>$_</th>" 
      }
      else {
        "<th class='$gCls'>$_</th>" 
      }
    }) -join '' + '</tr>'
  $header = '<thead>' + $groupRow + $fieldHeaderRow + '</thead>'
  $table += $header
  $table += '<tbody>'
  foreach ($p in $CAExport) {
    $safePolicyId = [System.Web.HttpUtility]::HtmlEncode([string]$p.PolicyId)
    $rowTds = @()
    $colIndex = 0
    foreach ($col in $htmlColumns) {
      $colIndex++
      $raw = $null
      # Use direct property access instead of Match() which uses wildcards
      if ($p.PSObject.Properties.Name -contains $col) {
        $raw = $p.$col 
      }
      
      if ($col -eq 'RawJson') {
        # Lazy JSON: store minimal placeholder with data attributes; JSON filled client-side on first expansion
        $hasOrig = ($RawPolicyIndex -and $p.PolicyId -and $RawPolicyIndex.ContainsKey($p.PolicyId))
        $origFlag = if ($hasOrig) { 'true' } else { 'false' }
        $rowTds += "<td class='$idHiddenClass'><details class='raw-json lazy-json' data-policyid='$safePolicyId' data-has-orig='$origFlag'><summary>View</summary><div class='json-toggle-bar'><button type='button' class='json-btn active' data-mode='mut' aria-pressed='true'>Mutated</button><button type='button' class='json-btn' data-mode='orig' aria-pressed='false'$(if($hasOrig){''}else{' disabled title=\"No original snapshot\"'})>Original</button><button type='button' class='json-copy util-button' data-copy='mut' title='Copy displayed JSON' aria-label='Copy JSON'>Copy</button></div><pre class='json-block' data-mode='mut' data-loaded='false'></pre><pre class='json-block' data-mode='orig' style='display:none;' data-loaded='false'></pre></details></td>"
      }
      elseif ($col -eq 'Name' -and $p.PolicyId) {
        $plink = $LinkURL + [uri]::EscapeDataString([string]$p.PolicyId)
        $rawName = [string]$raw
        $safe = [System.Web.HttpUtility]::HtmlEncode($rawName)
        # Get status for label
        $statusVal = [string]$p.Status
        # Generate status label
        $statusLabelClass = ''
        $statusLabelText = ''
        switch -Regex ($statusVal) {
          '^enabled$' {
            $statusLabelClass = 'status-label status-enabled'
            $statusLabelText = 'Enabled'
            break
          }
          'enabledForReportingButNotEnforced' {
            $statusLabelClass = 'status-label status-report'
            $statusLabelText = 'Report'
            break
          }
          'report only' {
            $statusLabelClass = 'status-label status-report'
            $statusLabelText = 'Report'
            break
          }
          '^disabled$' {
            $statusLabelClass = 'status-label status-disabled'
            $statusLabelText = 'Disabled'
            break
          }
          default {
            $statusLabelClass = 'status-label'
            $statusLabelText = $statusVal
          }
        }
        $safeStatus = [System.Web.HttpUtility]::HtmlEncode($statusVal)
        $safeStatusLabel = [System.Web.HttpUtility]::HtmlEncode($statusLabelText)
        $statusLabel = "<span class='$statusLabelClass' title='$safeStatus'>$safeStatusLabel</span>"
        
        $nameGuidance = ''
        $hints = @($NamingHintIndex[$p.PolicyId])
        if ($hints.Count) {
          $hintItems = ($hints | ForEach-Object { "<li>$([System.Web.HttpUtility]::HtmlEncode($_))</li>" }) -join ''
          $nameGuidance = "<details class='naming-guidance'><summary>$($hints.Count) naming hint(s)</summary><p>Potential name/configuration mismatch; review intent before changing anything.</p><ul>$hintItems</ul></details>"
        }
        if ($p.'Recommended Name' -and $p.'Recommended Name' -ne $p.Name) {
          $recNameSafe = [System.Web.HttpUtility]::HtmlEncode([string]$p.'Recommended Name')
          $nameGuidance += "<div class='suggested-name' hidden>Suggested name (heuristic): $recNameSafe</div>"
        }
        
        $rowTds += "<td class='sticky-name name-col'>$statusLabel<div class='name-content'><span class='policy-name'>$safe</span><a class='policy-link' href='$plink' target='_blank' aria-label='Open policy in new tab' title='Open policy in new tab'><svg class='icon-ext' viewBox='0 0 24 24' focusable='false' aria-hidden='true'><path d='M14 3h7v7m0-7L10 14' stroke-linecap='round' stroke-linejoin='round'/><path d='M21 21H3V3h7' stroke-linecap='round' stroke-linejoin='round'/></svg><span class='visually-hidden'></span></a>$nameGuidance</div></td>"
      }
      elseif ($col -eq 'Description') {
        if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
          $rowTds += "<td class='desc-cell'><span class='placeholder' aria-label='None'>—</span></td>"
        }
        else {
          $safe = [System.Web.HttpUtility]::HtmlEncode([string]$raw)
          $descId = 'desc-' + [guid]::NewGuid().ToString('N')
          $rowTds += "<td class='desc-cell'><div class='desc-text' id='$descId'>$safe</div><button type='button' class='desc-expand' aria-label='Expand description' aria-expanded='false' aria-controls='$descId'>More</button></td>"
        }
      }
      elseif ($col -eq 'Block') {
        if ($raw -eq 'True' -or $raw -eq $true) {
          $rowTds += "<td class='$idHiddenClass'><span class='pill-block' aria-label='Block'>BLOCK</span></td>" 
        }
        else {
          $rowTds += "<td class='$idHiddenClass'><span class='pill-grant' title='No Block grant in this policy. This does not grant access or override other policies.'>NON-BLOCKING</span></td>"
        }
      }
      elseif ($AssignmentTokenIndex[$p.PolicyId].ContainsKey($col) -and $AssignmentTokenIndex[$p.PolicyId][$col].Count) {
        $tokensHtml = foreach ($token in $AssignmentTokenIndex[$p.PolicyId][$col]) {
          $tokenType = [System.Web.HttpUtility]::HtmlEncode($token.Type)
          $tokenId = [System.Web.HttpUtility]::HtmlEncode($token.Id.ToLowerInvariant())
          $tokenLabel = [System.Web.HttpUtility]::HtmlEncode("$($token.Type): $($token.Label)")
          $tokenTitle = [System.Web.HttpUtility]::HtmlEncode("$($token.Direction) $($token.Type): $($token.Id)")
          "<button type='button' class='entity-token' data-entity-type='$tokenType' data-entity-id='$tokenId' data-direction='$($token.Direction)' title='$tokenTitle' aria-pressed='false'>$tokenLabel</button>"
        }
        $rowTds += "<td class='$idHiddenClass'>$($tokensHtml -join '')</td>"
      }
      elseif ($boolColumns -contains $col) {
        if ($raw -eq 'True' -or $raw -eq $true) {
          $rowTds += "<td class='bool-col$idHiddenClass'><span class='bool-yes' aria-label='True'>✔</span></td>" 
        }
        else {
          $rowTds += "<td class='bool-col$idHiddenClass'><span class='bool-no' aria-label='False'>—</span></td>" 
        }
      }
      else {
        if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
          $rowTds += "<td class='$idHiddenClass'><span class='placeholder' aria-label='None'>—</span></td>" 
        }
        else {
          $rawString = [string]$raw
          $safe = [System.Web.HttpUtility]::HtmlEncode($rawString)
          
          # Check if this is an assignment column that should be truncatable
          $truncatableColumns = @('Included Users', 'Excluded Users', 'Included Applications', 'Excluded Applications', 'Included Locations', 'Excluded Locations')
          if ($truncatableColumns -contains $col -and $rawString.Length -gt 100) {
            $assignId = 'assign-' + [guid]::NewGuid().ToString('N')
            $rowTds += "<td class='assignment-cell $idHiddenClass'><div class='assignment-text' id='$assignId'>$safe</div><button type='button' class='assignment-expand' aria-label='Expand' aria-expanded='false' aria-controls='$assignId'>More</button></td>"
          }
          else {
            $rowTds += "<td class='$idHiddenClass'>$safe</td>"
          }
        }
      }
      if ($col -notin @('Name', 'RawJson')) {
        $compareValue = [string]$raw
        if ($boolColumns -contains $col) {
          $compareValue = if ($raw -eq 'True' -or $raw -eq $true) { 'True' } else { 'False' }
        }
        elseif ($AssignmentTokenIndex[$p.PolicyId].ContainsKey($col)) {
          $keys = @($AssignmentTokenIndex[$p.PolicyId][$col] | ForEach-Object {
            "$($_.Type):$($_.Id.ToLowerInvariant())"
          } | Sort-Object -Unique)
          $compareValue = ConvertTo-Json -InputObject $keys -Compress
        }
        $safeValue = [System.Web.HttpUtility]::HtmlEncode($compareValue)
        $safeColumn = [System.Web.HttpUtility]::HtmlEncode($col)
        $rowTds[-1] = $rowTds[-1].Insert(4, "data-compare-field='$safeColumn' data-compare-value='$safeValue' ")
        $compareButton = "<button type='button' class='compare-cell' aria-label='Compare $safeColumn values' aria-pressed='false'>Compare</button>"
        $rowTds[-1] = $rowTds[-1] -replace '</td>$', "$compareButton</td>"
      }
    }
    $deployment = switch ($p.Status) { 'enabled' { 'enabled' } 'report only' { 'report' } 'disabled' { 'disabled' } default { 'unknown' } }
    $table += "<tr data-deployment='$deployment'>" + ($rowTds -join '') + '</tr>'
  }
  $table += '</tbody></table></div><p id="policy-empty" class="summary-note" hidden>No policies match the current search, deployment-state and comparison filters.</p>'
  # (Summary panel already populated above)
  # Append policies table (inside policies panel) exactly once, then close policies panel
  $HtmlParts += ($table -join "`n")
  $HtmlParts += '</div>'  # close policy-export / panel-policies
  # Append recommendations panel as a sibling (was previously nested causing it to be hidden and table omitted)
  $HtmlParts += $SecurityCheck
  $HtmlParts += Get-NamedLocationHtmlFragment -Inventory $NamedLocationInventory

  # Escape HTML-significant characters so tenant strings cannot terminate the script element.
  $policyDataJson = ConvertTo-Json -InputObject @($CAExport | Select-Object * -ExcludeProperty RawJson) -Depth 100 -Compress -EscapeHandling EscapeHtml -WarningAction Stop
  $rawIndexJson = ConvertTo-Json -InputObject @($CAPolicy) -Depth 100 -Compress -EscapeHandling EscapeHtml -WarningAction Stop
  $fallbackToggle = @"
<script>
window.CAExportData = $policyDataJson;
// Build RawPolicyDataIndex keyed by id for O(1) lookup during lazy load
window.RawPolicyDataIndex = (function(){
  var arr = $rawIndexJson;
  var idx = Object.create(null);
  for(var i=0;i<arr.length;i++){ var it=arr[i]; if(it && it.id){ idx[it.id]=it; } }
  return idx;
})();
// Accessible report navigation and interactions
document.addEventListener('DOMContentLoaded', function(){
  var tabs = Array.prototype.slice.call(document.querySelectorAll('.view-toggle-group [role=tab]'));
  var hasRecommendations = !!document.getElementById('tab-recommendations');
  var panels = {};
  tabs.forEach(function(t){ panels[t.id] = document.getElementById(t.getAttribute('aria-controls')); });
  // Explicit panel element references for later logic (table lookup, etc.)
  var policies = panels['tab-policies'];
  var summary  = panels['tab-summary'];
  var recs     = hasRecommendations ? panels['tab-recommendations'] : null;
  var searchInput = document.getElementById('policy-search');
  var searchClear = document.getElementById('policy-search-clear');
  var backToTop = document.getElementById('back-to-top');
  var recToggle = document.getElementById('toggle-recs');
  var boolModeToggle = document.getElementById('bool-mode-toggle');
  var comparison = null;
  var matchOnly = document.getElementById('match-only');
  var clearComparison = document.getElementById('clear-comparison');
  var comparisonStatus = document.getElementById('comparison-status');
  var suggestedNames = document.getElementById('show-suggested-names');
  suggestedNames.addEventListener('change', function(){
    policies.querySelectorAll('.suggested-name').forEach(function(name){ name.hidden = !suggestedNames.checked; });
    filterPolicies();
  });

  function updateComparison(){
    var counts = { include:0, exclude:0, enabled:0, report:0, disabled:0, unknown:0, total:0 };
    table.querySelectorAll('tbody tr').forEach(function(row){
      var included = false, excluded = false, matched = false;
      row.querySelectorAll('.entity-token').forEach(function(token){
        var same = !!(comparison && comparison.kind === 'entity' &&
          token.dataset.entityType === comparison.type && token.dataset.entityId === comparison.id);
        token.classList.toggle('match-include', same && token.dataset.direction === 'include');
        token.classList.toggle('match-exclude', same && token.dataset.direction === 'exclude');
        token.setAttribute('aria-pressed', same ? 'true' : 'false');
        if(same){ matched = true; if(token.dataset.direction === 'include'){ included = true; } else { excluded = true; } }
      });
      row.querySelectorAll('td[data-compare-field]').forEach(function(cell){
        var same = !!(comparison && comparison.kind === 'cell' &&
          cell.dataset.compareField === comparison.field && cell.dataset.compareValue === comparison.value);
        cell.classList.toggle('cell-match', same);
        cell.querySelector('.compare-cell').setAttribute('aria-pressed', same ? 'true' : 'false');
        if(same){ matched = true; }
      });
      row.dataset.matches = matched ? 'true' : 'false';
      if(included){ counts.include++; }
      if(excluded){ counts.exclude++; }
      if(matched){ counts.total++; counts[row.dataset.deployment]++; }
    });
    matchOnly.disabled = !comparison;
    clearComparison.disabled = !comparison;
    if(!comparison){
      matchOnly.checked = false;
      comparisonStatus.textContent = 'Click an assignment to match its type and original ID across includes/excludes, or Compare to match an exact value in the same field.';
    } else {
      comparisonStatus.textContent = comparison.label + ' | ' + counts.total + ' matching policies in this export' +
        (comparison.kind === 'entity' ? ' | Included: ' + counts.include + '; Excluded: ' + counts.exclude + ' (a policy can appear in both)' : '') +
        ' | Enabled: ' + counts.enabled + '; Report-only: ' + counts.report + '; Disabled: ' + counts.disabled + '; Unknown state: ' + counts.unknown +
        '. Totals precede search/deployment filters; visible count is shown above.';
    }
    filterPolicies();
  }
  matchOnly.addEventListener('change', filterPolicies);
  clearComparison.addEventListener('click', function(){ comparison = null; updateComparison(); });

  function updateHeaderHeight(){
    var groupHeaderRow = policies ? policies.querySelector('.group-header') : null;
    if(!groupHeaderRow) return;
    var height = groupHeaderRow.getBoundingClientRect().height;
    if(height > 0){ groupHeaderRow.closest('table').style.setProperty('--group-header-height', height + 'px'); }
  }

  function activate(tabId){
    if(!panels[tabId]){ if(window.console) console.warn('Activate called for unknown tabId', tabId); }
    tabs.forEach(function(t){
      var on = (t.id === tabId);
      t.classList.toggle('active', on);
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      t.tabIndex = on ? 0 : -1;
      var panel = panels[t.id];
      if(panel){
        panel.style.display = on ? 'block' : 'none';
        panel.setAttribute('aria-hidden', on ? 'false':'true');
      }
    });
    if(tabId === 'tab-policies'){ updateHeaderHeight(); }
    if(window.console){ console.log('Tab activated:', tabId, 'summary visible?', summary && summary.style.display, 'policies visible?', policies && policies.style.display); }
    if(searchInput){
      var isPolicies = (tabId === 'tab-policies');
      searchInput.disabled = !isPolicies;
      searchInput.title = isPolicies ? 'Search policies' : 'Search only available in Policy Details view';
    }
  var hash = '#policies';
  if(hasRecommendations && tabId === 'tab-recommendations') hash = '#recommendations';
    if(tabId === 'tab-summary') hash = '#summary';
    if(tabId === 'tab-locations') hash = '#locations';
  try { history.replaceState(null, '', hash); } catch(e){}
  try { localStorage.setItem('caexport.activeTab', tabId); } catch(e){}
    window.scrollTo({top:0});
  }

  // Click handling
  tabs.forEach(function(t){ t.addEventListener('click', function(){ activate(t.id); }); });

  // Keyboard navigation (Left/Right/Home/End)
  function tabKeyHandler(e){
    var key = e.key;
    var currentIndex = tabs.indexOf(document.activeElement);
    if(currentIndex === -1) return;
    if(key === 'ArrowRight'){ e.preventDefault(); var next = (currentIndex+1) % tabs.length; tabs[next].focus(); activate(tabs[next].id); }
    else if(key === 'ArrowLeft'){ e.preventDefault(); var prev = (currentIndex-1+tabs.length) % tabs.length; tabs[prev].focus(); activate(tabs[prev].id); }
    else if(key === 'Home'){ e.preventDefault(); tabs[0].focus(); activate(tabs[0].id); }
    else if(key === 'End'){ e.preventDefault(); tabs[tabs.length-1].focus(); activate(tabs[tabs.length-1].id); }
  }
  tabs.forEach(function(t){ t.addEventListener('keydown', tabKeyHandler); });

  // Hash routing
  var storedTab = null; try { storedTab = localStorage.getItem('caexport.activeTab'); } catch(e){}
  if(window.location.hash){
    var h = window.location.hash.toLowerCase();
    if(hasRecommendations && h.includes('recommend')) activate('tab-recommendations');
    else if(h.includes('locations')) activate('tab-locations');
    else if(h.includes('summary')) activate('tab-summary');
    else activate('tab-policies');
  } else if(storedTab && panels[storedTab]) { activate(storedTab); } else { activate('tab-policies'); }

  // Table interactions (guarded)
  var table = policies ? policies.querySelector('table') : null;
  if(table){
    var groupHeaderRow = table.querySelector('.group-header');
    new ResizeObserver(updateHeaderHeight).observe(groupHeaderRow);
    updateHeaderHeight();
    table.addEventListener('click', function(e){
      var entity = e.target.closest('.entity-token');
      if(entity){
        comparison = { kind:'entity', type:entity.dataset.entityType, id:entity.dataset.entityId, label:entity.textContent + ' [' + entity.title + ']' };
        updateComparison();
        return;
      }
      var compareButton = e.target.closest('.compare-cell');
      if(compareButton){
        var cell = compareButton.closest('td');
        comparison = { kind:'cell', field:cell.dataset.compareField, value:cell.dataset.compareValue, label:'Exact field match: ' + cell.dataset.compareField + ' = ' + (cell.dataset.compareValue || '(empty)') };
        updateComparison();
        return;
      }
      // Lazy JSON population on first expand
      var summary = e.target.closest('summary');
      if(summary){
        var details = summary.parentElement;
        if(details && details.classList.contains('lazy-json')){
          var opened = details.hasAttribute('open');
          // Delay execution until the element is actually opened (after default toggle)
          setTimeout(function(){
            if(details.hasAttribute('open')){
              var mutPre = details.querySelector("pre.json-block[data-mode='mut']");
              if(mutPre && mutPre.getAttribute('data-loaded')==='false'){
                var pid = details.getAttribute('data-policyid');
                try {
                  var policyRow = window.CAExportData ? window.CAExportData.find(function(x){ return x.PolicyId===pid; }) : null;
                  if(policyRow){
                    var shallow = Object.assign({}, policyRow);
                    delete shallow.RawJson;
                    mutPre.textContent = JSON.stringify(shallow, null, 2);
                    mutPre.setAttribute('data-loaded','true');
                  }
                  var hasOrig = details.getAttribute('data-has-orig')==='true';
                  if(hasOrig){
                    var origPre = details.querySelector("pre.json-block[data-mode='orig']");
                    if(origPre && origPre.getAttribute('data-loaded')==='false' && window.RawPolicyDataIndex){
                      var origObj = window.RawPolicyDataIndex[pid];
                      if(origObj){ origPre.textContent = JSON.stringify(origObj, null, 2); origPre.setAttribute('data-loaded','true'); }
                    }
                  }
                } catch(ex){ console.warn('Lazy JSON load failed', ex); }
              }
            }
          }, 0);
        }
      }
      // JSON toggle buttons
      var jsonBtn = e.target.closest('button.json-btn');
      if(jsonBtn){
        var td = jsonBtn.closest('td');
        var btns = td.querySelectorAll('button.json-btn');
        for(var i=0;i<btns.length;i++){ btns[i].classList.remove('active'); btns[i].setAttribute('aria-pressed','false'); }
        jsonBtn.classList.add('active'); jsonBtn.setAttribute('aria-pressed','true');
        var mode = jsonBtn.getAttribute('data-mode');
        var blocks = td.querySelectorAll('pre.json-block');
        for(var b=0;b<blocks.length;b++){ blocks[b].style.display = (blocks[b].getAttribute('data-mode') === mode) ? 'block' : 'none'; }
        // Update copy button data attribute to reflect current mode
        var copyBtn = td.querySelector('button.json-copy');
        if(copyBtn){ copyBtn.setAttribute('data-copy', mode); }
        return; // don't fall through to row select logic for the toggle click
      }
      // JSON copy button
      var jsonCopy = e.target.closest('button.json-copy');
      if(jsonCopy){
        var td = jsonCopy.closest('td');
        var mode = jsonCopy.getAttribute('data-copy') || 'mut';
        var block = td.querySelector("pre.json-block[data-mode='"+mode+"']");
        if(block){
          var text = block.textContent || '';
          navigator.clipboard.writeText(text).then(function(){
            var lr = document.getElementById('live-region');
            if(lr){ lr.textContent = 'JSON copied to clipboard ('+mode+').'; }
            jsonCopy.textContent = 'Copied';
            setTimeout(function(){ jsonCopy.textContent='Copy'; }, 1500);
          }).catch(function(){
            var lr = document.getElementById('live-region');
            if(lr){ lr.textContent = 'Copy failed'; }
          });
        }
        return;
      }
      if(e.target.closest('a, button, input, select, summary, details')) return;
      var tr = e.target.closest('tr');
      if(tr && tr.parentElement.tagName !== 'THEAD'){
        tr.classList.toggle('selected');
      }
    });

  }

  // Status filter functionality
  var statusFilter = document.getElementById('status-filter');
  if(statusFilter && table){
    function applyStatusFilter(){
      filterPolicies();
    }
    statusFilter.addEventListener('change', applyStatusFilter);
    // Restore filter state
    try {
      var storedFilter = localStorage.getItem('caexport.statusFilter');
      if(storedFilter){ statusFilter.value = storedFilter; applyStatusFilter(); }
    } catch(e){}
    // Save filter state on change
    statusFilter.addEventListener('change', function(){
      try { localStorage.setItem('caexport.statusFilter', statusFilter.value); } catch(e){}
    });
  }

  // Recommendation expand/collapse (simple show/hide of recommendation cards)
  if(recToggle){
    recToggle.addEventListener('click', function(){
      var recPanel = document.getElementById('panel-recommendations');
      if(!recPanel) return;
      var items = recPanel.querySelectorAll('.recommendation');
      var collapsing = recToggle.getAttribute('data-collapsed') !== 'true';
      items.forEach(function(el){ el.style.display = collapsing ? 'none' : ''; });
      recToggle.setAttribute('data-collapsed', collapsing ? 'true' : 'false');
      recToggle.textContent = collapsing ? 'Show Recs' : 'Hide Recs';
      recToggle.classList.toggle('active', !collapsing);
    });
  }

  // Keep navigation by purpose separate from optional assessment-state filtering.
  (function(){
    var recPanel = document.getElementById('panel-recommendations');
    if(!recPanel) return;
    var buttons = Array.prototype.slice.call(recPanel.querySelectorAll('.rec-filter-group button'));
    var assessment = document.getElementById('rec-assessment-filter');
    var cards = Array.prototype.slice.call(recPanel.querySelectorAll('details.recommendation-card'));
    var resultCount = document.getElementById('rec-result-count');
    var empty = document.getElementById('rec-filter-empty');
    var views = buttons.map(function(b){ return b.dataset.filter; });
    var states = Array.prototype.map.call(assessment.options, function(o){ return o.value; });
    var view = 'fail';
    var storedView = null;
    var storedAssessment = null;
    try {
      storedView = localStorage.getItem('caexport.recFilter');
      storedAssessment = localStorage.getItem('caexport.recAssessment');
    } catch(e){}
    if(views.indexOf(storedView)>-1){
      view = storedView;
    } else if(storedView === 'pass' || (storedView !== 'all' && states.indexOf(storedView)>-1)){
      view = 'all';
      storedAssessment = storedView === 'pass' ? 'enforced' : storedView;
    }
    assessment.value = states.indexOf(storedAssessment)>-1 ? storedAssessment : 'all';

    function applyRecommendationFilters(persist){
      var selectedButton = null;
      buttons.forEach(function(b){
        var active = b.dataset.filter === view;
        b.classList.toggle('active', active);
        b.setAttribute('aria-pressed', active ? 'true' : 'false');
        if(active){ selectedButton = b; }
      });
      var visibleCount = 0;
      cards.forEach(function(card){
        var inView = view === 'all' || (view === 'fail' && card.dataset.attention === 'true') ||
          (view === 'templates' && card.dataset.kind === 'template') ||
          (view === 'optional' && ['review','emergency'].indexOf(card.dataset.kind)>-1);
        var inAssessment = assessment.value === 'all' || card.dataset.assessment === assessment.value;
        var show = inView && inAssessment;
        card.style.display = show ? '' : 'none';
        if(show){ visibleCount++; }
      });
      resultCount.textContent = 'Showing ' + visibleCount + ' of ' + cards.length + ' assessments: ' +
        selectedButton.textContent + ' / ' + assessment.options[assessment.selectedIndex].textContent;
      empty.hidden = visibleCount > 0;
      if(persist){
        try {
          localStorage.setItem('caexport.recFilter', view);
          localStorage.setItem('caexport.recAssessment', assessment.value);
        } catch(e){}
      }
    }
    buttons.forEach(function(b){
      b.addEventListener('click', function(){ view = b.dataset.filter; applyRecommendationFilters(true); });
    });
    assessment.addEventListener('change', function(){ applyRecommendationFilters(true); });
    applyRecommendationFilters(false);
  })();

  // Boolean icon/text toggle with persistence
  // Stored key: caexport.boolMode -> 'icon' (default) or 'text'
  if(boolModeToggle && table){
    var storedBoolMode = null; try { storedBoolMode = localStorage.getItem('caexport.boolMode'); } catch(e){}
    function applyBoolMode(mode){
      var toIcons = (mode === 'icon');
      boolModeToggle.setAttribute('data-mode', mode);
      boolModeToggle.textContent = 'Boolean: ' + (toIcons ? 'Icons' : 'Text');
      table.querySelectorAll('td.bool-col span.bool-yes, td.bool-col span.bool-no').forEach(function(span){
        span.textContent = span.classList.contains('bool-yes') ? (toIcons?'✔':'True') : (toIcons?'—':'False');
      });
    }
    applyBoolMode(storedBoolMode && ['icon','text'].indexOf(storedBoolMode)>-1 ? storedBoolMode : 'icon');
    boolModeToggle.addEventListener('click', function(){
      var current = boolModeToggle.getAttribute('data-mode');
      var next = current === 'icon' ? 'text' : 'icon';
      applyBoolMode(next);
      try { localStorage.setItem('caexport.boolMode', next); } catch(e){}
    });
  }

  // Description expand/collapse (aria-expanded + aria-controls for accessibility)
  if(table){
    table.addEventListener('click', function(e){
      var btn = e.target.closest('.desc-expand');
      if(!btn) return;
      var cell = btn.closest('.desc-cell');
      var expanded = cell.classList.toggle('expanded');
      btn.textContent = expanded ? 'Less' : 'More';
  btn.setAttribute('aria-label', expanded ? 'Collapse description' : 'Expand description');
  btn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    });
  }

  // Assignment cell expand/collapse
  if(table){
    table.addEventListener('click', function(e){
      var btn = e.target.closest('.assignment-expand');
      if(!btn) return;
      var cell = btn.closest('.assignment-cell');
      var expanded = cell.classList.toggle('expanded');
      btn.textContent = expanded ? 'Less' : 'More';
      btn.setAttribute('aria-label', expanded ? 'Collapse' : 'Expand');
      btn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    });
  }

  // Policy search filter
  function filterPolicies(){
    if(!table) return; // nothing to do
    var q = (searchInput.value || '').toLowerCase();
    if(searchInput.parentElement){
      if(q.length>0){ searchInput.parentElement.classList.add('has-value'); }
      else { searchInput.parentElement.classList.remove('has-value'); }
    }
    var rows = table.querySelectorAll('tbody tr');
    var status = statusFilter ? statusFilter.value : 'all';
    var visibleCount = 0;
    for(var i=0;i<rows.length;i++){
      var row = rows[i];
      var label = row.querySelector('.status-label');
      var show = status === 'all' || !!(label && label.classList.contains('status-' + status));
      if(show && comparison && matchOnly.checked){ show = row.dataset.matches === 'true'; }
      if(show && q.length>0){
        var cells = row.children;
        var match = false;
        for(var c=0;c<cells.length;c++){
          var searchable = cells[c].cloneNode(true);
          searchable.querySelectorAll('.compare-cell, .suggested-name[hidden], .lazy-json').forEach(function(control){ control.remove(); });
          if(/^block(?:s|ed|ing)?$/.test(q)){
            searchable.querySelectorAll('.pill-grant').forEach(function(control){ control.remove(); });
          }
          if(searchable.textContent && searchable.textContent.toLowerCase().indexOf(q) > -1){ match = true; break; }
        }
        show = match;
      }
      row.style.display = show ? '' : 'none';
      if(show){ visibleCount++; }
    }
    var countText = visibleCount + ' of ' + rows.length + ' policies';
    var counter = document.getElementById('policy-count');
    if(counter){ counter.textContent = countText; }
    var empty = document.getElementById('policy-empty');
    if(empty){ empty.hidden = visibleCount > 0; }
    var liveRegion = document.getElementById('live-region');
    if(liveRegion){ liveRegion.textContent = 'Showing ' + countText; }
  }
  if(searchInput){
    searchInput.addEventListener('input', filterPolicies);
    searchInput.addEventListener('keydown', function(e){ if(e.key==='Escape'){ searchInput.value=''; filterPolicies(); searchInput.blur(); } });
  }
  if(searchClear){
    searchClear.addEventListener('click', function(){ searchInput.value=''; filterPolicies(); searchInput.focus(); });
  }
  filterPolicies();

  // Back to top visibility
  function updateBackToTop(){
    if(!backToTop) return;
    if(window.scrollY > 300){ backToTop.style.display='block'; } else { backToTop.style.display='none'; }
  }
  window.addEventListener('scroll', updateBackToTop);
  if(backToTop){
    backToTop.addEventListener('click', function(){ window.scrollTo({top:0, behavior:'smooth'}); });
  }
  updateBackToTop();

  // Ensure search state matches initial view (handled in activate but safeguard)
  if(searchInput && searchInput.disabled && window.location.hash.toLowerCase().indexOf('policies')>-1){ searchInput.disabled=false; }
});
</script>
"@
  $fullPage = $htmlDoc + $HtmlParts + $fallbackToggle + '</body></html>'
  $fullPage | Out-File $Launch -Encoding UTF8
  if (-not $NoBrowser) {
    if ($IsWindows) {
      Start-Process $Launch
    }
    elseif ($IsMacOS) {
      & open $Launch
    }
    elseif ($IsLinux) {
      & xdg-open $Launch
    }
    else {
      Write-Info "HTML report saved to: $Launch (auto-open not supported on this platform)"
    }
  }
}
if ($JsonExport) {
  Write-Info 'Saving to File: JSON (enriched policies)'
  $LaunchJson = Join-Path -Path $ExportLocation -ChildPath $JsonFileName
  try {
    $CAExport | ConvertTo-Json -Depth 12 | Out-File $LaunchJson
    Write-Info "JSON (enriched) saved: $LaunchJson"
  }
  catch {
    Write-Warn "Failed to save enriched JSON: $_" 
  }
}
if ($CsvExport) {
  Write-Info 'Saving to File: CSV (one row per policy)'
  $LaunchCsv = Join-Path -Path $ExportLocation -ChildPath $CsvFileName
  $defaultColumns = @(
    'Name', 'PolicyId', 'Status', 'Naming Hints', 'Modified', 'Created',
    'Included Users', 'Excluded Users',
    'Included Applications', 'Excluded Applications',
    'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations', 'User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows', 'Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA',
    'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session'
  )
  $chosenColumns = if ($CsvColumns) {
    $CsvColumns 
  }
  else {
    $defaultColumns 
  }
  $available = if ($CAExport.Count -gt 0) {
    $CAExport[0].PSObject.Properties.Name 
  }
  else {
    @() 
  }
  $missingCols = @(); $finalCols = @()
  foreach ($c in $chosenColumns) {
    if ($available -contains $c) {
      $finalCols += $c 
    }
    else {
      $missingCols += $c 
    } 
  }
  if ($missingCols) {
    Write-Warn "Ignoring unknown CSV columns: $($missingCols -join ', ')" 
  }
  $exportSet = if ($finalCols) {
    $CAExport | Select-Object -Property $finalCols 
  }
  else {
    $null 
  }
  if (-not $exportSet) {
    Write-Warn 'No CAExport data found; exporting raw policies.'; $CAPolicy | Select-Object * | Export-Csv -Path $LaunchCsv
  }
  else {
    $exportSet | Export-Csv -Path $LaunchCsv
  }
  Write-Info "CSV saved: $LaunchCsv"
}
if ($CsvPivotExport) {
  Write-Info 'Saving to File: Pivot CSV'
  $LaunchCsvPivot = Join-Path -Path $ExportLocation -ChildPath $CsvPivotFileName
  $pivotToExport = $pivot
  if (-not $pivotToExport -or $pivotToExport.Count -eq 0) {
    Write-Warn 'Pivot empty; skipping pivot CSV.' 
  }
  else {
    try {
      $pivotToExport | Export-Csv -Path $LaunchCsvPivot
      Write-Info "Pivot CSV saved: $LaunchCsvPivot"
    }
    catch {
      Write-Warn "Failed to write pivot CSV: $($_.Exception.Message)"
    }
  }
}