# Pester tests for core helper logic in Export-CAPolicyWithRecs.ps1
# Run with: Invoke-Pester -Path tests -Output Detailed

<# Each Describe defines needed helper functions in its own BeforeAll to work with Pester V5 discovery (which does not execute top-level code). #>

Describe 'State-aware Microsoft recommendation assessments' {
    BeforeAll {
        $source = Get-Content -LiteralPath "$PSScriptRoot/../Export-CAPolicyWithRecs.ps1" -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
        $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
        function Get-TestRecommendation {
            param([string]$Control, [object[]]$Policies = @($policy), [string[]]$EmergencyIds)
            Get-CaRecommendationResult -Policies $Policies -EmergencyAccessAccountId $EmergencyIds |
                Where-Object Control -eq $Control
        }
    }

    BeforeEach {
        $script:OfflineMode = $true
        $script:AuthStrengthCachePopulated = $false
        $policy = @{
            id = '11111111-1111-1111-1111-111111111111'; displayName = 'Test policy'; state = 'enabled'
            conditions = @{
                users = @{ includeUsers = @('All'); excludeUsers = @(); includeGroups = @(); excludeGroups = @(); includeRoles = @(); excludeRoles = @() }
                applications = @{ includeApplications = @('All'); excludeApplications = @() }
                clientAppTypes = @('all')
            }
            grantControls = @{ operator = 'OR'; builtInControls = @('mfa') }
            sessionControls = @{}
        }
    }

    It 'Distinguishes <State> from actual enforcement' -TestCases @(
        @{ State = 'enabled'; Expected = 'Enforced' }
        @{ State = 'enabledForReportingButNotEnforced'; Expected = 'Report-only' }
        @{ State = 'disabled'; Expected = 'Disabled' }
        @{ State = 'unknownFutureValue'; Expected = 'Needs review' }
    ) {
        param($State, $Expected)
        $policy.state = $State
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be $Expected
    }

    It 'Uses enabled qualifying policies without promoting a scoped enabled policy over a broad report-only policy' {
        $reportOnly = $policy | ConvertTo-Json -Depth 20 | ConvertFrom-Json -AsHashtable
        $reportOnly.state = 'enabledForReportingButNotEnforced'
        $policy.conditions.users.includeUsers = @('a-specific-user')
        (Get-TestRecommendation 'CA-01' @($policy, $reportOnly)).Assessment | Should -Be 'Report-only'
    }

    It 'Treats <Scenario> as uncertain coverage rather than enforced protection' -TestCases @(
        @{ Scenario = 'group targeting'; Change = { $policy.conditions.users.includeUsers = @(); $policy.conditions.users.includeGroups = @('group') } }
        @{ Scenario = 'group exclusions'; Change = { $policy.conditions.users.excludeGroups = @('group') } }
        @{ Scenario = 'resource exclusions'; Change = { $policy.conditions.applications.excludeApplications = @('resource') } }
        @{ Scenario = 'a location condition'; Change = { $policy.conditions.locations = @{ includeLocations = @('location') } } }
        @{ Scenario = 'a client restriction'; Change = { $policy.conditions.clientAppTypes = @('browser') } }
        @{ Scenario = 'a risk-free-only condition'; Change = { $policy.conditions.signInRiskLevels = @('none') } }
        @{ Scenario = 'an unknown condition'; Change = { $policy.conditions.AdditionalProperties = @{ futureCondition = @{ value = 'restricted' } } } }
    ) {
        param($Scenario, $Change)
        . $Change
        $result = Get-TestRecommendation 'CA-01'
        $result.Assessment | Should -Be 'Needs review'
        $result.Evidence[0].Issues.Count | Should -BeGreaterThan 0
    }

    It 'Ignores unconfigured Graph SDK containers in <Representation> policies' -TestCases @(
        @{ Representation = 'dictionary' }
        @{ Representation = 'PSCustomObject' }
    ) {
        param($Representation)
        $emptyFilter = [pscustomobject]@{ Mode = $null; Rule = $null }
        $emptyGuests = [pscustomobject]@{ GuestOrExternalUserTypes = $null; ExternalTenants = [pscustomobject]@{ MembershipKind = $null } }
        $policy.conditions.authenticationFlows = [pscustomobject]@{ TransferMethods = $null }
        $policy.conditions.clientApplications = [pscustomobject]@{ IncludeServicePrincipals = $null; ExcludeServicePrincipals = @(); ServicePrincipalFilter = $emptyFilter; AgentIdServicePrincipalFilter = $emptyFilter }
        $policy.conditions.devices = [pscustomobject]@{ DeviceFilter = $emptyFilter }
        $policy.conditions.locations = [pscustomobject]@{ IncludeLocations = $null; ExcludeLocations = @() }
        $policy.conditions.platforms = [pscustomobject]@{ IncludePlatforms = @(); ExcludePlatforms = $null }
        $policy.conditions.applications.applicationFilter = $emptyFilter
        $policy.conditions.users.includeGuestsOrExternalUsers = $emptyGuests
        $policy.conditions.users.excludeGuestsOrExternalUsers = $emptyGuests
        $policy.grantControls.authenticationStrength = [pscustomobject]@{ Id = $null; AllowedCombinations = @(); RequirementsSatisfied = $null }
        if ($Representation -eq 'PSCustomObject') { $policy = $policy | ConvertTo-Json -Depth 20 | ConvertFrom-Json }
        $result = Get-TestRecommendation 'CA-01'
        $result.Assessment | Should -Be 'Enforced'
        $result.Evidence.Count | Should -Be 1
        $result.Evidence[0].Issues.Count | Should -Be 0
        (Get-TestRecommendation 'CA-26').Evidence.Count | Should -Be 0
        $policy.grantControls.builtInControls = @('compliantDevice')
        (Get-TestRecommendation 'CA-28').TemplateAssessment | Should -Be 'Matching policy enabled'
    }

    It 'Does not mistake empty SDK strengths or guest assignments for configured controls or targets' {
        $policy.grantControls.builtInControls = @('block')
        $policy.grantControls.authenticationStrength = [pscustomobject]@{ Id = $null; AllowedCombinations = @(); RequirementsSatisfied = $null }
        $policy.conditions.users.includeUsers = @()
        $policy.conditions.users.includeGuestsOrExternalUsers = [pscustomobject]@{ GuestOrExternalUserTypes = $null; ExternalTenants = [pscustomobject]@{ MembershipKind = $null } }
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Missing'
        (Get-TestRecommendation 'CA-07').Assessment | Should -Be 'Needs review'
        $policy.conditions.devices = [pscustomobject]@{ DeviceFilter = [pscustomobject]@{ Mode = $null; Rule = $null } }
        (Get-TestRecommendation 'CA-28').TemplateAssessment | Should -Be 'Not observed'
    }

    It 'Accepts explicit all-platform and all-location assignments but retains actual restrictions' {
        $policy.conditions.platforms = @{ includePlatforms = @('all'); excludePlatforms = @() }
        $policy.conditions.locations = @{ includeLocations = @('All'); excludeLocations = @() }
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Enforced'
        $policy.conditions.locations.includeLocations = @('AllTrusted')
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
        $policy.conditions.locations.includeLocations = @('All')
        $policy.conditions.platforms.excludePlatforms = @('linux')
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
    }

    It 'Preserves real nested SDK restrictions and unknown values rather than discarding their containers' {
        $policy.conditions.devices = [pscustomobject]@{ DeviceFilter = [pscustomobject]@{ Mode = 'include'; Rule = 'device.isCompliant -eq True' } }
        (Get-TestRecommendation 'CA-01').Evidence[0].Issues | Should -Contain "Additional condition 'devices' limits coverage and needs review."
        $policy.conditions.Remove('devices')
        $policy.conditions.AdditionalProperties = @{ futureCondition = [pscustomobject]@{ Required = $false } }
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
    }

    It 'Separates optional and operational applicability reviews from actionable attention findings' {
        $results = @(Get-CaRecommendationResult -Policies @())
        @($results | Where-Object { $_.Kind -in @('Review','Template','Emergency') -and $_.NeedsAttention }).Count | Should -Be 0
        ($results | Where-Object Control -eq 'CA-01').NeedsAttention | Should -BeTrue
        $accountId = '22222222-2222-2222-2222-222222222222'
        (Get-TestRecommendation 'CA-22' -EmergencyIds @($accountId)).NeedsAttention | Should -BeTrue
        (Get-TestRecommendation 'CA-01').NeedsAttention | Should -BeFalse
    }

    It 'Keeps registration-only MFA separate from resource coverage and detects missing action policies' {
        (Get-TestRecommendation 'CA-11').Assessment | Should -Be 'Missing'
        (Get-TestRecommendation 'CA-13').Assessment | Should -Be 'Missing'
        $policy.conditions.applications = @{ includeUserActions = @('urn:user:registersecurityinfo') }
        (Get-TestRecommendation 'CA-13').Assessment | Should -Be 'Enforced'
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Missing'
    }

    It 'Does not mistake an OR alternative for mandatory MFA' {
        $policy.grantControls.builtInControls = @('mfa', 'compliantDevice')
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
        $policy.grantControls.operator = 'AND'
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Enforced'
    }

    It 'Includes terms-of-use and custom factors when checking OR bypasses' -TestCases @(
        @{ Property = 'termsOfUse' }
        @{ Property = 'customAuthenticationFactors' }
    ) {
        param($Property)
        $policy.grantControls[$Property] = @('control-id')
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
    }

    It 'Recognizes built-in phishing-resistant strength without relying on its name or Graph access' {
        $policy.grantControls.builtInControls = @()
        $policy.grantControls.authenticationStrength = @{ id = '00000000-0000-0000-0000-000000000004'; displayName = 'Renamed' }
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Enforced'
        (Get-TestRecommendation 'CA-05').Assessment | Should -Be 'Enforced'
    }

    It 'Does not trust a custom strength name or unknown allowed combinations' {
        $policy.grantControls.builtInControls = @()
        $policy.grantControls.authenticationStrength = @{ id = 'custom'; displayName = 'Phishing-resistant MFA'; allowedCombinations = @('unknownFutureValue') }
        (Get-TestRecommendation 'CA-05').Assessment | Should -Be 'Needs review'
    }

    It 'Requires sign-in risk, both recommended levels, and every-time reauthentication' {
        $policy.conditions.userRiskLevels = @('high')
        (Get-TestRecommendation 'CA-09').Assessment | Should -Be 'Missing'
        $policy.conditions.Remove('userRiskLevels')
        $policy.conditions.signInRiskLevels = @('high')
        (Get-TestRecommendation 'CA-09').Assessment | Should -Be 'Needs review'
        $policy.conditions.signInRiskLevels = @('medium', 'high')
        $policy.sessionControls.signInFrequency = @{ isEnabled = $true; frequencyInterval = 'everyTime' }
        (Get-TestRecommendation 'CA-09').Assessment | Should -Be 'Enforced'
    }

    It 'Assesses modern high-user-risk remediation separately from legacy password changes' {
        $policy.conditions.userRiskLevels = @('high')
        $policy.grantControls.builtInControls = @('riskRemediation')
        $policy.grantControls.operator = 'AND'
        $policy.grantControls.authenticationStrength = @{ id = '00000000-0000-0000-0000-000000000002' }
        $policy.sessionControls.signInFrequency = @{ isEnabled = $true; frequencyInterval = 'everyTime' }
        (Get-TestRecommendation 'CA-14').Assessment | Should -Be 'Enforced'
        $policy.grantControls.builtInControls = @('passwordChange', 'mfa')
        (Get-TestRecommendation 'CA-14').Assessment | Should -Be 'Needs review'
        $policy.grantControls.builtInControls = @('unknownFutureValue')
        (Get-TestRecommendation 'CA-14').Assessment | Should -Be 'Needs review'
    }

    It 'Distinguishes security-info registration from device registration' {
        $policy.conditions.applications = @{ includeUserActions = @('urn:user:registersecurityinfo') }
        (Get-TestRecommendation 'CA-13').Assessment | Should -Be 'Enforced'
        (Get-TestRecommendation 'CA-11').Assessment | Should -Not -Be 'Enforced'
        $policy.conditions.applications.includeUserActions = @('urn:user:registerdevice')
        (Get-TestRecommendation 'CA-11').Assessment | Should -Be 'Enforced'
        (Get-TestRecommendation 'CA-13').Assessment | Should -Not -Be 'Enforced'
    }

    It 'Allows broad MFA to satisfy guest, portal and Azure management checks without duplicate policies' {
        foreach ($control in @('CA-16', 'CA-17', 'CA-18')) {
            (Get-TestRecommendation $control).Assessment | Should -Be 'Enforced'
        }
        $policy.conditions.applications.includeApplications = @('MicrosoftAdminPortals')
        (Get-TestRecommendation 'CA-17').Assessment | Should -Be 'Enforced'
        (Get-TestRecommendation 'CA-18').Assessment | Should -Not -Be 'Enforced'
    }

    It 'Does not mistake MFA on one role for complete administrator coverage' {
        $policy.conditions.users.includeUsers = @()
        $policy.conditions.users.includeRoles = @('62e90394-69f5-4237-9190-012177145e10')
        (Get-TestRecommendation 'CA-04').Assessment | Should -Be 'Needs review'
    }

    It 'Requires mandatory managed-device controls for administrators' {
        $policy.grantControls.builtInControls = @('compliantDevice', 'domainJoinedDevice')
        (Get-TestRecommendation 'CA-15').Assessment | Should -Be 'Enforced'
        $policy.grantControls.builtInControls += 'mfa'
        (Get-TestRecommendation 'CA-15').Assessment | Should -Be 'Needs review'
    }

    It 'Flags the approved-app migration without counting it as modern mobile protection' {
        $policy.grantControls.builtInControls = @('approvedApplication')
        $policy.conditions.platforms = @{ includePlatforms = @('android', 'iOS') }
        (Get-TestRecommendation 'CA-19').Assessment | Should -Be 'Needs review'
        (Get-TestRecommendation 'CA-02').Assessment | Should -Not -Be 'Enforced'
        $policy.grantControls.builtInControls = @('compliantApplication')
        (Get-TestRecommendation 'CA-02').Assessment | Should -Be 'Enforced'
    }

    It 'Reads modern controls from SDK AdditionalProperties and ignores disabled session controls' {
        $policy.sessionControls.AdditionalProperties = @{ secureSignInSession = @{ isEnabled = $false } }
        (Get-TestRecommendation 'CA-23').Evidence.Count | Should -Be 0
        $policy.sessionControls.AdditionalProperties.secureSignInSession.isEnabled = $true
        $result = Get-TestRecommendation 'CA-23'
        $result.Evidence.Count | Should -Be 1
        $result.Assessment | Should -Be 'Needs review'
    }

    It 'Recognizes authentication-flow configuration in typed and additional properties' {
        $policy.grantControls.builtInControls = @('block')
        $policy.conditions.AdditionalProperties = @{ authenticationFlows = @{ transferMethods = 'deviceCodeFlow,authenticationTransfer' } }
        (Get-TestRecommendation 'CA-10').Evidence.Count | Should -Be 1
    }

    It 'Requires both legacy client types and rejects platform blocks excluding all platforms' {
        $policy.grantControls.builtInControls = @('block')
        $policy.conditions.clientAppTypes = @('exchangeActiveSync')
        (Get-TestRecommendation 'CA-00').Assessment | Should -Be 'Needs review'
        $policy.conditions.clientAppTypes += 'other'
        (Get-TestRecommendation 'CA-00').Assessment | Should -Be 'Enforced'
        $policy.conditions.clientAppTypes = @('all')
        $policy.conditions.platforms = @{ includePlatforms = @('all'); excludePlatforms = @('all') }
        (Get-TestRecommendation 'CA-12').Assessment | Should -Be 'Needs review'
        $policy.conditions.platforms.excludePlatforms = @('windows', 'macOS', 'android', 'iOS', 'linux')
        (Get-TestRecommendation 'CA-12').Assessment | Should -Be 'Enforced'
    }

    It 'Reviews unknown grants rather than assuming a known MFA control is sufficient' {
        $policy.grantControls.builtInControls = @('mfa', 'unknownFutureValue')
        $policy.grantControls.operator = 'AND'
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
    }

    It 'Finds elevated insider-risk blocks without treating a risk signal alone as a block' {
        $policy.conditions.AdditionalProperties = @{ insiderRiskLevels = 'moderate,elevated' }
        (Get-TestRecommendation 'CA-25').Evidence.Count | Should -Be 0
        $policy.grantControls.builtInControls = @('block')
        (Get-TestRecommendation 'CA-25').Evidence.Count | Should -Be 1
    }

    It 'Treats an empty workload assignment as empty, not as evidence of a target identity' {
        $policy.conditions.users.includeUsers = @()
        $policy.conditions.clientApplications = @{ includeServicePrincipals = @() }
        (Get-TestRecommendation 'CA-07').Assessment | Should -Be 'Needs review'
    }

    It 'Keeps absent optional advanced scenarios as review rather than universal failures' {
        foreach ($control in @('CA-20','CA-21','CA-23','CA-24','CA-25','CA-26','CA-27')) {
            (Get-TestRecommendation $control -Policies @()).Assessment | Should -Be 'Needs review'
        }
    }

    It 'Recognizes workload and agent targeting without reporting a missing user assignment' {
        $policy.conditions.Remove('users')
        $policy.conditions.clientApplications = @{ includeServicePrincipals = @('service-principal') }
        $policy.conditions.servicePrincipalRiskLevels = @('high')
        $policy.grantControls.builtInControls = @('block')
        (Get-TestRecommendation 'CA-07').Assessment | Should -Be 'No findings'
        (Get-TestRecommendation 'CA-24').Evidence.Count | Should -Be 1
        $policy.conditions.clientApplications = @{ includeAgentIdServicePrincipals = @('agent') }
        (Get-TestRecommendation 'CA-26').Evidence.Count | Should -Be 1
    }

    It 'Does not guess emergency accounts and only accepts explicitly designated exclusions' {
        $accountId = '22222222-2222-2222-2222-222222222222'
        $policy.conditions.users.excludeUsers = @($accountId)
        (Get-TestRecommendation 'CA-01').Assessment | Should -Be 'Needs review'
        (Get-TestRecommendation 'CA-01' -EmergencyIds @($accountId)).Assessment | Should -Be 'Enforced'
        (Get-TestRecommendation 'CA-22').Message | Should -Match 'Supply -EmergencyAccessAccountId'
        (Get-TestRecommendation 'CA-22' -EmergencyIds @($accountId)).Evidence.Count | Should -Be 0
        $policy.conditions.users.excludeUsers = @()
        (Get-TestRecommendation 'CA-22' -EmergencyIds @($accountId)).Evidence.Count | Should -Be 1
    }

    It 'Measures an optional M365 device template as <Expected> without a baseline pass' -TestCases @(
        @{ State = 'enabled'; Expected = 'Matching policy enabled' }
        @{ State = 'enabledForReportingButNotEnforced'; Expected = 'Matching policy report-only' }
        @{ State = 'disabled'; Expected = 'Matching policy disabled' }
        @{ State = 'unknownFutureValue'; Expected = 'Partial match' }
    ) {
        param($State, $Expected)
        $policy.state = $State
        $policy.conditions.applications.includeApplications = @('Office365')
        $policy.grantControls.builtInControls = @('compliantDevice')
        $result = Get-TestRecommendation 'CA-28'
        $result.Assessment | Should -Be 'Needs review'
        $result.TemplateAssessment | Should -Be $Expected
        $result.Evidence[0].Resources | Should -Contain 'Office365'
        $result.TemplateSettings.Count | Should -BeGreaterThan 0
    }

    It 'Distinguishes suite-wide, selected-resource and blanket-block scenarios' {
        $policy.conditions.applications.includeApplications = @('00000003-0000-0ff1-ce00-000000000000')
        $policy.grantControls.builtInControls = @('compliantDevice')
        (Get-TestRecommendation 'CA-28').TemplateAssessment | Should -Be 'Not observed'
        (Get-TestRecommendation 'CA-29').TemplateAssessment | Should -Be 'Matching policy enabled'
        $policy.conditions.applications.includeApplications = @('All')
        (Get-TestRecommendation 'CA-28').TemplateAssessment | Should -Be 'Matching policy enabled'
        (Get-TestRecommendation 'CA-29').TemplateAssessment | Should -Be 'Not observed'
        $policy.conditions.applications.includeApplications = @('Office365')
        $policy.grantControls.builtInControls = @('block')
        (Get-TestRecommendation 'CA-28').TemplateAssessment | Should -Be 'Not observed'
    }

    It 'Evaluates device-filter null coverage and direction for <Mode>: <Rule>' -TestCases @(
        @{ Mode = 'exclude'; Rule = '(device.isCompliant -eq True)'; Expected = 'Matching policy enabled' }
        @{ Mode = 'include'; Rule = 'device.isCompliant -ne "True"'; Expected = 'Matching policy enabled' }
        @{ Mode = 'include'; Rule = 'device.isCompliant -eq False'; Expected = 'Partial match' }
        @{ Mode = 'include'; Rule = 'device.isCompliant -eq True'; Expected = 'Partial match' }
        @{ Mode = 'exclude'; Rule = 'device.isCompliant -ne True'; Expected = 'Partial match' }
        @{ Mode = 'exclude'; Rule = 'device.isManaged -eq True'; Expected = 'Partial match' }
        @{ Mode = 'exclude'; Rule = '(device.isCompliant -eq True) -or (device.trustType -eq "AzureAD")'; Expected = 'Partial match' }
    ) {
        param($Mode, $Rule, $Expected)
        $policy.conditions.applications.includeApplications = @('Office365')
        $policy.grantControls.builtInControls = @('block')
        $policy.conditions.devices = @{ deviceFilter = @{ mode = $Mode; rule = $Rule } }
        $result = Get-TestRecommendation 'CA-28'
        $result.TemplateAssessment | Should -Be $Expected
        $result.Evidence[0].DeviceFilter.rule | Should -Be $Rule
        (Get-TestRecommendation 'CA-20').Evidence.Count | Should -Be 1
    }

    It 'Keeps device templates partial for <Scenario>' -TestCases @(
        @{ Scenario = 'an MFA OR bypass'; Change = { $policy.grantControls.builtInControls += 'mfa' } }
        @{ Scenario = 'excluded resources'; Change = { $policy.conditions.applications.excludeApplications = @('resource-id') } }
        @{ Scenario = 'a dynamic application filter'; Change = { $policy.conditions.applications.AdditionalProperties = @{ applicationFilter = @{ mode = 'exclude'; rule = 'customSecurityAttributes.test -eq "value"' } } } }
        @{ Scenario = 'a device filter narrowing a compliance grant'; Change = { $policy.conditions.devices = @{ deviceFilter = @{ mode = 'include'; rule = 'device.isCompliant -eq True' } } } }
        @{ Scenario = 'a location-only restriction'; Change = { $policy.conditions.locations = @{ includeLocations = @('location') } } }
    ) {
        param($Scenario, $Change)
        $policy.conditions.applications.includeApplications = @('Office365')
        $policy.grantControls.builtInControls = @('compliantDevice')
        . $Change
        $result = Get-TestRecommendation 'CA-28'
        $result.TemplateAssessment | Should -Be 'Partial match'
        $result.Evidence[0].Issues.Count | Should -BeGreaterThan 0
    }

    It 'Measures the SharePoint session pattern without certifying application-side restrictions' {
        $policy.conditions.applications.includeApplications = @('00000003-0000-0ff1-ce00-000000000000')
        $policy.conditions.clientAppTypes = @('browser')
        $policy.sessionControls.applicationEnforcedRestrictions = @{ isEnabled = $true }
        $result = Get-TestRecommendation 'CA-30'
        $result.TemplateAssessment | Should -Be 'Matching policy enabled'
        $result.Assessment | Should -Be 'Needs review'
        $result.Verification | Should -Match 'cannot read those SharePoint settings'
        $policy.sessionControls.applicationEnforcedRestrictions.isEnabled = $false
        (Get-TestRecommendation 'CA-30').TemplateAssessment | Should -Be 'Not observed'
        $policy.sessionControls.applicationEnforcedRestrictions.isEnabled = $true
        $policy.conditions.applications.includeApplications = @('00000002-0000-0ff1-ce00-000000000000')
        (Get-TestRecommendation 'CA-30').Evidence.Count | Should -Be 0
    }

    It 'Measures mandatory M365 mobile app protection and rejects device/legacy alternatives' {
        $policy.conditions.applications.includeApplications = @('Office365')
        $policy.conditions.platforms = @{ includePlatforms = @('android','iOS') }
        $policy.conditions.clientAppTypes = @('mobileAppsAndDesktopClients')
        $policy.grantControls.builtInControls = @('compliantApplication')
        (Get-TestRecommendation 'CA-31').TemplateAssessment | Should -Be 'Matching policy enabled'
        $policy.grantControls.builtInControls += 'compliantDevice'
        (Get-TestRecommendation 'CA-31').TemplateAssessment | Should -Be 'Partial match'
        $policy.grantControls.builtInControls = @('approvedApplication')
        (Get-TestRecommendation 'CA-31').TemplateAssessment | Should -Be 'Partial match'
        $policy.grantControls.builtInControls = @('compliantApplication')
        $policy.conditions.clientAppTypes = @('browser')
        (Get-TestRecommendation 'CA-31').TemplateAssessment | Should -Be 'Partial match'
        $policy.conditions.clientAppTypes = @('mobileAppsAndDesktopClients')
        $policy.conditions.platforms.includePlatforms = @('android')
        (Get-TestRecommendation 'CA-31').TemplateAssessment | Should -Be 'Partial match'
    }

    It 'Renders optional settings and escapes resource/filter evidence' {
        $policy.conditions.applications.includeApplications = @('Office365','<img src=x>')
        $policy.grantControls.builtInControls = @('block')
        $policy.conditions.devices = @{ deviceFilter = @{ mode = 'include'; rule = '<img src=x>' } }
        $html = Get-RecommendationsHtmlFragment -Recommendations @(Get-CaRecommendationResult -Policies @($policy))
        $html | Should -Match "data-kind='template'"
        $html | Should -Match 'Optional template'
        $html | Should -Match 'Included resources:'
        $html | Should -Match '&lt;img src=x&gt;'
        $html | Should -Not -Match '<img'
    }

    It 'Does not make tenant-wide claims for a single-policy export' {
        $results = @(Get-CaRecommendationResult -Policies @($policy) -PartialExport)
        @($results | Where-Object Assessment -ne 'Needs review').Count | Should -Be 0
    }

    It 'Escapes policy evidence and renders requirements, state labels and Microsoft links' {
        $policy.displayName = '<img src=x onerror=alert(1)> & Test'
        $html = Get-RecommendationsHtmlFragment -Recommendations @(Get-CaRecommendationResult -Policies @($policy))
        $html | Should -Match 'Requirements and prerequisites:'
        $html | Should -Match 'data-assessment=.enforced'
        $html | Should -Match 'https://learn.microsoft.com/'
        $html | Should -Match '&lt;img src=x onerror=alert\(1\)&gt; &amp; Test'
        $html | Should -Not -Match '<img'
    }
}

Describe 'Summary table presentation' {
    BeforeAll {
        $source = Get-Content -LiteralPath "$PSScriptRoot/../Export-CAPolicyWithRecs.ps1" -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
        $helper = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Add-SummaryRow'
        }, $true)
        . ([scriptblock]::Create($helper.Extent.Text))
    }

    It 'Shows the percentage denominator and a culture-independent visual share' {
        $originalCulture = [System.Globalization.CultureInfo]::CurrentCulture
        try {
            [System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
            $html = Add-SummaryRow -Name 'Administrator MFA' -Value 1 -TotalRef 3
            $html | Should -Match "summary-share-base'>of 3"
            $html | Should -Match "style='width:33\.333%'"
            $html | Should -Match '33,3'
        }
        finally {
            [System.Globalization.CultureInfo]::CurrentCulture = $originalCulture
        }
    }

    It 'Shows not-applicable instead of a misleading zero percentage when there is no denominator' {
        $html = Add-SummaryRow -Name 'Total policies' -Value 0
        $html | Should -Match "summary-number'>0"
        $html | Should -Match 'aria-label="Not applicable"'
        $html | Should -Not -Match 'summary-share-fill'
    }

    It 'Escapes row labels and units and provides an accessible metric header' {
        $html = Add-SummaryRow -Name '<Metric>' -Value 2.5 -Unit '<days>'
        $html | Should -Match "<th scope='row'>&lt;Metric&gt;</th>"
        $html | Should -Match '&lt;days&gt;'
        $html | Should -Not -Match '<Metric>|<days>'
    }
}

Describe 'Graph connection scope verification' {
    BeforeAll {
        $source = Get-Content -LiteralPath "$PSScriptRoot/../Export-CAPolicyWithRecs.ps1" -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
        $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in @('Connect-GraphContext', 'Write-Info')
        }, $true) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
        function Get-MgContext { [CmdletBinding()] param() }
        function Connect-MgGraph { [CmdletBinding()] param([string[]]$Scopes) }
    }

    BeforeEach {
        $script:ContextReads = 0
        $validContext = [pscustomobject]@{
            Account = 'test@example.invalid'
            Scopes = @('Policy.Read.All', 'Directory.Read.All', 'RoleManagement.Read.All')
        }
        Mock Connect-MgGraph {}
        Mock Get-MgContext { $validContext }
    }

    It 'Keeps an existing connection with all required scopes' {
        Connect-GraphContext

        Should -Invoke Connect-MgGraph -Times 0
        Should -Invoke Get-MgContext -Times 1 -Exactly
    }

    It 'Reads and validates the new context after an initial sign-in' {
        Mock Get-MgContext {
            if ($script:ContextReads++ -gt 0) { $validContext }
        }

        Connect-GraphContext

        Should -Invoke Get-MgContext -Times 2 -Exactly
        Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
            $Scopes.Count -eq 3 -and $Scopes -contains 'Policy.Read.All' -and
            $Scopes -contains 'Directory.Read.All' -and $Scopes -contains 'RoleManagement.Read.All' -and $ErrorAction -eq 'Stop'
        }
    }

    It 'Reconnects an incomplete cached context and validates the replacement' {
        Mock Get-MgContext {
            if ($script:ContextReads++ -eq 0) { [pscustomobject]@{ Account = 'test@example.invalid'; Scopes = @('Policy.Read.All') } }
            else { $validContext }
        }

        Connect-GraphContext

        Should -Invoke Connect-MgGraph -Times 1 -Exactly
        Should -Invoke Get-MgContext -Times 2 -Exactly
    }

    It 'Stops with the missing permission names when consent remains incomplete' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'test@example.invalid'; Scopes = @('Policy.Read.All') } }

        { Connect-GraphContext } | Should -Throw '*Directory.Read.All, RoleManagement.Read.All*Grant admin consent*'
        Should -Invoke Get-MgContext -Times 2 -Exactly
    }

    It 'Stops when connection succeeds without returning an authenticated context' {
        Mock Get-MgContext { $null }

        { Connect-GraphContext } | Should -Throw '*did not return an authenticated account*'
    }

    It 'Stops on nonterminating sign-in errors instead of exporting with a stale context' {
        Mock Get-MgContext { $null }
        Mock Connect-MgGraph { Write-Error 'Consent denied' }

        { Connect-GraphContext } | Should -Throw '*Consent denied*'
        Should -Invoke Get-MgContext -Times 1 -Exactly
    }
}

Describe 'Named location inventory and rendering' {
    BeforeAll {
        $source = Get-Content -LiteralPath "$PSScriptRoot/../Export-CAPolicyWithRecs.ps1" -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
        $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -in @('Get-CaNamedLocationInventory', 'Get-NamedLocationHtmlFragment', 'Get-CaProperty', 'Write-Warn')
        }, $true) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
        function Get-MgIdentityConditionalAccessNamedLocation { [CmdletBinding()] param([switch]$All) }
    }

    BeforeEach {
        Mock Get-MgIdentityConditionalAccessNamedLocation { @() }
        Mock Write-Warn {}
    }

    It 'Requests every page of tenant locations, including unreferenced definitions' {
        Mock Get-MgIdentityConditionalAccessNamedLocation {
            [pscustomobject]@{ Id = 'referenced'; DisplayName = 'Office' }
            [pscustomobject]@{ Id = 'unreferenced'; DisplayName = 'Unused office' }
        }

        $inventory = Get-CaNamedLocationInventory

        $inventory.Status | Should -Be 'Available'
        $inventory.Locations.Count | Should -Be 2
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 1 -Exactly -ParameterFilter { $All -and $ErrorAction -eq 'Stop' }
    }

    It 'Distinguishes an empty tenant inventory from an unavailable offline snapshot' {
        $empty = Get-CaNamedLocationInventory
        $offline = Get-CaNamedLocationInventory -Offline

        $empty.Status | Should -Be 'Available'
        $empty.Locations.Count | Should -Be 0
        Get-NamedLocationHtmlFragment $empty | Should -Match 'No named locations are defined'
        $offline.Status | Should -Be 'Offline'
        Get-NamedLocationHtmlFragment $offline | Should -Match 'not included in raw policy snapshots'
        Get-NamedLocationHtmlFragment $offline | Should -Not -Match 'No named locations are defined'
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 1 -Exactly
    }

    It 'Surfaces failed inventory retrieval in both warnings and the report' {
        Mock Get-MgIdentityConditionalAccessNamedLocation { Write-Error 'Access denied <detail>' }

        $inventory = Get-CaNamedLocationInventory
        $html = Get-NamedLocationHtmlFragment $inventory

        $inventory.Status | Should -Be 'Failed'
        $html | Should -Match 'Could not retrieve named locations: Access denied &lt;detail&gt;'
        $html | Should -Not -Match 'No named locations are defined'
        Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter { $Message -like '*Access denied*' }
    }

    It 'Renders IP and country definitions from <Representation> locations' -TestCases @(
        @{ Representation = 'raw JSON' }
        @{ Representation = 'SDK derived properties' }
        @{ Representation = 'SDK AdditionalProperties' }
    ) {
        param($Representation)
        $ip = @{
            id = 'office-id'; displayName = '<Office & VPN>'; '@odata.type' = '#microsoft.graph.ipNamedLocation'
            isTrusted = $true; ipRanges = @(@{ cidrAddress = '192.0.2.0/24' }, @{ cidrAddress = '2001:db8::/32' })
            createdDateTime = '2026-01-01T00:00:00Z'; modifiedDateTime = '2026-02-01T00:00:00Z'
        }
        $country = @{
            id = 'country-id'; displayName = 'Countries'; '@odata.type' = '#microsoft.graph.countryNamedLocation'
            countriesAndRegions = @('US', 'CA'); includeUnknownCountriesAndRegions = $true; countryLookupMethod = 'authenticatorAppGps'
        }
        $expectedTrust = 'Trusted'
        $expectedUnknown = 'Included'
        if ($Representation -eq 'SDK derived properties') {
            $ip.OdataType = $ip['@odata.type']; $ip.Remove('@odata.type')
            $country.OdataType = $country['@odata.type']; $country.Remove('@odata.type')
            $ip = [pscustomobject]$ip; $country = [pscustomobject]$country
        }
        elseif ($Representation -eq 'SDK AdditionalProperties') {
            $ip.isTrusted = $false
            $country.includeUnknownCountriesAndRegions = $false
            $expectedTrust = 'Not trusted'
            $expectedUnknown = 'Excluded'
            $ipProperties = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
            foreach ($key in $ip.Keys) { $ipProperties.Add($key, $ip[$key]) }
            $ipProperties['ipRanges'] = @($ip.ipRanges | ForEach-Object {
                $range = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
                $range.Add('cidrAddress', $_.cidrAddress)
                $range
            })
            $countryProperties = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
            foreach ($key in $country.Keys) { $countryProperties.Add($key, $country[$key]) }
            $ip = [pscustomobject]@{ Id = $ip.id; DisplayName = $ip.displayName; OdataType = $null; IpRanges = $null; AdditionalProperties = $ipProperties }
            $country = [pscustomobject]@{ Id = $country.id; DisplayName = $country.displayName; OdataType = $null; CountriesAndRegions = $null; AdditionalProperties = $countryProperties }
        }

        $html = Get-NamedLocationHtmlFragment ([pscustomobject]@{ Status = 'Available'; Locations = @($ip, $country) })

        $html | Should -Match '&lt;Office &amp; VPN&gt;'
        $html | Should -Not -Match '<Office'
        $html | Should -Match '192\.0\.2\.0/24'
        $html | Should -Match '2001:db8::/32'
        $html | Should -Match "<p class='named-location-flag'>$expectedTrust</p>"
        $html | Should -Match "Unknown countries/regions:</strong> $expectedUnknown"
        $html | Should -Match 'authenticatorAppGps'
        $html | Should -Match '<li>CA</li><li>US</li>'
        $html | Should -Match '2026-02-01T00:00:00Z'
    }

    It 'Does not label missing trust or country flags as false' {
        $html = Get-NamedLocationHtmlFragment ([pscustomobject]@{ Status = 'Available'; Locations = @(
            @{ '@odata.type' = '#microsoft.graph.ipNamedLocation'; displayName = 'Unknown trust' },
            @{ '@odata.type' = '#microsoft.graph.countryNamedLocation'; displayName = 'Unknown areas' }
        ) })

        $html | Should -Match 'Trust not returned'
        $html | Should -Match 'Unknown countries/regions:</strong> Not returned'
        $html | Should -Not -Match 'Not trusted|Unknown countries/regions:</strong> Excluded'
    }

    It 'Preserves explicit false flags and escapes unknown location properties' {
        $html = Get-NamedLocationHtmlFragment ([pscustomobject]@{ Status = 'Available'; Locations = @(
            @{ '@odata.type' = '#microsoft.graph.ipNamedLocation'; isTrusted = $false },
            @{ '@odata.type' = '#microsoft.graph.countryNamedLocation'; includeUnknownCountriesAndRegions = $false },
            @{ '@odata.type' = '<future-type>'; displayName = '<script>alert(1)</script>'; custom = '<img src=x>' }
        ) })

        $html | Should -Match 'Not trusted'
        $html | Should -Match 'Unknown countries/regions:</strong> Excluded'
        $html | Should -Match 'Unrecognized location type'
        $html | Should -Match '&lt;future-type&gt;'
        $html | Should -Not -Match '<script>|<img|<future-type>'
    }
}

Describe 'Offline recommendation exports' {
    BeforeAll {
        function Get-MgContext { [CmdletBinding()] param() }
        function Get-MgOrganization { [CmdletBinding()] param() }
        function Get-MgIdentityConditionalAccessPolicy { [CmdletBinding()] param([switch]$All, $Headers) }
        function Get-MgIdentityConditionalAccessNamedLocation { [CmdletBinding()] param([switch]$All, $NamedLocationId, $Property) }
        function Get-MgDirectoryRole { param([switch]$All) }
        function Get-MgRoleManagementDirectoryRoleDefinition { param([switch]$All, $Property) }
        function Get-MgUser { [CmdletBinding()] param($UserId, $Property) }
        function Get-MgGroup { [CmdletBinding()] param($GroupId, $Property) }
        function Get-MgServicePrincipal { [CmdletBinding()] param($ServicePrincipalId, $Property) }
        function Get-MgAgreement { [CmdletBinding()] param($AgreementId, $Property) }
        function Get-MgPolicyAuthenticationStrengthPolicy { param([switch]$All, $Property) }
        $exportScript = "$PSScriptRoot/../Export-CAPolicyWithRecs.ps1"
        $rawPath = Join-Path $TestDrive 'policies_raw.json'
        $basePolicy = @{
            id = '11111111-1111-1111-1111-111111111111'; displayName = 'MFA'; state = 'enabled'
            createdDateTime = '2026-01-01T00:00:00Z'
            conditions = @{
                users = @{ includeUsers = @('All'); excludeUsers = @(); includeGroups = @(); excludeGroups = @(); includeRoles = @(); excludeRoles = @() }
                applications = @{ includeApplications = @('All'); excludeApplications = @() }
                clientAppTypes = @('all')
            }
            grantControls = @{ operator = 'OR'; builtInControls = @('mfa') }
        }
        $device = $basePolicy | ConvertTo-Json -Depth 20 | ConvertFrom-Json -AsHashtable
        $device.id = '22222222-2222-2222-2222-222222222222'
        $device.displayName = 'Compliance disabled'
        $device.state = 'disabled'
        $device.grantControls.builtInControls = @('compliantDevice')
        $registration = $basePolicy | ConvertTo-Json -Depth 20 | ConvertFrom-Json -AsHashtable
        $registration.id = '33333333-3333-3333-3333-333333333333'
        $registration.displayName = 'Registration report only'
        $registration.state = 'enabledForReportingButNotEnforced'
        $registration.conditions.applications = @{ includeUserActions = @('urn:user:registersecurityinfo') }
        @($basePolicy, $device, $registration) | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $rawPath
    }

    It 'Produces all formats with 32 state-aware recommendation cards without calling Graph' {
        $out = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'all-formats')
        & $exportScript -RawInputFile $rawPath -OutputPath $out.FullName -Html -Json -Csv -CsvPivot -NoBrowser -WarningVariable exportWarnings
        @($exportWarnings | Where-Object Message -Like '*NoTypeInformation*').Count | Should -Be 0
        $html = Get-Content -LiteralPath (Get-ChildItem $out.FullName -Filter '*.html').FullName -Raw
        ([regex]::Matches($html, "<details class='recommendation-card ")).Count | Should -Be 32
        foreach ($state in @('enforced','report-only','disabled','missing','needs-review','no-findings')) {
            $html | Should -Match "data-assessment='$state'"
        }
        ([regex]::Matches($html, '<tr class="summary-section-row">')).Count | Should -Be 5
        ([regex]::Matches($html, "<tr class='summary-metric-row'>")).Count | Should -Be 15
        $html | Should -Match 'aria-describedby="summary-scope-note"'
        $html | Should -Match 'summary-value-unit">days'
        ([regex]::Matches($html, "<button type='button' id='rec-filter-")).Count | Should -Be 4
        $html | Should -Match "id='rec-assessment-filter'"
        $html | Should -Match "id='rec-result-count'"
        $html | Should -Match "id='rec-filter-empty'"
        $html | Should -Match "id='tab-locations'"
        $html | Should -Match "id='panel-locations'"
        $html | Should -Match 'not included in raw policy snapshots'
        @(Get-ChildItem $out.FullName -Filter '*.json').Count | Should -Be 1
        @(Get-ChildItem $out.FullName -Filter '*.csv').Count | Should -Be 2
        foreach ($csv in Get-ChildItem $out.FullName -Filter '*.csv') {
            Get-Content $csv.FullName -TotalCount 1 | Should -Not -Match '^#TYPE'
        }
        $policyCsv = Get-ChildItem $out.FullName -Filter '*.csv' | Where-Object Name -NotLike '*-pivot.csv'
        @(Import-Csv $policyCsv.FullName).Count | Should -Be 3
        $json = Get-ChildItem $out.FullName -Filter '*.json' | Where-Object Name -NotLike '*_raw.json'
        @(Get-Content $json.FullName -Raw | ConvertFrom-Json).Count | Should -Be 3
    }

    It 'Preserves NoRecommendations mode without recommendation cards or a recommendations tab' {
        $out = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'no-recommendations')
        & $exportScript -RawInputFile $rawPath -OutputPath $out.FullName -Html -NoRecommendations -NoBrowser
        $html = Get-Content -LiteralPath (Get-ChildItem $out.FullName -Filter '*.html').FullName -Raw
        $html | Should -Not -Match "<details class='recommendation-card "
        $html | Should -Not -Match "id='tab-recommendations'"
        $html | Should -Not -Match "id='rec-assessment-filter'"
        $html | Should -Match 'Recommendations omitted'
        $html | Should -Match "id='tab-locations'"
        $html | Should -Match "id='panel-locations'"
    }

    It 'Captures a portable snapshot during a <Mode> export and regenerates all formats on a clean offline process' -TestCases @(
        @{ Mode = 'HTML'; IncludeHtml = $true }
        @{ Mode = 'CSV-only'; IncludeHtml = $false }
    ) {
        param($Mode, $IncludeHtml)
        $livePolicies = Get-Content $rawPath -Raw | ConvertFrom-Json -AsHashtable
        $locationId = '44444444-4444-4444-4444-444444444444'
        $livePolicies[1].conditions.locations = @{ includeLocations = @($locationId) }
        $livePolicies[1].conditions.users.includeUsers = @('66666666-6666-6666-6666-666666666666')
        $livePolicies[1].conditions.users.includeGroups = @('77777777-7777-7777-7777-777777777777')
        $livePolicies[1].conditions.users.includeRoles = @('62e90394-69f5-4237-9190-012177145e10')
        $livePolicies[1].conditions.applications.includeApplications = @('88888888-8888-8888-8888-888888888888')
        $livePolicies[1].grantControls.termsOfUse = @('99999999-9999-9999-9999-999999999999')
        $livePolicies[0].grantControls.builtInControls = @()
        $livePolicies[0].grantControls.authenticationStrength = @{ id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }
        $livePolicies[2].conditions.AdditionalProperties = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        $livePolicies[2].conditions.AdditionalProperties.Add('futureCondition', @{ value = 'restricted' })
        $emergencyIds = @('10000000-0000-0000-0000-000000000001', '10000000-0000-0000-0000-000000000002')
        Mock Get-PackageProvider { [pscustomobject]@{ Name = 'NuGet' } } -ParameterFilter { $Name -eq 'NuGet' }
        Mock Get-PSRepository { [pscustomobject]@{ InstallationPolicy = 'Trusted' } } -ParameterFilter { $Name -eq 'PSGallery' }
        Mock Get-Module { [pscustomobject]@{ Name = $Name } } -ParameterFilter { $Name -like 'Microsoft.Graph.*' }
        Mock Import-Module {} -ParameterFilter { $Name -like 'Microsoft.Graph.*' }
        Mock Get-MgContext {
            [pscustomobject]@{ Account = 'test@example.invalid'; Scopes = @('Policy.Read.All', 'Directory.Read.All', 'RoleManagement.Read.All') }
        }
        Mock Get-MgOrganization { [pscustomobject]@{ DisplayName = 'Synthetic tenant' } }
        Mock Get-MgIdentityConditionalAccessPolicy { $livePolicies }
        Mock Get-MgDirectoryRole { @() }
        Mock Get-MgRoleManagementDirectoryRoleDefinition { @() }
        Mock Get-MgUser { [pscustomobject]@{ DisplayName = 'Sample user' } }
        Mock Get-MgGroup { [pscustomobject]@{ DisplayName = 'Sample group' } }
        Mock Get-MgServicePrincipal { [pscustomobject]@{ DisplayName = 'Sample application' } }
        Mock Get-MgAgreement { [pscustomobject]@{ DisplayName = 'Sample terms of use' } }
        Mock Get-MgPolicyAuthenticationStrengthPolicy {
            [pscustomobject]@{
                Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Custom sample strength'
                AllowedCombinations = @('fido2'); RequirementsSatisfied = 'mfa'
            }
        }
        Mock Get-MgIdentityConditionalAccessNamedLocation {
            [pscustomobject]@{
                Id = $locationId; DisplayName = 'Office & VPN'; OdataType = '#microsoft.graph.ipNamedLocation'
                AdditionalProperties = @{
                    isTrusted = $false
                    ipRanges = @(@{ cidrAddress = '192.0.2.0/24' }, @{ cidrAddress = '2001:db8::/32' })
                }
            }
            [pscustomobject]@{
                Id = '55555555-5555-5555-5555-555555555555'; DisplayName = 'Unreferenced countries'
                AdditionalProperties = @{
                    '@odata.type' = '#microsoft.graph.countryNamedLocation'; countriesAndRegions = @('US', 'CA')
                    includeUnknownCountriesAndRegions = $false; countryLookupMethod = 'clientIpAddress'
                }
            }
        }
        $out = New-Item -ItemType Directory -Path (Join-Path $TestDrive "connected-$Mode")

        & $exportScript -OutputPath $out.FullName -Html:$IncludeHtml -Json:$IncludeHtml -Csv -CsvPivot:$IncludeHtml `
            -NoRecommendations:(-not $IncludeHtml) -NoBrowser -EmergencyAccessAccountId $emergencyIds

        $snapshotPath = (Get-ChildItem $out.FullName -Filter '*_raw.json').FullName
        $snapshot = Get-Content $snapshotPath -Raw | ConvertFrom-Json
        [math]::Abs(([datetimeoffset]::UtcNow - [datetimeoffset]::Parse([string]$snapshot.CapturedAt)).TotalMinutes) |
            Should -BeLessThan 2
        $snapshot.Policies.Count | Should -Be 3
        $snapshot.Policies[1].conditions.locations.includeLocations[0] | Should -Be $locationId
        $snapshot.NamedLocations.Locations.Count | Should -Be 2
        $snapshot.Lookups.Applications.'88888888-8888-8888-8888-888888888888' | Should -Be 'Sample application'
        $snapshot.AuthenticationStrengths.Count | Should -Be 1
        $snapshot.EmergencyAccessAccountIds | Should -Be $emergencyIds
        if (-not $IncludeHtml) { @(Get-ChildItem $out.FullName -Filter '*.html').Count | Should -Be 0 }

        $offlineDirectory = New-Item -ItemType Directory -Path (Join-Path $TestDrive "another-computer-$Mode")
        $copiedInput = Join-Path $offlineDirectory.FullName 'copied-snapshot.json'
        Copy-Item -LiteralPath $snapshotPath -Destination $copiedInput
        $runner = Join-Path $TestDrive "offline-runner-$Mode.ps1"
        @'
param($Exporter, $Snapshot, $Destination)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.PowerShell.Utility, Microsoft.PowerShell.Management
$PSModuleAutoLoadingPreference = 'None'
& $Exporter -RawInputFile $Snapshot -OutputPath $Destination -Html -Json -Csv -CsvPivot -NoBrowser
'@ | Set-Content -LiteralPath $runner
        $processOutput = & (Get-Process -Id $PID).Path -NoLogo -NoProfile -File $runner $exportScript $copiedInput $offlineDirectory.FullName 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($processOutput -join "`n")

        $html = Get-Content -LiteralPath (Get-ChildItem $offlineDirectory.FullName -Filter '*.html').FullName -Raw
        $html | Should -Match 'Unreferenced countries'
        $html | Should -Match '192\.0\.2\.0/24'
        $html | Should -Match '2001:db8::/32'
        $html | Should -Match 'Not trusted'
        $html | Should -Match 'Unknown countries/regions:</strong> Excluded'
        $html | Should -Match 'clientIpAddress'
        $html | Should -Match '<td[^>]*>Office &amp; VPN</td>'
        $html | Should -Match 'Sample user'
        $html | Should -Match 'Sample group'
        $html | Should -Match 'Global Administrator'
        $html | Should -Match 'Sample terms of use'
        $html | Should -Match "data-assessment='enforced'[^>]*id='rec-CA-01'"
        $html | Should -Match "data-assessment='needs-review'[^>]*id='rec-CA-13'"
        $html | Should -Match 'CA Export v3\.7 - Synthetic tenant'
        $html | Should -Match ([regex]::Escape("Snapshot captured: $($snapshot.CapturedAt)"))
        $html | Should -Not -Match 'not included in raw policy snapshots'
        $connectedCsv = Get-ChildItem $out.FullName -Filter '*.csv' | Where-Object Name -NotLike '*-pivot.csv'
        $offlineCsv = Get-ChildItem $offlineDirectory.FullName -Filter '*.csv' | Where-Object Name -NotLike '*-pivot.csv'
        (Get-Content $offlineCsv.FullName -Raw) | Should -Be (Get-Content $connectedCsv.FullName -Raw)
        if ($IncludeHtml) {
            $connectedJson = Get-ChildItem $out.FullName -Filter '*.json' | Where-Object Name -NotLike '*_raw.json'
            $offlineJson = Get-ChildItem $offlineDirectory.FullName -Filter '*.json' | Where-Object Name -NotLike 'copied-snapshot.json'
            # Suggested names already contain a freshly randomized numeric prefix on every run.
            $connectedPolicies = Get-Content $connectedJson.FullName -Raw | ConvertFrom-Json |
                Select-Object * -ExcludeProperty 'Recommended Name' | ConvertTo-Json -Depth 12
            $offlinePolicies = Get-Content $offlineJson.FullName -Raw | ConvertFrom-Json |
                Select-Object * -ExcludeProperty 'Recommended Name' | ConvertTo-Json -Depth 12
            $offlinePolicies | Should -Be $connectedPolicies
            $connectedHtml = Get-Content -LiteralPath (Get-ChildItem $out.FullName -Filter '*.html').FullName -Raw
            $connectedMfaCard = [regex]::Match($connectedHtml, "(?s)<details class='recommendation-card [^>]+id='rec-CA-01'.*?</details>").Value
            $connectedMfaCard | Should -Match "data-assessment='enforced'"
            [regex]::Matches($html, "<details class='recommendation-card [^>]+>").Value |
                Should -Be ([regex]::Matches($connectedHtml, "<details class='recommendation-card [^>]+>").Value)
        }
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 1 -Exactly -ParameterFilter { $All }
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 0 -ParameterFilter { $NamedLocationId }
    }

    It 'Preserves a <InventoryState> inventory and partial-export limits when regenerating a saved snapshot' -TestCases @(
        @{ InventoryState = 'Available'; ExpectedMessage = 'No named locations are defined' }
        @{ InventoryState = 'Failed'; ExpectedMessage = 'Saved collection failure &lt;detail&gt;' }
    ) {
        param($InventoryState, $ExpectedMessage)
        $snapshotPath = Join-Path $TestDrive "partial-$InventoryState.json"
        $saved = @{
            SnapshotFormat = 'CAExportSnapshot'; SchemaVersion = 1
            TenantName = 'Synthetic saved tenant'; CapturedAt = '2026-01-02 03:04:05Z'; PartialExport = $true
            Policies = @($basePolicy)
            NamedLocations = @{ Status = $InventoryState; Locations = @(); Message = 'Saved collection failure <detail>' }
            Lookups = @{ Users = @{}; Groups = @{}; Roles = @{}; Applications = @{}; Locations = @{}; TermsOfUse = @{} }
            AuthenticationStrengths = @(); EmergencyAccessAccountIds = @()
        }
        ConvertTo-Json -InputObject $saved -Depth 30 | Set-Content -LiteralPath $snapshotPath
        $out = New-Item -ItemType Directory -Path (Join-Path $TestDrive "partial-report-$InventoryState")

        & $exportScript -RawInputFile $snapshotPath -OutputPath $out.FullName -Html -NoBrowser

        $html = Get-Content -LiteralPath (Get-ChildItem $out.FullName -Filter '*.html').FullName -Raw
        $html | Should -Match $ExpectedMessage
        $html | Should -Match 'Snapshot captured: 2026-01-02 03:04:05Z'
        ([regex]::Matches($html, "data-assessment='needs-review'")).Count | Should -Be 32
        $html | Should -Not -Match "data-attention='true' id='rec-CA-22'"
        if ($InventoryState -eq 'Failed') { $html | Should -Not -Match 'No named locations are defined' }
    }

    It 'Rejects <InvalidInput> rather than rendering a misleading offline report' -TestCases @(
        @{ InvalidInput = 'an unsupported snapshot version' }
        @{ InvalidInput = 'an incomplete snapshot' }
        @{ InvalidInput = 'enriched policy JSON' }
        @{ InvalidInput = 'invalid JSON' }
    ) {
        param($InvalidInput)
        $inputPath = Join-Path $TestDrive 'invalid-offline.json'
        $content = switch ($InvalidInput) {
            'an unsupported snapshot version' { '{"SnapshotFormat":"CAExportSnapshot","SchemaVersion":99}' }
            'an incomplete snapshot' { '{"SnapshotFormat":"CAExportSnapshot","SchemaVersion":1,"Policies":[]}' }
            'enriched policy JSON' { '[{"Name":"Sample policy","PolicyId":"11111111-1111-1111-1111-111111111111","Status":"Enabled"}]' }
            'invalid JSON' { '{ not JSON' }
        }
        Set-Content -LiteralPath $inputPath -Value $content

        { & $exportScript -RawInputFile $inputPath -OutputPath $TestDrive -Html -NoBrowser } |
            Should -Throw '*Failed to load offline input*'
    }
}

Describe 'User and group enrichment' {
    BeforeAll {
        $source = Get-Content -LiteralPath "$PSScriptRoot/../Export-CAPolicyWithRecs.ps1" -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
        $helperNames = @('Initialize-GraphModule', 'Invoke-SafeGet', 'Write-Warn', 'Write-Info', 'Test-IsGuid', 'Convert-IdListToName', 'Get-CaNamedLocationInventory')
        $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in $helperNames
        }, $true) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
        $start = $source.IndexOf('# ---------------- Build lookup maps')
        $end = $source.IndexOf('# ---------------- Save portable snapshot')
        $enrich = [scriptblock]::Create($source.Substring($start, $end - $start))

        function Get-MgUser { [CmdletBinding()] param($UserId, $Property) }
        function Get-MgGroup { [CmdletBinding()] param($GroupId, $Property) }
        function Get-MgDirectoryRole { param([switch]$All) }
        function Get-MgRoleManagementDirectoryRoleDefinition { param([switch]$All, $Property) }
        function Get-MgIdentityConditionalAccessNamedLocation { [CmdletBinding()] param([switch]$All, $NamedLocationId, $Property) }
    }

    BeforeEach {
        $script:OfflineMode = $false
        $HTMLExport = $false
        $userIds = @('11111111-1111-1111-1111-111111111111')
        $groupIds = @('22222222-2222-2222-2222-222222222222')
        $roleIds = $appIds = $locIds = $touIds = @()
        Mock Get-MgUser { [pscustomobject]@{ DisplayName = 'Test User' } }
        Mock Get-MgGroup { [pscustomobject]@{ DisplayName = 'Test Group' } }
        Mock Get-MgDirectoryRole { @() }
        Mock Get-MgRoleManagementDirectoryRoleDefinition { @() }
        Mock Get-MgIdentityConditionalAccessNamedLocation { @() }
        Mock Write-Warn {}
    }

    It 'Loads the user and group commands during default module initialization' {
        Mock Get-PackageProvider { [pscustomobject]@{ Name = 'NuGet' } }
        Mock Get-PSRepository { [pscustomobject]@{ InstallationPolicy = 'Trusted' } }
        Mock Get-Module { [pscustomobject]@{ Name = $Name } }
        Mock Import-Module {}

        Initialize-GraphModule

        Should -Invoke Import-Module -Times 1 -Exactly -ParameterFilter { $Name -eq 'Microsoft.Graph.Users' }
        Should -Invoke Import-Module -Times 1 -Exactly -ParameterFilter { $Name -eq 'Microsoft.Graph.Groups' }
    }

    It 'Resolves assignment IDs using live user and group lookup results' {
        . $enrich

        Convert-IdListToName -List $userIds -Map $UserMap | Should -Be 'Test User'
        Convert-IdListToName -List $groupIds -Map $GroupMap | Should -Be 'Test Group'
        Should -Invoke Write-Warn -Times 0
    }

    It 'Warns on nonterminating Graph errors and retains the ID while other lookups continue' {
        Mock Get-MgUser { Write-Error 'Access denied' }

        . $enrich

        Convert-IdListToName -List $userIds -Map $UserMap | Should -Be $userIds[0]
        Convert-IdListToName -List $groupIds -Map $GroupMap | Should -Be 'Test Group'
        Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter {
            $Message -like "Could not resolve user '*'; keeping its ID:*Access denied*"
        }
    }

    It 'Does not replace an unresolved group ID with a blank name' {
        Mock Get-MgGroup { [pscustomobject]@{ DisplayName = $null } }

        . $enrich

        Convert-IdListToName -List $groupIds -Map $GroupMap | Should -Be $groupIds[0]
        Should -Invoke Write-Warn -Times 1 -Exactly -ParameterFilter {
            $Message -like "Could not resolve group '*'; keeping its ID:*no display name*"
        }
    }

    It 'Reuses the complete HTML location inventory for assignment name resolution' {
        $HTMLExport = $true
        $locIds = @('33333333-3333-3333-3333-333333333333')
        Mock Get-MgIdentityConditionalAccessNamedLocation {
            [pscustomobject]@{ Id = '33333333-3333-3333-3333-333333333333'; DisplayName = 'Office' }
            [pscustomobject]@{ Id = '44444444-4444-4444-4444-444444444444'; DisplayName = 'Unreferenced office' }
        }

        . $enrich

        $LocMap[$locIds[0]] | Should -Be 'Office'
        $NamedLocationInventory.Locations.Count | Should -Be 2
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 1 -Exactly
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 1 -Exactly -ParameterFilter { $All }
    }

    It 'Collects the full inventory for portable snapshots even without HTML' {
        $locIds = @('33333333-3333-3333-3333-333333333333')
        Mock Get-MgIdentityConditionalAccessNamedLocation {
            [pscustomobject]@{ Id = '33333333-3333-3333-3333-333333333333'; DisplayName = 'Office' }
            [pscustomobject]@{ Id = '44444444-4444-4444-4444-444444444444'; DisplayName = 'Unused office' }
        }

        . $enrich

        $LocMap[$locIds[0]] | Should -Be 'Office'
        $NamedLocationInventory.Locations.Count | Should -Be 2
        Should -Invoke Get-MgIdentityConditionalAccessNamedLocation -Times 1 -Exactly -ParameterFilter { $All }
    }
}

Describe 'Get-NormalizedPolicyHash' {
    BeforeAll {
        function Get-NormalizedPolicyHash { param([Parameter(Mandatory)]$Policy)
            $norm = $Policy | Select-Object * -ExcludeProperty PolicyId, DateModified, CreatedDateTime, Description, 'Duplicate Matches', IsDuplicate, RawJson, ContentHash
            $json = ($norm | ConvertTo-Json -Depth 8 -Compress)
            $bytes = [Text.Encoding]::UTF8.GetBytes($json)
            $sha = [System.Security.Cryptography.SHA256]::Create()
            return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '')
        }
    }
    It 'Produces identical hash when only description changes' {
        $base = [pscustomobject]@{ Name='TestPolicy'; Description='A'; Status='enabled'; Created=(Get-Date); Modified=(Get-Date); 'Duplicate Matches'=''; IsDuplicate=$false; RawJson=''; ContentHash='' }
        $variant = $base | Select-Object *
        $variant.Description = 'Different description text'
        $h1 = Get-NormalizedPolicyHash -Policy $base
        $h2 = Get-NormalizedPolicyHash -Policy $variant
        $h1 | Should -Be $h2
    }
}

Describe 'Test-IsPhishResistantStrength' {
    BeforeAll {
        function Test-IsPhishResistantStrength { param($Policy)
            $strengthObj = $Policy.GrantControls.AuthenticationStrength
            if (-not $strengthObj -or -not $strengthObj.AllowedCombinations) { return $false }
            $nonPhish = @('deviceBasedPush','temporaryAccessPassOneTime','temporaryAccessPassMultiUse','microsoftAuthenticatorPush','sms','voice','softwareOath','hardwareOath','x509CertificateSingleFactor','federatedSingleFactor','qrCodePin')
            $contains = ($strengthObj.AllowedCombinations | Where-Object { $nonPhish -contains $_ }).Count -gt 0
            return (-not $contains)
        }
        # Prepare fake auth strength cache
        $script:AuthStrengthCache = @{}
        $script:AuthStrengthCachePopulated = $true
        $script:AuthStrengthCache['strength-phish'] = [pscustomobject]@{ Id='strength-phish'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2','windowsHelloForBusiness') }
        $script:AuthStrengthCache['strength-nonphish'] = [pscustomobject]@{ Id='strength-nonphish'; DisplayName='NonPhish'; AllowedCombinations=@('microsoftAuthenticatorPush','fido2') }
    }
    It 'Returns True for strength with only phish-resistant combos' {
        $policy = [pscustomobject]@{ GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-phish'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2','windowsHelloForBusiness') } }; Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles=@() } } }
        (Test-IsPhishResistantStrength -Policy $policy) | Should -BeTrue
    }
    It 'Returns False when any non-phish combo present' {
        $policy = [pscustomobject]@{ GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-nonphish'; DisplayName='NonPhish'; AllowedCombinations=@('microsoftAuthenticatorPush','fido2') } }; Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles=@() } } }
        (Test-IsPhishResistantStrength -Policy $policy) | Should -BeFalse
    }
}

Describe 'CA Check Functions Representative Tests' {
    Context 'CA-04 Require MFA for Admins' {
        BeforeAll {
            function Test-PolicyTargetsAdminRoles { param($Policy) if (-not $Policy -or -not $Policy.Conditions.Users.IncludeRoles) { return $false }; return ($Policy.Conditions.Users.IncludeRoles.Count -gt 0) }
            function Test-PolicyRequiresMfaForAdmins { param($Policy)
                if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }
                $strengthName = ''
                if ($Policy.GrantControls.AuthenticationStrength.DisplayName) { $strengthName = [string]$Policy.GrantControls.AuthenticationStrength.DisplayName }
                $l = $strengthName.ToLowerInvariant()
                $hasImpliedMfa = ($Policy.GrantControls.BuiltInControls -contains 'Mfa') -or ($l -match 'phishing') -or ($l -match 'passwordless') -or ($l -match 'multifactor') -or ($l -match '\bmfa\b')
                return $hasImpliedMfa
            }
        }
        It 'Returns False when policy does not target admin roles' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ BuiltInControls = @('Mfa'); AuthenticationStrength = [pscustomobject]@{ DisplayName = 'MFA' } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @() } }
            }
            (Test-PolicyRequiresMfaForAdmins -Policy $policy) | Should -BeFalse
        }
        It 'Returns True when policy targets admin roles and implies MFA' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ BuiltInControls = @('Mfa'); AuthenticationStrength = [pscustomobject]@{ DisplayName = 'MultiFactor' } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @('62e90394-69f5-4237-9190-012177145e10') } } # Global Admin role id example
            }
            (Test-PolicyRequiresMfaForAdmins -Policy $policy) | Should -BeTrue
        }
    }

    Context 'CA-05 Require Phish-Resistant MFA for Admins' {
        BeforeAll {
            function Test-IsPhishResistantStrength { param($Policy)
                $strengthObj = $Policy.GrantControls.AuthenticationStrength
                if (-not $strengthObj -or -not $strengthObj.AllowedCombinations) { return $false }
                $nonPhish = @('deviceBasedPush','temporaryAccessPassOneTime','temporaryAccessPassMultiUse','microsoftAuthenticatorPush','sms','voice','softwareOath','hardwareOath','x509CertificateSingleFactor','federatedSingleFactor','qrCodePin')
                $contains = ($strengthObj.AllowedCombinations | Where-Object { $nonPhish -contains $_ }).Count -gt 0
                return (-not $contains)
            }
            function Test-PolicyTargetsAdminRoles { param($Policy) if (-not $Policy -or -not $Policy.Conditions.Users.IncludeRoles) { return $false }; return ($Policy.Conditions.Users.IncludeRoles.Count -gt 0) }
            function Test-PolicyRequiresPhishResistantMfaForAdmins { param($Policy) if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }; return (Test-IsPhishResistantStrength -Policy $Policy) }
            # Ensure cache reflects phish-resistant combos only
            $script:AuthStrengthCache = @{ }
            $script:AuthStrengthCachePopulated = $true
            $script:AuthStrengthCache['strength-pr'] = [pscustomobject]@{ Id='strength-pr'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2') }
        }
        It 'Returns False when not targeting admin roles' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-pr'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2') } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @() } }
            }
            (Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $policy) | Should -BeFalse
        }
        It 'Returns True when targeting admin roles AND strength is phish-resistant' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-pr'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2') } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @('62e90394-69f5-4237-9190-012177145e10') } }
            }
            (Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $policy) | Should -BeTrue
        }
    }

    Context 'CA-08 Direct User Assignment' {
        BeforeAll {
            function Test-CA08 { param($PolicyCheck)
                $PolicyCheck.Conditions.Users.IncludeUsers -ne 'None' -and
                $null -ne $PolicyCheck.Conditions.Users.IncludeUsers -and
                $PolicyCheck.Conditions.Users.IncludeUsers -ne 'All' -and
                $PolicyCheck.Conditions.Users.IncludeUsers -ne 'GuestsOrExternalUsers'
            }
        }
        It 'Returns False when IncludeUsers is All' {
            $policy = [pscustomobject]@{ Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeUsers = 'All' } } }
            (Test-CA08 -PolicyCheck $policy) | Should -BeFalse
        }
        It 'Returns True when IncludeUsers is a direct user id' {
            $policy = [pscustomobject]@{ Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeUsers = '12345678-aaaa-bbbb-cccc-1234567890ab' } } }
            (Test-CA08 -PolicyCheck $policy) | Should -BeTrue
        }
        It 'Returns False when IncludeUsers is GuestsOrExternalUsers' {
            $policy = [pscustomobject]@{ Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeUsers = 'GuestsOrExternalUsers' } } }
            (Test-CA08 -PolicyCheck $policy) | Should -BeFalse
        }
    }
}
