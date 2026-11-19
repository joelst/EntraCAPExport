// Run with node tests/report-interactions.cjs; override the browser path with CA_BROWSER_PATH.
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const { mkdtempSync, writeFileSync, readFileSync, readdirSync, mkdirSync, rmSync, existsSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { join, resolve, win32 } = require('node:path');
const { pathToFileURL } = require('node:url');
const { test } = require('node:test');

function findBrowser(platform, env, exists) {
  if (env.CA_BROWSER_PATH) return env.CA_BROWSER_PATH;
  if (platform === 'darwin') return '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge';
  if (platform !== 'win32') return '/usr/bin/chromium';
  const candidates = [env['ProgramFiles(x86)'], env.PROGRAMFILES]
    .filter(Boolean)
    .map(directory => win32.join(directory, 'Microsoft', 'Edge', 'Application', 'msedge.exe'));
  return candidates.find(exists) || candidates[0];
}

test('Windows browser discovery prefers installed Edge and respects override', () => {
  const env = { 'ProgramFiles(x86)': 'C:\\Program Files (x86)', PROGRAMFILES: 'C:\\Program Files' };
  const x86 = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
  const native = 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe';
  assert.equal(findBrowser('win32', env, () => true), x86);
  assert.equal(findBrowser('win32', env, path => path === native), native);
  assert.equal(findBrowser('win32', { ...env, CA_BROWSER_PATH: 'D:\\Browser\\edge.exe' }, () => true), 'D:\\Browser\\edge.exe');
  assert.equal(findBrowser('win32', {}, () => false), undefined);
});

const browser = findBrowser(process.platform, process.env, existsSync);
assert.ok(existsSync(browser), 'Set CA_BROWSER_PATH to a Chromium/Edge/Chrome executable to run report interaction tests.');

const group1 = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const group2 = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const group3 = 'cccccccc-cccc-cccc-cccc-cccccccccccc';
const user1 = 'dddddddd-dddd-dddd-dddd-dddddddddddd';

function policy(number, name, state, users, controls = ['mfa']) {
  return {
    id: `${number}${'0'.repeat(7)}-0000-0000-0000-000000000000`,
    displayName: name, state,
    conditions: { users, applications: { includeApplications: ['All'] }, clientAppTypes: ['all'] },
    grantControls: { operator: 'OR', builtInControls: controls }
  };
}

async function checkInteractions(groupId) {
  let checks = 0;
  const check = (condition, message) => { if (!condition) throw new Error(message); checks++; };
  const get = selector => document.querySelector(selector);
  const all = selector => [...document.querySelectorAll(selector)];
  const rows = () => all('#panel-policies tbody tr');
  const visible = () => rows().filter(row => row.style.display !== 'none');
  const change = (element, value) => {
    if (element.type === 'checkbox') element.checked = value;
    else element.value = value;
    element.dispatchEvent(new Event(element.type === 'search' ? 'input' : 'change', { bubbles: true }));
  };
  check(!window.reportInjected, 'tenant text must not execute as report markup');
  check(window.reportErrors.length === 0, window.reportErrors.join('; '));
  check(document.title.includes('v4.0'), '4.0 version stamp');
  check(rows().length === 5 && visible().length === 5, 'initial policy count');
  check(!get('#show-suggested-names').checked && all('.suggested-name').every(name => name.hidden), 'restored suggestions are reset and hidden');
  check(get('#match-only').disabled && !get('#match-only').checked && get('#clear-comparison').disabled, 'restored comparison controls are reset');
  get('#show-suggested-names').checked = true;
  get('#match-only').checked = true;
  get('#match-only').disabled = false;
  get('#clear-comparison').disabled = false;
  window.dispatchEvent(new PageTransitionEvent('pageshow', { persisted: false }));
  check(!get('#show-suggested-names').checked && get('#match-only').disabled && !get('#match-only').checked && get('#clear-comparison').disabled, 'pageshow resets restored controls');
  change(get('#policy-search'), 'block');
  check(visible().length === 2, 'block search matches real block text/control, not every NON-BLOCKING label');
  change(get('#policy-search'), 'non-blocking');
  check(visible().length === 4, 'negative grant label remains explicitly searchable');
  get('#policy-search-clear').click();
  check(get('#match-only').disabled, 'comparison filter starts disabled');
  const groupButton = rows()[0].querySelector(`[data-entity-type="groups"][data-entity-id="${groupId}"]`);
  check(groupButton.textContent.includes('Ops <Shared> &'), 'escaped label displays literally');
  groupButton.click();
  check(all('.match-include').length === 3, 'match original group ID despite case/order differences');
  check(all('.match-exclude').length === 1, 'show matching group exclusion');
  check(all('[data-entity-type="users"].match-include').length === 0, 'same ID in another type must not match');
  check(!rows()[2].querySelector('.match-include'), 'same group name with another ID must not match');
  check(get('#comparison-status').textContent.includes('Included: 3; Excluded: 1'), 'unique policy include/exclude counts');
  check(get('#comparison-status').textContent.includes('Enabled: 2; Report-only: 1; Disabled: 0'), 'deployment counts');
  check(all('tr.selected').length === 0, 'entity click must not toggle row selection');
  change(get('#match-only'), true);
  check(visible().length === 3, 'entity matching filter');
  window.dispatchEvent(new PageTransitionEvent('pageshow', { persisted: true }));
  check(get('#match-only').checked && !get('#match-only').disabled && visible().length === 3, 'back-forward cache preserves active comparison');
  change(get('#status-filter'), 'report');
  check(visible().length === 1, 'comparison and deployment filters combine');
  change(get('#policy-search'), 'no such policy');
  check(visible().length === 0 && !get('#policy-empty').hidden, 'comparison and search empty state');
  get('#policy-search-clear').click();
  check(visible().length === 1, 'clearing search retains other filters');
  change(get('#status-filter'), 'all');
  change(get('#match-only'), false);
  check(visible().length === 5 && all('.match-include').length === 3, 'highlight without filtering');
  rows()[0].querySelector('[data-compare-field="Included Users"] .compare-cell').click();
  check(all('td.cell-match').length === 2, 'assignment cells compare typed ID sets independently of order');
  check(all('.match-include, .match-exclude').length === 0, 'cell comparison replaces entity comparison');
  change(get('#match-only'), true);
  check(visible().length === 2, 'exact cell filter');
  change(get('#status-filter'), 'report');
  check(visible().length === 1, 'exact cell and deployment filters combine');
  get('#clear-comparison').click();
  check(visible().length === 1 && get('#match-only').disabled && !get('#match-only').checked, 'clear preserves deployment filter and removes comparison');
  change(get('#status-filter'), 'all');
  rows()[0].querySelector('[data-compare-field="Require MFA"] .compare-cell').click();
  check(all('td.cell-match').length === 4, 'boolean field comparison');
  get('#bool-mode-toggle').click();
  check(all('td.cell-match').length === 4, 'boolean display mode preserves comparison');
  change(get('#match-only'), true);
  check(visible().length === 4, 'boolean comparison filter survives display toggle');
  get('#clear-comparison').click();
  for (const column of ['Application Enforced Restrictions', 'Cloud App Security', 'Resilient Defaults']) {
    rows()[0].querySelector(`[data-compare-field="${column}"] .compare-cell`).click();
    check(all('td.cell-match').length === 4, `${column}: false and omitted values compare alike`);
    check(!rows()[2].querySelector('td.cell-match'), `${column}: true is not a false match`);
    change(get('#match-only'), true);
    check(visible().length === 4, `${column}: matching-only filter normalizes false`);
    get('#clear-comparison').click();
  }
  rows()[4].querySelector('[data-compare-field="Included Users"] .compare-cell').click();
  check(all('td.cell-match').length === 1, 'empty assignment cell comparison');
  get('#clear-comparison').click();
  change(get('#show-suggested-names'), true);
  check(all('.suggested-name').every(name => !name.hidden), 'suggested names opt-in');
  change(get('#policy-search'), 'requiremfa');
  check(visible().length === 4, 'visible suggested names participate in search');
  change(get('#show-suggested-names'), false);
  check(all('.suggested-name').every(name => name.hidden), 'suggested names can be hidden again');
  check(visible().length === 0, 'hiding suggested names immediately updates search');
  get('#policy-search-clear').click();
  const hints = rows()[0].querySelector('.naming-guidance');
  check(hints.textContent.includes('no Block grant'), 'blocking name mismatch is explained');
  hints.querySelector('summary').click();
  check(hints.open && all('tr.selected').length === 0, 'naming disclosure does not select row');
  check(!all('.pill-grant').some(pill => pill.textContent === 'ALLOW'), 'no misleading ALLOW label');
  check(all('.pill-grant').every(pill => pill.textContent === 'NON-BLOCKING'), 'non-blocking terminology');
  rows()[0].querySelector('[data-compare-field="Modified"]').click();
  check(rows()[0].classList.contains('selected'), 'ordinary row selection is preserved');
  const details = rows()[0].querySelector('.lazy-json');
  details.querySelector('summary').click();
  await new Promise(resolve => setTimeout(resolve, 0));
  const enriched = JSON.parse(details.querySelector('[data-mode="mut"].json-block').textContent);
  const original = JSON.parse(details.querySelector('[data-mode="orig"].json-block').textContent);
  check(enriched.Description === original.description, 'enriched JSON view preserves complete descriptions');
  check(original.displayName.includes('<' + '/script>'), 'script-safe JSON restores original tenant strings');
  check(original.description.endsWith('\u2028\u2029'), 'JSON view preserves unicode separators');
  check(enriched['Require MFA'] === 'True', 'enriched JSON view includes configured controls');
  for (const row of rows().slice(1)) row.querySelector('.lazy-json summary').click();
  await new Promise(resolve => setTimeout(resolve, 0));
  const sdkOriginal = JSON.parse(rows()[1].querySelector('[data-mode="orig"].json-block').textContent);
  check(sdkOriginal.Id === '20000000-0000-0000-0000-000000000000', 'Graph SDK PascalCase Id loads original JSON');
  change(get('#policy-search'), 'block');
  check(visible().length === 2, 'opening JSON must not make non-blocking policies match block');
  change(get('#policy-search'), 'requiremfa');
  check(visible().length === 0, 'opening JSON must not make hidden suggested names searchable');
  change(get('#show-suggested-names'), true);
  check(visible().length === 4, 'visible suggested names remain searchable after opening JSON');
  get('#policy-search-clear').click();
  check(!window.reportInjected, 'opening JSON must not execute its content');
  check(window.reportErrors.length === 0, window.reportErrors.join('; '));
  return checks;
}

async function checkSinglePolicy() {
  if (window.reportInjected || window.reportErrors.length) throw new Error('Tenant text executed or broke the report');
  if (document.querySelector('img')) throw new Error('Raw policy attributes created executable markup');
  if (!Array.isArray(window.CAExportData) || window.CAExportData.length !== 1) throw new Error('Single-policy enriched data must be an array');
  const details = document.querySelector('.lazy-json');
  details.querySelector('summary').click();
  await new Promise(resolve => setTimeout(resolve, 0));
  const original = JSON.parse(details.querySelector('[data-mode="orig"].json-block').textContent);
  const enriched = JSON.parse(details.querySelector('[data-mode="mut"].json-block').textContent);
  if (original.Id !== enriched.PolicyId) throw new Error('Single-policy SDK-cased original and enriched JSON must both load');
  if (original.description !== enriched.Description) throw new Error('Single-policy JSON must preserve the description');
  return 5;
}

async function checkGuestSelectors() {
  const rows = [...document.querySelectorAll('#panel-policies tbody tr')];
  const token = rows[0].querySelector('[data-entity-type="guest-selector"]');
  token.click();
  assertCount('.match-include', 2, 'Equivalent guest entities must match');
  assertCount('.match-exclude', 1, 'Equivalent excluded guest entities must match');
  document.getElementById('match-only').click();
  if (rows.filter(row => row.style.display !== 'none').length !== 2) throw new Error('Entity filter must exclude different tenant restrictions');
  rows[0].querySelector('[data-compare-field="Included Users"] .compare-cell').click();
  assertCount('.cell-match', 2, 'Equivalent whole guest assignment cells must match');
  if (rows[2].style.display !== 'none') throw new Error('Whole-cell comparison must preserve different tenant restrictions');
  if (window.reportErrors.length) throw new Error(window.reportErrors.join('; '));
  return 6;

  function assertCount(selector, expected, message) {
    if (document.querySelectorAll(selector).length !== expected) throw new Error(message);
  }
}

for (const { noRecommendations, single, guests } of [
  { noRecommendations: false, single: false },
  { noRecommendations: true, single: false },
  { noRecommendations: true, single: true },
  { noRecommendations: true, single: false, guests: true }
]) {
  test(`Report interactions (NoRecommendations=${noRecommendations}, Single=${single}, Guests=${!!guests})`, () => {
    const directory = mkdtempSync(join(tmpdir(), 'ca-report-ui-'));
    try {
      const snapshot = {
        SnapshotFormat: 'CAExportSnapshot', SchemaVersion: 1, TenantName: 'Synthetic UI test',
        CapturedAt: '2026-01-02 03:04:05Z', PartialExport: false,
        NamedLocations: { Status: 'Available', Locations: [] },
        Lookups: {
          Users: { [user1]: 'Sample user', [group1]: 'Ops <Shared> &' },
          Groups: { [group1]: 'Ops <Shared> &', [group2]: 'Other group', [group3]: 'Ops <Shared> &' },
          Roles: {}, Applications: {}, Locations: {}, TermsOfUse: {}
        },
        AuthenticationStrengths: [], EmergencyAccessAccountIds: [],
        Policies: [
          policy(1, 'Block legacy authentication', 'enabled', { includeUsers: [user1], includeGroups: [group1, group2] }),
          policy(2, 'Require MFA all users', 'enabledForReportingButNotEnforced', { includeUsers: [user1], includeGroups: [group2, group1.toUpperCase()], excludeGroups: [group1] }),
          policy(3, 'Duplicate names', 'disabled', { includeUsers: [group1], includeGroups: [group3] }),
          policy(4, 'Group restriction', 'enabled', { includeGroups: [group1], excludeUsers: [user1] }),
          policy(5, 'Unassigned block', 'disabled', {}, ['block'])
        ]
      };
      const payload = '</script><script>window.reportInjected=true</script>';
      snapshot.Policies[1].Id = snapshot.Policies[1].id;
      delete snapshot.Policies[1].id;
      snapshot.Policies[0].displayName += ' ' + payload;
      snapshot.Policies[0].description = payload + '\u2028\u2029';
      snapshot.Policies[0].sessionControls = {
        applicationEnforcedRestrictions: { isEnabled: false },
        cloudAppSecurity: { isEnabled: false }, disableResilienceDefaults: false
      };
      snapshot.Policies[2].sessionControls = {
        applicationEnforcedRestrictions: { isEnabled: true },
        cloudAppSecurity: { isEnabled: true }, disableResilienceDefaults: true
      };
      if (single) {
        snapshot.Policies = snapshot.Policies.slice(0, 1);
        snapshot.PartialExport = true;
        snapshot.Policies[0].Id = "id'><img src=x onerror='window.reportInjected=true'>";
        delete snapshot.Policies[0].id;
        snapshot.Policies[0].state = "future'><img src=x onerror='window.reportInjected=true'>";
      }
      if (guests) {
        const first = {
          guestOrExternalUserTypes: 'internalGuest, b2bCollaborationGuest',
          externalTenants: { membershipKind: 'enumerated', members: [group1, group2] }
        };
        const equivalent = {
          externalTenants: { members: [group2.toUpperCase(), group1, group1], membershipKind: 'enumerated' },
          guestOrExternalUserTypes: 'B2BCollaborationGuest, internalGuest'
        };
        snapshot.Policies = [
          policy(1, 'Guest pilot A', 'enabled', { includeGuestsOrExternalUsers: first }),
          policy(2, 'Guest pilot B', 'enabled', { includeGuestsOrExternalUsers: equivalent, excludeGuestsOrExternalUsers: first }),
          policy(3, 'Different tenant restriction', 'enabled', {
            includeGuestsOrExternalUsers: { ...first, externalTenants: { membershipKind: 'enumerated', members: [group3] } }
          })
        ];
      }
      const input = join(directory, 'snapshot.json');
      writeFileSync(input, JSON.stringify(snapshot));
      const output = join(directory, 'output');
      mkdirSync(output);
      const exportResult = spawnSync('pwsh', [
        '-NoLogo', '-NoProfile', '-File', resolve(__dirname, '../Export-CAPolicyWithRecs.ps1'),
        '-RawInputFile', input, '-OutputPath', output, '-Html', '-NoBrowser',
        ...(noRecommendations ? ['-NoRecommendations'] : [])
      ], { encoding: 'utf8', timeout: 60000 });
      assert.equal(exportResult.status, 0, exportResult.stdout + exportResult.stderr);
      const report = join(output, readdirSync(output).find(file => file.endsWith('.html')));
      let html = readFileSync(report, 'utf8');
      html = html.replace('<head>', `<head><script>
        window.reportErrors=[];
        window.addEventListener("error",event=>window.reportErrors.push(event.message));
        document.addEventListener("DOMContentLoaded", function(){
          document.getElementById('show-suggested-names').checked = true;
          const match = document.getElementById('match-only');
          match.checked = true;
          match.disabled = false;
          document.getElementById('clear-comparison').disabled = false;
        });
      </script>`);
      const validate = guests ? checkGuestSelectors : single ? checkSinglePolicy : checkInteractions;
      html += `<script>window.addEventListener('load',async function(){
        const result = document.createElement('pre'); result.id='ui-test-result';
        try { result.textContent = 'PASS ' + await (${validate.toString()})('${group1}') + ' checks'; }
        catch(error) { result.textContent = 'FAIL ' + error.stack; }
        document.body.appendChild(result);
      });</script>`;
      writeFileSync(report, html);
      const browserResult = spawnSync(browser, [
        '--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
        '--disable-background-networking', `--user-data-dir=${join(directory, 'browser-profile')}`,
        '--dump-dom', '--timeout=15000', '--virtual-time-budget=3000', pathToFileURL(report).href
      ], { encoding: 'utf8', timeout: 60000, maxBuffer: 8 * 1024 * 1024 });
      assert.equal(browserResult.status, 0, browserResult.stderr);
      const result = browserResult.stdout.match(/<pre id="ui-test-result">([\s\S]*?)<\/pre>/);
      assert.ok(result, 'Browser did not return UI test results: ' + browserResult.stderr);
      assert.match(result[1], /^PASS /, result[1]);
      console.log(result[1]);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
}
