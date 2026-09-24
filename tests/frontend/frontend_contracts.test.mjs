import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const read = name => readFileSync(new URL(`../../dashboard/${name}`, import.meta.url), 'utf8');
const editor = read('editor.html');
const index = read('index.html');
const converter = read('converter.html');
const help = read('help.html');
const pptxGuide = read('pptx-guide.html');
const resetPassword = read('reset-password.html');
const headers = read('_headers');
const readme = readFileSync(new URL('../../README.md', import.meta.url), 'utf8');

function inlineScripts(html) {
  return [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)].map(match => match[1]);
}

test('inline frontend scripts compile', () => {
  for (const [name, html] of [
    ['editor.html', editor],
    ['index.html', index],
    ['converter.html', converter],
    ['reset-password.html', resetPassword],
    ['apple-welcome.html', read('apple-welcome.html')],
    ['apple-delete.html', read('apple-delete.html')],
  ]) {
    const scripts = inlineScripts(html);
    assert.ok(scripts.length, `${name} should contain an inline script`);
    scripts.forEach((source, i) => assert.doesNotThrow(
      () => new vm.Script(source, { filename: `${name}#script-${i}` })
    ));
  }
});

test('polling is sequential and bounded on both generation surfaces', () => {
  for (const html of [editor, converter]) {
    assert.doesNotMatch(html, /\bsetInterval\s*\(/);
    assert.match(html, /POLL_TIMEOUT_MS\s*=\s*10\s*\*\s*60\s*\*\s*1000/);
    assert.match(html, /pollTimer\s*=\s*setTimeout\(poll,\s*POLL_INTERVAL_MS\)/);
    assert.match(html, /if\s*\(!res\.ok\)/);
    assert.match(html, /\['processing', 'complete', 'error'\]\.includes\(data\.status\)/);
  }
});

test('editor conflict and recovery dialogs preserve local work and one-time secrets', () => {
  assert.match(editor, /cancelValue:\s*'keep'/);
  assert.match(editor, /key:\s*'keep',\s*label:\s*'Keep editing'/);
  assert.match(editor, /key:\s*'reload',\s*label:\s*'Discard local changes'/);
  assert.match(editor, /key:\s*'overwrite',\s*label:\s*'Overwrite server'/);
  assert.match(editor, /function showRecoveryCode[\s\S]*?dismissible:\s*false/);
  assert.match(editor, /const modalQueue\s*=\s*\[\]/);
  assert.match(editor, /beforeunload[\s\S]*?activeModal\s*&&\s*!activeModal\.dismissible/);
});

test('output gating retains checkbox preferences and resyncs membership changes', () => {
  assert.doesNotMatch(editor, /if\s*\(!ok\)\s*box\.checked\s*=\s*false/);
  assert.match(editor, /function togglePlayIncluded[\s\S]*?syncGenerateChecks\(\)/);
  assert.match(editor, /!chkOffenseCoach\.disabled\s*&&\s*chkOffenseCoach\.checked/);
});

test('drawing tools keep independent defaults', () => {
  assert.match(editor, /const toolStyles\s*=\s*\{/);
  assert.match(editor, /route:\s*\{\s*color:\s*'#FF0000',\s*dash:\s*false/);
  assert.match(editor, /line:\s*\{\s*color:\s*'#000000',\s*dash:\s*true,\s*end:\s*'arrow'/);
  assert.match(editor, /if\s*\(toolStyles\[t\]\)\s*style\s*=\s*\{\s*\.\.\.toolStyles\[t\]\s*\}/);
});

test('imports and permanent save failures are bounded', () => {
  assert.match(editor, /MAX_IMPORT_BYTES\s*=\s*1000000/);
  assert.match(editor, /file\.size\s*>\s*MAX_IMPORT_BYTES/);
  assert.match(editor, /err\.permanent\s*=\s*res\.status\s*>=\s*400/);
  assert.match(editor, /Save rejected/);
});

test('unsaved editor state is scoped to an immutable account ID', () => {
  assert.match(editor, /docOwner\s*=\s*null;\s*\/\/ immutable user ID/);
  assert.match(editor, /docOwner\s*===\s*userId/);
  assert.doesNotMatch(editor, /docOwner\s*===\s*email/);
  assert.match(editor, /ownerId:\s*savingOwner/);
  assert.match(editor, /activeUserId\s*!==\s*savingOwner/);
});

test('account loads and first saves are race-safe', () => {
  assert.match(editor, /let authEpoch\s*=\s*0/);
  assert.match(editor, /playbookLoadRequest\.abort\(\)/);
  assert.match(editor, /epoch\s*!==\s*authEpoch\s*\|\|\s*activeUserId\s*!==\s*userId/);
  assert.match(editor, /if\s*\(!overwriteNext\)\s*body\.baseUpdatedAt\s*=\s*serverUpdatedAt/);
});

test('editor exposes an accessible named-playbook manager and immutable save identity', () => {
  assert.match(editor, /<label id="playbookManagerLabel" for="playbookSelect">Playbook<\/label>/);
  assert.match(editor, /<select id="playbookSelect" disabled><\/select>/);
  assert.match(editor, /id="newPlaybookBtn"[^>]*>\+ New<\/button>/);
  assert.match(editor, /id="renamePlaybookBtn"[^>]*>Rename<\/button>/);
  assert.match(editor, /id="playbookStatus" role="status" aria-live="polite"/);
  assert.match(editor, /\.playbook-manager, \.new-play-format \{ grid-column: 1 \/ -1;/);
  assert.match(editor, /function showPlaybookForm[\s\S]*?label\.htmlFor = 'playbookNameInput'/);
  assert.match(editor, /function showPlaybookForm[\s\S]*?document\.createElement\('fieldset'\)/);
  assert.match(editor, /t\.tagName === 'SELECT'/);
  assert.match(editor, /function captureDocumentContext\(\)[\s\S]*?playbookId: docPlaybookId[\s\S]*?epoch: documentEpoch/);
  assert.match(editor, /function saveDoc\(\)[\s\S]*?captureDocumentContext\(\)[\s\S]*?doSave\(context\)/);
  assert.match(editor, /fetch\(playbookDocumentUrl\(savingPlaybookId\)/);
  assert.match(editor, /body\.baseRevision = serverRevision/);
  assert.match(editor, /body\.forceOverwrite = true/);
  assert.doesNotMatch(editor, /fetch\(['"]\/api\/plays['"]/);
});

test('authentication and account actions cannot apply stale responses', () => {
  assert.match(editor, /let authBusy\s*=\s*false/);
  assert.match(editor, /if\s*\(authBusy\)\s*return/);
  assert.match(editor, /verifyAuthenticatedAccount\(data\.userId,\s*epoch\)/);
  assert.match(editor, /menuEpoch\s*!==\s*authEpoch\s*\|\|\s*menuUserId\s*!==\s*activeUserId/);
  assert.match(editor, /setAccountRequestBusy\(true\)/);
  assert.match(editor, /requestDeletionUntilSettled/);
  assert.match(editor, /res\.status\s*===\s*423\s*&&\s*data\.deletionPending/);
  assert.match(editor, /res\.status\s*===\s*401\s*&&\s*sawPending/);
});

test('download labels are built without an HTML injection sink', () => {
  assert.doesNotMatch(converter, /\.innerHTML\s*=/);
  assert.match(converter, /a\.appendChild\(document\.createTextNode/);
});

test('converter presents bounded PPTX assumptions separately from errors', () => {
  assert.match(converter, /id="conversionWarning" role="status" aria-live="polite"/);
  assert.match(converter, /function showWarnings\(warnings\)/);
  assert.match(converter, /warning\.code === 'assumed_offense_before_defense'/);
  assert.match(converter, /warning\.code === 'skipped_before_first_divider'/);
  assert.match(converter, /warning\.code === 'skipped_no_field_rectangle'/);
  assert.match(converter, /Number\.isSafeInteger\(count\)/);
  assert.match(converter, /conversionWarning\.textContent\s*=/);
  assert.match(converter, /showDownloads\(jobId, data\.files\);\s*showWarnings\(data\.warnings\);/);
  assert.match(converter, /if \(clearUI\)[\s\S]*?hideWarnings\(\)/);
  assert.doesNotMatch(converter, /errorMsg\.textContent\s*=\s*[^;]*assumed_offense_before_defense/);
});

test('responsive and keyboard contracts stay present', () => {
  assert.match(editor, /@media \(max-width: 820px\)/);
  assert.match(editor, /\.renum-btn \{ width: auto;/);
  assert.match(editor, /@media \(hover: none\) \{ \.play-row\.active \.row-btns \{ display: flex; \} \}/);
  assert.match(index, /\.auth-layout > \.card \{ order: -1;/);
  assert.match(converter, /id="dropZone" role="button" tabindex="0"/);
  assert.match(converter, /dropZone\.addEventListener\('keydown'/);
  assert.match(editor, /role="dialog" aria-modal="true"/);
});

test('home is an authenticated two-choice hub, not an upload surface', () => {
  const converterChoice = index.match(/<a class="choice converter"[\s\S]*?<\/a>/);

  assert.match(index, /<title>GSS Playbook Editor/);
  assert.match(index, /<form id="signinForm"/);
  assert.match(index, /authSubmit\('\/api\/auth\/login'\)/);
  assert.match(index, /authSubmit\('\/api\/auth\/register'\)/);
  assert.match(index, /fetch\('\/api\/auth\/recover'/);
  assert.match(index, /class="choice editor-choice" href="\/editor"/);
  assert.match(index, /class="choice converter" href="\/converter"/);
  assert.ok(converterChoice, 'missing PowerPoint (PPTX) Import choice');
  assert.doesNotMatch(converterChoice[0], /pptx-guide/);
  assert.match(index, /class="choice editor-choice"[\s\S]*?<h2>Playbook Editor<\/h2>[\s\S]*?Choose or create plays &rarr;/);
  assert.match(index, /class="choice converter"[\s\S]*?<h2>PowerPoint \(PPTX\) Import<\/h2>[\s\S]*?Use a PowerPoint playbook &rarr;/);
  assert.match(index, /\.choice-grid\s*\{[^}]*grid-template-columns: minmax\(0, 1fr\);[^}]*max-width: 620px;/);
  assert.match(index, /\.hub-sub\s*\{[^}]*font-size: 1rem;[^}]*line-height: 1\.6;/);
  assert.match(index, /\.hub-help\s*\{[^}]*font-size: 1rem;[^}]*line-height: 1\.6;/);
  assert.match(index, /<div class="hub-help">[\s\S]*?<p>Need help\?<\/p>[\s\S]*?quick-start guide[\s\S]*?href="\/pptx-guide">Read the PowerPoint \(PPTX\) Import Guide<\/a>/);
  assert.match(index, /<header class="topbar">[\s\S]*?id="hubAccount"[^>]*hidden>[\s\S]*?id="userEmail"[\s\S]*?id="signOutBtn"[\s\S]*?<\/header>/);
  assert.match(index, /\.account-email\s*\{[^}]*min-width: 0;[^}]*overflow: hidden;[^}]*text-overflow: ellipsis;[^}]*white-space: nowrap;/);
  assert.match(index, /@media \(max-width: 620px\)[\s\S]*?\.top-account \{ flex: 1 0 100%; justify-content: flex-end;/);
  assert.match(index, /function showAuth\([\s\S]*?\$\('hubAccount'\)\.hidden = true;/);
  assert.match(index, /function showHub\([\s\S]*?\$\('hubAccount'\)\.hidden = false;/);
  assert.doesNotMatch(index, /id="dropZone"/);
  assert.doesNotMatch(index, /['"]\/api\/upload['"]/);
});

test('home presents account creation as a clear, dedicated mode', () => {
  assert.match(index, /id="startRegisterBtn"[^>]*>Create a free account<\/button>/);
  assert.match(index, /function setAuthMode\(mode, focusHeading = false\)/);
  assert.match(index, /Create your free account/);
  assert.match(index, /recovery code after signup/i);
  assert.match(index, /authPassword\.autocomplete = registering \? 'new-password' : 'current-password'/);
  assert.match(index, /authPassword\.minLength = 8/);
  assert.match(index, /if \(authMode === 'register'\)[\s\S]*?authSubmit\('\/api\/auth\/register'\)/);
  assert.match(index, /Already have an account\?[^<]*<a id="signinInsteadLink"/);
});

test('PowerPoint import naming and the allowlisted post-auth return stay consistent', () => {
  assert.match(index, /id="converterSigninNotice"[^>]*hidden>[^<]*Sign in or create an account first\.[\s\S]*?return you automatically to <strong>PowerPoint \(PPTX\) Import<\/strong>/);
  assert.match(index, /const POST_AUTH_DESTINATIONS = Object\.freeze\(\{ converter: '\/converter', editor: '\/editor', apple: '\/apple-welcome' \}\)/);
  assert.match(index, /Object\.prototype\.hasOwnProperty\.call\([\s\S]*?POST_AUTH_DESTINATIONS,[\s\S]*?requestedPostAuthDestination[\s\S]*?\? POST_AUTH_DESTINATIONS\[requestedPostAuthDestination\] : null/);
  assert.match(index, /function routeAfterAuthentication\(account, focusHeading = false\)[\s\S]*?window\.location\.replace\(postAuthDestination\)[\s\S]*?showHub\(account, focusHeading\)/);
  assert.doesNotMatch(index, /window\.location\.assign\(postAuthDestination\)/);
  assert.match(index, /async function authSubmit\(path\)[\s\S]*?routeAfterAuthentication\(verified, true\)/);
  assert.match(index, /\$\('recoverForm'\)\.addEventListener[\s\S]*?routeAfterAuthentication\(verified, true\)/);
  assert.match(index, /async function init\(\)[\s\S]*?routeAfterAuthentication\(data\)/);
  assert.match(index, /if \(requestedPostAuthDestination === 'apple'\)\s*\{\s*showAuth\(\);\s*return;/,
    'explicit old-account proof must show the login form instead of looping on a stale session');

  assert.match(converter, /<title>GSS Playbook Editor PowerPoint \(PPTX\) Import/);
  assert.match(converter, /<h1>GSS Playbook Editor &mdash; PowerPoint \(PPTX\) Import<\/h1>/);
  assert.match(converter, /id="generateBtn" disabled>Create coach card<\/button>/);
  assert.match(converter, /href="\/pptx-guide">&#128209; PowerPoint \(PPTX\) Import Guide<\/a>/);
  assert.match(pptxGuide, /<h1>PowerPoint \(PPTX\) Import Guide<\/h1>/);
  assert.match(pptxGuide, /class="converter-cta" href="\/converter">Open PowerPoint \(PPTX\) Import &rarr;<\/a>/);
});

test('optional Apple sign-in leaves the email form and signup visible', () => {
  assert.match(index, /\$\('legacyLogin'\)\.open = true;/);
  assert.match(index, /\$\('signupPrompt'\)\.hidden = false;/);
  assert.doesNotMatch(index, /\$\('signupPrompt'\)\.hidden = [^;]*appleEnabled/);
  assert.match(index, /Or use email and password/);
  assert.match(help, /either Apple or your existing email and password/);
});

test('playbook deletion and portable backup guidance matches the named-book model', () => {
  for (const doc of [help, readme]) {
    assert.match(doc, /original playbook[^.]*cannot be deleted|original can instead be renamed and reused/i);
    assert.match(doc, /exact name/i);
    assert.match(doc, /backup[^.]*source playbook name/i);
    assert.match(doc, /another coach/i);
    assert.match(doc, /another playbook or account/i);
    assert.match(doc, /Export(?:[^.]*before| first)/i);
  }
});

test('editor keeps 5v5 and 6v6 formats per play', () => {
  assert.match(editor, /const OFFENSE_5_CHIP_KEYS = \['1', '2', '3', 'C', 'QB'\]/);
  assert.match(editor, /const DEFENSE_5_CHIP_KEYS = \['1', '2', '3', '4', '5'\]/);
  assert.match(editor, /isFivePlayerDefense && routeChip === 'N'/);
  assert.match(editor, /const PLAYER_FORMATS = \{/);
  assert.match(editor, /const DEFAULT_NEW_PLAYERS_PER_SIDE = 5/);
  assert.match(editor, /const playersPerSide = normalizePlayersPerSide\(p\.playersPerSide, 6\)/);
  assert.match(editor, /function inferDefaultPlayersPerSide\(d, normalizedDoc\)/);
  assert.match(editor, /if \(five \+ six === 0\) return DEFAULT_NEW_PLAYERS_PER_SIDE/);
  assert.match(editor, /return five > six \? 5 : 6/);
  assert.match(editor, /out\.defaultPlayersPerSide = inferDefaultPlayersPerSide\(d, out\)/);
  assert.match(editor, /function makePlay\(section\)[\s\S]*?DEFAULT_NEW_PLAYERS_PER_SIDE/);
  assert.match(editor, /Saved for next time\. Existing plays keep their format\./);
  assert.match(editor, /function setDefaultPlayersPerSide\(value\)[\s\S]*?doc\.defaultPlayersPerSide = count;[\s\S]*?markDirty\(\)/);
  assert.match(editor, /className = 'format-badge'/);
  assert.match(editor, /title: 'What format are you coaching\?'/);
  assert.match(editor, /choices:\s*\[[\s\S]*?label: '5v5'[\s\S]*?label: '6v6'/);
  assert.match(editor, /dismissible:\s*false/);
  assert.match(editor, /schema:\s*2/);
});

test('home preserves one-time recovery codes until explicit acknowledgement', () => {
  assert.match(index, /function showRecoveryCode/);
  assert.match(index, /recoveryModalOpen\s*=\s*true/);
  assert.match(index, /if \(event\.key === 'Escape'\)[\s\S]*?event\.preventDefault\(\)/);
  assert.match(index, /saveRecoveryBtn\.addEventListener\('click'/);
  assert.match(index, /beforeunload[\s\S]*?recoveryModalOpen/);
  assert.match(index, /const focusable = \[recoveryCode, copyRecoveryBtn, saveRecoveryBtn\]/);
  assert.match(index, /document\.querySelector\('main'\)\.inert = true/);
  assert.match(index, /verifyAccount\(data\.userId, epoch\)/);
});

test('email password reset is primary while the offline code stays available', () => {
  assert.match(index, /id="forgotLink" href="\/\?forgot=1">Forgot password\?<\/a>/);
  assert.match(index, /id="resetRequestForm"/);
  assert.match(index, /Email me a reset link/);
  assert.match(index, /fetch\('\/api\/auth\/password-reset\/request'/);
  assert.match(index, /If an active account exists for that email/);
  assert.match(index, /id="useRecoveryCodeLink"[^>]*>Have a recovery code\? Use it instead<\/a>/);
  assert.match(index, /fetch\('\/api\/auth\/recover'/);
  assert.match(editor, /id="emailResetLink" href="\/\?forgot=1" target="_blank" rel="noopener"/);
  assert.match(editor, /Use an offline recovery code instead/);
  assert.doesNotMatch(help, /recovery code[^.]*the <strong>only<\/strong> way/i);
});

test('emailed reset credentials stay in the fragment and are cleared before network work', () => {
  assert.match(resetPassword, /new URLSearchParams\(window\.location\.hash\.slice\(1\)\)/);
  assert.match(resetPassword, /history\.replaceState\(null, '', window\.location\.pathname\)/);
  assert.match(resetPassword, /\^\[A-Za-z0-9_-\]\{43\}\$/);
  assert.match(resetPassword, /id="newPassword"[^>]*autocomplete="new-password"/);
  assert.match(resetPassword, /id="confirmPassword"[^>]*autocomplete="new-password"/);
  assert.match(resetPassword, /fetch\('\/api\/auth\/password-reset\/complete'/);
  assert.match(resetPassword, /body: JSON\.stringify\(\{ email: resetEmail, token: resetToken, newPassword \}\)/);
  assert.match(resetPassword, /resetToken = ''/);
  assert.match(resetPassword, /id="savedButton"[^>]*>I saved it/);
  assert.match(resetPassword, /beforeunload[\s\S]*?resetComplete/);
  assert.doesNotMatch(resetPassword, /(?:localStorage|sessionStorage|\.innerHTML\s*=)/);
  assert.match(headers, /\/reset-password\s+[\s\S]*?Cache-Control: no-store[\s\S]*?X-Robots-Tag: noindex, nofollow/);
});

test('cold signed-out editor visits use home while in-editor reauthentication stays local', () => {
  assert.match(editor, /finish-deletion/);
  assert.match(editor, /res\.status === 401 && !finishingDeletion/);
  assert.match(editor, /window\.location\.replace\('\/'\)/);
  assert.match(editor, /function showAuthView/);
  assert.match(editor, /doc && saveState !== 'saved' && docOwner === userId/);
  assert.match(index, /href="\/editor\?finish-deletion=1"/);
});

test('converter stays hidden until auth succeeds and returns signed-out users home', () => {
  assert.match(converter, /id="converterApp" hidden/);
  assert.match(converter, /fetch\('\/api\/auth\/me', \{ cache: 'no-store' \}\)/);
  assert.match(converter, /res\.status === 401/);
  assert.match(converter, /const target = '\/\?signin=converter'/);
  assert.match(converter, /window\.location\.replace\(target\)/);
  assert.match(converter, /converterApp\.hidden = false/);
  assert.match(converter, /generationFetch\('\/api\/upload'/);
});

test('help names the Arrow and Block controls shown by the editor', () => {
  assert.match(help, /Choose <strong>Arrow<\/strong>/);
  assert.match(help, /<strong>Block<\/strong>/);
  assert.match(help, /Arrow \/ Ball \/ Block \/ None/);
});

test('coach help explains the two workflows, printing, and safe sharing in plain language', () => {
  for (const html of [help, pptxGuide, converter]) {
    assert.match(html, /printable PDFs, not editable plays/);
    assert.match(html, /laminate before cutting/i);
  }
  assert.match(help, /This shares a copy, not a live shared playbook/);
  assert.match(help, /Cut around each whole insert, not each individual play/);
  assert.match(help, /Fewer plays per page means larger drawings/);
  assert.match(help, /reprint old cards if their numbers change/);
  assert.match(help, /Email delivery is not yet reliable/);
  for (const html of [help, pptxGuide]) {
    assert.doesNotMatch(html, /deterministic|AI-mediated|top-level|rectangle-family|waypoint|polls automatically/i);
    assert.match(html, /<details>[\s\S]*?<summary>/);
    assert.doesNotMatch(html, /offense <strong>1&ndash;16|defense <strong>A&ndash;F/);
  }
});

test('help navigation points to real pages and section anchors', () => {
  const pages = new Map([
    ['/', index], ['/help', help], ['/pptx-guide', pptxGuide], ['/editor', editor], ['/converter', converter],
  ]);
  for (const [path, html] of [['/help', help], ['/pptx-guide', pptxGuide]]) {
    for (const [, href] of html.matchAll(/href="([^"]+)"/g)) {
      const link = new URL(href, 'https://example.test' + path);
      if (link.origin !== 'https://example.test') continue;
      assert.ok(pages.has(link.pathname), href);
      if (link.hash) assert.ok(pages.get(link.pathname).includes('id="' + link.hash.slice(1) + '"'), href);
    }
  }
});

test('wristband compatibility guidance is accurate and links safely', () => {
  for (const html of [converter, help]) {
    assert.match(html, /4\.40 &times; 2\.09 in/);
    assert.match(html, /2 5\/8 &times; 4 5\/8 in/);
    assert.doesNotMatch(html, /2\.75 &times; 4\.75 in/);
    assert.match(html, /https:\/\/www\.amazon\.com\/dp\/B07QHCVV7M\?th=1/);
    assert.match(html, /https:\/\/www\.amazon\.com\/Fiskars-SureCut-Portable-Paper-Trimmer\/dp\/B000OMYB18\//);
    assert.match(html, /target="_blank" rel="noopener noreferrer"/);
    assert.match(html, /(?:Print at <strong>100%|<strong>Print at 100%)/i);
    assert.match(html, /laminate before cutting/i);
    assert.match(html, /not affiliated with or endorsed by/);
    assert.match(html, /aria-label="WristCoaches adult three-panel wrist coach on Amazon \(opens in a new tab\)"/);
    assert.match(html, /aria-label="Fiskars SureCut portable paper trimmer on Amazon \(opens in a new tab\)"/);
  }
});

test('both creation surfaces offer independent paginated coach layouts', () => {
  for (const html of [editor, converter]) {
    for (const section of ['offense','defense']) {
      const select = html.match(new RegExp('<select id="' + section + 'PerPage"[^>]*>([\\s\\S]*?)</select>'));
      assert.ok(select);
      assert.deepEqual([...select[1].matchAll(/value="(\d+)"/g)].map(m => Number(m[1])), [1,2,4,6,9,16]);
      assert.match(html, new RegExp(section + '_plays_per_page'));
    }
    assert.match(html, /Unused spaces stay blank/);
  }
  assert.match(editor, /MAX_INCLUDED_OFFENSE = 64, MAX_INCLUDED_DEFENSE = 24/);
  assert.match(editor, /plays\.slice\(start, start \+ perPage\)/);
});

test('PPTX guidance matches the deterministic multi-page converter', () => {
  for (const html of [converter, help]) {
    assert.match(html, /href="\/pptx-guide"/);
  }
  assert.match(converter, /For offense only, no title slides are needed/i);
  assert.match(converter, /DEFENSE title slide but no OFFENSE title slide/);
  assert.match(help, /does not use AI to interpret or change your routes/i);
  assert.match(help, /64 offense/);
  assert.match(help, /24 defense/);
  assert.match(help, /16 plays per page/);
  assert.match(help, /independently of offense/);
  assert.match(help, /up to eight plays in each insert/);

  assert.match(pptxGuide, /Up to 64 offense \/ 24 defense plays/);
  assert.match(pptxGuide, /No OFFENSE title slide\?/);
  assert.match(pptxGuide, /plays before the first DEFENSE title slide are treated as offense/);
  assert.match(pptxGuide, /message after conversion tells you when this happens/);
  assert.match(pptxGuide, /OFFENSE title slide, put it before the first offense play/i);
  assert.match(pptxGuide, /rounded and snipped-corner rectangle variants/i);
  assert.match(pptxGuide, /does not use AI to interpret your routes/);
  assert.match(pptxGuide, /class="converter-cta" href="\/converter">Open PowerPoint \(PPTX\) Import/);
  assert.match(pptxGuide, /href="\/">&larr; GSS Playbook Editor home/);
  assert.doesNotMatch(pptxGuide, /<script\b/i);
  assert.doesNotMatch(pptxGuide, /(?:src|href)="https?:/i);
});
