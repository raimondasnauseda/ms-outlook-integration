/* Offline tests: link validation (ported from the Google add-on), meeting block,
 * function-command flow with a mocked Office.js, and source hygiene. Run: npm test */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const L = require('../docs/link.js');

let passed = 0, failed = 0;
const pending = [];
function test(name, fn) {
  pending.push(Promise.resolve().then(fn).then(
    () => { passed++; console.log('  ok   ' + name); },
    (e) => { failed++; console.log('  FAIL ' + name + '\n       ' + (e && e.message)); }
  ));
}
function eq(a, b, msg) { if (a !== b) throw new Error((msg || '') + ` expected ${JSON.stringify(b)}, got ${JSON.stringify(a)}`); }
function ok(c, msg) { if (!c) throw new Error(msg || 'assertion failed'); }

// ---------- Validation (same cases as the Google repo) ----------
const accepted = [
  ['https://plus.interactio.com/link/x', 'https://plus.interactio.com/link/x'],
  ['https://meeting.interactio.com/link/x', 'https://meeting.interactio.com/link/x'],
  ['meeting.interactio.com/link/x', 'https://meeting.interactio.com/link/x'],
  ['  https://meeting.interactio.com/link/x  ', 'https://meeting.interactio.com/link/x'],
  ['https://INTERACTIO.io', 'https://interactio.io'],
  ['https://Meeting.Interactio.COM/link/AbC', 'https://meeting.interactio.com/link/AbC'],
  ['https://interactio.com', 'https://interactio.com'],
  ['https://interactio.com./link/x', 'https://interactio.com/link/x'],
  ['https://a.b.interactio.io:8443/p#f', 'https://a.b.interactio.io:8443/p#f'],
  ['https://interactio.io/x?y=1', 'https://interactio.io/x?y=1'],
];
for (const [inp, out] of accepted) test(`accepts ${inp.trim()}`, () => {
  const r = L.normalizeInteractioLink(inp); ok(r.ok, r.error); eq(r.link, out);
});
const rejected = [
  '', null, undefined, '   ',
  'http://meeting.interactio.com/link/x',
  'https://evil.com/interactio.io',
  'https://interactio.io.evil.com/x',
  'https://evilinteractio.io/x',
  'https://evil.com@interactio.io/x',
  'https://user:pw@interactio.io/x',
  'javascript:alert(1)',
  'ftp://interactio.io',
  'https://meeting.interactio.com/link/a b',
  'https://interactio.com/' + 'a'.repeat(2050),
  'https://zoom.us/j/123',
  'https://interactio.com:99999999/x',
];
for (const inp of rejected) test(`rejects ${JSON.stringify(inp && inp.length > 60 ? inp.slice(0, 40) + '…' : inp)}`, () => {
  const r = L.normalizeInteractioLink(inp); ok(!r.ok, 'should be rejected'); ok(r.error, 'error message');
});

// ---------- Meeting block ----------
const LINK = 'https://meeting.interactio.com/link/raimondas';
test('block contains link, Join text and id', () => {
  const b = L.buildMeetingBlock(LINK);
  ok(b.includes(`href="${LINK}"`)); ok(b.includes('Join Interactio meeting')); ok(b.includes('id="interactio-meeting"'));
});
test('link appears exactly once as visible text (plain-text invites)', () => {
  const b = L.buildMeetingBlock(LINK);
  const text = b.replace(/<[^>]+>/g, ' ');
  eq(text.split(LINK).length - 1, 1, 'visible URL count');
  eq(b.split(`href="${LINK}"`).length - 1, 1, 'one href');
  ok(!/_{41,}/.test(b), 'divider lines are short');
});
test('block HTML-escapes the link', () => {
  const r = L.normalizeInteractioLink('https://interactio.com/x"><script>alert(1)</script>');
  ok(r.ok, 'validator accepts it (no spaces)');
  const b = L.buildMeetingBlock(r.link);
  ok(!b.includes('<script>'), 'no raw script tag'); ok(b.includes('x&quot;&gt;&lt;script&gt;'), 'escaped');
});
test('merge appends to an existing body', () => {
  const out = L.mergeMeetingBlock('<p>Agenda</p>', LINK);
  ok(out.startsWith('<p>Agenda</p>')); ok(out.includes(LINK));
});
test('merge is a no-op when the same link is already there', () => {
  const once = L.mergeMeetingBlock('<p>Agenda</p>', LINK);
  eq(L.mergeMeetingBlock(once, LINK), null);
});
test('merge replaces an older Interactio block instead of adding a second', () => {
  const old = L.mergeMeetingBlock('<p>Agenda</p>', 'https://meeting.interactio.com/link/old') + '<p>Footer</p>';
  const out = L.mergeMeetingBlock(old, LINK);
  ok(!out.includes('/link/old'), 'old link removed'); ok(out.includes(LINK));
  eq(out.split('Join Interactio meeting').length - 1, 1, 'one block');
  ok(out.startsWith('<p>Agenda</p>') && out.endsWith('<p>Footer</p>'), 'surrounding content kept');
});
test('merge replaces a block that Outlook rewrote (x_ id, spans)', () => {
  const rewritten = '<html><body><p>Hi</p><div id="x_interactio-meeting"><br><div><span>' + '_'.repeat(80) +
    '</span></div><div><b>Interactio meeting</b></div><div><a href="https://interactio.io/old">Join Interactio meeting</a></div>' +
    '<div>https://interactio.io/old</div><div><span>' + '_'.repeat(80) + '</span></div></div></body></html>';
  const out = L.mergeMeetingBlock(rewritten, LINK);
  ok(!out.includes('interactio.io/old')); ok(out.includes(LINK)); ok(out.endsWith('</body></html>'), 'body/html kept');
});
test('merge handles empty/null body', () => {
  ok(L.mergeMeetingBlock('', LINK).includes(LINK)); ok(L.mergeMeetingBlock(null, LINK).includes(LINK));
});

// ---------- Function command with mocked Office.js ----------
function loadCommands({ saved, body = '<p>Agenda</p>', dialogReply, failGet, failSet }) {
  const store = new Map(saved ? [[L.SETTINGS_KEY, saved]] : []);
  const state = { body, completed: null, dialogOpened: false, notifications: [], savedAsync: 0 };
  const Office = {
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    CoercionType: { Html: 'html' },
    EventType: { DialogMessageReceived: 'dialogMessageReceived', DialogEventReceived: 'dialogEventReceived' },
    MailboxEnums: { ItemNotificationMessageType: { ErrorMessage: 'errorMessage', InformationalMessage: 'informationalMessage' } },
    onReady: () => Promise.resolve(),
    actions: { associate: (n, f) => { state.fn = f; } },
    context: {
      roamingSettings: {
        get: (k) => store.get(k), set: (k, v) => store.set(k, v), remove: (k) => store.delete(k),
        saveAsync: (cb) => { state.savedAsync++; cb({ status: 'succeeded' }); },
      },
      mailbox: { item: {
        body: {
          getAsync: (t, cb) => cb(failGet ? { status: 'failed', error: { message: 'x' } } : { status: 'succeeded', value: state.body }),
          setAsync: (v, o, cb) => { if (failSet) return cb({ status: 'failed', error: { message: 'x' } }); state.body = v; cb({ status: 'succeeded' }); },
        },
        notificationMessages: { replaceAsync: (k, m, cb) => { state.notifications.push(m); cb && cb({ status: 'succeeded' }); } },
      } },
      ui: {
        displayDialogAsync: (url, opts, cb) => {
          state.dialogOpened = url;
          const handlers = {};
          const dialog = { addEventHandler: (t, h) => { handlers[t] = h; }, close: () => { state.dialogClosed = true; } };
          cb({ status: 'succeeded', value: dialog });
          setTimeout(() => {
            if (dialogReply === undefined) handlers.dialogEventReceived && handlers.dialogEventReceived({ error: 12006 });
            else handlers.dialogMessageReceived && handlers.dialogMessageReceived({ message: dialogReply });
          }, 0);
        },
      },
    },
  };
  const ctx = { Office, console, setTimeout, window: { location: { href: 'https://example.github.io/addin/commands.html' } } };
  ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../docs/link.js'), 'utf8'), ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../docs/commands.js'), 'utf8'), ctx);
  const run = () => new Promise((resolve) => {
    state.fn({ completed: (o) => { state.completed = o; resolve(state); } });
  });
  return { state, store, run };
}

test('fn: saved link → inserted, allowEvent true', async () => {
  const { run } = loadCommands({ saved: LINK });
  const s = await run();
  eq(s.completed.allowEvent, true); ok(s.body.includes(LINK)); ok(!s.dialogOpened, 'no dialog');
});
test('fn: no saved link → settings dialog, saves reply, inserts', async () => {
  const { run, store } = loadCommands({ saved: null, dialogReply: JSON.stringify({ link: 'meeting.interactio.com/link/new' }) });
  const s = await run();
  ok(String(s.dialogOpened).includes('settings.html'), 'dialog opened');
  eq(s.completed.allowEvent, true); eq(store.get(L.SETTINGS_KEY), 'https://meeting.interactio.com/link/new');
  ok(s.body.includes('https://meeting.interactio.com/link/new'));
});
test('fn: dialog returns non-Interactio link → rejected, not saved', async () => {
  const { run, store } = loadCommands({ saved: null, dialogReply: JSON.stringify({ link: 'https://evil.com/x' }) });
  const s = await run();
  eq(s.completed.allowEvent, false); eq(store.get(L.SETTINGS_KEY), undefined); ok(!s.body.includes('evil'));
});
test('fn: dialog closed → allowEvent false, body untouched', async () => {
  const { run } = loadCommands({ saved: null });
  const s = await run();
  eq(s.completed.allowEvent, false); eq(s.body, '<p>Agenda</p>');
});
test('fn: tampered saved value is re-validated', async () => {
  const { run } = loadCommands({ saved: 'https://evil.com/x', dialogReply: undefined });
  const s = await run();
  ok(s.dialogOpened, 'falls back to setup'); ok(!s.body.includes('evil'));
});
test('fn: already present → no duplicate, allowEvent true', async () => {
  const { run } = loadCommands({ saved: LINK, body: L.mergeMeetingBlock('<p>x</p>', LINK) });
  const s = await run();
  eq(s.completed.allowEvent, true); eq(s.body.split(LINK).length - 1, 2, 'href + text, once');
});
test('fn: body read failure → allowEvent false', async () => {
  const s = await loadCommands({ saved: LINK, failGet: true }).run(); eq(s.completed.allowEvent, false);
});
test('fn: body write failure → allowEvent false', async () => {
  const s = await loadCommands({ saved: LINK, failSet: true }).run(); eq(s.completed.allowEvent, false);
});
test('per-user isolation: user B never sees user A link', async () => {
  const a = loadCommands({ saved: LINK }); await a.run();
  const b = loadCommands({ saved: null }); const s = await b.run();
  ok(!s.body.includes(LINK), 'B body has no A link'); ok(s.dialogOpened, 'B gets setup');
});

// ---------- Source hygiene ----------
test('no hard-coded personal link in sources', () => {
  const files = ['docs/link.js', 'docs/commands.js', 'docs/settings.js', 'docs/settings.html', 'docs/commands.html', 'manifest.template.xml'];
  for (const f of files) {
    const src = fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
    const links = src.match(/interactio\.(?:com|io)\/link\/[a-z0-9_-]+/gi) || [];
    for (const l of links) ok(/\/link\/(your-name|x)$/i.test(l), `${f}: personal link ${l}`);
  }
});

Promise.all(pending).then(() => {
  console.log(`\n${passed} passed, ${failed} failed`);
  process.exit(failed ? 1 : 0);
});
