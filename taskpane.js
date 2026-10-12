// Thin shell, same contract as the browser extension: the task pane IS the
// Prisme.ai chat (loaded in an iframe from the instance's /embed/chat). The
// add-in's only job is to keep the agent's context in sync with the document
// or mail you have open, and to execute the agent's office tools through
// Office.js — so you just type any question ("summarize", "rewrite this
// selection", "what does this sheet say?") and the agent already knows the file.
const LOADER_SOURCE = 'prisme-embed'; // loader → iframe
const FRAME_SOURCE = 'prisme-embed-frame'; // iframe → loader

const $ = (id) => document.getElementById(id);
const els = {
  config: $('config'),
  frame: $('frame'),
  gear: $('gear'),
  newchat: $('newchat'),
  quote: $('quote'),
  instance: $('cfg-instance'),
  api: $('cfg-api'),
  agent: $('cfg-agent'),
  org: $('cfg-org'),
  officeTools: $('cfg-office-tools'),
  auth: $('cfg-auth'),
  sso: $('sso-fields'),
  ssoProvider: $('cfg-sso-provider'),
  ssoClient: $('cfg-sso-client'),
  ssoTenant: $('cfg-sso-tenant'),
  error: $('error'),
  connect: $('connect'),
};

// Agents all live in the shared agent-factory workspace, so the slug is fixed
// and not exposed in the UI.
const WORKSPACE_SLUG = 'slug:agent-factory';
const STORAGE_KEY = 'prisme-embed-cfg';

// Icons matching the Prisme.ai platform codebase (lucide: PenSquare / Settings / Quote).
const svg = (paths) =>
  '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
  paths +
  '</svg>';
els.newchat.innerHTML = svg(
  '<path d="M12 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.12 2.12 0 0 1 3 3L12 15l-4 1 1-4Z"/>'
);
els.gear.innerHTML = svg(
  '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>'
);
els.quote.innerHTML = svg(
  '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M8 12a2 2 0 0 0 2-2V8H8"/><path d="M14 12a2 2 0 0 0 2-2V8h-2"/>'
);

let cfg = null;
let frameOrigin = null;
let frameReady = false;
let lastContext = null;
let host = null; // Office.HostType.* — null when opened outside Office (plain chat)

const DEFAULTS = {
  instance: '',
  api: '',
  agent: '',
  org: '',
  // Workspace exposing the office_* tool-calls. Left blank → the add-in is a
  // plain document-context chat (no tool execution).
  officeTools: '',
  // 'anonymous', or 'sso': silent Microsoft Entra sign-in traded for a
  // nominative Prisme.ai session (see the single sign-on section below).
  auth: 'anonymous',
  ssoProvider: '',
  ssoClient: '',
  ssoTenant: '',
};

// Config precedence: managed (query string of the task pane URL, written into
// the manifest by the admin — the Office equivalent of chrome.storage.managed)
// wins over the user's saved config, which wins over defaults. Never put a
// secret in these parameters: the manifest is readable by every assigned user.
const params = new URLSearchParams(location.search);
const MANAGED = {
  instance: params.get('instance'),
  api: params.get('api'),
  agent: params.get('agent'),
  org: params.get('org'),
  officeTools: params.get('tools'),
  auth: params.get('auth'),
  ssoProvider: params.get('sso_provider'),
  ssoClient: params.get('sso_client'),
  ssoTenant: params.get('sso_tenant'),
};

// localStorage can be unavailable or partitioned in a webview — never rely on it.
function loadSaved() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || null;
  } catch {
    return null;
  }
}
function save(value) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(value));
  } catch {
    /* config stays in memory for this session */
  }
}

function boot() {
  const saved = loadSaved();
  const c = { ...DEFAULTS, ...(saved || {}) };
  Object.keys(MANAGED).forEach((key) => {
    if (!MANAGED[key]) return;
    c[key] = MANAGED[key];
    els[key].disabled = true; // locked by the deployment
  });
  els.instance.value = c.instance;
  els.api.value = c.api;
  els.agent.value = c.agent;
  els.org.value = c.org;
  els.officeTools.value = c.officeTools;
  els.auth.value = c.auth === 'sso' ? 'sso' : 'anonymous';
  els.ssoProvider.value = c.ssoProvider;
  els.ssoClient.value = c.ssoClient;
  els.ssoTenant.value = c.ssoTenant;
  toggleSsoFields();
  // Auto-connect when the manifest supplies an agent, or the user saved one.
  if (c.agent && (MANAGED.agent || (saved && saved.agent))) connect();
}

function readForm() {
  return {
    instance: els.instance.value.trim().replace(/\/+$/, ''),
    api: els.api.value.trim(),
    agent: els.agent.value.trim(),
    org: els.org.value.trim(),
    officeTools: els.officeTools.value.trim(),
    auth: els.auth.value,
    ssoProvider: els.ssoProvider.value.trim(),
    ssoClient: els.ssoClient.value.trim(),
    ssoTenant: els.ssoTenant.value.trim(),
  };
}

function toggleSsoFields() {
  els.sso.classList.toggle('hidden', els.auth.value !== 'sso');
}

// window.alert is blocked in Office webviews → inline error.
function showError(message) {
  els.error.textContent = message || '';
}

function connect() {
  cfg = readForm();
  if (!cfg.instance || !cfg.agent) {
    showError('Platform URL and Agent ID are required.');
    return;
  }
  try {
    frameOrigin = new URL(cfg.instance).origin;
  } catch {
    showError('Invalid Platform URL.');
    return;
  }
  if (cfg.auth === 'sso' && !(cfg.api && cfg.ssoProvider && cfg.ssoClient && cfg.ssoTenant)) {
    showError('Single sign-on needs the API URL, the SSO provider, the Entra client ID and the tenant ID.');
    return;
  }
  showError('');
  save(cfg);
  loadChat();
  els.config.classList.add('hidden');
  els.frame.classList.remove('hidden');
  els.gear.classList.remove('hidden');
  els.newchat.classList.remove('hidden');
  if (host) els.quote.classList.remove('hidden');
}

function showSettings() {
  els.config.classList.remove('hidden');
  els.frame.classList.add('hidden');
  els.gear.classList.add('hidden');
  els.newchat.classList.add('hidden');
  els.quote.classList.add('hidden');
}

// (Re)load the embed → a fresh conversation.
function loadChat() {
  frameReady = false;
  els.frame.src = cfg.instance + '/embed/chat?t=' + Date.now();
}

function postToFrame(message) {
  if (els.frame.contentWindow && frameOrigin) {
    els.frame.contentWindow.postMessage({ source: LOADER_SOURCE, ...message }, frameOrigin);
  }
}

function sendConfig() {
  postToFrame({
    type: 'config',
    config: {
      workspaceSlug: WORKSPACE_SLUG,
      agentId: cfg.agent,
      apiUrl: cfg.api,
      // SSO: the embed asks us for the bearer ('token-request') instead of
      // minting an anonymous session.
      auth: cfg.auth === 'sso' ? 'token' : 'anonymous',
      orgSlug: cfg.org || undefined,
      theme: 'light',
      features: {
        enableFiles: true,
        enableFeedback: true,
        enableToolCalls: true,
        enableArtifacts: true,
        // Voice-to-voice is not ready yet → keep only speech-to-text (dictation).
        enableVoice: false,
        enableSpeechToText: true,
      },
      context: lastContext || undefined,
      // When set, the iframe opens the workspace's tool.exec.* stream and
      // relays execution back here (handleToolExec).
      toolsWorkspace: cfg.officeTools || undefined,
    },
  });
}

// Push the current document context to the chat. It rides along with every
// message the user sends (parts[].metadata.contextSnippet), so no special UI
// is needed.
function syncContext(ctx) {
  lastContext = ctx || lastContext;
  if (frameReady && lastContext) postToFrame({ type: 'setContext', context: lastContext });
}

window.addEventListener('message', (event) => {
  if (event.origin !== frameOrigin) return;
  const data = event.data;
  if (!data || data.source !== FRAME_SOURCE) return;
  if (data.type === 'ready') {
    frameReady = true;
    sendConfig();
    refreshContext(); // prime immediately
  } else if (data.type === 'token-request') {
    answerTokenRequest();
  } else if (data.type === 'tool-exec') {
    // The embed relayed a tool.exec.request from the agent's run (received on
    // its session event stream). Execute it here and return the result.
    handleToolExec(data);
  }
});

// ── Single sign-on (auth: 'sso') ─────────────────────────────────────────────
// No popup, no redirect: the user is already signed in to Microsoft in Office,
// so MSAL's Nested App Authentication hands us an Entra id_token through the
// Office host, and the instance trades it for a nominative Prisme.ai session
// (POST /login/token-exchange). The Entra app must be the one the instance's
// SSO provider already uses: Entra's `sub` differs per app registration, and
// that is what the instance matches users on.
let msalApp = null;
let session = null; // { token, expires (epoch ms), servedAt (epoch ms) }

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const el = document.createElement('script');
    el.src = src;
    el.onload = resolve;
    el.onerror = () => reject(new Error('Could not load ' + src));
    document.head.appendChild(el);
  });
}

async function entraIdToken() {
  if (!host) throw new Error('Single sign-on only works inside an Office application.');
  if (!Office.context.requirements.isSetSupported('NestedAppAuth', '1.1')) {
    throw new Error('This version of Office does not support single sign-on. Update Office and try again.');
  }
  if (!msalApp) {
    // Vendored (Microsoft no longer serves MSAL from a CDN) and only loaded
    // when SSO is on.
    if (!window.msal) await loadScript('vendor/msal-browser.min.js');
    msalApp = await window.msal.createNestablePublicClientApplication({
      auth: {
        clientId: cfg.ssoClient,
        authority: 'https://login.microsoftonline.com/' + cfg.ssoTenant,
      },
    });
  }
  const request = { scopes: ['User.Read', 'openid', 'profile'] };
  let result;
  try {
    result = await msalApp.acquireTokenSilent(request);
  } catch {
    // First use, or consent needed: the Office host shows Microsoft's prompt.
    result = await msalApp.acquireTokenPopup(request);
  }
  if (!result || !result.idToken) throw new Error('Microsoft did not return an ID token.');
  return result.idToken;
}

async function openSession() {
  const idToken = await entraIdToken();
  let res;
  try {
    res = await fetch(cfg.api.replace(/\/+$/, '') + '/login/token-exchange', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'omit', // the bearer is all we need; no cross-site cookie
      body: JSON.stringify({ provider: cfg.ssoProvider, subject_token: idToken }),
    });
  } catch {
    throw new Error(
      'Could not reach the Prisme.ai API. Check the API URL and that this add-in host is allowed by the API (CORS).'
    );
  }
  const body = await res.json().catch(() => ({}));
  if (!res.ok || !body.token) {
    const detail = body.message || body.error || '';
    throw new Error(`The Prisme.ai instance refused the sign-in (${res.status})${detail ? ': ' + detail : ''}`);
  }
  return { token: body.token, expires: Date.parse(body.expires) || Date.now() + 3600000, servedAt: 0 };
}

// The embed asks for the bearer at mount and again whenever the API answers
// 401. Reuse the session while it is valid — except when it was handed over a
// moment ago: a request that soon means the instance rejected it.
async function answerTokenRequest() {
  try {
    const now = Date.now();
    const reusable = session && session.expires - now > 60000 && now - session.servedAt > 15000;
    if (!reusable) session = await openSession();
    session.servedAt = Date.now();
    postToFrame({ type: 'token', token: session.token });
  } catch (err) {
    console.error('[prisme-office] sign-in:', err);
    session = null;
    postToFrame({ type: 'token', token: null });
    showSettings();
    showError(String((err && err.message) || err));
  }
}

// ── Office.js helpers ────────────────────────────────────────────────────────
const HOST_LABELS = { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint', Outlook: 'Outlook' };
function hostName() {
  const T = Office.HostType;
  if (host === T.Word) return HOST_LABELS.Word;
  if (host === T.Excel) return HOST_LABELS.Excel;
  if (host === T.PowerPoint) return HOST_LABELS.PowerPoint;
  if (host === T.Outlook) return HOST_LABELS.Outlook;
  return null;
}

// Callback-style Office API → promise (rejects on a failed AsyncResult).
function officeAsync(call) {
  return new Promise((resolve, reject) => {
    call((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(new Error((res.error && res.error.message) || 'Office call failed'));
    });
  });
}

// Mail fields are plain values in read mode and async getters in compose mode.
function mailField(field) {
  if (field && typeof field.getAsync === 'function') {
    return officeAsync((cb) => field.getAsync(cb)).catch(() => undefined);
  }
  return Promise.resolve(field);
}

function formatAddress(a) {
  if (!a) return '';
  return a.displayName && a.emailAddress ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress || a.displayName || '';
}

function rowsToText(values) {
  return values.map((row) => row.join('\t')).join('\n');
}

function documentTitle() {
  const url = (Office.context.document && Office.context.document.url) || '';
  // Local path or web URL → keep the file name only.
  const name = decodeURIComponent(url.split(/[\\/]/).pop().split('?')[0] || '');
  return name || hostName() + ' document';
}

async function readWord() {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load('text');
    await ctx.sync();
    return body.text;
  });
}

// Active sheet's used range, capped so a huge workbook doesn't flood the context.
const MAX_ROWS = 200;
const MAX_COLS = 50;
async function readExcel() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
    await ctx.sync();
    if (used.isNullObject) return { sheet: sheet.name, address: '', values: [] };
    const range = sheet.getRangeByIndexes(
      used.rowIndex,
      used.columnIndex,
      Math.min(used.rowCount, MAX_ROWS),
      Math.min(used.columnCount, MAX_COLS)
    );
    range.load(['address', 'values']);
    await ctx.sync();
    return { sheet: sheet.name, address: range.address, values: range.values };
  });
}

// Shape types that never carry a text frame — loading their text would fail the batch.
const NO_TEXT_SHAPES = ['Image', 'Group', 'Line', 'Table', 'Chart', 'Unsupported'];
const MAX_SLIDES = 50;
async function readPowerPoint() {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items');
    await ctx.sync();
    const items = slides.items.slice(0, MAX_SLIDES);
    items.forEach((slide) => slide.shapes.load('items/type'));
    await ctx.sync();
    const out = [];
    for (let i = 0; i < items.length; i++) {
      const shapes = items[i].shapes.items.filter((s) => !NO_TEXT_SHAPES.includes(s.type));
      const texts = [];
      // One shape per sync: a shape without a text frame rejects the whole
      // batch, and its type isn't always enough to tell.
      for (const shape of shapes) {
        try {
          const range = shape.textFrame.textRange;
          range.load('text');
          await ctx.sync();
          if (range.text && range.text.trim()) texts.push(range.text.trim());
        } catch {
          /* no text in this shape */
        }
      }
      out.push({ slide: i + 1, text: texts.join('\n') });
    }
    return out;
  });
}

async function readMail() {
  const item = Office.context.mailbox.item;
  if (!item) return null;
  const [subject, from, to, body] = await Promise.all([
    mailField(item.subject),
    mailField(item.from),
    mailField(item.to),
    officeAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb)).catch(() => ''),
  ]);
  return {
    subject: subject || '',
    from: formatAddress(from),
    to: (to || []).map(formatAddress).filter(Boolean),
    body: body || '',
  };
}

async function selectedText() {
  const T = Office.HostType;
  if (host === T.Excel) {
    // Text coercion only handles a single cell → read the selected range.
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(['values', 'cellCount']);
      await ctx.sync();
      return range.cellCount > 1 || range.values[0][0] !== '' ? rowsToText(range.values) : '';
    });
  }
  if (host === T.Outlook) {
    const item = Office.context.mailbox.item;
    // Only compose mode exposes the selection.
    if (!item || typeof item.getSelectedDataAsync !== 'function') return '';
    const res = await officeAsync((cb) => item.getSelectedDataAsync(Office.CoercionType.Text, cb));
    return (res && res.data) || '';
  }
  const text = await officeAsync((cb) =>
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, cb)
  );
  return String(text || '');
}

// Full readable content of the open document / mail, as { title, content }.
async function readDocument() {
  const T = Office.HostType;
  if (host === T.Word) return { title: documentTitle(), content: await readWord() };
  if (host === T.Excel) {
    const { sheet, address, values } = await readExcel();
    return { title: documentTitle(), content: `Sheet: ${sheet}\nRange: ${address}\n\n${rowsToText(values)}` };
  }
  if (host === T.PowerPoint) {
    const slides = await readPowerPoint();
    return {
      title: documentTitle(),
      content: slides.map((s) => `--- Slide ${s.slide} ---\n${s.text}`).join('\n\n'),
    };
  }
  if (host === T.Outlook) {
    const mail = await readMail();
    if (!mail) return { title: 'Outlook', content: '' };
    return {
      title: mail.subject || 'Outlook message',
      content: `From: ${mail.from}\nTo: ${mail.to.join(', ')}\n\n${mail.body}`,
    };
  }
  throw new Error('not running inside an Office application');
}

// ── Ambient document context ────────────────────────────────────────────────
// Same shape as the extension's page context ({ title, url, selection, content })
// so the embed serializes it identically.
let contextTimer;
function refreshContext() {
  if (!host) return;
  clearTimeout(contextTimer);
  contextTimer = setTimeout(async () => {
    try {
      const [doc, selection] = await Promise.all([readDocument(), selectedText().catch(() => '')]);
      syncContext({
        title: `${doc.title} (${hostName()})`,
        selection: selection.slice(0, 4000) || undefined,
        content: doc.content.slice(0, 8000),
      });
    } catch (err) {
      console.warn('[prisme-office] context:', err);
    }
  }, 800);
}

function watchDocument() {
  if (host === Office.HostType.Outlook) {
    // Fires when the pane is pinned and the user opens another message.
    if (typeof Office.context.mailbox.addHandlerAsync === 'function') {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, refreshContext);
    }
  } else {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refreshContext);
  }
  window.addEventListener('focus', refreshContext);
}

// ── Action executor (tool-calls) ────────────────────────────────────────────
// The agent (server) calls office_* tools on the tools workspace; those relay
// to us through the embed. Contract (MCP-aligned): correlate by correlationId,
// and return `result` XOR `error: {code, message}` — no `ok` flag.
function requireHost(type, tool) {
  if (host !== type) throw new Error(`${tool} is not available in ${hostName() || 'this context'}`);
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}

async function insertText({ text, location }) {
  if (typeof text !== 'string') throw new Error('text is required');
  const T = Office.HostType;
  if (host === T.Word) {
    await Word.run(async (ctx) => {
      if (location === 'start' || location === 'end') {
        ctx.document.body.insertText(text, location === 'start' ? 'Start' : 'End');
      } else {
        ctx.document.getSelection().insertText(text, 'Replace');
      }
      await ctx.sync();
    });
    return { ok: true };
  }
  if (host === T.Outlook) {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.body.setSelectedDataAsync !== 'function') {
      throw new Error('the message is open in read mode — use outlook_draft_reply');
    }
    await officeAsync((cb) =>
      item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, cb)
    );
    return { ok: true };
  }
  // Excel (active cell) and PowerPoint (selected text box).
  await officeAsync((cb) =>
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, cb)
  );
  return { ok: true };
}

const MAX_CELLS = 5000;
function targetSheet(ctx, name) {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

async function excelReadRange({ address, sheet }) {
  return Excel.run(async (ctx) => {
    const range = address ? targetSheet(ctx, sheet).getRange(address) : ctx.workbook.getSelectedRange();
    range.load(['address', 'cellCount']);
    await ctx.sync();
    if (range.cellCount > MAX_CELLS) {
      throw new Error(`range ${range.address} has ${range.cellCount} cells (max ${MAX_CELLS}) — narrow it`);
    }
    range.load(['values', 'formulas']);
    await ctx.sync();
    return { address: range.address, values: range.values, formulas: range.formulas };
  });
}

async function excelWriteRange({ address, values, sheet }) {
  if (!address) throw new Error('address is required');
  const rows = Array.isArray(values) ? values : null;
  if (!rows || !rows.length || !rows.every((r) => Array.isArray(r) && r.length === rows[0].length)) {
    throw new Error('values must be a non-empty rectangular 2D array');
  }
  return Excel.run(async (ctx) => {
    // `address` is the top-left anchor; the range grows to fit `values`.
    const range = targetSheet(ctx, sheet)
      .getRange(address)
      .getCell(0, 0)
      .getResizedRange(rows.length - 1, rows[0].length - 1);
    range.values = rows; // a string starting with "=" is written as a formula
    range.load('address');
    await ctx.sync();
    return { ok: true, address: range.address };
  });
}

async function outlookDraftReply({ body, replyAll }) {
  if (typeof body !== 'string') throw new Error('body is required');
  const item = Office.context.mailbox.item;
  if (!item || typeof item.displayReplyForm !== 'function') {
    throw new Error('no message open in read mode');
  }
  // Opens Outlook's own reply form, unsent — the user reviews and sends.
  if (replyAll) item.displayReplyAllForm(escapeHtml(body));
  else item.displayReplyForm(escapeHtml(body));
  return { ok: true, sent: false };
}

async function execTool(tool, args) {
  args = args || {};
  if (!host) throw new Error('not running inside an Office application');
  const T = Office.HostType;
  switch (tool) {
    case 'office_read':
      return { host: hostName(), ...(await readDocument()) };
    case 'office_get_selection':
      return { host: hostName(), text: await selectedText() };
    case 'office_insert_text':
      return insertText(args);
    case 'excel_read_range':
      requireHost(T.Excel, tool);
      return excelReadRange(args);
    case 'excel_write_range':
      requireHost(T.Excel, tool);
      return excelWriteRange(args);
    case 'powerpoint_read_slides':
      requireHost(T.PowerPoint, tool);
      return { slides: await readPowerPoint() };
    case 'outlook_read_item':
      requireHost(T.Outlook, tool);
      return readMail();
    case 'outlook_draft_reply':
      requireHost(T.Outlook, tool);
      return outlookDraftReply(args);
    default:
      throw new Error('unknown tool ' + tool);
  }
}

function handleToolExec({ correlationId, name, args }) {
  const reply = (payload) => postToFrame(Object.assign({ type: 'tool-result', correlationId }, payload));
  execTool(name, args).then(
    (result) => reply({ result }),
    (err) => reply({ error: { code: 1, message: String((err && err.message) || err) } })
  );
}

// ── UI ──────────────────────────────────────────────────────────────────────
els.connect.addEventListener('click', connect);
els.newchat.addEventListener('click', () => {
  loadChat(); // reload the iframe → new conversation (config re-posted on ready)
});
els.gear.addEventListener('click', showSettings);
els.auth.addEventListener('change', toggleSsoFields);
// Selection → quoted into the composer (the task pane's "Ask {agent}" pill).
els.quote.addEventListener('click', async () => {
  const text = (await selectedText().catch(() => '')).trim();
  if (frameReady) postToFrame({ type: 'setQuote', text: text || null });
});

// Office.onReady resolves with a null host when the page is opened in a plain
// browser tab → the pane still works as a context-less chat (handy for dev).
if (window.Office) {
  Office.onReady((info) => {
    host = (info && info.host) || null;
    if (host && !hostName()) host = null; // unsupported host (OneNote, Project…)
    boot();
    if (host) watchDocument();
  });
} else {
  boot();
}
