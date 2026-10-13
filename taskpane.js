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
// Debug mode — on a development host, or with `debug=1` in the task pane URL:
// each tool result then carries the time it took in Office and what it cost.
const DEBUG = params.get('debug') === '1' || ['localhost', '127.0.0.1'].includes(location.hostname);

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
  conversation = { reported: false, id: null };
  ignoredCalls = 0;
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

// The conversation of this pane's chat. A user with Word, Excel and PowerPoint
// open has one pane per application, and every pane receives every tool call
// of that user: each one must only answer the calls of its own conversation,
// or the first to reply — maybe the wrong document — wins.
// `reported`: the chat has told us its conversation at least once (not every
// version of the chat does, whatever it announces); `id`: null until the
// first message of a fresh conversation.
let conversation = { reported: false, id: null };
let ignoredCalls = 0; // calls left to another pane, for the debug figures
const CONVERSATION_CHANGED = 'chat:conversation-changed';
const bareId = (id) => String(id).replace(/^ctx_/, '');

// Whether a relayed tool call is for this pane to answer.
function ownsToolCall({ name, contextId }) {
  // Only a conversation the chat really gave us can rule a call out. The chat
  // gives the id of a new conversation at the end of its first answer, after
  // the tool calls of that answer: until then the pane cannot tell, and must
  // not stay silent.
  if (conversation.id !== null && contextId) {
    return bareId(contextId) === bareId(conversation.id);
  }
  // No way to tell the conversations apart: at least leave a tool of another
  // application to the pane of that application.
  const tool = Object.prototype.hasOwnProperty.call(TOOLS, name) ? TOOLS[name] : null;
  return !(tool && tool.host && host && tool.host !== hostName());
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
  } else if (data.type === 'event' && data.name === CONVERSATION_CHANGED) {
    const state = (data.data && data.data.state) || {};
    conversation = { reported: true, id: typeof state.conversationId === 'string' ? state.conversationId : null };
  } else if (data.type === 'tool-exec') {
    // The embed relayed a tool.exec.request from the agent's run (received on
    // its session event stream). Execute it here and return the result —
    // unless it belongs to the conversation of another pane, which answers it.
    if (ownsToolCall(data)) handleToolExec(data);
    else ignoredCalls++;
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

// ── Ambient document context ────────────────────────────────────────────────
// Same shape as the extension's page context ({ title, url, selection, content })
// so the embed serializes it identically.
// Tells the agent where this content lives: without it, a request such as
// "change the title of slide 2" is taken for a deck of its own to generate.
function contextNote(more) {
  const what = host === Office.HostType.Outlook ? 'email' : 'document';
  return (
    `[This is the ${what} the user has open in ${hostName()}, shared by the Prisme.ai Office add-in. ` +
    `Requests about "this ${what}", "this slide" or "this sheet" refer to it. ` +
    (more
      ? 'Only its beginning is shown here: call office_outline for its structure, office_read for a given part, office_search to find a passage. '
      : '') +
    'To read more of it or to change it, call the Office add-in tools, not a canvas or file-generation tool.]'
  );
}

let contextTimer;
function refreshContext() {
  if (!host) return;
  clearTimeout(contextTimer);
  contextTimer = setTimeout(async () => {
    try {
      // One Office call at a time: the context refresh and the agent's tools
      // would otherwise read the document concurrently.
      const [doc, selection] = await exclusive(async () => [await readDocument(), await selectedText().catch(() => '')]);
      syncContext({
        title: `${doc.title} (${hostName()})`,
        selection: selection.slice(0, 4000) || undefined,
        content: `${contextNote(doc.more || doc.content.length > CONTEXT_CHARS)}\n\n${doc.content.slice(0, CONTEXT_CHARS)}`,
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

// Below the relay's own timeout (30 s), so a stuck Office call is reported as
// such instead of as a missing add-in.
const TOOL_TIMEOUT_MS = 25000;

// Runs Office work one task at a time. A task that never settles is given up
// on after the tool timeout, so it can't block the ones queued behind it.
let officeQueue = Promise.resolve();
function exclusive(task) {
  const previous = officeQueue;
  const run = previous.then(task);
  const giveUp = previous.then(() => new Promise((resolve) => setTimeout(resolve, TOOL_TIMEOUT_MS)));
  officeQueue = Promise.race([run.catch(() => {}), giveUp]);
  return run;
}

// The agent (server) calls a tool on the tools workspace, which relays it to
// us through the embed. Contract (MCP-aligned): correlate by correlationId,
// and answer `result` XOR `error: {code, message}` — exactly once.
function handleToolExec({ correlationId, name, args }) {
  const reply = (payload) => postToFrame(Object.assign({ type: 'tool-result', correlationId }, payload));
  const fail = (message) => reply({ error: { code: 1, message } });
  const deadline = new Promise((_, reject) =>
    setTimeout(
      () => reject(new Error(`${name} did not complete in ${hostName()} within ${TOOL_TIMEOUT_MS / 1000} s`)),
      TOOL_TIMEOUT_MS
    )
  );
  const arrived = Date.now();
  let started = 0;
  const run = () => {
    started = Date.now();
    toolDiag = {};
    officeStep = '';
    return execTool(name, args);
  };
  // `waited`: time spent behind another Office task before this one started.
  const figures = () => ({
    ms: Date.now() - started,
    waited: started - arrived,
    conversation: conversation.reported ? conversation.id || 'none yet' : 'not reported',
    ignored: ignoredCalls || undefined,
    ...toolDiag,
  });
  Promise.race([exclusive(run), deadline]).then(
    (result) => {
      if (DEBUG && result && typeof result === 'object') result.diag = figures();
      const size = JSON.stringify(result === undefined ? null : result).length;
      if (size > MAX_RESULT_CHARS) {
        fail(`${name} returned ${size} characters (max ${MAX_RESULT_CHARS}) — ask for a smaller part`);
      } else {
        reply({ result });
      }
    },
    (err) => {
      const step = officeStep ? ` (while ${officeStep})` : '';
      const diag = DEBUG && started ? ` | diag ${JSON.stringify(figures())}` : '';
      officeStep = '';
      fail(String((err && err.message) || err) + step + diag);
    }
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
