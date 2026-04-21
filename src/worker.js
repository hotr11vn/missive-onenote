/**
 * Missive → OneNote Integration
 * Cloudflare Worker — Standalone (no build step required)
 *
 * Features:
 *  - Missive sidebar app (iFrame) using Missive JS SDK
 *  - Microsoft OAuth 2.0 via popup (works inside iFrames)
 *  - Fetches email content client-side via Missive JS API
 *  - Creates OneNote pages via Microsoft Graph API
 *  - Preserves full email HTML formatting
 *  - Notebook + Section picker with lazy loading
 *  - Remembers last saved location per user
 *
 * Environment secrets (set via: wrangler secret put <NAME>):
 *  - MICROSOFT_CLIENT_ID
 *  - MICROSOFT_CLIENT_SECRET
 *
 * KV binding: ONENOTE_KV (set in wrangler.toml)
 */

// ─── Constants ────────────────────────────────────────────────────────────────

const MS_AUTH_URL   = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
const MS_TOKEN_URL  = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const GRAPH_BASE    = 'https://graph.microsoft.com/v1.0';
const MS_SCOPES     = 'Notes.ReadWrite offline_access User.Read';

// ─── Helpers: Session & Randomness ───────────────────────────────────────────

function randomId(len = 32) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let out = '';
  for (let i = 0; i < len; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

// ─── Helpers: Token Storage (Cloudflare KV) ───────────────────────────────────

async function storeTokens(env, sessionId, tokens) {
  await env.ONENOTE_KV.put(
    `session:${sessionId}`,
    JSON.stringify(tokens),
    { expirationTtl: 60 * 60 * 24 * 30 } // 30 days
  );
}

async function getTokens(env, sessionId) {
  if (!sessionId) return null;
  return env.ONENOTE_KV.get(`session:${sessionId}`, 'json');
}

async function refreshAccessToken(env, sessionId, refreshToken, workerUrl) {
  const res = await fetch(MS_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id:     env.MICROSOFT_CLIENT_ID,
      client_secret: env.MICROSOFT_CLIENT_SECRET,
      refresh_token: refreshToken,
      grant_type:    'refresh_token',
      scope:         MS_SCOPES,
      redirect_uri:  `${workerUrl}/auth/callback`,
    }),
  });
  if (!res.ok) return null;
  const tokens = await res.json();
  tokens.expires_at = Date.now() + tokens.expires_in * 1000;
  await storeTokens(env, sessionId, tokens);
  return tokens;
}

async function getValidToken(env, sessionId, workerUrl) {
  const tokens = await getTokens(env, sessionId);
  if (!tokens) return null;
  // Refresh 5 min before expiry
  if (tokens.expires_at && Date.now() > tokens.expires_at - 300_000) {
    const refreshed = await refreshAccessToken(env, sessionId, tokens.refresh_token, workerUrl);
    return refreshed?.access_token ?? null;
  }
  return tokens.access_token;
}

// ─── Helpers: Microsoft Graph ─────────────────────────────────────────────────

async function graph(accessToken, path, opts = {}) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...opts,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...opts.headers,
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

// ─── Helpers: Last Location (per MS user) ────────────────────────────────────

async function getLastLocation(env, userId) {
  return env.ONENOTE_KV.get(`last_location:${userId}`, 'json');
}

async function setLastLocation(env, userId, loc) {
  await env.ONENOTE_KV.put(
    `last_location:${userId}`,
    JSON.stringify(loc),
    { expirationTtl: 60 * 60 * 24 * 365 }
  );
}

// ─── Session ID from request header ──────────────────────────────────────────

function getSessionId(request) {
  return request.headers.get('X-Session-ID') || null;
}

// ─── HTML: Sidebar Page ───────────────────────────────────────────────────────

function renderSidebar({ workerUrl, lastLocation = null, error = '' } = {}) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Save to OneNote</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      font-size: 13px;
      color: #1a1a1a;
      background: #fff;
      padding: 14px 16px;
    }

    /* ── Header ── */
    .header {
      display: flex;
      align-items: center;
      gap: 9px;
      margin-bottom: 16px;
      padding-bottom: 12px;
      border-bottom: 1px solid #ebebeb;
    }
    .header-title {
      font-weight: 700;
      font-size: 14px;
      color: #7719AA;
      letter-spacing: -0.2px;
    }

    /* ── Forms ── */
    .field { margin-bottom: 11px; }
    .field label {
      display: block;
      font-weight: 600;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #666;
      margin-bottom: 4px;
    }
    select, .email-info {
      width: 100%;
      padding: 7px 10px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 13px;
      color: #1a1a1a;
      background: #fff;
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%23999'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 10px center;
      cursor: pointer;
    }
    select:focus { outline: none; border-color: #7719AA; box-shadow: 0 0 0 2px rgba(119,25,170,0.12); }
    select:disabled { background-color: #f7f7f7; color: #aaa; cursor: not-allowed; }

    .hint {
      font-size: 11px;
      color: #999;
      margin-top: 3px;
    }
    .hint strong { color: #666; }

    /* ── Email preview ── */
    .email-info {
      background: #f9f9f9;
      font-size: 12px;
      color: #444;
      line-height: 1.5;
      min-height: 38px;
      cursor: default;
    }
    .email-info.empty { color: #bbb; font-style: italic; }

    /* ── Buttons ── */
    .btn {
      display: block;
      width: 100%;
      padding: 9px 14px;
      border: none;
      border-radius: 7px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      text-align: center;
      transition: background 0.15s, transform 0.1s;
      text-decoration: none;
    }
    .btn:active { transform: scale(0.98); }
    .btn-primary { background: #7719AA; color: #fff; }
    .btn-primary:hover:not(:disabled) { background: #6510a0; }
    .btn-primary:disabled { background: #c099d9; cursor: not-allowed; }
    .btn-ghost {
      background: none;
      color: #999;
      font-size: 11px;
      font-weight: 500;
      padding: 6px;
      text-decoration: underline;
      cursor: pointer;
      border: none;
      width: auto;
      display: inline-block;
    }
    .btn-ghost:hover { color: #555; }

    /* ── Status messages ── */
    .status {
      padding: 9px 12px;
      border-radius: 6px;
      font-size: 12px;
      line-height: 1.5;
      margin-top: 10px;
    }
    .status-success { background: #f0faf4; color: #1a7f45; border: 1px solid #bde2cc; }
    .status-error   { background: #fff5f5; color: #c0392b; border: 1px solid #f5c6cb; }
    .status-info    { background: #f4f0ff; color: #5a2d82; border: 1px solid #d9c6f0; }

    /* ── Connect screen ── */
    .connect-screen {
      text-align: center;
      padding: 24px 0 16px;
    }
    .connect-screen p {
      color: #666;
      line-height: 1.6;
      margin-bottom: 18px;
      font-size: 13px;
    }
    .connect-btn {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 10px 20px;
      background: #fff;
      border: 1.5px solid #ddd;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      color: #333;
      cursor: pointer;
      transition: border-color 0.15s, box-shadow 0.15s;
      text-decoration: none;
    }
    .connect-btn:hover {
      border-color: #7719AA;
      box-shadow: 0 2px 8px rgba(119,25,170,0.15);
    }

    /* ── Divider ── */
    hr { border: none; border-top: 1px solid #f0f0f0; margin: 14px 0; }

    /* ── Save link ── */
    .open-link {
      display: block;
      text-align: center;
      font-size: 12px;
      color: #7719AA;
      text-decoration: none;
      margin-top: 8px;
    }
    .open-link:hover { text-decoration: underline; }

    /* ── Footer ── */
    .footer {
      text-align: right;
      margin-top: 14px;
    }

    /* ── Spinner ── */
    @keyframes spin { to { transform: rotate(360deg); } }
    .spinner {
      display: inline-block;
      width: 14px;
      height: 14px;
      border: 2px solid #e0c9f5;
      border-top-color: #7719AA;
      border-radius: 50%;
      animation: spin 0.7s linear infinite;
      vertical-align: middle;
      margin-right: 6px;
    }
  </style>
</head>
<body>

  <!-- Header -->
  <div class="header">
    <svg width="26" height="26" viewBox="0 0 26 26" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect width="26" height="26" rx="6" fill="#7719AA"/>
      <path d="M7 8h4.5l2.5 6 2.5-6H21v10h-2.2V12l-2.3 6h-1.8L12.4 12v6H10V8H7z" fill="white"/>
      <path d="M5 17h3v1.5H5V17z" fill="white" opacity="0.5"/>
    </svg>
    <span class="header-title">Save to OneNote</span>
  </div>

  ${error ? `<div class="status status-error">${error}</div>` : ''}

  <!-- Auth screen (shown when not connected) -->
  <div id="connectScreen" class="connect-screen" style="display:none">
    <p>Connect your Microsoft account to save Missive emails directly into OneNote.</p>
    <button class="connect-btn" onclick="openAuthPopup()">
      <svg width="18" height="18" viewBox="0 0 18 18" fill="none"><rect width="8.5" height="8.5" fill="#F25022"/><rect x="9.5" width="8.5" height="8.5" fill="#7FBA00"/><rect y="9.5" width="8.5" height="8.5" fill="#00A4EF"/><rect x="9.5" y="9.5" width="8.5" height="8.5" fill="#FFB900"/></svg>
      Connect Microsoft Account
    </button>
  </div>

  <!-- Main app (shown when connected) -->
  <div id="appScreen" style="display:none">

    <div class="field">
      <label>Email</label>
      <div class="email-info empty" id="emailPreview">Open an email in Missive…</div>
    </div>

    <div class="field">
      <label>Notebook</label>
      <select id="notebookSel" onchange="onNotebookChange()" disabled>
        <option value="">Loading…</option>
      </select>
      <div class="hint" id="lastNotebookHint"></div>
    </div>

    <div class="field">
      <label>Section</label>
      <select id="sectionSel" onchange="onSectionChange()" disabled>
        <option value="">Select a notebook first</option>
      </select>
      <div class="hint" id="lastSectionHint"></div>
    </div>

    <hr>

    <button class="btn btn-primary" id="saveBtn" onclick="saveEmail()" disabled>
      Save Email to OneNote
    </button>

    <div id="statusArea"></div>
    <a id="openLink" class="open-link" href="#" target="_blank" rel="noopener" style="display:none">
      Open in OneNote ↗
    </a>

    <div class="footer">
      <button class="btn-ghost" onclick="disconnect()">Disconnect account</button>
    </div>
  </div>

  <!-- Missive SDK -->
  <script src="https://missiveapp.com/include/api.js"></script>

  <script>
    // ── State ──────────────────────────────────────────────────────────────────
    const WORKER_URL = '${workerUrl}';
    const LAST_LOC   = ${JSON.stringify(lastLocation)};

    let sessionId    = sessionStorage.getItem('mn_session') || null;
    let currentEmail = null; // { subject, fromName, fromEmail, toFields, ccFields, date, bodyHtml }
    let notebooks    = [];
    let sections     = [];
    let msUserId     = null;

    // ── Boot ───────────────────────────────────────────────────────────────────
    async function boot() {
      if (sessionId) {
        // Verify token is still valid
        try {
          const r = await apiFetch('/api/me');
          if (r.ok) {
            const me = await r.json();
            msUserId = me.id;
            showApp();
            loadNotebooks();
            return;
          }
        } catch(e) {}
        // Token invalid — clear session
        sessionId = null;
        sessionStorage.removeItem('mn_session');
      }
      showConnect();
    }

    function showConnect() {
      document.getElementById('connectScreen').style.display = '';
      document.getElementById('appScreen').style.display = 'none';
    }

    function showApp() {
      document.getElementById('connectScreen').style.display = 'none';
      document.getElementById('appScreen').style.display = '';
    }

    // ── Auth popup ─────────────────────────────────────────────────────────────
    function openAuthPopup() {
      const popup = window.open(
        WORKER_URL + '/auth/login',
        'msAuth',
        'width=520,height=640,left=100,top=100,scrollbars=yes'
      );
      window.addEventListener('message', async function onMsg(evt) {
        if (evt.data?.type !== 'ms_auth_success') return;
        window.removeEventListener('message', onMsg);
        if (evt.data.sessionId) {
          sessionId = evt.data.sessionId;
          sessionStorage.setItem('mn_session', sessionId);
          try { popup.close(); } catch(e) {}
          const r = await apiFetch('/api/me');
          if (r.ok) {
            const me = await r.json();
            msUserId = me.id;
          }
          showApp();
          loadNotebooks();
        }
      });
    }

    // ── Disconnect ─────────────────────────────────────────────────────────────
    async function disconnect() {
      if (!confirm('Disconnect your Microsoft account?')) return;
      await apiFetch('/auth/disconnect', { method: 'POST' });
      sessionId = null;
      sessionStorage.removeItem('mn_session');
      currentEmail = null;
      showConnect();
    }

    // ── API helper ─────────────────────────────────────────────────────────────
    function apiFetch(path, opts = {}) {
      return fetch(WORKER_URL + path, {
        ...opts,
        headers: {
          'X-Session-ID': sessionId || '',
          'Content-Type': 'application/json',
          ...(opts.headers || {}),
        },
      });
    }

    // ── Missive JS SDK ─────────────────────────────────────────────────────────
    Missive.on('change:conversations', async (ids) => {
      if (!ids || ids.length === 0) {
        currentEmail = null;
        updateEmailPreview();
        updateSaveButton();
        return;
      }
      try {
        const [conv] = await Missive.fetchConversations(ids);
        if (!conv) return;

        // Fetch the latest message for full body
        const msgIds = (conv.messages || []).map(m => m.id).slice(-3); // last 3
        const msgs   = msgIds.length ? await Missive.fetchMessages(msgIds) : [];
        const latest = msgs[msgs.length - 1] || {};
        const thread = msgs.slice(0, -1);

        currentEmail = {
          subject:   conv.subject || latest.subject || '(No subject)',
          fromName:  latest.from_field?.name  || '',
          fromEmail: latest.from_field?.address || '',
          toFields:  latest.to_fields  || [],
          ccFields:  latest.cc_fields  || [],
          date:      latest.delivered_at
                       ? new Date(latest.delivered_at * 1000).toLocaleString()
                       : new Date().toLocaleString(),
          bodyHtml:  latest.body || latest.preview || '',
          thread:    thread.map(m => ({
            fromName:  m.from_field?.name || m.from_field?.address || 'Unknown',
            date:      m.delivered_at ? new Date(m.delivered_at * 1000).toLocaleString() : '',
            bodyHtml:  m.body || m.preview || '',
          })),
        };
        updateEmailPreview();
        updateSaveButton();
      } catch(e) {
        console.error('Missive fetch error:', e);
      }
    });

    function updateEmailPreview() {
      const el = document.getElementById('emailPreview');
      if (!currentEmail) {
        el.textContent = 'Open an email in Missive…';
        el.className = 'email-info empty';
        return;
      }
      el.innerHTML =
        '<strong>' + esc(currentEmail.subject) + '</strong><br>' +
        'From: ' + esc(currentEmail.fromName || currentEmail.fromEmail);
      el.className = 'email-info';
    }

    // ── Notebooks ──────────────────────────────────────────────────────────────
    async function loadNotebooks() {
      const sel = document.getElementById('notebookSel');
      sel.innerHTML = '<option value="">Loading notebooks…</option>';
      sel.disabled = true;

      try {
        const r = await apiFetch('/api/notebooks');
        if (!r.ok) throw new Error('Could not load notebooks');
        const data = await r.json();
        notebooks = data.value || [];

        sel.innerHTML = '<option value="">Select a notebook…</option>';
        notebooks.forEach(nb => {
          const o = document.createElement('option');
          o.value = nb.id;
          o.textContent = nb.displayName;
          if (LAST_LOC?.notebookId === nb.id) o.selected = true;
          sel.appendChild(o);
        });
        sel.disabled = false;

        // Show last-used hint
        if (LAST_LOC?.notebookName) {
          document.getElementById('lastNotebookHint').innerHTML =
            'Last used: <strong>' + esc(LAST_LOC.notebookName) + '</strong>';
        }

        // Auto-load sections if last location matches
        if (LAST_LOC?.notebookId && sel.value === LAST_LOC.notebookId) {
          await loadSections(LAST_LOC.sectionId);
        }
      } catch(e) {
        sel.innerHTML = '<option value="">Error loading notebooks</option>';
        showStatus('error', 'Could not load notebooks: ' + e.message);
      }
    }

    function onNotebookChange() {
      loadSections();
    }

    async function loadSections(preselectId) {
      const notebookId = document.getElementById('notebookSel').value;
      const sel = document.getElementById('sectionSel');

      if (!notebookId) {
        sel.innerHTML = '<option value="">Select a notebook first</option>';
        sel.disabled = true;
        updateSaveButton();
        return;
      }

      sel.innerHTML = '<option value="">Loading sections…</option>';
      sel.disabled = true;
      updateSaveButton();

      try {
        const r = await apiFetch('/api/sections?notebookId=' + encodeURIComponent(notebookId));
        if (!r.ok) throw new Error('Could not load sections');
        const data = await r.json();
        sections = data.value || [];

        sel.innerHTML = '<option value="">Select a section…</option>';
        sections.forEach(sec => {
          const o = document.createElement('option');
          o.value = sec.id;
          o.textContent = sec.displayName;
          if ((preselectId || LAST_LOC?.sectionId) === sec.id &&
              document.getElementById('notebookSel').value === (LAST_LOC?.notebookId || '')) {
            o.selected = true;
          }
          sel.appendChild(o);
        });
        sel.disabled = false;

        // Show last-used hint
        if (LAST_LOC?.sectionName && notebookId === LAST_LOC?.notebookId) {
          document.getElementById('lastSectionHint').innerHTML =
            'Last used: <strong>' + esc(LAST_LOC.sectionName) + '</strong>';
        } else {
          document.getElementById('lastSectionHint').innerHTML = '';
        }

        updateSaveButton();
      } catch(e) {
        sel.innerHTML = '<option value="">Error loading sections</option>';
        showStatus('error', 'Could not load sections: ' + e.message);
      }
    }

    function onSectionChange() {
      updateSaveButton();
    }

    function updateSaveButton() {
      const notebookId = document.getElementById('notebookSel')?.value;
      const sectionId  = document.getElementById('sectionSel')?.value;
      const btn = document.getElementById('saveBtn');
      if (btn) btn.disabled = !(notebookId && sectionId && currentEmail);
    }

    // ── Save to OneNote ────────────────────────────────────────────────────────
    async function saveEmail() {
      const notebookSel = document.getElementById('notebookSel');
      const sectionSel  = document.getElementById('sectionSel');
      const notebookId  = notebookSel.value;
      const sectionId   = sectionSel.value;
      const notebookName = notebookSel.options[notebookSel.selectedIndex].text;
      const sectionName  = sectionSel.options[sectionSel.selectedIndex].text;
      const btn = document.getElementById('saveBtn');
      const linkEl = document.getElementById('openLink');

      if (!notebookId || !sectionId || !currentEmail) return;

      btn.disabled = true;
      btn.innerHTML = '<span class="spinner"></span>Saving…';
      linkEl.style.display = 'none';
      showStatus('info', 'Creating OneNote page…');

      try {
        const r = await apiFetch('/api/save', {
          method: 'POST',
          body: JSON.stringify({
            sectionId,
            notebookId,
            notebookName,
            sectionName,
            email: currentEmail,
          }),
        });

        const data = await r.json();
        if (!r.ok) throw new Error(data.error || 'Save failed');

        showStatus('success', 'Saved to <strong>' + esc(notebookName) + ' / ' + esc(sectionName) + '</strong>');
        if (data.pageUrl) {
          linkEl.href = data.pageUrl;
          linkEl.style.display = '';
        }
        btn.innerHTML = 'Save Email to OneNote';
        btn.disabled = false;
      } catch(e) {
        showStatus('error', 'Save failed: ' + e.message);
        btn.innerHTML = 'Save Email to OneNote';
        btn.disabled = false;
      }
    }

    // ── Utilities ──────────────────────────────────────────────────────────────
    function showStatus(type, html) {
      const el = document.getElementById('statusArea');
      el.innerHTML = '<div class="status status-' + type + '">' + html + '</div>';
    }

    function esc(str) {
      return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    // ── Init ───────────────────────────────────────────────────────────────────
    boot();
  </script>
</body>
</html>`;
}

// ─── HTML: Auth Popup Close Page ──────────────────────────────────────────────

function renderAuthSuccess(sessionId) {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Connected!</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      display: flex; align-items: center; justify-content: center;
      min-height: 100vh; margin: 0; background: #faf8ff;
    }
    .box {
      text-align: center; padding: 40px;
      background: #fff; border-radius: 12px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.08);
      max-width: 320px;
    }
    .check { font-size: 48px; margin-bottom: 16px; }
    h2 { color: #7719AA; font-size: 20px; margin-bottom: 8px; }
    p  { color: #666; font-size: 13px; }
  </style>
</head>
<body>
  <div class="box">
    <div class="check">✓</div>
    <h2>Connected!</h2>
    <p>You can close this window.</p>
  </div>
  <script>
    // Notify the parent sidebar
    if (window.opener) {
      window.opener.postMessage(
        { type: 'ms_auth_success', sessionId: ${JSON.stringify(sessionId)} },
        '*'
      );
      setTimeout(() => window.close(), 1500);
    }
  </script>
</body>
</html>`;
}

// ─── OneNote Page Builder ─────────────────────────────────────────────────────

function buildOneNotePage(email) {
  const {
    subject   = '(No subject)',
    fromName  = '',
    fromEmail = '',
    toFields  = [],
    ccFields  = [],
    date      = '',
    bodyHtml  = '',
    thread    = [],
  } = email;

  const toStr = toFields
    .map(f => (f.name ? `${f.name} <${f.address}>` : f.address))
    .join(', ');
  const ccStr = ccFields
    .map(f => (f.name ? `${f.name} <${f.address}>` : f.address))
    .join(', ');

  const fromStr = fromName ? `${fromName} &lt;${xmlEsc(fromEmail)}&gt;` : xmlEsc(fromEmail);

  // Previous messages thread (oldest first, collapsed styling)
  const threadHtml = thread.length
    ? thread.map(m => `
      <div style="border-left:3px solid #ccc;margin:12px 0;padding:8px 14px;background:#f9f9f9;">
        <p style="margin:0 0 6px;font-size:12px;color:#888;">
          <strong>${xmlEsc(m.fromName)}</strong> &mdash; ${xmlEsc(m.date)}
        </p>
        ${m.bodyHtml || ''}
      </div>`).join('')
    : '';

  return `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml"
      xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <title>${xmlEsc(subject)}</title>
  <meta name="created" content="${new Date().toISOString()}" />
</head>
<body>
  <div style="font-family:Calibri,Arial,sans-serif;max-width:860px;">

    <!-- Subject -->
    <h1 style="font-size:20px;color:#1a1a1a;margin:0 0 14px;padding-bottom:10px;border-bottom:3px solid #7719AA;">
      ${xmlEsc(subject)}
    </h1>

    <!-- Metadata table -->
    <table style="border-collapse:collapse;width:100%;margin-bottom:16px;font-size:13px;">
      <tr>
        <td style="padding:3px 10px 3px 0;color:#777;font-weight:700;white-space:nowrap;width:50px;">From</td>
        <td style="padding:3px 0;">${fromStr}</td>
      </tr>
      ${toStr ? `<tr>
        <td style="padding:3px 10px 3px 0;color:#777;font-weight:700;">To</td>
        <td style="padding:3px 0;">${xmlEsc(toStr)}</td>
      </tr>` : ''}
      ${ccStr ? `<tr>
        <td style="padding:3px 10px 3px 0;color:#777;font-weight:700;">CC</td>
        <td style="padding:3px 0;">${xmlEsc(ccStr)}</td>
      </tr>` : ''}
      <tr>
        <td style="padding:3px 10px 3px 0;color:#777;font-weight:700;">Date</td>
        <td style="padding:3px 0;">${xmlEsc(date)}</td>
      </tr>
    </table>

    <hr style="border:none;border-top:1px solid #e0e0e0;margin:0 0 16px;" />

    <!-- Email body (preserves original HTML formatting) -->
    <div style="line-height:1.65;font-size:14px;">
      ${bodyHtml}
    </div>

    ${threadHtml ? `
      <hr style="border:none;border-top:1px solid #e0e0e0;margin:24px 0 14px;" />
      <h3 style="font-size:13px;font-weight:700;color:#888;margin:0 0 10px;text-transform:uppercase;letter-spacing:0.5px;">
        Earlier messages in thread
      </h3>
      ${threadHtml}
    ` : ''}

    <hr style="border:none;border-top:1px solid #efefef;margin:24px 0 8px;" />
    <p style="font-size:11px;color:#bbb;">Saved via ZAGO Missive→OneNote integration</p>

  </div>
</body>
</html>`;
}

function xmlEsc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ─── Main Worker ──────────────────────────────────────────────────────────────

export default {
  async fetch(request, env) {
    const url       = new URL(request.url);
    const workerUrl = `${url.protocol}//${url.host}`;
    const path      = url.pathname;

    // ── CORS preflight (for cross-origin fetches from Missive iframe) ──────────
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: corsHeaders(request),
      });
    }

    try {
      // ── Health check ───────────────────────────────────────────────────────
      if (path === '/health') {
        return json({ status: 'ok', timestamp: new Date().toISOString() });
      }

      // ── Auth: initiate Microsoft OAuth ─────────────────────────────────────
      if (path === '/auth/login') {
        const state     = randomId(32);
        const sessionId = randomId(32);

        // Store state → sessionId mapping (expires in 10 min)
        await env.ONENOTE_KV.put(`oauth_state:${state}`, sessionId, { expirationTtl: 600 });

        const authUrl = new URL(MS_AUTH_URL);
        authUrl.searchParams.set('client_id',     env.MICROSOFT_CLIENT_ID);
        authUrl.searchParams.set('response_type', 'code');
        authUrl.searchParams.set('redirect_uri',  `${workerUrl}/auth/callback`);
        authUrl.searchParams.set('scope',         MS_SCOPES);
        authUrl.searchParams.set('state',         state);
        authUrl.searchParams.set('response_mode', 'query');
        authUrl.searchParams.set('prompt',        'select_account');

        return Response.redirect(authUrl.toString(), 302);
      }

      // ── Auth: handle OAuth callback ────────────────────────────────────────
      if (path === '/auth/callback') {
        const code  = url.searchParams.get('code');
        const state = url.searchParams.get('state');
        const err   = url.searchParams.get('error');
        const errDesc = url.searchParams.get('error_description') || err;

        if (err) {
          return html500(`Microsoft sign-in was cancelled or failed: ${errDesc}`);
        }

        const sessionId = await env.ONENOTE_KV.get(`oauth_state:${state}`);
        if (!sessionId) {
          return html500('Auth state expired or invalid. Please try again.');
        }
        await env.ONENOTE_KV.delete(`oauth_state:${state}`);

        // Exchange code for tokens
        const tokenRes = await fetch(MS_TOKEN_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            client_id:     env.MICROSOFT_CLIENT_ID,
            client_secret: env.MICROSOFT_CLIENT_SECRET,
            code,
            redirect_uri:  `${workerUrl}/auth/callback`,
            grant_type:    'authorization_code',
            scope:         MS_SCOPES,
          }),
        });

        if (!tokenRes.ok) {
          const body = await tokenRes.text();
          return html500(`Failed to exchange token: ${body.slice(0, 200)}`);
        }

        const tokens = await tokenRes.json();
        tokens.expires_at = Date.now() + tokens.expires_in * 1000;
        await storeTokens(env, sessionId, tokens);

        // Persist MS user ID for last-location keying
        try {
          const me = await graph(tokens.access_token, '/me?$select=id,displayName');
          await env.ONENOTE_KV.put(
            `user_id:${sessionId}`,
            me.id,
            { expirationTtl: 60 * 60 * 24 * 30 }
          );
        } catch (_) {}

        return htmlPage(renderAuthSuccess(sessionId));
      }

      // ── Auth: disconnect ───────────────────────────────────────────────────
      if (path === '/auth/disconnect' && request.method === 'POST') {
        const sessionId = getSessionId(request);
        if (sessionId) {
          const userId = await env.ONENOTE_KV.get(`user_id:${sessionId}`);
          await env.ONENOTE_KV.delete(`session:${sessionId}`);
          await env.ONENOTE_KV.delete(`user_id:${sessionId}`);
          if (userId) await env.ONENOTE_KV.delete(`last_location:${userId}`);
        }
        return json({ ok: true }, 200, request);
      }

      // ── API: current MS user (for session validation) ──────────────────────
      if (path === '/api/me') {
        const sessionId   = getSessionId(request);
        const accessToken = await getValidToken(env, sessionId, workerUrl);
        if (!accessToken) return json({ error: 'Not authenticated' }, 401, request);

        const me = await graph(accessToken, '/me?$select=id,displayName,mail');
        return json(me, 200, request);
      }

      // ── API: list notebooks ────────────────────────────────────────────────
      if (path === '/api/notebooks') {
        const sessionId   = getSessionId(request);
        const accessToken = await getValidToken(env, sessionId, workerUrl);
        if (!accessToken) return json({ error: 'Not authenticated' }, 401, request);

        const data = await graph(
          accessToken,
          '/me/onenote/notebooks?$orderby=displayName&$select=id,displayName,links'
        );
        return json(data, 200, request);
      }

      // ── API: list sections ─────────────────────────────────────────────────
      if (path === '/api/sections') {
        const sessionId   = getSessionId(request);
        const accessToken = await getValidToken(env, sessionId, workerUrl);
        if (!accessToken) return json({ error: 'Not authenticated' }, 401, request);

        const notebookId = url.searchParams.get('notebookId');
        if (!notebookId) return json({ error: 'notebookId required' }, 400, request);

        const data = await graph(
          accessToken,
          `/me/onenote/notebooks/${notebookId}/sections?$orderby=displayName&$select=id,displayName`
        );
        return json(data, 200, request);
      }

      // ── API: save email to OneNote ─────────────────────────────────────────
      if (path === '/api/save' && request.method === 'POST') {
        const sessionId   = getSessionId(request);
        const accessToken = await getValidToken(env, sessionId, workerUrl);
        if (!accessToken) return json({ error: 'Not authenticated' }, 401, request);

        let body;
        try { body = await request.json(); }
        catch { return json({ error: 'Invalid JSON body' }, 400, request); }

        const { sectionId, notebookId, notebookName, sectionName, email } = body;
        if (!sectionId || !email) {
          return json({ error: 'sectionId and email are required' }, 400, request);
        }

        // Build OneNote XHTML page from email data
        const pageHtml = buildOneNotePage(email);

        // POST to OneNote API
        const createRes = await fetch(
          `${GRAPH_BASE}/me/onenote/sections/${sectionId}/pages`,
          {
            method: 'POST',
            headers: {
              Authorization:  `Bearer ${accessToken}`,
              'Content-Type': 'application/xhtml+xml',
            },
            body: pageHtml,
          }
        );

        if (!createRes.ok) {
          const errText = await createRes.text();
          return json(
            { error: `OneNote API error ${createRes.status}: ${errText.slice(0, 300)}` },
            500,
            request
          );
        }

        const page = await createRes.json();

        // Persist last-used location
        const userId = await env.ONENOTE_KV.get(`user_id:${sessionId}`);
        if (userId) {
          await setLastLocation(env, userId, {
            notebookId,
            notebookName,
            sectionId,
            sectionName,
          });
        }

        return json({
          success:  true,
          pageId:   page.id,
          pageUrl:  page.links?.oneNoteWebUrl?.href
                 ?? page.links?.oneNoteClientUrl?.href
                 ?? null,
          title:    page.title,
        }, 201, request);
      }

      // ── Sidebar UI (main app page) ─────────────────────────────────────────
      if (path === '/' || path === '') {
        const sessionId   = getSessionId(request) || request.headers.get('X-Session-ID');
        let lastLocation  = null;

        if (sessionId) {
          const userId = await env.ONENOTE_KV.get(`user_id:${sessionId}`);
          if (userId) lastLocation = await getLastLocation(env, userId);
        }

        return htmlPage(renderSidebar({ workerUrl, lastLocation }));
      }

      return new Response('Not found', { status: 404 });

    } catch (err) {
      console.error('Worker error:', err);
      return json({ error: err.message }, 500, request);
    }
  },
};

// ─── Response Helpers ─────────────────────────────────────────────────────────

function corsHeaders(request) {
  return {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Session-ID',
    'Access-Control-Max-Age':       '86400',
  };
}

function json(data, status = 200, request) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...corsHeaders(request || {}),
    },
  });
}

function htmlPage(content) {
  return new Response(content, {
    headers: {
      'Content-Type': 'text/html; charset=utf-8',
      // Allow embedding in Missive's iframe
      'X-Frame-Options':         'ALLOWALL',
      'Content-Security-Policy': "frame-ancestors 'self' https://missiveapp.com https://*.missiveapp.com;",
    },
  });
}

function html500(message) {
  return new Response(
    `<!DOCTYPE html><html><body style="font-family:sans-serif;padding:40px;text-align:center;">
      <h2 style="color:#c0392b">Something went wrong</h2>
      <p style="color:#666">${xmlEsc(message)}</p>
      <p style="margin-top:20px;font-size:13px;color:#999">Close this window and try again.</p>
    </body></html>`,
    { status: 500, headers: { 'Content-Type': 'text/html; charset=utf-8' } }
  );
}
