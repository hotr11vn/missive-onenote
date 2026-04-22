# Missive → OneNote Integration

A Cloudflare Worker that adds a **sidebar app** and **quick action** to Missive, letting you save any email directly into a Microsoft OneNote notebook — with full HTML formatting preserved, and memory of your last-used notebook/section.

Works on **desktop and mobile (iOS/Android)**.

---

## How it works

1. The Worker serves a sidebar UI that Missive loads inside an iFrame.
2. When you open an email in Missive, the sidebar reads the subject, sender, recipients, date, and full HTML body using the Missive JS SDK.
3. A **quick action** ("Save to OneNote") also appears in the conversation context menu — tap it to open the sidebar instantly.
4. You pick a OneNote **Notebook** and **Section** (dropdowns auto-populate from your account). Your last choice is remembered.
5. Click **Save** — the Worker creates a OneNote page via Microsoft Graph API, preserving all formatting.

---

## Fixes applied (vs. original)

| # | Issue | Root cause | Fix |
|---|-------|-----------|-----|
| 1 | **Mobile/iOS blank screen** | `frame-ancestors` CSP header blocked the iframe on iOS (Missive iOS loads from `localhost`) | Removed `frame-ancestors` directive entirely per [Missive iOS docs](https://missiveapp.com/docs/developers/ui-iframe-integrations#ios-compatibility) |
| 2 | **Quick action missing** | `Missive.setActions()` was never called | Added `setActions()` registering a "Save to OneNote" action in the `conversation` context |
| 3 | **SDK URL outdated** | Used deprecated `https://missiveapp.com/include/api.js` | Updated to `https://integrations.missiveapp.com/missive.js` |
| 4 | **iOS OAuth popup blocked** | `window.open()` popups are blocked inside iframes on iOS | Replaced with `Missive.initiateCallback()` — the official iOS-compatible OAuth flow; also switched session storage from `sessionStorage` to `Missive.storeSet/storeGet` |

---

## Setup

### 1. Azure App Registration (Microsoft)

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: `Missive OneNote Integration` (or anything you like)
3. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
4. Redirect URI:
   - Platform: **Web**
   - URI: `https://missive-onenote.YOUR-ACCOUNT.workers.dev/auth/callback`
5. Click **Register**

After registering:

- Copy the **Application (client) ID** → `MICROSOFT_CLIENT_ID`
- Go to **Certificates & secrets** → **New client secret** → copy the **Value** → `MICROSOFT_CLIENT_SECRET`
- Go to **API permissions** → **Microsoft Graph** → **Delegated** → add:
  - `Notes.ReadWrite`
  - `User.Read`
  - `offline_access`

---

### 2. Cloudflare Setup

```bash
npm install
npx wrangler login
```

Create KV namespace:

```bash
npm run kv:create
```

Paste the returned `id` into `wrangler.toml`, then set secrets:

```bash
npx wrangler secret put MICROSOFT_CLIENT_ID
npx wrangler secret put MICROSOFT_CLIENT_SECRET
```

---

### 3. Deploy

```bash
npm run deploy
```

Update the Azure redirect URI to your Worker URL:
`https://missive-onenote.YOUR-ACCOUNT.workers.dev/auth/callback`

---

### 4. Add to Missive

1. Missive → **Settings** → **Integrations** → **Custom integrations** → **New integration**
2. Type: **Sidebar app (iFrame)**
3. URL: `https://missive-onenote.YOUR-ACCOUNT.workers.dev/`
4. Name: `Save to OneNote`

Once added, the **quick action** ("Save to OneNote") will automatically appear in conversation context menus on both desktop and mobile.

---

## Environment variables

| Variable | Description |
|---|---|
| `MICROSOFT_CLIENT_ID` | Azure App Registration Client ID |
| `MICROSOFT_CLIENT_SECRET` | Azure App Registration Client Secret |

KV namespace binding `ONENOTE_KV` is configured in `wrangler.toml`.

---

## Useful commands

```bash
npm run deploy    # Deploy to Cloudflare
npm run dev       # Run locally
npm run logs      # Tail Worker logs
```
