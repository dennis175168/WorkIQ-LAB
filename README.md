# WorkIQ LAB

A small test harness for the **Microsoft 365 Work IQ REST API** (Copilot Chat API).
It signs a user in with Microsoft Entra (delegated auth via the **device-code flow**),
creates a conversation, and lets you chat with Microsoft 365 Copilot from your terminal —
in either synchronous or streamed mode.

> Work IQ only supports **delegated** permissions (signed-in user). App-only / client-credential
> flows are **not** supported.

## API summary

| Action              | Method & path                                                                 |
| ------------------- | ----------------------------------------------------------------------------- |
| Base URL            | `https://workiq.svc.cloud.microsoft/rest` (prod) / `.../rest/beta` (beta) |
| Create conversation | `POST /conversations` with body `{}`                                      |
| Chat (sync)         | `POST /conversations/{id}/chat`                                             |
| Chat (streamed)     | `POST /conversations/{id}/chatOverStream`                                   |
| Auth scope          | `api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask`                          |

Docs: [https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/rest/overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/rest/overview)

## Prerequisites

1. **Work IQ enabled** in your tenant by an admin.
   See [Enable Work IQ](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/enable-work-iq).
2. A **Microsoft Entra app registration** (public client):
   - Add delegated API permission **`WorkIQAgent.Ask`** (search for the *Work IQ* API, app ID URI `api://workiq.svc.cloud.microsoft`) and grant **admin consent**.
   - Under **Authentication → Advanced settings**, set **Allow public client flows** to **Yes** (required for device-code flow).
3. **Node.js 18+**.

## Setup

```powershell
npm install
Copy-Item .env.example .env
# then edit .env with your TENANT_ID and CLIENT_ID
```

`.env` values:

| Variable      | Description                                                        |
| ------------- | ------------------------------------------------------------------ |
| `TENANT_ID` | Directory (tenant) ID, or`organizations`                         |
| `CLIENT_ID` | Application (client) ID of your app registration                   |
| `USE_BETA`  | `true` to hit the `/beta` endpoint, else `false`             |
| `TIME_ZONE` | IANA time zone sent as`locationHint` (e.g. `America/New_York`) |

## Run

```powershell
# synchronous chat REPL
npm start

# start in streaming mode
npm start -- --stream
```

On first run you'll see a device-code message — open [https://microsoft.com/devicelogin](https://microsoft.com/devicelogin),
enter the code, and sign in. The token is cached in `.token-cache.json` so later runs are silent.

### REPL commands

- `/stream` — switch to the streamed endpoint
- `/sync` — switch to the synchronous endpoint
- `/exit` — quit

## Web demo

A simple browser-based chat UI (Express backend + static frontend) is included under
[`web/`](./web). The backend handles delegated sign-in (device-code flow) and proxies
requests to the Work IQ API, so the browser never calls the API directly (avoids CORS
and keeps tokens server-side).

```powershell
npm run web
```

Then open **[http://localhost:3000](http://localhost:3000)** in your browser.

- If you've already signed in via `npm start`, the cached token is reused and you go
  straight to the chat screen.
- Otherwise click **登入 / Sign in** — the page shows a device code and a link
  ([https://login.microsoft.com/device](https://login.microsoft.com/device)). After you sign in, the chat screen appears
  automatically.

Features: markdown + citation rendering, a **web-search** on/off toggle (single-turn),
**new conversation**, and **sign out**. Set a custom port with the `PORT` env var
(e.g. `PORT=8080`).

Web demo files:

```
web/
  server.ts          # Express backend: auth + Work IQ proxy
  public/index.html  # chat UI markup
  public/app.js      # frontend logic (device-code login + chat)
  public/styles.css  # styling
```

## Manual testing without Node

Open [`requests.http`](./requests.http) with the VS Code **REST Client** extension,
paste a delegated access token, and fire requests directly.

## Project layout

```
src/
  auth.ts     # MSAL device-code flow + token cache
  workiq.ts   # typed Work IQ Chat API client (create / chat / chatOverStream)
  index.ts    # interactive CLI
web/          # browser-based chat demo (see "Web demo" above)
scripts/
  provisionWorkIqSp.ts  # one-off: create the Work IQ service principal in the tenant
requests.http # manual request collection
```

## Known limitations (from the API docs)

- Text responses only; no file creation, email, meeting scheduling, code interpreter, or image tools.
- No long-running tasks (prone to gateway timeouts).
- Web + enterprise search grounding are on by default; disabling web search is a per-message action.
- `/beta` APIs may change and aren't supported for production.

## Work IQ official docs

Microsoft Learn documentation for the Work IQ REST API and related setup:

- [Work IQ REST API overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/rest/overview)
- [Enable Work IQ](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/enable-work-iq)
- [Work IQ extensibility overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/overview)
- [Authentication & delegated permissions (Microsoft Entra)](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-device-code)
