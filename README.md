# WorkIQ LAB

Two browser demos for the **Microsoft 365 Work IQ API**, showing the two ways an
app can use Work IQ:

| Demo | Folder | What it is | Work IQ surface |
| --- | --- | --- | --- |
| **Demo 1 — Chat** | [`demo-rest/`](./demo-rest) | A website that calls the Work IQ **REST Chat API** directly. Pure grounded Q&A. | `POST /conversations/{id}/chat` |
| **Demo 2 — Agent** | [`demo-agent/`](./demo-agent) | A website with a custom **LangChain.js agent** whose tools are the Work IQ **MCP** server. The LLM decides when to answer vs. take an action. | MCP `ask` + `do_action` |

Both sign the user in with Microsoft Entra (delegated **device-code flow**) and keep
tokens server-side. Work IQ only supports **delegated** permissions (signed-in user).

```
User ──▶ Demo 1 website ──▶ Work IQ REST Chat API      (ask questions, text back)

User ──▶ Demo 2 website ──▶ LangChain agent (LLM brain) ──▶ Work IQ MCP
                                                              ├─ ask        → grounded chat
                                                              └─ do_action  → real actions (send mail)
```

---

## Prerequisites

1. **Work IQ enabled** in your tenant by an admin —
   [Enable Work IQ](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/enable-work-iq).
2. A **Microsoft Entra app registration** (public client):
   - Delegated API permission **`WorkIQAgent.Ask`** (Work IQ API, app ID URI
     `api://workiq.svc.cloud.microsoft`), with **admin consent**.
   - **Authentication → Allow public client flows → Yes** (for device-code flow).
3. **Node.js 18+**.
4. **Demo 2 only:** an LLM with tool-calling (Microsoft Foundry / Azure OpenAI /
   OpenAI / GitHub Models).

---

## Setup

```bash
npm install
cp .env.example .env      # then edit .env (see below)
```

### `.env` — shared (both demos)

| Variable | Description |
| --- | --- |
| `TENANT_ID` | Directory (tenant) ID, or `organizations` |
| `CLIENT_ID` | Application (client) ID of your Entra app registration |
| `USE_BETA` | `true` to hit the `/beta` Chat endpoint, else `false` (Demo 1) |
| `TIME_ZONE` | IANA time zone sent as `locationHint` (e.g. `America/New_York`) |

### `.env` — Demo 2 agent brain

Pick a provider with `AGENT_PROVIDER` and fill its variables:

```bash
# Microsoft Foundry (OpenAI-compatible /openai/v1 endpoint + API key)
AGENT_PROVIDER=foundry
FOUNDRY_ENDPOINT=https://<resource>.services.ai.azure.com/openai/v1   # note: /openai/v1, NOT /responses
FOUNDRY_DEPLOYMENT=<deployment-name>                                  # e.g. gpt-4o / gpt-chat-latest
FOUNDRY_API_KEY=<key from the Foundry "Call model" panel>

# — or — OpenAI          AGENT_PROVIDER=openai   OPENAI_API_KEY=...  AGENT_MODEL=gpt-4o-mini
# — or — GitHub Models   AGENT_PROVIDER=github   GITHUB_TOKEN=...    AGENT_MODEL=gpt-4o-mini
# — or — Azure OpenAI    AGENT_PROVIDER=azure    AZURE_OPENAI_API_KEY / _INSTANCE_NAME / _DEPLOYMENT_NAME / _VERSION
```

> Tip: prefer a fast chat model (e.g. `gpt-4o`). Reasoning models like `gpt-5` are
> slow (30–40 s/turn) and more prone to rate limits.

---

## How to start the demos

### Demo 1 — Work IQ REST Chat  →  http://localhost:3000

```bash
npm run web:rest
```

### Demo 2 — Agent (LangChain + Work IQ MCP)  →  http://localhost:3001

```bash
npm run web:agent
```

Run both at once in two terminals. Change the port with `PORT`, e.g.
`PORT=8080 npm run web:rest`.

**First-time sign-in:** the page shows a device code and a link. Open
[https://microsoft.com/devicelogin](https://microsoft.com/devicelogin), enter the
code, and sign in. The token is cached in `.token-cache.json` (shared by both
demos), so later runs go straight to the chat screen.

### CLI (optional, Demo 1 style)

```bash
npm start            # interactive REST chat REPL
npm start -- --stream
```

REPL commands: `/stream`, `/sync`, `/exit`.

---

## Using the demos

**Demo 1** — ask questions grounded in your Microsoft 365 data, e.g.
"我今天有哪些會議？", "摘要我未讀的 Outlook 郵件". It only returns text.

**Demo 2** — same questions work (the agent calls the MCP `ask` tool), **and** it can
take actions from natural language, e.g. "發一封提醒信件給我自己記得報帳" → the agent
calls `do_action /me/sendMail`.

### Actions need a tenant policy (Demo 2)

Work IQ MCP **blocks all mutations (create/update/delete/action, incl. sending mail)
by default**. To allow `do_action /me/sendMail`, a tenant admin must enable the mail
mutation scenario in the Microsoft 365 admin center:
**Agents → Tools → Work IQ MCP → Policy** (can take up to 24 h to apply). Until then,
send attempts return `Path is not in the policy allowlist`. No code change is needed
once the policy is enabled. See
[Policy governance for Work IQ MCP](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/policy-governance-mcp).

---

## Manual REST testing (no Node)

Open [`requests.http`](./requests.http) with the VS Code **REST Client** extension,
paste a delegated access token, and fire requests directly at the Work IQ REST API.

---

## Project layout

```
src/                     # shared clients
  auth.ts                #   MSAL device-code flow + token cache (WorkIQAgent.Ask)
  workiq.ts              #   Work IQ REST Chat client (Demo 1)
  mcp.ts                 #   Work IQ MCP client (Demo 2)
  index.ts               #   CLI (Demo 1 chat REPL)
demo-rest/               # Demo 1 website — HTML + Work IQ REST Chat API
  server.ts              #   Express backend (auth + Chat proxy)
  public/                #   index.html, app.js, styles.css
demo-agent/              # Demo 2 website — HTML + LangChain agent + Work IQ MCP
  server.ts              #   Express backend (auth + agent)
  agent.ts               #   LangChain tool-calling agent over MCP (ask / do_action)
  public/                #   index.html, app.js, styles.css
scripts/
  provisionWorkIqSp.ts   #   one-off: create the Work IQ service principal (Node)
  provision-workiq-sp.ps1#   one-off: same, PowerShell
requests.http            # manual REST request collection (Demo 1)
```

### npm scripts

| Script | Does |
| --- | --- |
| `npm run web:rest` | Start Demo 1 (REST Chat) — default port 3000 |
| `npm run web:agent` | Start Demo 2 (Agent) — default port 3001 |
| `npm start` | CLI Demo 1 chat REPL |
| `npm run typecheck` | `tsc --noEmit` |

---

## Work IQ API cheat-sheet

| Surface | Endpoint / call | Used by |
| --- | --- | --- |
| REST Chat | `POST /conversations`, `/conversations/{id}/chat`, `/chatOverStream` | Demo 1 |
| MCP | `https://workiq.svc.cloud.microsoft/mcp` — tools `ask`, `do_action`, `fetch`, … | Demo 2 |
| Auth scope | `api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask` (delegated) | both |

The Work IQ **REST API has no action/tool endpoints** — it's chat only. Actions
(the "Work IQ Tool API") are exposed **only through MCP** today.

## Docs

- [Work IQ API overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/api-overview)
- [Work IQ REST Chat API](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/rest/overview)
- [Work IQ MCP overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/overview) ·
  [tool reference](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/tool-reference) ·
  [policy governance](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/policy-governance-mcp)
- [Enable Work IQ](https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/enable-work-iq)
