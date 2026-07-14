import "dotenv/config";
import express from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promises as fs } from "node:fs";
import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
} from "@azure/msal-node";
import { WorkIqMcpClient } from "../src/mcp.ts";
import { WorkIqAgent } from "./agent.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CACHE_PATH = path.join(__dirname, "..", ".token-cache.json");
const WORK_IQ_SCOPE = "api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask";

const TENANT_ID = process.env.TENANT_ID!;
const CLIENT_ID = process.env.CLIENT_ID!;
const TIME_ZONE = process.env.TIME_ZONE ?? "America/New_York";
const PORT = Number(process.env.PORT ?? 3001);

if (!TENANT_ID || !CLIENT_ID) {
  console.error("Missing TENANT_ID / CLIENT_ID in .env");
  process.exit(1);
}

// ---- MSAL public client with file-backed token cache (shared with Demo 1) ----
function cachePlugin() {
  return {
    beforeCacheAccess: async (ctx: any) => {
      try {
        ctx.tokenCache.deserialize(await fs.readFile(CACHE_PATH, "utf-8"));
      } catch {
        /* first run */
      }
    },
    afterCacheAccess: async (ctx: any) => {
      if (ctx.cacheHasChanged) {
        await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize(), "utf-8");
      }
    },
  };
}

const config: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
  cache: { cachePlugin: cachePlugin() },
  system: {
    loggerOptions: {
      loggerCallback: (lvl, msg) => {
        if (lvl === LogLevel.Error) console.error(msg);
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Warning,
    },
  },
};
const pca = new PublicClientApplication(config);

async function getSilentToken(): Promise<string | null> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts.length) return null;
  try {
    const r = await pca.acquireTokenSilent({
      account: accounts[0],
      scopes: [WORK_IQ_SCOPE],
    });
    return r?.accessToken ?? null;
  } catch {
    return null;
  }
}

async function currentUsername(): Promise<string | null> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0]?.username ?? null;
}

// ---- Device-code sign-in state (single-user demo) ----
type DeviceFlow = {
  state: "pending" | "done" | "error";
  userCode?: string;
  verificationUri?: string;
  message?: string;
  error?: string;
};
let deviceFlow: DeviceFlow | null = null;

// ---- Work IQ MCP client + LangChain agent (lazily created after sign-in) ----
const mcp = new WorkIqMcpClient(async () => {
  const token = await getSilentToken();
  if (!token) throw new Error("Not signed in.");
  return token;
});
let agent: WorkIqAgent | null = null;

async function getAgent(): Promise<WorkIqAgent> {
  if (!agent) {
    agent = new WorkIqAgent(mcp, await currentUsername(), TIME_ZONE);
  }
  return agent;
}

// ---------------------------------------------------------------------------
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/status", async (_req, res) => {
  const token = await getSilentToken();
  const accounts = await pca.getTokenCache().getAllAccounts();
  res.json({
    authenticated: !!token,
    account: accounts[0]?.username ?? null,
    provider: process.env.AGENT_PROVIDER ?? "azure",
    timeZone: TIME_ZONE,
  });
});

app.post("/api/login", async (_req, res) => {
  if (deviceFlow?.state === "pending" && deviceFlow.userCode) {
    return res.json(deviceFlow);
  }
  deviceFlow = { state: "pending" };
  await new Promise<void>((resolve) => {
    pca
      .acquireTokenByDeviceCode({
        scopes: [WORK_IQ_SCOPE],
        deviceCodeCallback: (r) => {
          deviceFlow = {
            state: "pending",
            userCode: r.userCode,
            verificationUri: r.verificationUri,
            message: r.message,
          };
          resolve();
        },
      })
      .then(() => {
        deviceFlow = { state: "done" };
      })
      .catch((e) => {
        deviceFlow = { state: "error", error: String(e?.message ?? e) };
        resolve();
      });
  });
  res.json(deviceFlow);
});

app.get("/api/login/poll", (_req, res) => {
  res.json(deviceFlow ?? { state: "error", error: "No login in progress." });
});

app.post("/api/logout", async (_req, res) => {
  const cache = pca.getTokenCache();
  for (const a of await cache.getAllAccounts()) await cache.removeAccount(a);
  await fs.rm(CACHE_PATH, { force: true });
  agent = null;
  deviceFlow = null;
  res.json({ ok: true });
});

app.post("/api/new-conversation", async (_req, res) => {
  agent = null;
  res.json({ ok: true });
});

app.post("/api/chat", async (req, res) => {
  const text = String(req.body?.text ?? "").trim();
  if (!text) return res.status(400).json({ error: "Empty message." });

  try {
    const a = await getAgent();
    const reply = await a.chat(text);
    res.json({ text: reply });
  } catch (e) {
    res.status(500).json({ error: String((e as Error).message) });
  }
});

app.listen(PORT, () => {
  console.log(`\n  Work IQ Agent demo running at  http://localhost:${PORT}\n`);
});
