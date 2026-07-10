import {
  PublicClientApplication,
  LogLevel,
  type Configuration,
  type AuthenticationResult,
} from "@azure/msal-node";
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CACHE_PATH = path.join(__dirname, "..", ".token-cache.json");

/** Work IQ delegated scope. Admin consent required in the tenant. */
export const WORK_IQ_SCOPE = "api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask";

/**
 * Simple file-backed token cache so you don't have to complete the device-code
 * flow on every run.
 */
function buildCachePlugin() {
  return {
    beforeCacheAccess: async (ctx: {
      tokenCache: { deserialize: (s: string) => void };
    }) => {
      try {
        const data = await fs.readFile(CACHE_PATH, "utf-8");
        ctx.tokenCache.deserialize(data);
      } catch {
        // No cache yet — first run.
      }
    },
    afterCacheAccess: async (ctx: {
      cacheHasChanged: boolean;
      tokenCache: { serialize: () => string };
    }) => {
      if (ctx.cacheHasChanged) {
        await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize(), "utf-8");
      }
    },
  };
}

function buildClient(tenantId: string, clientId: string): PublicClientApplication {
  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin: buildCachePlugin() },
    system: {
      loggerOptions: {
        loggerCallback: (level, message) => {
          if (level === LogLevel.Error) console.error(message);
        },
        piiLoggingEnabled: false,
        logLevel: LogLevel.Warning,
      },
    },
  };
  return new PublicClientApplication(config);
}

/**
 * Acquire a delegated access token for the Work IQ API.
 * Tries the silent cache first, then falls back to the device-code flow.
 */
export async function getAccessToken(
  tenantId: string,
  clientId: string
): Promise<string> {
  const pca = buildClient(tenantId, clientId);
  const scopes = [WORK_IQ_SCOPE];

  // Attempt silent acquisition from the cached account.
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const silent = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes,
      });
      if (silent?.accessToken) return silent.accessToken;
    } catch {
      // Silent failed (expired/no refresh token) — fall through to device code.
    }
  }

  const result: AuthenticationResult | null = await pca.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      console.log("\n" + "=".repeat(60));
      console.log(response.message);
      console.log("=".repeat(60) + "\n");
    },
  });

  if (!result?.accessToken) {
    throw new Error("Failed to acquire an access token.");
  }
  return result.accessToken;
}
