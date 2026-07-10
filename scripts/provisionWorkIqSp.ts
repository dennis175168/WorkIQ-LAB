/**
 * One-off admin utility: provision the Work IQ service principal in the tenant.
 *
 * Uses the Microsoft Graph Command Line Tools public client (device-code flow)
 * to get a delegated Graph token with Application.ReadWrite.All, then POSTs the
 * Work IQ appId to /servicePrincipals.
 *
 * Run:  node --import tsx scripts/provisionWorkIqSp.ts
 * Sign in as a Global Administrator / Application Administrator.
 */
import { PublicClientApplication, type Configuration } from "@azure/msal-node";

const TENANT_ID = "1a632370-87d8-4768-a8d0-7a9a728dd03d";
const WORK_IQ_APP_ID = "fdcc1f02-fc51-4226-8753-f668596af7f7";
// Microsoft Graph Command Line Tools — first-party public client.
const GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

async function main() {
  const config: Configuration = {
    auth: {
      clientId: GRAPH_CLI_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    },
  };
  const pca = new PublicClientApplication(config);

  const result = await pca.acquireTokenByDeviceCode({
    scopes: ["https://graph.microsoft.com/Application.ReadWrite.All"],
    deviceCodeCallback: (r) => {
      console.log("\n" + "=".repeat(60));
      console.log(r.message);
      console.log("=".repeat(60) + "\n");
    },
  });

  if (!result?.accessToken) throw new Error("Failed to get a Graph token.");
  console.log(`Signed in as: ${result.account?.username}`);

  const res = await fetch("https://graph.microsoft.com/v1.0/servicePrincipals", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${result.accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ appId: WORK_IQ_APP_ID }),
  });

  const bodyText = await res.text();
  if (res.status === 201) {
    const sp = JSON.parse(bodyText);
    console.log(`\n✅ Created Work IQ service principal. objectId=${sp.id}`);
  } else if (
    res.status === 409 ||
    /already exists|conflict/i.test(bodyText)
  ) {
    console.log("\n✅ Work IQ service principal already exists (conflict).");
  } else {
    console.error(`\n❌ Graph returned ${res.status}:\n${bodyText}`);
    process.exit(1);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
