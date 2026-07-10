import "dotenv/config";
import * as readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { getAccessToken } from "./auth.ts";
import { WorkIqClient, lastAssistantText } from "./workiq.ts";

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value || value.startsWith("your-")) {
    console.error(
      `Missing env var ${name}. Copy .env.example to .env and fill it in.`
    );
    process.exit(1);
  }
  return value;
}

async function main() {
  const tenantId = requireEnv("TENANT_ID");
  const clientId = requireEnv("CLIENT_ID");
  const useBeta = (process.env.USE_BETA ?? "false").toLowerCase() === "true";
  const timeZone = process.env.TIME_ZONE ?? "America/New_York";
  const stream = process.argv.includes("--stream");

  console.log("Signing in to Microsoft Entra (Work IQ delegated scope)...");
  const client = new WorkIqClient(
    () => getAccessToken(tenantId, clientId),
    useBeta
  );

  console.log(`Creating a conversation (${useBeta ? "beta" : "prod"})...`);
  const conversation = await client.createConversation();
  console.log(`Conversation id: ${conversation.id}\n`);

  const rl = readline.createInterface({ input, output });
  console.log(
    "Type a prompt and press Enter. Commands: /stream, /sync, /exit.\n" +
      `Streaming mode: ${stream ? "on" : "off"}\n`
  );

  let streaming = stream;
  try {
    while (true) {
      const prompt = (await rl.question("you > ")).trim();
      if (!prompt) continue;
      if (prompt === "/exit" || prompt === "/quit") break;
      if (prompt === "/stream") {
        streaming = true;
        console.log("(streaming on)\n");
        continue;
      }
      if (prompt === "/sync") {
        streaming = false;
        console.log("(streaming off)\n");
        continue;
      }

      try {
        if (streaming) {
          process.stdout.write("copilot > ");
          for await (const data of client.chatOverStream(
            conversation.id,
            prompt,
            { timeZone }
          )) {
            process.stdout.write(data);
          }
          process.stdout.write("\n\n");
        } else {
          const result = await client.chat(conversation.id, prompt, {
            timeZone,
          });
          console.log(`copilot > ${lastAssistantText(result)}\n`);
        }
      } catch (err) {
        console.error(`\n[error] ${(err as Error).message}\n`);
      }
    }
  } finally {
    rl.close();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
