/**
 * Minimal Work IQ MCP client (Streamable HTTP transport) for performing real
 * actions — the "Work IQ Tool API". Work IQ's Chat REST API only returns text;
 * actions such as sending mail are exposed through the Work IQ MCP server via
 * generic tools (`do_action`, `create_entity`, `fetch`, ...).
 *
 * This client speaks JSON-RPC 2.0 over a single HTTP endpoint and reuses the
 * same delegated Work IQ token (scope WorkIQAgent.Ask) as the Chat client.
 *
 * Docs:
 *  - https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/overview
 *  - https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/mcp/tool-reference
 */

const DEFAULT_ENDPOINT = "https://workiq.svc.cloud.microsoft/mcp";
const PROTOCOL_VERSION = "2025-06-18";

export interface McpToolResult {
  content?: Array<{ type: string; text?: string }>;
  structuredContent?: { data?: unknown; statusCode?: number } & Record<string, unknown>;
  isError?: boolean;
}

export interface SendMailOptions {
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bcc?: string[];
  /** Send the body as HTML instead of plain text. */
  html?: boolean;
  /** Save a copy in Sent Items (default true). */
  saveToSentItems?: boolean;
}

function recipients(addresses?: string[]) {
  return (addresses ?? [])
    .map((address) => address.trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));
}

export class WorkIqMcpClient {
  private sessionId: string | null = null;
  private negotiatedVersion = PROTOCOL_VERSION;
  private nextId = 1;
  private initialized = false;

  constructor(
    private readonly getToken: () => Promise<string>,
    private readonly endpoint = process.env.WORK_IQ_MCP_ENDPOINT ?? DEFAULT_ENDPOINT
  ) {}

  /** Send a single JSON-RPC message and return its `result` (null for notifications). */
  private async send(
    method: string,
    params?: unknown,
    notification = false
  ): Promise<any> {
    const token = await this.getToken();
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      "MCP-Protocol-Version": this.negotiatedVersion,
    };
    if (this.sessionId) headers["Mcp-Session-Id"] = this.sessionId;

    const message: Record<string, unknown> = { jsonrpc: "2.0", method };
    if (!notification) message.id = this.nextId++;
    if (params !== undefined) message.params = params;

    const res = await fetch(this.endpoint, {
      method: "POST",
      headers,
      body: JSON.stringify(message),
    });

    const sid = res.headers.get("Mcp-Session-Id");
    if (sid) this.sessionId = sid;

    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      throw new Error(
        `Work IQ MCP ${method} failed: ${res.status} ${res.statusText}\n${detail}`
      );
    }

    if (notification || res.status === 202) return null;

    const rpc = await this.readJsonRpc(res);
    if (rpc?.error) {
      throw new Error(
        `Work IQ MCP ${method} error ${rpc.error.code}: ${rpc.error.message}`
      );
    }
    return rpc?.result;
  }

  /** Read a JSON-RPC response from either a JSON body or an SSE stream. */
  private async readJsonRpc(res: Response): Promise<any> {
    const contentType = res.headers.get("content-type") ?? "";
    if (!contentType.includes("text/event-stream")) {
      const text = await res.text();
      return text ? JSON.parse(text) : null;
    }

    const reader = res.body!.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let result: any = null;

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      let idx: number;
      while ((idx = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, idx).trimEnd();
        buffer = buffer.slice(idx + 1);
        if (!line.startsWith("data:")) continue;
        const data = line.slice(5).trim();
        if (!data) continue;
        try {
          const parsed = JSON.parse(data);
          // The JSON-RPC response carries an "id"; skip keep-alives/notifications.
          if (parsed && parsed.jsonrpc && "id" in parsed) result = parsed;
        } catch {
          /* ignore non-JSON keep-alive lines */
        }
      }
      if (result) break;
    }
    try {
      await reader.cancel();
    } catch {
      /* already closed */
    }
    return result;
  }

  /** MCP handshake: initialize + notifications/initialized. Runs once. */
  async initialize(): Promise<void> {
    if (this.initialized) return;
    const result = await this.send("initialize", {
      protocolVersion: PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: "workiq-lab", version: "1.0.0" },
    });
    if (result?.protocolVersion) this.negotiatedVersion = result.protocolVersion;
    // Best-effort: the Work IQ MCP server sometimes 500s on this notification,
    // yet the session is still usable, so a failure here must not be fatal.
    try {
      await this.send("notifications/initialized", undefined, true);
    } catch {
      /* non-fatal */
    }
    this.initialized = true;
  }

  /** List the tools exposed by the Work IQ MCP server. */
  async listTools(): Promise<{ tools: Array<{ name: string; description?: string }> }> {
    await this.initialize();
    return this.send("tools/list", {});
  }

  /** Invoke a Work IQ MCP tool by name. */
  async callTool(
    name: string,
    args: Record<string, unknown>
  ): Promise<McpToolResult> {
    await this.initialize();
    return (await this.send("tools/call", {
      name,
      arguments: args,
    })) as McpToolResult;
  }

  /** Send an email via the `do_action` tool (POST /me/sendMail). */
  async sendMail(opts: SendMailOptions): Promise<McpToolResult> {
    const jsonBody = JSON.stringify({
      message: {
        subject: opts.subject,
        body: {
          contentType: opts.html ? "HTML" : "Text",
          content: opts.body,
        },
        toRecipients: recipients(opts.to),
        ccRecipients: recipients(opts.cc),
        bccRecipients: recipients(opts.bcc),
      },
      saveToSentItems: opts.saveToSentItems ?? true,
    });
    return this.callTool("do_action", { actionUrl: "/me/sendMail", jsonBody });
  }
}
