/**
 * Minimal typed client for the Microsoft 365 Work IQ Chat REST API.
 * Docs: https://learn.microsoft.com/microsoft-365/copilot/extensibility/work-iq/rest/overview
 */

const PROD_BASE = "https://workiq.svc.cloud.microsoft/rest";
const BETA_BASE = "https://workiq.svc.cloud.microsoft/rest/beta";

export interface CopilotConversation {
  id: string;
  createdDateTime: string;
  displayName: string;
  status?: string;
  state?: string;
  turnCount: number;
  messages?: CopilotResponseMessage[];
}

export interface CopilotResponseMessage {
  "@odata.type"?: string;
  id: string;
  text: string;
  createdDateTime: string;
  adaptiveCards?: unknown[];
  attributions?: Attribution[];
}

export interface Attribution {
  attributionType?: string;
  providerDisplayName?: string;
  attributionSource?: string;
  seeMoreWebUrl?: string;
}

export interface ChatOptions {
  /** IANA time zone, e.g. "America/New_York". */
  timeZone: string;
  /** OneDrive / SharePoint file URLs to use as grounding context. */
  fileUris?: string[];
  /** Set false to disable web search grounding for this single turn. */
  webSearch?: boolean;
  /** Extra free-text grounding snippets. */
  additionalContext?: string[];
}

export class WorkIqClient {
  private readonly base: string;

  constructor(
    private readonly getToken: () => Promise<string>,
    useBeta = false
  ) {
    this.base = useBeta ? BETA_BASE : PROD_BASE;
  }

  private async request<T>(
    method: string,
    urlPath: string,
    body?: unknown
  ): Promise<T> {
    const token = await this.getToken();
    const res = await fetch(`${this.base}${urlPath}`, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });

    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      throw new Error(
        `Work IQ API ${method} ${urlPath} failed: ${res.status} ${res.statusText}\n${detail}`
      );
    }
    return (await res.json()) as T;
  }

  /** Create a new conversation and return its id. */
  async createConversation(): Promise<CopilotConversation> {
    return this.request<CopilotConversation>("POST", "/conversations", {});
  }

  /** Build the chat request body from a prompt + options. */
  private buildBody(text: string, opts: ChatOptions) {
    const body: Record<string, unknown> = {
      message: { text },
      locationHint: { timeZone: opts.timeZone },
    };

    if (opts.fileUris?.length || opts.webSearch === false) {
      const contextualResources: Record<string, unknown> = {};
      if (opts.fileUris?.length) {
        contextualResources.files = opts.fileUris.map((uri) => ({ uri }));
      }
      if (opts.webSearch === false) {
        contextualResources.webSearchOptions = { isEnabled: false };
      }
      body.contextualResources = contextualResources;
    }

    if (opts.additionalContext?.length) {
      body.additionalContext = opts.additionalContext.map((text) => ({ text }));
    }
    return body;
  }

  /** Send a message on the synchronous endpoint and return the full conversation. */
  async chat(
    conversationId: string,
    text: string,
    opts: ChatOptions
  ): Promise<CopilotConversation> {
    return this.request<CopilotConversation>(
      "POST",
      `/conversations/${conversationId}/chat`,
      this.buildBody(text, opts)
    );
  }

  /**
   * Send a message on the streamed endpoint. Yields raw Server-Sent-Event data
   * lines as they arrive so callers can render partial output.
   */
  async *chatOverStream(
    conversationId: string,
    text: string,
    opts: ChatOptions
  ): AsyncGenerator<string> {
    const token = await this.getToken();
    const res = await fetch(
      `${this.base}/conversations/${conversationId}/chatOverStream`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          Accept: "text/event-stream",
        },
        body: JSON.stringify(this.buildBody(text, opts)),
      }
    );

    if (!res.ok || !res.body) {
      const detail = await res.text().catch(() => "");
      throw new Error(
        `Work IQ stream failed: ${res.status} ${res.statusText}\n${detail}`
      );
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      let idx: number;
      while ((idx = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, idx).trimEnd();
        buffer = buffer.slice(idx + 1);
        if (line.startsWith("data:")) {
          yield line.slice(5).trim();
        }
      }
    }
    if (buffer.startsWith("data:")) yield buffer.slice(5).trim();
  }
}

/** Extract just the assistant's reply text from a conversation response. */
export function lastAssistantText(conv: CopilotConversation): string {
  const msgs = conv.messages ?? [];
  // The final message is the assistant reply; the prior one echoes the prompt.
  for (let i = msgs.length - 1; i >= 0; i--) {
    if (msgs[i].text) return msgs[i].text;
  }
  return "(no response text)";
}
