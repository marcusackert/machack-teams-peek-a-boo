const GRAPH = "https://graph.microsoft.com/v1.0";

export interface Team {
  id: string;
  displayName: string;
  tenantId?: string;
}

export interface Channel {
  id: string;
  displayName: string;
  membershipType?: string;
  webUrl?: string;
}

export interface Identity {
  id?: string;
  displayName?: string;
  userIdentityType?: string;
  tenantId?: string;
}

export interface ChatMessage {
  id: string;
  replyToId?: string | null;
  messageType: string;
  createdDateTime: string;
  lastModifiedDateTime?: string;
  lastEditedDateTime?: string | null;
  deletedDateTime?: string | null;
  subject?: string | null;
  importance?: string;
  webUrl?: string;
  from?: { user?: Identity | null; application?: Identity | null } | null;
  body: { contentType: "html" | "text"; content: string };
  attachments?: { id: string; contentType: string; contentUrl?: string; name?: string | null; content?: string | null }[];
  reactions?: { reactionType: string; displayName?: string; user: { user?: Identity } }[];
  replies?: ChatMessage[];
  "replies@odata.nextLink"?: string;
}

export interface Page<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export class GraphError extends Error {
  constructor(
    public status: number,
    public code: string,
    message: string,
  ) {
    super(message);
  }
  /** Missing permission or consent rather than a transient failure. */
  get isAuthz(): boolean {
    return this.status === 401 || this.status === 403;
  }
}

export type TokenFn = () => Promise<string>;

export async function graphGet<T>(token: TokenFn, pathOrUrl: string): Promise<T> {
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : GRAPH + pathOrUrl;
  if (!url.startsWith(GRAPH)) throw new Error("Refusing to send token to non-Graph URL");
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${await token()}` } });
    if (res.status === 429 && attempt < 3) {
      const wait = Number(res.headers.get("Retry-After") ?? "5");
      await new Promise((r) => setTimeout(r, Math.min(wait, 30) * 1000));
      continue;
    }
    if (!res.ok) {
      let code = String(res.status);
      let message = res.statusText;
      try {
        const err = (await res.json()).error;
        code = err?.code ?? code;
        message = err?.message ?? message;
      } catch {
        /* non-JSON error */
      }
      throw new GraphError(res.status, code, message);
    }
    return (await res.json()) as T;
  }
}

async function getAll<T>(token: TokenFn, path: string): Promise<T[]> {
  const out: T[] = [];
  let next: string | undefined = path;
  while (next) {
    const page: Page<T> = await graphGet<Page<T>>(token, next);
    out.push(...page.value);
    next = page["@odata.nextLink"];
  }
  return out;
}

export const listJoinedTeams = (token: TokenFn) =>
  getAll<Team>(token, "/me/joinedTeams?$select=id,displayName,tenantId");

export const listChannels = (token: TokenFn, teamId: string) =>
  getAll<Channel>(token, `/teams/${teamId}/channels?$select=id,displayName,membershipType,webUrl`);

export const getChannel = (token: TokenFn, teamId: string, channelId: string) =>
  graphGet<Channel>(token, `/teams/${teamId}/channels/${encodeURIComponent(channelId)}`);

/** Newest root posts first page (or a nextLink for older ones), replies expanded. */
export const listMessages = (token: TokenFn, teamId: string, channelId: string, nextLink?: string) =>
  graphGet<Page<ChatMessage>>(
    token,
    nextLink ?? `/teams/${teamId}/channels/${encodeURIComponent(channelId)}/messages?$top=30&$expand=replies`,
  );

export const listReplies = (token: TokenFn, teamId: string, channelId: string, messageId: string) =>
  getAll<ChatMessage>(token, `/teams/${teamId}/channels/${encodeURIComponent(channelId)}/messages/${messageId}/replies?$top=50`);

/** Downloads Graph-hosted content (inline images) as an object URL. */
export async function fetchBlobUrl(token: TokenFn, url: string): Promise<string> {
  if (!url.startsWith(GRAPH)) throw new Error("Not a Graph URL");
  const res = await fetch(url, { headers: { Authorization: `Bearer ${await token()}` } });
  if (!res.ok) throw new GraphError(res.status, String(res.status), res.statusText);
  return URL.createObjectURL(await res.blob());
}
