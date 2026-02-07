/**
 * Read messages from Microsoft Teams conversations via Graph API.
 *
 * Supports both:
 * - Chat (1:1, group): GET /chats/{chatId}/messages
 * - Channel:           GET /teams/{teamId}/channels/{channelId}/messages
 *
 * Required Graph permissions:
 * - ChannelMessage.Read.All (channels)
 * - Chat.Read.All (DMs / group chats)
 */

import type { MSTeamsAccessTokenProvider } from "./attachments/types.js";
import { GRAPH_ROOT } from "./attachments/shared.js";

// ── Graph API response types ────────────────────────────────────────────

/** Subset of the Graph ChatMessage resource we care about. */
export type GraphChatMessage = {
  id?: string;
  createdDateTime?: string;
  messageType?: string;
  from?: {
    user?: { id?: string; displayName?: string } | null;
    application?: { id?: string; displayName?: string } | null;
  } | null;
  body?: { contentType?: string; content?: string } | null;
  subject?: string | null;
  attachments?: Array<{ id?: string; name?: string; contentType?: string }>;
};

/** Normalized message returned to the agent. */
export type NormalizedTeamsMessageAttachment = {
  contentType?: string;
  contentUrl?: string;
  name?: string;
  thumbnailUrl?: string;
};

export type NormalizedTeamsMessage = {
  id: string;
  from: string;
  body: string;
  createdDateTime: string;
  messageType: string;
  attachments?: NormalizedTeamsMessageAttachment[];
};

export type ReadMessagesResult = {
  messages: NormalizedTeamsMessage[];
  /** Opaque cursor for fetching the next page (from @odata.nextLink). */
  nextCursor?: string;
};

// ── Target specification ────────────────────────────────────────────────

export type ChatTarget = { kind: "chat"; chatId: string };
export type ChannelTarget = { kind: "channel"; teamId: string; channelId: string };
export type ReadTarget = ChatTarget | ChannelTarget;

// ── Bot Framework → Graph chat ID resolution ─────────────────────────────

/**
 * Bot Framework conversation IDs (e.g. `a:1hgI...`) are NOT valid Graph API
 * chat IDs. Graph expects `19:...@unq.gbl.spaces` for 1:1 / group chats.
 *
 * Resolution strategy (in order):
 * 1. Query the user's installed Teams apps to find the bot's installation,
 *    then get the associated chat via the installedApps/chat endpoint.
 *    Requires `TeamsAppInstallation.ReadForUser.All` (application).
 * 2. Fall back to listing the user's chats and matching by bot membership.
 *    Requires `Chat.Read.All` (application).
 */
export async function resolveGraphChatId(params: {
  tokenProvider: MSTeamsAccessTokenProvider;
  userAadObjectId: string;
  /** Bot app registration ID (client ID), used for the installedApps lookup. */
  botAppId?: string;
  /** Bot AAD object ID, used for the /users/{id}/chats member-matching fallback. */
  botAadObjectId?: string;
  fetchFn?: typeof fetch;
}): Promise<string> {
  const { tokenProvider, userAadObjectId, botAppId, botAadObjectId, fetchFn = fetch } = params;
  const accessToken = await tokenProvider.getAccessToken("https://graph.microsoft.com");

  // Strategy 1: installedApps → chat (most reliable, purpose-built API)
  if (botAppId) {
    const chatId = await resolveViaInstalledApps({
      accessToken,
      userAadObjectId,
      botAppId,
      fetchFn,
    });
    if (chatId) {
      return chatId;
    }
  }

  // Strategy 2: list user's chats and match by bot membership
  const chatId = await resolveViaUserChats({
    accessToken,
    userAadObjectId,
    botAadObjectId,
    fetchFn,
  });
  if (chatId) {
    return chatId;
  }

  throw new Error(
    `Could not resolve Graph chat ID for user ${userAadObjectId}. ` +
      `Ensure TeamsAppInstallation.ReadForUser.All or Chat.Read.All permission ` +
      `is granted with admin consent.`,
  );
}

/**
 * Resolve Graph chat ID via the installedApps endpoint.
 * GET /users/{userId}/teamwork/installedApps?$filter=...&$expand=teamsApp
 * GET /users/{userId}/teamwork/installedApps/{installId}/chat
 */
async function resolveViaInstalledApps(params: {
  accessToken: string;
  userAadObjectId: string;
  botAppId: string;
  fetchFn: typeof fetch;
}): Promise<string | null> {
  const { accessToken, userAadObjectId, botAppId, fetchFn } = params;

  // Step 1: find the app installation for this bot
  const listUrl =
    `${GRAPH_ROOT}/users/${encodeURIComponent(userAadObjectId)}/teamwork/installedApps` +
    `?$expand=teamsApp` +
    `&$filter=teamsApp/externalId eq '${botAppId}'`;

  const listRes = await fetchFn(listUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!listRes.ok) {
    // Log the error for debugging but don't throw — fall through to next strategy
    const errBody = await listRes.text().catch(() => "");
    console.warn(
      `[msteams:read] installedApps lookup failed (${listRes.status}): ${errBody.slice(0, 200)}`,
    );
    return null;
  }

  type InstalledApp = { id?: string; teamsApp?: { externalId?: string } };
  type ListResponse = { value?: InstalledApp[] };

  const listData = (await listRes.json()) as ListResponse;
  const apps = listData.value ?? [];

  if (apps.length === 0 || !apps[0]?.id) {
    return null;
  }

  const installId = apps[0].id;

  // Step 2: get the chat associated with this installation
  const chatUrl =
    `${GRAPH_ROOT}/users/${encodeURIComponent(userAadObjectId)}` +
    `/teamwork/installedApps/${encodeURIComponent(installId)}/chat`;

  const chatRes = await fetchFn(chatUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!chatRes.ok) {
    const errBody = await chatRes.text().catch(() => "");
    console.warn(
      `[msteams:read] installedApps/chat lookup failed (${chatRes.status}): ${errBody.slice(0, 200)}`,
    );
    return null;
  }

  type ChatResponse = { id?: string };
  const chatData = (await chatRes.json()) as ChatResponse;

  return chatData.id ?? null;
}

/**
 * Resolve Graph chat ID by listing the user's chats and matching bot membership.
 * GET /users/{userId}/chats?$filter=chatType eq 'oneOnOne'&$expand=members
 */
async function resolveViaUserChats(params: {
  accessToken: string;
  userAadObjectId: string;
  botAadObjectId?: string;
  fetchFn: typeof fetch;
}): Promise<string | null> {
  const { accessToken, userAadObjectId, botAadObjectId, fetchFn } = params;

  const url =
    `${GRAPH_ROOT}/users/${encodeURIComponent(userAadObjectId)}/chats` +
    `?$filter=chatType eq 'oneOnOne'` +
    `&$expand=members` +
    `&$select=id,chatType` +
    `&$top=50`;

  const res = await fetchFn(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ConsistencyLevel: "eventual",
    },
  });

  if (!res.ok) {
    const errBody = await res.text().catch(() => "");
    console.warn(
      `[msteams:read] user chats lookup failed (${res.status}): ${errBody.slice(0, 200)}`,
    );
    return null;
  }

  type GraphChat = {
    id?: string;
    chatType?: string;
    members?: Array<{ userId?: string; displayName?: string }>;
  };
  type GraphListResponse = { value?: GraphChat[] };

  const data = (await res.json()) as GraphListResponse;
  const chats = data.value ?? [];

  // Match the chat where the bot is a member
  if (botAadObjectId) {
    const botId = botAadObjectId.toLowerCase();
    for (const chat of chats) {
      const members = chat.members ?? [];
      const hasBot = members.some((m) => m.userId?.toLowerCase() === botId);
      if (hasBot && chat.id) {
        return chat.id;
      }
    }
  }

  // Fallback: if only one 1:1 chat exists, use it
  if (chats.length === 1 && chats[0]?.id) {
    return chats[0].id;
  }

  return null;
}

// ── Core implementation ─────────────────────────────────────────────────

const MAX_LIMIT = 50;
const DEFAULT_LIMIT = 20;

/**
 * Build the Graph API URL for listing messages in a conversation.
 */
export function buildReadMessagesUrl(target: ReadTarget, limit: number, cursor?: string): string {
  if (cursor) {
    // @odata.nextLink is a full URL — use it directly
    return cursor;
  }

  const top = Math.min(Math.max(1, limit), MAX_LIMIT);

  if (target.kind === "channel") {
    return (
      `${GRAPH_ROOT}/teams/${encodeURIComponent(target.teamId)}` +
      `/channels/${encodeURIComponent(target.channelId)}` +
      `/messages?$top=${top}&$orderby=createdDateTime desc`
    );
  }

  return (
    `${GRAPH_ROOT}/chats/${encodeURIComponent(target.chatId)}` +
    `/messages?$top=${top}&$orderby=createdDateTime desc`
  );
}

/**
 * Strip HTML tags from a string (best-effort, for body.content with contentType "html").
 */
function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .trim();
}

/**
 * Normalize a Graph ChatMessage into a simple shape for the agent.
 */
function normalizeMessage(msg: GraphChatMessage): NormalizedTeamsMessage | null {
  const id = msg.id ?? "";
  if (!id) {
    return null;
  }

  // Determine sender display name
  const fromUser = msg.from?.user?.displayName;
  const fromApp = msg.from?.application?.displayName;
  const from = fromUser ?? fromApp ?? "unknown";

  // Normalize body — strip HTML if needed
  let body = "";
  if (msg.body?.content) {
    body =
      msg.body.contentType?.toLowerCase() === "html"
        ? stripHtml(msg.body.content)
        : msg.body.content.trim();
  }

  return {
    id,
    from,
    body,
    createdDateTime: msg.createdDateTime ?? "",
    messageType: msg.messageType ?? "message",
  };
}

/**
 * Fetch messages from a Teams conversation via the Graph API.
 */
export async function readMSTeamsMessages(params: {
  tokenProvider: MSTeamsAccessTokenProvider;
  target: ReadTarget;
  limit?: number;
  cursor?: string;
  fetchFn?: typeof fetch;
}): Promise<ReadMessagesResult> {
  const { tokenProvider, target, fetchFn = fetch } = params;
  const limit = params.limit ?? DEFAULT_LIMIT;

  const accessToken = await tokenProvider.getAccessToken("https://graph.microsoft.com");
  const url = buildReadMessagesUrl(target, limit, params.cursor);

  const res = await fetchFn(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!res.ok) {
    const status = res.status;
    if (status === 403) {
      throw new Error(
        "Graph API returned 403 Forbidden. Ensure the app has " +
          "ChannelMessage.Read.All (channels) or Chat.Read.All (chats) " +
          "permissions with admin consent. " +
          "See https://docs.openclaw.ai/channels/msteams#rsc-vs-graph-api",
      );
    }
    if (status === 404) {
      throw new Error(
        `Conversation not found (404). The chat or channel may not exist ` +
          `or the app may lack access.`,
      );
    }
    if (status === 429) {
      throw new Error("Graph API rate limit exceeded (429). Please try again later.");
    }
    const body = await res.text().catch(() => "");
    throw new Error(`Graph API error ${status}: ${body.slice(0, 300)}`);
  }

  type GraphListResponse = {
    value?: GraphChatMessage[];
    "@odata.nextLink"?: string;
  };

  let data: GraphListResponse;
  try {
    data = (await res.json()) as GraphListResponse;
  } catch {
    throw new Error("Failed to parse Graph API response");
  }

  const raw = Array.isArray(data.value) ? data.value : [];
  const messages = raw.map(normalizeMessage).filter((m): m is NormalizedTeamsMessage => m !== null);

  const nextCursor = data["@odata.nextLink"] ?? undefined;

  return { messages, nextCursor };
}
