import {
  resolveChannelMediaMaxBytes,
  type OpenClawConfig,
  type PluginRuntime,
} from "openclaw/plugin-sdk";
import type { MSTeamsAccessTokenProvider } from "./attachments/types.js";
import type {
  MSTeamsConversationStore,
  StoredConversationReference,
} from "./conversation-store.js";
import type { MSTeamsAdapter } from "./messenger.js";
import { createMSTeamsConversationStoreFs } from "./conversation-store-fs.js";
import { resolveGraphChatId, type ReadTarget } from "./graph-read.js";
import { getMSTeamsRuntime } from "./runtime.js";
import { createMSTeamsAdapter, loadMSTeamsSdkWithAuth } from "./sdk.js";
import { resolveMSTeamsCredentials } from "./token.js";

export type MSTeamsConversationType = "personal" | "groupChat" | "channel";

export type MSTeamsReadContext = {
  tokenProvider: MSTeamsAccessTokenProvider;
  target: ReadTarget;
  conversationType: MSTeamsConversationType;
};

export type MSTeamsProactiveContext = {
  appId: string;
  conversationId: string;
  ref: StoredConversationReference;
  adapter: MSTeamsAdapter;
  log: ReturnType<PluginRuntime["logging"]["getChildLogger"]>;
  /** The type of conversation: personal (1:1), groupChat, or channel */
  conversationType: MSTeamsConversationType;
  /** Token provider for Graph API / OneDrive operations */
  tokenProvider: MSTeamsAccessTokenProvider;
  /** SharePoint site ID for file uploads in group chats/channels */
  sharePointSiteId?: string;
  /** Resolved media max bytes from config (default: 100MB) */
  mediaMaxBytes?: number;
};

/**
 * Parse the target value into a conversation reference lookup key.
 * Supported formats:
 * - conversation:19:abc@thread.tacv2 → lookup by conversation ID
 * - user:aad-object-id → lookup by user AAD object ID
 * - 19:abc@thread.tacv2 → direct conversation ID
 */
function parseRecipient(to: string): {
  type: "conversation" | "user";
  id: string;
} {
  const trimmed = to.trim();
  const finalize = (type: "conversation" | "user", id: string) => {
    const normalized = id.trim();
    if (!normalized) {
      throw new Error(`Invalid target value: missing ${type} id`);
    }
    return { type, id: normalized };
  };
  if (trimmed.startsWith("conversation:")) {
    return finalize("conversation", trimmed.slice("conversation:".length));
  }
  if (trimmed.startsWith("user:")) {
    return finalize("user", trimmed.slice("user:".length));
  }
  // Assume it's a conversation ID if it looks like one
  if (trimmed.startsWith("19:") || trimmed.includes("@thread")) {
    return finalize("conversation", trimmed);
  }
  // Otherwise treat as user ID
  return finalize("user", trimmed);
}

/**
 * Find a stored conversation reference for the given recipient.
 */
async function findConversationReference(recipient: {
  type: "conversation" | "user";
  id: string;
  store: MSTeamsConversationStore;
}): Promise<{
  conversationId: string;
  ref: StoredConversationReference;
} | null> {
  if (recipient.type === "conversation") {
    const ref = await recipient.store.get(recipient.id);
    if (ref) {
      return { conversationId: recipient.id, ref };
    }
    return null;
  }

  const found = await recipient.store.findByUserId(recipient.id);
  if (!found) {
    return null;
  }
  return { conversationId: found.conversationId, ref: found.reference };
}

/**
 * Create a lightweight token provider that acquires Graph API tokens via
 * OAuth2 client credentials for a specific tenant.
 *
 * Used when the user's tenant differs from the bot's registration tenant
 * (cross-tenant / multi-tenant app scenario). The app registration MUST
 * be configured as multi-tenant in Azure AD and have admin consent on the
 * target tenant for this to work.
 */
async function createTenantGraphTokenProvider(params: {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  fetchFn?: typeof fetch;
}): Promise<MSTeamsAccessTokenProvider> {
  const { clientId, clientSecret, tenantId, fetchFn = fetch } = params;
  let cachedToken: { token: string; expiresAt: number } | null = null;

  return {
    getAccessToken: async (_scope: string): Promise<string> => {
      // Return cached token if still valid (with 5 min buffer)
      if (cachedToken && cachedToken.expiresAt > Date.now() + 300_000) {
        return cachedToken.token;
      }

      const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      });

      const res = await fetchFn(tokenUrl, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      });

      if (!res.ok) {
        const text = await res.text().catch(() => "");
        throw new Error(
          `Failed to acquire Graph token for tenant ${tenantId} (${res.status}). ` +
            `Ensure the app registration is configured as multi-tenant and has ` +
            `admin consent on the target tenant. ${text.slice(0, 200)}`,
        );
      }

      const data = (await res.json()) as { access_token?: string; expires_in?: number };
      if (!data.access_token) {
        throw new Error("Token response missing access_token");
      }

      cachedToken = {
        token: data.access_token,
        expiresAt: Date.now() + (data.expires_in ?? 3600) * 1000,
      };

      return data.access_token;
    },
  };
}

/**
 * Lightweight resolver for Graph read operations.
 * Resolves credentials, token provider, and conversation target
 * without creating a full adapter or proactive context.
 *
 * When the user's tenant (from the conversation reference) differs from the
 * bot's registration tenant, a cross-tenant token provider is created so
 * Graph API calls target the correct tenant where Teams is provisioned.
 */
export async function resolveGraphReadContext(params: {
  cfg: OpenClawConfig;
  to: string;
}): Promise<MSTeamsReadContext> {
  const msteamsCfg = params.cfg.channels?.msteams;

  if (!msteamsCfg?.enabled) {
    throw new Error("msteams provider is not enabled");
  }

  const creds = resolveMSTeamsCredentials(msteamsCfg);
  if (!creds) {
    throw new Error("msteams credentials not configured");
  }

  const store = createMSTeamsConversationStoreFs();
  const recipient = parseRecipient(params.to);
  const found = await findConversationReference({ ...recipient, store });

  if (!found) {
    throw new Error(
      `No conversation reference found for ${recipient.type}:${recipient.id}. ` +
        `The bot must receive a message from this conversation before it can read history.`,
    );
  }

  const { ref } = found;
  const conversationType = resolveConversationType(ref);

  // Determine which tenant to use for Graph API calls.
  // If the user's tenant differs from the bot's registration tenant,
  // create a cross-tenant token provider targeting the user's tenant.
  const userTenantId = ref.conversation?.tenantId;
  const botTenantId = creds.tenantId;
  const isCrossTenant = userTenantId && userTenantId !== botTenantId;

  let tokenProvider: MSTeamsAccessTokenProvider;
  if (isCrossTenant) {
    tokenProvider = await createTenantGraphTokenProvider({
      clientId: creds.appId,
      clientSecret: creds.appPassword,
      tenantId: userTenantId,
    });
  } else {
    // Same tenant — use the standard SDK token provider
    const { sdk, authConfig } = await loadMSTeamsSdkWithAuth(creds);
    tokenProvider = new sdk.MsalTokenProvider(authConfig) as MSTeamsAccessTokenProvider;
  }

  const target = await resolveReadTarget(ref, conversationType, tokenProvider, creds.appId);

  return { tokenProvider, target, conversationType };
}

/** Determine conversation type from stored reference. */
function resolveConversationType(ref: StoredConversationReference): MSTeamsConversationType {
  const stored = ref.conversation?.conversationType?.toLowerCase() ?? "";
  if (stored === "personal") {
    return "personal";
  }
  if (stored === "channel") {
    return "channel";
  }
  return "groupChat";
}

/**
 * Check whether a conversation ID looks like a Graph API chat ID.
 * Graph chat IDs follow the `19:...@unq.gbl.spaces` or `19:...@thread.v2` pattern.
 * Bot Framework IDs typically start with `a:` or similar prefixes.
 */
function isGraphChatId(id: string): boolean {
  return id.startsWith("19:") && id.includes("@");
}

/** Build a ReadTarget from a stored conversation reference. */
async function resolveReadTarget(
  ref: StoredConversationReference,
  conversationType: MSTeamsConversationType,
  tokenProvider: MSTeamsAccessTokenProvider,
  botAppId?: string,
): Promise<ReadTarget> {
  const conversationId = ref.conversation?.id;
  if (!conversationId) {
    throw new Error("Stored conversation reference has no conversation ID");
  }

  if (conversationType === "channel") {
    const teamId = ref.teamId;
    if (!teamId) {
      throw new Error(
        "Cannot read channel messages: no teamId in stored conversation reference. " +
          "The bot needs to receive a message in this channel first.",
      );
    }
    return { kind: "channel", teamId, channelId: conversationId };
  }

  // personal or groupChat → use /chats/{id}/messages
  // Bot Framework conversation IDs (a:...) are NOT valid Graph chat IDs.
  // We need to resolve the real Graph chat ID via the Graph API.
  if (isGraphChatId(conversationId)) {
    return { kind: "chat", chatId: conversationId };
  }

  // Check if we have a cached Graph chat ID from inbound channelData
  if (ref.graphChatId && isGraphChatId(ref.graphChatId)) {
    return { kind: "chat", chatId: ref.graphChatId };
  }

  // Last resort: resolve Bot Framework ID → Graph chat ID via Graph API
  // (requires Chat.Read.All and Teams provisioned on the tenant)
  const userAadObjectId = ref.user?.aadObjectId;
  if (!userAadObjectId) {
    throw new Error(
      "Cannot resolve Graph chat ID: no user aadObjectId in stored conversation reference. " +
        "Send a message in Teams first so the bot can capture the Graph chat ID from channelData.",
    );
  }

  const botAadObjectId = ref.agent?.aadObjectId ?? ref.bot?.id;
  const graphChatId = await resolveGraphChatId({
    tokenProvider,
    userAadObjectId,
    botAppId,
    botAadObjectId,
  });

  return { kind: "chat", chatId: graphChatId };
}

export async function resolveMSTeamsSendContext(params: {
  cfg: OpenClawConfig;
  to: string;
}): Promise<MSTeamsProactiveContext> {
  const msteamsCfg = params.cfg.channels?.msteams;

  if (!msteamsCfg?.enabled) {
    throw new Error("msteams provider is not enabled");
  }

  const creds = resolveMSTeamsCredentials(msteamsCfg);
  if (!creds) {
    throw new Error("msteams credentials not configured");
  }

  const store = createMSTeamsConversationStoreFs();

  // Parse recipient and find conversation reference
  const recipient = parseRecipient(params.to);
  const found = await findConversationReference({ ...recipient, store });

  if (!found) {
    throw new Error(
      `No conversation reference found for ${recipient.type}:${recipient.id}. ` +
        `The bot must receive a message from this conversation before it can send proactively.`,
    );
  }

  const { conversationId, ref } = found;
  const core = getMSTeamsRuntime();
  const log = core.logging.getChildLogger({ name: "msteams:send" });

  const { sdk, authConfig } = await loadMSTeamsSdkWithAuth(creds);
  const adapter = createMSTeamsAdapter(authConfig, sdk);

  // Create token provider for Graph API / OneDrive operations
  const tokenProvider = new sdk.MsalTokenProvider(authConfig) as MSTeamsAccessTokenProvider;

  // Determine conversation type from stored reference
  const conversationType = resolveConversationType(ref);

  // Get SharePoint site ID from config (required for file uploads in group chats/channels)
  const sharePointSiteId = msteamsCfg.sharePointSiteId;

  // Resolve media max bytes from config
  const mediaMaxBytes = resolveChannelMediaMaxBytes({
    cfg: params.cfg,
    resolveChannelLimitMb: ({ cfg }) => cfg.channels?.msteams?.mediaMaxMb,
  });

  return {
    appId: creds.appId,
    conversationId,
    ref,
    adapter: adapter as unknown as MSTeamsAdapter,
    log,
    conversationType,
    tokenProvider,
    sharePointSiteId,
    mediaMaxBytes,
  };
}
