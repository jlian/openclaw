/**
 * Local message history store for MSTeams conversations.
 *
 * Captures inbound and outbound messages as they flow through the Bot
 * Framework webhook, enabling message history reads without Graph API.
 *
 * Storage: one JSON file per conversation, auto-purged by age and capacity.
 */

/** Default retention: 12 months. */
export const DEFAULT_HISTORY_TTL_MS = 365 * 24 * 60 * 60 * 1000;

/** Default max messages per conversation. */
export const DEFAULT_MAX_MESSAGES = 200_000;

export type StoredMessage = {
  id: string;
  from: string;
  body: string;
  createdDateTime: string;
  messageType: string;
  attachments?: Array<{
    contentType?: string;
    contentUrl?: string;
    name?: string;
    thumbnailUrl?: string;
  }>;
};

export type MessageHistoryStoreData = {
  version: 1;
  conversationId: string;
  messages: StoredMessage[];
};

export type ReadOpts = { limit?: number; cursor?: string; before?: string; after?: string };
export type ReadResult = { messages: StoredMessage[]; nextCursor?: string };

export type MessageHistoryStore = {
  /** Append a message (deduped by id, auto-prunes old entries). */
  append(conversationId: string, message: StoredMessage): Promise<void>;
  /** Read recent messages (newest first). Supports limit, cursor, before/after. */
  read(conversationId: string, opts?: ReadOpts): Promise<ReadResult>;
  /** Read from the most recently modified store file (fallback when conversation ID is unknown). */
  readDefault(opts?: ReadOpts): Promise<ReadResult>;
};
