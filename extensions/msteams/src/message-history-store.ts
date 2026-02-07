/**
 * Local message history store for MSTeams conversations.
 *
 * Captures inbound and outbound messages as they flow through
 * the Bot Framework webhook, enabling readMessages without Graph API
 * permissions or cross-tenant consent.
 *
 * Storage: one JSON file per conversation, auto-purged by age and capacity.
 */

import type { NormalizedTeamsMessage } from "./graph-read.js";

/** Default retention: 12 months. */
export const DEFAULT_HISTORY_TTL_MS = 365 * 24 * 60 * 60 * 1000;

/** Default max messages per conversation. At ~300 bytes/msg, 500 ≈ 150 KB. */
export const DEFAULT_MAX_MESSAGES = 500;

export type MessageHistoryStoreData = {
  version: 1;
  conversationId: string;
  messages: NormalizedTeamsMessage[];
};

export type MessageHistoryStore = {
  /** Append a message to the conversation's history. Auto-prunes old entries. */
  append(conversationId: string, message: NormalizedTeamsMessage): Promise<void>;
  /** Read recent messages (newest first). Supports limit + offset-based cursor. */
  read(
    conversationId: string,
    opts?: { limit?: number; cursor?: string },
  ): Promise<{ messages: NormalizedTeamsMessage[]; nextCursor?: string }>;
};
