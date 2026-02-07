/**
 * File-system backed message history store.
 *
 * One JSON file per conversation: `msteams-history-{hash}.json`
 * Uses the same lock/write utilities as the conversation store.
 */

import crypto from "node:crypto";
import type { NormalizedTeamsMessage } from "./graph-read.js";
import type { MessageHistoryStore, MessageHistoryStoreData } from "./message-history-store.js";
import { DEFAULT_HISTORY_TTL_MS, DEFAULT_MAX_MESSAGES } from "./message-history-store.js";
import { resolveMSTeamsStorePath } from "./storage.js";
import { readJsonFile, withFileLock, writeJsonFile } from "./store-fs.js";

const DEFAULT_LIMIT = 20;
const MAX_LIMIT = 50;

/** Deterministic short hash for filenames. */
function hashConversationId(id: string): string {
  return crypto.createHash("sha256").update(id).digest("hex").slice(0, 16);
}

function emptyStore(conversationId: string): MessageHistoryStoreData {
  return { version: 1, conversationId, messages: [] };
}

/** Parse a numeric cursor (offset into the messages array). */
function parseCursor(cursor: string | undefined): number {
  if (!cursor) {
    return 0;
  }
  const n = Number.parseInt(cursor, 10);
  return Number.isFinite(n) && n >= 0 ? n : 0;
}

export function createMessageHistoryStoreFs(params?: {
  env?: NodeJS.ProcessEnv;
  homedir?: () => string;
  stateDir?: string;
  ttlMs?: number;
  maxMessages?: number;
}): MessageHistoryStore {
  const ttlMs = params?.ttlMs ?? DEFAULT_HISTORY_TTL_MS;
  const maxMessages = params?.maxMessages ?? DEFAULT_MAX_MESSAGES;

  function resolveFilePath(conversationId: string): string {
    const hash = hashConversationId(conversationId);
    return resolveMSTeamsStorePath({
      filename: `msteams-history-${hash}.json`,
      env: params?.env,
      homedir: params?.homedir,
      stateDir: params?.stateDir,
    });
  }

  async function readStore(
    filePath: string,
    conversationId: string,
  ): Promise<MessageHistoryStoreData> {
    const { value } = await readJsonFile<MessageHistoryStoreData>(
      filePath,
      emptyStore(conversationId),
    );
    if (value.version !== 1 || !Array.isArray(value.messages)) {
      return emptyStore(conversationId);
    }
    return value;
  }

  /** Remove expired messages and cap to maxMessages (keep newest). */
  function prune(messages: NormalizedTeamsMessage[]): NormalizedTeamsMessage[] {
    const cutoff = Date.now() - ttlMs;
    const fresh = messages.filter((m) => {
      const ts = Date.parse(m.createdDateTime);
      return Number.isFinite(ts) && ts > cutoff;
    });
    // Keep newest messages if over capacity
    if (fresh.length > maxMessages) {
      return fresh.slice(fresh.length - maxMessages);
    }
    return fresh;
  }

  const append: MessageHistoryStore["append"] = async (conversationId, message) => {
    const filePath = resolveFilePath(conversationId);
    const empty = emptyStore(conversationId);

    await withFileLock(filePath, empty, async () => {
      const store = await readStore(filePath, conversationId);
      // Dedupe by message ID (idempotent append)
      if (message.id && store.messages.some((m) => m.id === message.id)) {
        return;
      }
      store.messages.push(message);
      store.messages = prune(store.messages);
      await writeJsonFile(filePath, store);
    });
  };

  const read: MessageHistoryStore["read"] = async (conversationId, opts) => {
    const filePath = resolveFilePath(conversationId);
    const store = await readStore(filePath, conversationId);

    // Messages are stored oldest-first; reverse for newest-first output
    const allMessages = store.messages.toReversed();
    const limit = Math.min(Math.max(1, opts?.limit ?? DEFAULT_LIMIT), MAX_LIMIT);
    const offset = parseCursor(opts?.cursor);

    const page = allMessages.slice(offset, offset + limit);
    const hasMore = offset + limit < allMessages.length;

    return {
      messages: page,
      nextCursor: hasMore ? String(offset + limit) : undefined,
    };
  };

  return { append, read };
}
