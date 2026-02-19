/**
 * File-system backed message history store.
 *
 * One JSON file per conversation: `msteams-history-{hash}.json`
 */

import crypto from "node:crypto";
import type {
  MessageHistoryStore,
  MessageHistoryStoreData,
  ReadOpts,
  ReadResult,
  StoredMessage,
} from "./message-history-store.js";
import { DEFAULT_HISTORY_TTL_MS, DEFAULT_MAX_MESSAGES } from "./message-history-store.js";
import { resolveMSTeamsStorePath } from "./storage.js";
import { readJsonFile, withFileLock, writeJsonFile } from "./store-fs.js";

const DEFAULT_LIMIT = 20;
const MAX_LIMIT = 50;

function hashConversationId(id: string): string {
  return crypto.createHash("sha256").update(id).digest("hex").slice(0, 16);
}

function emptyStore(conversationId: string): MessageHistoryStoreData {
  return { version: 1, conversationId, messages: [] };
}

function parseCursor(cursor: string | undefined): number {
  if (!cursor) return 0;
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

  function prune(messages: StoredMessage[]): StoredMessage[] {
    const cutoff = Date.now() - ttlMs;
    const fresh = messages.filter((m) => {
      const ts = Date.parse(m.createdDateTime);
      return Number.isFinite(ts) && ts > cutoff;
    });
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
      if (message.id && store.messages.some((m) => m.id === message.id)) {
        return;
      }
      store.messages.push(message);
      const last = store.messages.at(-2);
      if (last && message.createdDateTime < last.createdDateTime) {
        store.messages.sort((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));
      }
      store.messages = prune(store.messages);
      await writeJsonFile(filePath, store);
    });
  };

  const read: MessageHistoryStore["read"] = async (conversationId, opts) => {
    const filePath = resolveFilePath(conversationId);
    const store = await readStore(filePath, conversationId);

    let filtered = store.messages;
    if (opts?.after) {
      filtered = filtered.filter((m) => m.createdDateTime > opts.after!);
    }
    if (opts?.before) {
      filtered = filtered.filter((m) => m.createdDateTime < opts.before!);
    }

    const allMessages = filtered.toReversed();
    const limit = Math.min(Math.max(1, opts?.limit ?? DEFAULT_LIMIT), MAX_LIMIT);
    const offset = parseCursor(opts?.cursor);
    const page = allMessages.slice(offset, offset + limit);
    const hasMore = offset + limit < allMessages.length;

    return {
      messages: page,
      nextCursor: hasMore ? String(offset + limit) : undefined,
    };
  };

  const readDefault: MessageHistoryStore["readDefault"] = async (opts) => {
    // Find the most recently modified history file
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    const storeDir = resolveMSTeamsStorePath({
      filename: ".",
      env: params?.env,
      homedir: params?.homedir,
      stateDir: params?.stateDir,
    });
    try {
      const files = await fs.readdir(storeDir);
      const historyFiles = files.filter(
        (f) => f.startsWith("msteams-history-") && f.endsWith(".json"),
      );
      if (historyFiles.length === 0) {
        return { messages: [] };
      }
      // Pick most recently modified
      let best = historyFiles[0];
      let bestMtime = 0;
      for (const f of historyFiles) {
        const stat = await fs.stat(path.join(storeDir, f));
        if (stat.mtimeMs > bestMtime) {
          bestMtime = stat.mtimeMs;
          best = f;
        }
      }
      const filePath = path.join(storeDir, best);
      const { value } = await readJsonFile<MessageHistoryStoreData>(filePath, emptyStore(""));
      if (value.version !== 1 || !Array.isArray(value.messages)) {
        return { messages: [] };
      }
      return read(value.conversationId, opts);
    } catch {
      return { messages: [] };
    }
  };

  return { append, read, readDefault };
}
