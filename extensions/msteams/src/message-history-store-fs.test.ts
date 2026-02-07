import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import type { NormalizedTeamsMessage } from "./graph-read.js";
import { createMessageHistoryStoreFs } from "./message-history-store-fs.js";

function makeMsg(
  id: string,
  body: string,
  from = "Alice",
  isoDate?: string,
): NormalizedTeamsMessage {
  return {
    id,
    from,
    body,
    createdDateTime: isoDate ?? new Date().toISOString(),
    messageType: "message",
  };
}

describe("createMessageHistoryStoreFs", () => {
  const tmpDirs: string[] = [];

  function createTmpDir(): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "msteams-history-test-"));
    tmpDirs.push(dir);
    return dir;
  }

  afterEach(() => {
    for (const dir of tmpDirs) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
    tmpDirs.length = 0;
  });

  it("appends and reads messages (newest first)", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir });

    const m1 = makeMsg("1", "hello", "Alice", "2026-02-01T10:00:00Z");
    const m2 = makeMsg("2", "world", "Bob", "2026-02-01T10:01:00Z");
    const m3 = makeMsg("3", "bye", "Alice", "2026-02-01T10:02:00Z");

    await store.append("conv-1", m1);
    await store.append("conv-1", m2);
    await store.append("conv-1", m3);

    const result = await store.read("conv-1");
    expect(result.messages).toHaveLength(3);
    // Newest first
    expect(result.messages[0]?.body).toBe("bye");
    expect(result.messages[1]?.body).toBe("world");
    expect(result.messages[2]?.body).toBe("hello");
  });

  it("deduplicates by message ID", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir });

    const m1 = makeMsg("same-id", "first");
    await store.append("conv-1", m1);
    await store.append("conv-1", { ...m1, body: "duplicate" });

    const result = await store.read("conv-1");
    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]?.body).toBe("first");
  });

  it("respects limit parameter", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir });

    for (let i = 0; i < 10; i++) {
      await store.append(
        "conv-1",
        makeMsg(String(i), `msg ${i}`, "User", `2026-02-01T10:0${i}:00Z`),
      );
    }

    const result = await store.read("conv-1", { limit: 3 });
    expect(result.messages).toHaveLength(3);
    expect(result.nextCursor).toBeDefined();
  });

  it("paginates with cursor", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir });

    for (let i = 0; i < 5; i++) {
      await store.append(
        "conv-1",
        makeMsg(String(i), `msg ${i}`, "User", `2026-02-01T10:0${i}:00Z`),
      );
    }

    const page1 = await store.read("conv-1", { limit: 2 });
    expect(page1.messages).toHaveLength(2);
    expect(page1.nextCursor).toBe("2");

    const page2 = await store.read("conv-1", { limit: 2, cursor: page1.nextCursor });
    expect(page2.messages).toHaveLength(2);
    expect(page2.nextCursor).toBe("4");

    const page3 = await store.read("conv-1", { limit: 2, cursor: page2.nextCursor });
    expect(page3.messages).toHaveLength(1);
    expect(page3.nextCursor).toBeUndefined();
  });

  it("prunes messages beyond maxMessages", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir, maxMessages: 5 });

    for (let i = 0; i < 10; i++) {
      await store.append(
        "conv-1",
        makeMsg(String(i), `msg ${i}`, "User", `2026-02-01T10:${String(i).padStart(2, "0")}:00Z`),
      );
    }

    const result = await store.read("conv-1", { limit: 50 });
    expect(result.messages).toHaveLength(5);
    // Should keep the 5 newest (msg 5-9)
    expect(result.messages[0]?.body).toBe("msg 9");
    expect(result.messages[4]?.body).toBe("msg 5");
  });

  it("prunes expired messages", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir, ttlMs: 1000 });

    const old = makeMsg("old", "ancient", "User", new Date(Date.now() - 5000).toISOString());
    const fresh = makeMsg("new", "recent", "User", new Date().toISOString());

    await store.append("conv-1", old);
    await store.append("conv-1", fresh);

    const result = await store.read("conv-1");
    // Only the fresh message should remain (old was pruned on write)
    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]?.body).toBe("recent");
  });

  it("isolates conversations by ID", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir });

    await store.append("conv-a", makeMsg("1", "msg A"));
    await store.append("conv-b", makeMsg("2", "msg B"));

    const resultA = await store.read("conv-a");
    expect(resultA.messages).toHaveLength(1);
    expect(resultA.messages[0]?.body).toBe("msg A");

    const resultB = await store.read("conv-b");
    expect(resultB.messages).toHaveLength(1);
    expect(resultB.messages[0]?.body).toBe("msg B");
  });

  it("returns empty for unknown conversation", async () => {
    const stateDir = createTmpDir();
    const store = createMessageHistoryStoreFs({ stateDir });

    const result = await store.read("nonexistent");
    expect(result.messages).toHaveLength(0);
    expect(result.nextCursor).toBeUndefined();
  });
});
