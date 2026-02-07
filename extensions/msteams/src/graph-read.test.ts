import { describe, expect, it, vi } from "vitest";
import type { MSTeamsAccessTokenProvider } from "./attachments/types.js";
import {
  buildReadMessagesUrl,
  readMSTeamsMessages,
  resolveGraphChatId,
  type GraphChatMessage,
  type ReadTarget,
} from "./graph-read.js";

// ── URL builder ─────────────────────────────────────────────────────────

describe("buildReadMessagesUrl", () => {
  it("builds a chat URL with default ordering", () => {
    const target: ReadTarget = { kind: "chat", chatId: "19:abc@thread.v2" };
    const url = buildReadMessagesUrl(target, 10);
    expect(url).toContain("/chats/19%3Aabc%40thread.v2/messages");
    expect(url).toContain("$top=10");
    expect(url).toContain("$orderby=createdDateTime desc");
  });

  it("builds a channel URL with teamId and channelId", () => {
    const target: ReadTarget = {
      kind: "channel",
      teamId: "team-123",
      channelId: "19:chan@thread.tacv2",
    };
    const url = buildReadMessagesUrl(target, 25);
    expect(url).toContain("/teams/team-123/channels/19%3Achan%40thread.tacv2/messages");
    expect(url).toContain("$top=25");
  });

  it("clamps limit to MAX_LIMIT (50)", () => {
    const target: ReadTarget = { kind: "chat", chatId: "chat1" };
    const url = buildReadMessagesUrl(target, 100);
    expect(url).toContain("$top=50");
  });

  it("clamps limit to minimum 1", () => {
    const target: ReadTarget = { kind: "chat", chatId: "chat1" };
    const url = buildReadMessagesUrl(target, -5);
    expect(url).toContain("$top=1");
  });

  it("returns cursor URL directly when cursor is provided", () => {
    const target: ReadTarget = { kind: "chat", chatId: "chat1" };
    const cursor = "https://graph.microsoft.com/v1.0/chats/chat1/messages?$skiptoken=abc123";
    const url = buildReadMessagesUrl(target, 10, cursor);
    expect(url).toBe(cursor);
  });
});

// ── Shared fixtures ─────────────────────────────────────────────────────

const mockTokenProvider: MSTeamsAccessTokenProvider = {
  getAccessToken: vi.fn(async () => "mock-token"),
};

// ── readMSTeamsMessages ─────────────────────────────────────────────────

describe("readMSTeamsMessages", () => {
  function createFetchMock(
    body: { value?: GraphChatMessage[]; "@odata.nextLink"?: string },
    status = 200,
  ): typeof fetch {
    return vi.fn(async () => ({
      ok: status >= 200 && status < 300,
      status,
      json: async () => body,
      text: async () => JSON.stringify(body),
    })) as unknown as typeof fetch;
  }

  it("fetches and normalizes chat messages", async () => {
    const messages: GraphChatMessage[] = [
      {
        id: "msg-1",
        createdDateTime: "2026-02-07T10:00:00Z",
        messageType: "message",
        from: { user: { id: "u1", displayName: "Alice" } },
        body: { contentType: "text", content: "Hello world" },
      },
      {
        id: "msg-2",
        createdDateTime: "2026-02-07T09:55:00Z",
        messageType: "message",
        from: { application: { id: "app1", displayName: "Bot" } },
        body: { contentType: "html", content: "<p>Hi <b>there</b></p>" },
      },
    ];

    const fetchFn = createFetchMock({ value: messages });
    const target: ReadTarget = { kind: "chat", chatId: "chat-abc" };

    const result = await readMSTeamsMessages({
      tokenProvider: mockTokenProvider,
      target,
      limit: 10,
      fetchFn,
    });

    expect(result.messages).toHaveLength(2);
    expect(result.messages[0]).toEqual({
      id: "msg-1",
      from: "Alice",
      body: "Hello world",
      createdDateTime: "2026-02-07T10:00:00Z",
      messageType: "message",
    });
    // HTML should be stripped
    expect(result.messages[1]?.body).toBe("Hi there");
    expect(result.messages[1]?.from).toBe("Bot");
    expect(result.nextCursor).toBeUndefined();
  });

  it("returns nextCursor from @odata.nextLink", async () => {
    const nextLink = "https://graph.microsoft.com/v1.0/chats/x/messages?$skiptoken=tok123";
    const fetchFn = createFetchMock({
      value: [
        {
          id: "m1",
          createdDateTime: "2026-01-01T00:00:00Z",
          messageType: "message",
          from: { user: { displayName: "Bob" } },
          body: { contentType: "text", content: "hey" },
        },
      ],
      "@odata.nextLink": nextLink,
    });

    const result = await readMSTeamsMessages({
      tokenProvider: mockTokenProvider,
      target: { kind: "chat", chatId: "x" },
      fetchFn,
    });

    expect(result.nextCursor).toBe(nextLink);
    expect(result.messages).toHaveLength(1);
  });

  it("filters out messages with no id", async () => {
    const fetchFn = createFetchMock({
      value: [
        { id: "", from: null, body: { content: "empty id" } },
        {
          id: "valid",
          createdDateTime: "2026-01-01T00:00:00Z",
          messageType: "message",
          from: null,
          body: { content: "ok" },
        },
      ],
    });

    const result = await readMSTeamsMessages({
      tokenProvider: mockTokenProvider,
      target: { kind: "chat", chatId: "c" },
      fetchFn,
    });

    expect(result.messages).toHaveLength(1);
    expect(result.messages[0]?.id).toBe("valid");
    expect(result.messages[0]?.from).toBe("unknown");
  });

  it("throws on 403 with permission hint", async () => {
    const fetchFn = createFetchMock({}, 403);
    await expect(
      readMSTeamsMessages({
        tokenProvider: mockTokenProvider,
        target: { kind: "channel", teamId: "t", channelId: "c" },
        fetchFn,
      }),
    ).rejects.toThrow(/403 Forbidden/);
  });

  it("throws on 404", async () => {
    const fetchFn = createFetchMock({}, 404);
    await expect(
      readMSTeamsMessages({
        tokenProvider: mockTokenProvider,
        target: { kind: "chat", chatId: "missing" },
        fetchFn,
      }),
    ).rejects.toThrow(/not found.*404/i);
  });

  it("throws on 429 rate limit", async () => {
    const fetchFn = createFetchMock({}, 429);
    await expect(
      readMSTeamsMessages({
        tokenProvider: mockTokenProvider,
        target: { kind: "chat", chatId: "x" },
        fetchFn,
      }),
    ).rejects.toThrow(/rate limit/i);
  });

  it("sends correct Authorization header", async () => {
    const fetchFn = createFetchMock({ value: [] });
    await readMSTeamsMessages({
      tokenProvider: mockTokenProvider,
      target: { kind: "chat", chatId: "test" },
      fetchFn,
    });

    expect(fetchFn).toHaveBeenCalledWith(
      expect.stringContaining("/chats/test/messages"),
      expect.objectContaining({
        headers: { Authorization: "Bearer mock-token" },
      }),
    );
  });

  it("uses cursor URL directly when provided", async () => {
    const fetchFn = createFetchMock({ value: [] });
    const cursor = "https://graph.microsoft.com/v1.0/chats/x/messages?$skiptoken=abc";

    await readMSTeamsMessages({
      tokenProvider: mockTokenProvider,
      target: { kind: "chat", chatId: "x" },
      cursor,
      fetchFn,
    });

    expect(fetchFn).toHaveBeenCalledWith(cursor, expect.any(Object));
  });

  it("uses default limit of 20", async () => {
    const fetchFn = createFetchMock({ value: [] });
    await readMSTeamsMessages({
      tokenProvider: mockTokenProvider,
      target: { kind: "chat", chatId: "x" },
      fetchFn,
    });

    expect(fetchFn).toHaveBeenCalledWith(expect.stringContaining("$top=20"), expect.any(Object));
  });
});

// ── resolveGraphChatId ──────────────────────────────────────────────────

describe("resolveGraphChatId", () => {
  // Helper: mock fetch that responds differently per URL pattern
  function createRoutedFetch(
    routes: Record<
      string,
      { ok: boolean; json?: () => Promise<unknown>; status?: number; text?: () => Promise<string> }
    >,
  ) {
    return vi.fn(async (url: string) => {
      for (const [pattern, response] of Object.entries(routes)) {
        if (url.includes(pattern)) {
          return response;
        }
      }
      return { ok: false, status: 404, text: async () => "not found" };
    });
  }

  it("resolves via installedApps when botAppId is provided", async () => {
    const fetchFn = createRoutedFetch({
      "teamwork/installedApps?": {
        ok: true,
        json: async () => ({
          value: [{ id: "install-123", teamsApp: { externalId: "bot-app-id" } }],
        }),
      },
      "installedApps/install-123/chat": {
        ok: true,
        json: async () => ({ id: "19:resolved-chat@unq.gbl.spaces" }),
      },
    });

    const result = await resolveGraphChatId({
      tokenProvider: mockTokenProvider,
      userAadObjectId: "user-aad-1",
      botAppId: "bot-app-id",
      fetchFn,
    });

    expect(result).toBe("19:resolved-chat@unq.gbl.spaces");
    // Should NOT have called /users/{id}/chats
    expect(fetchFn).not.toHaveBeenCalledWith(
      expect.stringContaining("/chats?$filter=chatType"),
      expect.anything(),
    );
  });

  it("falls back to user chats when installedApps returns empty", async () => {
    const fetchFn = createRoutedFetch({
      "teamwork/installedApps?": {
        ok: true,
        json: async () => ({ value: [] }),
      },
      "/chats?": {
        ok: true,
        json: async () => ({
          value: [
            {
              id: "19:fallback@unq.gbl.spaces",
              chatType: "oneOnOne",
              members: [{ userId: "bot-aad-id" }],
            },
          ],
        }),
      },
    });

    const result = await resolveGraphChatId({
      tokenProvider: mockTokenProvider,
      userAadObjectId: "user-aad-1",
      botAppId: "bot-app-id",
      botAadObjectId: "bot-aad-id",
      fetchFn,
    });

    expect(result).toBe("19:fallback@unq.gbl.spaces");
  });

  it("falls back to user chats when installedApps request fails", async () => {
    const fetchFn = createRoutedFetch({
      "teamwork/installedApps?": {
        ok: false,
        status: 403,
        text: async () => "forbidden",
      },
      "/chats?": {
        ok: true,
        json: async () => ({
          value: [{ id: "19:from-chats@unq.gbl.spaces", chatType: "oneOnOne", members: [] }],
        }),
      },
    });

    const result = await resolveGraphChatId({
      tokenProvider: mockTokenProvider,
      userAadObjectId: "user-aad-1",
      botAppId: "bot-app-id",
      fetchFn,
    });

    expect(result).toBe("19:from-chats@unq.gbl.spaces");
  });

  it("matches bot AAD ID case-insensitively in user chats", async () => {
    const fetchFn = createRoutedFetch({
      "/chats?": {
        ok: true,
        json: async () => ({
          value: [
            {
              id: "19:chat@unq.gbl.spaces",
              chatType: "oneOnOne",
              members: [{ userId: "BOT-AAD-ID" }],
            },
          ],
        }),
      },
    });

    const result = await resolveGraphChatId({
      tokenProvider: mockTokenProvider,
      userAadObjectId: "user-aad-1",
      botAadObjectId: "bot-aad-id",
      fetchFn,
    });

    expect(result).toBe("19:chat@unq.gbl.spaces");
  });

  it("throws when both strategies fail", async () => {
    const fetchFn = createRoutedFetch({
      "teamwork/installedApps?": { ok: false, status: 403, text: async () => "forbidden" },
      "/chats?": { ok: false, status: 403, text: async () => "forbidden" },
    });

    await expect(
      resolveGraphChatId({
        tokenProvider: mockTokenProvider,
        userAadObjectId: "user-aad-1",
        botAppId: "bot-app-id",
        fetchFn,
      }),
    ).rejects.toThrow("Could not resolve Graph chat ID");
  });

  it("skips installedApps strategy when no botAppId", async () => {
    const fetchFn = createRoutedFetch({
      "/chats?": {
        ok: true,
        json: async () => ({
          value: [{ id: "19:only@unq.gbl.spaces", chatType: "oneOnOne", members: [] }],
        }),
      },
    });

    const result = await resolveGraphChatId({
      tokenProvider: mockTokenProvider,
      userAadObjectId: "user-aad-1",
      fetchFn,
    });

    expect(result).toBe("19:only@unq.gbl.spaces");
    expect(fetchFn).not.toHaveBeenCalledWith(
      expect.stringContaining("installedApps"),
      expect.anything(),
    );
  });
});
