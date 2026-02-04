import fs from "node:fs";
import type { SessionEntry, SessionSystemPromptReport } from "../../config/sessions/types.js";
import type { CostUsageSummary, SessionCostSummary } from "../../infra/session-cost-usage.js";
import type { GatewayRequestHandlers } from "./types.js";
import { loadConfig } from "../../config/config.js";
import { resolveSessionFilePath } from "../../config/sessions/paths.js";
import { loadProviderUsageSummary } from "../../infra/provider-usage.js";
import {
  loadCostUsageSummary,
  loadSessionCostSummary,
  loadSessionUsageTimeSeries,
  discoverAllSessions,
} from "../../infra/session-cost-usage.js";
import {
  ErrorCodes,
  errorShape,
  formatValidationErrors,
  validateSessionsUsageParams,
} from "../protocol/index.js";
import { loadCombinedSessionStoreForGateway, loadSessionEntry } from "../session-utils.js";

const COST_USAGE_CACHE_TTL_MS = 30_000;

type DateRange = { startMs: number; endMs: number };

type CostUsageCacheEntry = {
  summary?: CostUsageSummary;
  updatedAt?: number;
  inFlight?: Promise<CostUsageSummary>;
};

const costUsageCache = new Map<string, CostUsageCacheEntry>();

/**
 * Parse a date string (YYYY-MM-DD) to start of day timestamp in UTC.
 * Returns undefined if invalid.
 */
const parseDateToMs = (raw: unknown): number | undefined => {
  if (typeof raw !== "string" || !raw.trim()) {
    return undefined;
  }
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(raw.trim());
  if (!match) {
    return undefined;
  }
  const [, year, month, day] = match;
  // Use UTC to ensure consistent behavior across timezones
  const ms = Date.UTC(parseInt(year), parseInt(month) - 1, parseInt(day));
  if (Number.isNaN(ms)) {
    return undefined;
  }
  return ms;
};

const parseDays = (raw: unknown): number | undefined => {
  if (typeof raw === "number" && Number.isFinite(raw)) {
    return Math.floor(raw);
  }
  if (typeof raw === "string" && raw.trim() !== "") {
    const parsed = Number(raw);
    if (Number.isFinite(parsed)) {
      return Math.floor(parsed);
    }
  }
  return undefined;
};

/**
 * Get date range from params (startDate/endDate or days).
 * Falls back to last 30 days if not provided.
 */
const parseDateRange = (params: {
  startDate?: unknown;
  endDate?: unknown;
  days?: unknown;
}): DateRange => {
  const now = new Date();
  // Use UTC for consistent date handling
  const todayStartMs = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate());
  const todayEndMs = todayStartMs + 24 * 60 * 60 * 1000 - 1;

  const startMs = parseDateToMs(params.startDate);
  const endMs = parseDateToMs(params.endDate);

  if (startMs !== undefined && endMs !== undefined) {
    // endMs should be end of day
    return { startMs, endMs: endMs + 24 * 60 * 60 * 1000 - 1 };
  }

  const days = parseDays(params.days);
  if (days !== undefined) {
    const clampedDays = Math.max(1, days);
    const start = todayStartMs - (clampedDays - 1) * 24 * 60 * 60 * 1000;
    return { startMs: start, endMs: todayEndMs };
  }

  // Default to last 30 days
  const defaultStartMs = todayStartMs - 29 * 24 * 60 * 60 * 1000;
  return { startMs: defaultStartMs, endMs: todayEndMs };
};

async function loadCostUsageSummaryCached(params: {
  startMs: number;
  endMs: number;
  config: ReturnType<typeof loadConfig>;
}): Promise<CostUsageSummary> {
  const cacheKey = `${params.startMs}-${params.endMs}`;
  const now = Date.now();
  const cached = costUsageCache.get(cacheKey);
  if (cached?.summary && cached.updatedAt && now - cached.updatedAt < COST_USAGE_CACHE_TTL_MS) {
    return cached.summary;
  }

  if (cached?.inFlight) {
    if (cached.summary) {
      return cached.summary;
    }
    return await cached.inFlight;
  }

  const entry: CostUsageCacheEntry = cached ?? {};
  const inFlight = loadCostUsageSummary({
    startMs: params.startMs,
    endMs: params.endMs,
    config: params.config,
  })
    .then((summary) => {
      costUsageCache.set(cacheKey, { summary, updatedAt: Date.now() });
      return summary;
    })
    .catch((err) => {
      if (entry.summary) {
        return entry.summary;
      }
      throw err;
    })
    .finally(() => {
      const current = costUsageCache.get(cacheKey);
      if (current?.inFlight === inFlight) {
        current.inFlight = undefined;
        costUsageCache.set(cacheKey, current);
      }
    });

  entry.inFlight = inFlight;
  costUsageCache.set(cacheKey, entry);

  if (entry.summary) {
    return entry.summary;
  }
  return await inFlight;
}

export type SessionUsageEntry = {
  key: string;
  label?: string;
  sessionId?: string;
  updatedAt?: number;
  usage: SessionCostSummary | null;
  contextWeight?: SessionSystemPromptReport | null;
};

export type SessionsUsageResult = {
  updatedAt: number;
  startDate: string;
  endDate: string;
  sessions: SessionUsageEntry[];
  totals: CostUsageSummary["totals"];
};

export const usageHandlers: GatewayRequestHandlers = {
  "usage.status": async ({ respond }) => {
    const summary = await loadProviderUsageSummary();
    respond(true, summary, undefined);
  },
  "usage.cost": async ({ respond, params }) => {
    const config = loadConfig();
    const { startMs, endMs } = parseDateRange({
      startDate: params?.startDate,
      endDate: params?.endDate,
      days: params?.days,
    });
    const summary = await loadCostUsageSummaryCached({ startMs, endMs, config });
    respond(true, summary, undefined);
  },
  "sessions.usage": async ({ respond, params }) => {
    if (!validateSessionsUsageParams(params)) {
      respond(
        false,
        undefined,
        errorShape(
          ErrorCodes.INVALID_REQUEST,
          `invalid sessions.usage params: ${formatValidationErrors(validateSessionsUsageParams.errors)}`,
        ),
      );
      return;
    }

    const p = params;
    const config = loadConfig();
    const { startMs, endMs } = parseDateRange({
      startDate: p.startDate,
      endDate: p.endDate,
    });
    const limit = typeof p.limit === "number" && Number.isFinite(p.limit) ? p.limit : 50;
    const includeContextWeight = p.includeContextWeight ?? false;
    const specificKey = typeof p.key === "string" ? p.key.trim() : null;

    // Load session store for named sessions
    const { store } = loadCombinedSessionStoreForGateway(config);
    const now = Date.now();

    // Merge discovered sessions with store entries
    type MergedEntry = {
      key: string;
      sessionId: string;
      sessionFile: string;
      label?: string;
      updatedAt: number;
      storeEntry?: SessionEntry;
      firstUserMessage?: string;
    };

    const mergedEntries: MergedEntry[] = [];

    // Optimization: If a specific key is requested, skip full directory scan
    if (specificKey) {
      // Check if it's a named session in the store
      const storeEntry = store[specificKey];
      let sessionId = storeEntry?.sessionId ?? specificKey;

      // Resolve the session file path
      const sessionFile = resolveSessionFilePath(sessionId, storeEntry);

      try {
        const stats = fs.statSync(sessionFile);
        if (stats.isFile()) {
          mergedEntries.push({
            key: specificKey,
            sessionId,
            sessionFile,
            label: storeEntry?.label,
            updatedAt: storeEntry?.updatedAt ?? stats.mtimeMs,
            storeEntry,
          });
        }
      } catch {
        // File doesn't exist - no results for this key
      }
    } else {
      // Full discovery for list view
      const discoveredSessions = await discoverAllSessions({
        startMs,
        endMs,
      });

      // Build a map of sessionId -> store entry for quick lookup
      const storeBySessionId = new Map<string, { key: string; entry: SessionEntry }>();
      for (const [key, entry] of Object.entries(store)) {
        if (entry?.sessionId) {
          storeBySessionId.set(entry.sessionId, { key, entry });
        }
      }

      for (const discovered of discoveredSessions) {
        const storeMatch = storeBySessionId.get(discovered.sessionId);
        if (storeMatch) {
          // Named session from store
          mergedEntries.push({
            key: storeMatch.key,
            sessionId: discovered.sessionId,
            sessionFile: discovered.sessionFile,
            label: storeMatch.entry.label,
            updatedAt: storeMatch.entry.updatedAt ?? discovered.mtime,
            storeEntry: storeMatch.entry,
          });
        } else {
          // Unnamed session - use session ID as key, no label
          mergedEntries.push({
            key: discovered.sessionId,
            sessionId: discovered.sessionId,
            sessionFile: discovered.sessionFile,
            label: undefined, // No label for unnamed sessions
            updatedAt: discovered.mtime,
          });
        }
      }
    }

    // Sort by most recent first
    mergedEntries.sort((a, b) => b.updatedAt - a.updatedAt);

    // Apply limit
    const limitedEntries = mergedEntries.slice(0, limit);

    // Load usage for each session
    const sessions: SessionUsageEntry[] = [];
    const aggregateTotals = {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: 0,
      totalCost: 0,
      inputCost: 0,
      outputCost: 0,
      cacheReadCost: 0,
      cacheWriteCost: 0,
      missingCostEntries: 0,
    };

    for (const merged of limitedEntries) {
      const usage = await loadSessionCostSummary({
        sessionId: merged.sessionId,
        sessionEntry: merged.storeEntry,
        sessionFile: merged.sessionFile,
        config,
        startMs,
        endMs,
      });

      if (usage) {
        aggregateTotals.input += usage.input;
        aggregateTotals.output += usage.output;
        aggregateTotals.cacheRead += usage.cacheRead;
        aggregateTotals.cacheWrite += usage.cacheWrite;
        aggregateTotals.totalTokens += usage.totalTokens;
        aggregateTotals.totalCost += usage.totalCost;
        aggregateTotals.inputCost += usage.inputCost;
        aggregateTotals.outputCost += usage.outputCost;
        aggregateTotals.cacheReadCost += usage.cacheReadCost;
        aggregateTotals.cacheWriteCost += usage.cacheWriteCost;
        aggregateTotals.missingCostEntries += usage.missingCostEntries;
      }

      sessions.push({
        key: merged.key,
        label: merged.label,
        sessionId: merged.sessionId,
        updatedAt: merged.updatedAt,
        usage,
        contextWeight: includeContextWeight
          ? (merged.storeEntry?.systemPromptReport ?? null)
          : undefined,
      });
    }

    // Format dates back to YYYY-MM-DD strings
    const formatDateStr = (ms: number) => {
      const d = new Date(ms);
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
    };

    const result: SessionsUsageResult = {
      updatedAt: now,
      startDate: formatDateStr(startMs),
      endDate: formatDateStr(endMs),
      sessions,
      totals: aggregateTotals,
    };

    respond(true, result, undefined);
  },
  "sessions.usage.timeseries": async ({ respond, params }) => {
    const key = typeof params?.key === "string" ? params.key.trim() : null;
    if (!key) {
      respond(
        false,
        undefined,
        errorShape(ErrorCodes.INVALID_REQUEST, "key is required for timeseries"),
      );
      return;
    }

    const config = loadConfig();
    const { entry } = loadSessionEntry(key);

    // For discovered sessions (not in store), try using key as sessionId directly
    const sessionId = entry?.sessionId ?? key;
    const sessionFile = entry?.sessionFile ?? resolveSessionFilePath(key);

    const timeseries = await loadSessionUsageTimeSeries({
      sessionId,
      sessionEntry: entry,
      sessionFile,
      config,
      maxPoints: 200,
    });

    if (!timeseries) {
      respond(
        false,
        undefined,
        errorShape(ErrorCodes.INVALID_REQUEST, `No transcript found for session: ${key}`),
      );
      return;
    }

    respond(true, timeseries, undefined);
  },
  "sessions.usage.logs": async ({ respond, params }) => {
    const key = typeof params?.key === "string" ? params.key.trim() : null;
    if (!key) {
      respond(false, undefined, errorShape(ErrorCodes.INVALID_REQUEST, "key is required for logs"));
      return;
    }

    const limit =
      typeof params?.limit === "number" && Number.isFinite(params.limit)
        ? Math.min(params.limit, 1000)
        : 200;

    const config = loadConfig();
    const { entry } = loadSessionEntry(key);

    // For discovered sessions (not in store), try using key as sessionId directly
    const sessionId = entry?.sessionId ?? key;
    const sessionFile = entry?.sessionFile ?? resolveSessionFilePath(key);

    const { loadSessionLogs } = await import("../../infra/session-cost-usage.js");
    const logs = await loadSessionLogs({
      sessionId,
      sessionEntry: entry,
      sessionFile,
      config,
      limit,
    });

    respond(true, { logs: logs ?? [] }, undefined);
  },
};
