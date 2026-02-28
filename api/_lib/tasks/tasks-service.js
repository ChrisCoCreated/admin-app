const { createGraphDelegatedClient } = require("./graph-delegated-client");
const { resolveUserUpn } = require("./identity");
const { mergeTasksWithOverlays, sortUnifiedTasks } = require("./merge-sort");
const {
  clearOverlayUserCache,
  createOverlayItem,
  listPinnedOverlaysByUser,
  listOverlaysByUser,
  patchOverlayFieldsBatch,
  patchOverlayFields,
} = require("./overlay-repository");
const { applyOverlayBehaviorRules } = require("./overlay-rules");
const { fetchPlannerTasks } = require("./planner-provider");
const { fetchTodoTasks } = require("./todo-provider");
const {
  buildTaskKey,
  fromOverlayFields,
  normalizeExternalTaskId,
  normalizeProvider,
  sanitizeOverlayPatch,
  toGraphOverlayFields,
} = require("./unified-model");

const unifiedTasksCache = new Map();
const whiteboardTasksCache = new Map();
const whiteboardSyncStateByUser = new Map();
const GRAPH_TASKS_DEBUG = process.env.GRAPH_TASKS_DEBUG === "1";

function unifiedTasksCacheKey(userUpn) {
  return String(userUpn || "").trim().toLowerCase();
}

function getUnifiedTasksCacheTtlMs() {
  const configured = Number(process.env.TASKS_UNIFIED_CACHE_TTL_MS || 90000);
  if (!Number.isFinite(configured) || configured < 0) {
    return 90000;
  }
  return Math.floor(configured);
}

function clearUnifiedTasksCache(userUpn) {
  if (userUpn) {
    unifiedTasksCache.delete(unifiedTasksCacheKey(userUpn));
    return;
  }
  unifiedTasksCache.clear();
}

function whiteboardCacheKey(userUpn) {
  return String(userUpn || "").trim().toLowerCase();
}

function getWhiteboardCacheTtlMs() {
  const configured = Number(process.env.TASKS_WHITEBOARD_CACHE_TTL_MS || 30000);
  if (!Number.isFinite(configured) || configured < 0) {
    return 30000;
  }
  return Math.floor(configured);
}

function getWhiteboardSyncStaleMs() {
  const configured = Number(process.env.TASKS_WHITEBOARD_SYNC_STALE_MS || 300000);
  if (!Number.isFinite(configured) || configured < 1000) {
    return 300000;
  }
  return Math.floor(configured);
}

function getWhiteboardSyncCooldownMs() {
  const configured = Number(process.env.TASKS_WHITEBOARD_SYNC_COOLDOWN_MS || 90000);
  if (!Number.isFinite(configured) || configured < 0) {
    return 90000;
  }
  return Math.floor(configured);
}

function clearWhiteboardTasksCache(userUpn) {
  if (userUpn) {
    whiteboardTasksCache.delete(whiteboardCacheKey(userUpn));
    return;
  }
  whiteboardTasksCache.clear();
}

function buildUnifiedTasks(graphClient, userUpn) {
  const providerErrors = {};

  const todoPromise = fetchTodoTasks(graphClient)
    .then((tasks) => ({ ok: true, tasks }))
    .catch((error) => ({ ok: false, error }));
  const plannerPromise = fetchPlannerTasks(graphClient)
    .then((tasks) => ({ ok: true, tasks }))
    .catch((error) => ({ ok: false, error }));
  const overlaysPromise = listOverlaysByUser(graphClient, userUpn)
    .then((bundle) => ({ ok: true, bundle }))
    .catch((error) => ({ ok: false, error }));

  return Promise.all([todoPromise, plannerPromise, overlaysPromise]).then(
    ([todoResult, plannerResult, overlaysResult]) => {
      const todoTasks = todoResult.ok
        ? todoResult.tasks.filter((task) => task && task.isCompleted !== true)
        : [];
      if (!todoResult.ok) {
        providerErrors.todo = toProviderError(todoResult.error);
      }

      const plannerTasks = plannerResult.ok
        ? plannerResult.tasks.filter((task) => task && task.isCompleted !== true)
        : [];
      if (!plannerResult.ok) {
        providerErrors.planner = toProviderError(plannerResult.error);
      }

      const overlaysBundle = overlaysResult.ok
        ? overlaysResult.bundle
        : { siteId: "", listId: "", overlays: [], byKey: new Map() };
      if (!overlaysResult.ok) {
        providerErrors.overlay = toProviderError(overlaysResult.error);
      }

      if (!todoResult.ok && !plannerResult.ok) {
        const error = new Error("Both To Do and Planner providers failed.");
        error.status = 502;
        error.code = "UNIFIED_PROVIDERS_FAILED";
        error.retryable = false;
        throw error;
      }

      const allTasks = [...todoTasks, ...plannerTasks];
      const merged = mergeTasksWithOverlays(allTasks, overlaysBundle.byKey);
      const sortedTasks = sortUnifiedTasks(merged.tasks);
      const partial = Object.keys(providerErrors).length > 0;

      return {
        tasks: sortedTasks,
        meta: {
          total: sortedTasks.length,
          todoCount: todoTasks.length,
          plannerCount: plannerTasks.length,
          overlayMatchedCount: merged.overlayMatchedCount,
          overlayOrphanCount: merged.overlayOrphanCount,
          partial,
          providerErrors,
        },
      };
    }
  );
}

function mapGraphError(error) {
  const message = error?.message || "Graph request failed.";
  const status = Number(error?.status) || 502;
  let code = String(error?.code || "GRAPH_REQUEST_FAILED");
  const retryable = Boolean(error?.retryable);
  const lowerCode = code.toLowerCase();
  if (
    status === 401 ||
    lowerCode.includes("invalidauthenticationtoken") ||
    lowerCode.includes("token")
  ) {
    code = "TOKEN_EXPIRED_OR_INVALID";
  }
  if (/planner/i.test(message) && status === 403) {
    code = "FORBIDDEN_PLANNER";
  }

  const correlationId = `tasks_${Date.now().toString(36)}_${Math.random()
    .toString(36)
    .slice(2, 8)}`;

  return {
    status,
    payload: {
      error: {
        code,
        message,
        retryable,
        correlationId,
      },
    },
  };
}

function toProviderError(error) {
  return {
    code: String(error?.code || "PROVIDER_FETCH_FAILED"),
    status: Number(error?.status) || 502,
    message: error?.message || "Provider fetch failed.",
    retryable: Boolean(error?.retryable),
  };
}

function logTasksDebug(message, details) {
  if (!GRAPH_TASKS_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[tasks-service] ${message}`, details);
    return;
  }
  console.log(`[tasks-service] ${message}`);
}

function looksOpaqueTaskId(value) {
  const text = String(value || "").trim();
  if (!text || /\s/.test(text)) {
    return false;
  }
  return text.length >= 28;
}

function shouldBackfillOverlayTitle(overlayTitle, externalTaskId, graphTitle) {
  const normalizedGraphTitle = String(graphTitle || "").trim();
  if (!normalizedGraphTitle) {
    return false;
  }

  const normalizedOverlayTitle = String(overlayTitle || "").trim();
  if (!normalizedOverlayTitle) {
    return true;
  }
  if (normalizedOverlayTitle === normalizedGraphTitle) {
    return false;
  }
  if ((externalTaskId && normalizedOverlayTitle === externalTaskId) || looksOpaqueTaskId(normalizedOverlayTitle)) {
    return true;
  }
  return false;
}

async function getUnifiedTasks({ graphAccessToken, claims }) {
  const userUpn = resolveUserUpn(claims);
  const graphClient = createGraphDelegatedClient(graphAccessToken);
  return buildUnifiedTasks(graphClient, userUpn);
}

function maxLastExternalSyncAt(overlays) {
  let maxMs = 0;
  let maxIso = null;
  for (const overlay of overlays || []) {
    const value = String(overlay?.lastExternalSyncAt || "").trim();
    if (!value) {
      continue;
    }
    const ms = Date.parse(value);
    if (!Number.isFinite(ms)) {
      continue;
    }
    if (ms > maxMs) {
      maxMs = ms;
      maxIso = new Date(ms).toISOString();
    }
  }
  return maxIso;
}

function mapWhiteboardTasksFromOverlay(overlays) {
  return (overlays || []).map((overlay) => {
    return {
      provider: String(overlay.provider || "").trim().toLowerCase(),
      externalTaskId: String(overlay.externalTaskId || "").trim(),
      externalContainerId: "",
      title: String(overlay.title || "").trim() || "Untitled task",
      createdDateTimeUtc: null,
      dueDateTimeUtc: overlay?.lastKnownDueDateUtc || null,
      isCompleted: overlay?.lastKnownCompleted === true,
      completedDateTimeUtc: null,
      source: {
        rawId: String(overlay.itemId || "").trim(),
      },
      overlay: {
        itemId: overlay.itemId,
        workingStatus: overlay.workingStatus,
        workType: overlay.workType,
        tags: overlay.tags,
        activeStartedAt: overlay.activeStartedAt,
        lastWorkedAt: overlay.lastWorkedAt,
        energy: overlay.energy,
        effortMinutes: overlay.effortMinutes,
        impact: overlay.impact,
        overlayNotes: overlay.overlayNotes,
        pinned: overlay.pinned === true,
        layout: overlay.layout || "",
        category: overlay.category || "",
        externalState: overlay.externalState || "",
        lastExternalSyncAt: overlay.lastExternalSyncAt || null,
        lastKnownDueDateUtc: overlay.lastKnownDueDateUtc || null,
        lastKnownCompleted: overlay.lastKnownCompleted === true,
        lastOverlayUpdatedAt: overlay.lastOverlayUpdatedAt,
      },
    };
  });
}

async function getWhiteboardTasks({ graphAccessToken, claims }) {
  const userUpn = resolveUserUpn(claims);
  const key = whiteboardCacheKey(userUpn);
  const now = Date.now();
  const cached = whiteboardTasksCache.get(key);
  if (cached?.value && cached.expiresAt > now) {
    return {
      ...cached.value,
      meta: {
        ...(cached.value?.meta || {}),
        cached: true,
      },
    };
  }

  const graphClient = createGraphDelegatedClient(graphAccessToken);
  const overlaysBundle = await listPinnedOverlaysByUser(graphClient, userUpn);
  const overlays = Array.isArray(overlaysBundle?.overlays) ? overlaysBundle.overlays : [];
  const totalByProvider = {
    todo: 0,
    planner: 0,
    other: 0,
  };

  for (const overlay of overlaysBundle?.allOverlays || []) {
    const provider = String(overlay?.provider || "").trim().toLowerCase();
    if (provider === "todo") {
      totalByProvider.todo += 1;
      continue;
    }
    if (provider === "planner") {
      totalByProvider.planner += 1;
      continue;
    }
    totalByProvider.other += 1;
  }

  const tasks = mapWhiteboardTasksFromOverlay(overlays);

  const pinnedByProvider = {
    todo: 0,
    planner: 0,
    other: 0,
  };
  for (const task of tasks) {
    const provider = String(task?.provider || "").trim().toLowerCase();
    if (provider === "todo") {
      pinnedByProvider.todo += 1;
      continue;
    }
    if (provider === "planner") {
      pinnedByProvider.planner += 1;
      continue;
    }
    pinnedByProvider.other += 1;
  }

  const lastSyncedAt = maxLastExternalSyncAt(overlays);
  const lastSyncedMs = lastSyncedAt ? Date.parse(lastSyncedAt) : 0;
  const syncStale = !lastSyncedAt || !Number.isFinite(lastSyncedMs) || now - lastSyncedMs > getWhiteboardSyncStaleMs();

  const payload = {
    tasks,
    meta: {
      total: tasks.length,
      source: "taskoverlay_fast",
      cached: false,
      totalOverlayRows: Number(overlaysBundle?.totalRows || overlays.length),
      requestedUserUpn: userUpn,
      totalByProvider,
      pinnedByProvider,
      lastSyncedAt,
      syncStale,
    },
  };

  whiteboardTasksCache.set(key, {
    value: payload,
    expiresAt: Date.now() + getWhiteboardCacheTtlMs(),
  });

  return payload;
}

async function syncWhiteboardTasks({ graphAccessToken, claims, body }) {
  const userUpn = resolveUserUpn(claims);
  const force = Boolean(body?.force);
  const key = whiteboardCacheKey(userUpn);
  const now = Date.now();
  const state = whiteboardSyncStateByUser.get(key);
  const cooldownMs = getWhiteboardSyncCooldownMs();

  if (state?.inFlight && !force) {
    return {
      statusCode: 202,
      meta: {
        alreadyRunning: true,
        cooldownMs,
      },
    };
  }

  if (!force && state?.lastFinishedAt && now - state.lastFinishedAt < cooldownMs) {
    return {
      statusCode: 202,
      meta: {
        skippedCooldown: true,
        cooldownMs,
        nextAllowedAt: new Date(state.lastFinishedAt + cooldownMs).toISOString(),
      },
    };
  }

  const startSync = async () => {
    const startedAt = Date.now();
    const nowIso = new Date().toISOString();
    const graphClient = createGraphDelegatedClient(graphAccessToken);
    const overlaysBundle = await listPinnedOverlaysByUser(graphClient, userUpn);
    const overlays = Array.isArray(overlaysBundle?.overlays) ? overlaysBundle.overlays : [];
    if (overlays.length === 0) {
      clearWhiteboardTasksCache(userUpn);
      return {
        statusCode: 200,
        meta: {
          scanned: 0,
          matched: 0,
          updated: 0,
          unpinnedMissing: 0,
          failed: 0,
          durationMs: Date.now() - startedAt,
          partial: false,
          providerErrors: {},
        },
      };
    }

    const [todoResult, plannerResult] = await Promise.allSettled([
      fetchTodoTasks(graphClient),
      fetchPlannerTasks(graphClient),
    ]);

    const graphTasksByKey = new Map();
    const providerErrors = {};

    if (todoResult.status === "fulfilled") {
      for (const task of todoResult.value || []) {
        graphTasksByKey.set(buildTaskKey(task.provider, task.externalTaskId), task);
      }
    } else {
      providerErrors.todo = toProviderError(todoResult.reason);
    }

    if (plannerResult.status === "fulfilled") {
      for (const task of plannerResult.value || []) {
        graphTasksByKey.set(buildTaskKey(task.provider, task.externalTaskId), task);
      }
    } else {
      providerErrors.planner = toProviderError(plannerResult.reason);
    }

    const patches = [];
    let matched = 0;
    let unpinnedMissing = 0;
    for (const overlay of overlays) {
      const keyValue = buildTaskKey(overlay.provider, overlay.externalTaskId);
      const graphTask = graphTasksByKey.get(keyValue);
      const fields = {
        LastExternalSyncAt: nowIso,
      };

      if (graphTask) {
        matched += 1;
        fields.ExternalState = "ok";
        fields.LastKnownDueDateUtc = graphTask?.dueDateTimeUtc || null;
        fields.LastKnownCompleted = graphTask?.isCompleted === true;
        const nextTitle = String(graphTask?.title || "").trim();
        if (shouldBackfillOverlayTitle(overlay?.title, overlay?.externalTaskId, nextTitle)) {
          fields.Title = nextTitle;
        }
      } else {
        fields.ExternalState = "missing";
        fields.Pinned = false;
        unpinnedMissing += 1;
        logTasksDebug("Auto-unpin whiteboard task missing in Graph.", {
          userUpn,
          taskKey: keyValue,
        });
      }

      patches.push({
        itemId: overlay.itemId,
        fields,
      });
    }

    const patchResults = await patchOverlayFieldsBatch(
      graphClient,
      overlaysBundle.siteId,
      overlaysBundle.listId,
      patches,
      overlaysBundle.fieldMap,
      4
    );

    let updated = 0;
    let failed = 0;
    for (const result of patchResults) {
      if (result?.ok) {
        updated += 1;
        continue;
      }
      failed += 1;
      if (result?.error) {
        logTasksDebug("Whiteboard sync row patch failed.", {
          userUpn,
          status: result.error?.status || 0,
          code: result.error?.code || "",
          message: result.error?.message || "",
        });
      }
    }

    clearOverlayUserCache(userUpn);
    clearWhiteboardTasksCache(userUpn);
    const durationMs = Date.now() - startedAt;
    return {
      statusCode: Object.keys(providerErrors).length > 0 ? 207 : 200,
      meta: {
        scanned: overlays.length,
        matched,
        updated,
        unpinnedMissing,
        failed,
        durationMs,
        partial: Object.keys(providerErrors).length > 0,
        providerErrors,
      },
    };
  };

  const inFlight = startSync()
    .finally(() => {
      const current = whiteboardSyncStateByUser.get(key);
      whiteboardSyncStateByUser.set(key, {
        inFlight: null,
        lastFinishedAt: Date.now(),
        lastResult: current?.lastResult || null,
      });
    });

  whiteboardSyncStateByUser.set(key, {
    inFlight,
    lastFinishedAt: state?.lastFinishedAt || 0,
    lastResult: state?.lastResult || null,
  });

  const result = await inFlight;
  whiteboardSyncStateByUser.set(key, {
    inFlight: null,
    lastFinishedAt: Date.now(),
    lastResult: result,
  });

  return result;
}

async function getUnifiedTasksCached({ graphAccessToken, claims }) {
  const userUpn = resolveUserUpn(claims);
  const key = unifiedTasksCacheKey(userUpn);
  const now = Date.now();
  const cached = unifiedTasksCache.get(key);

  if (cached && cached.value && cached.expiresAt > now) {
    return cached.value;
  }

  const startRefresh = () => {
    const current = unifiedTasksCache.get(key);
    if (current?.inFlight) {
      return current.inFlight;
    }

    const graphClient = createGraphDelegatedClient(graphAccessToken);
    const inFlight = buildUnifiedTasks(graphClient, userUpn)
      .then((value) => {
        unifiedTasksCache.set(key, {
          value,
          expiresAt: Date.now() + getUnifiedTasksCacheTtlMs(),
          inFlight: null,
        });
        return value;
      })
      .catch((error) => {
        const previous = unifiedTasksCache.get(key);
        unifiedTasksCache.set(key, {
          value: previous?.value || null,
          expiresAt: previous?.expiresAt || 0,
          inFlight: null,
        });
        throw error;
      });

    unifiedTasksCache.set(key, {
      value: cached?.value || null,
      expiresAt: cached?.expiresAt || 0,
      inFlight,
    });

    return inFlight;
  };

  if (cached?.value) {
    void startRefresh();
    return cached.value;
  }

  return startRefresh();
}

async function upsertOverlay({ graphAccessToken, claims, body }) {
  const provider = normalizeProvider(body?.provider);
  const externalTaskId = normalizeExternalTaskId(body?.externalTaskId);
  const patch = sanitizeOverlayPatch(body?.patch || {});
  const nowIso = new Date().toISOString();

  const userUpn = resolveUserUpn(claims);
  const graphClient = createGraphDelegatedClient(graphAccessToken);

  const overlaysBundle = await listOverlaysByUser(graphClient, userUpn);
  const key = buildTaskKey(provider, externalTaskId);
  const existing = overlaysBundle.byKey.get(key);

  const nextPatch = applyOverlayBehaviorRules(existing || null, patch, nowIso);
  const payload = {
    ...nextPatch,
    userUpn,
    userUpnLookupId: overlaysBundle?.userLookupId || null,
    provider,
    externalTaskId,
    lastOverlayUpdatedAt: nowIso,
  };

  const graphFields = toGraphOverlayFields(payload);
  let itemId = "";
  let created = false;

  if (existing) {
    itemId = String(existing.itemId || "").trim();
    await patchOverlayFields(
      graphClient,
      overlaysBundle.siteId,
      overlaysBundle.listId,
      itemId,
      graphFields,
      overlaysBundle.fieldMap
    );
  } else {
    const createdItem = await createOverlayItem(
      graphClient,
      overlaysBundle.siteId,
      overlaysBundle.listId,
      graphFields,
      overlaysBundle.fieldMap
    );
    itemId = String(createdItem?.id || "").trim();
    created = true;
  }

  clearOverlayUserCache(userUpn);
  clearUnifiedTasksCache(userUpn);
  clearWhiteboardTasksCache(userUpn);

  const overlay = fromOverlayFields({
    id: itemId,
    fields: graphFields,
  });

  return {
    overlay: {
      itemId: overlay?.itemId || itemId,
      provider,
      externalTaskId,
      workingStatus: overlay?.workingStatus || "",
      workType: overlay?.workType || "",
      title: overlay?.title || "",
      tags: overlay?.tags || [],
      activeStartedAt: overlay?.activeStartedAt || null,
      lastWorkedAt: overlay?.lastWorkedAt || null,
      energy: overlay?.energy || "",
      effortMinutes: overlay?.effortMinutes ?? null,
      impact: overlay?.impact || "",
      overlayNotes: overlay?.overlayNotes || "",
      pinned: overlay?.pinned === true,
      layout: overlay?.layout || "",
      category: overlay?.category || "",
      lastOverlayUpdatedAt: overlay?.lastOverlayUpdatedAt || nowIso,
    },
    meta: {
      created,
      itemId,
      updatedAt: nowIso,
    },
  };
}

module.exports = {
  clearWhiteboardTasksCache,
  clearUnifiedTasksCache,
  getUnifiedTasks,
  getUnifiedTasksCached,
  getWhiteboardTasks,
  mapGraphError,
  syncWhiteboardTasks,
  upsertOverlay,
};
