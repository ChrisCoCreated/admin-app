const { createGraphDelegatedClient } = require("./graph-delegated-client");
const { resolveUserUpn } = require("./identity");
const { mergeTasksWithOverlays, sortUnifiedTasks } = require("./merge-sort");
const {
  backfillMissingOverlayUsers,
  clearOverlayUserCache,
  createOverlayItem,
  listOverlaysByUser,
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

function mapLimit(items, concurrency, mapper) {
  return new Promise((resolve, reject) => {
    if (!Array.isArray(items) || items.length === 0) {
      resolve([]);
      return;
    }

    const results = new Array(items.length);
    let nextIndex = 0;
    let active = 0;
    let completed = 0;

    function launch() {
      if (completed >= items.length) {
        resolve(results);
        return;
      }

      while (active < concurrency && nextIndex < items.length) {
        const currentIndex = nextIndex;
        nextIndex += 1;
        active += 1;

        Promise.resolve(mapper(items[currentIndex], currentIndex))
          .then((value) => {
            results[currentIndex] = value;
            active -= 1;
            completed += 1;
            launch();
          })
          .catch(reject);
      }
    }

    launch();
  });
}

function looksOpaqueTaskId(value) {
  const text = String(value || "").trim();
  if (!text || /\s/.test(text)) {
    return false;
  }
  return text.length >= 28;
}

function pickDisplayTitle(graphTask, overlay) {
  const graphTitle = String(graphTask?.title || "").trim();
  if (graphTitle) {
    return graphTitle;
  }

  const overlayTitle = String(overlay?.title || "").trim();
  const externalTaskId = String(overlay?.externalTaskId || "").trim();
  if (!overlayTitle) {
    return "Untitled task";
  }
  if ((externalTaskId && overlayTitle === externalTaskId) || looksOpaqueTaskId(overlayTitle)) {
    return "Untitled task";
  }
  return overlayTitle;
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

async function getWhiteboardTasks({ graphAccessToken, claims }) {
  const userUpn = resolveUserUpn(claims);
  const graphClient = createGraphDelegatedClient(graphAccessToken);
  const backfillResult = await backfillMissingOverlayUsers(graphClient, userUpn);
  if (Number(backfillResult?.updatedRows || 0) > 0) {
    clearOverlayUserCache(userUpn);
  }
  const overlaysBundle = await listOverlaysByUser(graphClient, userUpn);
  const overlays = Array.isArray(overlaysBundle?.overlays) ? overlaysBundle.overlays : [];
  const providerErrors = {};
  const graphTasksByKey = new Map();

  const [todoResult, plannerResult] = await Promise.allSettled([
    fetchTodoTasks(graphClient),
    fetchPlannerTasks(graphClient),
  ]);

  if (todoResult.status === "fulfilled") {
    for (const task of todoResult.value || []) {
      const key = buildTaskKey(task.provider, task.externalTaskId);
      graphTasksByKey.set(key, task);
    }
  } else {
    providerErrors.todo = toProviderError(todoResult.reason);
  }

  if (plannerResult.status === "fulfilled") {
    for (const task of plannerResult.value || []) {
      const key = buildTaskKey(task.provider, task.externalTaskId);
      graphTasksByKey.set(key, task);
    }
  } else {
    providerErrors.planner = toProviderError(plannerResult.reason);
  }

  const titleBackfillOps = [];
  for (const overlay of overlays) {
    const key = buildTaskKey(overlay.provider, overlay.externalTaskId);
    const graphTask = graphTasksByKey.get(key);
    const graphTitle = String(graphTask?.title || "").trim();
    if (!shouldBackfillOverlayTitle(overlay?.title, overlay?.externalTaskId, graphTitle)) {
      continue;
    }

    overlay.title = graphTitle;
    titleBackfillOps.push({
      itemId: String(overlay?.itemId || "").trim(),
      title: graphTitle,
    });
  }

  if (titleBackfillOps.length > 0) {
    await mapLimit(titleBackfillOps, 3, async (entry) => {
      if (!entry.itemId || !entry.title) {
        return;
      }
      try {
        await patchOverlayFields(
          graphClient,
          overlaysBundle.siteId,
          overlaysBundle.listId,
          entry.itemId,
          { Title: entry.title },
          overlaysBundle.fieldMap
        );
      } catch (error) {
        logTasksDebug("TaskOverlay title backfill failed for whiteboard row.", {
          itemId: entry.itemId,
          status: error?.status || 0,
          code: error?.code || "",
          message: error?.message || "",
        });
      }
    });
    if (titleBackfillOps.length > 0) {
      clearOverlayUserCache(userUpn);
    }
  }

  const totalByProvider = {
    todo: 0,
    planner: 0,
    other: 0,
  };

  for (const overlay of overlays) {
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

  let graphMatchedCount = 0;
  const tasks = overlays
    .filter((overlay) => overlay?.pinned === true)
    .map((overlay) => {
      const key = buildTaskKey(overlay.provider, overlay.externalTaskId);
      const graphTask = graphTasksByKey.get(key);
      if (graphTask) {
        graphMatchedCount += 1;
      }
      return {
        provider: String(overlay.provider || "").trim().toLowerCase(),
        externalTaskId: String(overlay.externalTaskId || "").trim(),
        externalContainerId: String(graphTask?.externalContainerId || "").trim(),
        title: pickDisplayTitle(graphTask, overlay),
        createdDateTimeUtc: graphTask?.createdDateTimeUtc || null,
        dueDateTimeUtc: graphTask?.dueDateTimeUtc || null,
        isCompleted: graphTask?.isCompleted === true,
        completedDateTimeUtc: graphTask?.completedDateTimeUtc || null,
        source: {
          rawId: String(graphTask?.source?.rawId || overlay.itemId || "").trim(),
          etag: graphTask?.source?.etag,
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
          lastOverlayUpdatedAt: overlay.lastOverlayUpdatedAt,
        },
      };
    });

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

  return {
    tasks,
    meta: {
      total: tasks.length,
      source: "taskoverlay",
      totalOverlayRows: overlays.length,
      requestedUserUpn: userUpn,
      totalByProvider,
      pinnedByProvider,
      graphMatchedCount,
      graphMissCount: Math.max(0, tasks.length - graphMatchedCount),
      titleBackfilledCount: titleBackfillOps.length,
      partial: Object.keys(providerErrors).length > 0,
      providerErrors,
    },
  };
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
  clearUnifiedTasksCache,
  getUnifiedTasks,
  getUnifiedTasksCached,
  getWhiteboardTasks,
  mapGraphError,
  upsertOverlay,
};
