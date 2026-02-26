const { createGraphDelegatedClient } = require("./graph-delegated-client");
const { resolveUserUpn } = require("./identity");
const { mergeTasksWithOverlays, sortUnifiedTasks } = require("./merge-sort");
const {
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

async function getUnifiedTasks({ graphAccessToken, claims }) {
  const userUpn = resolveUserUpn(claims);
  const graphClient = createGraphDelegatedClient(graphAccessToken);

  const [todoTasks, plannerTasks, overlaysBundle] = await Promise.all([
    fetchTodoTasks(graphClient),
    fetchPlannerTasks(graphClient),
    listOverlaysByUser(graphClient, userUpn),
  ]);

  const allTasks = [...todoTasks, ...plannerTasks];
  const merged = mergeTasksWithOverlays(allTasks, overlaysBundle.byKey);
  const sortedTasks = sortUnifiedTasks(merged.tasks);

  return {
    tasks: sortedTasks,
    meta: {
      total: sortedTasks.length,
      todoCount: todoTasks.length,
      plannerCount: plannerTasks.length,
      overlayMatchedCount: merged.overlayMatchedCount,
      overlayOrphanCount: merged.overlayOrphanCount,
    },
  };
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
    provider,
    externalTaskId,
    lastOverlayUpdatedAt: nowIso,
  };

  const graphFields = toGraphOverlayFields(payload);
  let itemId = "";
  let created = false;

  if (existing) {
    itemId = String(existing.itemId || "").trim();
    await patchOverlayFields(graphClient, overlaysBundle.siteId, overlaysBundle.listId, itemId, graphFields);
  } else {
    const createdItem = await createOverlayItem(
      graphClient,
      overlaysBundle.siteId,
      overlaysBundle.listId,
      graphFields
    );
    itemId = String(createdItem?.id || "").trim();
    created = true;
  }

  clearOverlayUserCache(userUpn);

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
      tags: overlay?.tags || [],
      activeStartedAt: overlay?.activeStartedAt || null,
      lastWorkedAt: overlay?.lastWorkedAt || null,
      energy: overlay?.energy || "",
      effortMinutes: overlay?.effortMinutes ?? null,
      impact: overlay?.impact || "",
      overlayNotes: overlay?.overlayNotes || "",
      pinned: overlay?.pinned === true,
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
  getUnifiedTasks,
  mapGraphError,
  upsertOverlay,
};
