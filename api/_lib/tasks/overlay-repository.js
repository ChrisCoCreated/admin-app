const { OVERLAY_FIELD_MAP, buildTaskKey, fromOverlayFields } = require("./unified-model");

const OVERLAY_PAGE_SIZE = 200;
const OVERLAY_CACHE_TTL_MS = 45 * 1000;

const siteCache = {
  key: "",
  siteId: "",
  listId: "",
};

const userOverlayCache = new Map();

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function requireSharePointConfig() {
  const siteUrlValue = process.env.SHAREPOINT_SITE_URL;
  const listName = process.env.SHAREPOINT_TASK_OVERLAY_LIST_NAME || "TaskOverlay";

  if (!siteUrlValue) {
    const error = new Error("Missing SHAREPOINT_SITE_URL.");
    error.status = 500;
    error.code = "SHAREPOINT_CONFIG_MISSING";
    throw error;
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    const error = new Error("SHAREPOINT_SITE_URL must include a site path.");
    error.status = 500;
    error.code = "SHAREPOINT_CONFIG_INVALID";
    throw error;
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    listName,
  };
}

async function resolveSiteAndListIds(graphClient) {
  const { hostName, sitePath, listName } = requireSharePointConfig();
  const cacheKey = `${hostName}|${sitePath}|${listName}`;

  if (siteCache.key === cacheKey && siteCache.siteId && siteCache.listId) {
    return {
      siteId: siteCache.siteId,
      listId: siteCache.listId,
    };
  }

  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`;
  const sitePayload = await graphClient.fetchJson(siteUrl);
  const siteId = String(sitePayload?.id || "").trim();

  if (!siteId) {
    const error = new Error("Could not resolve SharePoint site id.");
    error.status = 404;
    error.code = "SHAREPOINT_SITE_NOT_FOUND";
    throw error;
  }

  const params = new URLSearchParams({
    $select: "id,displayName",
    $filter: `displayName eq ${quoteODataString(listName)}`,
    $top: "1",
  });
  const listUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`;
  const listPayload = await graphClient.fetchJson(listUrl);
  const list = Array.isArray(listPayload?.value) ? listPayload.value[0] : null;
  const listId = String(list?.id || "").trim();

  if (!listId) {
    const error = new Error(`Could not find SharePoint list '${listName}'.`);
    error.status = 404;
    error.code = "TASK_OVERLAY_LIST_NOT_FOUND";
    throw error;
  }

  siteCache.key = cacheKey;
  siteCache.siteId = siteId;
  siteCache.listId = listId;

  return { siteId, listId };
}

function userCacheKey(userUpn) {
  return String(userUpn || "").trim().toLowerCase();
}

function pickOverlaySelectableFields() {
  return [
    OVERLAY_FIELD_MAP.userUpn,
    OVERLAY_FIELD_MAP.provider,
    OVERLAY_FIELD_MAP.externalTaskId,
    OVERLAY_FIELD_MAP.workingStatus,
    OVERLAY_FIELD_MAP.workType,
    OVERLAY_FIELD_MAP.tags,
    OVERLAY_FIELD_MAP.activeStartedAt,
    OVERLAY_FIELD_MAP.lastWorkedAt,
    OVERLAY_FIELD_MAP.energy,
    OVERLAY_FIELD_MAP.effortMinutes,
    OVERLAY_FIELD_MAP.impact,
    OVERLAY_FIELD_MAP.overlayNotes,
    OVERLAY_FIELD_MAP.pinned,
    OVERLAY_FIELD_MAP.lastOverlayUpdatedAt,
  ].join(",");
}

function mapOverlays(items) {
  const overlays = [];
  const byKey = new Map();

  for (const item of items) {
    const overlay = fromOverlayFields(item);
    if (!overlay) {
      continue;
    }

    const key = buildTaskKey(overlay.provider, overlay.externalTaskId);
    overlay.key = key;
    overlays.push(overlay);

    if (!byKey.has(key)) {
      byKey.set(key, overlay);
    }
  }

  return { overlays, byKey };
}

async function listOverlaysByUser(graphClient, userUpn) {
  const key = userCacheKey(userUpn);
  const cached = userOverlayCache.get(key);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.value;
  }

  const { siteId, listId } = await resolveSiteAndListIds(graphClient);
  const selectFields = pickOverlaySelectableFields();
  const params = new URLSearchParams({
    $expand: `fields($select=${selectFields})`,
    $filter: `fields/${OVERLAY_FIELD_MAP.userUpn} eq ${quoteODataString(userUpn)}`,
    $top: String(OVERLAY_PAGE_SIZE),
  });

  const initialUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(initialUrl);

  const mapped = mapOverlays(items);
  const value = {
    siteId,
    listId,
    overlays: mapped.overlays,
    byKey: mapped.byKey,
  };

  userOverlayCache.set(key, {
    expiresAt: Date.now() + OVERLAY_CACHE_TTL_MS,
    value,
  });

  return value;
}

function clearOverlayUserCache(userUpn) {
  userOverlayCache.delete(userCacheKey(userUpn));
}

async function patchOverlayFields(graphClient, siteId, listId, itemId, fields) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
  await graphClient.fetchJson(url, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(fields),
  });
}

async function createOverlayItem(graphClient, siteId, listId, fields) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  const created = await graphClient.fetchJson(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields }),
  });

  return created;
}

module.exports = {
  clearOverlayUserCache,
  createOverlayItem,
  listOverlaysByUser,
  patchOverlayFields,
};
