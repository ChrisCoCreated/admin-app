const { OVERLAY_FIELD_MAP, buildTaskKey, fromOverlayFields } = require("./unified-model");

const OVERLAY_PAGE_SIZE = 200;
const GRAPH_TASKS_DEBUG = process.env.GRAPH_TASKS_DEBUG === "1";

const siteCache = {
  key: "",
  siteId: "",
  listId: "",
  fieldMap: null,
};

const userOverlayCache = new Map();

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function logOverlayDebug(message, details) {
  if (!GRAPH_TASKS_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[overlay-repo] ${message}`, details);
    return;
  }
  console.log(`[overlay-repo] ${message}`);
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

function normalizeToken(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

async function resolveListColumns(graphClient, siteId, listId) {
  const columns = [];
  let nextUrl =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}` +
    `/columns?$select=name,displayName&$top=200`;

  while (nextUrl) {
    const payload = await graphClient.fetchJson(nextUrl);
    const values = Array.isArray(payload?.value) ? payload.value : [];
    columns.push(...values);
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }

  return columns;
}

function resolveOverlayFieldMap(columns) {
  const resolved = {};
  const byNameToken = new Map();
  const byDisplayToken = new Map();

  for (const column of columns) {
    const name = String(column?.name || "").trim();
    const displayName = String(column?.displayName || "").trim();
    if (!name) {
      continue;
    }

    const nameToken = normalizeToken(name);
    const displayToken = normalizeToken(displayName);
    if (nameToken && !byNameToken.has(nameToken)) {
      byNameToken.set(nameToken, name);
    }
    if (displayToken && !byDisplayToken.has(displayToken)) {
      byDisplayToken.set(displayToken, name);
    }
  }

  for (const [logicalKey, expectedFieldName] of Object.entries(OVERLAY_FIELD_MAP)) {
    const token = normalizeToken(expectedFieldName);
    resolved[logicalKey] =
      byDisplayToken.get(token) ||
      byNameToken.get(token) ||
      expectedFieldName;
  }

  return resolved;
}

async function resolveSiteAndListIds(graphClient) {
  const { hostName, sitePath, listName } = requireSharePointConfig();
  const cacheKey = `${hostName}|${sitePath}|${listName}`;

  if (siteCache.key === cacheKey && siteCache.siteId && siteCache.listId && siteCache.fieldMap) {
    return {
      siteId: siteCache.siteId,
      listId: siteCache.listId,
      fieldMap: siteCache.fieldMap,
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

  const columns = await resolveListColumns(graphClient, siteId, listId);
  const fieldMap = resolveOverlayFieldMap(columns);

  siteCache.key = cacheKey;
  siteCache.siteId = siteId;
  siteCache.listId = listId;
  siteCache.fieldMap = fieldMap;

  logOverlayDebug("Resolved TaskOverlay field map.", fieldMap);

  return { siteId, listId, fieldMap };
}

function userCacheKey(userUpn) {
  return String(userUpn || "").trim().toLowerCase();
}

function normalizeMatchToken(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function overlayMatchesUser(overlayUserUpn, currentUserUpn) {
  const overlayToken = normalizeMatchToken(overlayUserUpn);
  const currentToken = normalizeMatchToken(currentUserUpn);
  if (!overlayToken || !currentToken) {
    return false;
  }

  if (overlayToken === currentToken) {
    return true;
  }

  const currentLocal = currentToken.split("@")[0] || "";
  const overlayLocal = overlayToken.split("@")[0] || "";
  if (!currentLocal || !overlayLocal) {
    return false;
  }

  // Person fields sometimes resolve to display-like text instead of full UPN.
  if (overlayToken === currentLocal || overlayLocal === currentLocal) {
    return true;
  }
  if (overlayToken.includes(currentLocal) || currentLocal.includes(overlayToken)) {
    return true;
  }

  return false;
}

function getOverlayCacheTtlMs() {
  const configured = Number(process.env.TASKS_OVERLAY_CACHE_TTL_MS || 120000);
  if (!Number.isFinite(configured) || configured < 0) {
    return 120000;
  }
  return Math.floor(configured);
}

function mapOverlays(items, fieldMap) {
  const overlays = [];
  const byKey = new Map();

  for (const item of items) {
    const rawFields = item?.fields || {};
    const normalizedFields = { ...rawFields };
    for (const [logicalKey, canonicalDisplayName] of Object.entries(OVERLAY_FIELD_MAP)) {
      const actualFieldName = fieldMap?.[logicalKey] || canonicalDisplayName;
      if (
        Object.prototype.hasOwnProperty.call(rawFields, actualFieldName) &&
        !Object.prototype.hasOwnProperty.call(normalizedFields, canonicalDisplayName)
      ) {
        normalizedFields[canonicalDisplayName] = rawFields[actualFieldName];
      }
    }

    const overlay = fromOverlayFields({
      ...item,
      fields: normalizedFields,
    });
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

function inferLookupIdFromItems(items, userUpnField) {
  const lookupField = `${userUpnField}LookupId`;
  const ids = new Set();
  for (const item of items) {
    const raw = item?.fields?.[lookupField];
    const n = Number(raw);
    if (Number.isFinite(n) && n > 0) {
      ids.add(n);
    }
  }
  if (ids.size === 1) {
    return Array.from(ids)[0];
  }
  return null;
}

function buildOverlayRowDiagnostics({ items, mappedOverlays, userUpnField, requestedUserUpn }) {
  const requested = String(requestedUserUpn || "").trim().toLowerCase();
  const mappedById = new Map();
  for (const overlay of mappedOverlays) {
    mappedById.set(String(overlay?.itemId || ""), overlay);
  }

  return items.map((item) => {
    const itemId = String(item?.id || "");
    const fields = item?.fields || {};
    const mapped = mappedById.get(itemId) || null;
    const normalizedUser = String(mapped?.userUpn || "").trim().toLowerCase();
    const matchResult = overlayMatchesUser(normalizedUser, requested);
    return {
      itemId,
      title: String(fields?.Title || mapped?.title || "").trim(),
      provider: String(fields?.Provider || fields?.field_1 || mapped?.provider || "").trim().toLowerCase(),
      externalTaskId: String(fields?.ExternalTaskId || fields?.field_2 || mapped?.externalTaskId || "").trim(),
      rawUserField: fields?.[userUpnField],
      rawUserLookupId: fields?.[`${userUpnField}LookupId`],
      normalizedUserToken: normalizedUser,
      matchResult,
    };
  });
}

async function listOverlaysByUser(graphClient, userUpn) {
  const key = userCacheKey(userUpn);
  const cached = userOverlayCache.get(key);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.value;
  }

  const startRefresh = async () => {
    const current = userOverlayCache.get(key);
    if (current?.inFlight) {
      return current.inFlight;
    }

    const inFlight = (async () => {
      const { siteId, listId, fieldMap } = await resolveSiteAndListIds(graphClient);
      const userUpnField = fieldMap?.userUpn || OVERLAY_FIELD_MAP.userUpn;
      const params = new URLSearchParams({
        $expand: "fields",
        $filter: `fields/${userUpnField} eq ${quoteODataString(userUpn)}`,
        $top: String(OVERLAY_PAGE_SIZE),
      });

      const initialUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
      let items = [];

      try {
        items = await graphClient.fetchAllPages(initialUrl);
      } catch (error) {
        const code = String(error?.code || "").toLowerCase();
        const isFilterSyntaxIssue =
          Number(error?.status) === 400 && (code.includes("invalidrequest") || code.includes("badrequest"));

        if (!isFilterSyntaxIssue) {
          throw error;
        }

        // Fallback for Person/Group field filtering quirks in Graph list-item queries.
        logOverlayDebug("UserUPN filtered query failed; falling back to in-memory user filter.", {
          status: error?.status || 0,
          code: error?.code || "",
          message: error?.message || "",
        });

        const fallbackParams = new URLSearchParams({
          $expand: "fields",
          $top: String(OVERLAY_PAGE_SIZE),
        });
        const fallbackUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${fallbackParams.toString()}`;
        items = await graphClient.fetchAllPages(fallbackUrl, { maxPages: 25 });
      }

      if (items.length === 0) {
        // Person/Group equality filters can return no rows even when records exist.
        logOverlayDebug("UserUPN filtered query returned zero rows; retrying with unfiltered scan.", {
          requestedUserUpn: String(userUpn || "").trim().toLowerCase(),
          userUpnField,
        });
        const fallbackParams = new URLSearchParams({
          $expand: "fields",
          $top: String(OVERLAY_PAGE_SIZE),
        });
        const fallbackUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${fallbackParams.toString()}`;
        items = await graphClient.fetchAllPages(fallbackUrl, { maxPages: 25 });
      }

      const mapped = mapOverlays(items, fieldMap);
      let filteredOverlays = mapped.overlays.filter((overlay) => {
        return overlayMatchesUser(overlay?.userUpn || "", userUpn);
      });
      const inferredLookupId = inferLookupIdFromItems(items, userUpnField);

      if (filteredOverlays.length === 0 && Number.isFinite(inferredLookupId) && inferredLookupId > 0) {
        const lookupByItemId = new Map();
        const lookupField = `${userUpnField}LookupId`;
        for (const item of items) {
          const itemId = String(item?.id || "").trim();
          if (!itemId) {
            continue;
          }
          const lookupId = Number(item?.fields?.[lookupField]);
          if (Number.isFinite(lookupId) && lookupId > 0) {
            lookupByItemId.set(itemId, lookupId);
          }
        }

        const lookupMatched = mapped.overlays.filter((overlay) => {
          const itemId = String(overlay?.itemId || "").trim();
          return lookupByItemId.get(itemId) === inferredLookupId;
        });

        if (lookupMatched.length > 0) {
          logOverlayDebug("Using UserUPN lookup-id fallback match for overlays.", {
            requestedUserUpn: String(userUpn || "").trim().toLowerCase(),
            userUpnField,
            inferredLookupId,
            matchedRowsCount: lookupMatched.length,
          });
          filteredOverlays = lookupMatched;
        }
      }

      if (filteredOverlays.length === 0 && mapped.overlays.length > 0) {
        const rowDiagnostics = buildOverlayRowDiagnostics({
          items,
          mappedOverlays: mapped.overlays,
          userUpnField,
          requestedUserUpn: userUpn,
        });
        logOverlayDebug("No strict user match for overlays.", {
          requestedUserUpn: String(userUpn || "").trim().toLowerCase(),
          userUpnField,
          totalRowsLoaded: mapped.overlays.length,
          matchedRowsCount: 0,
          rowDiagnostics,
        });
      }
      const byKey = new Map();
      for (const overlay of filteredOverlays) {
        if (!byKey.has(overlay.key)) {
          byKey.set(overlay.key, overlay);
        }
      }

      const value = {
        siteId,
        listId,
        fieldMap,
        userLookupId: inferredLookupId,
        overlays: filteredOverlays,
        byKey,
      };

      userOverlayCache.set(key, {
        expiresAt: Date.now() + getOverlayCacheTtlMs(),
        value,
        inFlight: null,
      });

      return value;
    })();

    userOverlayCache.set(key, {
      expiresAt: cached?.expiresAt || 0,
      value: cached?.value || null,
      inFlight,
    });

    try {
      return await inFlight;
    } catch (error) {
      const previous = userOverlayCache.get(key);
      userOverlayCache.set(key, {
        expiresAt: previous?.expiresAt || 0,
        value: previous?.value || null,
        inFlight: null,
      });
      throw error;
    }
  };

  if (cached?.value) {
    void startRefresh();
    return cached.value;
  }

  if (cached?.inFlight) {
    return cached.inFlight;
  }
  return startRefresh();
}

function clearOverlayUserCache(userUpn) {
  userOverlayCache.delete(userCacheKey(userUpn));
}

function translateFieldsForWrite(fields, fieldMap) {
  const translated = {};

  for (const [key, value] of Object.entries(fields || {})) {
    if (key.endsWith("LookupId")) {
      const canonicalBase = key.slice(0, -8);
      const logicalEntryByBase = Object.entries(OVERLAY_FIELD_MAP).find(([, canonical]) => canonical === canonicalBase);
      if (logicalEntryByBase) {
        const logicalKey = logicalEntryByBase[0];
        const actualBaseName = fieldMap?.[logicalKey] || canonicalBase;
        translated[`${actualBaseName}LookupId`] = value;
        continue;
      }
      translated[key] = value;
      continue;
    }

    const logicalEntry = Object.entries(OVERLAY_FIELD_MAP).find(([, canonical]) => canonical === key);
    if (!logicalEntry) {
      translated[key] = value;
      continue;
    }

    const logicalKey = logicalEntry[0];
    const actualName = fieldMap?.[logicalKey] || key;
    translated[actualName] = value;
  }

  return translated;
}

async function patchOverlayFields(graphClient, siteId, listId, itemId, fields, fieldMap) {
  const translatedFields = translateFieldsForWrite(fields, fieldMap);
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
  await graphClient.fetchJson(url, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(translatedFields),
  });
}

function getOverlayPatchConcurrency() {
  const configured = Number(process.env.TASKS_OVERLAY_PATCH_CONCURRENCY || 4);
  if (!Number.isFinite(configured) || configured < 1) {
    return 4;
  }
  return Math.min(Math.floor(configured), 12);
}

async function patchOverlayFieldsBatch(graphClient, siteId, listId, patches, fieldMap, concurrency) {
  const queue = Array.isArray(patches) ? patches.filter(Boolean) : [];
  if (!queue.length) {
    return [];
  }

  const limit = Number.isFinite(concurrency) && concurrency > 0
    ? Math.min(Math.floor(concurrency), 12)
    : getOverlayPatchConcurrency();

  const results = new Array(queue.length);
  let nextIndex = 0;
  let active = 0;
  let completed = 0;

  return new Promise((resolve) => {
    const launch = () => {
      if (completed >= queue.length) {
        resolve(results);
        return;
      }

      while (active < limit && nextIndex < queue.length) {
        const currentIndex = nextIndex;
        nextIndex += 1;
        active += 1;
        const entry = queue[currentIndex];

        Promise.resolve()
          .then(async () => {
            await patchOverlayFields(
              graphClient,
              siteId,
              listId,
              String(entry?.itemId || "").trim(),
              entry?.fields || {},
              fieldMap
            );
            results[currentIndex] = { ok: true };
          })
          .catch((error) => {
            results[currentIndex] = { ok: false, error };
          })
          .finally(() => {
            active -= 1;
            completed += 1;
            launch();
          });
      }
    };

    launch();
  });
}

async function createOverlayItem(graphClient, siteId, listId, fields, fieldMap) {
  const translatedFields = translateFieldsForWrite(fields, fieldMap);
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  const created = await graphClient.fetchJson(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields: translatedFields }),
  });

  return created;
}

async function backfillMissingOverlayUsers(graphClient, userUpn) {
  const { siteId, listId, fieldMap } = await resolveSiteAndListIds(graphClient);
  const userUpnField = fieldMap?.userUpn || OVERLAY_FIELD_MAP.userUpn;
  const userUpnLookupField = `${userUpnField}LookupId`;
  const fallbackParams = new URLSearchParams({
    $expand: "fields",
    $top: String(OVERLAY_PAGE_SIZE),
  });
  const fallbackUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${fallbackParams.toString()}`;
  const items = await graphClient.fetchAllPages(fallbackUrl, { maxPages: 25 });

  const inferredLookupId = inferLookupIdFromItems(items, userUpnField);
  let updated = 0;
  let skipped = 0;

  for (const item of items) {
    const itemId = String(item?.id || "").trim();
    if (!itemId) {
      continue;
    }
    const fields = item?.fields || {};
    const hasUserValue = Boolean(fields[userUpnField]);
    const hasLookupId = Number.isFinite(Number(fields[userUpnLookupField])) && Number(fields[userUpnLookupField]) > 0;
    if (hasUserValue || hasLookupId) {
      skipped += 1;
      continue;
    }

    const patch = {
      [userUpnField]: String(userUpn || "").trim().toLowerCase(),
    };
    if (Number.isFinite(inferredLookupId) && inferredLookupId > 0) {
      patch[userUpnLookupField] = inferredLookupId;
    }

    await patchOverlayFields(graphClient, siteId, listId, itemId, patch, fieldMap);
    updated += 1;
  }

  if (updated > 0) {
    logOverlayDebug("Backfilled missing TaskOverlay UserUPN values.", {
      requestedUserUpn: String(userUpn || "").trim().toLowerCase(),
      userUpnField,
      inferredLookupId,
      updatedRows: updated,
      skippedRows: skipped,
    });
  }

  return {
    updatedRows: updated,
    skippedRows: skipped,
    inferredLookupId,
  };
}

async function listPinnedOverlaysByUser(graphClient, userUpn) {
  const bundle = await listOverlaysByUser(graphClient, userUpn);
  const overlays = Array.isArray(bundle?.overlays) ? bundle.overlays : [];
  const totalRows = overlays.length;
  const pinnedOverlays = overlays.filter((overlay) => overlay?.pinned === true);
  const byKey = new Map();
  for (const overlay of pinnedOverlays) {
    if (!byKey.has(overlay.key)) {
      byKey.set(overlay.key, overlay);
    }
  }
  return {
    ...bundle,
    totalRows,
    allOverlays: overlays,
    overlays: pinnedOverlays,
    byKey,
  };
}

module.exports = {
  backfillMissingOverlayUsers,
  clearOverlayUserCache,
  createOverlayItem,
  listPinnedOverlaysByUser,
  listOverlaysByUser,
  patchOverlayFieldsBatch,
  patchOverlayFields,
};
