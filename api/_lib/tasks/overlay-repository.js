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
      const filteredOverlays = mapped.overlays.filter((overlay) => {
        return overlayMatchesUser(overlay?.userUpn || "", userUpn);
      });

      if (filteredOverlays.length === 0 && mapped.overlays.length > 0) {
        const sampleUsers = mapped.overlays
          .map((overlay) => String(overlay?.userUpn || "").trim())
          .filter(Boolean)
          .slice(0, 6);
        const rawSamples = items
          .map((item) => item?.fields?.[userUpnField])
          .filter((entry) => entry !== undefined && entry !== null)
          .slice(0, 6);
        logOverlayDebug("No strict user match for overlays.", {
          requestedUserUpn: String(userUpn || "").trim().toLowerCase(),
          userUpnField,
          overlayRowsLoaded: mapped.overlays.length,
          sampleOverlayUserValues: sampleUsers,
          sampleRawUserFieldValues: rawSamples,
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

module.exports = {
  clearOverlayUserCache,
  createOverlayItem,
  listOverlaysByUser,
  patchOverlayFields,
};
