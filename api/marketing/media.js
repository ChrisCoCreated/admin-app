const crypto = require("crypto");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const { requireApiAuth } = require("../_lib/require-api-auth");

const MARKETING_MEDIA_DEBUG = process.env.MARKETING_MEDIA_DEBUG === "1";
const STORE_DIR = process.env.MARKETING_MEDIA_STORE_DIR
  ? path.resolve(process.env.MARKETING_MEDIA_STORE_DIR)
  : path.join(process.cwd(), "data", "marketing-media");
const STORE_FILES_DIR = path.join(STORE_DIR, "files");
const STORE_INDEX_PATH = path.join(STORE_DIR, "index.json");
const PRUNE_INTERVAL_MS = 5 * 60 * 1000;

let storeReadyPromise = null;
let indexWriteLock = Promise.resolve();
let lastPruneAt = 0;

function logMediaDebug(message, details) {
  if (!MARKETING_MEDIA_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[marketing-media] ${message}`, details);
    return;
  }
  console.log(`[marketing-media] ${message}`);
}

function getSharePointSiteParts() {
  const raw = String(process.env.SHAREPOINT_SITE_URL || "").trim();
  if (!raw) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }
  const siteUrl = new URL(raw);
  return {
    hostName: siteUrl.hostname,
    sitePath: siteUrl.pathname.replace(/\/$/, ""),
  };
}

function getAllowedSharePointPrefix() {
  const { hostName, sitePath } = getSharePointSiteParts();
  return `https://${hostName}${sitePath}/`;
}

function quoteODataString(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
}

function decodeLoose(value) {
  let output = String(value || "").replace(/\+/g, " ");
  for (let index = 0; index < 3; index += 1) {
    try {
      const next = decodeURIComponent(output);
      if (next === output) {
        break;
      }
      output = next;
    } catch {
      break;
    }
  }
  return output;
}

function parseAttachmentPath(targetUrl, sitePath) {
  const decodedPath = decodeLoose(String(targetUrl.pathname || ""));
  if (!decodedPath.startsWith(sitePath)) {
    return null;
  }
  const relative = decodedPath.slice(sitePath.length).replace(/^\/+/, "");
  const match = /^Lists\/(.+)\/Attachments\/(\d+)\/([^/]+)$/i.exec(relative);
  if (!match) {
    return null;
  }
  return {
    listName: decodeLoose(match[1]).trim(),
    itemId: Number(match[2]),
    fileName: decodeLoose(match[3]).trim(),
    normalizedPath: relative,
  };
}

function getIncomingBearerToken(req) {
  const authHeader = String(req.headers.authorization || "");
  const match = /^Bearer\s+(.+)$/i.exec(authHeader);
  return match ? String(match[1] || "").trim() : "";
}

function getPruneDays() {
  const configured = Number(process.env.MARKETING_MEDIA_PRUNE_DAYS || "0");
  if (!Number.isFinite(configured) || configured <= 0) {
    return 0;
  }
  return Math.floor(configured);
}

function getMaxBytes() {
  const configured = Number(process.env.MARKETING_MEDIA_MAX_BYTES || "0");
  if (!Number.isFinite(configured) || configured <= 0) {
    return 0;
  }
  return Math.floor(configured);
}

async function ensureStoreReady() {
  if (!storeReadyPromise) {
    storeReadyPromise = (async () => {
      await fsp.mkdir(STORE_FILES_DIR, { recursive: true });
      try {
        await fsp.access(STORE_INDEX_PATH, fs.constants.F_OK);
      } catch {
        const initial = { entries: {} };
        await fsp.writeFile(STORE_INDEX_PATH, JSON.stringify(initial, null, 2), "utf8");
      }
    })();
  }
  await storeReadyPromise;
}

async function readIndex() {
  await ensureStoreReady();
  try {
    const raw = await fsp.readFile(STORE_INDEX_PATH, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || typeof parsed.entries !== "object") {
      return { entries: {} };
    }
    return {
      entries: parsed.entries,
    };
  } catch {
    return { entries: {} };
  }
}

function atomicJsonWrite(filePath, payload) {
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  const body = JSON.stringify(payload, null, 2);
  return fsp
    .writeFile(tmpPath, body, "utf8")
    .then(() => fsp.rename(tmpPath, filePath))
    .catch(async (error) => {
      try {
        await fsp.unlink(tmpPath);
      } catch {
        // ignore temp cleanup failure
      }
      throw error;
    });
}

function withIndexLock(action) {
  const next = indexWriteLock.then(() => action());
  indexWriteLock = next.catch(() => undefined);
  return next;
}

function normalizeSourceUrl(targetUrl) {
  const normalized = new URL(targetUrl.href);
  normalized.hash = "";
  normalized.hostname = normalized.hostname.toLowerCase();
  return normalized.toString();
}

function buildCacheKey(targetUrl, attachment) {
  const normalizedUrl = normalizeSourceUrl(targetUrl);
  const attachmentId = Number.isFinite(attachment?.itemId) ? String(attachment.itemId) : "";
  const digest = crypto
    .createHash("sha256")
    .update(`${normalizedUrl}|${attachmentId}`)
    .digest("hex");
  return {
    key: digest,
    normalizedUrl,
  };
}

async function touchEntry(cacheKey) {
  await withIndexLock(async () => {
    const index = await readIndex();
    const entry = index.entries?.[cacheKey];
    if (!entry) {
      return;
    }
    entry.lastAccessedAt = new Date().toISOString();
    await atomicJsonWrite(STORE_INDEX_PATH, index);
  });
}

async function getCachedMedia(cacheKey) {
  const index = await readIndex();
  const entry = index.entries?.[cacheKey];
  if (!entry || !entry.fileName) {
    return null;
  }
  const filePath = path.join(STORE_FILES_DIR, entry.fileName);
  try {
    const buffer = await fsp.readFile(filePath);
    void touchEntry(cacheKey);
    return {
      buffer,
      mimeType: String(entry.contentType || "application/octet-stream"),
    };
  } catch {
    return null;
  }
}

async function persistCachedMedia(cacheKey, sourceUrl, buffer, contentType) {
  await ensureStoreReady();
  const safeType = String(contentType || "application/octet-stream").split(";")[0].trim();
  const extByType = {
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/png": "png",
    "image/webp": "webp",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "video/mp4": "mp4",
    "video/webm": "webm",
    "video/quicktime": "mov",
  };
  const ext = extByType[safeType.toLowerCase()] || "bin";
  const fileName = `${cacheKey}.${ext}`;
  const filePath = path.join(STORE_FILES_DIR, fileName);
  const tmpFilePath = `${filePath}.${process.pid}.${Date.now()}.tmp`;

  try {
    await fsp.writeFile(tmpFilePath, buffer);
    await fsp.rename(tmpFilePath, filePath);
  } catch (error) {
    try {
      await fsp.unlink(tmpFilePath);
    } catch {
      // ignore
    }
    throw error;
  }

  const nowIso = new Date().toISOString();
  await withIndexLock(async () => {
    const index = await readIndex();
    index.entries[cacheKey] = {
      sourceUrl,
      contentType: safeType,
      size: buffer.length,
      fileName,
      createdAt: index.entries?.[cacheKey]?.createdAt || nowIso,
      lastAccessedAt: nowIso,
    };
    await atomicJsonWrite(STORE_INDEX_PATH, index);
  });
}

async function pruneStoreIfNeeded() {
  const now = Date.now();
  if (now - lastPruneAt < PRUNE_INTERVAL_MS) {
    return;
  }
  lastPruneAt = now;

  const pruneDays = getPruneDays();
  const maxBytes = getMaxBytes();
  if (!pruneDays && !maxBytes) {
    return;
  }

  await withIndexLock(async () => {
    const index = await readIndex();
    const entries = Object.entries(index.entries || {}).map(([key, value]) => ({
      key,
      ...value,
      lastAccessedMs: Date.parse(value?.lastAccessedAt || value?.createdAt || 0) || 0,
      createdMs: Date.parse(value?.createdAt || 0) || 0,
      size: Number(value?.size) || 0,
    }));

    const keysToRemove = new Set();

    if (pruneDays > 0) {
      const minAccessMs = now - pruneDays * 24 * 60 * 60 * 1000;
      for (const entry of entries) {
        const markMs = entry.lastAccessedMs || entry.createdMs;
        if (markMs > 0 && markMs < minAccessMs) {
          keysToRemove.add(entry.key);
        }
      }
    }

    let totalBytes = entries
      .filter((entry) => !keysToRemove.has(entry.key))
      .reduce((sum, entry) => sum + entry.size, 0);

    if (maxBytes > 0 && totalBytes > maxBytes) {
      const byOldest = entries
        .filter((entry) => !keysToRemove.has(entry.key))
        .sort((a, b) => (a.lastAccessedMs || a.createdMs) - (b.lastAccessedMs || b.createdMs));
      for (const entry of byOldest) {
        if (totalBytes <= maxBytes) {
          break;
        }
        keysToRemove.add(entry.key);
        totalBytes -= entry.size;
      }
    }

    if (!keysToRemove.size) {
      return;
    }

    for (const key of keysToRemove) {
      const fileName = String(index.entries?.[key]?.fileName || "").trim();
      if (fileName) {
        const filePath = path.join(STORE_FILES_DIR, fileName);
        try {
          await fsp.unlink(filePath);
        } catch {
          // ignore missing file
        }
      }
      delete index.entries[key];
    }

    await atomicJsonWrite(STORE_INDEX_PATH, index);
    logMediaDebug("cache-prune", {
      removed: keysToRemove.size,
      pruneDays,
      maxBytes,
    });
  });
}

async function exchangeDelegatedToken(assertion, scope, label) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_API_CLIENT_ID;
  const clientSecret = process.env.AZURE_API_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    requested_token_use: "on_behalf_of",
    assertion: String(assertion || ""),
    scope: String(scope || ""),
  });
  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }
  if (!response.ok || !payload?.access_token) {
    throw new Error(payload?.error_description || payload?.error || `${label} token exchange failed.`);
  }
  return payload.access_token;
}

function buildTokenContext(req) {
  const incomingToken = getIncomingBearerToken(req);
  if (!incomingToken) {
    throw new Error("Missing incoming API bearer token.");
  }
  const claims = req.authUser?.claims || {};
  logMediaDebug("incoming-api-token-aud", {
    aud: String(claims?.aud || ""),
    iss: String(claims?.iss || ""),
    scp: String(claims?.scp || ""),
  });
  return {
    incomingToken,
    graphTokenPromise: null,
    sharePointTokenPromise: null,
  };
}

async function getDelegatedSharePointToken(ctx, hostName) {
  if (!ctx.sharePointTokenPromise) {
    logMediaDebug("obo-sharepoint-token-start", { scope: `https://${hostName}/.default` });
    ctx.sharePointTokenPromise = exchangeDelegatedToken(
      ctx.incomingToken,
      `https://${hostName}/.default`,
      "SharePoint OBO"
    ).then((token) => {
      logMediaDebug("obo-sharepoint-token-success");
      return token;
    });
  }
  return ctx.sharePointTokenPromise;
}

async function getDelegatedGraphToken(ctx) {
  if (!ctx.graphTokenPromise) {
    logMediaDebug("obo-graph-token-start", { scope: "https://graph.microsoft.com/.default" });
    ctx.graphTokenPromise = exchangeDelegatedToken(
      ctx.incomingToken,
      "https://graph.microsoft.com/.default",
      "Graph OBO"
    ).then((token) => {
      logMediaDebug("obo-graph-token-success");
      return token;
    });
  }
  return ctx.graphTokenPromise;
}

async function fetchGraphJson(url, token) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }
  if (!response.ok) {
    throw new Error(payload?.error?.message || text || `Graph request failed (${response.status}).`);
  }
  return payload;
}

async function fetchGraphBinary(url, token, label) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`${label} failed (${response.status}): ${detail.slice(0, 300)}`);
  }
  return response;
}

async function fetchGraphSiteId(token, hostName, sitePath) {
  const payload = await fetchGraphJson(
    `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`,
    token
  );
  if (!payload?.id) {
    throw new Error("Could not resolve Graph site id.");
  }
  return payload.id;
}

async function fetchGraphListIdByName(token, siteId, listName) {
  const normalizeListKey = (value) =>
    decodeLoose(String(value || ""))
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  const listPathTail = (webUrl) => {
    try {
      const pathname = decodeLoose(new URL(String(webUrl || "")).pathname || "");
      const parts = pathname.split("/").filter(Boolean);
      return parts.length ? parts[parts.length - 1] : "";
    } catch {
      return "";
    }
  };

  const wanted = normalizeListKey(listName);
  const rows = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName,webUrl&$top=200`;
  while (nextUrl) {
    const payload = await fetchGraphJson(nextUrl, token);
    rows.push(...(Array.isArray(payload?.value) ? payload.value : []));
    nextUrl = String(payload?.["@odata.nextLink"] || "").trim();
  }
  const match = rows.find((row) => {
    const display = normalizeListKey(row?.displayName);
    const tail = normalizeListKey(listPathTail(row?.webUrl));
    return display === wanted || tail === wanted;
  });
  if (!match?.id) {
    throw new Error(`Could not find list '${listName}'.`);
  }
  return match.id;
}

async function fetchGraphListDriveId(token, siteId, listId) {
  const payload = await fetchGraphJson(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/drive?$select=id,webUrl`,
    token
  );
  if (!payload?.id) {
    throw new Error("Could not resolve list drive id.");
  }
  return {
    id: payload.id,
    webUrl: String(payload.webUrl || "").trim(),
  };
}

async function fetchViaDelegatedSharePointAttachment(targetUrl, ctx) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const parsed = parseAttachmentPath(targetUrl, sitePath);
  if (!parsed || !parsed.listName || !Number.isFinite(parsed.itemId)) {
    return null;
  }

  logMediaDebug("delegated-sharepoint-obo-start", {
    listName: parsed.listName,
    itemId: parsed.itemId,
    fileName: parsed.fileName,
  });

  const spToken = await getDelegatedSharePointToken(ctx, hostName);
  const attachmentsUrl = `https://${hostName}${sitePath}/_api/web/lists/GetByTitle(${quoteODataString(
    parsed.listName
  )})/items(${parsed.itemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`;
  const metaResponse = await fetch(attachmentsUrl, {
    headers: {
      Authorization: `Bearer ${spToken}`,
      Accept: "application/json;odata=nometadata",
    },
  });
  const metaText = await metaResponse.text();
  let metaPayload = null;
  try {
    metaPayload = metaText ? JSON.parse(metaText) : null;
  } catch {
    metaPayload = null;
  }
  if (!metaResponse.ok) {
    throw new Error(`SharePoint attachment metadata failed (${metaResponse.status}).`);
  }

  const rows = Array.isArray(metaPayload?.value)
    ? metaPayload.value
    : Array.isArray(metaPayload?.d?.results)
      ? metaPayload.d.results
      : [];
  if (!rows.length) {
    throw new Error("No attachment rows returned from SharePoint.");
  }

  const normalizeName = (value) => decodeLoose(String(value || "")).trim().toLowerCase();
  const requested = normalizeName(parsed.fileName);
  const exact = rows.find((row) => normalizeName(row?.FileName) === requested);
  const picked = exact || rows[0];
  const serverRelative = String(picked?.ServerRelativeUrl || "").trim();
  if (!serverRelative) {
    throw new Error("Attachment row missing ServerRelativeUrl.");
  }

  const absolute = `https://${hostName}${serverRelative.startsWith("/") ? serverRelative : `/${serverRelative}`}`;
  logMediaDebug("delegated-sharepoint-obo-content-url", { absolute });
  const fileResponse = await fetch(absolute, {
    headers: {
      Authorization: `Bearer ${spToken}`,
    },
  });
  if (!fileResponse.ok) {
    const detail = await fileResponse.text();
    throw new Error(`SharePoint attachment content failed (${fileResponse.status}): ${detail.slice(0, 200)}`);
  }
  logMediaDebug("delegated-sharepoint-obo-success", { status: fileResponse.status });
  return fileResponse;
}

function toGraphDrivePathForAttachments(targetUrl, sitePath) {
  const decodedPath = decodeLoose(String(targetUrl.pathname || ""));
  if (!decodedPath.startsWith(sitePath)) {
    return "";
  }
  const relative = decodedPath.slice(sitePath.length).replace(/^\/+/, "");
  if (!/^Lists\/.+\/Attachments\/\d+\/.+/i.test(relative)) {
    return "";
  }
  return relative;
}

async function fetchViaDelegatedListDrive(targetUrl, ctx) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const parsed = parseAttachmentPath(targetUrl, sitePath);
  if (!parsed || !parsed.listName || !parsed.fileName || !Number.isFinite(parsed.itemId)) {
    return null;
  }

  const graphToken = await getDelegatedGraphToken(ctx);
  logMediaDebug("delegated-list-drive-start", {
    listName: parsed.listName,
    itemId: parsed.itemId,
    fileName: parsed.fileName,
  });
  const siteId = await fetchGraphSiteId(graphToken, hostName, sitePath);
  const listId = await fetchGraphListIdByName(graphToken, siteId, parsed.listName);
  const drive = await fetchGraphListDriveId(graphToken, siteId, listId);
  const encodedSegments = ["Attachments", String(parsed.itemId), parsed.fileName]
    .map((part) => encodeURIComponent(part))
    .join("/");
  const contentUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodedSegments}:/content`;
  logMediaDebug("delegated-list-drive-content-url", { contentUrl, driveWebUrl: drive.webUrl });
  const response = await fetchGraphBinary(contentUrl, graphToken, "Delegated list drive fetch");
  logMediaDebug("delegated-list-drive-success", { status: response.status });
  return response;
}

async function fetchViaDelegatedSiteDrive(targetUrl, ctx) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const graphDrivePath = toGraphDrivePathForAttachments(targetUrl, sitePath);
  if (!graphDrivePath) {
    return null;
  }

  const graphToken = await getDelegatedGraphToken(ctx);
  logMediaDebug("delegated-site-drive-start", { graphDrivePath });
  const siteId = await fetchGraphSiteId(graphToken, hostName, sitePath);
  const encodedPath = graphDrivePath
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
  const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/content`;
  logMediaDebug("delegated-site-drive-content-url", { contentUrl });
  const response = await fetchGraphBinary(contentUrl, graphToken, "Delegated site drive fetch");
  logMediaDebug("delegated-site-drive-success", { status: response.status });
  return response;
}

function toGraphShareTokenFromUrl(absoluteUrl) {
  const base64 = Buffer.from(String(absoluteUrl || ""), "utf8").toString("base64");
  const base64Url = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${base64Url}`;
}

async function fetchViaDelegatedShares(targetUrl, ctx) {
  const graphToken = await getDelegatedGraphToken(ctx);
  const shareToken = toGraphShareTokenFromUrl(targetUrl.href);
  const contentUrl = `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/content`;
  logMediaDebug("delegated-shares-start", { contentUrl });
  const response = await fetchGraphBinary(contentUrl, graphToken, "Delegated shares fetch");
  logMediaDebug("delegated-shares-success", { status: response.status });
  return response;
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "marketing"] }))) {
    return;
  }

  const rawUrl = String(req.query.url || "").trim();
  if (!rawUrl) {
    res.status(400).json({ error: "Missing url query parameter." });
    return;
  }

  try {
    const targetUrl = new URL(rawUrl);
    const allowedPrefix = getAllowedSharePointPrefix();
    if (!targetUrl.href.startsWith(allowedPrefix)) {
      res.status(403).json({ error: "URL not allowed." });
      return;
    }

    await pruneStoreIfNeeded();

    const { sitePath } = getSharePointSiteParts();
    const parsedAttachment = parseAttachmentPath(targetUrl, sitePath);
    if (parsedAttachment) {
      logMediaDebug("normalized-attachment-tuple", {
        listName: parsedAttachment.listName,
        itemId: parsedAttachment.itemId,
        fileName: parsedAttachment.fileName,
      });
    }

    const cacheIdentity = buildCacheKey(targetUrl, parsedAttachment);
    const cached = await getCachedMedia(cacheIdentity.key);
    if (cached) {
      logMediaDebug("cache-hit", { key: cacheIdentity.key, sourceUrl: cacheIdentity.normalizedUrl });
      const dataBase64 = cached.buffer.toString("base64");
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        mimeType: cached.mimeType,
        dataBase64,
      });
      return;
    }

    logMediaDebug("cache-miss", { key: cacheIdentity.key, sourceUrl: cacheIdentity.normalizedUrl });

    const ctx = buildTokenContext(req);
    const resolverErrors = [];
    let upstream = null;

    try {
      upstream = await fetchViaDelegatedSharePointAttachment(targetUrl, ctx);
    } catch (error) {
      resolverErrors.push(`sharepoint-obo:${error?.message || String(error)}`);
      logMediaDebug("delegated-sharepoint-obo-error", { detail: error?.message || String(error) });
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedListDrive(targetUrl, ctx);
      } catch (error) {
        resolverErrors.push(`list-drive:${error?.message || String(error)}`);
        logMediaDebug("delegated-list-drive-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedSiteDrive(targetUrl, ctx);
      } catch (error) {
        resolverErrors.push(`site-drive:${error?.message || String(error)}`);
        logMediaDebug("delegated-site-drive-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedShares(targetUrl, ctx);
      } catch (error) {
        resolverErrors.push(`shares:${error?.message || String(error)}`);
        logMediaDebug("delegated-shares-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      const normalizedPath = parsedAttachment?.normalizedPath || decodeLoose(String(targetUrl.pathname || ""));
      res.status(401).json({
        error: "Upstream request failed.",
        detail: `All delegated resolvers failed for '${normalizedPath}'. ${resolverErrors.join(" | ")}`.slice(
          0,
          500
        ),
      });
      return;
    }

    const mimeType = upstream.headers.get("content-type") || "application/octet-stream";
    const arrayBuffer = await upstream.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    try {
      await persistCachedMedia(cacheIdentity.key, cacheIdentity.normalizedUrl, buffer, mimeType);
      logMediaDebug("cache-write-success", {
        key: cacheIdentity.key,
        sourceUrl: cacheIdentity.normalizedUrl,
        bytes: buffer.length,
      });
    } catch (cacheWriteError) {
      logMediaDebug("cache-write-error", {
        key: cacheIdentity.key,
        detail: cacheWriteError?.message || String(cacheWriteError),
      });
    }

    const dataBase64 = buffer.toString("base64");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      mimeType,
      dataBase64,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
