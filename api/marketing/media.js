const { requireApiAuth } = require("../_lib/require-api-auth");

const MARKETING_MEDIA_DEBUG = process.env.MARKETING_MEDIA_DEBUG === "1";

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

    const { sitePath } = getSharePointSiteParts();
    const parsedAttachment = parseAttachmentPath(targetUrl, sitePath);
    if (parsedAttachment) {
      logMediaDebug("normalized-attachment-tuple", {
        listName: parsedAttachment.listName,
        itemId: parsedAttachment.itemId,
        fileName: parsedAttachment.fileName,
      });
    }

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
    const dataBase64 = Buffer.from(arrayBuffer).toString("base64");
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
