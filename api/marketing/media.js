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

async function getSharePointAccessToken(hostName) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_API_CLIENT_ID;
  const clientSecret = process.env.AZURE_API_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: `https://${hostName}/.default`,
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const payload = await response.json();
  if (!response.ok || !payload?.access_token) {
    const errorText = payload?.error_description || payload?.error || "Could not get SharePoint token.";
    throw new Error(errorText);
  }

  return payload.access_token;
}

async function getGraphAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_API_CLIENT_ID;
  const clientSecret = process.env.AZURE_API_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const payload = await response.json();
  if (!response.ok || !payload?.access_token) {
    const errorText = payload?.error_description || payload?.error || "Could not get Graph token.";
    throw new Error(errorText);
  }

  return payload.access_token;
}

function getAllowedSharePointPrefix() {
  const raw = String(process.env.SHAREPOINT_SITE_URL || "").trim();
  if (!raw) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }
  const siteUrl = new URL(raw);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  return `https://${siteUrl.hostname}${sitePath}/`;
}

function quoteODataString(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
}

function getSharePointSiteParts() {
  const raw = String(process.env.SHAREPOINT_SITE_URL || "").trim();
  if (!raw) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }
  const siteUrl = new URL(raw);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  return {
    hostName: siteUrl.hostname,
    sitePath,
  };
}

async function fetchSharePointJson(url, token) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
    },
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`SharePoint API failed (${response.status}): ${text.slice(0, 300)}`);
  }
  return text ? JSON.parse(text) : null;
}

function parseAttachmentPath(targetUrl, sitePath) {
  const decodeLoose = (value) => {
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
  };

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
    listName: decodeLoose(String(match[1] || "")).trim(),
    itemId: Number(match[2]),
    fileName: decodeLoose(String(match[3] || "")).trim(),
  };
}

async function fetchAttachmentViaSharePointItemApi(targetUrl, sharePointToken) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const attachment = parseAttachmentPath(targetUrl, sitePath);
  if (!attachment || !attachment.listName || !Number.isFinite(attachment.itemId)) {
    return null;
  }

  const attachmentsUrl = `https://${hostName}${sitePath}/_api/web/lists/GetByTitle(${quoteODataString(
    attachment.listName
  )})/items(${attachment.itemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`;
  logMediaDebug("sharepoint-item-fallback-lookup", {
    attachmentsUrl,
    listName: attachment.listName,
    itemId: attachment.itemId,
  });

  const payload = await fetchSharePointJson(attachmentsUrl, sharePointToken);
  const rows = Array.isArray(payload?.value)
    ? payload.value
    : Array.isArray(payload?.d?.results)
      ? payload.d.results
      : [];

  if (!rows.length) {
    return null;
  }

  const normalizeName = (value) => String(value || "").trim().toLowerCase();
  const requestedName = normalizeName(attachment.fileName);
  const exact = rows.find((row) => normalizeName(row?.FileName) === requestedName);
  const fallback = exact || rows[0];
  const rel = String(fallback?.ServerRelativeUrl || "").trim();
  if (!rel) {
    return null;
  }
  const absolute = `https://${hostName}${rel.startsWith("/") ? rel : `/${rel}`}`;
  logMediaDebug("sharepoint-item-fallback-content", { absolute });
  const response = await fetch(absolute, {
    headers: {
      Authorization: `Bearer ${sharePointToken}`,
    },
  });
  return response;
}

async function fetchGraphSiteId(graphToken, hostName, sitePath) {
  const url = `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${graphToken}`,
      Accept: "application/json",
    },
  });
  const payload = await response.json();
  if (!response.ok || !payload?.id) {
    throw new Error(payload?.error?.message || `Could not resolve site id (${response.status}).`);
  }
  return payload.id;
}

async function fetchGraphListIdByName(graphToken, siteId, listName) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName&$filter=displayName eq ${encodeURIComponent(
    quoteODataString(listName)
  )}&$top=1`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${graphToken}`,
      Accept: "application/json",
    },
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload?.error?.message || `Could not resolve list (${response.status}).`);
  }
  const row = Array.isArray(payload?.value) ? payload.value[0] : null;
  if (!row?.id) {
    throw new Error(`Could not find list '${listName}'.`);
  }
  return row.id;
}

async function fetchGraphListDriveId(graphToken, siteId, listId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/drive?$select=id,webUrl`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${graphToken}`,
      Accept: "application/json",
    },
  });
  const payload = await response.json();
  if (!response.ok || !payload?.id) {
    throw new Error(payload?.error?.message || `Could not resolve list drive (${response.status}).`);
  }
  return {
    id: payload.id,
    webUrl: String(payload.webUrl || "").trim(),
  };
}

function toDecodedPathname(urlObj) {
  try {
    return decodeURIComponent(String(urlObj.pathname || ""));
  } catch {
    return String(urlObj.pathname || "");
  }
}

function toGraphDrivePathForAttachments(targetUrl, sitePath) {
  const decodedPath = toDecodedPathname(targetUrl);
  if (!decodedPath.startsWith(sitePath)) {
    return "";
  }
  const relative = decodedPath.slice(sitePath.length).replace(/^\/+/, "");
  if (!/^Lists\/.+\/Attachments\/\d+\/.+/i.test(relative)) {
    return "";
  }
  return relative;
}

async function fetchViaGraphListDriveAttachment(targetUrl) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const parsed = parseAttachmentPath(targetUrl, sitePath);
  if (!parsed || !parsed.listName || !parsed.fileName || !Number.isFinite(parsed.itemId)) {
    return null;
  }

  const graphToken = await getGraphAccessToken();
  const siteId = await fetchGraphSiteId(graphToken, hostName, sitePath);
  const listId = await fetchGraphListIdByName(graphToken, siteId, parsed.listName);
  const drive = await fetchGraphListDriveId(graphToken, siteId, listId);
  const encodedSegments = ["Attachments", String(parsed.itemId), parsed.fileName]
    .map((part) => encodeURIComponent(part))
    .join("/");
  const contentUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodedSegments}:/content`;
  logMediaDebug("graph-list-drive-fallback", {
    listName: parsed.listName,
    itemId: parsed.itemId,
    contentUrl,
    driveWebUrl: drive.webUrl,
  });

  const response = await fetch(contentUrl, {
    headers: {
      Authorization: `Bearer ${graphToken}`,
    },
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Graph list drive fallback failed (${response.status}): ${detail.slice(0, 300)}`);
  }
  return response;
}

async function fetchViaGraphSiteDrive(targetUrl) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const graphDrivePath = toGraphDrivePathForAttachments(targetUrl, sitePath);
  if (!graphDrivePath) {
    return null;
  }

  const graphToken = await getGraphAccessToken();
  const siteId = await fetchGraphSiteId(graphToken, hostName, sitePath);
  const encodedPath = graphDrivePath
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
  const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/content`;
  logMediaDebug("graph-drive-fallback", { contentUrl });

  const response = await fetch(contentUrl, {
    headers: {
      Authorization: `Bearer ${graphToken}`,
    },
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Graph fallback failed (${response.status}): ${detail.slice(0, 300)}`);
  }
  return response;
}

function toGraphShareTokenFromUrl(absoluteUrl) {
  const base64 = Buffer.from(String(absoluteUrl || ""), "utf8").toString("base64");
  const base64Url = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${base64Url}`;
}

async function fetchViaGraphShares(targetUrl) {
  const graphToken = await getGraphAccessToken();
  const shareToken = toGraphShareTokenFromUrl(targetUrl.href);
  const contentUrl = `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/content`;
  logMediaDebug("graph-shares-fallback", { contentUrl });
  const response = await fetch(contentUrl, {
    headers: {
      Authorization: `Bearer ${graphToken}`,
    },
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Graph shares fallback failed (${response.status}): ${detail.slice(0, 300)}`);
  }
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

    const sharePointToken = await getSharePointAccessToken(targetUrl.hostname);
    let upstream = await fetch(targetUrl.href, {
      headers: {
        Authorization: `Bearer ${sharePointToken}`,
      },
    });

    if (upstream.status === 401 || upstream.status === 403) {
      logMediaDebug("sharepoint-direct-failed", {
        status: upstream.status,
        url: targetUrl.href,
      });

      try {
        const viaItemApi = await fetchAttachmentViaSharePointItemApi(targetUrl, sharePointToken);
        if (viaItemApi && viaItemApi.ok) {
          upstream = viaItemApi;
        } else if (viaItemApi) {
          logMediaDebug("sharepoint-item-fallback-non-ok", {
            status: viaItemApi.status,
            url: targetUrl.href,
          });
        }
      } catch (itemFallbackError) {
        logMediaDebug("sharepoint-item-fallback-error", {
          detail: itemFallbackError?.message || String(itemFallbackError),
          url: targetUrl.href,
        });
      }

      if (upstream.status === 401 || upstream.status === 403) {
        try {
          const graphFallback = await fetchViaGraphSiteDrive(targetUrl);
          if (graphFallback) {
            upstream = graphFallback;
          }
        } catch (fallbackError) {
          logMediaDebug("graph-fallback-error", {
            detail: fallbackError?.message || String(fallbackError),
            url: targetUrl.href,
          });
        }
      }

      if (upstream.status === 401 || upstream.status === 403 || upstream.status === 404) {
        try {
          const graphListDriveFallback = await fetchViaGraphListDriveAttachment(targetUrl);
          if (graphListDriveFallback) {
            upstream = graphListDriveFallback;
          }
        } catch (listDriveError) {
          logMediaDebug("graph-list-drive-fallback-error", {
            detail: listDriveError?.message || String(listDriveError),
            url: targetUrl.href,
          });
        }
      }

      if (upstream.status === 401 || upstream.status === 403 || upstream.status === 404) {
        try {
          const graphSharesFallback = await fetchViaGraphShares(targetUrl);
          if (graphSharesFallback) {
            upstream = graphSharesFallback;
          }
        } catch (sharesError) {
          logMediaDebug("graph-shares-fallback-error", {
            detail: sharesError?.message || String(sharesError),
            url: targetUrl.href,
          });
        }
      }
    }

    if (!upstream.ok) {
      const detail = await upstream.text();
      res.status(upstream.status).json({
        error: "Upstream request failed.",
        detail: detail.slice(0, 400),
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
      detail: error && error.message ? error.message : String(error),
    });
  }
};
