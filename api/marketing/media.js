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
