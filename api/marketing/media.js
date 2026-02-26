const { requireApiAuth } = require("../_lib/require-api-auth");

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

function getAllowedSharePointPrefix() {
  const raw = String(process.env.SHAREPOINT_SITE_URL || "").trim();
  if (!raw) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }
  const siteUrl = new URL(raw);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  return `https://${siteUrl.hostname}${sitePath}/`;
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
    const upstream = await fetch(targetUrl.href, {
      headers: {
        Authorization: `Bearer ${sharePointToken}`,
      },
    });

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
