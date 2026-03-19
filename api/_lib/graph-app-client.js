const tokenCache = {
  accessToken: "",
  expiresAt: 0,
};

function parseJsonSafe(text) {
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

async function getGraphAppAccessToken() {
  if (tokenCache.accessToken && tokenCache.expiresAt - 60000 > Date.now()) {
    return tokenCache.accessToken;
  }

  const tenantId = String(process.env.AZURE_TENANT_ID || "").trim();
  const clientId = String(process.env.AZURE_API_CLIENT_ID || "").trim();
  const clientSecret = String(process.env.AZURE_API_CLIENT_SECRET || "").trim();

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
      }),
    }
  );

  const text = await response.text();
  const payload = parseJsonSafe(text);
  const accessToken = String(payload?.access_token || "").trim();
  if (!response.ok || !accessToken) {
    const error = new Error(
      payload?.error_description || payload?.error || `Could not get Graph app token (${response.status}).`
    );
    error.status = response.status || 500;
    error.code = "GRAPH_APP_TOKEN_FAILED";
    throw error;
  }

  const expiresInSeconds = Number(payload?.expires_in || 3600);
  tokenCache.accessToken = accessToken;
  tokenCache.expiresAt = Date.now() + Math.max(60000, expiresInSeconds * 1000);
  return tokenCache.accessToken;
}

async function fetchGraphJson(url, options = {}) {
  const token = await getGraphAppAccessToken();
  const response = await fetch(url, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...(options.headers || {}),
    },
    body: options.body,
  });

  const text = await response.text();
  const payload = parseJsonSafe(text);

  if (!response.ok) {
    const error = new Error(payload?.error?.message || text || `Graph request failed (${response.status}).`);
    error.status = response.status;
    error.code = payload?.error?.code || "GRAPH_REQUEST_FAILED";
    throw error;
  }

  return payload;
}

function createGraphAppClient() {
  return {
    fetchJson(url, options = {}) {
      return fetchGraphJson(url, options);
    },
    async fetchAllPages(initialUrl) {
      const rows = [];
      let nextUrl = initialUrl;

      while (nextUrl) {
        const payload = await fetchGraphJson(nextUrl);
        rows.push(...(Array.isArray(payload?.value) ? payload.value : []));
        nextUrl = String(payload?.["@odata.nextLink"] || "").trim();
      }

      return rows;
    },
  };
}

module.exports = {
  createGraphAppClient,
  getGraphAppAccessToken,
};
