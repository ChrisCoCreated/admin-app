function normalizeText(value) {
  return String(value || "").trim();
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

async function getOboToken(assertion, scope) {
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
    scope,
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
    const error = new Error(payload?.error_description || payload?.error || `OBO token failed (${response.status}).`);
    error.status = response.status;
    throw error;
  }

  return payload.access_token;
}

async function fetchSharePointJson(url, token, options = {}) {
  const response = await fetch(url, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=verbose",
      ...(options.headers || {}),
    },
    body: options.body,
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  if (!response.ok) {
    const message =
      payload?.error?.message?.value ||
      payload?.odata?.error?.message?.value ||
      payload?.error_description ||
      text ||
      `SharePoint request failed (${response.status}).`;
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }

  return payload;
}

async function getFormDigest(siteBaseUrl, token) {
  const payload = await fetchSharePointJson(`${siteBaseUrl}/_api/contextinfo`, token, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
    },
  });
  return normalizeText(payload?.d?.GetContextWebInformation?.FormDigestValue);
}

async function resolveListInfo(siteBaseUrl, listName, token) {
  const url = `${siteBaseUrl}/_api/web/lists?$select=Id,Title,ListItemEntityTypeFullName&$filter=Title eq ${encodeURIComponent(
    quoteODataString(listName)
  )}`;
  const payload = await fetchSharePointJson(url, token);
  const list = Array.isArray(payload?.d?.results) ? payload.d.results[0] : null;
  if (!list?.Id || !list?.ListItemEntityTypeFullName) {
    throw new Error(`Could not resolve SharePoint list '${listName}' over REST.`);
  }
  return {
    id: normalizeText(list.Id),
    entityTypeName: normalizeText(list.ListItemEntityTypeFullName),
  };
}

async function patchSharePointUrlField({
  incomingToken,
  siteBaseUrl,
  hostName,
  listName,
  itemId,
  fieldInternalName,
  urlValue,
  description,
}) {
  const scope = `https://${hostName}/.default`;
  const token = await getOboToken(incomingToken, scope);
  const digest = await getFormDigest(siteBaseUrl, token);
  const listInfo = await resolveListInfo(siteBaseUrl, listName, token);
  const payload = {
    __metadata: { type: listInfo.entityTypeName },
    [fieldInternalName]: normalizeText(urlValue)
      ? {
          __metadata: { type: "SP.FieldUrlValue" },
          Url: normalizeText(urlValue),
          Description: normalizeText(description) || normalizeText(urlValue),
        }
      : null,
  };

  await fetchSharePointJson(`${siteBaseUrl}/_api/web/lists(guid'${listInfo.id}')/items(${encodeURIComponent(itemId)})`, token, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE",
    },
    body: JSON.stringify(payload),
  });
}

module.exports = {
  patchSharePointUrlField,
};
