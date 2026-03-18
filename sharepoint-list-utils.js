export function normalizeServerRelativePath(path) {
  return decodeURIComponent(String(path || "")).toLowerCase();
}

export function escapeODataString(value) {
  return String(value || "").replace(/'/g, "''");
}

function firstSentence(value) {
  const text = String(value || "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) {
    return "";
  }

  const dotIndex = text.indexOf(".");
  if (dotIndex > 0) {
    return text.slice(0, dotIndex + 1);
  }

  return text;
}

export function formatSharePointError(status, bodyText) {
  const trimmed = String(bodyText || "").trim();
  if (!trimmed) {
    return `SharePoint request failed (${status}).`;
  }

  try {
    const parsed = JSON.parse(trimmed);
    const message =
      parsed?.error?.message?.value ||
      parsed?.odata?.error?.message?.value ||
      parsed?.message ||
      firstSentence(trimmed);
    return `${message} (HTTP ${status})`;
  } catch {
    return `${firstSentence(trimmed)} (HTTP ${status})`;
  }
}

export function createSharePointApi({ siteUrl, getToken }) {
  let formDigest = "";
  let digestAt = 0;
  const ensuredUserByEmail = new Map();

  function toAbsoluteUrl(pathOrUrl) {
    if (/^https?:\/\//i.test(pathOrUrl)) {
      return pathOrUrl;
    }
    return `${siteUrl}${pathOrUrl}`;
  }

  async function request(pathOrUrl, options = {}) {
    const token = await getToken();
    const response = await fetch(toAbsoluteUrl(pathOrUrl), {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=verbose",
        ...(options.headers || {}),
      },
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(formatSharePointError(response.status, text));
    }

    if (response.status === 204) {
      return null;
    }

    return response.json();
  }

  async function getFormDigest() {
    const now = Date.now();
    if (formDigest && now - digestAt < 22 * 60 * 1000) {
      return formDigest;
    }

    const data = await request("/_api/contextinfo", {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=verbose",
      },
    });

    formDigest = data.d.GetContextWebInformation.FormDigestValue;
    digestAt = now;
    return formDigest;
  }

  async function ensureCurrentUserId(email) {
    const normalized = String(email || "").trim().toLowerCase();
    if (!normalized) {
      throw new Error("No signed-in user found.");
    }

    if (ensuredUserByEmail.has(normalized)) {
      return ensuredUserByEmail.get(normalized);
    }

    const digest = await getFormDigest();
    const ensureData = await request("/_api/web/ensureuser", {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
      },
      body: JSON.stringify({
        logonName: `i:0#.f|membership|${normalized}`,
      }),
    });

    const resolvedId = Number(ensureData?.d?.Id);
    if (!Number.isFinite(resolvedId) || resolvedId <= 0) {
      throw new Error("Could not resolve current user in SharePoint.");
    }

    ensuredUserByEmail.set(normalized, resolvedId);
    return resolvedId;
  }

  async function resolveListByPath(listPath) {
    const listData = await request(
      "/_api/web/lists?$select=Id,Title,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder",
    );

    const targetPath = normalizeServerRelativePath(listPath);
    const resolved = listData?.d?.results?.find((entry) => {
      const rootPath = normalizeServerRelativePath(entry?.RootFolder?.ServerRelativeUrl);
      return rootPath === targetPath;
    });

    if (!resolved?.Id) {
      throw new Error(`List not found at path ${listPath}.`);
    }

    return resolved;
  }

  async function getListFields(listId, options = {}) {
    const includeHidden = Boolean(options?.includeHidden);
    const hiddenFilter = includeHidden ? "" : "&$filter=Hidden eq false";
    const data = await request(
      `/_api/web/lists(guid'${listId}')/fields?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField,Choices,LookupList,LookupField,AllowMultipleValues${hiddenFilter}`,
    );
    return Array.isArray(data?.d?.results) ? data.d.results : [];
  }

  async function createListItem(listId, payload) {
    const digest = await getFormDigest();
    return request(`/_api/web/lists(guid'${listId}')/items`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
      },
      body: JSON.stringify(payload),
    });
  }

  async function updateListItem(listId, itemId, entityTypeName, payloadFields) {
    const digest = await getFormDigest();
    const payload = {
      __metadata: { type: entityTypeName },
      ...(payloadFields || {}),
    };

    return request(`/_api/web/lists(guid'${listId}')/items(${itemId})`, {
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

  async function deleteListItem(listId, itemId) {
    const digest = await getFormDigest();
    return request(`/_api/web/lists(guid'${listId}')/items(${itemId})`, {
      method: "POST",
      headers: {
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
      },
    });
  }

  async function addAttachment(listId, itemId, fileName, buffer) {
    const digest = await getFormDigest();
    const safeName = escapeODataString(fileName);

    return request(`/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/add(FileName='${safeName}')`, {
      method: "POST",
      headers: {
        "Content-Type": "application/octet-stream",
        "X-RequestDigest": digest,
      },
      body: buffer,
    });
  }

  function clearCaches() {
    formDigest = "";
    digestAt = 0;
    ensuredUserByEmail.clear();
  }

  return {
    request,
    getFormDigest,
    ensureCurrentUserId,
    resolveListByPath,
    getListFields,
    createListItem,
    updateListItem,
    deleteListItem,
    addAttachment,
    clearCaches,
  };
}

export function buildFieldMaps(fields) {
  const titleToField = new Map();
  const internalToField = new Map();

  for (const field of fields || []) {
    if (!field?.Title || !field?.InternalName) {
      continue;
    }
    titleToField.set(field.Title.trim().toLowerCase(), field);
    internalToField.set(field.InternalName.trim().toLowerCase(), field);
  }

  return { titleToField, internalToField };
}

export function fieldByTitle(fieldMap, candidates) {
  for (const title of candidates) {
    const match = fieldMap?.titleToField?.get(title.trim().toLowerCase());
    if (match) {
      return match;
    }
  }
  return null;
}

export function fieldByTitleWritable(fieldMap, candidates) {
  for (const title of candidates) {
    const match = fieldMap?.titleToField?.get(title.trim().toLowerCase());
    if (match && !match.ReadOnlyField) {
      return match;
    }
  }
  return null;
}

export function fieldByInternalName(fieldMap, candidates) {
  for (const name of candidates) {
    const match = fieldMap?.internalToField?.get(name.trim().toLowerCase());
    if (match) {
      return match;
    }
  }
  return null;
}
