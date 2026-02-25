const DEFAULT_BASE_URL = "https://api-uk.onetouchhealth.net/connect/c2/v1";

let tokenCache = {
  accessToken: "",
  expiresAtMs: 0,
};
let tokenInFlight = null;

function maskValue(value) {
  const raw = String(value || "");
  if (!raw) {
    return "";
  }
  if (raw.length <= 2) {
    return `${raw[0] || ""}*`;
  }
  return `${raw.slice(0, 2)}***${raw.slice(-1)}`;
}

function getRequiredEnv() {
  const baseUrl = String(process.env.ONETOUCH_BASE_URL || DEFAULT_BASE_URL).trim().replace(/\/+$/, "");
  const username = String(process.env.ONETOUCH_USERNAME || "").trim();
  const password = String(process.env.ONETOUCH_PASSWORD || "").trim();

  if (!username || !password) {
    throw new Error(
      "Missing OneTouch credentials. Set ONETOUCH_USERNAME and ONETOUCH_PASSWORD."
    );
  }

  return {
    baseUrl,
    username,
    password,
  };
}

async function fetchJson(url) {
  const response = await fetch(url, {
    headers: {
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
    const detail = payload?.error || payload?.message || text || `HTTP ${response.status}`;
    throw new Error(`OneTouch request failed: ${detail}`);
  }

  return payload;
}

function resolveToken(payload) {
  if (!payload || typeof payload !== "object") {
    return "";
  }

  const candidates = [
    payload.access_token,
    payload.token,
    payload?.data?.access_token,
    payload?.data?.token,
    payload?.result?.access_token,
    payload?.result?.token,
  ];

  const token = candidates.find((value) => typeof value === "string" && value.trim());
  return token ? token.trim() : "";
}

function resolveTokenTtlSeconds(payload) {
  const candidates = [
    payload?.expires_in,
    payload?.data?.expires_in,
    payload?.result?.expires_in,
    payload?.expires,
  ];

  for (const value of candidates) {
    const numeric = Number(value);
    if (Number.isFinite(numeric) && numeric > 0) {
      return numeric;
    }
  }

  return 60 * 25;
}

async function getAccessToken() {
  const now = Date.now();
  if (tokenCache.accessToken && tokenCache.expiresAtMs > now + 20_000) {
    return tokenCache.accessToken;
  }

  if (tokenInFlight) {
    return tokenInFlight;
  }

  tokenInFlight = (async () => {
    const { baseUrl, username, password } = getRequiredEnv();
    const loginUrl = `${baseUrl}/auth`;

    console.info("[OneTouch] Auth request", {
      url: loginUrl,
      usernameMasked: maskValue(username),
      hasPassword: Boolean(password),
    });

    async function requestToken() {
      const response = await fetch(loginUrl, {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          username,
          password,
        }),
      });

      const text = await response.text();
      let payload = null;
      try {
        payload = text ? JSON.parse(text) : null;
      } catch {
        payload = null;
      }

      if (!response.ok) {
        const detail = payload?.error || payload?.message || text || `HTTP ${response.status}`;
        console.warn("[OneTouch] Auth failed", {
          status: response.status,
          detail,
          usernameMasked: maskValue(username),
        });
        throw new Error(`OneTouch auth failed (${response.status}): ${detail}`);
      }

      const accessToken = resolveToken(payload);
      if (!accessToken) {
        throw new Error("OneTouch login succeeded but no access token was returned.");
      }

      const ttlMs = resolveTokenTtlSeconds(payload) * 1000;
      tokenCache = {
        accessToken,
        expiresAtMs: Date.now() + ttlMs,
      };

      return accessToken;
    }

    try {
      return await requestToken();
    } catch (firstError) {
      console.warn("[OneTouch] Auth retrying once after failure", {
        reason: firstError?.message || String(firstError),
        usernameMasked: maskValue(username),
      });
      await new Promise((resolve) => setTimeout(resolve, 200));
      return requestToken().catch(() => {
        throw firstError;
      });
    }
  })().finally(() => {
    tokenInFlight = null;
  });

  return tokenInFlight;
}

async function callOneTouch(endpointPath, query = {}) {
  const { baseUrl } = getRequiredEnv();
  const accessToken = await getAccessToken();

  const endpoint = String(endpointPath || "").replace(/^\/+/, "");
  const url = new URL(`${baseUrl}/${endpoint}`);

  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === "") {
      continue;
    }
    url.searchParams.set(key, String(value));
  }

  const response = await fetch(url, {
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
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
    const detail = payload?.error || payload?.message || text || `HTTP ${response.status}`;
    throw new Error(`OneTouch request failed (${response.status}): ${detail}`);
  }

  return payload;
}

function resolveRecords(payload, explicitKeys = []) {
  const dottedKeyList = [
    ...explicitKeys,
    "data",
    "results",
    "items",
    "list",
    "value",
    "data.visits",
    "data.timesheets",
    "data.clients",
    "data.carers",
  ];

  for (const keyPath of dottedKeyList) {
    const segments = String(keyPath).split(".");
    let list = payload;
    for (const segment of segments) {
      list = list?.[segment];
    }
    if (Array.isArray(list)) {
      return list;
    }
  }

  if (Array.isArray(payload)) {
    return payload;
  }

  return [];
}

function asString(value) {
  return String(value || "").trim();
}

function buildName(record, firstKeys, lastKeys, fallbackKeys) {
  const fallback = fallbackKeys.map((key) => asString(record?.[key])).find(Boolean);
  if (fallback) {
    return fallback;
  }

  const first = firstKeys.map((key) => asString(record?.[key])).find(Boolean);
  const last = lastKeys.map((key) => asString(record?.[key])).find(Boolean);
  return [first, last].filter(Boolean).join(" ").trim();
}

function normalizeClient(record) {
  const id = asString(record?.id || record?.client_id || record?.external_id || record?.ID);
  const name = buildName(record, ["first_name", "firstname"], ["last_name", "lastname"], ["name", "full_name"]);

  return {
    id,
    name,
    address: asString(record?.address || record?.address_1 || record?.address1),
    town: asString(record?.town || record?.city),
    county: asString(record?.county || record?.region),
    postcode: asString(record?.postcode || record?.post_code || record?.zip),
    email: asString(record?.email || record?.email_address),
    phone: asString(record?.phone || record?.mobile),
    status: asString(record?.status || record?.client_status),
    raw: record,
  };
}

function normalizeCarer(record) {
  const id = asString(record?.id || record?.carer_id || record?.external_id || record?.ID);
  const name = buildName(record, ["first_name", "firstname"], ["last_name", "lastname"], ["name", "full_name"]);

  return {
    id,
    name,
    email: asString(record?.email || record?.email_address),
    phone: asString(record?.phone || record?.mobile),
    postcode: asString(record?.postcode || record?.post_code),
    status: asString(record?.status || record?.carer_status),
    raw: record,
  };
}

function normalizeVisit(record) {
  return {
    id: asString(record?.id || record?.visit_id),
    clientId: asString(record?.client_id || record?.clientid || record?.client),
    carerId: asString(record?.carer_id || record?.carerid || record?.carer),
    startAt:
      asString(record?.start) || asString(record?.start_time) || asString(record?.start_datetime),
    raw: record,
  };
}

async function listClients() {
  const payload = await callOneTouch("clients/all");
  const records = resolveRecords(payload, ["clients"]);

  return records.map(normalizeClient).filter((client) => client.id);
}

async function listCarers() {
  const payload = await callOneTouch("carers/all");
  const records = resolveRecords(payload, ["carers"]);

  return records.map(normalizeCarer).filter((carer) => carer.id);
}

async function listVisits() {
  const payload = await callOneTouch("visits");
  const records = resolveRecords(payload, ["visits"]);

  return records.map(normalizeVisit).filter((visit) => visit.clientId && visit.carerId);
}

module.exports = {
  listCarers,
  listClients,
  listVisits,
};
