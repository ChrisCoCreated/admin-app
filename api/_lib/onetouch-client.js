const DEFAULT_BASE_URL = "https://api-uk.onetouchhealth.net/connect/c2/v1";

let tokenCache = {
  accessToken: "",
  expiresAtMs: 0,
};

function getRequiredEnv() {
  const baseUrl = String(process.env.ONETOUCH_BASE_URL || DEFAULT_BASE_URL).trim().replace(/\/+$/, "");
  const account = String(process.env.ONETOUCH_ACCOUNT || process.env.ONETOUCH_ACCOUNT_CODE || "").trim();
  const username = String(process.env.ONETOUCH_USERNAME || "").trim();
  const password = String(process.env.ONETOUCH_PASSWORD || "").trim();

  if (!account || !username || !password) {
    throw new Error(
      "Missing OneTouch credentials. Set ONETOUCH_ACCOUNT (or ONETOUCH_ACCOUNT_CODE), ONETOUCH_USERNAME, and ONETOUCH_PASSWORD."
    );
  }

  return {
    baseUrl,
    account,
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

  const { baseUrl, account, username, password } = getRequiredEnv();

  const loginUrl = new URL(`${baseUrl}/login`);
  loginUrl.searchParams.set("account", account);
  loginUrl.searchParams.set("username", username);
  loginUrl.searchParams.set("password", password);

  const payload = await fetchJson(loginUrl);
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

async function callOneTouch(endpointPath, query = {}) {
  const { baseUrl, account } = getRequiredEnv();
  const accessToken = await getAccessToken();

  const endpoint = String(endpointPath || "").replace(/^\/+/, "");
  const url = new URL(`${baseUrl}/${endpoint}`);
  url.searchParams.set("account", account);
  url.searchParams.set("access_token", accessToken);

  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === "") {
      continue;
    }
    url.searchParams.set(key, String(value));
  }

  return fetchJson(url);
}

function resolveRecords(payload, explicitKeys = []) {
  const keyList = [...explicitKeys, "data", "results", "items", "list", "value"];

  for (const key of keyList) {
    const list = payload?.[key];
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
