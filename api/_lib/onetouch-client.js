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

function parseDateOfBirth(value) {
  const raw = asString(value);
  if (!raw) {
    return "";
  }

  const datePart = raw.includes("T") ? raw.split("T")[0] : raw;
  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(datePart);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const slashMatch = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(datePart);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const yearRaw = Number(slashMatch[3]);
    const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return [
      String(parsed.getUTCFullYear()).padStart(4, "0"),
      String(parsed.getUTCMonth() + 1).padStart(2, "0"),
      String(parsed.getUTCDate()).padStart(2, "0"),
    ].join("-");
  }

  return "";
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
  const tags = Array.isArray(record?.tags)
    ? record.tags
        .map((tag) => asString(tag?.tag || tag?.name || tag))
        .filter(Boolean)
    : [];

  return {
    id,
    name,
    dateOfBirth: parseDateOfBirth(
      record?.date_of_birth || record?.dob || record?.birth_date || record?.birthdate
    ),
    address: asString(record?.address || record?.address_1 || record?.address1),
    town: asString(record?.town || record?.city),
    county: asString(record?.county || record?.region),
    postcode: asString(record?.postcode || record?.post_code || record?.zip),
    email: asString(record?.email || record?.email_address),
    phone: asString(record?.phone || record?.mobile),
    status: asString(record?.status || record?.client_status),
    tags,
    raw: record,
  };
}

function normalizeCarer(record) {
  const id = asString(record?.id || record?.carer_id || record?.external_id || record?.ID);
  const name = buildName(record, ["first_name", "firstname"], ["last_name", "lastname"], ["name", "full_name"]);
  const tags = Array.isArray(record?.tags)
    ? record.tags
        .map((tag) => asString(tag?.tag || tag?.name || tag))
        .filter(Boolean)
    : [];

  return {
    id,
    name,
    email: asString(record?.primary_email || record?.email || record?.email_address),
    phone: asString(record?.phone_mobile || record?.phone || record?.mobile),
    postcode: asString(record?.postcode || record?.post_code || record?.postCode || record?.zip),
    status: asString(record?.status || record?.carer_status),
    tags,
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
  const firstPayload = await callOneTouch("clients/all", { page: 1 });
  const allRecords = [...resolveRecords(firstPayload, ["clients"])];

  const firstCurrentPage = Number(firstPayload?.current_page || 1);
  const firstLastPage = Number(firstPayload?.last_page || firstCurrentPage);
  const hasNextUrl = Boolean(firstPayload?.next_page_url);
  const hasMorePages =
    hasNextUrl || (Number.isFinite(firstLastPage) && firstLastPage > firstCurrentPage);

  if (hasMorePages && Number.isFinite(firstLastPage) && firstLastPage > 1) {
    const pageNumbers = [];
    for (let page = 2; page <= firstLastPage; page += 1) {
      pageNumbers.push(page);
    }

    const concurrencyRaw = Number(process.env.ONETOUCH_CLIENTS_PAGE_CONCURRENCY || 4);
    const concurrency = Number.isFinite(concurrencyRaw)
      ? Math.max(1, Math.min(Math.floor(concurrencyRaw), 10))
      : 4;
    let nextIndex = 0;

    const workers = Array.from(
      { length: Math.min(concurrency, pageNumbers.length) },
      async () => {
        while (true) {
          const index = nextIndex;
          nextIndex += 1;
          if (index >= pageNumbers.length) {
            return;
          }
          const payload = await callOneTouch("clients/all", { page: pageNumbers[index] });
          allRecords.push(...resolveRecords(payload, ["clients"]));
        }
      }
    );

    await Promise.all(workers);
  }

  return allRecords.map(normalizeClient).filter((client) => client.id);
}

async function listCarers() {
  const firstPayload = await callOneTouch("carers/all", { page: 1 });
  const allRecords = [...resolveRecords(firstPayload, ["carers"])];

  const firstCurrentPage = Number(firstPayload?.current_page || 1);
  const firstLastPage = Number(firstPayload?.last_page || firstCurrentPage);
  const hasNextUrl = Boolean(firstPayload?.next_page_url);
  const hasMorePages =
    hasNextUrl || (Number.isFinite(firstLastPage) && firstLastPage > firstCurrentPage);

  if (hasMorePages && Number.isFinite(firstLastPage) && firstLastPage > 1) {
    const pageNumbers = [];
    for (let page = 2; page <= firstLastPage; page += 1) {
      pageNumbers.push(page);
    }

    const concurrencyRaw = Number(process.env.ONETOUCH_CARERS_PAGE_CONCURRENCY || 4);
    const concurrency = Number.isFinite(concurrencyRaw)
      ? Math.max(1, Math.min(Math.floor(concurrencyRaw), 10))
      : 4;
    let nextIndex = 0;

    const workers = Array.from(
      { length: Math.min(concurrency, pageNumbers.length) },
      async () => {
        while (true) {
          const index = nextIndex;
          nextIndex += 1;
          if (index >= pageNumbers.length) {
            return;
          }
          const payload = await callOneTouch("carers/all", { page: pageNumbers[index] });
          allRecords.push(...resolveRecords(payload, ["carers"]));
        }
      }
    );

    await Promise.all(workers);
  }

  return allRecords.map(normalizeCarer).filter((carer) => carer.id);
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
