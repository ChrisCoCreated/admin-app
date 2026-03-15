const DEFAULT_BASE_URL = "https://api-uk.onetouchhealth.net/connect/c2/v1";
const ONETOUCH_CREATE_DEBUG = process.env.ONETOUCH_CREATE_DEBUG === "1";
const ALLOWED_CARER_CREATE_FIELDS = new Set([
  "external_id",
  "increment",
  "strict",
  "title",
  "firstname",
  "lastname",
  "known_as",
  "sex",
  "dob",
  "pps",
  "ni_number",
  "primary_email",
  "secondary_email",
  "phone_mobile",
  "phone_home",
  "phone_work",
  "address",
  "town",
  "county",
  "postcode",
  "country",
  "nationality",
  "date_start",
  "transport_type",
  "recruitment_source",
  "position",
  "location",
  "branch",
  "area",
  "status",
  "comment",
  "bank_account_name",
  "bank_sort_code",
  "bank_account_number",
  "bank_iban",
  "profile_photo",
]);

let tokenCache = {
  accessToken: "",
  expiresAtMs: 0,
};
let tokenInFlight = null;
const generalLocationAreaCache = {
  expiresAtMs: 0,
  locations: [],
  areas: [],
};

function clearTokenCache() {
  tokenCache = {
    accessToken: "",
    expiresAtMs: 0,
  };
  tokenInFlight = null;
}

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

function createOneTouchError({ endpoint = "", status = 0, payload = null, text = "" } = {}) {
  const detail = payload?.error || payload?.message || text || `HTTP ${status || 0}`;
  const error = new Error(`OneTouch request failed (${status || 0}): ${detail}`);
  error.endpoint = endpoint;
  error.status = status || 0;
  error.payload = payload;
  error.raw = text;
  return error;
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
  const endpoint = String(endpointPath || "").replace(/^\/+/, "");
  const url = new URL(`${baseUrl}/${endpoint}`);

  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === "") {
      continue;
    }
    url.searchParams.set(key, String(value));
  }

  async function requestWithToken(accessToken, options = {}) {
    const method = String(options.method || "GET").trim().toUpperCase();
    const headers = {
      Accept: "application/json",
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers && typeof options.headers === "object" ? options.headers : {}),
    };
    const requestOptions = {
      method,
      headers,
    };
    if (options.body !== undefined) {
      requestOptions.body =
        typeof options.body === "string" ? options.body : JSON.stringify(options.body);
      if (!headers["Content-Type"]) {
        headers["Content-Type"] = "application/json";
      }
    }

    const response = await fetch(url, {
      ...requestOptions,
    });

    const text = await response.text();
    let payload = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = null;
    }

    return { response, payload, text };
  }

  async function callWithRetryOnce(options = {}) {
    const firstToken = await getAccessToken();
    const first = await requestWithToken(firstToken, options);
    if (first.response.ok) {
      return first.payload;
    }

    if (first.response.status === 401) {
      console.warn("[OneTouch] Received 401, clearing cached token and retrying request once.", {
        endpoint,
      });
      clearTokenCache();
      const retryToken = await getAccessToken();
      const retry = await requestWithToken(retryToken, options);
      if (retry.response.ok) {
        return retry.payload;
      }
      const retryDetail = retry.payload?.error || retry.payload?.message || retry.text || `HTTP ${retry.response.status}`;
      throw new Error(`OneTouch request failed (${retry.response.status}): ${retryDetail}`);
    }

    const detail = first.payload?.error || first.payload?.message || first.text || `HTTP ${first.response.status}`;
    throw new Error(`OneTouch request failed (${first.response.status}): ${detail}`);
  }

  return callWithRetryOnce();
}

async function postOneTouch(endpointPath, body = {}, query = {}) {
  const { baseUrl } = getRequiredEnv();
  const endpoint = String(endpointPath || "").replace(/^\/+/, "");
  const url = new URL(`${baseUrl}/${endpoint}`);

  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === "") {
      continue;
    }
    url.searchParams.set(key, String(value));
  }

  async function requestWithToken(accessToken) {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body && typeof body === "object" ? body : {}),
    });

    const text = await response.text();
    let payload = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = null;
    }

    return { response, payload, text };
  }

  const firstToken = await getAccessToken();
  const first = await requestWithToken(firstToken);
  if (first.response.ok) {
    return first.payload;
  }

  if (first.response.status === 401) {
    console.warn("[OneTouch] Received 401, clearing cached token and retrying request once.", {
      endpoint,
    });
    clearTokenCache();
    const retryToken = await getAccessToken();
    const retry = await requestWithToken(retryToken);
    if (retry.response.ok) {
      return retry.payload;
    }
    throw createOneTouchError({
      endpoint,
      status: retry.response.status,
      payload: retry.payload,
      text: retry.text,
    });
  }

  throw createOneTouchError({
    endpoint,
    status: first.response.status,
    payload: first.payload,
    text: first.text,
  });
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

function normalizeToken(value) {
  return asString(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function stripTrailingUkPostcode(value) {
  const raw = asString(value);
  if (!raw) {
    return "";
  }
  const normalized = raw.replace(/\s+/g, " ").trim();
  const trimmed = normalized.replace(/[\s,;-]+([A-Z]{1,2}\d[A-Z\d]{0,2})$/i, "").trim();
  return trimmed || normalized;
}

function scoreMatch(candidate, lookup) {
  if (!candidate || !lookup) {
    return 0;
  }
  if (candidate === lookup) {
    return 100;
  }
  if (candidate.startsWith(lookup) || lookup.startsWith(candidate)) {
    return 75;
  }
  if (candidate.includes(lookup) || lookup.includes(candidate)) {
    return 50;
  }
  return 0;
}

function pickBestByName(items, value) {
  const lookupText = stripTrailingUkPostcode(value);
  const lookup = normalizeToken(lookupText);
  if (!lookup) {
    return null;
  }

  let best = null;
  let bestScore = 0;
  for (const item of items || []) {
    const candidate = normalizeToken(item?.name || "");
    const score = scoreMatch(candidate, lookup);
    if (score > bestScore) {
      best = item;
      bestScore = score;
    }
  }
  return bestScore > 0 ? best : null;
}

function normalizeGeneralLocation(record) {
  const name = asString(
    record?.name ||
      record?.location ||
      record?.location_name ||
      record?.title ||
      record?.label
  );
  const area = asString(
    record?.area ||
      record?.area_name ||
      record?.areaName ||
      record?.area_title ||
      record?.areaLabel ||
      record?.area?.name
  );
  return { name, area };
}

function normalizeGeneralArea(record) {
  const name = asString(
    record?.name ||
      record?.area ||
      record?.area_name ||
      record?.title ||
      record?.label
  );
  return { name };
}

async function fetchGeneralLocationsAndAreas() {
  const now = Date.now();
  if (generalLocationAreaCache.expiresAtMs > now && generalLocationAreaCache.locations.length) {
    return {
      locations: generalLocationAreaCache.locations,
      areas: generalLocationAreaCache.areas,
    };
  }

  const [locationsPayload, areasPayload] = await Promise.all([
    callOneTouch("general/locations"),
    callOneTouch("general/get-areas"),
  ]);

  const locationRecords = resolveRecords(locationsPayload, ["locations", "data.locations", "result.locations"]);
  const areaRecords = resolveRecords(areasPayload, ["areas", "data.areas", "result.areas"]);
  const locations = locationRecords.map(normalizeGeneralLocation).filter((row) => row.name);
  const areas = areaRecords.map(normalizeGeneralArea).filter((row) => row.name);

  generalLocationAreaCache.expiresAtMs = now + 5 * 60 * 1000;
  generalLocationAreaCache.locations = locations;
  generalLocationAreaCache.areas = areas;

  return { locations, areas };
}

async function getOneTouchLocationAreaOptions() {
  const { locations, areas } = await fetchGeneralLocationsAndAreas();
  const uniqueLocations = Array.from(new Set(locations.map((row) => asString(row?.name)).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b)
  );
  const uniqueAreas = Array.from(new Set(areas.map((row) => asString(row?.name)).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b)
  );
  return {
    locations: uniqueLocations,
    areas: uniqueAreas,
  };
}

async function getOneTouchAreaOptions() {
  const payload = await callOneTouch("general/get-areas");
  const areaRecords = resolveRecords(payload, ["areas", "data.areas", "result.areas"]);
  const areas = areaRecords.map(normalizeGeneralArea).map((row) => asString(row?.name)).filter(Boolean);
  return Array.from(new Set(areas)).sort((a, b) => a.localeCompare(b));
}

function normalizeRecruitmentSource(record) {
  const name = asString(
    record?.name ||
      record?.source ||
      record?.recruitment_source ||
      record?.recruitmentSource ||
      record?.title ||
      record?.label
  );
  return { name };
}

async function getOneTouchRecruitmentSourceOptions() {
  const payload = await callOneTouch("carer/recruitment-sources");
  const sourceRecords = resolveRecords(payload, [
    "sources",
    "recruitment_sources",
    "data.sources",
    "data.recruitment_sources",
    "result.sources",
    "result.recruitment_sources",
  ]);
  const sources = sourceRecords.map(normalizeRecruitmentSource).map((row) => asString(row?.name)).filter(Boolean);
  return Array.from(new Set(sources)).sort((a, b) => a.localeCompare(b));
}

function normalizePositionOption(record) {
  const name = asString(
    record?.name ||
      record?.position ||
      record?.position_name ||
      record?.title ||
      record?.label
  );
  return { name };
}

async function getOneTouchPositionOptions() {
  const payload = await callOneTouch("carer/get-positions");
  const records = resolveRecords(payload, ["positions", "data.positions", "result.positions"]);
  const values = records.map(normalizePositionOption).map((row) => asString(row?.name)).filter(Boolean);
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b));
}

function normalizeStatusOption(record) {
  const name = asString(
    record?.name ||
      record?.status ||
      record?.status_name ||
      record?.title ||
      record?.label
  );
  return { name };
}

async function getOneTouchStatusOptions() {
  const payload = await callOneTouch("carer/statuses");
  const records = resolveRecords(payload, ["statuses", "data.statuses", "result.statuses"]);
  const values = records.map(normalizeStatusOption).map((row) => asString(row?.name)).filter(Boolean);
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b));
}

async function resolveOneTouchLocationArea({ location = "", livesIn = "", explicitArea = "" } = {}) {
  const locationHint = stripTrailingUkPostcode(location) || asString(location);
  const livesInHint = stripTrailingUkPostcode(livesIn) || asString(livesIn);
  const explicitAreaHint = asString(explicitArea);

  try {
    const { locations, areas } = await fetchGeneralLocationsAndAreas();

    const matchedLocation =
      pickBestByName(locations, locationHint) ||
      pickBestByName(locations, livesInHint) ||
      null;

    const areaFromLocation = asString(matchedLocation?.area);
    const matchedArea =
      pickBestByName(areas, explicitAreaHint) ||
      pickBestByName(areas, areaFromLocation) ||
      pickBestByName(areas, locationHint) ||
      pickBestByName(areas, livesInHint) ||
      null;

    const resolvedLocation = asString(matchedLocation?.name) || locationHint || livesInHint;
    const resolvedArea = asString(matchedArea?.name) || areaFromLocation || explicitAreaHint || "";

    return {
      location: resolvedLocation,
      area: resolvedArea,
    };
  } catch {
    return {
      location: locationHint || livesInHint,
      area: explicitAreaHint || "",
    };
  }
}

function collectNonEmptyValues(values) {
  return values.map((value) => asString(value)).filter(Boolean);
}

function dedupeEmails(values) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const key = String(value || "").trim().toLowerCase();
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    output.push(String(value).trim());
  }
  return output;
}

function normalizePhoneForCompare(value) {
  return String(value || "").replace(/[^\d+]/g, "").toLowerCase();
}

function dedupePhones(values) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const raw = String(value || "").trim();
    const key = normalizePhoneForCompare(raw);
    if (!raw || !key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    output.push(raw);
  }
  return output;
}

function buildCombinedEmail(record) {
  const values = collectNonEmptyValues([
    record?.primary_email,
    record?.secondary_email,
    record?.email,
    record?.email_address,
  ]);
  return dedupeEmails(values).join("; ");
}

function buildCombinedPhone(record) {
  const values = collectNonEmptyValues([
    record?.phone_mobile,
    record?.phone_home,
    record?.phone,
    record?.mobile,
  ]);
  return dedupePhones(values).join(" / ");
}

function resolveCareType(record) {
  const explicit = collectNonEmptyValues([
    record?.care_type,
    record?.careType,
    record?.service_type,
    record?.serviceType,
    record?.support_type,
    record?.supportType,
  ])[0];
  if (explicit) {
    const normalized = explicit.toLowerCase();
    if (normalized.includes("companion")) {
      return "Companionship";
    }
    if (normalized.includes("care")) {
      return "Care";
    }
    return explicit;
  }

  const tags = Array.isArray(record?.tags)
    ? record.tags.map((tag) => asString(tag?.tag || tag?.name || tag)).filter(Boolean)
    : [];
  for (const tag of tags) {
    const normalized = tag.toLowerCase();
    if (normalized.includes("companion")) {
      return "Companionship";
    }
    if (normalized.includes("care")) {
      return "Care";
    }
  }

  return "";
}

function resolveLocation(record) {
  return (
    collectNonEmptyValues([
      record?.location,
      record?.area,
      record?.zone,
      record?.patch,
      record?.town,
      record?.county,
    ])[0] || ""
  );
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
    emailPrimary: asString(record?.primary_email),
    emailSecondary: asString(record?.secondary_email),
    phoneMobile: asString(record?.phone_mobile),
    phoneHome: asString(record?.phone_home),
    dateOfBirth: parseDateOfBirth(
      record?.date_of_birth || record?.dob || record?.birth_date || record?.birthdate
    ),
    location: resolveLocation(record),
    careType: resolveCareType(record),
    address: asString(record?.address || record?.address_1 || record?.address1),
    town: asString(record?.town || record?.city),
    county: asString(record?.county || record?.region),
    postcode: asString(record?.postcode || record?.post_code || record?.zip),
    email: buildCombinedEmail(record),
    phone: buildCombinedPhone(record),
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
  const lowerCareComp = tags.find((tag) => {
    const token = tag.toLowerCase();
    return token.includes("companion") || token.includes("care");
  });
  const careCompanionshipTag = lowerCareComp || "";
  const otherTags = tags.filter((tag) => tag !== careCompanionshipTag);
  const area = asString(
    record?.organisation?.area?.name ||
      record?.area ||
      record?.location ||
      record?.town ||
      record?.county
  );
  const contractedHours = Number(
    record?.contracted_hrs?.total_weekly ??
      record?.contracted_hours?.total_weekly ??
      record?.contracted_hrs_total ??
      0
  );

  return {
    id,
    name,
    email: asString(record?.primary_email || record?.email || record?.email_address),
    phone: asString(record?.phone_mobile || record?.phone || record?.mobile),
    postcode: asString(record?.postcode || record?.post_code || record?.postCode || record?.zip),
    status: asString(record?.status || record?.carer_status),
    tags,
    area,
    contractedHours: Number.isFinite(contractedHours) ? contractedHours : 0,
    careCompanionshipTag,
    otherTags,
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

async function getCarerByExternalId(externalId) {
  const id = asString(externalId);
  if (!id) {
    return null;
  }
  const payload = await callOneTouch(`carer/get/${encodeURIComponent(id)}`);
  return normalizeCarer(payload || {});
}

async function listCarersDetailed() {
  const carers = await listCarers();
  const concurrencyRaw = Number(process.env.ONETOUCH_CARER_DETAIL_CONCURRENCY || 4);
  const concurrency = Number.isFinite(concurrencyRaw)
    ? Math.max(1, Math.min(Math.floor(concurrencyRaw), 10))
    : 4;
  let nextIndex = 0;

  const enriched = new Array(carers.length);
  const workers = Array.from(
    { length: Math.min(concurrency, carers.length) },
    async () => {
      while (true) {
        const index = nextIndex;
        nextIndex += 1;
        if (index >= carers.length) {
          return;
        }
        const carer = carers[index];
        try {
          const detail = await getCarerByExternalId(carer.id);
          enriched[index] = detail || carer;
        } catch {
          enriched[index] = carer;
        }
      }
    }
  );

  await Promise.all(workers);
  return enriched.filter(Boolean);
}

async function listVisits() {
  const payload = await callOneTouch("visits");
  const records = resolveRecords(payload, ["visits"]);

  return records.map(normalizeVisit).filter((visit) => visit.clientId && visit.carerId);
}

function normalizeTimesheet(record) {
  return {
    date: asString(record?.date),
    carerName: asString(record?.carer_name || record?.carerName),
    clientName: asString(record?.client_name || record?.clientName),
    carerId: asString(record?.carer_id || record?.carerId),
    clientId: asString(record?.client_id || record?.clientId),
    externalCarer: Boolean(record?.external_carer ?? record?.externalCarer),
    jobType: asString(record?.jobtype || record?.job_type || record?.jobType),
    dueIn: asString(record?.scheduled_start || record?.due_in || record?.dueIn),
    dueOut: asString(record?.scheduled_finish || record?.due_out || record?.dueOut),
    logIn: asString(record?.full_start || record?.log_in || record?.logIn || record?.actual_start),
    logOut: asString(record?.full_finish || record?.log_out || record?.logOut || record?.actual_finish),
    timeConfirmed: Boolean(
      record?.time_confirmed ??
        record?.timeConfirmed ??
        String(record?.shift_status || "").trim().toLowerCase() === "complete"
    ),
    billing: asString(record?.billing),
    pay: asString(record?.pay),
    travelPay: asString(record?.travel_pay || record?.travelPay),
    branch: asString(record?.branch),
    area: asString(record?.area),
    shiftStatus: asString(record?.shift_status || record?.shiftStatus),
    raw: record,
  };
}

function getNextPageNumber(payload) {
  const nextUrl = asString(payload?.next_page_url || payload?.nextPageUrl);
  if (!nextUrl) {
    return null;
  }

  try {
    const parsed = new URL(nextUrl);
    const page = Number(parsed.searchParams.get("page"));
    return Number.isFinite(page) && page > 0 ? page : null;
  } catch {
    return null;
  }
}

async function listTimesheets({ carerId = "", date = "", dateStart = "", dateFinish = "", perPage = 200 } = {}) {
  const query = {
    carer_id: asString(carerId),
    carerId: asString(carerId),
    date: asString(date),
    datestart: asString(dateStart),
    datefinish: asString(dateFinish),
    date_start: asString(dateStart),
    date_finish: asString(dateFinish),
    per_page: Number.isFinite(Number(perPage)) ? Math.max(1, Math.min(Number(perPage), 500)) : 200,
    page: 1,
  };

  if (!query.carer_id) {
    throw new Error("A carer id is required to fetch timesheets.");
  }

  let pagePayload = await callOneTouch("finance/summary", query);
  let records = resolveRecords(pagePayload, ["data"]).map(normalizeTimesheet);
  let nextPage = getNextPageNumber(pagePayload);
  let pagesFetched = 1;

  while (nextPage && pagesFetched < 50) {
    pagePayload = await callOneTouch("finance/summary", {
      ...query,
      page: nextPage,
    });
    records = records.concat(resolveRecords(pagePayload, ["data"]).map(normalizeTimesheet));
    nextPage = getNextPageNumber(pagePayload);
    pagesFetched += 1;
  }

  console.info("[OneTouch] finance/summary response", {
    carerId: query.carer_id,
    date: query.date,
    dateStart: query.datestart || query.date_start,
    dateFinish: query.datefinish || query.date_finish,
    records: records.length,
    pagesFetched,
  });

  return {
    timesheets: records,
    total: Number(pagePayload?.total || records.length) || records.length,
    pageCount: pagesFetched,
    filters: {
      carerId: query.carer_id,
      date: query.date,
      dateStart: query.datestart,
      dateFinish: query.datefinish,
      perPage: query.per_page,
    },
  };
}

function pickFirstNonEmpty(values) {
  for (const value of values) {
    const text = asString(value);
    if (text) {
      return text;
    }
  }
  return "";
}

function splitName(name) {
  const raw = asString(name);
  if (!raw) {
    return { first: "", last: "" };
  }
  const parts = raw.split(/\s+/).filter(Boolean);
  if (parts.length === 1) {
    return { first: parts[0], last: "Unknown" };
  }
  return {
    first: parts[0],
    last: parts.slice(1).join(" "),
  };
}

function sanitizeCarerCreatePayload(payload = {}) {
  const candidateName = pickFirstNonEmpty([payload.full_name, payload.name, payload.candidate_name]);
  const split = splitName(candidateName);
  const externalId = pickFirstNonEmpty([payload.external_id, payload.externalId, payload.id]);

  const base = {
    external_id: externalId,
    firstname: pickFirstNonEmpty([payload.firstname, payload.first_name, split.first]),
    lastname: pickFirstNonEmpty([payload.lastname, payload.last_name, split.last]),
    known_as: pickFirstNonEmpty([payload.known_as, payload.knownAs, split.first]),
    primary_email: pickFirstNonEmpty([payload.primary_email, payload.email]),
    phone_mobile: pickFirstNonEmpty([payload.phone_mobile, payload.phone, payload.phone_number]),
    town: pickFirstNonEmpty([payload.town, payload.location]),
    area: pickFirstNonEmpty([payload.area, payload.location]),
    recruitment_source: pickFirstNonEmpty([payload.recruitment_source, payload.source]),
    status: pickFirstNonEmpty([payload.status, payload.carer_status]),
    comment: pickFirstNonEmpty([payload.comment, payload.notes]),
    position: pickFirstNonEmpty([payload.position, "Carer"]),
  };

  const out = {};
  for (const [key, value] of Object.entries(base)) {
    const text = asString(value);
    if (!text) {
      continue;
    }
    out[key] = text;
  }
  return out;
}

function pickObjectKeys(input, keys) {
  const out = {};
  for (const key of keys) {
    if (input[key] === undefined || input[key] === null) {
      continue;
    }
    const value = asString(input[key]);
    if (!value) {
      continue;
    }
    out[key] = value;
  }
  return out;
}

function enforceAllowedCarerCreateFields(payload = {}) {
  const out = {};
  for (const [key, value] of Object.entries(payload || {})) {
    if (!ALLOWED_CARER_CREATE_FIELDS.has(key)) {
      continue;
    }
    if (value === undefined || value === null) {
      continue;
    }
    if (typeof value === "string") {
      const text = asString(value);
      if (!text) {
        continue;
      }
      out[key] = text;
      continue;
    }
    out[key] = value;
  }
  return out;
}

async function createCarer(payload = {}) {
  const requestId = `otc_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
  let createPayload = sanitizeCarerCreatePayload(payload);
  const explicitLocation = asString(payload?.location);
  const explicitArea = asString(payload?.area);
  const hasExplicitArea = Boolean(explicitArea);

  if (hasExplicitArea) {
    if (explicitLocation) {
      createPayload.location = explicitLocation;
    }
    createPayload.area = explicitArea;
    if (!createPayload.town && createPayload.location) {
      createPayload.town = explicitLocation;
    }
  } else {
    const resolved = await resolveOneTouchLocationArea({
      location: createPayload.location,
      livesIn: payload?.livesIn || payload?.lives_in || "",
      explicitArea: createPayload.area,
    });
    if (resolved.location) {
      createPayload.location = resolved.location;
      if (!createPayload.town) {
        createPayload.town = resolved.location;
      }
    }
    if (resolved.area) {
      createPayload.area = resolved.area;
    }
  }
  if (!createPayload.firstname || !createPayload.lastname) {
    throw new Error("OneTouch create requires candidate name.");
  }

  createPayload = enforceAllowedCarerCreateFields(createPayload);

  console.info("[OneTouch] carer/create request", {
    requestId,
    external_id: createPayload.external_id || "",
    firstname: createPayload.firstname || "",
    lastname: createPayload.lastname || "",
    area: createPayload.area || "",
    recruitment_source: createPayload.recruitment_source || "",
    position: createPayload.position || "",
    status: createPayload.status || "",
    location: createPayload.location || "",
    town: createPayload.town || "",
  });

  const attempts = [
    { label: "full", payload: { ...createPayload } },
    { label: "without_status", payload: (() => {
      const copy = { ...createPayload };
      delete copy.status;
      return copy;
    })() },
    { label: "without_status_location", payload: (() => {
      const copy = { ...createPayload };
      delete copy.status;
      delete copy.location;
      return copy;
    })() },
    {
      label: "minimal",
      payload: pickObjectKeys(createPayload, [
        "external_id",
        "firstname",
        "lastname",
        "known_as",
        "primary_email",
        "phone_mobile",
        "town",
        "area",
        "recruitment_source",
        "position",
        "comment",
      ]),
    },
  ];

  let response = null;
  let lastError = null;
  for (const attempt of attempts) {
    try {
      console.info("[OneTouch] carer/create attempt", {
        requestId,
        attempt: attempt.label,
        keys: Object.keys(attempt.payload).sort(),
      });
      response = await postOneTouch("carer/create", attempt.payload);
      console.info("[OneTouch] carer/create attempt success", {
        requestId,
        attempt: attempt.label,
      });
      break;
    } catch (error) {
      lastError = error;
      const isQueryError = /query error/i.test(error?.message || "");
      console.warn("[OneTouch] carer/create attempt failed", {
        requestId,
        attempt: attempt.label,
        status: error?.status || 0,
        isQueryError,
        message: error?.message || String(error),
      });
      if (!isQueryError) {
        break;
      }
    }
  }

  if (!response) {
    const error = lastError || new Error("OneTouch create failed.");
    console.error("[OneTouch] carer/create failed", {
      requestId,
      endpoint: error?.endpoint || "carer/create",
      status: error?.status || 0,
      payload: error?.payload ?? null,
      message: error?.message || String(error),
      request: {
        external_id: createPayload.external_id || "",
        firstname: createPayload.firstname || "",
        lastname: createPayload.lastname || "",
        area: createPayload.area || "",
        recruitment_source: createPayload.recruitment_source || "",
        position: createPayload.position || "",
        status: createPayload.status || "",
        location: createPayload.location || "",
        town: createPayload.town || "",
      },
    });
    const message = error?.message || "OneTouch create failed.";
    throw new Error(`[${requestId}] ${message}`);
  }

  const success = response?.success === true || response?.status === true;
  const id = asString(response?.id || response?.carer_id || response?.data?.id || response?.result?.id);
  if (!success || !id) {
    console.error("[OneTouch] carer/create unexpected success payload", {
      requestId,
      payload: response,
    });
    throw new Error("OneTouch create did not return a successful carer id.");
  }
  console.info("[OneTouch] carer/create success", {
    requestId,
    id,
    external_id: createPayload.external_id || "",
  });
  if (ONETOUCH_CREATE_DEBUG) {
    console.info("[OneTouch] carer/create success payload", {
      requestId,
      payload: response,
    });
  }
  return {
    id,
    raw: response,
  };
}

module.exports = {
  listCarers,
  listCarersDetailed,
  listClients,
  listVisits,
  listTimesheets,
  createCarer,
  resolveOneTouchLocationArea,
  getOneTouchLocationAreaOptions,
  getOneTouchAreaOptions,
  getOneTouchRecruitmentSourceOptions,
  getOneTouchPositionOptions,
  getOneTouchStatusOptions,
};
