const { URL } = require("url");

function getSupabaseConfig() {
  const url = String(process.env.SUPABASE_URL || "").trim().replace(/\/+$/, "");
  const serviceRoleKey = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();

  if (!url || !serviceRoleKey) {
    throw new Error("Server missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY.");
  }

  return { url, serviceRoleKey };
}

function buildUrl(pathname, query = {}) {
  const { url } = getSupabaseConfig();
  const target = new URL(`${url}/rest/v1/${pathname.replace(/^\/+/, "")}`);
  for (const [key, value] of Object.entries(query || {})) {
    if (value === undefined || value === null || value === "") {
      continue;
    }
    target.searchParams.set(key, String(value));
  }
  return target;
}

async function parseSupabaseError(response) {
  const text = await response.text();
  let payload = null;

  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  const detail =
    payload?.message ||
    payload?.error_description ||
    payload?.details ||
    payload?.hint ||
    text ||
    `HTTP ${response.status}`;

  const error = new Error(detail);
  error.status = response.status;
  error.detail = detail;
  error.payload = payload;
  throw error;
}

async function supabaseRestFetch(pathname, options = {}) {
  const { serviceRoleKey } = getSupabaseConfig();
  const response = await fetch(buildUrl(pathname, options.query), {
    method: options.method || "GET",
    headers: {
      Accept: "application/json",
      apikey: serviceRoleKey,
      Authorization: `Bearer ${serviceRoleKey}`,
      ...(options.headers || {}),
    },
    body: options.body === undefined ? undefined : JSON.stringify(options.body),
  });

  if (!response.ok) {
    await parseSupabaseError(response);
  }

  if (response.status === 204) {
    return null;
  }

  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

module.exports = {
  supabaseRestFetch,
};
