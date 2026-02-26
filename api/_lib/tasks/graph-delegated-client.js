function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const GRAPH_TASKS_DEBUG = process.env.GRAPH_TASKS_DEBUG === "1";

function logGraphTasksDebug(message, details) {
  if (!GRAPH_TASKS_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[graph-tasks] ${message}`, details);
    return;
  }
  console.log(`[graph-tasks] ${message}`);
}

function parseRetryAfter(retryAfterValue) {
  const value = String(retryAfterValue || "").trim();
  if (!value) {
    return 0;
  }

  const asSeconds = Number(value);
  if (Number.isFinite(asSeconds) && asSeconds >= 0) {
    return Math.round(asSeconds * 1000);
  }

  const asDateMs = Date.parse(value);
  if (Number.isFinite(asDateMs)) {
    return Math.max(0, asDateMs - Date.now());
  }

  return 0;
}

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

function buildErrorFromResponse(response, payload, fallbackCode) {
  const graphCode = payload?.error?.code || fallbackCode;
  const graphMessage = payload?.error?.message || `Graph request failed (${response.status}).`;
  const retryable = response.status === 429 || response.status === 503 || response.status === 504;
  const error = new Error(graphMessage);
  error.status = response.status;
  error.code = graphCode;
  error.retryable = retryable;
  return error;
}

function buildHeaders(token, extraHeaders = {}) {
  return {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
    ...extraHeaders,
  };
}

function nextBackoffMs(attempt) {
  const jitter = Math.floor(Math.random() * 200);
  const base = Math.min(4000, 300 * 2 ** Math.max(0, attempt - 1));
  return base + jitter;
}

async function fetchJsonWithRetry(url, token, options = {}) {
  const {
    method = "GET",
    headers = {},
    body,
    maxAttempts = 5,
  } = options;

  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    let response;
    let payload;

    try {
      response = await fetch(url, {
        method,
        headers: buildHeaders(token, headers),
        body,
      });
    } catch (error) {
      logGraphTasksDebug("Network failure when calling Graph.", {
        method,
        url,
        attempt,
        maxAttempts,
        message: error?.message || String(error),
      });
      lastError = error;
      if (attempt >= maxAttempts) {
        throw error;
      }
      await sleep(nextBackoffMs(attempt));
      continue;
    }

    const text = await response.text();
    payload = parseJsonSafe(text);

    if (response.ok) {
      return payload;
    }

    const currentError = buildErrorFromResponse(response, payload, "GRAPH_REQUEST_FAILED");
    lastError = currentError;
    logGraphTasksDebug("Graph request returned non-OK response.", {
      method,
      url,
      status: response.status,
      graphCode: currentError?.code || "",
      graphMessage: currentError?.message || "",
      attempt,
      maxAttempts,
      retryable: currentError.retryable,
    });

    if (!(currentError.retryable && attempt < maxAttempts)) {
      throw currentError;
    }

    const retryAfterMs = parseRetryAfter(response.headers.get("Retry-After"));
    const waitMs = retryAfterMs > 0 ? retryAfterMs : nextBackoffMs(attempt);
    await sleep(waitMs);
  }

  throw lastError || new Error("Graph request failed.");
}

async function fetchAllPages(initialUrl, token, options = {}) {
  const items = [];
  let nextUrl = initialUrl;
  let pages = 0;
  const maxPages = Number.isFinite(options.maxPages) ? Math.max(1, options.maxPages) : 1000;

  while (nextUrl) {
    pages += 1;
    if (pages > maxPages) {
      const error = new Error("Graph pagination exceeded max pages.");
      error.status = 502;
      error.code = "GRAPH_PAGINATION_LIMIT";
      throw error;
    }

    const payload = await fetchJsonWithRetry(nextUrl, token, {
      method: "GET",
      maxAttempts: options.maxAttempts || 5,
    });

    const value = Array.isArray(payload?.value) ? payload.value : [];
    items.push(...value);
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }

  return items;
}

function createGraphDelegatedClient(token) {
  if (!token) {
    const error = new Error("Missing delegated Graph token.");
    error.status = 401;
    error.code = "TOKEN_EXPIRED_OR_INVALID";
    throw error;
  }

  return {
    fetchJson: (url, options = {}) => fetchJsonWithRetry(url, token, options),
    fetchAllPages: (url, options = {}) => fetchAllPages(url, token, options),
  };
}

module.exports = {
  createGraphDelegatedClient,
};
