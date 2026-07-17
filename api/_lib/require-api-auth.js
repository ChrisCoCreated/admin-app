const crypto = require("crypto");
const { getAuthorizedUsersMap } = require("./authorized-users");

const openIdConfigCache = new Map();
const jwksCache = new Map();
const authorizedUsers = getAuthorizedUsersMap();
const API_AUTH_DEBUG = process.env.API_AUTH_DEBUG === "1";
const ROLE_PAGES = {
  admin: [
    "clients",
    "clientdata",
    "carers",
    "timesheets",
    "recruitment",
    "wellbeingintake",
    "enquiries",
    "agendas",
    "problems",
    "scorecard",
    "scorecarddefinitions",
    "scorecardgoals",
    "whiteboard",
    "simpletasks",
    "tasks",
    "taskstest",
    "enquiryremindertest",
    "mapping",
    "drivetime",
    "reports",
    "finance",
    "functions",
    "emailtemplates",
    "suppliers",
    "consultant",
    "marketing",
    "marketingreports",
    "photolayout",
  ],
  care_manager: [
    "clients",
    "clientdata",
    "carers",
    "timesheets",
    "recruitment",
    "wellbeingintake",
    "enquiries",
    "agendas",
    "scorecard",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "functions",
    "emailtemplates",
    "suppliers",
  ],
  operations: [
    "clients",
    "clientdata",
    "carers",
    "timesheets",
    "recruitment",
    "wellbeingintake",
    "enquiries",
    "agendas",
    "scorecard",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "functions",
    "emailtemplates",
    "suppliers",
  ],
  finance: ["finance"],
  consultant: ["consultant", "agendas"],
  director: ["agendas", "finance", "scorecard", "scorecarddefinitions", "scorecardgoals", "suppliers", "wellbeingintake"],
  marketing: ["marketing", "marketingreports", "photolayout", "functions", "emailtemplates", "agendas"],
  photo_layout: ["photolayout", "agendas"],
  time_only: ["timesheets", "mapping", "drivetime", "agendas"],
  hr_only: ["carers", "timesheets", "recruitment", "agendas"],
  clients_only: ["clients", "clientdata", "agendas"],
  enquiries_only: ["enquiries", "agendas"],
  hr_clients: ["clients", "clientdata", "carers", "timesheets", "recruitment", "agendas"],
  time_clients: ["clients", "clientdata", "timesheets", "mapping", "drivetime", "agendas"],
  time_hr: ["carers", "timesheets", "recruitment", "mapping", "drivetime", "agendas"],
  time_hr_clients: ["clients", "clientdata", "carers", "timesheets", "recruitment", "mapping", "drivetime", "agendas"],
  logged_in: ["mapping"],
};
const ACCESS_PAGE_EXPANSIONS = {
  marketing: ["marketing", "marketingreports", "photolayout", "functions", "emailtemplates", "agendas"],
  photolayout: ["photolayout", "agendas"],
  finance: ["finance"],
  mapping: ["timesheets", "mapping", "drivetime", "agendas"],
  drivetime: ["timesheets", "mapping", "drivetime", "agendas"],
  carers: ["carers", "timesheets", "recruitment", "agendas"],
  clients: ["clients", "clientdata", "agendas"],
  enquiries: ["enquiries", "agendas"],
  consultant: ["consultant", "agendas"],
};

function logApiAuthDebug(message, details) {
  if (!API_AUTH_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[api-auth] ${message}`, details);
    return;
  }
  console.log(`[api-auth] ${message}`);
}

function normalizeRole(role) {
  return String(role || "").trim().toLowerCase();
}

function getDynamicAccessiblePages(role) {
  const normalizedRole = normalizeRole(role);
  if (!normalizedRole.startsWith("pages:")) {
    return [];
  }

  const pages = normalizedRole
    .slice("pages:".length)
    .split(",")
    .map((page) => String(page || "").trim().toLowerCase())
    .filter(Boolean);
  const accessible = new Set();

  for (const page of pages) {
    const expandedPages = ACCESS_PAGE_EXPANSIONS[page] || [page];
    for (const expandedPage of expandedPages) {
      accessible.add(expandedPage);
    }
  }

  return Array.from(accessible);
}

function getAccessiblePages(role) {
  const normalizedRole = normalizeRole(role);
  const pages = ROLE_PAGES[normalizedRole] || getDynamicAccessiblePages(normalizedRole);
  if (!Array.isArray(pages)) {
    return [];
  }
  const accessiblePages = [...pages];
  for (const sharedPage of ["mapping", "kpis"]) {
    if (!accessiblePages.includes(sharedPage)) {
      accessiblePages.push(sharedPage);
    }
  }
  return accessiblePages;
}

function roleMatchesAllowedRoles(role, allowedRoles) {
  const normalizedRole = normalizeRole(role);
  if (allowedRoles.includes(normalizedRole)) {
    return true;
  }

  const userPages = new Set(getAccessiblePages(normalizedRole));
  if (!userPages.size) {
    return false;
  }

  return allowedRoles.some((allowedRole) => {
    const requiredPages = getAccessiblePages(allowedRole);
    return requiredPages.length > 0 && requiredPages.every((page) => userPages.has(page));
  });
}

function base64UrlDecode(input) {
  const normalized = input.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  return Buffer.from(padded, "base64");
}

function parseJwt(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) {
    throw new Error("Invalid token format.");
  }

  const [headerB64, payloadB64, signatureB64] = parts;
  const header = JSON.parse(base64UrlDecode(headerB64).toString("utf8"));
  const payload = JSON.parse(base64UrlDecode(payloadB64).toString("utf8"));
  const signature = base64UrlDecode(signatureB64);

  return {
    signingInput: `${headerB64}.${payloadB64}`,
    header,
    payload,
    signature,
  };
}

async function fetchJson(url) {
  const response = await fetch(url);
  const data = await response.json();

  if (!response.ok) {
    const detail = data?.error_description || data?.error || `HTTP ${response.status}`;
    throw new Error(detail);
  }

  return data;
}

async function getOpenIdConfig(tenantId) {
  if (openIdConfigCache.has(tenantId)) {
    return openIdConfigCache.get(tenantId);
  }

  const promise = fetchJson(
    `https://login.microsoftonline.com/${tenantId}/v2.0/.well-known/openid-configuration`
  );
  openIdConfigCache.set(tenantId, promise);
  return promise;
}

async function getJwks(jwksUri) {
  if (jwksCache.has(jwksUri)) {
    return jwksCache.get(jwksUri);
  }

  const promise = fetchJson(jwksUri);
  jwksCache.set(jwksUri, promise);
  return promise;
}

function ensureClaimWindow(payload) {
  const now = Math.floor(Date.now() / 1000);

  if (typeof payload.exp === "number" && now >= payload.exp) {
    throw new Error("Token expired.");
  }

  if (typeof payload.nbf === "number" && now < payload.nbf) {
    throw new Error("Token not yet valid.");
  }
}

function ensureAudience(payload) {
  const configuredAudience = process.env.AZURE_API_AUDIENCE;
  const fallbackAudience = process.env.AZURE_API_CLIENT_ID;
  const accepted = new Set([configuredAudience, fallbackAudience].filter(Boolean));

  if (!accepted.size) {
    throw new Error("Server missing AZURE_API_AUDIENCE or AZURE_API_CLIENT_ID.");
  }

  if (!accepted.has(payload.aud)) {
    throw new Error("Token audience mismatch.");
  }
}

function ensureIssuer(payload, tenantId) {
  const validIssuers = new Set([
    `https://login.microsoftonline.com/${tenantId}/v2.0`,
    `https://sts.windows.net/${tenantId}/`,
  ]);

  if (!validIssuers.has(payload.iss)) {
    throw new Error("Token issuer mismatch.");
  }
}

function ensureScope(payload) {
  const required = (process.env.AZURE_REQUIRED_SCOPE || "client.read").trim();
  if (!required) {
    return;
  }

  const scopes = String(payload.scp || "").split(/\s+/).filter(Boolean);
  if (!scopes.includes(required)) {
    throw new Error("Required scope missing.");
  }
}

function verifySignature(signingInput, signature, jwk) {
  if (!jwk || jwk.kty !== "RSA") {
    throw new Error("No suitable signing key.");
  }

  const keyObject = crypto.createPublicKey({
    key: jwk,
    format: "jwk",
  });

  const verifier = crypto.createVerify("RSA-SHA256");
  verifier.update(signingInput);
  verifier.end();

  if (!verifier.verify(keyObject, signature)) {
    throw new Error("Invalid token signature.");
  }
}

async function validateBearerToken(token) {
  const tenantId = process.env.AZURE_TENANT_ID;
  if (!tenantId) {
    throw new Error("Server missing AZURE_TENANT_ID.");
  }

  const { signingInput, header, payload, signature } = parseJwt(token);
  if (header.alg !== "RS256") {
    throw new Error("Unsupported token algorithm.");
  }

  const openIdConfig = await getOpenIdConfig(tenantId);
  const jwks = await getJwks(openIdConfig.jwks_uri);
  const keys = Array.isArray(jwks?.keys) ? jwks.keys : [];
  const jwk = keys.find((key) => key.kid === header.kid);

  verifySignature(signingInput, signature, jwk);
  ensureIssuer(payload, tenantId);
  ensureAudience(payload);
  ensureClaimWindow(payload);
  ensureScope(payload);

  return payload;
}

function resolveUserEmail(claims) {
  return String(claims?.preferred_username || claims?.email || claims?.upn || "")
    .trim()
    .toLowerCase();
}

async function requireApiAuth(req, res, options = {}) {
  const allowedRoles = Array.isArray(options.allowedRoles)
    ? options.allowedRoles.map((role) => String(role).trim().toLowerCase()).filter(Boolean)
    : null;
  const authHeader = String(req.headers.authorization || "");
  const match = /^Bearer\s+(.+)$/i.exec(authHeader);

  if (!match) {
    logApiAuthDebug("Authorization header missing bearer token.");
    res.status(401).json({ error: "Missing bearer token." });
    return null;
  }

  try {
    const claims = await validateBearerToken(match[1]);
    const email = resolveUserEmail(claims);
    const role = authorizedUsers.get(email);
    logApiAuthDebug("Validated bearer token.", {
      email,
      role: role || "",
      aud: claims?.aud || "",
      iss: claims?.iss || "",
      scp: claims?.scp || "",
    });
    if (!email || !role) {
      logApiAuthDebug("Token resolved to an unauthorized email.", {
        email,
        knownUser: Boolean(role),
      });
      res.status(403).json({ error: "Forbidden." });
      return null;
    }
    if (allowedRoles && allowedRoles.length > 0 && !roleMatchesAllowedRoles(role, allowedRoles)) {
      logApiAuthDebug("Authenticated role is not permitted for route.", {
        email,
        role,
        allowedRoles,
      });
      res.status(403).json({ error: "Forbidden." });
      return null;
    }
    req.authUser = {
      email,
      role,
      claims,
    };
    return claims;
  } catch (error) {
    logApiAuthDebug("Bearer token rejected.", {
      reason: error?.message || String(error),
      routeMethod: req?.method || "",
    });
    res.status(401).json({ error: "Unauthorized." });
    return null;
  }
}

module.exports = {
  requireApiAuth,
};
