const { getAuthorizedUsersMap } = require("./authorized-users");

const GRAPH_AUTH_DEBUG = process.env.GRAPH_AUTH_DEBUG === "1";
const authorizedUsers = getAuthorizedUsersMap();
const ROLE_PAGES = {
  admin: [
    "clients",
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
    "finance",
    "emailtemplates",
    "suppliers",
    "consultant",
    "marketing",
    "photolayout",
  ],
  finance: ["finance"],
  consultant: ["consultant", "agendas"],
  director: ["agendas", "finance", "scorecard", "scorecarddefinitions", "scorecardgoals", "suppliers", "wellbeingintake"],
  marketing: ["marketing", "photolayout", "emailtemplates", "agendas"],
  photo_layout: ["photolayout", "agendas"],
  time_only: ["timesheets", "mapping", "drivetime", "agendas"],
  hr_only: ["carers", "timesheets", "recruitment", "agendas"],
  clients_only: ["clients", "agendas"],
  enquiries_only: ["enquiries", "agendas"],
  hr_clients: ["clients", "carers", "timesheets", "recruitment", "agendas"],
  time_clients: ["clients", "timesheets", "mapping", "drivetime", "agendas"],
  time_hr: ["carers", "timesheets", "recruitment", "mapping", "drivetime", "agendas"],
  time_hr_clients: ["clients", "carers", "timesheets", "recruitment", "mapping", "drivetime", "agendas"],
  logged_in: ["drivetime"],
};
const ACCESS_PAGE_EXPANSIONS = {
  marketing: ["marketing", "photolayout", "emailtemplates", "agendas"],
  photolayout: ["photolayout", "agendas"],
  finance: ["finance"],
  mapping: ["timesheets", "mapping", "drivetime", "agendas"],
  drivetime: ["timesheets", "mapping", "drivetime", "agendas"],
  carers: ["carers", "timesheets", "recruitment", "agendas"],
  clients: ["clients", "agendas"],
  enquiries: ["enquiries", "agendas"],
  consultant: ["consultant", "agendas"],
};

function logGraphAuthDebug(message, details) {
  if (!GRAPH_AUTH_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[graph-auth] ${message}`, details);
    return;
  }
  console.log(`[graph-auth] ${message}`);
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
  for (const sharedPage of ["drivetime", "kpis"]) {
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

function parseJwtPayload(token) {
  try {
    const parts = String(token || "").split(".");
    if (parts.length !== 3) {
      return null;
    }
    const payload = JSON.parse(base64UrlDecode(parts[1]).toString("utf8"));
    return payload;
  } catch {
    return null;
  }
}

async function fetchGraphMe(token) {
  const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,mail", {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  const text = await response.text();
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = null;
  }

  if (!response.ok) {
    const detail = data?.error?.message || text || `Graph /me failed (${response.status}).`;
    const error = new Error(detail);
    error.status = response.status;
    error.code = data?.error?.code || "GRAPH_ME_FAILED";
    throw error;
  }

  return data;
}

function resolveUserEmail({ graphProfile, tokenPayload }) {
  const fromGraph = String(graphProfile?.userPrincipalName || graphProfile?.mail || "")
    .trim()
    .toLowerCase();
  if (fromGraph) {
    return fromGraph;
  }

  return String(
    tokenPayload?.preferred_username || tokenPayload?.email || tokenPayload?.upn || ""
  )
    .trim()
    .toLowerCase();
}

async function requireGraphAuth(req, res, options = {}) {
  const allowedRoles = Array.isArray(options.allowedRoles)
    ? options.allowedRoles.map((role) => String(role).trim().toLowerCase()).filter(Boolean)
    : null;
  const authHeader = String(req.headers.authorization || "");
  const match = /^Bearer\s+(.+)$/i.exec(authHeader);

  if (!match) {
    logGraphAuthDebug("Authorization header missing bearer token.");
    res.status(401).json({
      error: { code: "UNAUTHORIZED", message: "Missing bearer token." },
    });
    return null;
  }

  const token = match[1];
  const tokenPayload = parseJwtPayload(token);

  logGraphAuthDebug("Checking delegated Graph token via /me.", {
    aud: tokenPayload?.aud || "",
    iss: tokenPayload?.iss || "",
    upn: tokenPayload?.upn || tokenPayload?.preferred_username || tokenPayload?.email || "",
    scp: tokenPayload?.scp || "",
  });

  try {
    const graphProfile = await fetchGraphMe(token);
    const email = resolveUserEmail({ graphProfile, tokenPayload });
    const role = authorizedUsers.get(email);

    if (!email || !role) {
      logGraphAuthDebug("Graph-authenticated user is not authorized in app allowlist.", {
        email,
        graphUserId: graphProfile?.id || "",
      });
      res.status(403).json({
        error: { code: "FORBIDDEN", message: "Forbidden." },
      });
      return null;
    }

    if (allowedRoles && allowedRoles.length > 0 && !roleMatchesAllowedRoles(role, allowedRoles)) {
      logGraphAuthDebug("Graph-authenticated user role not permitted for route.", {
        email,
        role,
        allowedRoles,
      });
      res.status(403).json({
        error: { code: "FORBIDDEN", message: "Forbidden." },
      });
      return null;
    }

    req.authUser = {
      email,
      role,
      claims: tokenPayload || {},
      graphProfile,
      graphAccessToken: token,
    };

    return req.authUser.claims;
  } catch (error) {
    logGraphAuthDebug("Delegated Graph token rejected by Graph /me.", {
      reason: error?.message || String(error),
      code: error?.code || "",
      status: error?.status || 0,
      routeMethod: req?.method || "",
    });

    res.status(401).json({
      error: {
        code: "TOKEN_EXPIRED_OR_INVALID",
        message: GRAPH_AUTH_DEBUG
          ? `Unauthorized: ${error?.message || "Graph token rejected."}`
          : "Unauthorized.",
      },
    });
    return null;
  }
}

module.exports = {
  requireGraphAuth,
};
