const crypto = require("crypto");
const authorizedUsersConfig = require("../../data/authorized-users.json");

const openIdConfigCache = new Map();
const jwksCache = new Map();
const GRAPH_AUTH_DEBUG = process.env.GRAPH_AUTH_DEBUG === "1";
const authorizedUsers = new Map(
  (Array.isArray(authorizedUsersConfig?.users) ? authorizedUsersConfig.users : [])
    .map((entry) => {
      const email = String(entry?.email || "")
        .trim()
        .toLowerCase();
      const role = String(entry?.role || "")
        .trim()
        .toLowerCase();
      if (!email || !role) {
        return null;
      }
      return [email, role];
    })
    .filter(Boolean)
);

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

function ensureIssuer(payload, tenantId) {
  const validIssuers = new Set([
    `https://login.microsoftonline.com/${tenantId}/v2.0`,
    `https://sts.windows.net/${tenantId}/`,
  ]);

  if (!validIssuers.has(payload.iss)) {
    throw new Error("Token issuer mismatch.");
  }
}

function ensureGraphAudience(payload) {
  const configured = String(process.env.GRAPH_TOKEN_AUDIENCE || "").trim();
  const accepted = new Set(
    [
      configured,
      "https://graph.microsoft.com",
      "00000003-0000-0000-c000-000000000000",
    ].filter(Boolean)
  );

  if (!accepted.has(String(payload.aud || ""))) {
    throw new Error("Token audience mismatch for Graph.");
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
  logGraphAuthDebug("Validating delegated Graph token claims.", {
    aud: payload?.aud || "",
    iss: payload?.iss || "",
    upn: payload?.upn || payload?.preferred_username || payload?.email || "",
    scp: payload?.scp || "",
    rolesCount: Array.isArray(payload?.roles) ? payload.roles.length : 0,
  });
  if (header.alg !== "RS256") {
    throw new Error("Unsupported token algorithm.");
  }

  const openIdConfig = await getOpenIdConfig(tenantId);
  const jwks = await getJwks(openIdConfig.jwks_uri);
  const keys = Array.isArray(jwks?.keys) ? jwks.keys : [];
  const jwk = keys.find((key) => key.kid === header.kid);

  verifySignature(signingInput, signature, jwk);
  ensureIssuer(payload, tenantId);
  ensureGraphAudience(payload);
  ensureClaimWindow(payload);

  return payload;
}

function resolveUserEmail(claims) {
  return String(claims?.preferred_username || claims?.email || claims?.upn || "")
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

  try {
    const claims = await validateBearerToken(match[1]);
    const email = resolveUserEmail(claims);
    const role = authorizedUsers.get(email);

    if (!email || !role) {
      logGraphAuthDebug("Validated token belongs to unauthorized user.", {
        email,
      });
      res.status(403).json({
        error: { code: "FORBIDDEN", message: "Forbidden." },
      });
      return null;
    }

    if (allowedRoles && allowedRoles.length > 0 && !allowedRoles.includes(role)) {
      logGraphAuthDebug("Validated token user role not permitted for route.", {
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
      claims,
      graphAccessToken: match[1],
    };

    return claims;
  } catch (error) {
    logGraphAuthDebug("Delegated Graph token rejected.", {
      reason: error?.message || String(error),
      routeMethod: req?.method || "",
    });
    res.status(401).json({
      error: {
        code: "TOKEN_EXPIRED_OR_INVALID",
        message: GRAPH_AUTH_DEBUG
          ? `Unauthorized: ${error?.message || "Token validation failed."}`
          : "Unauthorized.",
      },
    });
    return null;
  }
}

module.exports = {
  requireGraphAuth,
};
