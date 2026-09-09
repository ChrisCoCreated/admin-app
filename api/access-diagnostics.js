const { requireApiAuth } = require("./_lib/require-api-auth");
const { getAccessConfigDiagnostics } = require("./_lib/authorized-users");

const DIAGNOSTICS_ADMIN_EMAIL = "chris@planwithcare.co.uk";
const APP_VERSION = "2026.09.09.1";

function cleanText(value) {
  return String(value || "").trim();
}

function getIdentityCandidates(claims = {}) {
  return Array.from(
    new Set(
      [claims.preferred_username, claims.email, claims.upn]
        .map((value) => cleanText(value).toLowerCase())
        .filter(Boolean)
    )
  );
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  const claims = await requireApiAuth(req, res);
  if (!claims) {
    return;
  }

  const signedInEmail = cleanText(req.authUser?.email).toLowerCase();
  if (signedInEmail !== DIAGNOSTICS_ADMIN_EMAIL) {
    res.status(403).json({ error: "Forbidden." });
    return;
  }
  const checkedEmail = cleanText(req.query?.email || signedInEmail).toLowerCase();

  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    appVersion: APP_VERSION,
    deployment: {
      environment: cleanText(process.env.VERCEL_ENV) || "local",
      gitCommit: cleanText(process.env.VERCEL_GIT_COMMIT_SHA),
      deploymentId: cleanText(process.env.VERCEL_DEPLOYMENT_ID),
    },
    signedInIdentity: {
      resolvedEmail: signedInEmail,
      role: cleanText(req.authUser?.role),
      tokenEmailCandidates: getIdentityCandidates(claims),
    },
    accessCheck: {
      email: checkedEmail,
      configuration: getAccessConfigDiagnostics(checkedEmail),
    },
  });
};
