const { requireApiAuth } = require("./_lib/require-api-auth");
const { getScorecardView, normalizeReviewPeriod, saveScorecardReview } = require("./_lib/scorecard-repository");

const ALLOWED_ROLES = ["admin", "care_manager", "operations"];

module.exports = async (req, res) => {
  const method = String(req.method || "GET").toUpperCase();
  if (method !== "GET" && method !== "PUT") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  const claims = await requireApiAuth(req, res, { allowedRoles: ALLOWED_ROLES });
  if (!claims) {
    return;
  }

  try {
    const reviewPeriod =
      method === "GET" ? normalizeReviewPeriod(req.query?.reviewPeriod) : normalizeReviewPeriod(req.body?.reviewPeriod);

    const payload =
      method === "GET"
        ? await getScorecardView(reviewPeriod)
        : await saveScorecardReview(reviewPeriod, req.body?.scorecard || {}, req.authUser?.email || "");

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    const status = Number(error?.status) >= 400 ? Number(error.status) : 400;
    res.status(status).json({
      error: "Scorecard request failed.",
      detail: error?.detail || error?.message || "Unknown error.",
    });
  }
};
