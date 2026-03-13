const { requireApiAuth } = require("./_lib/require-api-auth");
const { getScorecard, normalizeReviewPeriod, upsertScorecard } = require("./_lib/scorecard-repository");

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
    if (method === "GET") {
      const reviewPeriod = normalizeReviewPeriod(req.query?.reviewPeriod);
      const record = await getScorecard(reviewPeriod);
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        reviewPeriod,
        scorecard: record?.payload || null,
        meta: record
          ? {
              createdAt: record.created_at || null,
              updatedAt: record.updated_at || null,
              createdBy: record.created_by || null,
              updatedBy: record.updated_by || null,
            }
          : null,
      });
      return;
    }

    const reviewPeriod = normalizeReviewPeriod(req.body?.reviewPeriod);
    const record = await upsertScorecard({
      reviewPeriod,
      payload: req.body?.scorecard,
      email: req.authUser?.email || "",
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      reviewPeriod,
      scorecard: record?.payload || null,
      meta: record
        ? {
            createdAt: record.created_at || null,
            updatedAt: record.updated_at || null,
            createdBy: record.created_by || null,
            updatedBy: record.updated_by || null,
          }
        : null,
    });
  } catch (error) {
    const status = Number(error?.status) >= 400 ? Number(error.status) : 400;
    res.status(status).json({
      error: "Scorecard request failed.",
      detail: error?.detail || error?.message || "Unknown error.",
    });
  }
};
