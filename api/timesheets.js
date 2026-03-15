const { listTimesheets } = require("./_lib/onetouch-client");
const { requireApiAuth } = require("./_lib/require-api-auth");

const MAX_RANGE_DAYS = 60;

function parseDateValue(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    throw new Error("Dates must use YYYY-MM-DD format.");
  }
  const date = new Date(`${raw}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) {
    throw new Error("Invalid date supplied.");
  }
  return date;
}

function getInclusiveDayCount(startDate, endDate) {
  const diffMs = endDate.getTime() - startDate.getTime();
  return Math.floor(diffMs / 86_400_000) + 1;
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (
    !(await requireApiAuth(req, res, {
      allowedRoles: [
        "admin",
        "care_manager",
        "operations",
        "hr_only",
        "hr_clients",
        "time_only",
        "time_clients",
        "time_hr",
        "time_hr_clients",
      ],
    }))
  ) {
    return;
  }

  try {
    const carerId = String(req.query.carer_id || req.query.carerId || "").trim();
    if (!carerId) {
      res.status(400).json({
        error: "Bad Request",
        detail: "carer_id is required.",
      });
      return;
    }

    const date = String(req.query.date || "").trim();
    const dateStart = String(req.query.datestart || req.query.dateStart || "").trim();
    const dateFinish = String(req.query.datefinish || req.query.dateFinish || "").trim();
    const perPageRaw = Number(req.query.per_page || req.query.perPage || "200");
    const perPage = Number.isFinite(perPageRaw) ? Math.max(1, Math.min(perPageRaw, 500)) : 200;

    const hasDate = Boolean(date);
    const hasRangeStart = Boolean(dateStart);
    const hasRangeFinish = Boolean(dateFinish);

    if (hasDate && (hasRangeStart || hasRangeFinish)) {
      res.status(400).json({
        error: "Bad Request",
        detail: "Use either date or datestart/datefinish, not both.",
      });
      return;
    }

    if (hasRangeStart !== hasRangeFinish) {
      res.status(400).json({
        error: "Bad Request",
        detail: "datestart and datefinish must be supplied together.",
      });
      return;
    }

    if (hasRangeStart && hasRangeFinish) {
      const startDate = parseDateValue(dateStart);
      const endDate = parseDateValue(dateFinish);
      if (startDate.getTime() > endDate.getTime()) {
        res.status(400).json({
          error: "Bad Request",
          detail: "datestart must be on or before datefinish.",
        });
        return;
      }
      if (getInclusiveDayCount(startDate, endDate) > MAX_RANGE_DAYS) {
        res.status(400).json({
          error: "Bad Request",
          detail: "Date ranges must be 60 days or fewer.",
        });
        return;
      }
    }

    if (hasDate) {
      parseDateValue(date);
    }

    const payload = await listTimesheets({
      carerId,
      date,
      dateStart,
      dateFinish,
      perPage,
    });

    console.info("[timesheets] API result", {
      carerId,
      date,
      dateStart,
      dateFinish,
      total: payload?.total || 0,
      rows: Array.isArray(payload?.timesheets) ? payload.timesheets.length : 0,
    });

    res.setHeader("Cache-Control", "private, max-age=30");
    res.status(200).json(payload);
  } catch (error) {
    const status = error?.status && Number.isFinite(error.status) ? error.status : 500;
    res.status(status).json({
      error: status >= 500 ? "Server error" : "Request failed",
      detail: error?.message || String(error),
    });
  }
};
