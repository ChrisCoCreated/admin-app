const { assignTagsToCarer } = require("../../_lib/onetouch-client");
const { readCarersDirectoryData } = require("../../_lib/directory-source");
const { requireApiAuth } = require("../../_lib/require-api-auth");

function normalizeIds(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return Array.from(
    new Set(
      values
        .map((value) => String(value || "").trim())
        .filter(Boolean)
    )
  );
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
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
        "time_hr",
        "time_hr_clients",
      ],
    }))
  ) {
    return;
  }

  const recordIds = normalizeIds(req.body?.recordIds);
  const tagId = Number(req.body?.tagId);

  if (!recordIds.length) {
    res.status(400).json({ error: "recordIds must contain at least one carer id." });
    return;
  }

  if (!Number.isFinite(tagId) || tagId <= 0) {
    res.status(400).json({ error: "tagId must be a positive number." });
    return;
  }

  try {
    const directory = await readCarersDirectoryData();
    const byId = new Map(directory.carers.map((carer) => [String(carer.id || "").trim(), carer]));
    const results = [];

    for (const recordId of recordIds) {
      const carer = byId.get(recordId);
      if (!carer) {
        results.push({
          id: recordId,
          name: "",
          ok: false,
          error: "Carer not found in current OneTouch directory.",
        });
        continue;
      }

      try {
        await assignTagsToCarer(recordId, { tagsToAssign: [tagId] });
        results.push({
          id: recordId,
          name: carer.name || "",
          ok: true,
        });
      } catch (error) {
        results.push({
          id: recordId,
          name: carer.name || "",
          ok: false,
          error: error?.message || String(error),
        });
      }
    }

    const succeeded = results.filter((item) => item.ok).length;
    const failed = results.length - succeeded;

    res.status(200).json({
      attempted: results.length,
      succeeded,
      failed,
      results,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
