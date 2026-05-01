const { readCarersDirectoryData } = require("../_lib/directory-source");
const { createSharePointColleague, listSharePointColleagues } = require("../_lib/colleagues-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

function normalizeText(value) {
  return String(value || "").trim();
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    const carerId = normalizeText(req.body?.carerId);
    if (!carerId) {
      res.status(400).json({
        error: "Missing carerId.",
      });
      return;
    }

    const directory = await readCarersDirectoryData();
    const carer = Array.isArray(directory?.carers)
      ? directory.carers.find((entry) => normalizeText(entry?.id) === carerId)
      : null;

    if (!carer) {
      res.status(404).json({
        error: "Carer not found.",
      });
      return;
    }

    const colleagues = await listSharePointColleagues();
    const existing = colleagues.find((entry) => normalizeText(entry?.oneTouchId) === carerId) || null;
    if (existing) {
      res.status(200).json({
        success: true,
        alreadyExists: true,
        carerId,
        sharePointItemId: normalizeText(existing.itemId),
      });
      return;
    }

    const created = await createSharePointColleague({
      name: normalizeText(carer.name),
      oneTouchId: carerId,
      archived: false,
    });

    res.status(200).json({
      success: true,
      alreadyExists: false,
      carerId,
      sharePointItemId: normalizeText(created?.itemId),
    });
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: error?.message || "Could not add carer to SharePoint.",
    });
  }
};
