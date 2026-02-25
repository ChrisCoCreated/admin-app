const { persistClientLocationFields } = require("../_lib/clients-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

function toValue(input) {
  return String(input || "").trim();
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  const targetId = String(req.query.id || req.body?.id || "").trim();
  if (!targetId) {
    res.status(400).json({ error: "Missing client id." });
    return;
  }

  const locationFields = {
    location: toValue(req.body?.location),
    address: toValue(req.body?.address),
  };

  if (!Object.values(locationFields).some(Boolean)) {
    res.status(400).json({ error: "No location fields to update." });
    return;
  }

  try {
    const updated = await persistClientLocationFields(targetId, locationFields);
    res.status(200).json({
      ok: true,
      updatedFields: updated.updatedFields,
    });
  } catch (error) {
    const statusCode = Number(error?.statusCode) || 500;
    res.status(statusCode).json({
      error: "Could not persist location fields.",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
