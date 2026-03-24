const { requireApiAuth } = require("../_lib/require-api-auth");
const { getAgendaShortcutPhoto, mapAgendaError } = require("../_lib/agendas/shortcuts");

module.exports = async (req, res) => {
  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin"] }))) {
    return;
  }

  try {
    if (req.method !== "GET") {
      res.status(405).json({
        error: { code: "METHOD_NOT_ALLOWED", message: "Method Not Allowed" },
      });
      return;
    }

    const photo = await getAgendaShortcutPhoto(req.query?.email || "");
    if (!photo) {
      res.status(404).json({
        error: { code: "SHORTCUT_PHOTO_NOT_FOUND", message: "Shortcut photo not found." },
      });
      return;
    }

    res.setHeader("Cache-Control", "private, max-age=604800");
    res.setHeader("Content-Type", photo.mimeType);
    res.status(200).end(photo.buffer);
  } catch (error) {
    const mapped = mapAgendaError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
