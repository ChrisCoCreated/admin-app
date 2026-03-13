const { requireApiAuth } = require("../_lib/require-api-auth");
const {
  createDefinition,
  getDefinitionById,
  listDefinitions,
  normalizeEntityType,
  updateDefinition,
} = require("../_lib/scorecard-repository");

const ALLOWED_ROLES = ["admin", "care_manager", "operations"];

module.exports = async (req, res) => {
  const method = String(req.method || "GET").toUpperCase();
  if (!["GET", "POST", "PATCH"].includes(method)) {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  const claims = await requireApiAuth(req, res, { allowedRoles: ALLOWED_ROLES });
  if (!claims) {
    return;
  }

  try {
    if (method === "GET") {
      const entityType = normalizeEntityType(req.query?.entityType);
      const definitions = await listDefinitions(entityType, {
        activeOnly: String(req.query?.activeOnly || "").trim().toLowerCase() === "true",
      });
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        entityType,
        definitions,
      });
      return;
    }

    if (method === "POST") {
      const entityType = normalizeEntityType(req.body?.entityType);
      const definition = await createDefinition(entityType, req.body?.definition || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        entityType,
        definition,
      });
      return;
    }

    const entityType = normalizeEntityType(req.body?.entityType);
    const definition = await updateDefinition(entityType, req.body?.id, req.body?.patch || {});
    const hydrated = await getDefinitionById(entityType, definition.id);
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      entityType,
      definition: hydrated || definition,
    });
  } catch (error) {
    const status = Number(error?.status) >= 400 ? Number(error.status) : 400;
    res.status(status).json({
      error: "Scorecard definitions request failed.",
      detail: error?.detail || error?.message || "Unknown error.",
    });
  }
};
