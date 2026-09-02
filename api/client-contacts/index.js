const { readDirectoryData } = require("../_lib/directory-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

const ALLOWED_ROLES = [
  "admin",
  "care_manager",
  "operations",
  "clients_only",
  "hr_clients",
  "time_clients",
  "time_hr_clients",
];

function isActive(status) {
  return String(status || "").trim().toLowerCase() === "active";
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ALLOWED_ROLES }))) {
    return;
  }

  try {
    const directory = await readDirectoryData();
    const contacts = directory.clients
      .filter((client) => isActive(client?.status))
      .map((client) => ({
        name: String(client?.name || "").trim(),
        email: String(client?.email || "").trim(),
      }))
      .filter((client) => client.name)
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));

    res.setHeader("Cache-Control", "private, max-age=30");
    res.status(200).json({
      contacts,
      total: contacts.length,
      warnings: Array.isArray(directory.warnings) ? directory.warnings : [],
    });
  } catch (error) {
    res.status(500).json({ error: "Server error", detail: error?.message || String(error) });
  }
};
