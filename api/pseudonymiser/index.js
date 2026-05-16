const { execFile } = require("child_process");
const path = require("path");
const { requireApiAuth } = require("../_lib/require-api-auth");

const PYTHON_BIN = process.env.PSEUDONYMISER_PYTHON || process.env.PYTHON || "python3";
const RUNNER_PATH = path.join(__dirname, "..", "_lib", "pseudonymiser", "runner.py");
const MAX_TEXT_LENGTH = 250000;

function normalizeAction(value) {
  return String(value || "").replace(/^\/+|\/+$/g, "");
}

function getRouteAction(req) {
  const queryAction = normalizeAction(req.query?.action);
  if (queryAction) {
    return queryAction;
  }

  const pathname = String(req.url || "").split("?")[0] || "";
  const matched = /\/api\/pseudonymiser\/([^/]+)\/?$/.exec(pathname);
  return normalizeAction(matched?.[1]);
}

function validateText(value) {
  if (typeof value !== "string" || !value.trim()) {
    return "Text is required.";
  }
  if (value.length > MAX_TEXT_LENGTH) {
    return "Text is too long.";
  }
  return "";
}

function parseRunnerError(stderr) {
  const raw = String(stderr || "").trim();
  if (!raw) {
    return "Pseudonymiser failed.";
  }
  try {
    const parsed = JSON.parse(raw.split(/\r?\n/).pop());
    return String(parsed?.error || "Pseudonymiser failed.");
  } catch {
    return "Pseudonymiser failed.";
  }
}

function runPseudonymiser(action, payload = {}) {
  return new Promise((resolve, reject) => {
    const child = execFile(
      PYTHON_BIN,
      [RUNNER_PATH, action],
      {
        cwd: path.dirname(RUNNER_PATH),
        encoding: "utf8",
        maxBuffer: 8 * 1024 * 1024,
        timeout: 15000,
      },
      (error, stdout, stderr) => {
        if (error) {
          reject(new Error(parseRunnerError(stderr)));
          return;
        }

        try {
          resolve(JSON.parse(stdout));
        } catch {
          reject(new Error("Pseudonymiser returned invalid JSON."));
        }
      }
    );

    child.stdin.end(JSON.stringify(payload));
  });
}

module.exports = async function pseudonymiserHandler(req, res) {
  const action = getRouteAction(req);

  if (action === "health") {
    res.status(200).json(await runPseudonymiser("health"));
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed." });
    return;
  }

  const claims = await requireApiAuth(req, res, { allowedRoles: ["clients_only"] });
  if (!claims) {
    return;
  }

  if (action !== "pseudonymise" && action !== "safety-check") {
    res.status(404).json({ error: "Not Found" });
    return;
  }

  const validationError = validateText(req.body?.text);
  if (validationError) {
    res.status(400).json({ error: validationError });
    return;
  }

  try {
    res.status(200).json(await runPseudonymiser(action, { text: req.body.text }));
  } catch (error) {
    res.status(500).json({ error: error?.message || "Pseudonymiser failed." });
  }
};
