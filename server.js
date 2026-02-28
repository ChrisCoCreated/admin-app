const fs = require("fs");
const fsp = require("fs/promises");
const http = require("http");
const path = require("path");
const { URL } = require("url");

const clientsIndexHandler = require("./api/clients/index");
const clientsByIdHandler = require("./api/clients/[id]");
const clientsReconcilePreviewHandler = require("./api/clients/reconcile/preview");
const clientsReconcileApplyHandler = require("./api/clients/reconcile/apply");
const carersIndexHandler = require("./api/carers/index");
const oneTouchClientsHandler = require("./api/onetouch/clients");
const authMeHandler = require("./api/auth/me");
const routesRunHandler = require("./api/routes/run");
const marketingPhotosHandler = require("./api/marketing/photos");
const marketingMediaHandler = require("./api/marketing/media");
const tasksUnifiedHandler = require("./api/tasks/unified");
const tasksOverlayHandler = require("./api/tasks/overlay");
const tasksWhiteboardHandler = require("./api/tasks/whiteboard");
const tasksWhiteboardSyncHandler = require("./api/tasks/whiteboard-sync");
const tasksCreateHandler = require("./api/tasks/create");

function loadEnvFile(envPath) {
  let raw = "";
  try {
    raw = fs.readFileSync(envPath, "utf8");
  } catch {
    return;
  }

  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const eqIndex = trimmed.indexOf("=");
    if (eqIndex <= 0) {
      continue;
    }

    const key = trimmed.slice(0, eqIndex).trim();
    if (!key || process.env[key] !== undefined) {
      continue;
    }

    let value = trimmed.slice(eqIndex + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    process.env[key] = value;
  }
}

loadEnvFile(path.join(process.cwd(), ".env"));

const PORT = Number(process.env.PORT || 8081);
const HOST = process.env.HOST || "127.0.0.1";
const ROOT_DIR = process.cwd();

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".webmanifest": "application/manifest+json; charset=utf-8",
};

function sendText(res, statusCode, body) {
  res.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
  });
  res.end(body);
}

function getStaticFilePath(pathname) {
  const safePath = pathname === "/" ? "/index.html" : pathname;
  const decoded = decodeURIComponent(safePath);
  const normalized = path.normalize(decoded).replace(/^([.][.][/\\])+/, "");
  return path.join(ROOT_DIR, normalized);
}

async function serveStatic(pathname, res) {
  const absolutePath = getStaticFilePath(pathname);
  if (!absolutePath.startsWith(ROOT_DIR)) {
    sendText(res, 403, "Forbidden");
    return;
  }

  let stat;
  try {
    stat = await fsp.stat(absolutePath);
  } catch {
    sendText(res, 404, "Not Found");
    return;
  }

  if (stat.isDirectory()) {
    sendText(res, 404, "Not Found");
    return;
  }

  const ext = path.extname(absolutePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || "application/octet-stream";

  res.writeHead(200, {
    "Content-Type": contentType,
  });

  const stream = fs.createReadStream(absolutePath);
  stream.on("error", () => sendText(res, 500, "Internal Server Error"));
  stream.pipe(res);
}

function createApiResponseAdapter(nodeRes) {
  let statusCode = 200;

  return {
    status(code) {
      statusCode = Number(code) || 200;
      return this;
    },
    setHeader(name, value) {
      nodeRes.setHeader(name, value);
      return this;
    },
    json(payload) {
      if (!nodeRes.headersSent) {
        nodeRes.writeHead(statusCode, {
          "Content-Type": "application/json; charset=utf-8",
          "Cache-Control": "no-store",
        });
      }
      nodeRes.end(JSON.stringify(payload));
    },
  };
}

async function readJsonBody(req) {
  const chunks = [];
  let totalBytes = 0;
  const maxBytes = 1024 * 1024;

  for await (const chunk of req) {
    totalBytes += chunk.length;
    if (totalBytes > maxBytes) {
      throw new Error("Request body too large.");
    }
    chunks.push(chunk);
  }

  if (!chunks.length) {
    return {};
  }

  const raw = Buffer.concat(chunks).toString("utf8").trim();
  if (!raw) {
    return {};
  }

  return JSON.parse(raw);
}

async function handleApi(req, res, reqUrl) {
  const query = Object.fromEntries(reqUrl.searchParams.entries());
  const apiReq = {
    method: req.method,
    headers: req.headers,
    query,
    body: {},
  };
  const apiRes = createApiResponseAdapter(res);

  if (reqUrl.pathname === "/api/clients") {
    await clientsIndexHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/clients/reconcile/preview") {
    await clientsReconcilePreviewHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/clients/reconcile/apply") {
    if (req.method === "POST") {
      apiReq.body = await readJsonBody(req);
    }
    await clientsReconcileApplyHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/carers") {
    await carersIndexHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/onetouch/clients") {
    await oneTouchClientsHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/auth/me") {
    await authMeHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/routes/run") {
    if (req.method === "POST") {
      apiReq.body = await readJsonBody(req);
    }
    await routesRunHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/marketing/photos") {
    await marketingPhotosHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/marketing/media") {
    await marketingMediaHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/tasks/unified") {
    await tasksUnifiedHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/tasks/overlay") {
    if (req.method === "POST") {
      apiReq.body = await readJsonBody(req);
    }
    await tasksOverlayHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/tasks/whiteboard") {
    await tasksWhiteboardHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/tasks/whiteboard-sync") {
    if (req.method === "POST") {
      apiReq.body = await readJsonBody(req);
    }
    await tasksWhiteboardSyncHandler(apiReq, apiRes);
    return true;
  }

  if (reqUrl.pathname === "/api/tasks/create") {
    if (req.method === "POST") {
      apiReq.body = await readJsonBody(req);
    }
    await tasksCreateHandler(apiReq, apiRes);
    return true;
  }

  const byIdMatch = /^\/api\/clients\/([^/]+)$/.exec(reqUrl.pathname);
  if (byIdMatch) {
    apiReq.query.id = decodeURIComponent(byIdMatch[1]);
    await clientsByIdHandler(apiReq, apiRes);
    return true;
  }

  return false;
}

const server = http.createServer(async (req, res) => {
  try {
    const reqUrl = new URL(req.url, `http://${req.headers.host || "localhost"}`);

    if (reqUrl.pathname.startsWith("/api/")) {
      const handled = await handleApi(req, res, reqUrl);
      if (!handled) {
        sendText(res, 404, "Not Found");
      }
      return;
    }

    if (req.method !== "GET") {
      sendText(res, 405, "Method Not Allowed");
      return;
    }

    await serveStatic(reqUrl.pathname, res);
  } catch (error) {
    res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
    res.end(
      JSON.stringify({
        error: "Server error",
        detail: error && error.message ? error.message : String(error),
      })
    );
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Admin app listening on http://${HOST}:${PORT}`);
});
