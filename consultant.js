import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const actionStatus = document.getElementById("actionStatus");
const clientSelect = document.getElementById("clientSelect");
const clientAddressDisplay = document.getElementById("clientAddressDisplay");
const consultantNameInput = document.getElementById("consultantNameInput");
const notesInput = document.getElementById("notesInput");
const anonymiseBtn = document.getElementById("anonymiseBtn");
const anonOutput = document.getElementById("anonOutput");
const copyAnonBtn = document.getElementById("copyAnonBtn");
const reportEditor = document.getElementById("reportEditor");
const downloadDocxBtn = document.getElementById("downloadDocxBtn");
const deanonymiseReportBtn = document.getElementById("deanonymiseReportBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let allClients = [];
let selectedClient = null;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setActionStatus(message, isError = false) {
  actionStatus.textContent = message;
  actionStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "consultant").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function normalizeStatus(value) {
  return String(value || "").trim().toLowerCase();
}

function buildClientAddress(client) {
  if (!client || typeof client !== "object") {
    return "";
  }

  const directAddress = String(client.address || "").trim();
  if (directAddress) {
    return directAddress;
  }

  const fallbackParts = [client.town, client.county, client.postcode]
    .map((part) => String(part || "").trim())
    .filter(Boolean);
  return fallbackParts.join(", ");
}

function escapeRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceWithBoundary(sourceText, matcher) {
  return sourceText.replace(matcher, (full, lead) => `${lead || ""}P`);
}

function anonymiseText(rawText, clientName) {
  let output = String(rawText || "");
  const cleanName = String(clientName || "").trim();
  if (!output || !cleanName) {
    return output;
  }

  const tokens = cleanName
    .split(/[^A-Za-z0-9'-]+/)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);

  const fullParts = cleanName
    .split(/[^A-Za-z0-9'-]+/)
    .map((token) => token.trim())
    .filter(Boolean);

  if (fullParts.length >= 2) {
    const fuzzyFull = fullParts.map((part) => escapeRegex(part)).join("[\\s,.'’_-]*");
    const fullMatcher = new RegExp(`(^|[^A-Za-z0-9])(${fuzzyFull})(?=[^A-Za-z0-9]|$)`, "gi");
    output = replaceWithBoundary(output, fullMatcher);
  }

  const uniqueTokens = Array.from(new Set(tokens)).sort((a, b) => b.length - a.length);
  for (const token of uniqueTokens) {
    const matcher = new RegExp(`(^|[^A-Za-z0-9])(${escapeRegex(token)})(?=[^A-Za-z0-9]|$)`, "gi");
    output = replaceWithBoundary(output, matcher);
  }

  return output;
}

function deriveDisplayNameFromEmail(email) {
  const local = String(email || "")
    .trim()
    .split("@")[0]
    .replace(/[._-]+/g, " ")
    .trim();
  if (!local) {
    return "";
  }
  return local
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.slice(0, 1).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function sanitizeReportHtml() {
  const wrapper = document.createElement("div");
  wrapper.innerHTML = reportEditor.innerHTML;

  const allowed = new Set(["P", "BR", "STRONG", "B", "EM", "I", "UL", "OL", "LI", "H1", "H2", "H3"]);

  function escapeHtml(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = String(node.textContent || "");
      return escapeHtml(text);
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return "";
    }

    const tag = node.tagName.toUpperCase();
    const children = Array.from(node.childNodes).map(walk).join("");

    if (!allowed.has(tag)) {
      if (tag === "DIV") {
        return `<p>${children}</p>`;
      }
      return children;
    }

    if (tag === "BR") {
      return "<br>";
    }

    const lower = tag.toLowerCase();
    return `<${lower}>${children}</${lower}>`;
  }

  const html = Array.from(wrapper.childNodes).map(walk).join("").trim();
  return html || "<p></p>";
}

async function copyTextToClipboard(text) {
  if (navigator?.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const input = document.createElement("textarea");
  input.value = text;
  input.setAttribute("readonly", "");
  input.style.position = "fixed";
  input.style.opacity = "0";
  document.body.appendChild(input);
  input.select();
  const copied = document.execCommand("copy");
  document.body.removeChild(input);
  if (!copied) {
    throw new Error("Copy failed.");
  }
}

function applyClientSelection(clientId) {
  selectedClient = allClients.find((client) => String(client.id || "") === String(clientId || "")) || null;
  const address = buildClientAddress(selectedClient);
  clientAddressDisplay.textContent = address || "-";
}

function renderClientOptions() {
  clientSelect.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = allClients.length ? "Select a client" : "No clients available";
  clientSelect.appendChild(placeholder);

  for (const client of allClients) {
    const option = document.createElement("option");
    option.value = String(client.id || "");
    option.textContent = `${String(client.name || "Unnamed")} (${String(client.id || "-")})`;
    clientSelect.appendChild(option);
  }
}

async function loadClients() {
  const payload = await directoryApi.listOneTouchClients({ limit: 500 });
  const clients = Array.isArray(payload?.clients) ? payload.clients : [];
  allClients = clients
    .filter((client) => String(client?.name || "").trim())
    .sort((a, b) =>
      String(a.name || "").localeCompare(String(b.name || ""), undefined, {
        sensitivity: "base",
      })
    );

  renderClientOptions();
}

function handleAnonymise() {
  if (!selectedClient) {
    setActionStatus("Select a client first.", true);
    return;
  }

  const notes = String(notesInput.value || "").trim();
  if (!notes) {
    setActionStatus("Add notes before anonymising.", true);
    return;
  }

  const output = anonymiseText(notes, selectedClient.name);
  anonOutput.value = output;
  setActionStatus("Notes anonymised.");
}

async function handleCopyAnonymised() {
  const value = String(anonOutput.value || "").trim();
  if (!value) {
    setActionStatus("No anonymised notes to copy.", true);
    return;
  }

  try {
    await copyTextToClipboard(value);
    setActionStatus("Anonymised notes copied.");
  } catch (error) {
    console.error(error);
    setActionStatus("Could not copy anonymised notes.", true);
  }
}

function runEditorCommand(cmd, value = null) {
  reportEditor.focus();
  if (cmd === "formatBlock") {
    const tag = String(value || "P").toUpperCase();
    document.execCommand("formatBlock", false, `<${tag}>`);
    return;
  }
  document.execCommand(cmd, false, value);
}

function deanonymiseReport() {
  if (!selectedClient || !String(selectedClient.name || "").trim()) {
    setActionStatus("Select a client first.", true);
    return;
  }

  const clientName = String(selectedClient.name || "").trim();
  const walker = document.createTreeWalker(reportEditor, NodeFilter.SHOW_TEXT);
  const targets = [];
  let node = walker.nextNode();
  while (node) {
    targets.push(node);
    node = walker.nextNode();
  }

  let changed = 0;
  for (const textNode of targets) {
    const original = String(textNode.nodeValue || "");
    const next = original.replace(/\bP\b/g, clientName);
    if (next !== original) {
      textNode.nodeValue = next;
      changed += 1;
    }
  }

  if (changed > 0) {
    setActionStatus("Report deanonymised.");
  } else {
    setActionStatus("No anonymised P tokens found in report text.");
  }
}

async function handleDownloadDocx() {
  if (!selectedClient) {
    setActionStatus("Select a client first.", true);
    return;
  }

  const consultantName = String(consultantNameInput.value || "").trim();
  if (!consultantName) {
    setActionStatus("Consultant name is required.", true);
    return;
  }

  const clientName = String(selectedClient.name || "").trim();
  const clientAddress = buildClientAddress(selectedClient);
  if (!clientAddress) {
    setActionStatus("Client address is required for export.", true);
    return;
  }

  const reportHtml = sanitizeReportHtml();
  const reportText = String(reportEditor.textContent || "").trim();
  if (!reportText) {
    setActionStatus("Report text is required.", true);
    return;
  }

  try {
    downloadDocxBtn.disabled = true;
    downloadDocxBtn.classList.add("is-generating");
    setActionStatus("Generating Word document...");

    const blob = await directoryApi.exportConsultantReportDocx({
      consultantName,
      clientName,
      clientAddress,
      reportHtml,
    });

    const now = new Date();
    const filename = `consultant-report-${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}.docx`;
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    setActionStatus("Downloaded Word document.");
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not generate Word document.", true);
  } finally {
    downloadDocxBtn.disabled = false;
    downloadDocxBtn.classList.remove("is-generating");
  }
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await fetchCurrentUser();
    const role = normalizeStatus(profile?.role);

    if (!canAccessPage(role, "consultant")) {
      redirectToUnauthorized("consultant");
      return;
    }

    renderTopNavigation({ role });

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");

    if (!String(consultantNameInput?.value || "").trim()) {
      const fallbackName = String(account?.name || "").trim() || deriveDisplayNameFromEmail(email);
      consultantNameInput.value = fallbackName;
    }

    await loadClients();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("consultant");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize consultant page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

clientSelect?.addEventListener("change", () => {
  applyClientSelection(clientSelect.value);
  setActionStatus("");
});

anonymiseBtn?.addEventListener("click", () => {
  handleAnonymise();
});

copyAnonBtn?.addEventListener("click", () => {
  void handleCopyAnonymised();
});

document.querySelectorAll(".editor-btn").forEach((button) => {
  button.addEventListener("click", () => {
    const cmd = String(button.getAttribute("data-cmd") || "");
    const value = button.getAttribute("data-value");
    if (!cmd) {
      return;
    }
    runEditorCommand(cmd, value);
  });
});

downloadDocxBtn?.addEventListener("click", () => {
  void handleDownloadDocx();
});

deanonymiseReportBtn?.addEventListener("click", () => {
  deanonymiseReport();
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
