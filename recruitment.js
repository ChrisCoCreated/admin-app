import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const searchInput = document.getElementById("searchInput");
const locationFilterSelect = document.getElementById("locationFilterSelect");
const statusFilterSelect = document.getElementById("statusFilterSelect");
const sourceFilterSelect = document.getElementById("sourceFilterSelect");
const recruitmentTableBody = document.getElementById("recruitmentTableBody");
const emptyState = document.getElementById("emptyState");
const statusMessage = document.getElementById("statusMessage");
const signOutBtn = document.getElementById("signOutBtn");
const detailRoot = document.getElementById("candidateDetail");
const sharePointListLink = document.getElementById("sharePointListLink");
const addMissingToOneTouchBtn = document.getElementById("addMissingToOneTouchBtn");

const detailFields = {
  candidateName: detailRoot?.querySelector('[data-field="candidateName"]'),
  location: detailRoot?.querySelector('[data-field="location"]'),
  status: detailRoot?.querySelector('[data-field="status"]'),
  source: detailRoot?.querySelector('[data-field="source"]'),
  phoneNumber: detailRoot?.querySelector('[data-field="phoneNumber"]'),
  interviewBooked: detailRoot?.querySelector('[data-field="interviewBooked"]'),
  interviewWith: detailRoot?.querySelector('[data-field="interviewWith"]'),
  keepInMind: detailRoot?.querySelector('[data-field="keepInMind"]'),
  livesIn: detailRoot?.querySelector('[data-field="livesIn"]'),
  firstInterviewDate: detailRoot?.querySelector('[data-field="firstInterviewDate"]'),
  earmarkedFor: detailRoot?.querySelector('[data-field="earmarkedFor"]'),
  created: detailRoot?.querySelector('[data-field="created"]'),
  oneTouchLink: detailRoot?.querySelector('[data-field="oneTouchLink"]'),
  notes: detailRoot?.querySelector('[data-field="notes"]'),
};

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let allCandidates = [];
let selectedCandidateId = "";
let addToOneTouchBusy = false;

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function cleanText(value) {
  return String(value || "").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setStatus(message, isError = false) {
  if (!statusMessage) {
    return;
  }
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function hasOneTouchLink(candidate) {
  return Boolean(cleanText(candidate?.oneTouchLink));
}

function setAddButtonsBusy(disabled) {
  addToOneTouchBusy = disabled;
  if (addMissingToOneTouchBtn) {
    addMissingToOneTouchBtn.disabled = disabled;
  }
}

function updateAddMissingButtonLabel() {
  if (!addMissingToOneTouchBtn) {
    return;
  }
  const missingCount = allCandidates.filter((candidate) => !hasOneTouchLink(candidate)).length;
  addMissingToOneTouchBtn.textContent =
    missingCount > 0 ? `Add Missing to OneTouch (${missingCount})` : "Add Missing to OneTouch";
}

function formatBoolean(value) {
  return value === true ? "Yes" : "No";
}

function formatDate(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "-";
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return raw;
  }
  return parsed.toLocaleDateString();
}

function setLinkField(node, url) {
  if (!node) {
    return;
  }
  const cleanUrl = cleanText(url);
  if (!cleanUrl) {
    node.textContent = "-";
    return;
  }
  const safeUrl = escapeHtml(cleanUrl);
  node.innerHTML = `<a href="${safeUrl}" target="_blank" rel="noopener noreferrer">Open link</a>`;
}

function setDetail(candidate) {
  if (!candidate) {
    detailFields.candidateName.textContent = "Select a candidate";
    detailFields.location.textContent = "-";
    detailFields.status.textContent = "-";
    detailFields.source.textContent = "-";
    detailFields.phoneNumber.textContent = "-";
    detailFields.interviewBooked.textContent = "-";
    detailFields.interviewWith.textContent = "-";
    detailFields.keepInMind.textContent = "-";
    detailFields.livesIn.textContent = "-";
    detailFields.firstInterviewDate.textContent = "-";
    detailFields.earmarkedFor.textContent = "-";
    detailFields.created.textContent = "-";
    detailFields.oneTouchLink.textContent = "-";
    detailFields.notes.textContent = "-";
    return;
  }

  detailFields.candidateName.textContent = cleanText(candidate.candidateName) || "-";
  detailFields.location.textContent = cleanText(candidate.location) || "-";
  detailFields.status.textContent = cleanText(candidate.status) || "-";
  detailFields.source.textContent = cleanText(candidate.source) || "-";
  detailFields.phoneNumber.textContent = cleanText(candidate.phoneNumber) || "-";
  detailFields.interviewBooked.textContent = formatBoolean(candidate.interviewBooked);
  detailFields.interviewWith.textContent = cleanText(candidate.interviewWith) || "-";
  detailFields.keepInMind.textContent = formatBoolean(candidate.keepInMind);
  detailFields.livesIn.textContent = cleanText(candidate.livesIn) || "-";
  detailFields.firstInterviewDate.textContent = formatDate(candidate.firstInterviewDate);
  detailFields.earmarkedFor.textContent = cleanText(candidate.earmarkedFor) || "-";
  detailFields.created.textContent = formatDate(candidate.created);
  setLinkField(detailFields.oneTouchLink, candidate.oneTouchLink);
  detailFields.notes.textContent = cleanText(candidate.notes) || "-";
}

function renderFilterOptions() {
  const locationOptions = Array.from(
    new Set(allCandidates.map((candidate) => cleanText(candidate.location)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const statusOptions = Array.from(
    new Set(allCandidates.map((candidate) => cleanText(candidate.status)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const sourceOptions = Array.from(
    new Set(allCandidates.map((candidate) => cleanText(candidate.source)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const selectedLocation = cleanText(locationFilterSelect.value || "all");
  const selectedStatus = cleanText(statusFilterSelect.value || "all");
  const selectedSource = cleanText(sourceFilterSelect.value || "all");

  locationFilterSelect.innerHTML = '<option value="all">All locations</option>';
  statusFilterSelect.innerHTML = '<option value="all">All statuses</option>';
  sourceFilterSelect.innerHTML = '<option value="all">All sources</option>';

  for (const location of locationOptions) {
    const option = document.createElement("option");
    option.value = location;
    option.textContent = location;
    locationFilterSelect.appendChild(option);
  }
  for (const status of statusOptions) {
    const option = document.createElement("option");
    option.value = status;
    option.textContent = status;
    statusFilterSelect.appendChild(option);
  }
  for (const source of sourceOptions) {
    const option = document.createElement("option");
    option.value = source;
    option.textContent = source;
    sourceFilterSelect.appendChild(option);
  }

  locationFilterSelect.value = locationOptions.includes(selectedLocation) ? selectedLocation : "all";
  statusFilterSelect.value = statusOptions.includes(selectedStatus) ? selectedStatus : "all";
  sourceFilterSelect.value = sourceOptions.includes(selectedSource) ? selectedSource : "all";
}

function getFilteredCandidates() {
  const query = normalizeText(searchInput.value);
  const selectedLocation = cleanText(locationFilterSelect.value || "all");
  const selectedStatus = cleanText(statusFilterSelect.value || "all");
  const selectedSource = cleanText(sourceFilterSelect.value || "all");

  return allCandidates.filter((candidate) => {
    if (selectedLocation !== "all" && cleanText(candidate.location) !== selectedLocation) {
      return false;
    }
    if (selectedStatus !== "all" && cleanText(candidate.status) !== selectedStatus) {
      return false;
    }
    if (selectedSource !== "all" && cleanText(candidate.source) !== selectedSource) {
      return false;
    }
    if (!query) {
      return true;
    }
    return (
      normalizeText(candidate.candidateName).includes(query) ||
      normalizeText(candidate.location).includes(query) ||
      normalizeText(candidate.status).includes(query) ||
      normalizeText(candidate.source).includes(query) ||
      normalizeText(candidate.phoneNumber).includes(query) ||
      normalizeText(candidate.livesIn).includes(query) ||
      normalizeText(candidate.notes).includes(query)
    );
  });
}

function renderCandidates() {
  const filtered = getFilteredCandidates();
  recruitmentTableBody.innerHTML = "";

  if (!filtered.length) {
    emptyState.hidden = false;
    setDetail(null);
    updateAddMissingButtonLabel();
    return;
  }

  emptyState.hidden = true;
  const selected = filtered.find((candidate) => candidate.id === selectedCandidateId) || filtered[0];
  selectedCandidateId = selected.id;

  for (const candidate of filtered) {
    const tr = document.createElement("tr");
    tr.classList.toggle("selected", candidate.id === selectedCandidateId);
    tr.innerHTML = `
      <td>${escapeHtml(cleanText(candidate.candidateName) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.location) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.status) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.source) || "-")}</td>
      <td>${escapeHtml(cleanText(candidate.phoneNumber) || "-")}</td>
      <td>
        ${
          hasOneTouchLink(candidate)
            ? '<span class="muted">Added</span>'
            : `<button type="button" class="secondary recruitment-add-btn"${addToOneTouchBusy ? " disabled" : ""}>Add to OneTouch</button>`
        }
      </td>
    `;

    tr.addEventListener("click", () => {
      selectedCandidateId = candidate.id;
      setDetail(candidate);
      renderCandidates();
    });

    const addBtn = tr.querySelector(".recruitment-add-btn");
    addBtn?.addEventListener("click", async (event) => {
      event.stopPropagation();
      if (addToOneTouchBusy) {
        return;
      }
      await addCandidateToOneTouch(candidate.id);
    });

    recruitmentTableBody.appendChild(tr);
  }

  setDetail(selected);
  updateAddMissingButtonLabel();
}

function upsertCandidateInCache(item) {
  if (!item || !item.id) {
    return;
  }
  const index = allCandidates.findIndex((candidate) => candidate.id === item.id);
  if (index < 0) {
    allCandidates.push(item);
    return;
  }
  allCandidates[index] = item;
}

async function addCandidateToOneTouch(itemId) {
  const cleanItemId = cleanText(itemId);
  if (!cleanItemId) {
    return;
  }
  setAddButtonsBusy(true);
  try {
    setStatus("Adding candidate to OneTouch...");
    const result = await directoryApi.addRecruitmentCandidateToOneTouch({
      itemId: cleanItemId,
    });
    if (result?.item) {
      upsertCandidateInCache(result.item);
    }
    renderCandidates();
    setStatus(`Candidate added to OneTouch (ID: ${cleanText(result?.oneTouchId) || "-"})`);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not add candidate to OneTouch.", true);
  } finally {
    setAddButtonsBusy(false);
  }
}

async function addAllMissingCandidatesToOneTouch() {
  const missing = allCandidates.filter((candidate) => !hasOneTouchLink(candidate));
  if (!missing.length) {
    setStatus("All active candidates already have a OneTouch link.");
    return;
  }

  setAddButtonsBusy(true);
  let successCount = 0;
  try {
    for (const candidate of missing) {
      setStatus(`Adding ${cleanText(candidate.candidateName) || "candidate"} to OneTouch...`);
      const result = await directoryApi.addRecruitmentCandidateToOneTouch({
        itemId: candidate.id,
      });
      if (result?.item) {
        upsertCandidateInCache(result.item);
      }
      successCount += 1;
      renderCandidates();
    }
    setStatus(`Added ${successCount} candidate(s) to OneTouch.`);
  } catch (error) {
    console.error(error);
    setStatus(
      `Added ${successCount} candidate(s), then failed: ${error?.message || "OneTouch request failed."}`,
      true
    );
  } finally {
    setAddButtonsBusy(false);
  }
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "recruitment").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();
    if (!canAccessPage(role, "recruitment")) {
      redirectToUnauthorized("recruitment");
      return;
    }

    renderTopNavigation({ role });
    setStatus("Loading active candidates...");

    const payload = await directoryApi.listRecruitment();
    allCandidates = Array.isArray(payload?.items) ? payload.items : [];
    if (sharePointListLink) {
      sharePointListLink.href = cleanText(payload?.listUrl) || "#";
    }

    renderFilterOptions();
    renderCandidates();
    setStatus(`Loaded ${allCandidates.length} active candidate(s).`);
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("recruitment");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not load recruitment candidates.", true);
    emptyState.hidden = false;
    setDetail(null);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

searchInput?.addEventListener("input", renderCandidates);
locationFilterSelect?.addEventListener("change", renderCandidates);
statusFilterSelect?.addEventListener("change", renderCandidates);
sourceFilterSelect?.addEventListener("change", renderCandidates);
addMissingToOneTouchBtn?.addEventListener("click", async () => {
  if (addToOneTouchBusy) {
    return;
  }
  await addAllMissingCandidatesToOneTouch();
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
