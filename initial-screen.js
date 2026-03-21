import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const signOutBtn = document.getElementById("signOutBtn");
const screenCandidateName = document.getElementById("screenCandidateName");
const screenCandidateFirstName = document.getElementById("screenCandidateFirstName");
const screenCandidateMeta = document.getElementById("screenCandidateMeta");
const screenStatusMessage = document.getElementById("screenStatusMessage");
const initialScreenForm = document.getElementById("initialScreenForm");
const saveInitialScreenBtn = document.getElementById("saveInitialScreenBtn");
const scoreChipGroups = Array.from(document.querySelectorAll(".score-chip-group"));

const fieldRefs = {
  q1NotesAvailability: document.getElementById("q1NotesAvailability"),
  q1Score: document.getElementById("q1Score"),
  q2NotesShortNotice: document.getElementById("q2NotesShortNotice"),
  q2Score: document.getElementById("q2Score"),
  q3NotesTravel: document.getElementById("q3NotesTravel"),
  q3Score: document.getElementById("q3Score"),
  q4NotesValuesFit: document.getElementById("q4NotesValuesFit"),
  q4Score: document.getElementById("q4Score"),
  q5NotesGoodCare: document.getElementById("q5NotesGoodCare"),
  q5Score: document.getElementById("q5Score"),
  q6NotesFlexibility: document.getElementById("q6NotesFlexibility"),
  q6Score: document.getElementById("q6Score"),
  q7NotesWellbeing: document.getElementById("q7NotesWellbeing"),
  q7Score: document.getElementById("q7Score"),
  initialCallSummary: document.getElementById("initialCallSummary"),
};

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let currentItemId = "";
let saveBusy = false;

function cleanText(value) {
  return String(value || "").trim();
}

function setStatus(message, isError = false) {
  if (!screenStatusMessage) {
    return;
  }
  screenStatusMessage.textContent = message;
  screenStatusMessage.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "recruitment").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function getItemIdFromUrl() {
  const url = new URL(window.location.href);
  return cleanText(url.searchParams.get("itemId"));
}

function setFormEnabled(enabled) {
  if (saveInitialScreenBtn) {
    saveInitialScreenBtn.disabled = !enabled || saveBusy;
  }
  for (const field of Object.values(fieldRefs)) {
    if (!field) {
      continue;
    }
    field.disabled = !enabled || saveBusy;
  }
  for (const group of scoreChipGroups) {
    for (const button of group.querySelectorAll(".score-chip")) {
      button.disabled = !enabled || saveBusy;
    }
  }
}

function syncScoreChipGroup(fieldId, value) {
  const group = document.querySelector(`.score-chip-group[data-score-field="${fieldId}"]`);
  if (!group) {
    return;
  }
  const selectedValue = cleanText(value);
  group.classList.toggle("has-selection", Boolean(selectedValue));
  for (const button of group.querySelectorAll(".score-chip")) {
    const isActive = cleanText(button.getAttribute("data-score-value")) === selectedValue;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
}

function fillForm(responses = {}) {
  fieldRefs.q1NotesAvailability.value = cleanText(responses.q1NotesAvailability);
  fieldRefs.q1Score.value = cleanText(responses.q1Score);
  fieldRefs.q2NotesShortNotice.value = cleanText(responses.q2NotesShortNotice);
  fieldRefs.q2Score.value = cleanText(responses.q2Score);
  fieldRefs.q3NotesTravel.value = cleanText(responses.q3NotesTravel);
  fieldRefs.q3Score.value = cleanText(responses.q3Score);
  fieldRefs.q4NotesValuesFit.value = cleanText(responses.q4NotesValuesFit);
  fieldRefs.q4Score.value = cleanText(responses.q4Score);
  fieldRefs.q5NotesGoodCare.value = cleanText(responses.q5NotesGoodCare);
  fieldRefs.q5Score.value = cleanText(responses.q5Score);
  fieldRefs.q6NotesFlexibility.value = cleanText(responses.q6NotesFlexibility);
  fieldRefs.q6Score.value = cleanText(responses.q6Score);
  fieldRefs.q7NotesWellbeing.value = cleanText(responses.q7NotesWellbeing);
  fieldRefs.q7Score.value = cleanText(responses.q7Score);
  fieldRefs.initialCallSummary.value = cleanText(responses.initialCallSummary);
  syncScoreChipGroup("q1Score", fieldRefs.q1Score.value);
  syncScoreChipGroup("q2Score", fieldRefs.q2Score.value);
  syncScoreChipGroup("q3Score", fieldRefs.q3Score.value);
  syncScoreChipGroup("q4Score", fieldRefs.q4Score.value);
  syncScoreChipGroup("q5Score", fieldRefs.q5Score.value);
  syncScoreChipGroup("q6Score", fieldRefs.q6Score.value);
  syncScoreChipGroup("q7Score", fieldRefs.q7Score.value);
}

function readForm() {
  return {
    q1NotesAvailability: fieldRefs.q1NotesAvailability.value,
    q1Score: fieldRefs.q1Score.value,
    q2NotesShortNotice: fieldRefs.q2NotesShortNotice.value,
    q2Score: fieldRefs.q2Score.value,
    q3NotesTravel: fieldRefs.q3NotesTravel.value,
    q3Score: fieldRefs.q3Score.value,
    q4NotesValuesFit: fieldRefs.q4NotesValuesFit.value,
    q4Score: fieldRefs.q4Score.value,
    q5NotesGoodCare: fieldRefs.q5NotesGoodCare.value,
    q5Score: fieldRefs.q5Score.value,
    q6NotesFlexibility: fieldRefs.q6NotesFlexibility.value,
    q6Score: fieldRefs.q6Score.value,
    q7NotesWellbeing: fieldRefs.q7NotesWellbeing.value,
    q7Score: fieldRefs.q7Score.value,
    initialCallSummary: fieldRefs.initialCallSummary.value,
  };
}

function renderCandidateHeader(item) {
  const candidateName = cleanText(item?.candidateName) || "Initial 10-Minute Call";
  const parts = [cleanText(item?.status), cleanText(item?.location), cleanText(item?.phoneNumber)].filter(Boolean);
  screenCandidateName.textContent = candidateName;
  screenCandidateFirstName.textContent = candidateName.split(/\s+/)[0] || "there";
  screenCandidateMeta.textContent = parts.length ? parts.join(" • ") : "Recruitment screening";
}

async function loadInitialScreen() {
  currentItemId = getItemIdFromUrl();
  if (!currentItemId) {
    setStatus("Missing recruitment item id.", true);
    return;
  }

  setFormEnabled(false);
  const payload = await directoryApi.getRecruitmentInitialScreen({ itemId: currentItemId });
  const item = payload?.item || null;
  if (!item?.itemId) {
    throw new Error("Candidate screening record could not be loaded.");
  }

  renderCandidateHeader(item);
  fillForm(item.responses || {});
  initialScreenForm.hidden = false;
  setFormEnabled(true);
  setStatus("Initial screening form loaded.");
}

async function saveInitialScreen(event) {
  event.preventDefault();
  if (saveBusy || !currentItemId) {
    return;
  }

  saveBusy = true;
  setFormEnabled(false);
  setStatus("Saving initial screening notes...");

  try {
    const result = await directoryApi.saveRecruitmentInitialScreen({
      itemId: currentItemId,
      responses: readForm(),
    });
    fillForm(result?.item?.responses || {});
    setStatus("Initial screening notes saved.");
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not save initial screening notes.", true);
  } finally {
    saveBusy = false;
    setFormEnabled(true);
  }
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
    await loadInitialScreen();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("recruitment");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not load the initial screening page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

initialScreenForm?.addEventListener("submit", saveInitialScreen);

for (const group of scoreChipGroups) {
  const fieldId = cleanText(group.getAttribute("data-score-field"));
  const input = fieldId ? document.getElementById(fieldId) : null;
  if (!(input instanceof HTMLInputElement)) {
    continue;
  }
  for (const button of group.querySelectorAll(".score-chip")) {
    button.addEventListener("click", () => {
      if (button.disabled) {
        return;
      }
      input.value = cleanText(button.getAttribute("data-score-value"));
      syncScoreChipGroup(fieldId, input.value);
    });
  }
}

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
