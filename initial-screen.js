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
const copyScreenSummaryBtn = document.getElementById("copyScreenSummaryBtn");
const scoreCountGreen = document.getElementById("scoreCountGreen");
const scoreCountAmber = document.getElementById("scoreCountAmber");
const scoreCountRed = document.getElementById("scoreCountRed");
const scoreCountUnscored = document.getElementById("scoreCountUnscored");
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
  screenOutcome: document.getElementById("screenOutcome"),
  screenNextSteps: document.getElementById("screenNextSteps"),
};

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let currentItemId = "";
let saveBusy = false;
let copyFeedbackTimer = 0;
const SCORE_FIELD_KEYS = ["q1Score", "q2Score", "q3Score", "q4Score", "q5Score", "q6Score", "q7Score"];

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

function getDraftStorageKey(itemId) {
  return `recruitment-initial-screen-draft:${cleanText(itemId)}`;
}

function loadLocalDraft(itemId) {
  const key = getDraftStorageKey(itemId);
  if (!key) {
    return null;
  }
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) {
      return null;
    }
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function saveLocalDraft() {
  if (!currentItemId) {
    return;
  }
  try {
    window.localStorage.setItem(
      getDraftStorageKey(currentItemId),
      JSON.stringify({
        updatedAt: new Date().toISOString(),
        responses: readForm(),
      })
    );
  } catch {
    // Ignore storage failures and keep the screening flow usable.
  }
}

function clearLocalDraft(itemId) {
  if (!itemId) {
    return;
  }
  try {
    window.localStorage.removeItem(getDraftStorageKey(itemId));
  } catch {
    // Ignore storage failures and keep the screening flow usable.
  }
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
  if (copyScreenSummaryBtn) {
    copyScreenSummaryBtn.disabled = !enabled || saveBusy;
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
  fieldRefs.screenOutcome.value = cleanText(responses.screenOutcome);
  fieldRefs.screenNextSteps.value = cleanText(responses.screenNextSteps);
  syncScoreChipGroup("q1Score", fieldRefs.q1Score.value);
  syncScoreChipGroup("q2Score", fieldRefs.q2Score.value);
  syncScoreChipGroup("q3Score", fieldRefs.q3Score.value);
  syncScoreChipGroup("q4Score", fieldRefs.q4Score.value);
  syncScoreChipGroup("q5Score", fieldRefs.q5Score.value);
  syncScoreChipGroup("q6Score", fieldRefs.q6Score.value);
  syncScoreChipGroup("q7Score", fieldRefs.q7Score.value);
  renderScoreSummary();
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
    screenOutcome: fieldRefs.screenOutcome.value,
    screenNextSteps: fieldRefs.screenNextSteps.value,
  };
}

function getScoreCounts() {
  const counts = {
    Green: 0,
    Amber: 0,
    Red: 0,
    Unscored: 0,
  };
  for (const key of SCORE_FIELD_KEYS) {
    const value = cleanText(fieldRefs[key]?.value);
    if (value === "Green" || value === "Amber" || value === "Red") {
      counts[value] += 1;
    } else {
      counts.Unscored += 1;
    }
  }
  return counts;
}

function renderScoreSummary() {
  const counts = getScoreCounts();
  if (scoreCountGreen) {
    scoreCountGreen.textContent = String(counts.Green);
  }
  if (scoreCountAmber) {
    scoreCountAmber.textContent = String(counts.Amber);
  }
  if (scoreCountRed) {
    scoreCountRed.textContent = String(counts.Red);
  }
  if (scoreCountUnscored) {
    scoreCountUnscored.textContent = String(counts.Unscored);
  }
}

function buildCopySummaryText() {
  const form = readForm();
  const counts = getScoreCounts();
  const lines = [
    `Initial Screen Summary: ${cleanText(screenCandidateName?.textContent) || "Candidate"}`,
    cleanText(screenCandidateMeta?.textContent) ? `Details: ${cleanText(screenCandidateMeta.textContent)}` : "",
    `Scores: Green ${counts.Green} | Amber ${counts.Amber} | Red ${counts.Red} | Unscored ${counts.Unscored}`,
    form.screenOutcome ? `Outcome: ${cleanText(form.screenOutcome)}` : "",
    form.screenNextSteps ? `Next steps: ${cleanText(form.screenNextSteps)}` : "",
    form.initialCallSummary ? `Call summary: ${cleanText(form.initialCallSummary)}` : "",
    `Link: ${window.location.href}`,
  ];
  return lines.filter(Boolean).join("\n");
}

function buildCopySummaryHtml() {
  const form = readForm();
  const counts = getScoreCounts();
  const candidateName = cleanText(screenCandidateName?.textContent) || "Candidate";
  const candidateMeta = cleanText(screenCandidateMeta?.textContent);
  const summaryText = cleanText(form.initialCallSummary);
  const screenOutcome = cleanText(form.screenOutcome);
  const screenNextSteps = cleanText(form.screenNextSteps);
  const pageUrl = window.location.href;

  return `
    <div style="font-family: Manrope, Segoe UI, Arial, sans-serif; color: #1c2433; line-height: 1.45;">
      <h2 style="margin: 0 0 8px; font-size: 20px;">Initial Screen Summary: ${escapeHtml(candidateName)}</h2>
      ${candidateMeta ? `<p style="margin: 0 0 12px; color: #5b6576;">${escapeHtml(candidateMeta)}</p>` : ""}
      <table style="border-collapse: collapse; margin: 0 0 14px;">
        <tr>
          <td style="padding: 6px 10px; border-radius: 999px; background: #ecfaf2; color: #0f6a3b; font-weight: 700;">Green: ${counts.Green}</td>
          <td style="width: 8px;"></td>
          <td style="padding: 6px 10px; border-radius: 999px; background: #fff6e5; color: #8a5a00; font-weight: 700;">Amber: ${counts.Amber}</td>
          <td style="width: 8px;"></td>
          <td style="padding: 6px 10px; border-radius: 999px; background: #fdecec; color: #a22a2a; font-weight: 700;">Red: ${counts.Red}</td>
          <td style="width: 8px;"></td>
          <td style="padding: 6px 10px; border-radius: 999px; background: #f5f7fa; color: #5c6676; font-weight: 700;">Unscored: ${counts.Unscored}</td>
        </tr>
      </table>
      ${
        screenOutcome
          ? `<p style="margin: 0 0 10px;"><strong>Outcome:</strong> ${escapeHtml(screenOutcome)}</p>`
          : ""
      }
      ${
        screenNextSteps
          ? `<div style="margin: 0 0 14px;">
              <div style="margin: 0 0 6px; font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; color: #1b5467; font-weight: 800;">Next Steps</div>
              <div style="padding: 12px 14px; border: 1px solid #d8e1ed; border-radius: 12px; background: #fbfdff; white-space: pre-wrap;">${escapeHtml(screenNextSteps)}</div>
            </div>`
          : ""
      }
      ${
        summaryText
          ? `<div style="margin: 0 0 14px;">
              <div style="margin: 0 0 6px; font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; color: #1b5467; font-weight: 800;">Call Summary</div>
              <div style="padding: 12px 14px; border: 1px solid #d8e1ed; border-radius: 12px; background: #fbfdff; white-space: pre-wrap;">${escapeHtml(summaryText)}</div>
            </div>`
          : ""
      }
      <p style="margin: 0;"><a href="${escapeHtml(pageUrl)}" style="color: #1f3f89; font-weight: 700;">Open Initial Screen</a></p>
    </div>
  `.trim();
}

function setCopyButtonState(state) {
  if (!copyScreenSummaryBtn) {
    return;
  }
  window.clearTimeout(copyFeedbackTimer);
  copyScreenSummaryBtn.classList.remove("is-success", "is-error");
  copyScreenSummaryBtn.textContent = "Copy Summary";
  if (state === "success") {
    copyScreenSummaryBtn.classList.add("is-success");
    copyScreenSummaryBtn.textContent = "Copied";
    copyFeedbackTimer = window.setTimeout(() => {
      copyScreenSummaryBtn.classList.remove("is-success");
      copyScreenSummaryBtn.textContent = "Copy Summary";
    }, 1800);
  } else if (state === "error") {
    copyScreenSummaryBtn.classList.add("is-error");
    copyScreenSummaryBtn.textContent = "Copy Failed";
    copyFeedbackTimer = window.setTimeout(() => {
      copyScreenSummaryBtn.classList.remove("is-error");
      copyScreenSummaryBtn.textContent = "Copy Summary";
    }, 1800);
  }
}

async function copyScreenSummary() {
  const text = buildCopySummaryText();
  const html = buildCopySummaryHtml();
  try {
    if (navigator.clipboard && window.ClipboardItem) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/plain": new Blob([text], { type: "text/plain" }),
          "text/html": new Blob([html], { type: "text/html" }),
        }),
      ]);
    } else {
      await navigator.clipboard.writeText(text);
    }
    setCopyButtonState("success");
    setStatus("Summary copied to clipboard.");
  } catch (error) {
    console.error(error);
    setCopyButtonState("error");
    setStatus("Could not copy summary to clipboard.", true);
  }
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
  const localDraft = loadLocalDraft(currentItemId);
  if (localDraft?.responses && typeof localDraft.responses === "object") {
    fillForm(localDraft.responses);
  }
  initialScreenForm.hidden = false;
  setFormEnabled(true);
  setStatus(localDraft ? "Initial screening form loaded. Restored local draft." : "Initial screening form loaded.");
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
    clearLocalDraft(currentItemId);
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
      renderScoreSummary();
      saveLocalDraft();
    });
  }
}

for (const field of Object.values(fieldRefs)) {
  if (!(field instanceof HTMLTextAreaElement) && !(field instanceof HTMLSelectElement)) {
    continue;
  }
  field.addEventListener("input", () => {
    renderScoreSummary();
    saveLocalDraft();
  });
  field.addEventListener("change", () => {
    renderScoreSummary();
    saveLocalDraft();
  });
}

copyScreenSummaryBtn?.addEventListener("click", async () => {
  if (copyScreenSummaryBtn.disabled) {
    return;
  }
  await copyScreenSummary();
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
