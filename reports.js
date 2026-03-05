import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const monthPresetSelect = document.getElementById("monthPresetSelect");
const contractCapacityLink = document.getElementById("contractCapacityLink");
const dateRangeMessage = document.getElementById("dateRangeMessage");

const CONTRACT_CAPACITY_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carer/contractCapacity.php";

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "reports").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function formatDateParam(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function getMonthOffsetForPreset(preset) {
  if (preset === "last") {
    return -1;
  }
  if (preset === "next") {
    return 1;
  }
  return 0;
}

function getMonthRangeForPreset(preset) {
  const today = new Date();
  const monthOffset = getMonthOffsetForPreset(preset);
  const year = today.getFullYear();
  const month = today.getMonth() + monthOffset;

  const start = new Date(year, month, 1);
  const end = new Date(year, month + 1, 0);

  return { start, end };
}

function updateContractCapacityLink() {
  const preset = String(monthPresetSelect?.value || "this").trim().toLowerCase();
  const { start, end } = getMonthRangeForPreset(preset);
  const datePickSt = formatDateParam(start);
  const datePickFn = formatDateParam(end);

  const url = new URL(CONTRACT_CAPACITY_BASE_URL);
  url.searchParams.set("datePickSt", datePickSt);
  url.searchParams.set("datePickFn", datePickFn);
  url.searchParams.set("hoursCapacity", "hours");

  if (contractCapacityLink) {
    contractCapacityLink.href = url.toString();
  }
  if (dateRangeMessage) {
    dateRangeMessage.textContent = `Date range: ${datePickSt} to ${datePickFn}`;
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
    const role = String(profile?.role || "").trim().toLowerCase();

    if (!canAccessPage(role, "reports")) {
      redirectToUnauthorized("reports");
      return;
    }

    renderTopNavigation({ role });

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");

    if (monthPresetSelect) {
      monthPresetSelect.value = "this";
    }
    updateContractCapacityLink();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("reports");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

monthPresetSelect?.addEventListener("change", () => {
  updateContractCapacityLink();
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
