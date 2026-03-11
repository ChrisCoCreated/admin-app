import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260311";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const periodPresetSelect = document.getElementById("periodPresetSelect");
const contractCapacityLink = document.getElementById("contractCapacityLink");
const availabilityCapacityLink = document.getElementById("availabilityCapacityLink");
const areaCapacityLink = document.getElementById("areaCapacityLink");
const clientHoursLink = document.getElementById("clientHoursLink");
const carerHoursLink = document.getElementById("carerHoursLink");
const dateRangeMessage = document.getElementById("dateRangeMessage");

const CONTRACT_CAPACITY_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carer/contractCapacity.php";
const AVAILABILITY_CAPACITY_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carer/availabilityCapacity.php";
const AREA_CAPACITY_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carer/areaCapacity.php";
const CLIENT_HOURS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/clientsHoursRpt.php";
const CARER_HOURS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carersHoursRpt.php";

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

function toStartOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function getMondayOfWeek(baseDate) {
  const date = toStartOfDay(baseDate);
  const day = date.getDay();
  const daysSinceMonday = (day + 6) % 7;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() - daysSinceMonday);
}

function addDays(baseDate, days) {
  return new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + days);
}

function getDateRangeForPreset(preset) {
  const today = new Date();

  if (preset === "last_month" || preset === "this_month" || preset === "next_month") {
    let monthOffset = 0;
    if (preset === "last_month") {
      monthOffset = -1;
    }
    if (preset === "next_month") {
      monthOffset = 1;
    }

    const year = today.getFullYear();
    const month = today.getMonth() + monthOffset;
    const start = new Date(year, month, 1);
    const end = new Date(year, month + 1, 0);
    return { start, end };
  }

  const thisWeekMonday = getMondayOfWeek(today);
  if (preset === "last_week") {
    const start = addDays(thisWeekMonday, -7);
    const end = addDays(start, 6);
    return { start, end };
  }
  if (preset === "next_week") {
    const start = addDays(thisWeekMonday, 7);
    const end = addDays(start, 6);
    return { start, end };
  }
  const start = thisWeekMonday;
  const end = addDays(start, 6);
  return { start, end };
}

function buildCapacityUrl(baseUrl, start, end) {
  const datePickSt = formatDateParam(start);
  const datePickFn = formatDateParam(end);

  const url = new URL(baseUrl);
  url.searchParams.set("datePickSt", datePickSt);
  url.searchParams.set("datePickFn", datePickFn);
  url.searchParams.set("hoursCapacity", "hours");
  return url.toString();
}

function buildAreaCapacityUrl(start, end) {
  const datePickSt = formatDateParam(start);
  const datePickFn = formatDateParam(end);

  const url = new URL(AREA_CAPACITY_BASE_URL);
  url.searchParams.set("datePickSt", datePickSt);
  url.searchParams.set("datePickFn", datePickFn);
  return url.toString();
}

function buildClientHoursUrl(start, end) {
  const startDate = formatDateParam(start);
  const endDate = formatDateParam(end);

  const url = new URL(CLIENT_HOURS_BASE_URL);
  url.searchParams.set("selectJob", "all");
  url.searchParams.set("startDate", startDate);
  url.searchParams.set("endDate", endDate);
  return url.toString();
}

function buildCarerHoursUrl(start, end) {
  const startDate = formatDateParam(start);
  const endDate = formatDateParam(end);

  const url = new URL(CARER_HOURS_BASE_URL);
  url.searchParams.set("jobtype", "All");
  url.searchParams.set("start", startDate);
  url.searchParams.set("finish", endDate);
  return url.toString();
}

function updateCapacityLinks() {
  const preset = String(periodPresetSelect?.value || "this_month").trim().toLowerCase();
  const { start, end } = getDateRangeForPreset(preset);
  const datePickSt = formatDateParam(start);
  const datePickFn = formatDateParam(end);

  if (contractCapacityLink) {
    contractCapacityLink.href = buildCapacityUrl(CONTRACT_CAPACITY_BASE_URL, start, end);
  }
  if (availabilityCapacityLink) {
    availabilityCapacityLink.href = buildCapacityUrl(AVAILABILITY_CAPACITY_BASE_URL, start, end);
  }
  if (areaCapacityLink) {
    areaCapacityLink.href = buildAreaCapacityUrl(start, end);
  }
  if (clientHoursLink) {
    clientHoursLink.href = buildClientHoursUrl(start, end);
  }
  if (carerHoursLink) {
    carerHoursLink.href = buildCarerHoursUrl(start, end);
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

    if (periodPresetSelect) {
      periodPresetSelect.value = "this_month";
    }
    updateCapacityLinks();
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

periodPresetSelect?.addEventListener("change", () => {
  updateCapacityLinks();
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
