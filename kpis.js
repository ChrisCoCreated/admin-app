import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260512";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("kpiStatusMessage");
const refreshKpisBtn = document.getElementById("refreshKpisBtn");
const generateKpiPdfBtn = document.getElementById("generateKpiPdfBtn");
const showKpiDetailsInput = document.getElementById("showKpiDetailsInput");
const showKpiTrendsInput = document.getElementById("showKpiTrendsInput");
const kpiListLink = document.getElementById("kpiListLink");
const latestWeekLabel = document.getElementById("latestWeekLabel");
const refreshedAtLabel = document.getElementById("refreshedAtLabel");
const sourceRowCount = document.getElementById("sourceRowCount");
const deliveryKpis = document.getElementById("deliveryKpis");
const utilisationKpis = document.getElementById("utilisationKpis");
const businessKpis = document.getElementById("businessKpis");
const businessTrendKpis = document.getElementById("businessTrendKpis");
const enquiriesKpis = document.getElementById("enquiriesKpis");
const enquiriesTrendKpis = document.getElementById("enquiriesTrendKpis");
const recruitmentKpis = document.getElementById("recruitmentKpis");
const trainingKpis = document.getElementById("trainingKpis");
const cqcKpis = document.getElementById("cqcKpis");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let loadingKpis = false;
const KPI_VIEW_PREFS_KEY = "thrive.kpis.viewPrefs";

function cleanText(value) {
  return String(value ?? "").trim();
}

function loadViewPrefs() {
  try {
    const raw = window.localStorage?.getItem(KPI_VIEW_PREFS_KEY);
    const prefs = raw ? JSON.parse(raw) : null;
    return {
      showDetails: prefs?.showDetails !== false,
      showTrends: prefs?.showTrends !== false,
    };
  } catch {
    return {
      showDetails: true,
      showTrends: true,
    };
  }
}

function saveViewPrefs(prefs) {
  try {
    window.localStorage?.setItem(KPI_VIEW_PREFS_KEY, JSON.stringify(prefs));
  } catch {
    // Local storage is optional; the switches still work for the current page load.
  }
}

function readViewPrefsFromControls() {
  return {
    showDetails: showKpiDetailsInput ? showKpiDetailsInput.checked : true,
    showTrends: showKpiTrendsInput ? showKpiTrendsInput.checked : true,
  };
}

function applyViewPrefs(prefs) {
  const nextPrefs = prefs || readViewPrefsFromControls();
  if (showKpiDetailsInput) {
    showKpiDetailsInput.checked = nextPrefs.showDetails;
  }
  if (showKpiTrendsInput) {
    showKpiTrendsInput.checked = nextPrefs.showTrends;
  }
  document.body.classList.toggle("kpi-hide-details", !nextPrefs.showDetails);
  document.body.classList.toggle("kpi-hide-trends", !nextPrefs.showTrends);
  saveViewPrefs(nextPrefs);
}

function escapeHtml(value) {
  return cleanText(value)
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  const text = cleanText(value).replace(/,/g, "");
  if (!text) {
    return null;
  }
  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) {
    return null;
  }
  const number = Number(match[0]);
  return Number.isFinite(number) ? number : null;
}

function parsePercent(value) {
  const number = parseNumber(value);
  if (number === null) {
    return null;
  }
  if (typeof value === "number" && number > 0 && number <= 1) {
    return number * 100;
  }
  return number;
}

function formatNumber(value, decimals = 0) {
  const number = parseNumber(value);
  if (number === null) {
    return "-";
  }
  return new Intl.NumberFormat("en-GB", {
    maximumFractionDigits: decimals,
    minimumFractionDigits: decimals,
  }).format(number);
}

function formatHours(value) {
  const number = parseNumber(value);
  if (number === null) {
    return "-";
  }
  const decimals = Number.isInteger(number) ? 0 : 1;
  return `${formatNumber(number, decimals)}h`;
}

function formatPercent(value) {
  const percent = parsePercent(value);
  if (percent === null) {
    return "-";
  }
  const decimals = Number.isInteger(percent) ? 0 : 1;
  return `${formatNumber(percent, decimals)}%`;
}

function formatDateTime(value) {
  const date = new Date(value);
  if (!value || Number.isNaN(date.getTime())) {
    return "-";
  }
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function removeYearFromDateLabel(value) {
  return cleanText(value).replace(/\s+\d{4}$/, "");
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setLoading(isLoading) {
  loadingKpis = isLoading;
  if (refreshKpisBtn) {
    refreshKpisBtn.disabled = isLoading;
    refreshKpisBtn.textContent = isLoading ? "Refreshing..." : "Refresh";
  }
}

function staleLabel(metric) {
  if (!metric?.stale) {
    return "";
  }
  return metric.sourceWeekLabel ? `Stale, from ${removeYearFromDateLabel(metric.sourceWeekLabel)}` : "Stale data";
}

function sourceLabel(metric, latestWeekLabelText) {
  if (metric?.sourceLabel) {
    return metric.sourceLabel;
  }
  if (!metric?.sourceWeekLabel) {
    return latestWeekLabelText ? `No value found through ${latestWeekLabelText}` : "No value found";
  }
  if (metric.stale) {
    return `Using most recent filled row: ${metric.sourceWeekLabel}`;
  }
  return `From latest row: ${metric.sourceWeekLabel}`;
}

export function createKpiMetricCard({
  title,
  value,
  metric,
  detail = "",
  tone = "default",
  latestWeekLabelText = "",
  wide = false,
  trendCompanion = false,
}) {
  const card = document.createElement("article");
  card.className = `kpi-card kpi-card-${tone}${wide ? " kpi-card-wide" : ""}`;
  if (trendCompanion) {
    card.classList.add("kpi-trend-companion-tile");
  }

  const stale = staleLabel(metric);
  const detailText = cleanText(detail);
  card.innerHTML = `
    <div class="kpi-card-topline">
      <h3>${escapeHtml(title)}</h3>
      ${stale ? `<span class="kpi-stale-pill">${escapeHtml(stale)}</span>` : ""}
    </div>
    <div class="kpi-card-value">${escapeHtml(value)}</div>
    ${detailText ? `<p class="kpi-card-detail">${escapeHtml(detailText)}</p>` : ""}
    <p class="kpi-card-source">${escapeHtml(sourceLabel(metric, latestWeekLabelText))}</p>
  `;
  return card;
}

export function createKpiNoteCard({ title, metric, emptyLabel = "No detail recorded", latestWeekLabelText = "" }) {
  const value = cleanText(metric?.value);
  const card = createKpiMetricCard({
    title,
    value: value || emptyLabel,
    metric,
    latestWeekLabelText,
    wide: true,
  });
  card.classList.add("kpi-note-card", "kpi-detail-tile");
  return card;
}

function getTrendValues(series, field) {
  return (Array.isArray(series) ? series : [])
    .map((row) => ({
      weekLabel: row.weekLabel,
      value: parseNumber(row[field]),
    }))
    .filter((point) => point.value !== null);
}

function summarizeTrend(points) {
  if (!points.length) {
    return {
      min: null,
      max: null,
      average: null,
      maxPoint: null,
      minPoint: null,
    };
  }
  const values = points.map((point) => point.value);
  const max = Math.max(...values);
  const min = Math.min(...values);
  return {
    min,
    max,
    average: values.reduce((sum, value) => sum + value, 0) / values.length,
    maxPoint: points.find((point) => point.value === max) || null,
    minPoint: points.find((point) => point.value === min) || null,
  };
}

function trendDeltaLabel(points, formatter) {
  if (points.length < 2) {
    return "Not enough data";
  }
  const first = points[0].value;
  const latest = points[points.length - 1].value;
  const delta = latest - first;
  const sign = delta > 0 ? "+" : "";
  return `${sign}${formatter(delta)} over quarter`;
}

function buildSparkline(points) {
  if (!points.length) {
    return '<div class="kpi-sparkline-empty">No trend data</div>';
  }
  const values = points.map((point) => point.value);
  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = max - min || 1;
  const width = 160;
  const height = 48;
  const step = points.length > 1 ? width / (points.length - 1) : width;
  const coords = points.map((point, index) => {
    const x = points.length > 1 ? index * step : width / 2;
    const y = height - ((point.value - min) / range) * (height - 8) - 4;
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  });
  return `
    <svg class="kpi-sparkline" viewBox="0 0 ${width} ${height}" role="img" aria-label="Quarter trend">
      <polyline points="${coords.join(" ")}" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></polyline>
    </svg>
  `;
}

export function createKpiTrendCard({ title, series, field, formatter = formatNumber }) {
  const points = getTrendValues(series, field);
  const latest = points.length ? points[points.length - 1] : null;
  const summary = summarizeTrend(points);
  const card = document.createElement("button");
  card.type = "button";
  card.className = "kpi-trend-card";
  card.setAttribute("aria-label", `Open detailed ${title}`);
  card.innerHTML = `
    <div>
      <h3>${escapeHtml(title)}</h3>
      <p class="kpi-trend-value">${latest ? escapeHtml(formatter(latest.value)) : "-"}</p>
      <p class="kpi-card-source">${escapeHtml(latest?.weekLabel || "No trend data")}</p>
    </div>
    <dl class="kpi-trend-stats">
      <div>
        <dt>Max</dt>
        <dd>${summary.max === null ? "-" : escapeHtml(formatter(summary.max))}</dd>
      </div>
      <div>
        <dt>Avg</dt>
        <dd>${summary.average === null ? "-" : escapeHtml(formatter(summary.average))}</dd>
      </div>
      <div>
        <dt>Min</dt>
        <dd>${summary.min === null ? "-" : escapeHtml(formatter(summary.min))}</dd>
      </div>
    </dl>
    <div class="kpi-trend-visual">
      ${buildSparkline(points)}
      <span>${escapeHtml(trendDeltaLabel(points, formatter))}${
        summary.maxPoint ? ` | Max ${escapeHtml(formatter(summary.max))} on ${escapeHtml(summary.maxPoint.weekLabel)}` : ""
      }</span>
    </div>
  `;
  card.addEventListener("click", () => {
    openTrendModal({ title, points, formatter, summary });
  });
  return card;
}

function ensureTrendModal() {
  let modal = document.getElementById("kpiTrendModal");
  if (modal) {
    return modal;
  }

  modal = document.createElement("div");
  modal.id = "kpiTrendModal";
  modal.className = "kpi-modal";
  modal.hidden = true;
  modal.innerHTML = `
    <div class="kpi-modal-backdrop" data-kpi-modal-close></div>
    <section class="kpi-modal-panel" role="dialog" aria-modal="true" aria-labelledby="kpiTrendModalTitle">
      <div class="kpi-modal-head">
        <div>
          <p class="kpi-kicker">Quarter detail</p>
          <h2 id="kpiTrendModalTitle"></h2>
        </div>
        <button class="secondary kpi-modal-close" type="button" data-kpi-modal-close>Close</button>
      </div>
      <div id="kpiTrendModalStats" class="kpi-modal-stats"></div>
      <div id="kpiTrendModalChart" class="kpi-modal-chart"></div>
      <div id="kpiTrendModalRows" class="kpi-modal-rows"></div>
    </section>
  `;
  modal.addEventListener("click", (event) => {
    if (event.target instanceof Element && event.target.closest("[data-kpi-modal-close]")) {
      closeTrendModal();
    }
  });
  document.body.appendChild(modal);
  return modal;
}

function closeTrendModal() {
  const modal = document.getElementById("kpiTrendModal");
  if (!modal) {
    return;
  }
  modal.hidden = true;
  document.body.classList.remove("kpi-modal-open");
}

function buildModalBars(points, formatter) {
  if (!points.length) {
    return '<p class="muted">No trend data available.</p>';
  }
  const values = points.map((point) => point.value);
  const max = Math.max(...values) || 1;
  return points
    .map((point) => {
      const width = max > 0 ? Math.max(point.value === 0 ? 0 : 2, (point.value / max) * 100) : 0;
      return `
        <div class="kpi-modal-bar-row">
          <span>${escapeHtml(point.weekLabel)}</span>
          <div class="kpi-modal-bar-track">
            <div class="kpi-modal-bar" style="width: ${width}%"></div>
          </div>
          <strong>${escapeHtml(formatter(point.value))}</strong>
        </div>
      `;
    })
    .join("");
}

function openTrendModal({ title, points, formatter, summary }) {
  const modal = ensureTrendModal();
  const titleNode = modal.querySelector("#kpiTrendModalTitle");
  const statsNode = modal.querySelector("#kpiTrendModalStats");
  const chartNode = modal.querySelector("#kpiTrendModalChart");
  const rowsNode = modal.querySelector("#kpiTrendModalRows");

  if (titleNode) {
    titleNode.textContent = title;
  }
  if (statsNode) {
    statsNode.innerHTML = `
      <div><span>Latest</span><strong>${points.length ? escapeHtml(formatter(points[points.length - 1].value)) : "-"}</strong></div>
      <div><span>Max</span><strong>${summary.max === null ? "-" : escapeHtml(formatter(summary.max))}</strong><small>${escapeHtml(
        summary.maxPoint?.weekLabel || ""
      )}</small></div>
      <div><span>Average</span><strong>${summary.average === null ? "-" : escapeHtml(formatter(summary.average))}</strong></div>
      <div><span>Min</span><strong>${summary.min === null ? "-" : escapeHtml(formatter(summary.min))}</strong><small>${escapeHtml(
        summary.minPoint?.weekLabel || ""
      )}</small></div>
    `;
  }
  if (chartNode) {
    chartNode.innerHTML = buildSparkline(points);
  }
  if (rowsNode) {
    rowsNode.innerHTML = buildModalBars(points, formatter);
  }

  modal.hidden = false;
  document.body.classList.add("kpi-modal-open");
  modal.querySelector(".kpi-modal-close")?.focus();
}

function appendChildren(parent, children) {
  if (!parent) {
    return;
  }
  parent.innerHTML = "";
  for (const child of children) {
    parent.appendChild(child);
  }
}

function metricValue(payload, key) {
  return payload?.values?.[key] || null;
}

function renderSummary(payload) {
  latestWeekLabel.textContent = payload?.latestWeekLabel || "-";
  refreshedAtLabel.textContent = formatDateTime(payload?.refreshedAt);
  sourceRowCount.textContent = Number.isFinite(Number(payload?.rowCount)) ? String(payload.rowCount) : "-";
  if (kpiListLink && payload?.listUrl) {
    kpiListLink.href = payload.listUrl;
    kpiListLink.hidden = false;
  }
}

function renderDelivery(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  appendChildren(deliveryKpis, [
    createKpiMetricCard({
      title: "Hours Delivered %",
      value: formatPercent(metricValue(payload, "hoursDeliveredPercent")?.value),
      metric: metricValue(payload, "hoursDeliveredPercent"),
      detail: `${formatHours(metricValue(payload, "hoursDelivered")?.value)} delivered against ${formatHours(
        metricValue(payload, "activeContractedHours")?.value
      )} contracted`,
      tone: "delivery",
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiTrendCard({
      title: "Delivery % Trend",
      series: payload?.trendSeries,
      field: "hoursDeliveredPercent",
      formatter: formatPercent,
    }),
    createKpiMetricCard({
      title: "Hours Delivered",
      value: formatHours(metricValue(payload, "hoursDelivered")?.value),
      metric: metricValue(payload, "hoursDelivered"),
      detail: `${formatHours(metricValue(payload, "activeContractedHours")?.value)} contracted`,
      tone: "delivery",
      latestWeekLabelText: latestLabel,
    }),
    createKpiNoteCard({
      title: "Dropped Hours Detail",
      metric: metricValue(payload, "droppedHoursReasons"),
      emptyLabel: "No dropped-hours reason recorded",
      latestWeekLabelText: latestLabel,
    }),
  ]);
}

function renderUtilisation(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  appendChildren(utilisationKpis, [
    createKpiMetricCard({
      title: "Utilisation %",
      value: formatPercent(metricValue(payload, "utilisationPercent")?.value),
      metric: metricValue(payload, "utilisationPercent"),
      tone: "delivery",
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiTrendCard({
      title: "Utilisation % Trend",
      series: payload?.trendSeries,
      field: "utilisationPercent",
      formatter: formatPercent,
    }),
    createKpiNoteCard({
      title: "Utilisation Notes",
      metric: metricValue(payload, "utilisationNotes"),
      emptyLabel: "No utilisation notes recorded",
      latestWeekLabelText: latestLabel,
    }),
  ]);
}

function renderBusiness(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  appendChildren(businessKpis, [
    createKpiMetricCard({
      title: "Total Contracted Hours",
      value: formatHours(metricValue(payload, "totalHours")?.value),
      detail: `${formatHours(metricValue(payload, "hoursDelivered")?.value)} delivered + ${formatHours(
        metricValue(payload, "subscriptionHours")?.value
      )} subscription`,
      metric: metricValue(payload, "totalHours"),
      tone: "business",
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiMetricCard({
      title: "Hours Won",
      value: formatHours(metricValue(payload, "hoursWon")?.value),
      metric: metricValue(payload, "hoursWon"),
      tone: "positive",
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiMetricCard({
      title: "Hours Lost",
      value: formatHours(metricValue(payload, "hoursLost")?.value),
      metric: metricValue(payload, "hoursLost"),
      tone: "risk",
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiMetricCard({
      title: "Pending Hours",
      value: formatHours(metricValue(payload, "pendingHours")?.value),
      metric: metricValue(payload, "pendingHours"),
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiNoteCard({
      title: "Pending Hours Detail",
      metric: metricValue(payload, "pendingHoursDetail"),
      emptyLabel: "No pending-hours detail recorded",
      latestWeekLabelText: latestLabel,
    }),
  ]);

  appendChildren(businessTrendKpis, [
    createKpiTrendCard({
      title: "Total Contracted Hours Trend",
      series: payload?.trendSeries,
      field: "totalHours",
      formatter: formatHours,
    }),
    createKpiTrendCard({
      title: "Hours Won Trend",
      series: payload?.trendSeries,
      field: "hoursWon",
      formatter: formatHours,
    }),
    createKpiTrendCard({
      title: "Hours Lost Trend",
      series: payload?.trendSeries,
      field: "hoursLost",
      formatter: formatHours,
    }),
    createKpiTrendCard({
      title: "Pending Hours Trend",
      series: payload?.trendSeries,
      field: "pendingHours",
      formatter: formatHours,
    }),
  ]);
}

function renderEnquiries(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  appendChildren(enquiriesKpis, [
    createKpiMetricCard({
      title: "Active Enquiries",
      value: formatNumber(metricValue(payload, "activeEnquiries")?.value),
      metric: metricValue(payload, "activeEnquiries"),
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiMetricCard({
      title: "Enquiries Total /wk",
      value: formatNumber(metricValue(payload, "enquiriesTotal")?.value),
      metric: metricValue(payload, "enquiriesTotal"),
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiMetricCard({
      title: "Solicitor Enquiries",
      value: formatNumber(metricValue(payload, "enquiriesSolicitor")?.value),
      metric: metricValue(payload, "enquiriesSolicitor"),
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
    createKpiMetricCard({
      title: "Consumer Enquiries",
      value: formatNumber(metricValue(payload, "enquiriesConsumer")?.value),
      metric: metricValue(payload, "enquiriesConsumer"),
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
  ]);

  appendChildren(enquiriesTrendKpis, [
    createKpiTrendCard({
      title: "Active Enquiries Trend",
      series: payload?.trendSeries,
      field: "activeEnquiries",
      formatter: formatNumber,
    }),
    createKpiTrendCard({
      title: "Total Enquiries Trend",
      series: payload?.trendSeries,
      field: "enquiriesTotal",
      formatter: formatNumber,
    }),
    createKpiTrendCard({
      title: "Solicitor Enquiries Trend",
      series: payload?.trendSeries,
      field: "enquiriesSolicitor",
      formatter: formatNumber,
    }),
    createKpiTrendCard({
      title: "Consumer Enquiries Trend",
      series: payload?.trendSeries,
      field: "enquiriesConsumer",
      formatter: formatNumber,
    }),
  ]);
}

function renderRecruitment(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  const onboarding = payload?.onboarding || {};
  const firstNames = Array.isArray(onboarding.firstNames) ? onboarding.firstNames : [];
  appendChildren(recruitmentKpis, [
    createKpiMetricCard({
      title: "1st Round Interviews",
      value: formatNumber(metricValue(payload, "firstRoundInterviews")?.value),
      metric: metricValue(payload, "firstRoundInterviews"),
      latestWeekLabelText: latestLabel,
    }),
    createKpiMetricCard({
      title: "Currently Onboarding",
      value: formatNumber(onboarding.count),
      detail: firstNames.length ? firstNames.join(", ") : "No accepted candidates listed",
      metric: { sourceLabel: "From recruitment section: status Accepted", stale: false },
      tone: "recruitment",
      latestWeekLabelText: latestLabel,
    }),
  ]);
}

function renderTraining(payload) {
  appendChildren(trainingKpis, [
    createKpiMetricCard({
      title: "Training Completion",
      value: formatPercent(metricValue(payload, "trainingCompletion")?.value),
      metric: metricValue(payload, "trainingCompletion"),
      tone: "training",
      latestWeekLabelText: payload?.latestWeekLabel || "",
    }),
  ]);
}

function renderCqc(payload) {
  appendChildren(cqcKpis, [
    createKpiNoteCard({
      title: "CQC Readiness",
      metric: metricValue(payload, "cqcReadiness"),
      emptyLabel: "No readiness note recorded",
      latestWeekLabelText: payload?.latestWeekLabel || "",
    }),
  ]);
}

function renderKpis(payload) {
  renderSummary(payload);
  renderDelivery(payload);
  renderUtilisation(payload);
  renderBusiness(payload);
  renderEnquiries(payload);
  renderRecruitment(payload);
  renderTraining(payload);
  renderCqc(payload);
}

async function loadKpis() {
  if (loadingKpis) {
    return;
  }
  setLoading(true);
  setStatus("Loading KPI snapshot...");
  try {
    const payload = await directoryApi.getKpis();
    renderKpis(payload);
    setStatus(payload.latestWeekLabel ? `Showing latest week: ${payload.latestWeekLabel}` : "KPI snapshot loaded.");
  } catch (error) {
    console.error("[kpis] Load failed", error);
    setStatus(error?.message || "Could not load KPI snapshot.", true);
  } finally {
    setLoading(false);
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
    const role = cleanText(profile?.role).toLowerCase();
    if (!canAccessPage(role, "kpis")) {
      window.location.href = "./unauthorized.html?page=kpis";
      return;
    }

    renderTopNavigation({ role });
    await loadKpis();
  } catch (error) {
    console.error("[kpis] Init failed", error);
    setStatus(error?.message || "Could not initialize KPIs.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

refreshKpisBtn?.addEventListener("click", () => {
  void loadKpis();
});

generateKpiPdfBtn?.addEventListener("click", () => {
  closeTrendModal();
  window.print();
});

showKpiDetailsInput?.addEventListener("change", () => {
  applyViewPrefs();
});

showKpiTrendsInput?.addEventListener("change", () => {
  applyViewPrefs();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeTrendModal();
  }
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

applyViewPrefs(loadViewPrefs());
void init();
