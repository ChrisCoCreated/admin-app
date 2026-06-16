import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

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
const marketingKpis = document.getElementById("marketingKpis");
const marketingTrendKpis = document.getElementById("marketingTrendKpis");
const recruitmentKpis = document.getElementById("recruitmentKpis");
const cqcReadinessKpis = document.getElementById("cqcReadinessKpis");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let loadingKpis = false;
let savingKpiField = "";
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
  onClick = null,
}) {
  const card = document.createElement("article");
  card.className = `kpi-card kpi-card-${tone}${wide ? " kpi-card-wide" : ""}`;
  if (trendCompanion) {
    card.classList.add("kpi-trend-companion-tile");
  }
  if (typeof onClick === "function") {
    card.classList.add("kpi-card-clickable");
    card.setAttribute("role", "button");
    card.setAttribute("tabindex", "0");
    card.setAttribute("aria-label", `Open detail for ${title}`);
    card.addEventListener("click", onClick);
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        onClick(event);
      }
    });
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

export function createKpiNoteCard({ title, metric, emptyLabel = "No detail recorded", latestWeekLabelText = "", wide = true }) {
  const value = cleanText(metric?.value);
  const card = createKpiMetricCard({
    title,
    value: value || emptyLabel,
    metric,
    latestWeekLabelText,
    wide,
  });
  card.classList.add("kpi-note-card", "kpi-detail-tile");
  return card;
}

function editableInitialValue(metric) {
  return cleanText(metric?.value);
}

function createEditableKpiCard({
  title,
  fieldKey,
  metric,
  displayValue,
  editValue = null,
  emptyLabel = "No detail recorded",
  latestWeekLabelText = "",
  tone = "default",
  multiline = true,
}) {
  const card = document.createElement("article");
  card.className = `kpi-card kpi-card-${tone} kpi-editable-card`;
  const inputId = `kpi-edit-${fieldKey}`;
  const source = sourceLabel(metric, latestWeekLabelText);
  const initialValue = cleanText(editValue ?? editableInitialValue(metric));
  const shownValue = cleanText(displayValue || initialValue || emptyLabel);
  const inputMarkup = multiline
    ? `<textarea id="${escapeHtml(inputId)}" class="kpi-edit-input" rows="5">${escapeHtml(initialValue)}</textarea>`
    : `<input id="${escapeHtml(inputId)}" class="kpi-edit-input" type="text" value="${escapeHtml(initialValue)}">`;

  card.innerHTML = `
    <div class="kpi-card-topline">
      <h3>${escapeHtml(title)}</h3>
      <div class="kpi-edit-actions">
        <button class="secondary kpi-edit-btn" type="button">Edit</button>
        <button class="primary kpi-save-btn" type="button" hidden>Save</button>
      </div>
    </div>
    <div class="kpi-card-value kpi-edit-display">${escapeHtml(shownValue)}</div>
    <div class="kpi-edit-field" hidden>${inputMarkup}</div>
    <p class="kpi-card-source">${escapeHtml(source)}</p>
  `;

  const displayNode = card.querySelector(".kpi-edit-display");
  const fieldWrap = card.querySelector(".kpi-edit-field");
  const input = card.querySelector(".kpi-edit-input");
  const editBtn = card.querySelector(".kpi-edit-btn");
  const saveBtn = card.querySelector(".kpi-save-btn");

  function setEditing(isEditing) {
    displayNode.hidden = isEditing;
    fieldWrap.hidden = !isEditing;
    editBtn.hidden = isEditing;
    saveBtn.hidden = !isEditing;
    card.classList.toggle("is-editing", isEditing);
    if (isEditing) {
      input.focus();
      input.select?.();
    }
  }

  async function save() {
    if (savingKpiField) {
      return;
    }
    savingKpiField = fieldKey;
    saveBtn.disabled = true;
    saveBtn.textContent = "Saving...";
    setStatus(`Saving ${title}...`);
    try {
      await directoryApi.updateKpis({ fields: { [fieldKey]: input.value } });
      setStatus(`${title} saved to the latest KPI row.`);
      await loadKpis();
    } catch (error) {
      console.error("[kpis] Save failed", error);
      setStatus(error?.message || `Could not save ${title}.`, true);
      saveBtn.disabled = false;
      saveBtn.textContent = "Save";
    } finally {
      savingKpiField = "";
    }
  }

  card.addEventListener("click", (event) => {
    if (event.target instanceof Element && event.target.closest("button, textarea, input")) {
      return;
    }
    setEditing(true);
  });
  editBtn.addEventListener("click", () => setEditing(true));
  saveBtn.addEventListener("click", save);
  input.addEventListener("keydown", (event) => {
    if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
      event.preventDefault();
      save();
    }
  });

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

export function createKpiTrendCard({ title, series, field, formatter = formatNumber, onClick = null, compact = false }) {
  const points = getTrendValues(series, field);
  const latest = points.length ? points[points.length - 1] : null;
  const summary = summarizeTrend(points);
  const card = document.createElement("button");
  card.type = "button";
  card.className = `kpi-trend-card${compact ? " kpi-trend-card-compact" : ""}`;
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
    if (typeof onClick === "function") {
      onClick({ title, points, formatter, summary });
      return;
    }
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

function buildActiveEnquiriesList(activeEnquiries) {
  if (activeEnquiries?.unavailable) {
    return `<p class="muted">${escapeHtml(activeEnquiries.warning || "Could not load active enquiries.")}</p>`;
  }
  const items = Array.isArray(activeEnquiries?.items) ? activeEnquiries.items : [];
  const initialItems = Array.isArray(activeEnquiries?.initialItems) ? activeEnquiries.initialItems : [];
  if (!items.length) {
    return `
      <p class="muted">No active enquiries found in Enquiries Log.</p>
      ${
        initialItems.length
          ? `<section class="kpi-modal-detail-section"><h3>Initial Enquiries</h3>${buildActiveEnquiryRows(initialItems)}</section>`
          : ""
      }
    `;
  }
  return `
    <section class="kpi-modal-detail-section">
      <h3>Active Enquiries</h3>
      ${buildActiveEnquiryRows(items)}
    </section>
    <section class="kpi-modal-detail-section">
      <h3>Initial Enquiries</h3>
      ${initialItems.length ? buildActiveEnquiryRows(initialItems) : '<p class="muted">No initial enquiries found.</p>'}
    </section>
  `;
}

function buildActiveEnquiryRows(items) {
  return `
    <div class="kpi-modal-detail-list">
      ${items.map((item) => buildEnquiryDetailRow(item, "Active")).join("")}
    </div>
  `;
}

function buildEnquiryDetailRow(item, fallbackStatus = "No status") {
  const tagName = item?.webUrl ? "a" : "div";
  const linkAttrs = item?.webUrl
    ? ` href="${escapeHtml(item.webUrl)}" target="_blank" rel="noopener noreferrer"`
    : "";
  return `
    <${tagName} class="kpi-modal-detail-row"${linkAttrs}>
      <div>
        <strong>${escapeHtml(item.title || "Untitled enquiry")}</strong>
        <span>${escapeHtml(item.status || fallbackStatus)}</span>
      </div>
      <small>${escapeHtml(item.modifiedLabel ? `Modified ${item.modifiedLabel}` : "")}</small>
    </${tagName}>
  `;
}

function openActiveEnquiriesModal(activeEnquiries = {}) {
  const modal = ensureTrendModal();
  const titleNode = modal.querySelector("#kpiTrendModalTitle");
  const statsNode = modal.querySelector("#kpiTrendModalStats");
  const chartNode = modal.querySelector("#kpiTrendModalChart");
  const rowsNode = modal.querySelector("#kpiTrendModalRows");

  if (titleNode) {
    titleNode.textContent = "Current Active Enquiries";
  }
  if (statsNode) {
    statsNode.innerHTML = `
      <div><span>Current Active</span><strong>${escapeHtml(formatNumber(activeEnquiries?.count ?? 0))}</strong><small>Status 1-6</small></div>
      <div><span>Initial Enquiries</span><strong>${escapeHtml(formatNumber(activeEnquiries?.initialCount ?? 0))}</strong><small>Status 7</small></div>
    `;
  }
  if (chartNode) {
    chartNode.innerHTML = "";
  }
  if (rowsNode) {
    rowsNode.innerHTML = `
      <section class="kpi-modal-detail-section">
        ${buildActiveEnquiriesList(activeEnquiries)}
      </section>
    `;
  }

  modal.hidden = false;
  document.body.classList.add("kpi-modal-open");
  modal.querySelector(".kpi-modal-close")?.focus();
}

function buildOutcomeDetailRows(items) {
  if (!Array.isArray(items) || !items.length) {
    return '<p class="muted">No enquiries in this group for the selected period.</p>';
  }
  return `
    <div class="kpi-modal-detail-list">
      ${items.map((item) => buildEnquiryDetailRow(item, "No status")).join("")}
    </div>
  `;
}

function openAssessmentOutcomeModal(assessmentOutcome) {
  const modal = ensureTrendModal();
  const titleNode = modal.querySelector("#kpiTrendModalTitle");
  const statsNode = modal.querySelector("#kpiTrendModalStats");
  const chartNode = modal.querySelector("#kpiTrendModalChart");
  const rowsNode = modal.querySelector("#kpiTrendModalRows");
  const detail = assessmentOutcome?.detail || {};
  const won = Number(assessmentOutcome?.won || 0);
  const lost = Number(assessmentOutcome?.lost || 0);
  const onHold = Number(assessmentOutcome?.onHold || 0);

  if (titleNode) {
    titleNode.textContent = "Won / Lost / On Hold Detail";
  }
  if (statsNode) {
    statsNode.innerHTML = `
      <div><span>Won</span><strong>${escapeHtml(formatNumber(won))}</strong></div>
      <div><span>Lost</span><strong>${escapeHtml(formatNumber(lost))}</strong></div>
      <div><span>On hold</span><strong>${escapeHtml(formatNumber(onHold))}</strong></div>
      <div><span>Period</span><strong>${escapeHtml(`${assessmentOutcome?.months || 3} months`)}</strong><small>${escapeHtml(
        assessmentOutcome?.startDateLabel ? `Since ${assessmentOutcome.startDateLabel}` : ""
      )}</small></div>
    `;
  }
  if (chartNode) {
    chartNode.innerHTML = "";
  }
  if (rowsNode) {
    rowsNode.innerHTML = `
      <section class="kpi-modal-detail-section">
        <h3>Won</h3>
        ${buildOutcomeDetailRows(detail.won)}
      </section>
      <section class="kpi-modal-detail-section">
        <h3>Lost</h3>
        ${buildOutcomeDetailRows(detail.lost)}
      </section>
      <section class="kpi-modal-detail-section">
        <h3>On hold</h3>
        ${buildOutcomeDetailRows(detail.onHold)}
      </section>
    `;
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
      wide: false,
    }),
    createKpiNoteCard({
      title: "Explanatory Notes",
      metric: metricValue(payload, "explanatoryNotes"),
      emptyLabel: "No explanatory notes recorded",
      latestWeekLabelText: latestLabel,
      wide: false,
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
    createEditableKpiCard({
      title: "Utilisation Notes",
      fieldKey: "utilisationNotes",
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
    createKpiNoteCard({
      title: "Pending Hours Detail",
      metric: metricValue(payload, "pendingHoursDetail"),
      emptyLabel: "No pending-hours detail recorded",
      latestWeekLabelText: latestLabel,
    }),
  ]);
}

function renderEnquiries(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  const assessmentOutcome = payload?.enquiryAssessmentOutcome || {};
  const assessedOutcomes = Number(assessmentOutcome.assessedOutcomes || 0);
  const wonCount = Number(assessmentOutcome.won || 0);
  const lostCount = Number(assessmentOutcome.lost || 0);
  const onHoldCount = Number(assessmentOutcome.onHold || 0);
  const wonOnHoldRatio = onHoldCount ? `${formatNumber(wonCount / onHoldCount, 1)}:1 won/on hold` : "No on-hold enquiries";
  const outcomeUnavailable = Boolean(assessmentOutcome.unavailable);
  const outcomeSource = assessmentOutcome.startDateLabel
    ? `From Enquiries Log since ${assessmentOutcome.startDateLabel}`
    : "From Enquiries Log";
  appendChildren(enquiriesKpis, [
    createKpiMetricCard({
      title: "Won vs Lost After Assessment",
      value: outcomeUnavailable
        ? "Unavailable"
        : `${formatNumber(wonCount)} won / ${formatNumber(lostCount)} lost`,
      detail: outcomeUnavailable
        ? assessmentOutcome.warning || "Could not load Enquiries Log."
        : assessedOutcomes
        ? `${formatPercent(assessmentOutcome.winPercent)} won from ${formatNumber(assessedOutcomes)} won/lost outcomes in the past 3 months`
        : "No won/lost outcomes modified in the past 3 months",
      metric: { sourceLabel: outcomeUnavailable ? "Enquiries Log unavailable" : outcomeSource, stale: false },
      tone: "positive",
      onClick: outcomeUnavailable ? null : () => openAssessmentOutcomeModal(assessmentOutcome),
    }),
    createKpiMetricCard({
      title: "On Hold After Assessment",
      value: outcomeUnavailable ? "Unavailable" : `${formatNumber(onHoldCount)} on hold`,
      detail: outcomeUnavailable ? assessmentOutcome.warning || "Could not load Enquiries Log." : wonOnHoldRatio,
      metric: { sourceLabel: outcomeUnavailable ? "Enquiries Log unavailable" : outcomeSource, stale: false },
      tone: "positive",
      onClick: outcomeUnavailable ? null : () => openAssessmentOutcomeModal(assessmentOutcome),
    }),
    createKpiMetricCard({
      title: "Active Enquiries",
      value: `${formatNumber(payload?.activeEnquiries?.count ?? 0)} active / ${formatNumber(
        payload?.activeEnquiries?.initialCount ?? 0
      )} initial`,
      detail: "Active statuses 1-6; Initial Enquiry status 7",
      metric: metricValue(payload, "activeEnquiries"),
      latestWeekLabelText: latestLabel,
      onClick: () => openActiveEnquiriesModal(payload?.activeEnquiries),
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
    createKpiMetricCard({
      title: "Enquiry Conversion",
      value: formatPercent(metricValue(payload, "enquiryConversion")?.value),
      metric: metricValue(payload, "enquiryConversion"),
      latestWeekLabelText: latestLabel,
      trendCompanion: true,
    }),
  ]);

  appendChildren(enquiriesTrendKpis, [
    createKpiTrendCard({
      title: "Active Enquiries by Week",
      series: payload?.trendSeries,
      field: "activeEnquiries",
      formatter: formatNumber,
    }),
    createKpiTrendCard({
      title: "New Enquiries by Week",
      series: payload?.trendSeries,
      field: "enquiriesTotal",
      formatter: formatNumber,
    }),
    createKpiTrendCard({
      title: "Solicitor Enquiries Trend",
      series: payload?.trendSeries,
      field: "enquiriesSolicitor",
      formatter: formatNumber,
      compact: true,
    }),
    createKpiTrendCard({
      title: "Consumer Enquiries Trend",
      series: payload?.trendSeries,
      field: "enquiriesConsumer",
      formatter: formatNumber,
      compact: true,
    }),
    createKpiTrendCard({
      title: "Enquiry Conversion Trend",
      series: payload?.trendSeries,
      field: "enquiryConversion",
      formatter: formatPercent,
    }),
  ]);
}

function renderMarketing(payload) {
  const latestLabel = payload?.latestWeekLabel || "";
  const marketingMetrics = [
    ["Instagram Followers", "instagramFollowers"],
    ["Facebook Followers", "facebookFollowers"],
    ["Newsletter Subscribers", "newsletterSubscribers"],
    ["Web Visits", "webVisits"],
    ["Domain Authority - Thrive", "domainAuthorityThrive"],
    ["Domain Authority - PWC", "domainAuthorityPwc"],
  ];

  appendChildren(
    marketingKpis,
    marketingMetrics.map(([title, key]) =>
      createKpiMetricCard({
        title,
        value: formatNumber(metricValue(payload, key)?.value),
        metric: metricValue(payload, key),
        latestWeekLabelText: latestLabel,
        trendCompanion: true,
      })
    )
  );

  appendChildren(
    marketingTrendKpis,
    marketingMetrics.map(([title, key]) =>
      createKpiTrendCard({
        title: `${title} Trend`,
        series: payload?.trendSeries,
        field: key,
        formatter: formatNumber,
      })
    )
  );
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

function renderCqcReadiness(payload) {
  appendChildren(cqcReadinessKpis, [
    createEditableKpiCard({
      title: "Training Completion",
      fieldKey: "trainingCompletion",
      displayValue: formatPercent(metricValue(payload, "trainingCompletion")?.value),
      editValue: formatPercent(metricValue(payload, "trainingCompletion")?.value),
      metric: metricValue(payload, "trainingCompletion"),
      emptyLabel: "No training completion recorded",
      tone: "training",
      latestWeekLabelText: payload?.latestWeekLabel || "",
      multiline: false,
    }),
    createEditableKpiCard({
      title: "CQC Readiness",
      fieldKey: "cqcReadiness",
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
  renderMarketing(payload);
  renderRecruitment(payload);
  renderCqcReadiness(payload);
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
