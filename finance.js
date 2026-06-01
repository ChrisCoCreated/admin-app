import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const periodPresetSelect = document.getElementById("periodPresetSelect");
const dateRangeMessage = document.getElementById("dateRangeMessage");
const financialAnalysisLink = document.getElementById("financialAnalysisLink");
const payrollLink = document.getElementById("payrollLink");
const clientHoursLink = document.getElementById("clientHoursLink");
const carerHoursLink = document.getElementById("carerHoursLink");
const expensesCsvInput = document.getElementById("expensesCsvInput");
const invoiceDateInput = document.getElementById("invoiceDateInput");
const dueDaysInput = document.getElementById("dueDaysInput");
const startingInvoiceNumberInput = document.getElementById("startingInvoiceNumberInput");
const excludeInvoicedInput = document.getElementById("excludeInvoicedInput");
const onlyChargeableInput = document.getElementById("onlyChargeableInput");
const convertExpensesBtn = document.getElementById("convertExpensesBtn");
const downloadSalesInvoiceBtn = document.getElementById("downloadSalesInvoiceBtn");
const expensesImportStatus = document.getElementById("expensesImportStatus");
const expensesImportSummary = document.getElementById("expensesImportSummary");
const expensesImportErrors = document.getElementById("expensesImportErrors");
const expensesImportPreviewBody = document.getElementById("expensesImportPreviewBody");

const FINANCIAL_ANALYSIS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/timesheet_analysis_newscale_getPay.php";
const PAYROLL_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carerPayroll.php";
const CLIENT_HOURS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/clientsHoursRpt.php";
const CARER_HOURS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carersHoursRpt.php";
const SALES_INVOICE_HEADERS = [
  "ContactID",
  "*ContactName",
  "EmailAddress",
  "POAddressLine1",
  "POAddressLine2",
  "POAddressLine3",
  "POAddressLine4",
  "POCity",
  "PORegion",
  "POPostalCode",
  "POCountry",
  "*InvoiceNumber",
  "Reference",
  "*InvoiceDate",
  "*DueDate",
  "Total",
  "InventoryItemCode",
  "*Description",
  "*Quantity",
  "*UnitAmount",
  "Discount",
  "*AccountCode",
  "*TaxType",
  "TaxAmount",
  "TrackingName1",
  "TrackingOption1",
  "TrackingName2",
  "TrackingOption2",
  "Currency",
  "BrandingTheme",
];

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);
let latestSalesInvoiceCsv = "";
let latestSalesInvoiceRows = [];

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setExpensesStatus(message, isError = false) {
  if (!expensesImportStatus) {
    return;
  }
  expensesImportStatus.textContent = message;
  expensesImportStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "finance").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function formatDateParam(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function formatReadableDate(date) {
  return date.toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function formatDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatSlashDate(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${day}/${month}/${date.getFullYear()}`;
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

function parseBoolean(value) {
  return String(value || "").trim().toLowerCase() === "true";
}

function parseCurrencyAmount(value) {
  const normalized = String(value || "")
    .replace(/[£,\s]/g, "")
    .trim();
  const amount = Number(normalized);
  return Number.isFinite(amount) ? amount : null;
}

function parseSourceDate(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }
  const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]) - 1;
    const year = Number(slashMatch[3]);
    const parsed = new Date(year, month, day);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function cleanCell(value) {
  return String(value || "").replace(/\r/g, "").trim();
}

function parseCsvText(text) {
  const raw = String(text || "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    const next = raw[index + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }

    value += char;
  }

  if (value.length || row.length) {
    row.push(value);
    rows.push(row);
  }

  const nonEmptyRows = rows.filter((candidate) => candidate.some((cell) => cleanCell(cell)));
  if (!nonEmptyRows.length) {
    return { headers: [], rows: [], errors: ["CSV file is empty."] };
  }

  const headers = nonEmptyRows[0].map((header) => cleanCell(header));
  const records = [];
  const errors = [];

  for (let rowIndex = 1; rowIndex < nonEmptyRows.length; rowIndex += 1) {
    const source = nonEmptyRows[rowIndex];
    const record = {};
    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      record[headers[columnIndex]] = cleanCell(source[columnIndex] || "");
    }
    if (source.length !== headers.length) {
      errors.push(`Row ${rowIndex + 1} has ${source.length} values; expected ${headers.length}.`);
    }
    records.push(record);
  }

  return { headers, rows: records, errors };
}

function escapeCsvValue(value) {
  const text = String(value ?? "");
  if (!/[",\n]/.test(text)) {
    return text;
  }
  return `"${text.replaceAll('"', '""')}"`;
}

function downloadCsv(filename, csvText) {
  const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function buildInvoiceNumber(startingNumber, index) {
  const baseNumber = String(startingNumber ?? "").trim();
  const numeric = Number(baseNumber);
  if (!Number.isInteger(numeric) || numeric < 1) {
    return String(index + 1);
  }
  const width = baseNumber.length;
  return String(numeric + index).padStart(width, "0");
}

function buildDescription(row) {
  const parts = [cleanCell(row.Description)];
  const details = cleanCell(row.Details);
  const journeys = cleanCell(row["List of Journeys"]);
  const claimant = cleanCell(row["Claim From"]);

  if (details) {
    parts.push(details);
  }
  if (journeys) {
    parts.push(`Journeys: ${journeys}`);
  }
  if (claimant) {
    parts.push(`Claim from: ${claimant}`);
  }

  return parts.filter(Boolean).join("\n\n");
}

function renderExpensesPreview(rows) {
  if (!expensesImportPreviewBody) {
    return;
  }
  expensesImportPreviewBody.innerHTML = "";

  if (!rows.length) {
    expensesImportPreviewBody.innerHTML = '<tr><td colspan="7" class="muted">No converted rows yet.</td></tr>';
    return;
  }

  for (const row of rows.slice(0, 8)) {
    const tr = document.createElement("tr");
    for (const value of [
      row.ContactID,
      row["*ContactName"],
      row["*InvoiceNumber"],
      row["*InvoiceDate"],
      row["*Description"],
      row["*UnitAmount"],
      row["*AccountCode"],
    ]) {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    }
    expensesImportPreviewBody.appendChild(tr);
  }
}

function resetExpensesOutput() {
  latestSalesInvoiceCsv = "";
  latestSalesInvoiceRows = [];
  downloadSalesInvoiceBtn?.setAttribute("disabled", "disabled");
  renderExpensesPreview([]);
}

function convertExpensesRows(sourceRows) {
  const invoiceDate = parseSourceDate(invoiceDateInput?.value) || new Date();
  const dueInDays = Math.max(0, Number(dueDaysInput?.value || 0) || 0);
  const startingInvoiceNumber = cleanCell(startingInvoiceNumberInput?.value);
  const excludeInvoiced = excludeInvoicedInput?.checked !== false;
  const onlyChargeable = onlyChargeableInput?.checked !== false;
  const groupedRows = new Map();
  const outputRows = [];
  const issues = [];
  let skippedAlreadyInvoiced = 0;
  let skippedNonChargeable = 0;

  sourceRows.forEach((row, index) => {
    const isInvoiced = parseBoolean(row.Invoiced);
    const isReimbursable = parseBoolean(row["Is Reimbursable"]);
    const spentOnClientsBehalf = parseBoolean(row["Spent on Clients Behalf"]);

    if (excludeInvoiced && isInvoiced) {
      skippedAlreadyInvoiced += 1;
      return;
    }

    if (onlyChargeable && !(isReimbursable || spentOnClientsBehalf)) {
      skippedNonChargeable += 1;
      return;
    }

    const amount = parseCurrencyAmount(row.Amount);
    const accountCode = cleanCell(row["Account Code - Invoice"]);
    const contactName = cleanCell(row.Client);
    const contactId = cleanCell(row["Client:XeroID"]);

    if (!contactName) {
      issues.push(`Row ${index + 2} skipped: missing Client.`);
      return;
    }
    if (!contactId) {
      issues.push(`Row ${index + 2} skipped: missing Client:XeroID.`);
      return;
    }
    if (amount === null) {
      issues.push(`Row ${index + 2} skipped: invalid Amount "${row.Amount || ""}".`);
      return;
    }
    if (!accountCode) {
      issues.push(`Row ${index + 2} skipped: missing Account Code - Invoice.`);
      return;
    }

    const description = buildDescription(row);
    const referenceParts = [cleanCell(row["Claim From"]), cleanCell(row["Client: Job Type"])]
      .filter(Boolean)
      .join(" - ");
    const existingGroup = groupedRows.get(contactId);

    if (!existingGroup) {
      groupedRows.set(contactId, {
        contactId,
        contactName,
        accountCode,
        total: amount,
        descriptions: [description],
        references: referenceParts ? [referenceParts] : [],
      });
      return;
    }

    if (existingGroup.accountCode !== accountCode) {
      issues.push(
        `Row ${index + 2} has a different account code for ContactID ${contactId}. Using ${existingGroup.accountCode}.`
      );
    }
    if (existingGroup.contactName !== contactName) {
      issues.push(`Row ${index + 2} has a different client name for ContactID ${contactId}. Using ${existingGroup.contactName}.`);
    }

    existingGroup.total += amount;
    existingGroup.descriptions.push(description);
    if (referenceParts) {
      existingGroup.references.push(referenceParts);
    }
  });

  Array.from(groupedRows.values())
    .sort((left, right) => left.contactName.localeCompare(right.contactName, undefined, { sensitivity: "base" }))
    .forEach((group, index) => {
      outputRows.push({
        ContactID: group.contactId,
        "*ContactName": group.contactName,
        EmailAddress: "",
        POAddressLine1: "",
        POAddressLine2: "",
        POAddressLine3: "",
        POAddressLine4: "",
        POCity: "",
        PORegion: "",
        POPostalCode: "",
        POCountry: "",
        "*InvoiceNumber": buildInvoiceNumber(startingInvoiceNumber, index),
        Reference: Array.from(new Set(group.references)).join(" | "),
        "*InvoiceDate": formatSlashDate(invoiceDate),
        "*DueDate": formatSlashDate(addDays(invoiceDate, dueInDays)),
        Total: group.total.toFixed(2),
        InventoryItemCode: "",
        "*Description": group.descriptions.filter(Boolean).join("\n\n"),
        "*Quantity": "1",
        "*UnitAmount": group.total.toFixed(2),
        Discount: "",
        "*AccountCode": group.accountCode,
        "*TaxType": "",
        TaxAmount: "",
        TrackingName1: "",
        TrackingOption1: "",
        TrackingName2: "",
        TrackingOption2: "",
        Currency: "",
        BrandingTheme: "",
      });
    });

  return {
    rows: outputRows,
    issues,
    skippedAlreadyInvoiced,
    skippedNonChargeable,
    invoiceDate,
  };
}

async function handleExpensesConvert() {
  try {
    const file = expensesCsvInput?.files?.[0];
    if (!file) {
      setExpensesStatus("Choose an expenses CSV first.", true);
      return;
    }

    const text = await file.text();
    const parsed = parseCsvText(text);
    if (parsed.errors.length) {
      setExpensesStatus("The CSV was read with warnings. Review the summary below.");
    } else {
      setExpensesStatus("Expenses CSV converted.");
    }

    const converted = convertExpensesRows(parsed.rows);
    latestSalesInvoiceRows = converted.rows;

    const csvLines = [SALES_INVOICE_HEADERS.map(escapeCsvValue).join(",")];
    for (const row of converted.rows) {
      csvLines.push(SALES_INVOICE_HEADERS.map((header) => escapeCsvValue(row[header] || "")).join(","));
    }
    latestSalesInvoiceCsv = csvLines.join("\n");

    if (expensesImportSummary) {
      expensesImportSummary.textContent =
        `Imported ${parsed.rows.length} source row(s). Prepared ${converted.rows.length} sales invoice(s). ` +
        `Skipped ${converted.skippedAlreadyInvoiced} already invoiced row(s) and ${converted.skippedNonChargeable} non-chargeable row(s).`;
    }

    const allIssues = [...parsed.errors, ...converted.issues];
    if (expensesImportErrors) {
      if (allIssues.length) {
        expensesImportErrors.hidden = false;
        expensesImportErrors.innerHTML = allIssues.map((issue) => `<div>${issue}</div>`).join("");
      } else {
        expensesImportErrors.hidden = true;
        expensesImportErrors.innerHTML = "";
      }
    }

    renderExpensesPreview(converted.rows);
    if (converted.rows.length) {
      downloadSalesInvoiceBtn?.removeAttribute("disabled");
    } else {
      downloadSalesInvoiceBtn?.setAttribute("disabled", "disabled");
    }
  } catch (error) {
    console.error("Expense CSV conversion failed", error);
    resetExpensesOutput();
    setExpensesStatus(error?.message || "Could not convert the expenses CSV.", true);
  }
}

function handleSalesInvoiceDownload() {
  if (!latestSalesInvoiceCsv || !latestSalesInvoiceRows.length) {
    setExpensesStatus("Convert a file before downloading.", true);
    return;
  }
  const today = new Date();
  const stamp = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, "0")}${String(today.getDate()).padStart(2, "0")}`;
  downloadCsv(`sales-invoices-import-${stamp}.csv`, latestSalesInvoiceCsv);
  setExpensesStatus("Sales invoice CSV downloaded. Next, upload it in Xero.");
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
    return {
      start: new Date(year, month, 1),
      end: new Date(year, month + 1, 0),
    };
  }

  const thisWeekMonday = getMondayOfWeek(today);
  if (preset === "last_week") {
    const start = addDays(thisWeekMonday, -7);
    return { start, end: addDays(start, 6) };
  }
  if (preset === "next_week") {
    const start = addDays(thisWeekMonday, 7);
    return { start, end: addDays(start, 6) };
  }

  const start = thisWeekMonday;
  return { start, end: addDays(start, 6) };
}

function buildFinancialAnalysisUrl(start, end) {
  const url = new URL(FINANCIAL_ANALYSIS_BASE_URL);
  url.searchParams.set("start", formatDateParam(start));
  url.searchParams.set("finish", formatDateParam(end));
  return url.toString();
}

function buildPayrollUrl(start, end) {
  const url = new URL(PAYROLL_BASE_URL);
  url.searchParams.set("datePickSt", formatDateParam(start));
  url.searchParams.set("datePickFn", formatDateParam(end));
  return url.toString();
}

function buildClientHoursUrl(start, end) {
  const url = new URL(CLIENT_HOURS_BASE_URL);
  url.searchParams.set("selectJob", "all");
  url.searchParams.set("startDate", formatDateParam(start));
  url.searchParams.set("endDate", formatDateParam(end));
  return url.toString();
}

function buildCarerHoursUrl(start, end) {
  const url = new URL(CARER_HOURS_BASE_URL);
  url.searchParams.set("jobtype", "All");
  url.searchParams.set("start", formatDateParam(start));
  url.searchParams.set("finish", formatDateParam(end));
  return url.toString();
}

function refreshLinks() {
  const preset = periodPresetSelect?.value || "last_month";
  const { start, end } = getDateRangeForPreset(preset);

  if (dateRangeMessage) {
    dateRangeMessage.textContent = `Selected period: ${formatReadableDate(start)} to ${formatReadableDate(end)}.`;
  }
  if (financialAnalysisLink) {
    financialAnalysisLink.href = buildFinancialAnalysisUrl(start, end);
  }
  if (payrollLink) {
    payrollLink.href = buildPayrollUrl(start, end);
  }
  if (clientHoursLink) {
    clientHoursLink.href = buildClientHoursUrl(start, end);
  }
  if (carerHoursLink) {
    carerHoursLink.href = buildCarerHoursUrl(start, end);
  }
}

async function init() {
  try {
    setStatus("Checking access...");
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();

    if (!canAccessPage(role, "finance")) {
      redirectToUnauthorized("finance");
      return;
    }

    renderTopNavigation({ role });
    document.body.classList.remove("auth-pending");
    setStatus(`Signed in as ${profile?.email || "unknown"} (${role || "unknown role"}).`);
    if (invoiceDateInput && !invoiceDateInput.value) {
      invoiceDateInput.value = formatDateInputValue(new Date());
    }
    resetExpensesOutput();
    refreshLinks();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("finance");
      return;
    }
    console.error("Finance page failed to initialise", error);
    setStatus(error?.message || "Unable to load the finance page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
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

periodPresetSelect?.addEventListener("change", refreshLinks);
convertExpensesBtn?.addEventListener("click", handleExpensesConvert);
downloadSalesInvoiceBtn?.addEventListener("click", handleSalesInvoiceDownload);
expensesCsvInput?.addEventListener("change", () => {
  resetExpensesOutput();
  setExpensesStatus("File selected. Convert it to prepare the Xero import CSV.");
  if (expensesImportSummary) {
    expensesImportSummary.textContent = "";
  }
  if (expensesImportErrors) {
    expensesImportErrors.hidden = true;
    expensesImportErrors.innerHTML = "";
  }
});

void init();
