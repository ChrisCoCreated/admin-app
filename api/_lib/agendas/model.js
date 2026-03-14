const { sanitizeLightHtml } = require("./html");

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function parseBoolean(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  return ["1", "true", "yes"].includes(String(value || "").trim().toLowerCase());
}

function parseNumber(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeAgendaType(value) {
  return String(value || "").trim().toLowerCase() === "one_to_one" ? "one_to_one" : "meeting";
}

function normalizeStageTag(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return new Set(["", "incubating", "ready", "decision", "blocked"]).has(normalized) ? normalized : "";
}

function parseStringArray(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => String(entry || "").trim()).filter(Boolean);
  }
  return String(value || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function dedupeParticipants(emails, names, ownerEmail) {
  const owner = normalizeEmail(ownerEmail);
  const emailList = Array.from(
    new Set(
      parseStringArray(emails)
        .map(normalizeEmail)
        .filter((entry) => entry && entry !== owner)
    )
  );
  const nameList = Array.from(new Set(parseStringArray(names)));
  return { emailList, nameList };
}

function sanitizeAgendaInput(input, ownerEmail, mode = "create") {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    const error = new Error("Agenda payload must be an object.");
    error.status = 400;
    error.code = "INVALID_AGENDA_PAYLOAD";
    throw error;
  }

  const title = String(input.title || "").trim();
  if (mode === "create" && !title) {
    const error = new Error("Agenda title is required.");
    error.status = 400;
    error.code = "AGENDA_TITLE_REQUIRED";
    throw error;
  }

  const { emailList, nameList } = dedupeParticipants(input.participantEmails, input.participantNames, ownerEmail);
  return {
    ...(title ? { title } : {}),
    ownerEmail: normalizeEmail(ownerEmail),
    participantEmails: emailList,
    participantNames: nameList,
    agendaType: normalizeAgendaType(input.agendaType),
    isPrivate: parseBoolean(input.isPrivate),
  };
}

function sanitizeAgendaItemInput(input, ownerEmail, mode = "create") {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    const error = new Error("Agenda item payload must be an object.");
    error.status = 400;
    error.code = "INVALID_AGENDA_ITEM_PAYLOAD";
    throw error;
  }

  const title = String(input.title || "").trim();
  if (mode === "create" && !title) {
    const error = new Error("Agenda item title is required.");
    error.status = 400;
    error.code = "AGENDA_ITEM_TITLE_REQUIRED";
    throw error;
  }

  const output = {
    ...(title ? { title } : {}),
    ownerEmail: normalizeEmail(ownerEmail),
    ...(Object.prototype.hasOwnProperty.call(input, "detailHtml")
      ? { detailHtml: sanitizeLightHtml(input.detailHtml) }
      : {}),
    ...(Object.prototype.hasOwnProperty.call(input, "isPrivate")
      ? { isPrivate: parseBoolean(input.isPrivate) }
      : {}),
    ...(Object.prototype.hasOwnProperty.call(input, "isUrgent")
      ? { isUrgent: parseBoolean(input.isUrgent) }
      : {}),
    ...(Object.prototype.hasOwnProperty.call(input, "isImportant")
      ? { isImportant: parseBoolean(input.isImportant) }
      : {}),
    ...(Object.prototype.hasOwnProperty.call(input, "sortOrder")
      ? { sortOrder: parseNumber(input.sortOrder, 0) }
      : {}),
    ...(Object.prototype.hasOwnProperty.call(input, "stageTag")
      ? { stageTag: normalizeStageTag(input.stageTag) }
      : {}),
  };

  if (mode === "create") {
    output.agendaId = String(input.agendaId || "").trim();
  }
  return output;
}

function mapAgendaRow(row, members = [], items = []) {
  const ownerEmail = normalizeEmail(row?.owner_email);
  const memberRows = Array.isArray(members) ? members : [];
  const participantRows = memberRows.filter((member) => normalizeEmail(member.user_email) !== ownerEmail);

  return {
    id: String(row?.id || "").trim(),
    title: String(row?.title || "").trim(),
    agendaType: normalizeAgendaType(row?.agenda_type),
    ownerEmail,
    isPrivate: row?.is_private === true,
    participantEmails: participantRows.map((member) => normalizeEmail(member.user_email)).filter(Boolean),
    participantNames: participantRows.map((member) => String(member.display_name || "").trim()).filter(Boolean),
    members: memberRows.map((member) => ({
      userEmail: normalizeEmail(member.user_email),
      displayName: String(member.display_name || "").trim(),
      isOwner: member.is_owner === true,
    })),
    items: Array.isArray(items) ? items : [],
    createdAt: row?.created_at || null,
    updatedAt: row?.updated_at || null,
  };
}

function mapAgendaItemRow(row) {
  return {
    id: String(row?.id || "").trim(),
    agendaId: String(row?.agenda_id || "").trim(),
    title: String(row?.title || "").trim(),
    detailHtml: sanitizeLightHtml(row?.detail_html),
    isPrivate: row?.is_private === true,
    isUrgent: row?.is_urgent === true,
    isImportant: row?.is_important === true,
    sortOrder: parseNumber(row?.sort_order, 0),
    stageTag: normalizeStageTag(row?.stage_tag),
    ownerEmail: normalizeEmail(row?.owner_email),
    createdAt: row?.created_at || null,
    updatedAt: row?.updated_at || null,
  };
}

module.exports = {
  mapAgendaItemRow,
  mapAgendaRow,
  normalizeEmail,
  sanitizeAgendaInput,
  sanitizeAgendaItemInput,
};
