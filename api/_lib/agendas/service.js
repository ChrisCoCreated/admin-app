const {
  createAgendaMembers,
  createAgendaItemRow,
  createAgendaRow,
  ensureUuid,
  listAgendaItemsByAgendaIds,
  listAgendaMembersByAgendaIds,
  listAgendasByIds,
  listMembershipsByUser,
  replaceAgendaMembers,
  updateAgendaItemRow,
  updateAgendaRow,
} = require("./repository");
const { mapAgendaItemRow, mapAgendaRow, normalizeEmail, sanitizeAgendaInput, sanitizeAgendaItemInput } = require("./model");

function buildMemberRows(agendaId, ownerEmail, participantEmails = [], participantNames = []) {
  const rows = [
    {
      agenda_id: agendaId,
      user_email: normalizeEmail(ownerEmail),
      display_name: "",
      is_owner: true,
    },
  ];

  participantEmails.forEach((email, index) => {
    rows.push({
      agenda_id: agendaId,
      user_email: normalizeEmail(email),
      display_name: participantNames[index] || "",
      is_owner: false,
    });
  });

  return rows;
}

async function loadVisibleAgendaContext(userEmail, options = {}) {
  const includeItems = options.includeItems !== false;
  const memberships = await listMembershipsByUser(userEmail);
  const agendaIds = Array.from(new Set(memberships.map((row) => String(row.agenda_id || "").trim()).filter(Boolean)));
  if (!agendaIds.length) {
    return { agendas: [] };
  }

  const fetches = [listAgendasByIds(agendaIds), listAgendaMembersByAgendaIds(agendaIds)];
  if (includeItems) {
    fetches.push(listAgendaItemsByAgendaIds(agendaIds));
  }
  const [agendaRows, memberRows, itemRows = []] = await Promise.all(fetches);

  const membersByAgenda = new Map();
  const itemsByAgenda = new Map();

  memberRows.forEach((row) => {
    const key = String(row.agenda_id || "").trim();
    if (!membersByAgenda.has(key)) {
      membersByAgenda.set(key, []);
    }
    membersByAgenda.get(key).push(row);
  });

  itemRows.forEach((row) => {
    const key = String(row.agenda_id || "").trim();
    if (!itemsByAgenda.has(key)) {
      itemsByAgenda.set(key, []);
    }
    itemsByAgenda.get(key).push(mapAgendaItemRow(row));
  });

  const agendas = agendaRows
    .map((row) => mapAgendaRow(row, membersByAgenda.get(String(row.id || "").trim()) || [], itemsByAgenda.get(String(row.id || "").trim()) || []))
    .filter((agenda) => agenda && agenda.id);

  return { agendas };
}

async function loadVisibleAgendaContextForIds(userEmail, agendaIds = [], options = {}) {
  const includeItems = options.includeItems !== false;
  const memberships = await listMembershipsByUser(userEmail);
  const visibleAgendaIds = new Set(memberships.map((row) => String(row.agenda_id || "").trim()).filter(Boolean));
  const requestedIds = (Array.isArray(agendaIds) ? agendaIds : [])
    .map((id) => String(id || "").trim())
    .filter((id) => visibleAgendaIds.has(id));

  if (!requestedIds.length) {
    return { agendas: [] };
  }

  const fetches = [listAgendasByIds(requestedIds), listAgendaMembersByAgendaIds(requestedIds)];
  if (includeItems) {
    fetches.push(listAgendaItemsByAgendaIds(requestedIds));
  }

  const [agendaRows, memberRows, itemRows = []] = await Promise.all(fetches);
  const membersByAgenda = new Map();
  const itemsByAgenda = new Map();

  memberRows.forEach((row) => {
    const key = String(row.agenda_id || "").trim();
    if (!membersByAgenda.has(key)) {
      membersByAgenda.set(key, []);
    }
    membersByAgenda.get(key).push(row);
  });

  itemRows.forEach((row) => {
    const key = String(row.agenda_id || "").trim();
    if (!itemsByAgenda.has(key)) {
      itemsByAgenda.set(key, []);
    }
    itemsByAgenda.get(key).push(mapAgendaItemRow(row));
  });

  const agendas = agendaRows
    .map((row) => mapAgendaRow(row, membersByAgenda.get(String(row.id || "").trim()) || [], itemsByAgenda.get(String(row.id || "").trim()) || []))
    .filter((agenda) => agenda && agenda.id);

  return { agendas };
}

function canViewAgenda(agenda, userEmail) {
  const normalizedUser = normalizeEmail(userEmail);
  const isMember = agenda.members.some((member) => normalizeEmail(member.userEmail) === normalizedUser);
  if (!isMember) {
    return false;
  }
  if (agenda.isPrivate) {
    return normalizeEmail(agenda.ownerEmail) === normalizedUser;
  }
  return true;
}

function canEditAgenda(agenda, userEmail) {
  return canViewAgenda(agenda, userEmail);
}

function canViewAgendaItem(item, agenda, userEmail) {
  const normalizedUser = normalizeEmail(userEmail);
  if (!canViewAgenda(agenda, normalizedUser)) {
    return false;
  }
  if (!item.isPrivate) {
    return true;
  }
  return normalizeEmail(item.ownerEmail) === normalizedUser || normalizeEmail(agenda.ownerEmail) === normalizedUser;
}

function canEditAgendaItem(item, agenda, userEmail) {
  return canViewAgendaItem(item, agenda, userEmail);
}

function sortAgendas(agendas) {
  return [...agendas].sort((a, b) => {
    const aUpdated = Date.parse(a.updatedAt || "") || 0;
    const bUpdated = Date.parse(b.updatedAt || "") || 0;
    if (aUpdated !== bUpdated) {
      return bUpdated - aUpdated;
    }
    return String(a.title || "").localeCompare(String(b.title || ""), undefined, { sensitivity: "base" });
  });
}

function sortItems(items) {
  return [...items].sort((a, b) => {
    const orderDelta = Number(a.sortOrder || 0) - Number(b.sortOrder || 0);
    if (orderDelta !== 0) {
      return orderDelta;
    }
    const aUpdated = Date.parse(a.updatedAt || "") || 0;
    const bUpdated = Date.parse(b.updatedAt || "") || 0;
    return bUpdated - aUpdated;
  });
}

async function listAgendasForUser(email) {
  const context = await loadVisibleAgendaContext(email);
  const visibleAgendas = sortAgendas(
    context.agendas
      .filter((agenda) => canViewAgenda(agenda, email))
      .map((agenda) => ({
        ...agenda,
        items: sortItems(agenda.items.filter((item) => canViewAgendaItem(item, agenda, email))),
      }))
  );

  return {
    agendas: visibleAgendas,
    meta: {
      totalAgendas: visibleAgendas.length,
      totalItems: visibleAgendas.reduce((sum, agenda) => sum + agenda.items.length, 0),
      currentUserEmail: normalizeEmail(email),
    },
  };
}

async function listAgendaSummariesForUser(email) {
  const context = await loadVisibleAgendaContext(email, { includeItems: false });
  const visibleAgendas = sortAgendas(
    context.agendas
      .filter((agenda) => canViewAgenda(agenda, email))
      .map((agenda) => ({
        ...agenda,
        items: [],
      }))
  );

  return {
    agendas: visibleAgendas,
    meta: {
      totalAgendas: visibleAgendas.length,
      currentUserEmail: normalizeEmail(email),
    },
  };
}

async function getAgendaDetailForUser(email, agendaId) {
  const context = await loadVisibleAgendaContextForIds(email, [agendaId], { includeItems: true });
  const agenda = context.agendas.find((entry) => entry.id === ensureUuid(agendaId, "Agenda"));
  if (!agenda || !canViewAgenda(agenda, email)) {
    const error = new Error("Agenda not found.");
    error.status = 404;
    error.code = "AGENDA_NOT_FOUND";
    throw error;
  }

  return {
    agenda: {
      ...agenda,
      items: sortItems(agenda.items.filter((item) => canViewAgendaItem(item, agenda, email))),
    },
  };
}

async function createAgendaForUser(email, body) {
  const payload = sanitizeAgendaInput(body, email, "create");
  const created = await createAgendaRow({
    title: payload.title,
    agenda_type: payload.agendaType,
    owner_email: payload.ownerEmail,
    is_private: payload.isPrivate,
  });
  const agendaId = ensureUuid(created?.id, "Agenda");
  await createAgendaMembers(buildMemberRows(agendaId, payload.ownerEmail, payload.participantEmails, payload.participantNames));
  return { ok: true, agendaId };
}

async function updateAgendaForUser(email, body) {
  const agendaId = ensureUuid(body?.agendaId, "Agenda");
  const current = await listAgendasForUser(email);
  const agenda = current.agendas.find((entry) => entry.id === agendaId);
  if (!agenda) {
    const error = new Error("Agenda not found.");
    error.status = 404;
    error.code = "AGENDA_NOT_FOUND";
    throw error;
  }
  if (!canEditAgenda(agenda, email) || normalizeEmail(agenda.ownerEmail) !== normalizeEmail(email)) {
    const error = new Error("Forbidden.");
    error.status = 403;
    error.code = "FORBIDDEN";
    throw error;
  }

  const patch = sanitizeAgendaInput({ ...agenda, ...body }, agenda.ownerEmail, "update");
  await updateAgendaRow(agendaId, {
    ...(patch.title ? { title: patch.title } : {}),
    agenda_type: patch.agendaType,
    is_private: patch.isPrivate,
    updated_at: new Date().toISOString(),
  });
  await replaceAgendaMembers(agendaId, buildMemberRows(agendaId, agenda.ownerEmail, patch.participantEmails, patch.participantNames).filter((member) => !member.is_owner));
  return { ok: true };
}

async function createAgendaItemForUser(email, body) {
  const agendaId = ensureUuid(body?.agendaId, "Agenda");
  const current = await listAgendasForUser(email);
  const agenda = current.agendas.find((entry) => entry.id === agendaId);
  if (!agenda) {
    const error = new Error("Agenda not found.");
    error.status = 404;
    error.code = "AGENDA_NOT_FOUND";
    throw error;
  }
  if (!canEditAgenda(agenda, email)) {
    const error = new Error("Forbidden.");
    error.status = 403;
    error.code = "FORBIDDEN";
    throw error;
  }

  const maxSortOrder = agenda.items.reduce((max, item) => Math.max(max, Number(item.sortOrder || 0)), 0);
  const payload = sanitizeAgendaItemInput(
    {
      ...body,
      agendaId,
      sortOrder: Object.prototype.hasOwnProperty.call(body || {}, "sortOrder") ? body.sortOrder : maxSortOrder + 100,
    },
    email,
    "create"
  );

  await createAgendaItemRow({
    agenda_id: agendaId,
    title: payload.title,
    detail_html: payload.detailHtml || "<p></p>",
    is_private: payload.isPrivate === true,
    is_urgent: payload.isUrgent === true,
    is_important: payload.isImportant === true,
    sort_order: payload.sortOrder,
    stage_tag: payload.stageTag || "",
    owner_email: payload.ownerEmail,
    task_map: Object.prototype.hasOwnProperty.call(payload, "taskMap") ? payload.taskMap : null,
  });

  await updateAgendaRow(agendaId, { updated_at: new Date().toISOString() });
  return { ok: true };
}

async function updateAgendaItemForUser(email, body) {
  const itemId = ensureUuid(body?.itemId, "Agenda item");
  const current = await listAgendasForUser(email);
  const agenda = current.agendas.find((entry) => entry.items.some((item) => item.id === itemId));
  if (!agenda) {
    const error = new Error("Agenda item not found.");
    error.status = 404;
    error.code = "AGENDA_ITEM_NOT_FOUND";
    throw error;
  }
  const item = agenda.items.find((entry) => entry.id === itemId);
  if (!canEditAgendaItem(item, agenda, email)) {
    const error = new Error("Forbidden.");
    error.status = 403;
    error.code = "FORBIDDEN";
    throw error;
  }

  const payload = sanitizeAgendaItemInput({ ...item, ...body }, item.ownerEmail, "update");
  await updateAgendaItemRow(itemId, {
    ...(payload.title ? { title: payload.title } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "detailHtml") ? { detail_html: payload.detailHtml } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "isPrivate") ? { is_private: payload.isPrivate === true } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "isUrgent") ? { is_urgent: payload.isUrgent === true } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "isImportant") ? { is_important: payload.isImportant === true } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "sortOrder") ? { sort_order: payload.sortOrder } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "stageTag") ? { stage_tag: payload.stageTag || "" } : {}),
    ...(Object.prototype.hasOwnProperty.call(payload, "taskMap") ? { task_map: payload.taskMap } : {}),
    updated_at: new Date().toISOString(),
  });
  await updateAgendaRow(agenda.id, { updated_at: new Date().toISOString() });
  return { ok: true };
}

function mapAgendaError(error) {
  return {
    status: Number(error?.status) || 500,
    payload: {
      error: {
        code: String(error?.code || "AGENDA_REQUEST_FAILED"),
        message: error?.message || "Agenda request failed.",
      },
    },
  };
}

module.exports = {
  createAgendaForUser,
  createAgendaItemForUser,
  getAgendaDetailForUser,
  listAgendasForUser,
  listAgendaSummariesForUser,
  mapAgendaError,
  updateAgendaForUser,
  updateAgendaItemForUser,
};
