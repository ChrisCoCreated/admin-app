const { supabaseRestFetch } = require("../supabase-rest");

function isUuid(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || ""));
}

function ensureUuid(value, label) {
  if (!isUuid(value)) {
    const error = new Error(`${label} is invalid.`);
    error.status = 400;
    error.code = "INVALID_UUID";
    throw error;
  }
  return String(value);
}

function buildInFilter(ids) {
  const safeIds = (Array.isArray(ids) ? ids : []).filter((id) => isUuid(id)).map(String);
  return safeIds.length ? `in.(${safeIds.join(",")})` : null;
}

async function listMembershipsByUser(userEmail) {
  const rows = await supabaseRestFetch("agenda_members", {
    query: {
      select: "agenda_id,user_email,display_name,is_owner,created_at",
      user_email: `eq.${String(userEmail || "").trim().toLowerCase()}`,
    },
  });
  return Array.isArray(rows) ? rows : [];
}

async function listAgendasByIds(ids) {
  const filter = buildInFilter(ids);
  if (!filter) {
    return [];
  }
  const rows = await supabaseRestFetch("agendas", {
    query: {
      select: "*",
      id: filter,
      order: "updated_at.desc,title.asc",
    },
  });
  return Array.isArray(rows) ? rows : [];
}

async function listAgendaMembersByAgendaIds(ids) {
  const filter = buildInFilter(ids);
  if (!filter) {
    return [];
  }
  const rows = await supabaseRestFetch("agenda_members", {
    query: {
      select: "*",
      agenda_id: filter,
      order: "created_at.asc",
    },
  });
  return Array.isArray(rows) ? rows : [];
}

async function listAgendaItemsByAgendaIds(ids) {
  const filter = buildInFilter(ids);
  if (!filter) {
    return [];
  }
  const rows = await supabaseRestFetch("agenda_items", {
    query: {
      select: "*",
      agenda_id: filter,
      order: "sort_order.asc,updated_at.desc",
    },
  });
  return Array.isArray(rows) ? rows : [];
}

async function createAgendaRow(row) {
  const rows = await supabaseRestFetch("agendas", {
    method: "POST",
    query: { select: "*" },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: [row],
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

async function updateAgendaRow(agendaId, patch) {
  const rows = await supabaseRestFetch("agendas", {
    method: "PATCH",
    query: {
      select: "*",
      id: `eq.${ensureUuid(agendaId, "Agenda")}`,
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: patch,
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

async function replaceAgendaMembers(agendaId, members) {
  const normalizedAgendaId = ensureUuid(agendaId, "Agenda");
  await supabaseRestFetch("agenda_members", {
    method: "DELETE",
    query: {
      agenda_id: `eq.${normalizedAgendaId}`,
      is_owner: "eq.false",
    },
    headers: {
      Prefer: "return=minimal",
    },
  });

  const rows = (Array.isArray(members) ? members : []).filter(Boolean);
  if (!rows.length) {
    return;
  }

  await supabaseRestFetch("agenda_members", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
    body: rows.map((member) => ({
      agenda_id: normalizedAgendaId,
      user_email: String(member.user_email || "").trim().toLowerCase(),
      display_name: member.display_name || null,
      is_owner: member.is_owner === true,
    })),
  });
}

async function createAgendaMembers(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return;
  }
  await supabaseRestFetch("agenda_members", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
    body: rows,
  });
}

async function createAgendaItemRow(row) {
  const rows = await supabaseRestFetch("agenda_items", {
    method: "POST",
    query: { select: "*" },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: [row],
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

async function updateAgendaItemRow(itemId, patch) {
  const rows = await supabaseRestFetch("agenda_items", {
    method: "PATCH",
    query: {
      select: "*",
      id: `eq.${ensureUuid(itemId, "Agenda item")}`,
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: patch,
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

module.exports = {
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
};
