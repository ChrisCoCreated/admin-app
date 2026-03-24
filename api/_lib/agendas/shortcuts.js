const { fetchGraphResponse } = require("../graph-app-client");
const { listAgendaSummariesForUser, mapAgendaError } = require("./service");
const { normalizeEmail } = require("./model");

const AGENDA_SHORTCUTS = [
  { label: "Huddle", type: "agenda", terms: ["huddle"] },
  { label: "Operations", type: "agenda", terms: ["operations"] },
  { label: "Leadership", type: "agenda", terms: ["leadership"] },
  { label: "£", type: "person", terms: ["laura"], userEmail: "laura@planwithcare.co.uk", displayName: "Laura" },
  { label: "N", type: "person", terms: ["nathan"], userEmail: "nathan@planwithcare.co.uk", displayName: "Nathan" },
  { label: "R", type: "person", terms: ["rebecca"], userEmail: "rebecca@planwithcare.co.uk", displayName: "Rebecca" },
  { label: "P", type: "person", terms: ["peter"], userEmail: "peter@planwithcare.co.uk", displayName: "Peter" },
  { label: "A", type: "person", terms: ["agota"], userEmail: "agota@planwithcare.co.uk", displayName: "Agota" },
  { label: "Al", type: "person", terms: ["alise"], userEmail: "alise@planwithcare.co.uk", displayName: "Alise" },
  { label: "M", type: "person", terms: ["miska", "michalina"], userEmail: "michalina@thrivehomecare.co.uk", displayName: "Miska" },
  { label: "C", type: "person", terms: ["claire"], userEmail: "claire@planwithcare.co.uk", displayName: "Claire" },
  { label: "Pa", type: "person", terms: ["paul"], userEmail: "paul@planwithcare.co.uk", displayName: "Paul" },
];

function agendaMembers(agenda) {
  return Array.isArray(agenda?.members) ? agenda.members : [];
}

function shortcutHaystack(agenda) {
  const members = agendaMembers(agenda);
  const memberEmails = members.map((member) => normalizeEmail(member.userEmail)).filter(Boolean);
  const memberNames = members.map((member) => String(member.displayName || "").trim().toLowerCase()).filter(Boolean);
  return [
    String(agenda?.title || "").trim().toLowerCase(),
    ...memberEmails,
    ...memberNames,
    ...(Array.isArray(agenda?.participantNames) ? agenda.participantNames.map((value) => String(value || "").trim().toLowerCase()) : []),
    ...(Array.isArray(agenda?.participantEmails) ? agenda.participantEmails.map((value) => normalizeEmail(value)) : []),
  ]
    .filter(Boolean)
    .join(" ");
}

function resolveShortcutAgendaId(shortcut, agendas) {
  const list = Array.isArray(agendas) ? agendas : [];
  if (!list.length) {
    return "";
  }
  const normalizedTerms = (Array.isArray(shortcut?.terms) ? shortcut.terms : [])
    .map((term) => String(term || "").trim().toLowerCase())
    .filter(Boolean);
  if (!normalizedTerms.length) {
    return "";
  }

  if (shortcut?.type === "person") {
    const expectedEmail = normalizeEmail(shortcut.userEmail);
    const exactEmailMatch = list.find((agenda) => {
      const members = agendaMembers(agenda);
      const hasMatchingEmail = members.some((member) => normalizeEmail(member.userEmail) === expectedEmail);
      return hasMatchingEmail && members.length <= 2;
    });
    if (exactEmailMatch?.id) {
      return exactEmailMatch.id;
    }

    const compactAgendaMatch = list.find((agenda) => {
      const members = agendaMembers(agenda);
      if (members.length > 2) {
        return false;
      }
      const title = String(agenda?.title || "").trim().toLowerCase();
      const haystack = shortcutHaystack(agenda);
      return normalizedTerms.some((term) => title.includes(term) || haystack.includes(term));
    });
    if (compactAgendaMatch?.id) {
      return compactAgendaMatch.id;
    }
  }

  const titleMatch = list.find((agenda) => {
    const title = String(agenda?.title || "").trim().toLowerCase();
    return normalizedTerms.some((term) => title.includes(term));
  });
  if (titleMatch?.id) {
    return titleMatch.id;
  }

  const fallbackMatch = list.find((agenda) => {
    const haystack = shortcutHaystack(agenda);
    return normalizedTerms.some((term) => haystack.includes(term));
  });
  return fallbackMatch?.id || "";
}

async function listAgendaShortcutsForUser(email) {
  const payload = await listAgendaSummariesForUser(email);
  const agendas = Array.isArray(payload?.agendas) ? payload.agendas : [];
  return {
    shortcuts: AGENDA_SHORTCUTS.map((shortcut) => ({
      label: shortcut.label,
      type: shortcut.type,
      terms: shortcut.terms,
      displayName: shortcut.displayName || shortcut.label,
      userEmail: shortcut.userEmail || "",
      targetAgendaId: resolveShortcutAgendaId(shortcut, agendas),
      photoUrl: shortcut.userEmail
        ? `/api/agendas/shortcut-photo?email=${encodeURIComponent(shortcut.userEmail)}`
        : "",
      hasPhoto: Boolean(shortcut.userEmail),
    })),
  };
}

async function getAgendaShortcutPhoto(email) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) {
    const error = new Error("Email is required.");
    error.status = 400;
    error.code = "SHORTCUT_EMAIL_REQUIRED";
    throw error;
  }

  const candidates = [
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(normalizedEmail)}/photo/$value`,
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(normalizedEmail)}/photos/48x48/$value`,
  ];

  for (const url of candidates) {
    try {
      const response = await fetchGraphResponse(url, {
        headers: {
          Accept: "image/*",
        },
      });
      const arrayBuffer = await response.arrayBuffer();
      return {
        buffer: Buffer.from(arrayBuffer),
        mimeType: String(response.headers.get("content-type") || "image/jpeg").trim() || "image/jpeg",
      };
    } catch (error) {
      if (Number(error?.status) === 404) {
        continue;
      }
      throw error;
    }
  }

  return null;
}

module.exports = {
  getAgendaShortcutPhoto,
  listAgendaShortcutsForUser,
  mapAgendaError,
};
