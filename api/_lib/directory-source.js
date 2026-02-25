const fs = require("fs/promises");
const path = require("path");
const { listCarers, listClients, listVisits } = require("./onetouch-client");

let directoryCache = {
  data: null,
  expiresAtMs: 0,
  inFlight: null,
};

function normalizeString(value) {
  return String(value || "").trim();
}

function sortByNameThenId(a, b) {
  const nameCompare = String(a.name || "").localeCompare(String(b.name || ""), undefined, {
    sensitivity: "base",
  });

  if (nameCompare !== 0) {
    return nameCompare;
  }

  return String(a.id || "").localeCompare(String(b.id || ""), undefined, {
    sensitivity: "base",
  });
}

async function readLocalClients() {
  const filePath = process.env.CLIENTS_DATA_FILE
    ? path.resolve(process.cwd(), process.env.CLIENTS_DATA_FILE)
    : path.join(process.cwd(), "data", "clients.json");

  const raw = await fs.readFile(filePath, "utf8");
  const parsed = JSON.parse(raw);

  if (!Array.isArray(parsed)) {
    throw new Error("Client data must be an array.");
  }

  return parsed.map((item) => ({
    id: normalizeString(item.id),
    name: normalizeString(item.name),
    address: normalizeString(item.address),
    town: normalizeString(item.town),
    county: normalizeString(item.county),
    postcode: normalizeString(item.postcode),
    email: normalizeString(item.email),
    phone: normalizeString(item.phone),
    status: normalizeString(item.status),
    raw: item,
  }));
}

function attachRelationships(clients, carers, visits) {
  const clientMap = new Map(clients.map((client) => [client.id, client]));
  const carerMap = new Map(carers.map((carer) => [carer.id, carer]));

  const clientLinks = new Map();
  const carerLinks = new Map();

  function ensureClientLink(clientId) {
    if (!clientLinks.has(clientId)) {
      clientLinks.set(clientId, {
        carerIds: new Set(),
        visitCount: 0,
        lastVisitAt: "",
      });
    }
    return clientLinks.get(clientId);
  }

  function ensureCarerLink(carerId) {
    if (!carerLinks.has(carerId)) {
      carerLinks.set(carerId, {
        clientIds: new Set(),
        visitCount: 0,
        lastVisitAt: "",
      });
    }
    return carerLinks.get(carerId);
  }

  for (const visit of visits) {
    if (!clientMap.has(visit.clientId) || !carerMap.has(visit.carerId)) {
      continue;
    }

    const clientLink = ensureClientLink(visit.clientId);
    const carerLink = ensureCarerLink(visit.carerId);

    clientLink.carerIds.add(visit.carerId);
    clientLink.visitCount += 1;
    if (visit.startAt && (!clientLink.lastVisitAt || visit.startAt > clientLink.lastVisitAt)) {
      clientLink.lastVisitAt = visit.startAt;
    }

    carerLink.clientIds.add(visit.clientId);
    carerLink.visitCount += 1;
    if (visit.startAt && (!carerLink.lastVisitAt || visit.startAt > carerLink.lastVisitAt)) {
      carerLink.lastVisitAt = visit.startAt;
    }
  }

  const carerById = new Map(carers.map((carer) => [carer.id, carer]));
  const clientById = new Map(clients.map((client) => [client.id, client]));

  const clientsWithRelationships = clients.map((client) => {
    const related = clientLinks.get(client.id);
    const carerIds = related ? Array.from(related.carerIds) : [];
    const relatedCarers = carerIds
      .map((carerId) => carerById.get(carerId))
      .filter(Boolean)
      .map((carer) => ({ id: carer.id, name: carer.name }))
      .sort(sortByNameThenId);

    return {
      ...client,
      relationships: {
        carerCount: relatedCarers.length,
        visitCount: related?.visitCount || 0,
        lastVisitAt: related?.lastVisitAt || "",
        carers: relatedCarers,
      },
    };
  });

  const carersWithRelationships = carers.map((carer) => {
    const related = carerLinks.get(carer.id);
    const clientIds = related ? Array.from(related.clientIds) : [];
    const relatedClients = clientIds
      .map((clientId) => clientById.get(clientId))
      .filter(Boolean)
      .map((client) => ({ id: client.id, name: client.name }))
      .sort(sortByNameThenId);

    return {
      ...carer,
      relationships: {
        clientCount: relatedClients.length,
        visitCount: related?.visitCount || 0,
        lastVisitAt: related?.lastVisitAt || "",
        clients: relatedClients,
      },
    };
  });

  return {
    clients: clientsWithRelationships,
    carers: carersWithRelationships,
  };
}

async function loadDirectoryData() {
  if (process.env.USE_LOCAL_CLIENTS_FALLBACK === "1") {
    const localClients = (await readLocalClients()).sort(sortByNameThenId);
    const withRelationships = localClients.map((client) => ({
      ...client,
      relationships: {
        carerCount: 0,
        visitCount: 0,
        lastVisitAt: "",
        carers: [],
      },
    }));

    return {
      source: "local-fallback",
      clients: withRelationships,
      carers: [],
      warnings: ["Using local client fallback data. OneTouch carers and relationships are unavailable."],
    };
  }

  const [clients, carers, visitsResult] = await Promise.all([
    listClients(),
    listCarers(),
    listVisits().then(
      (visits) => ({ visits, error: "" }),
      (error) => ({ visits: [], error: error?.message || String(error) })
    ),
  ]);

  const sortedClients = clients.sort(sortByNameThenId);
  const sortedCarers = carers.sort(sortByNameThenId);
  const relationshipData = attachRelationships(sortedClients, sortedCarers, visitsResult.visits);

  const warnings = [];
  if (visitsResult.error) {
    warnings.push(`Visits endpoint unavailable: ${visitsResult.error}`);
  }

  return {
    source: "onetouch",
    clients: relationshipData.clients,
    carers: relationshipData.carers,
    warnings,
  };
}

async function readDirectoryData() {
  const now = Date.now();

  if (directoryCache.data && directoryCache.expiresAtMs > now) {
    return directoryCache.data;
  }

  if (directoryCache.inFlight) {
    return directoryCache.inFlight;
  }

  directoryCache.inFlight = loadDirectoryData()
    .then((data) => {
      directoryCache.data = data;
      directoryCache.expiresAtMs = Date.now() + 45_000;
      return data;
    })
    .finally(() => {
      directoryCache.inFlight = null;
    });

  return directoryCache.inFlight;
}

module.exports = {
  readDirectoryData,
};
