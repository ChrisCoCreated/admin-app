const fs = require("fs/promises");
const path = require("path");
const { listCarers, listClients, listVisits } = require("./onetouch-client");

let directoryCache = {
  data: null,
  expiresAtMs: 0,
  inFlight: null,
};
let carersCache = {
  data: null,
  expiresAtMs: 0,
  inFlight: null,
};

function getDirectoryCacheTtlMs() {
  const configured = Number(process.env.ONETOUCH_DIRECTORY_CACHE_TTL_MS || 180000);
  if (!Number.isFinite(configured) || configured < 0) {
    return 180000;
  }
  return Math.floor(configured);
}

function getCarersCacheTtlMs() {
  const configured = Number(process.env.ONETOUCH_CARERS_CACHE_TTL_MS || getDirectoryCacheTtlMs());
  if (!Number.isFinite(configured) || configured < 0) {
    return getDirectoryCacheTtlMs();
  }
  return Math.floor(configured);
}

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

function withTimeout(promise, ms, label) {
  let timer = null;
  const timeoutPromise = new Promise((_, reject) => {
    timer = setTimeout(() => {
      reject(new Error(`${label} timed out after ${ms}ms`));
    }, ms);
  });

  return Promise.race([promise, timeoutPromise]).finally(() => {
    if (timer) {
      clearTimeout(timer);
    }
  });
}

async function timed(label, task) {
  const start = Date.now();
  try {
    const result = await task();
    console.info("[OneTouch] Load step complete", { step: label, elapsedMs: Date.now() - start });
    return result;
  } catch (error) {
    console.warn("[OneTouch] Load step failed", {
      step: label,
      elapsedMs: Date.now() - start,
      error: error?.message || String(error),
    });
    throw error;
  }
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

function attachCarerRelationshipsFromVisits(carers, visits) {
  const carerMap = new Map(carers.map((carer) => [carer.id, carer]));
  const carerLinks = new Map();

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
    if (!carerMap.has(visit.carerId)) {
      continue;
    }
    const link = ensureCarerLink(visit.carerId);
    if (visit.clientId) {
      link.clientIds.add(visit.clientId);
    }
    link.visitCount += 1;
    if (visit.startAt && (!link.lastVisitAt || visit.startAt > link.lastVisitAt)) {
      link.lastVisitAt = visit.startAt;
    }
  }

  return carers.map((carer) => {
    const related = carerLinks.get(carer.id);
    const clients = related
      ? Array.from(related.clientIds)
          .map((clientId) => ({ id: clientId, name: clientId }))
          .sort(sortByNameThenId)
      : [];

    return {
      ...carer,
      relationships: {
        clientCount: clients.length,
        visitCount: related?.visitCount || 0,
        lastVisitAt: related?.lastVisitAt || "",
        clients,
      },
    };
  });
}

async function loadDirectoryData() {
  const loadStartedAt = Date.now();
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

  const clientsTimeoutMs = Number(process.env.ONETOUCH_CLIENTS_TIMEOUT_MS || "12000");
  const carersTimeoutMs = Number(process.env.ONETOUCH_CARERS_TIMEOUT_MS || "12000");
  const visitsTimeoutMs = Number(process.env.ONETOUCH_VISITS_TIMEOUT_MS || "6000");

  const [clients, carers, visitsResult] = await Promise.all([
    timed("clients/all", () => withTimeout(listClients(), clientsTimeoutMs, "clients/all")),
    timed("carers/all", () => withTimeout(listCarers(), carersTimeoutMs, "carers/all")),
    timed("visits", () => withTimeout(listVisits(), visitsTimeoutMs, "visits")).then(
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

  console.info("[OneTouch] Directory load complete", {
    elapsedMs: Date.now() - loadStartedAt,
    clients: sortedClients.length,
    carers: sortedCarers.length,
    visits: visitsResult.visits.length,
    warnings: warnings.length,
  });

  return {
    source: "onetouch",
    clients: relationshipData.clients,
    carers: relationshipData.carers,
    warnings,
  };
}

async function loadCarersDirectoryData() {
  const loadStartedAt = Date.now();

  if (process.env.USE_LOCAL_CLIENTS_FALLBACK === "1") {
    return {
      source: "local-fallback",
      carers: [],
      warnings: ["Using local client fallback data. OneTouch carers are unavailable."],
    };
  }

  const carersTimeoutMs = Number(process.env.ONETOUCH_CARERS_TIMEOUT_MS || "12000");
  const visitsTimeoutMs = Number(process.env.ONETOUCH_CARERS_VISITS_TIMEOUT_MS || "1200");
  const carers = await timed("carers/all", () => withTimeout(listCarers(), carersTimeoutMs, "carers/all"));
  const visitsResult = await timed("visits", () => withTimeout(listVisits(), visitsTimeoutMs, "visits")).then(
    (visits) => ({ visits, error: "" }),
    (error) => ({ visits: [], error: error?.message || String(error) })
  );

  const sortedCarers = carers.sort(sortByNameThenId);
  const carersWithRelationships = attachCarerRelationshipsFromVisits(sortedCarers, visitsResult.visits);
  const warnings = [];
  if (visitsResult.error) {
    warnings.push(`Relationships unavailable in fast mode: ${visitsResult.error}`);
  }

  console.info("[OneTouch] Carers load complete", {
    elapsedMs: Date.now() - loadStartedAt,
    carers: sortedCarers.length,
    visits: visitsResult.visits.length,
    warnings: warnings.length,
  });

  return {
    source: "onetouch",
    carers: carersWithRelationships,
    warnings,
  };
}

async function readDirectoryData() {
  function refreshDirectoryData() {
    if (directoryCache.inFlight) {
      return directoryCache.inFlight;
    }

    directoryCache.inFlight = loadDirectoryData()
      .then((data) => {
        directoryCache.data = data;
        directoryCache.expiresAtMs = Date.now() + getDirectoryCacheTtlMs();
        return data;
      })
      .finally(() => {
        directoryCache.inFlight = null;
      });

    return directoryCache.inFlight;
  }

  const now = Date.now();

  if (directoryCache.data && directoryCache.expiresAtMs > now) {
    return directoryCache.data;
  }

  if (directoryCache.data) {
    void refreshDirectoryData();
    return directoryCache.data;
  }

  return refreshDirectoryData();
}

module.exports = {
  readCarersDirectoryData: async function readCarersDirectoryData() {
    function refreshCarersDirectoryData() {
      if (carersCache.inFlight) {
        return carersCache.inFlight;
      }

      carersCache.inFlight = loadCarersDirectoryData()
        .then((data) => {
          carersCache.data = data;
          carersCache.expiresAtMs = Date.now() + getCarersCacheTtlMs();
          return data;
        })
        .finally(() => {
          carersCache.inFlight = null;
        });

      return carersCache.inFlight;
    }

    const now = Date.now();

    if (carersCache.data && carersCache.expiresAtMs > now) {
      return carersCache.data;
    }

    if (carersCache.data) {
      void refreshCarersDirectoryData();
      return carersCache.data;
    }

    return refreshCarersDirectoryData();
  },
  readDirectoryData,
};
