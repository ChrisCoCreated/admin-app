const ROLE_PAGES = {
  admin: ["clients", "carers", "mapping", "marketing"],
  care_manager: ["clients", "carers", "mapping"],
  operations: ["clients", "carers", "mapping"],
  marketing: ["marketing"],
};

const PAGE_META = {
  clients: { href: "./clients.html", label: "Clients" },
  carers: { href: "./carers.html", label: "Carers" },
  mapping: { href: "./mapping.html", label: "Time Mapping" },
  marketing: { href: "./marketing.html", label: "Marketing" },
};

function normalizeRole(role) {
  return String(role || "").trim().toLowerCase();
}

function normalizePath(pathname) {
  const lastSegment = String(pathname || "").split("/").pop() || "";
  return lastSegment.toLowerCase();
}

export function getAccessiblePages(role) {
  const normalizedRole = normalizeRole(role);
  const pages = ROLE_PAGES[normalizedRole];
  return Array.isArray(pages) ? pages : [];
}

export function renderTopNavigation({ role, currentPathname = window.location.pathname } = {}) {
  const nav = document.getElementById("primaryNav");
  if (!nav) {
    return;
  }

  const pages = getAccessiblePages(role);
  const currentPath = normalizePath(currentPathname);
  nav.innerHTML = "";

  for (const pageKey of pages) {
    const page = PAGE_META[pageKey];
    if (!page) {
      continue;
    }
    const link = document.createElement("a");
    link.className = "topnav-link";
    link.href = page.href;
    link.textContent = page.label;
    if (normalizePath(page.href) === currentPath) {
      link.classList.add("active");
    }
    nav.appendChild(link);
  }
}
