const ROLE_PAGES = {
  admin: ["clients", "carers", "tasks", "mapping", "marketing"],
  care_manager: ["clients", "carers", "tasks", "mapping"],
  operations: ["clients", "carers", "tasks", "mapping"],
  marketing: ["marketing"],
};

const PAGE_META = {
  clients: { href: "./clients.html", label: "Clients" },
  carers: { href: "./carers.html", label: "Carers" },
  tasks: { href: "./tasks.html", label: "Tasks" },
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

export function canAccessPage(role, pageKey) {
  return getAccessiblePages(role).includes(String(pageKey || "").trim().toLowerCase());
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
