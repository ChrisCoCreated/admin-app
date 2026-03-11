const ROLE_PAGES = {
  admin: [
    "clients",
    "carers",
    "recruitment",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
    "marketing",
    "photolayout",
  ],
  care_manager: [
    "clients",
    "carers",
    "recruitment",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
  ],
  operations: [
    "clients",
    "carers",
    "recruitment",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
  ],
  marketing: ["marketing", "photolayout", "emailtemplates"],
  photo_layout: ["photolayout"],
  time_only: ["mapping", "drivetime"],
  hr_only: ["carers", "recruitment"],
  clients_only: ["clients"],
  hr_clients: ["clients", "carers", "recruitment"],
  time_clients: ["clients", "mapping", "drivetime"],
  time_hr: ["carers", "recruitment", "mapping", "drivetime"],
  time_hr_clients: ["clients", "carers", "recruitment", "mapping", "drivetime"],
};

const PAGE_META = {
  clients: { href: "./clients.html", label: "Clients" },
  carers: { href: "./carers.html", label: "Carers" },
  recruitment: { href: "./recruitment.html", label: "Recruitment" },
  whiteboard: { href: "./task-whiteboard.html", label: "Tasks" },
  simpletasks: { href: "./simple-tasks.html", label: "Tasks (Simple)" },
  tasks: { href: "./tasks.html", label: "Tasks (Advanced)" },
  mapping: { href: "./mapping.html", label: "Time Mapping" },
  drivetime: { href: "./drive-time-map.html", label: "Drive-Time Map" },
  reports: { href: "./reports.html", label: "Reports" },
  emailtemplates: { href: "./email-templates.html", label: "Email Templates" },
  marketing: { href: "./marketing.html", label: "Marketing" },
  photolayout: { href: "./photo-layout.html", label: "Photo Layout" },
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

export function getPageMeta(pageKey) {
  return PAGE_META[String(pageKey || "").trim().toLowerCase()] || null;
}

export function renderTopNavigation({ role, currentPathname = window.location.pathname } = {}) {
  const nav = document.getElementById("primaryNav");
  if (!nav) {
    return;
  }

  const pages = getAccessiblePages(role);
  const currentPath = normalizePath(currentPathname);
  nav.innerHTML = "";

  if (!pages.length) {
    return;
  }

  const menu = document.createElement("details");
  menu.className = "topnav-menu";
  const summary = document.createElement("summary");
  summary.className = "topnav-summary";
  summary.textContent = "Menu";
  menu.appendChild(summary);

  const panel = document.createElement("div");
  panel.className = "topnav-panel";

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
      link.setAttribute("aria-current", "page");
      summary.textContent = `Menu: ${page.label}`;
    }
    link.addEventListener("click", () => {
      menu.open = false;
    });
    panel.appendChild(link);
  }

  menu.appendChild(panel);
  nav.appendChild(menu);
}
