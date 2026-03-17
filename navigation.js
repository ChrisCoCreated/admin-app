import {
  getStoredActualRole,
  isLoggedInUserPreviewEnabled,
  setLoggedInUserPreviewEnabled,
} from "./role-preview.js?v=20260317";

const ROLE_PAGES = {
  admin: [
    "clients",
    "carers",
    "timesheets",
    "recruitment",
    "agendas",
    "scorecard",
    "scorecarddefinitions",
    "scorecardgoals",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
    "consultant",
    "marketing",
    "photolayout",
  ],
  care_manager: [
    "clients",
    "carers",
    "timesheets",
    "recruitment",
    "agendas",
    "scorecard",
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
    "timesheets",
    "recruitment",
    "agendas",
    "scorecard",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
  ],
  consultant: ["consultant", "agendas"],
  director: ["agendas", "scorecard", "scorecarddefinitions", "scorecardgoals"],
  marketing: ["marketing", "photolayout", "emailtemplates", "agendas"],
  photo_layout: ["photolayout", "agendas"],
  time_only: ["timesheets", "mapping", "drivetime", "agendas"],
  hr_only: ["carers", "timesheets", "recruitment", "agendas"],
  clients_only: ["clients", "agendas"],
  hr_clients: ["clients", "carers", "timesheets", "recruitment", "agendas"],
  time_clients: ["clients", "timesheets", "mapping", "drivetime", "agendas"],
  time_hr: ["carers", "timesheets", "recruitment", "mapping", "drivetime", "agendas"],
  time_hr_clients: ["clients", "carers", "timesheets", "recruitment", "mapping", "drivetime", "agendas"],
  logged_in: ["drivetime"],
};

const PAGE_META = {
  clients: { href: "./clients.html", label: "Clients" },
  carers: { href: "./carers.html", label: "Carers" },
  timesheets: { href: "./timesheets.html", label: "Timesheets" },
  recruitment: { href: "./recruitment.html", label: "Recruitment" },
  agendas: { href: "./agendas.html", label: "Agendas" },
  scorecard: { href: "./scorecard.html", label: "Performance Scorecard" },
  scorecarddefinitions: { href: "./scorecard-definitions.html", label: "Scorecard Setup" },
  scorecardgoals: { href: "./scorecard-goals.html", label: "Goal Setup" },
  whiteboard: { href: "./task-whiteboard.html", label: "Tasks" },
  simpletasks: { href: "./simple-tasks.html", label: "Tasks (Simple)" },
  tasks: { href: "./tasks.html", label: "Tasks (Advanced)" },
  mapping: { href: "./mapping.html", label: "Time Mapping" },
  drivetime: { href: "./drive-time-map.html", label: "Our Geography" },
  reports: { href: "./reports.html", label: "Reports" },
  emailtemplates: { href: "./email-templates.html", label: "Email Templates" },
  consultant: { href: "./consultant.html", label: "Consultant" },
  marketing: { href: "./marketing.html", label: "Marketing" },
  photolayout: { href: "./photo-layout.html", label: "Photo Layout" },
};

const ADMIN_HOME_PAGES = ["reports", "agendas", "recruitment", "emailtemplates"];

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
  if (!Array.isArray(pages)) {
    return [];
  }
  if (pages.includes("drivetime")) {
    return pages;
  }
  return [...pages, "drivetime"];
}

export function canAccessPage(role, pageKey) {
  return getAccessiblePages(role).includes(String(pageKey || "").trim().toLowerCase());
}

export function getPageMeta(pageKey) {
  return PAGE_META[String(pageKey || "").trim().toLowerCase()] || null;
}

export function getHomePageTiles(role) {
  const normalizedRole = normalizeRole(role);
  const accessiblePages = getAccessiblePages(normalizedRole);
  if (normalizedRole === "admin") {
    return ADMIN_HOME_PAGES.filter((pageKey) => accessiblePages.includes(pageKey));
  }
  if (accessiblePages.length <= 4) {
    return accessiblePages;
  }
  return [];
}

export function renderTopNavigation({ role, currentPathname = window.location.pathname } = {}) {
  const nav = document.getElementById("primaryNav");
  if (!nav) {
    return;
  }

  const pages = getAccessiblePages(role);
  const currentPath = normalizePath(currentPathname);
  const actualRole = getStoredActualRole();
  const canPreviewAsLoggedInUser = actualRole === "admin";
  nav.innerHTML = "";

  if (!pages.length && !canPreviewAsLoggedInUser) {
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

  if (canPreviewAsLoggedInUser) {
    const previewControl = document.createElement("label");
    previewControl.className = "topnav-preview-toggle";

    const previewInput = document.createElement("input");
    previewInput.type = "checkbox";
    previewInput.checked = isLoggedInUserPreviewEnabled();

    const previewCopy = document.createElement("span");
    previewCopy.className = "topnav-preview-copy";
    previewCopy.innerHTML =
      '<strong>View as logged-in user</strong><span>Hide admin-only permissions and pages until you switch this off.</span>';

    previewInput.addEventListener("change", () => {
      setLoggedInUserPreviewEnabled(previewInput.checked);
      const nextRole = previewInput.checked ? "logged_in" : actualRole;
      const currentPageKey = Object.entries(PAGE_META).find(([, page]) => normalizePath(page.href) === currentPath)?.[0] || "";
      menu.open = false;

      if (currentPageKey && !canAccessPage(nextRole, currentPageKey)) {
        window.location.href = "./index.html";
        return;
      }

      window.location.reload();
    });

    previewControl.appendChild(previewInput);
    previewControl.appendChild(previewCopy);
    panel.appendChild(previewControl);
  }

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
