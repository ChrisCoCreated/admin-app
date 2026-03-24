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
    "enquiries",
    "agendas",
    "problems",
    "scorecard",
    "scorecarddefinitions",
    "scorecardgoals",
    "whiteboard",
    "simpletasks",
    "tasks",
    "taskstest",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
    "suppliers",
    "consultant",
    "marketing",
    "photolayout",
  ],
  care_manager: [
    "clients",
    "carers",
    "timesheets",
    "recruitment",
    "enquiries",
    "agendas",
    "scorecard",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
    "suppliers",
  ],
  operations: [
    "clients",
    "carers",
    "timesheets",
    "recruitment",
    "enquiries",
    "agendas",
    "scorecard",
    "whiteboard",
    "simpletasks",
    "tasks",
    "mapping",
    "drivetime",
    "reports",
    "emailtemplates",
    "suppliers",
  ],
  consultant: ["consultant", "agendas"],
  director: ["agendas", "scorecard", "scorecarddefinitions", "scorecardgoals", "suppliers"],
  marketing: ["marketing", "photolayout", "emailtemplates", "agendas"],
  photo_layout: ["photolayout", "agendas"],
  time_only: ["timesheets", "mapping", "drivetime", "agendas"],
  hr_only: ["carers", "timesheets", "recruitment", "agendas"],
  clients_only: ["clients", "agendas"],
  enquiries_only: ["enquiries", "agendas"],
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
  enquiries: { href: "./enquiries.html", label: "Enquiries" },
  agendas: { href: "./agendas.html", label: "Agendas" },
  problems: { href: "./problems.html", label: "Problems to Solve" },
  scorecard: { href: "./scorecard.html", label: "Performance Scorecard" },
  scorecarddefinitions: { href: "./scorecard-definitions.html", label: "Scorecard Setup" },
  scorecardgoals: { href: "./scorecard-goals.html", label: "Goal Setup" },
  whiteboard: { href: "./task-whiteboard.html", label: "Tasks" },
  simpletasks: { href: "./simple-tasks.html", label: "Tasks (Simple)" },
  tasks: { href: "./tasks.html", label: "Tasks (Advanced)" },
  taskstest: { href: "./tasks-test.html", label: "Tasks Test" },
  mapping: { href: "./mapping.html", label: "Time Mapping" },
  drivetime: { href: "./drive-time-map.html", label: "Our Geography", shortcutLabel: "Map" },
  reports: { href: "./reports.html", label: "Reports" },
  emailtemplates: { href: "./email-templates.html", label: "Email Templates", shortcutLabel: "Email" },
  suppliers: { href: "./suppliers.html", label: "Suppliers & Experiences" },
  consultant: { href: "./consultant.html", label: "Consultant" },
  marketing: { href: "./marketing.html", label: "Marketing" },
  photolayout: { href: "./photo-layout.html", label: "Photo Layout" },
};

const ADMIN_HOME_PAGES = ["reports", "agendas", "recruitment", "emailtemplates", "drivetime"];
const MENU_GROUPS = [
  {
    title: "People",
    pages: ["clients", "carers", "recruitment", "enquiries", "suppliers", "consultant"],
  },
  {
    title: "Planning",
    pages: ["agendas", "whiteboard", "simpletasks", "tasks", "taskstest"],
  },
  {
    title: "Time & Geography",
    pages: ["timesheets", "mapping", "drivetime"],
  },
  {
    title: "Performance",
    pages: ["reports", "problems", "scorecard", "scorecarddefinitions", "scorecardgoals"],
  },
  {
    title: "Marketing & Content",
    pages: ["emailtemplates", "marketing", "photolayout"],
  },
];

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

function getGroupedPages(role) {
  const accessiblePages = getAccessiblePages(role);
  const accessibleSet = new Set(accessiblePages);
  const groupedSections = [];
  const usedPages = new Set();

  for (const group of MENU_GROUPS) {
    const pages = group.pages.filter((pageKey) => accessibleSet.has(pageKey));
    if (!pages.length) {
      continue;
    }
    pages.forEach((pageKey) => usedPages.add(pageKey));
    groupedSections.push({ title: group.title, pages });
  }

  const remainingPages = accessiblePages.filter((pageKey) => !usedPages.has(pageKey));
  if (remainingPages.length) {
    groupedSections.push({ title: "Other", pages: remainingPages });
  }

  return groupedSections;
}

export function renderTopNavigation({ role, currentPathname = window.location.pathname } = {}) {
  const nav = document.getElementById("primaryNav");
  if (!nav) {
    return;
  }

  const pages = getAccessiblePages(role);
  const groupedPages = getGroupedPages(role);
  const shortcutPages = getHomePageTiles(role);
  const currentPath = normalizePath(currentPathname);
  const actualRole = getStoredActualRole();
  const canPreviewAsLoggedInUser = actualRole === "admin";
  const actions = nav.parentElement;
  const topbarInner = actions?.parentElement;
  const existingShortcuts = topbarInner?.querySelector(".topbar-shortcuts");
  const signOutBtn = document.getElementById("signOutBtn");
  existingShortcuts?.remove();
  nav.innerHTML = "";

  if (!pages.length && !canPreviewAsLoggedInUser) {
    return;
  }

  if (topbarInner && actions && shortcutPages.length) {
    const shortcuts = document.createElement("div");
    shortcuts.className = "topbar-shortcuts";
    shortcuts.setAttribute("aria-label", "Quick links");

    for (const pageKey of shortcutPages) {
      const page = PAGE_META[pageKey];
      if (!page) {
        continue;
      }
      const link = document.createElement("a");
      link.className = "topbar-shortcut";
      link.href = page.href;
      link.textContent = page.shortcutLabel || page.label;
      if (normalizePath(page.href) === currentPath) {
        link.classList.add("active");
        link.setAttribute("aria-current", "page");
      }
      shortcuts.appendChild(link);
    }

    if (shortcuts.children.length) {
      topbarInner.insertBefore(shortcuts, actions);
    }
  }

  const menu = document.createElement("details");
  menu.className = "topnav-menu";
  const summary = document.createElement("summary");
  summary.className = "topnav-summary";
  summary.textContent = "Menu";
  menu.appendChild(summary);

  const panel = document.createElement("div");
  panel.className = "topnav-panel";
  let previewControl = null;

  if (canPreviewAsLoggedInUser) {
    previewControl = document.createElement("label");
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
  }

  for (const section of groupedPages) {
    const sectionElement = document.createElement("section");
    sectionElement.className = "topnav-section";

    const heading = document.createElement("h3");
    heading.className = "topnav-section-title";
    heading.textContent = section.title;
    sectionElement.appendChild(heading);

    const sectionLinks = document.createElement("div");
    sectionLinks.className = "topnav-section-links";

    for (const pageKey of section.pages) {
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
      }
      link.addEventListener("click", () => {
        menu.open = false;
      });
      sectionLinks.appendChild(link);
    }

    if (!sectionLinks.children.length) {
      continue;
    }

    sectionElement.appendChild(sectionLinks);
    panel.appendChild(sectionElement);
  }

  if (signOutBtn) {
    signOutBtn.hidden = false;
    signOutBtn.classList.add("topnav-signout");
    signOutBtn.classList.remove("secondary");

    const accountActions = document.createElement("div");
    accountActions.className = "topnav-account-actions";
    if (previewControl) {
      accountActions.appendChild(previewControl);
    }
    accountActions.appendChild(signOutBtn);
    panel.appendChild(accountActions);
  } else if (previewControl) {
    panel.appendChild(previewControl);
  }

  menu.appendChild(panel);
  nav.appendChild(menu);
}
