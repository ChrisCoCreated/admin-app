// Shared contract: keep this file identical in Admin and Associates.
const ROLES = Object.freeze({
  superadmin: "Superadmin", manager: "Manager", financeManager: "Finance Manager",
  consultant: "Consultant", careCoordinator: "Care Coordinator", care: "Care", marketing: "Marketing", hr: "HR",
});
const ROLE_DESCRIPTIONS = Object.freeze({
  superadmin: "All features and role management in both apps.",
  manager: "Full Admin access except role management; Manager features in Associates.",
  financeManager: "Finance in Admin; finance expense controls and PPE notifications in Associates.",
  consultant: "Consultant features in both apps; Photo Layout and Enquiries in Admin.",
  careCoordinator: "Operations in Admin; Care Coordinator features in Associates.",
  care: "Mapping and KPIs in Admin; Care and standard features in Associates.",
  marketing: "Marketing features in Admin; standard signed-in features in Associates.",
  hr: "Carers, Recruitment and Timesheets in Admin; standard signed-in features in Associates.",
});
function normalizeEmail(value) { return String(value || "").trim().toLowerCase(); }
function getProtectedEmail() { return normalizeEmail(process.env.SUPER_USER_EMAIL || "chris@planwithcare.co.uk"); }
function failure(message, status = 503) { return Object.assign(new Error(message), { status }); }
function validateRoles(roles) {
  if (!Array.isArray(roles) || roles.some(role => !Object.hasOwn(ROLES, role))) throw failure("Select valid shared roles.", 400);
  return Object.keys(ROLES).filter(role => roles.includes(role));
}
function validateAssignment(body, protectedEmail = getProtectedEmail(), actor = "") {
  const email = normalizeEmail(body?.email);
  if (email.length > 254 || !/^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(email)) throw failure("Enter a valid work email address.", 400);
  const roles = validateRoles(body?.roles);
  if (email === normalizeEmail(protectedEmail) && !roles.includes("superadmin")) throw failure("The configured superadmin must keep the Superadmin role.", 409);
  if (email === normalizeEmail(actor) && !roles.includes("superadmin")) throw failure("You cannot remove your own superadmin access.", 409);
  return { email, roles };
}
function effectiveRoles(email, assignments) {
  const roles = validateRoles(assignments[normalizeEmail(email)] || []);
  return normalizeEmail(email) === getProtectedEmail() ? validateRoles([...roles, "superadmin"]) : roles;
}
module.exports = { ROLES, ROLE_DESCRIPTIONS, normalizeEmail, getProtectedEmail, failure, validateRoles, validateAssignment, effectiveRoles };
