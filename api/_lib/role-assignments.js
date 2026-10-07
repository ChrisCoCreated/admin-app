const { loadSnapshot, profileFor, listAssignments, saveAssignment } = require('./shared-role-store');
const { ROLES } = require('./shared-roles');
const LEGACY_ROLE = { superadmin:'admin', manager:'admin', financeManager:'finance', consultant:'consultant', careCoordinator:'operations', care:'logged_in', marketing:'marketing', hr:'hr_only', careAdmin:'logged_in' };
function adminRole(roles) {
  if (roles.includes('superadmin') || roles.includes('manager')) return 'admin';
  return roles.length === 1 ? LEGACY_ROLE[roles[0]] : roles.length ? `shared:${roles.join(',')}` : '';
}
async function getRuntimeAuthorizedUsersMap() {
  const snapshot = await loadSnapshot();
  const list = await listAssignments(snapshot);
  const users = new Map();
  users.superadminEmails = new Set();
  users.rolesByEmail = new Map();
  users.canonicalByEmail = new Map();
  for (const user of list.users) {
    if (!user.roles.length) continue;
    for (const email of [user.email, ...user.aliases]) {
      users.set(email, adminRole(user.roles));
      users.rolesByEmail.set(email, user.roles);
      users.canonicalByEmail.set(email, user.email);
      if (user.isSuperadmin) users.superadminEmails.add(email);
    }
  }
  return users;
}
module.exports = { getRuntimeAuthorizedUsersMap, listAssignments, saveAssignment, adminRole, ROLES, profileFor };
