import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { renderTopNavigation } from "./navigation.js?v=20261007";
import { createRoleEditor } from "./shared-role-editor.js";
const auth = createAuthController({ tenantId: FRONTEND_CONFIG.tenantId, clientId: FRONTEND_CONFIG.spaClientId });
const api = createDirectoryApi(auth);
const editor = createRoleEditor({
  request: (method = "GET", body) => method === "PUT" ? api.saveRoleAssignment(body.email, body.roles) : api.getRoleAssignments(),
  ids: { status:"roleStatus", content:"roleContent", form:"assignRoleForm", email:"roleEmail", options:"roleOptions", search:"roleSearch", users:"roleUsers", reset:"resetRoleBtn" },
});
document.getElementById("signOutBtn").addEventListener("click", async () => {
  editor.clear();
  await auth.signOut({ redirectUri: new URL("./index.html", window.location.href).toString() });
  window.location.href = "./index.html";
});
(async () => {
  try {
    if (!(await auth.restoreSession())) { window.location.href = "./index.html"; return; }
    const profile = await api.getCurrentUser();
    if (!profile.isSuperadmin || profile.previewingLoggedInUser) { window.location.href = "./unauthorized.html?page=rolemanagement"; return; }
    renderTopNavigation({ role: profile.role });
    await editor.load();
  } catch (error) { editor.status(error.message || "Could not load shared role management.", true); }
  finally { document.body.classList.remove("auth-pending"); }
})();
