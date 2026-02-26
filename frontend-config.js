export const FRONTEND_CONFIG = {
  tenantId: "d838ff12-f626-4a30-b05e-aa4364197307",
  spaClientId: "2aa13448-1ea1-4889-9871-43d950b844e0",
  apiScope: "api://926265aa-d9d6-413b-a11e-40b9d66045c3/client.read",
  graphTaskScopes: ["User.Read", "Tasks.ReadWrite", "Group.Read.All", "Sites.ReadWrite.All"],
  // Keep empty for same-origin hosting. Set to backend origin for split deployments.
  apiBaseUrl: "",
};
