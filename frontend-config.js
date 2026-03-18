export const FRONTEND_CONFIG = {
  tenantId: "d838ff12-f626-4a30-b05e-aa4364197307",
  spaClientId: "2aa13448-1ea1-4889-9871-43d950b844e0",
  sharePointSpaClientId: "3330ff2a-78af-4aa4-a8e2-b50337520c01",
  apiScope: "api://926265aa-d9d6-413b-a11e-40b9d66045c3/client.read",
  graphTaskScopes: ["User.Read", "Tasks.ReadWrite", "Group.Read.All", "Sites.ReadWrite.All"],
  // Keep empty for same-origin hosting. Set to backend origin for split deployments.
  apiBaseUrl: "",
  useOfficeCatchmentMode: true,
  sharePoint: {
    supportTeamSiteUrl: "https://planwithcare.sharepoint.com/sites/SupportTeam",
    thriveCallsSiteUrl: "https://planwithcare.sharepoint.com/sites/ThriveCalls",
    wellbeingSiteUrl: "https://planwithcare.sharepoint.com/sites/Wellbeing",
    photosListPath: "/sites/SupportTeam/Lists/Photos and Lists",
    clientsListPath: "/sites/SupportTeam/Lists/Clients  One Touch",
    expensesListPath: "/sites/SupportTeam/Lists/Expenses",
    ppeListPath: "/sites/SupportTeam/Lists/PPE FirstAid Check",
    suppliersListPath: "/sites/Wellbeing/Lists/Suppliers Database",
    enquiriesListTitle: "Enquiries Log",
  },
  mapOffice: {
    name: "Canterbury Office",
    postcode: "CT1",
    lat: 51.2802,
    lng: 1.0789,
  },
};
