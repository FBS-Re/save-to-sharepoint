/*
  Authentication configuration.

  >>> FILL IN THESE TWO VALUES after you register the app in Entra ID (see DEPLOYMENT-GUIDE):
      CLIENT_ID  – the "Application (client) ID" of your app registration
      TENANT_ID  – your Microsoft 365 tenant/directory ID (or use "organizations")

  Everything else is standard.
*/
window.FBS_AUTH = {
  CLIENT_ID: "23c21ada-3801-475d-bf4b-359eeff5ebff",
  TENANT_ID: "76807519-6b03-4301-ae6c-3e4d68404d51",

  // Microsoft Graph permissions the add-in asks for (delegated – acts as the signed-in user)
  SCOPES: [
    "User.Read",            // who am I
    "Sites.ReadWrite.All",  // list sites/libraries, upload files, write metadata
    "User.ReadBasic.All"    // resolve people for person/author columns
  ],

  // Must EXACTLY match a redirect URI registered on the app (type: Single-page application)
  // This is the auth helper page in this same folder.
  get REDIRECT_URI() {
    return window.location.origin + window.location.pathname.replace(/[^/]*$/, "") + "auth.html";
  }
};

window.FBS_MSAL_CONFIG = {
  auth: {
    clientId: window.FBS_AUTH.CLIENT_ID,
    authority: "https://login.microsoftonline.com/" + window.FBS_AUTH.TENANT_ID,
    redirectUri: window.FBS_AUTH.REDIRECT_URI
  },
  cache: {
    cacheLocation: "localStorage", // survives the task pane reopening
    storeAuthStateInCookie: false
  }
};
