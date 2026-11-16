/*
  Authentication configuration.

  >>> FILL IN THESE TWO VALUES after you register the app in Entra ID (see DEPLOYMENT-GUIDE):
      CLIENT_ID  – the "Application (client) ID" of your app registration
      TENANT_ID  – your Microsoft 365 tenant/directory ID (or use "organizations")

  Everything else is standard.
*/
window.FBS_AUTH = {
  CLIENT_ID: "REPLACE-WITH-APPLICATION-CLIENT-ID",
  TENANT_ID: "REPLACE-WITH-DIRECTORY-TENANT-ID",

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
