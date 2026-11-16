# FBS Re — "Save to SharePoint" Outlook add-in

A button inside Outlook that opens a small pop-up where a user can:

1. tick which of the open email's **attachments** to save,
2. choose the **department site → document library** to save them into,
3. fill in that library's **required metadata** (loaded live from SharePoint, so it always matches the library), and
4. click **Save** — the files land in SharePoint with the metadata filled.

Everything runs in the user's browser/Outlook using Microsoft Graph; there is no server to maintain.

---

## What's in this folder

```
manifest.xml              <- the add-in definition you load into Outlook
assets/icon-*.png         <- app icons (the FBS ribbon mark)
src/taskpane.html         <- the pop-up
src/taskpane.css          <- styling (FBS brand)
src/taskpane.js           <- main logic
src/graph.js              <- Microsoft Graph / SharePoint helpers
src/authConfig.js         <- *** you fill in CLIENT_ID + TENANT_ID here ***
src/auth.html             <- sign-in helper page
DEPLOYMENT-GUIDE.md       <- this file
```

You will edit **two** things before it works:
- `src/authConfig.js` → `CLIENT_ID` and `TENANT_ID`
- `manifest.xml` → replace every `https://YOUR-HOST` with your real hosting URL

---

## Step 1 — Host the files (HTTPS)

An Outlook add-in's web files must be served over **HTTPS**. Two easy options — pick one.

### Option A — GitHub Pages (free, ~5 min, no Azure)
1. Create a new GitHub repository, e.g. `save-to-sharepoint`.
2. Upload the contents of this folder (keep the `src/` and `assets/` folders).
3. Repo **Settings → Pages →** Source = `main` branch, `/root`. Save.
4. After a minute your site is live at:
   `https://<your-user-or-org>.github.io/save-to-sharepoint`
   That URL is your **YOUR-HOST** value.

### Option B — Azure Static Web Apps (free tier; you're Global Admin so you have Azure)
1. Azure Portal → **Create a resource → Static Web App** (Free plan).
2. Deploy this folder (via GitHub Actions or the VS Code "Azure Static Web Apps" extension; app location `/`).
3. Your URL looks like `https://<name>.azurestaticapps.net` — that's **YOUR-HOST**.

> Whichever you pick, open `https://YOUR-HOST/src/taskpane.html` in a browser once to confirm it loads (it will say "Sign in").

---

## Step 2 — Register the app in Entra ID (Azure AD)

This lets the add-in sign users in and write to SharePoint.

1. **Entra admin center** (entra.microsoft.com) → **Applications → App registrations → New registration**.
   - Name: `FBS Re – Save to SharePoint`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: platform **Single-page application (SPA)** →
     `https://YOUR-HOST/src/auth.html`
   - Register.
2. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
3. **Authentication →** confirm the SPA redirect URI above is listed. Also add (SPA):
   `https://YOUR-HOST/src/taskpane.html`  ← harmless, helps some Outlook clients.
4. **API permissions → Add a permission:**
   - **Microsoft Graph → Delegated:** `User.Read`, `Sites.ReadWrite.All`, `User.ReadBasic.All`
   - **SharePoint → Delegated:** `AllSites.Write`  *(needed only for required **person** columns — see notes)*
5. Click **Grant admin consent for <tenant>** (you're Global Admin, so this is one click).

---

## Step 3 — Fill in the config

Open `src/authConfig.js` and set:

```js
CLIENT_ID: "the Application (client) ID from step 2",
TENANT_ID: "the Directory (tenant) ID from step 2",
```

Open `manifest.xml` and **replace every** `https://YOUR-HOST` with your real host
(there are several — in the icon URLs, `SourceLocation`, and the `Resources` section).
Save both files and re-upload them to your host (Step 1).

---

## Step 4 — Load the add-in into Outlook (sideload for testing)

You don't need to publish it org-wide yet. Sideload it for yourself/testers:

**Outlook on the web / new Outlook**
1. Open [outlook.office.com](https://outlook.office.com), open any email.
2. **… (More actions) → Get Add-ins** (or **Apps → Add-ins**).
3. **My add-ins → Custom add-ins → Add a custom add-in → Add from file…**
4. Choose your edited `manifest.xml`. Confirm.

**Classic Outlook (Windows desktop)** — same **Get Add-ins → My add-ins → Add a custom add-in → Add from file**.

> To give it to a few testers without each person sideloading: **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps**, upload `manifest.xml`, and assign it to a small test group. (This is also how you'd later roll it out org-wide.)

---

## Step 5 — Try it

1. Open an email that has an attachment.
2. On the ribbon you'll see **FBS Re → Save to SharePoint**. Click it.
3. Sign in (first time only — consent is already granted).
4. Tick the attachment(s) → choose a **site** → choose a **library** → fill the required fields → **Save to SharePoint**.
5. Open the library in SharePoint to confirm the file and its metadata.

---

## Notes, limits & good-to-knows (v1)

- **Which sites show up:** the picker lists sites you **follow** plus sites the search index returns. Tip: in each new department site click **Follow**, and testers will see it immediately.
- **Required fields are enforced by the pop-up.** Uploading via Graph doesn't trigger SharePoint's own "you must fill this in" prompt, so the add-in enforces it on the client side before saving — exactly what you asked for.
- **Person/people columns:** supported via SharePoint's `ensureuser` (that's what the SharePoint `AllSites.Write` permission is for). Enter the person's email. Single-person columns are handled; multi-person columns are a v2 item.
- **Column types handled:** Single line text, multiple lines, choice, number, currency, yes/no, date, hyperlink, and person. Managed-metadata (term store) and lookup columns are v2 items — if a new library needs them as *required*, tell me and I'll add them.
- **File size:** small files upload directly; files over 4 MB use a resumable upload session automatically.
- **Cloud attachments** (OneDrive links in an email) aren't downloaded in v1 — only real file attachments. Easy to add later.
- **Saving the email itself as PDF** isn't in v1 (you chose attachments + pick-each-time). It's a clean add-on when you want it.

## Ideas for v2 (when you're ready)
- Save the whole email as a PDF record.
- Remember each person's last-used site/library.
- Managed-metadata, lookup, and multi-person columns.
- Auto-suggest the department site from the sender's domain or a keyword.
- A "document type" step that routes to the right folder automatically.

---

*Built for FBS Reinsurance Limited · ....For Better Services*
