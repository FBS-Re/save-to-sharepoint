/*
  Thin helpers over Microsoft Graph (and a tiny bit of SharePoint REST for people fields).
  All functions take an access token and throw on HTTP errors with a readable message.
*/
window.FBSGraph = (function () {
  const GRAPH = "https://graph.microsoft.com/v1.0";

  async function http(token, method, url, opts = {}) {
    const headers = Object.assign(
      { Authorization: "Bearer " + token },
      opts.headers || {}
    );
    if (opts.json !== undefined) {
      headers["Content-Type"] = "application/json";
      opts.body = JSON.stringify(opts.json);
    }
    const res = await fetch(url, { method, headers, body: opts.body });
    if (!res.ok) {
      let detail = res.status + " " + res.statusText;
      try {
        const e = await res.json();
        if (e && e.error && e.error.message) detail = e.error.message;
      } catch (_) {}
      throw new Error(detail);
    }
    if (res.status === 204) return null;
    const ct = res.headers.get("content-type") || "";
    return ct.indexOf("application/json") >= 0 ? res.json() : res;
  }

  // ---- me ----
  async function getMe(token) {
    return http(token, "GET", GRAPH + "/me?$select=displayName,mail,userPrincipalName");
  }

  // ---- sites the user can see ----
  async function getSites(token) {
    const map = new Map();
    const add = (s) => {
      if (s && s.id && s.name !== undefined) {
        // skip personal OneDrive sites
        if ((s.webUrl || "").indexOf("/personal/") >= 0) return;
        map.set(s.id, { id: s.id, name: s.displayName || s.name, webUrl: s.webUrl });
      }
    };
    // Sites the user follows (usually their department sites)
    try {
      const f = await http(token, "GET", GRAPH + "/me/followedSites?$select=id,name,displayName,webUrl");
      (f.value || []).forEach(add);
    } catch (_) {}
    // Plus anything the search index knows about
    try {
      const s = await http(token, "GET", GRAPH + "/sites?search=*&$select=id,name,displayName,webUrl&$top=200");
      (s.value || []).forEach(add);
    } catch (_) {}
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
  }

  // ---- document libraries in a site (returns list id + drive id) ----
  async function getLibraries(token, siteId) {
    const lists = await http(
      token, "GET",
      GRAPH + "/sites/" + siteId + "/lists?$select=id,name,displayName,list&$top=200"
    );
    const libs = (lists.value || []).filter(
      (l) => l.list && l.list.template === "documentLibrary" && !l.list.hidden
    );
    // Attach drive ids (needed for upload)
    const out = [];
    for (const l of libs) {
      // Hide the system "Form Templates"/"Style Library" etc. by name where obvious
      if (["Form Templates", "Style Library", "Site Assets", "Site Pages"].indexOf(l.displayName) >= 0) continue;
      try {
        const drive = await http(token, "GET", GRAPH + "/sites/" + siteId + "/lists/" + l.id + "/drive?$select=id");
        out.push({ listId: l.id, driveId: drive.id, name: l.displayName || l.name });
      } catch (_) {
        /* library without a drive – skip */
      }
    }
    return out.sort((a, b) => a.name.localeCompare(b.name));
  }

  // ---- columns of a library; returns the ones a user must fill ----
  async function getRequiredColumns(token, siteId, listId) {
    const cols = await http(
      token, "GET",
      GRAPH + "/sites/" + siteId + "/lists/" + listId + "/columns"
    );
    const skip = new Set([
      "ContentType", "Attachments", "Edit", "LinkTitle", "LinkTitleNoMenu",
      "DocIcon", "ItemChildCount", "FolderChildCount", "_ComplianceFlags"
    ]);
    return (cols.value || [])
      .filter((c) =>
        c.required === true &&
        c.readOnly !== true &&
        c.hidden !== true &&
        !skip.has(c.name) &&
        // a type we can render (text/choice/number/etc.); detect by which sub-object is present
        (c.text || c.choice || c.number || c.currency || c.boolean || c.dateTime || c.personOrGroup || c.hyperlinkOrPicture || c.note)
      )
      .map((c) => ({
        internalName: c.name,
        displayName: c.displayName || c.name,
        type: columnType(c),
        choices: c.choice ? (c.choice.choices || []) : null,
        multi: c.choice ? c.choice.displayAs === "checkBoxes" : (c.personOrGroup ? c.personOrGroup.allowMultipleSelection : false),
        defaultValue: c.defaultValue ? c.defaultValue.value : null,
        raw: c
      }));
  }

  function columnType(c) {
    if (c.choice) return "choice";
    if (c.boolean) return "boolean";
    if (c.number) return "number";
    if (c.currency) return "currency";
    if (c.dateTime) return "dateTime";
    if (c.personOrGroup) return "person";
    if (c.hyperlinkOrPicture) return "url";
    if (c.note) return "note";
    return "text";
  }

  // ---- upload a file (handles small and large) ----
  // folderPath may be "" or "2026/Treaty". Returns the created driveItem.
  async function uploadFile(token, siteId, driveId, folderPath, fileName, bytes) {
    const base = GRAPH + "/sites/" + siteId + "/drives/" + driveId + "/root:/";
    const path = (folderPath ? folderPath.replace(/^\/+|\/+$/g, "") + "/" : "") + encodeURIComponent(fileName);

    if (bytes.byteLength <= 4 * 1024 * 1024) {
      return http(token, "PUT", base + path + ":/content", {
        headers: { "Content-Type": "application/octet-stream" },
        body: bytes
      });
    }
    // Large file: resumable upload session
    const session = await http(token, "POST", base + path + ":/createUploadSession", {
      json: { item: { "@microsoft.graph.conflictBehavior": "replace" } }
    });
    const url = session.uploadUrl;
    const CHUNK = 5 * 1024 * 1024;
    let start = 0, last = null;
    while (start < bytes.byteLength) {
      const end = Math.min(start + CHUNK, bytes.byteLength);
      const res = await fetch(url, {
        method: "PUT",
        headers: {
          "Content-Length": String(end - start),
          "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + bytes.byteLength
        },
        body: bytes.slice(start, end)
      });
      if (!res.ok && res.status !== 202) throw new Error("Upload failed at chunk " + start + " (" + res.status + ")");
      if (res.status === 200 || res.status === 201) last = await res.json();
      start = end;
    }
    return last;
  }

  // ---- get the list item id behind an uploaded driveItem ----
  async function getListItemId(token, siteId, driveId, driveItemId) {
    const li = await http(
      token, "GET",
      GRAPH + "/sites/" + siteId + "/drives/" + driveId + "/items/" + driveItemId + "/listItem?$select=id"
    );
    return li.id;
  }

  // ---- write metadata onto the list item ----
  async function patchFields(token, siteId, listId, itemId, fields) {
    return http(
      token, "PATCH",
      GRAPH + "/sites/" + siteId + "/lists/" + listId + "/items/" + itemId + "/fields",
      { json: fields }
    );
  }

  // ---- resolve a person to the site's user id (for person columns) ----
  // Uses SharePoint REST ensureuser with a SharePoint-scoped token.
  async function ensureSiteUserId(spoToken, webUrl, loginEmail) {
    const res = await fetch(webUrl + "/_api/web/ensureuser", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + spoToken,
        "Content-Type": "application/json;odata=nometadata",
        Accept: "application/json;odata=nometadata"
      },
      body: JSON.stringify({ logonName: loginEmail })
    });
    if (!res.ok) throw new Error("Could not resolve user " + loginEmail + " (" + res.status + ")");
    const u = await res.json();
    return u.Id;
  }

  return {
    getMe, getSites, getLibraries, getRequiredColumns,
    uploadFile, getListItemId, patchFields, ensureSiteUserId, columnType
  };
})();
