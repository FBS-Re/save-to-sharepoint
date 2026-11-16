/* Save to SharePoint — main task-pane logic */
(function () {
  "use strict";

  let account = null;
  const tokenCache = {};   // scopeKey -> { token, expiresOn(ms) }
  let sites = [];
  let libraries = [];
  let requiredCols = [];
  let currentSite = null;      // {id,name,webUrl}
  let currentLibrary = null;   // {listId,driveId,name}

  const $ = (id) => document.getElementById(id);

  const SESSION_KEY = "fbs_spo_session";

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) return;
    wire();
    // When the pane is pinned, reload for whichever email the user moves to.
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged); } catch (e) {}
    // Restore a previous sign-in so the user isn't asked to sign in again on reopen.
    restoreSession();
  });

  function wire() {
    $("signInBtn").onclick = signIn;
    $("signOutBtn").onclick = signOut;
    $("siteSelect").onchange = onSiteChange;
    $("librarySelect").onchange = onLibraryChange;
    $("saveBtn").onclick = onSave;
  }

  // ---------- auth ----------
  // Sign-in runs in an Office dialog (auth-dialog.html), because pop-ups are
  // blocked inside the task pane AND the dialog does not share storage with the
  // task pane. So the dialog does the Microsoft sign-in and hands the access
  // token straight back to us through messageParent; we keep it in memory here.
  function signIn() {
    setStatus('<span class="spin"></span>Opening sign-in…', "busy");
    getToken(window.FBS_AUTH.SCOPES).then(function () {
      setStatus("", "");
      onSignedIn();
    }).catch(function (e) {
      setStatus("Sign-in failed: " + (e && e.message ? e.message : e), "err");
    });
  }

  // Open the sign-in dialog for a specific set of scopes and get a token back.
  function openAuthDialog(scopes) {
    return new Promise(function (resolve, reject) {
      const base = window.location.href.split("?")[0].replace(/[^/]*$/, "");
      const url = base + "auth-dialog.html?scopes=" + encodeURIComponent(scopes.join(" "));
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 30, promptBeforeOpen: false }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Could not open the sign-in window: " + (res.error && res.error.message)));
          return;
        }
        const dlg = res.value;
        let settled = false;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          let msg = {};
          try { msg = JSON.parse(arg.message); } catch (e) {}
          settled = true;
          dlg.close();
          if (msg.status === "success" && msg.token) resolve(msg);
          else reject(new Error(msg.message || "Sign-in failed"));
        });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
          if (settled) return;
          reject(new Error(arg.error === 12006 ? "Sign-in was cancelled." : "The sign-in window closed before finishing."));
        });
      });
    });
  }

  // Return an access token for the given scopes, reusing a cached one when valid.
  function getToken(scopes) {
    const key = scopes.join(" ");
    const cached = tokenCache[key];
    if (cached && cached.expiresOn - Date.now() > 60000) return Promise.resolve(cached.token);
    return openAuthDialog(scopes).then(function (r) {
      tokenCache[key] = { token: r.token, expiresOn: r.expiresOn || (Date.now() + 50 * 60000) };
      if (r.username) account = { username: r.username };
      persistSession();
      return r.token;
    });
  }

  function signOut() {
    account = null;
    for (const k in tokenCache) delete tokenCache[k];
    clearSession();
    $("mainSection").classList.add("hidden");
    $("signedInAs").classList.add("hidden");
    $("signInBtn").classList.remove("hidden");
    setStatus("", "");
  }

  // ---------- session persistence ----------
  // Remember the signed-in account + tokens so the user stays signed in when the
  // pane is reopened (e.g. after moving to another email), until the token expires.
  function persistSession() {
    try { localStorage.setItem(SESSION_KEY, JSON.stringify({ account: account, tokenCache: tokenCache })); } catch (e) {}
  }
  function clearSession() {
    try { localStorage.removeItem(SESSION_KEY); } catch (e) {}
  }
  function restoreSession() {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      const key = window.FBS_AUTH.SCOPES.join(" ");
      const cached = data.tokenCache && data.tokenCache[key];
      if (data.account && cached && cached.token && cached.expiresOn - Date.now() > 60000) {
        account = data.account;
        Object.keys(data.tokenCache).forEach(function (k) { tokenCache[k] = data.tokenCache[k]; });
        onSignedIn();
      }
    } catch (e) {}
  }

  // Fired when the user selects a different email while the pane is pinned open.
  function onItemChanged() {
    if (!account) return;
    try { loadAttachments(); } catch (e) {}
  }

  async function onSignedIn() {
    $("signInBtn").classList.add("hidden");
    $("signedInAs").classList.remove("hidden");
    $("mainSection").classList.remove("hidden");
    try {
      const token = await getToken(window.FBS_AUTH.SCOPES);
      const me = await FBSGraph.getMe(token);
      $("userName").textContent = me.displayName || me.userPrincipalName;
      loadAttachments();
      loadSites(token);
    } catch (e) {
      setStatus("Could not start: " + e.message, "err");
    }
  }

  // ---------- attachments ----------
  function loadAttachments() {
    const item = Office.context.mailbox.item;
    const atts = (item.attachments || []).filter(
      (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline
    );
    const box = $("attList");
    if (!atts.length) {
      box.innerHTML = '<p class="muted">This email has no file attachments to save.</p>';
      return;
    }
    box.innerHTML = "";
    atts.forEach((a) => {
      const row = document.createElement("label");
      row.className = "att-row";
      row.innerHTML =
        '<input type="checkbox" class="att-cb" checked data-id="' + a.id + '" data-name="' +
        escAttr(a.name) + '">' +
        '<span class="att-name">' + escHtml(a.name) + "</span>" +
        '<span class="att-size">' + fmtSize(a.size) + "</span>";
      box.appendChild(row);
    });
    refreshSaveEnabled();
  }

  // ---------- sites & libraries ----------
  async function loadSites(token) {
    const sel = $("siteSelect");
    try {
      sites = await FBSGraph.getSites(token);
      if (!sites.length) {
        sel.innerHTML = '<option value="">No sites found</option>';
        return;
      }
      sel.innerHTML = '<option value="">Select a department site…</option>' +
        sites.map((s, i) => '<option value="' + i + '">' + escHtml(s.name) + "</option>").join("");
    } catch (e) {
      sel.innerHTML = '<option value="">Error loading sites</option>';
      setStatus("Sites: " + e.message, "err");
    }
  }

  async function onSiteChange() {
    const i = $("siteSelect").value;
    const libSel = $("librarySelect");
    resetMetaForm();
    currentLibrary = null;
    if (i === "") { libSel.disabled = true; libSel.innerHTML = '<option value="">Pick a site first</option>'; refreshSaveEnabled(); return; }
    currentSite = sites[+i];
    libSel.disabled = true;
    libSel.innerHTML = '<option value="">Loading libraries…</option>';
    try {
      const token = await getToken(window.FBS_AUTH.SCOPES);
      libraries = await FBSGraph.getLibraries(token, currentSite.id);
      if (!libraries.length) {
        libSel.innerHTML = '<option value="">No libraries in this site</option>';
        return;
      }
      libSel.innerHTML = '<option value="">Select a library…</option>' +
        libraries.map((l, k) => '<option value="' + k + '">' + escHtml(l.name) + "</option>").join("");
      libSel.disabled = false;
    } catch (e) {
      libSel.innerHTML = '<option value="">Error</option>';
      setStatus("Libraries: " + e.message, "err");
    }
    refreshSaveEnabled();
  }

  async function onLibraryChange() {
    const k = $("librarySelect").value;
    resetMetaForm();
    if (k === "") { currentLibrary = null; refreshSaveEnabled(); return; }
    currentLibrary = libraries[+k];
    const form = $("metaForm");
    form.innerHTML = '<p class="muted"><span class="spin"></span>Loading required fields…</p>';
    try {
      const token = await getToken(window.FBS_AUTH.SCOPES);
      requiredCols = await FBSGraph.getRequiredColumns(token, currentSite.id, currentLibrary.listId);
      renderMetaForm();
    } catch (e) {
      form.innerHTML = '<p class="muted">Could not load fields.</p>';
      setStatus("Fields: " + e.message, "err");
    }
    refreshSaveEnabled();
  }

  // ---------- metadata form ----------
  function resetMetaForm() {
    requiredCols = [];
    $("metaForm").innerHTML = '<p class="muted">Pick a library to load its required fields.</p>';
  }

  function renderMetaForm() {
    const form = $("metaForm");
    if (!requiredCols.length) {
      form.innerHTML = '<p class="muted">This library has no required fields — just save.</p>';
      return;
    }
    form.innerHTML = "";
    requiredCols.forEach((c, idx) => {
      const wrap = document.createElement("div");
      wrap.className = "field";
      const id = "f_" + idx;
      let control = "";
      const dv = c.defaultValue != null ? String(c.defaultValue) : "";

      if (c.type === "choice") {
        const opts = (c.choices || []).map((o) =>
          '<option value="' + escAttr(o) + '"' + (o === dv ? " selected" : "") + ">" + escHtml(o) + "</option>"
        ).join("");
        control = '<select id="' + id + '" class="input" data-type="choice">' +
          '<option value="">— select —</option>' + opts + "</select>";
      } else if (c.type === "boolean") {
        control = '<select id="' + id + '" class="input" data-type="boolean">' +
          '<option value="">— select —</option><option value="true">Yes</option><option value="false">No</option></select>';
      } else if (c.type === "dateTime") {
        control = '<input id="' + id + '" class="input" type="date" data-type="dateTime">';
      } else if (c.type === "number" || c.type === "currency") {
        control = '<input id="' + id + '" class="input" type="number" step="any" data-type="number" value="' + escAttr(dv) + '">';
      } else if (c.type === "note") {
        control = '<textarea id="' + id + '" class="input" rows="2" data-type="note"></textarea>';
      } else if (c.type === "person") {
        control = '<input id="' + id + '" class="input" type="email" placeholder="name@fbsre.ng" data-type="person">';
      } else if (c.type === "url") {
        control = '<input id="' + id + '" class="input" type="url" placeholder="https://…" data-type="url">';
      } else {
        control = '<input id="' + id + '" class="input" type="text" data-type="text" value="' + escAttr(dv) + '">';
      }

      wrap.innerHTML =
        '<label class="lbl" for="' + id + '">' + escHtml(c.displayName) +
        '<span class="req-mark">*</span>' +
        '<span class="chip">' + c.type + "</span></label>" + control;
      wrap.querySelector("[id='" + id + "']").dataset.internal = c.internalName;
      form.appendChild(wrap);
    });
    form.addEventListener("input", refreshSaveEnabled);
    form.addEventListener("change", refreshSaveEnabled);
    refreshSaveEnabled();
  }

  function collectMetadata() {
    // returns {fields:{...}, persons:[{internal,email}], missing:[displayName]}
    const fields = {};
    const persons = [];
    const missing = [];
    requiredCols.forEach((c, idx) => {
      const el = $("f_" + idx);
      const val = (el.value || "").trim();
      if (!val) { missing.push(c.displayName); return; }
      switch (c.type) {
        case "boolean": fields[c.internalName] = (val === "true"); break;
        case "number": case "currency": fields[c.internalName] = parseFloat(val); break;
        case "dateTime": fields[c.internalName] = new Date(val + "T00:00:00").toISOString(); break;
        case "person": persons.push({ internal: c.internalName, email: val }); break;
        default: fields[c.internalName] = val;
      }
    });
    return { fields, persons, missing };
  }

  function refreshSaveEnabled() {
    const anyAtt = document.querySelectorAll(".att-cb:checked").length > 0;
    const ready = !!(currentLibrary && anyAtt);
    $("saveBtn").disabled = !ready;
  }

  // ---------- save ----------
  async function onSave() {
    const checked = Array.from(document.querySelectorAll(".att-cb:checked"));
    if (!checked.length) { setStatus("Tick at least one attachment.", "err"); return; }
    if (!currentLibrary) { setStatus("Choose a destination library.", "err"); return; }

    const meta = collectMetadata();
    if (meta.missing.length) {
      setStatus("Please complete: " + meta.missing.join(", "), "err");
      return;
    }

    $("saveBtn").disabled = true;
    setStatus('<span class="spin"></span>Saving…', "busy");

    try {
      const token = await getToken(window.FBS_AUTH.SCOPES);

      // Resolve any person fields once (needs a SharePoint-scoped token)
      let personFields = {};
      if (meta.persons.length) {
        const spoRoot = new URL(currentSite.webUrl).origin; // https://tenant.sharepoint.com
        const spoToken = await getToken([spoRoot + "/AllSites.Write"]);
        for (const p of meta.persons) {
          const uid = await FBSGraph.ensureSiteUserId(spoToken, currentSite.webUrl, p.email);
          personFields[p.internal + "LookupId"] = String(uid);
        }
      }
      const allFields = Object.assign({}, meta.fields, personFields);

      const folder = ($("folderPath").value || "").trim();
      let done = 0;
      const total = checked.length;
      for (const cb of checked) {
        const name = cb.dataset.name;
        setStatus('<span class="spin"></span>Saving ' + escHtml(name) + " (" + (done + 1) + "/" + total + ")…", "busy");
        const bytes = await readAttachment(cb.dataset.id);
        const item = await FBSGraph.uploadFile(token, currentSite.id, currentLibrary.driveId, folder, name, bytes);
        if (Object.keys(allFields).length) {
          const itemId = await FBSGraph.getListItemId(token, currentSite.id, currentLibrary.driveId, item.id);
          await FBSGraph.patchFields(token, currentSite.id, currentLibrary.listId, itemId, allFields);
        }
        done++;
      }
      setStatus("✓ Saved " + done + (done === 1 ? " file" : " files") + " to " + escHtml(currentLibrary.name) + ".", "ok");
    } catch (e) {
      setStatus("Save failed: " + e.message, "err");
    } finally {
      refreshSaveEnabled();
    }
  }

  // read one attachment's bytes (base64 -> Uint8Array)
  function readAttachment(attachmentId) {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Could not read attachment")); return;
        }
        try {
          if (res.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
            reject(new Error("Unsupported attachment format")); return;
          }
          const bin = atob(res.value.content);
          const arr = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
          resolve(arr);
        } catch (e) { reject(e); }
      });
    });
  }

  // ---------- helpers ----------
  function setStatus(html, cls) {
    const s = $("status");
    s.className = "status " + (cls || "");
    s.innerHTML = html;
  }
  function fmtSize(b) {
    if (b == null) return "";
    if (b < 1024) return b + " B";
    if (b < 1048576) return (b / 1024).toFixed(0) + " KB";
    return (b / 1048576).toFixed(1) + " MB";
  }
  function escHtml(s) { return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function escAttr(s) { return escHtml(s).replace(/'/g, "&#39;"); }
})();
