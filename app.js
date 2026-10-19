/* Muir Receipts — PWA
   Flow: sign in (MSAL, redirect) → snap/pick receipt(s) → type + reason → upload to SharePoint
   `Incoming` as the signed-in user via Microsoft Graph, with a `_meta.json` sidecar.
   Offline: submissions queue in IndexedDB and flush when back online.
   Demo mode (?demo=1): no sign-in, uploads are simulated — for previews and UI testing. */
(() => {
  "use strict";
  const APP_VERSION = "0.3.2";
  const CFG = window.MUIR_CONFIG || {};
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const DEMO = new URLSearchParams(location.search).has("demo") || localStorage.getItem("muir.demo") === "1";

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const screens = { signin: $("s-signin"), capture: $("s-capture"), details: $("s-details"), done: $("s-done") };
  const el = {
    who: $("who"), version: $("version"), btnSignin: $("btn-signin"), signinErr: $("signin-err"),
    btnCamera: $("btn-camera"), btnGallery: $("btn-gallery"), btnFile: $("btn-file"),
    inCamera: $("in-camera"), inGallery: $("in-gallery"), inFile: $("in-file"),
    queueBanner: $("queue-banner"), netStatus: $("net-status"), recent: $("recent"),
    thumbs: $("thumbs"), types: $("types"), reason: $("reason"),
    attendeesWrap: $("attendees-wrap"), attendees: $("attendees"), nMuir: $("n-muir"), nClients: $("n-clients"),
    currencies: $("currencies"),
    projectWrap: $("project-wrap"), project: $("project"), projectList: $("project-list"),
    swPersonal: $("sw-personal"), swCash: $("sw-cash"), payLabel: $("pay-label"),
    sendErr: $("send-err"), btnBack: $("btn-back"), btnSend: $("btn-send"), btnLater: $("btn-later"),
    tofinishCard: $("tofinish-card"), tofinish: $("tofinish"),
    doneTitle: $("done-title"), doneText: $("done-text"), btnAnother: $("btn-another"), btnSignout: $("btn-signout"),
  };
  el.version.textContent = `Muir Receipts v${APP_VERSION}${DEMO ? " · demo" : ""}`;

  // ---------- State ----------
  const state = {
    account: null,                 // MSAL account
    user: null,                    // { name, upn }
    lists: structuredClone(CFG.defaults || {}),
    files: [],                     // [{ name, type, blob, previewUrl, kind: 'image'|'pdf' }]
    expenseType: null, currency: "AUD",
    personal: false, cash: false,
    sending: false,
    finishing: null,             // draft being completed: { prefix, files: [names], meta }
  };

  function show(name) {
    Object.values(screens).forEach((s) => s.classList.remove("active"));
    screens[name].classList.add("active");
    window.scrollTo(0, 0);
  }

  // ---------- Auth (MSAL) ----------
  let msal = null;
  function msalInstance() {
    if (msal) return msal;
    msal = new window.msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
        redirectUri: location.origin + location.pathname,
        navigateToLoginRequestUrl: false,
      },
      cache: { cacheLocation: "localStorage" },
    });
    return msal;
  }
  async function initAuth() {
    if (DEMO) {
      state.user = { name: "Demo User", upn: "demo@muir.com.au" };
      return true;
    }
    if (!CFG.clientId || CFG.clientId.startsWith("REPLACE")) {
      el.signinErr.textContent = "App not configured: clientId / tenantId missing in config.js.";
      el.signinErr.classList.remove("hidden");
      return false;
    }
    const m = msalInstance();
    await m.initialize();
    const res = await m.handleRedirectPromise().catch((e) => { showSigninError(e); return null; });
    let account = res?.account || m.getActiveAccount() || m.getAllAccounts()[0] || null;
    if (!account) return false;
    m.setActiveAccount(account);
    state.account = account;
    state.user = { name: account.name || account.username, upn: account.username };
    return true;
  }
  function signIn() {
    el.signinErr.classList.add("hidden");
    msalInstance().loginRedirect({ scopes: CFG.scopes, prompt: "select_account" }).catch(showSigninError);
  }
  function signOut() {
    if (DEMO) { localStorage.removeItem("muir.demo"); location.href = location.pathname; return; }
    msalInstance().logoutRedirect({ account: state.account, postLogoutRedirectUri: location.origin + location.pathname });
  }
  function showSigninError(e) {
    el.signinErr.textContent = "Sign-in failed: " + (e?.errorMessage || e?.message || e);
    el.signinErr.classList.remove("hidden");
  }
  // Silent only. Never navigates away mid-operation (that could lose an unsaved receipt);
  // callers get err.signin = true and the item stays queued until the user signs in again.
  async function token() {
    if (DEMO) {   // test hook: localStorage muir.demoFail = "signin" simulates an expired sign-in
      if (localStorage.getItem("muir.demoFail") === "signin") { const err = new Error("Sign-in expired"); err.signin = true; throw err; }
      return "demo-token";
    }
    const m = msalInstance();
    const req = { scopes: CFG.scopes, account: state.account };
    try {
      const r = await m.acquireTokenSilent(req);
      return r.accessToken;
    } catch (e) {
      if (e instanceof window.msal.InteractionRequiredAuthError) {
        const err = new Error("Sign-in expired"); err.signin = true; throw err;
      }
      throw e;
    }
  }
  // Called when the app opens (nothing unsaved on screen): refresh the sign-in up front.
  // With a live Microsoft session this is a quick bounce, usually with no password.
  async function ensureSignedIn() {
    if (DEMO) return true;
    try { await token(); return true; }
    catch (e) {
      if (e.signin) { await msalInstance().acquireTokenRedirect({ scopes: CFG.scopes, account: state.account }); return false; }
      return true; // network errors etc: carry on, uploads will queue
    }
  }

  // ---------- Graph ----------
  async function graph(path, opts = {}) {
    const t = await token();
    const r = await fetch(path.startsWith("http") ? path : GRAPH + path, {
      ...opts, headers: { Authorization: `Bearer ${t}`, ...(opts.headers || {}) },
    });
    if (!r.ok) {
      let detail = "";
      try { detail = (await r.json())?.error?.message || ""; } catch (_) {}
      throw new Error(`Graph ${r.status}${detail ? ": " + detail : ""}`);
    }
    return r;
  }
  // Upload one file into the Incoming folder. Simple PUT ≤ 4 MB, upload session above that.
  async function uploadFile(name, blob, conflict = "rename") {
    if (DEMO) { await token(); await sleep(400); console.log("[demo] uploaded", name, blob.size, conflict); return { name }; }
    const base = `/drives/${CFG.driveId}/items/${CFG.incomingFolderId}:/${encodeURIComponent(name)}:`;
    if (blob.size <= 4 * 1024 * 1024) {
      const r = await graph(`${base}/content?@microsoft.graph.conflictBehavior=${conflict}`, {
        method: "PUT", headers: { "Content-Type": blob.type || "application/octet-stream" }, body: blob,
      });
      return r.json();
    }
    const s = await (await graph(`${base}/createUploadSession`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename", name } }),
    })).json();
    const chunk = 5 * 320 * 1024; // multiple of 320 KiB as Graph requires
    let last = null;
    for (let start = 0; start < blob.size; start += chunk) {
      const end = Math.min(start + chunk, blob.size);
      const r = await fetch(s.uploadUrl, {
        method: "PUT",
        headers: { "Content-Length": String(end - start), "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` },
        body: blob.slice(start, end),
      });
      if (!r.ok && r.status !== 202) throw new Error(`Upload chunk failed: ${r.status}`);
      last = r;
    }
    return last.status === 202 ? {} : last.json();
  }
  async function loadRemoteConfig() {
    if (!CFG.configFileId || DEMO) return;
    try {
      const r = await graph(`/drives/${CFG.driveId}/items/${CFG.configFileId}/content`);
      const j = await r.json();
      state.lists = { ...state.lists, ...j };
    } catch (e) { console.warn("Remote config not loaded:", e.message); }
  }

  // ---------- Images ----------
  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }
  function loadImage(blob) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Could not read image")); };
      img.src = url;
    });
  }
  // Downscale to maxEdge and re-encode as JPEG. Browsers apply EXIF orientation when decoding
  // into an <img>, so the canvas output is upright.
  async function normaliseImage(file) {
    const { maxEdge = 1600, quality = 0.85 } = CFG.image || {};
    const img = await loadImage(file);
    const scale = Math.min(1, maxEdge / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.round(img.naturalWidth * scale), h = Math.round(img.naturalHeight * scale);
    const c = document.createElement("canvas"); c.width = w; c.height = h;
    c.getContext("2d").drawImage(img, 0, 0, w, h);
    const blob = await new Promise((res) => c.toBlob(res, "image/jpeg", quality));
    return blob || file;
  }

  // ---------- Capture → files ----------
  async function addFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    el.sendErr.classList.add("hidden");
    for (const f of files) {
      if (f.type === "application/pdf" || /\.pdf$/i.test(f.name)) {
        state.files.push({ name: f.name, type: "application/pdf", blob: f, kind: "pdf" });
      } else {
        let blob = f;
        try { blob = await normaliseImage(f); } catch (_) {}
        state.files.push({ name: f.name.replace(/\.[^.]+$/, "") + ".jpg", type: "image/jpeg", blob, kind: "image", previewUrl: URL.createObjectURL(blob) });
      }
    }
    renderThumbs();
    show("details");
    if (!state.expenseType) el.reason.blur();
  }
  function removeFile(i) {
    const f = state.files[i];
    if (f?.previewUrl) URL.revokeObjectURL(f.previewUrl);
    state.files.splice(i, 1);
    if (!state.files.length) { show("capture"); return; }
    renderThumbs();
  }
  function renderThumbs() {
    el.thumbs.innerHTML = "";
    state.files.forEach((f, i) => {
      const d = document.createElement("div"); d.className = "thumb";
      d.innerHTML = f.kind === "image" ? `<img alt="receipt ${i + 1}" src="${f.previewUrl}">` : `<div class="pdf">PDF</div>`;
      const x = document.createElement("button"); x.type = "button"; x.textContent = "×"; x.setAttribute("aria-label", "Remove");
      x.onclick = () => removeFile(i);
      d.appendChild(x); el.thumbs.appendChild(d);
    });
    const add = document.createElement("div"); add.className = "thumb add"; add.textContent = "+"; add.title = "Add another page";
    add.onclick = () => el.inCamera.click();
    el.thumbs.appendChild(add);
  }

  // ---------- Form ----------
  function renderTypes() {
    el.types.innerHTML = "";
    const lastUsed = localStorage.getItem("muir.lastType");
    const types = [...(state.lists.expenseTypes || [])];
    if (lastUsed && types.includes(lastUsed)) { types.splice(types.indexOf(lastUsed), 1); types.unshift(lastUsed); }
    types.forEach((t) => {
      const b = document.createElement("button"); b.type = "button"; b.className = "chip" + (state.expenseType === t ? " on" : ""); b.textContent = t;
      b.onclick = () => { state.expenseType = t; renderTypes(); updateConditional(); };
      el.types.appendChild(b);
    });
    el.reason.placeholder = state.lists.reasonPlaceholder || el.reason.placeholder;
    const projects = state.lists.projects || [];
    el.projectWrap.classList.toggle("hidden", projects.length === 0 && !state.lists.alwaysShowProject);
    el.projectList.innerHTML = projects.map((p) => `<option value="${p.replace(/"/g, "&quot;")}">`).join("");
  }
  function renderCurrencies() {
    el.currencies.innerHTML = "";
    const list = state.lists.currencies || ["AUD", "EUR", "USD", "GBP", "NZD", "SGD", "Other"];
    list.forEach((c) => {
      const b = document.createElement("button"); b.type = "button"; b.className = "chip" + (state.currency === c ? " on" : ""); b.textContent = c;
      b.onclick = () => { state.currency = c; renderCurrencies(); };
      el.currencies.appendChild(b);
    });
  }
  function updateConditional() {
    const needAttendees = (state.lists.attendeeTypes || []).includes(state.expenseType);
    el.attendeesWrap.classList.toggle("hidden", !needAttendees);
  }
  function setSwitch(btn, on) { btn.classList.toggle("on", on); btn.setAttribute("aria-checked", String(on)); }
  function resetForm() {
    state.files.forEach((f) => f.previewUrl && URL.revokeObjectURL(f.previewUrl));
    state.files = []; state.expenseType = null; state.personal = false; state.cash = false; state.currency = "AUD"; state.finishing = null;
    el.reason.value = ""; el.attendees.value = ""; el.project.value = ""; el.nMuir.value = ""; el.nClients.value = "";
    setSwitch(el.swPersonal, false); setSwitch(el.swCash, false); el.payLabel.textContent = "Credit Card";
    el.sendErr.classList.add("hidden");
    [el.inCamera, el.inGallery, el.inFile].forEach((i) => (i.value = ""));
    renderTypes(); renderCurrencies(); updateConditional();
  }
  function validate() {
    if (!state.files.length) return "Add at least one photo or PDF.";
    if (!state.expenseType) return "Pick an expense type.";
    if (el.reason.value.trim().length < 3) return "A short reason is required — finance will chase you for it otherwise.";
    if (!el.attendeesWrap.classList.contains("hidden") && !(Number(el.nMuir.value) + Number(el.nClients.value) > 0)) return "How many Muir employees and clients? (needed for FBT).";
    return null;
  }

  // ---------- Submission ----------
  function stamp(d = new Date()) {
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  }
  function buildSubmission(status = "complete") {
    const now = new Date();
    const local = (state.user.upn.split("@")[0] || "user").replace(/[^a-z0-9._-]/gi, "").toLowerCase();
    const fin = state.finishing;
    const prefix = fin ? fin.prefix : `${stamp(now)}_${local}`;
    const files = fin ? [] : state.files.map((f, i) => ({ name: `${prefix}_${i + 1}${f.kind === "pdf" ? ".pdf" : ".jpg"}`, type: f.type, blob: f.blob }));
    const fileNames = fin ? fin.files : files.map((f) => f.name);
    const meta = {
      schema: "muir-receipt/2", app: `muir-receipts-pwa ${APP_VERSION}`, status,
      submitted_by: { name: state.user.name, upn: state.user.upn },
      captured_at: fin ? fin.meta.captured_at : now.toISOString(), details_at: status === "complete" ? now.toISOString() : null,
      tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
      expense_type: state.expenseType, reason: el.reason.value.trim() || null,
      attendees: el.attendeesWrap.classList.contains("hidden") ? null : (el.attendees.value.trim() || null),
      n_muir_employees: el.attendeesWrap.classList.contains("hidden") ? null : Number(el.nMuir.value || 0),
      n_clients: el.attendeesWrap.classList.contains("hidden") ? null : Number(el.nClients.value || 0),
      project: el.project.value.trim() || null,
      currency: state.currency, payment_type: state.cash ? "Cash" : "Credit Card",
      to_be_reimbursed: state.personal, card: state.personal ? "personal" : "company",
      files: fileNames,
    };
    return { id: prefix + (status === "complete" && fin ? "-details" : ""), prefix, meta, files, metaOnly: !!fin };
  }
  async function transmit(sub) {
    for (const f of sub.files) await uploadFile(f.name, f.blob);
    const metaBlob = new Blob([JSON.stringify(sub.meta, null, 2)], { type: "application/json" });
    // Uploaded last so its presence marks the set complete; a details-later completion replaces the pending meta.
    await uploadFile(`${sub.prefix}_meta.json`, metaBlob, sub.metaOnly ? "replace" : "rename");
  }
  // ---- "No time now": send the photo immediately with status pending_details, keep a local draft to finish ----
  async function sendLater() {
    if (!state.files.length) { el.sendErr.textContent = "Add at least one photo or PDF."; el.sendErr.classList.remove("hidden"); return; }
    if (state.sending) return;
    state.sending = true; el.btnLater.disabled = true; el.btnLater.innerHTML = `<span class="spinner"></span> Sending photo…`;
    const sub = buildSubmission("pending_details");
    const draft = { prefix: sub.prefix, files: sub.files.map((f) => f.name), meta: sub.meta,
      thumb: state.files[0].kind === "image" ? state.files[0].blob : null, pages: state.files.length };
    await queuePut(sub); await draftPut(draft);   // saved on the phone first
    inflight.add(sub.id);
    try {
      if (!navigator.onLine) throw new Error("offline");
      await transmit(sub); await queueDel(sub.id); addRecent(sub, "sent");
    } catch (e) { addRecent(sub, "queued"); }
    finally { inflight.delete(sub.id); }
    state.sending = false; el.btnLater.disabled = false; el.btnLater.textContent = "No time now — send photo, add details later";
    resetForm(); refreshQueueBanner(); renderDrafts();
    showDone("Photo sent — details later", "It's in your To finish list on the home screen. Finance can see the receipt now; the reason and type are still needed.");
  }
  async function openDraft(prefix) {
    const d = (await draftsAll()).find((x) => x.prefix === prefix); if (!d) return;
    resetForm();
    state.finishing = { prefix: d.prefix, files: d.files, meta: d.meta };
    state.files = d.thumb ? [{ name: d.files[0], type: "image/jpeg", blob: d.thumb, kind: "image", previewUrl: URL.createObjectURL(d.thumb) }] : [{ name: d.files[0], type: "application/pdf", blob: new Blob(), kind: "pdf" }];
    renderThumbs(); el.thumbs.querySelectorAll("button, .add").forEach((x) => x.remove());
    el.btnLater.classList.add("hidden"); el.btnSend.textContent = "Send details";
    show("details");
  }
  async function send() {
    const err = validate();
    if (err) { el.sendErr.textContent = err; el.sendErr.classList.remove("hidden"); return; }
    if (state.sending) return;
    state.sending = true; el.btnSend.disabled = true; el.btnSend.innerHTML = `<span class="spinner"></span> Sending…`;
    const sub = buildSubmission("complete");
    const wasFinishing = !!state.finishing;
    localStorage.setItem("muir.lastType", state.expenseType);
    await queuePut(sub);              // saved on the phone before any network call
    inflight.add(sub.id);
    try {
      if (!navigator.onLine) throw new Error("offline");
      await transmit(sub);
      await queueDel(sub.id);
      addRecent(sub, "sent");
      showDone(wasFinishing ? "Details sent" : "Sent", `${sub.meta.expense_type} · ${sub.meta.reason}`);
    } catch (e) {
      addRecent(sub, "queued");
      const why = e.signin ? "Your sign-in has expired — close and reopen the app to sign in; it will then send automatically."
        : navigator.onLine ? "Upload failed (" + e.message + "); it will retry." : "You're offline.";
      showDone("Saved — will send shortly", `${sub.meta.expense_type} · ${sub.meta.reason}. ${why}`);
    } finally {
      inflight.delete(sub.id);
      if (wasFinishing) await draftDel(sub.prefix);
      state.sending = false; el.btnSend.disabled = false; el.btnSend.textContent = "Send to finance"; el.btnLater.classList.remove("hidden");
      resetForm(); refreshQueueBanner(); renderDrafts();
    }
  }
  function showDone(title, text) { el.doneTitle.textContent = title; el.doneText.textContent = text; show("done"); }

  // ---------- Offline queue (IndexedDB) ----------
  const DB_NAME = "muir-receipts", STORE = "queue", DRAFTS = "drafts";
  function db() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB_NAME, 2);
      r.onupgradeneeded = () => {
        const d = r.result;
        if (!d.objectStoreNames.contains(STORE)) d.createObjectStore(STORE, { keyPath: "id" });
        if (!d.objectStoreNames.contains(DRAFTS)) d.createObjectStore(DRAFTS, { keyPath: "prefix" });
      };
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  }
  async function queuePut(sub) {
    const d = await db();
    await new Promise((res, rej) => { const t = d.transaction(STORE, "readwrite"); t.objectStore(STORE).put(sub); t.oncomplete = res; t.onerror = () => rej(t.error); });
  }
  async function queueAll() {
    const d = await db();
    return new Promise((res, rej) => { const r = d.transaction(STORE).objectStore(STORE).getAll(); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  }
  async function queueDel(id) {
    const d = await db();
    await new Promise((res, rej) => { const t = d.transaction(STORE, "readwrite"); t.objectStore(STORE).delete(id); t.oncomplete = res; t.onerror = () => rej(t.error); });
  }
  async function draftPut(d) { const x = await db(); await new Promise((res, rej) => { const t = x.transaction(DRAFTS, "readwrite"); t.objectStore(DRAFTS).put(d); t.oncomplete = res; t.onerror = () => rej(t.error); }); }
  async function draftsAll() { const x = await db(); return new Promise((res, rej) => { const r = x.transaction(DRAFTS).objectStore(DRAFTS).getAll(); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
  async function draftDel(prefix) { const x = await db(); await new Promise((res, rej) => { const t = x.transaction(DRAFTS, "readwrite"); t.objectStore(DRAFTS).delete(prefix); t.oncomplete = res; t.onerror = () => rej(t.error); }); }
  async function renderDrafts() {
    const list = (await draftsAll().catch(() => [])).sort((p, q) => (p.meta.captured_at < q.meta.captured_at ? 1 : -1));
    el.tofinishCard.classList.toggle("hidden", list.length === 0);
    el.tofinish.innerHTML = "";
    list.forEach((d) => {
      const li = document.createElement("li"); li.style.cursor = "pointer";
      const when = new Date(d.meta.captured_at).toLocaleString("en-AU", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
      li.innerHTML = `<span>${when} · ${d.pages} page${d.pages > 1 ? "s" : ""}<br><span class="muted">Tap to add type and reason</span></span><span>›</span>`;
      li.onclick = () => openDraft(d.prefix);
      el.tofinish.appendChild(li);
    });
  }
  let flushing = false;
  const inflight = new Set();
  async function flushQueue() {
    if (flushing || !navigator.onLine || !state.user) return;
    flushing = true;
    try {
      for (const sub of await queueAll()) {
        if (inflight.has(sub.id)) continue;
        try { await transmit(sub); await queueDel(sub.id); addRecent(sub, "sent"); }
        catch (e) { console.warn("Queue item still pending:", sub.id, e.message); break; }
      }
    } finally { flushing = false; refreshQueueBanner(); }
  }
  async function refreshQueueBanner() {
    const n = (await queueAll().catch(() => [])).length;
    el.queueBanner.classList.toggle("hidden", n === 0);
    el.queueBanner.textContent = n ? `${n} receipt${n > 1 ? "s" : ""} waiting to send — will go automatically when online.` : "";
  }
  function setNet() {
    const on = navigator.onLine;
    el.netStatus.classList.toggle("off", !on);
    el.netStatus.querySelector("span:last-child").textContent = on ? "Online" : "Offline — receipts will queue";
    if (on) flushQueue();
  }

  // ---------- Recent list ----------
  function addRecent(sub, status) {
    const list = JSON.parse(localStorage.getItem("muir.recent") || "[]");
    const i = list.findIndex((x) => x.id === sub.id);
    const item = { id: sub.id, at: sub.meta.captured_at, type: sub.meta.expense_type, reason: sub.meta.reason, status };
    if (i >= 0) list[i] = item; else list.unshift(item);
    localStorage.setItem("muir.recent", JSON.stringify(list.slice(0, 8)));
    renderRecent();
  }
  function renderRecent() {
    const list = JSON.parse(localStorage.getItem("muir.recent") || "[]");
    el.recent.innerHTML = list.length ? "" : `<li><span class="muted">No receipts sent from this phone yet.</span></li>`;
    list.forEach((x) => {
      const li = document.createElement("li");
      const when = new Date(x.at).toLocaleDateString("en-AU", { day: "numeric", month: "short" });
      const label = x.type ? `${x.type}<br><span class="muted">${escapeHtml(x.reason || "")}</span>` : `Photo only<br><span class="muted">details still to add</span>`;
      li.innerHTML = `<span>${when} · ${label}</span><span>${x.status === "sent" ? "✓" : "⏳"}</span>`;
      el.recent.appendChild(li);
    });
  }
  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  // ---------- Wire up ----------
  el.btnSignin.onclick = signIn;
  el.btnSignout.onclick = (e) => { e.preventDefault(); signOut(); };
  el.btnCamera.onclick = () => el.inCamera.click();
  el.btnGallery.onclick = () => el.inGallery.click();
  el.btnFile.onclick = () => el.inFile.click();
  [el.inCamera, el.inGallery, el.inFile].forEach((i) => (i.onchange = () => addFiles(i.files)));
  el.swPersonal.onclick = () => { state.personal = !state.personal; setSwitch(el.swPersonal, state.personal); };
  el.swCash.onclick = () => { state.cash = !state.cash; setSwitch(el.swCash, state.cash); el.payLabel.textContent = state.cash ? "Cash" : "Credit Card"; };
  el.btnBack.onclick = () => { if (state.finishing) { resetForm(); el.btnLater.classList.remove("hidden"); el.btnSend.textContent = "Send to finance"; } show("capture"); };
  el.btnSend.onclick = send;
  el.btnLater.onclick = sendLater;
  el.btnAnother.onclick = () => { show("capture"); setTimeout(() => el.inCamera.click(), 50); };
  window.addEventListener("online", setNet); window.addEventListener("offline", setNet);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) flushQueue(); });

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((e) => console.warn("SW registration failed", e));
  }

  (async () => {
    const ok = await initAuth();
    if (!ok) { show("signin"); return; }
    if (!(await ensureSignedIn())) return;   // redirecting to Microsoft; app reloads after
    el.who.textContent = state.user.name;
    await loadRemoteConfig();
    renderTypes(); renderCurrencies(); updateConditional(); renderRecent(); setNet(); refreshQueueBanner(); renderDrafts();
    show("capture");
    flushQueue();
  })();

  // Exposed for tests / console
  window.__muir = { state, normaliseImage, buildSubmission, queueAll, flushQueue, draftsAll, APP_VERSION };
})();
