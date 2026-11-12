/* Muir Receipts — PWA
   Flow: sign in (MSAL, redirect) → snap/pick receipt(s) → type + reason → upload to SharePoint
   `Incoming` as the signed-in user via Microsoft Graph, with a `_meta.json` sidecar.
   Offline: submissions queue in IndexedDB and flush when back online.
   Demo mode (?demo=1): no sign-in, uploads are simulated — for previews and UI testing. */
(() => {
  "use strict";
  const APP_VERSION = "0.7.0";
  const CFG = window.MUIR_CONFIG || {};
  // v0.4.1 lists live here (not in config.js) so the deployed config.js with the real clientId stays untouched.
  const APP_LISTS = {
      // What the person paying can answer. The review maps each option to a Zoho GL using the cardholder's
      // area/state (Lookups) and domestic/overseas — see project doc 44. Keep names in step with Lookups › App types.
      expenseTypes: [
        "Flights", "Accommodation", "Taxi, rideshare, parking, tolls", "Meals while travelling",
        "Entertainment – staff only", "Entertainment – with clients",
        "Vehicle – fuel", "Vehicle – repairs, rego, other",
        "Software & IT", "Phone & internet", "Office, printing, postage, freight",
        "Tools, consumables, equipment", "Repairs, maintenance, hire",
        "Marketing & trade shows", "Training & courses", "Professional services",
        "Staff amenities & recruitment", "R&D", "Personal – card used by mistake", "Other"
      ],
      // Added to the end of the reason so it reads e.g. "Taxi to airport – travel"
      typeSuffix: {
        "Flights": "travel", "Accommodation": "travel", "Taxi, rideshare, parking, tolls": "travel", "Meals while travelling": "travel meal",
        "Entertainment – staff only": "staff entertainment", "Entertainment – with clients": "client entertainment",
        "Vehicle – fuel": "vehicle fuel", "Vehicle – repairs, rego, other": "vehicle",
        "Software & IT": "IT", "Phone & internet": "phone/internet", "Office, printing, postage, freight": "office",
        "Tools, consumables, equipment": "tools & consumables", "Repairs, maintenance, hire": "repairs & maintenance",
        "Marketing & trade shows": "marketing", "Training & courses": "training", "Professional services": "professional services",
        "Staff amenities & recruitment": "staff costs", "R&D": "R&D", "Personal – card used by mistake": "personal, to repay", "Other": "other"
      },
      // Types that ask how many Muir staff and clients (FBT)
      attendeeTypes: ["Entertainment – staff only", "Entertainment – with clients", "Meals while travelling"],
      // "Whose card or cash paid?" — defaults to the signed-in person; lets e.g. accounts photograph a receipt for someone else.
      // Keep in step with the register's Lookups › Cardholders.
      cardholders: [
        { name: "Andrew Buckley", upn: "andrew.buckley@muir.com.au" },
        { name: "Chris Evans", upn: "chris@muir.com.au" },
        { name: "Cory Ovlaque", upn: "corentin.ovlaque@muir.com.au" },
        { name: "Felipe Bernardes", upn: "felipe.bernardes@muir.com.au" },
        { name: "Grant Morgan", upn: "grant.morgan@muir.com.au" },
        { name: "Kelly Du Bois", upn: "kelly.dubois@muir.com.au" },
        { name: "Kusa Weir", upn: "kusa.weir@muir.com.au" },
        { name: "Max Buckley", upn: "max.buckley@muir.com.au" },
        { name: "Sam Crowle", upn: "sam.crowle@muir.com.au" },
        { name: "Shaun Hancock", upn: "shaun.hancock@muir.com.au" },
      ],
    reasonPlaceholder: "e.g. Taxi to airport, Dinner with Onboard Marine — who was there if paying for others",
  };

  const GRAPH = "https://graph.microsoft.com/v1.0";
  // Cardholder list: SharePoint › 13_Accounting › Corporate Cards › Settings › cardholders.json (schema muir-cardholders/1).
  // Edited on the Cardholders screen (admins only); everyone's app reads it at start; the receipts run syncs it to the register.
  const CARDHOLDERS_FILE_ID = "01NLAK5FAOAGJV3WVO4FHZMN4CUWKLZWOL";
  const DEMO = new URLSearchParams(location.search).has("demo") || localStorage.getItem("muir.demo") === "1";
  // Any unexpected error shows a red bar (with the version) instead of the app silently doing nothing — screenshot it for Max.
  function showCrash(msg) {
    const b = document.getElementById("crash"); if (!b) return;
    b.innerHTML = `Something went wrong (v${APP_VERSION}): ${String(msg).replace(/[<>&]/g, "").slice(0, 160)} — <a href="#" onclick="location.reload();return false">Refresh</a>`;
    b.classList.remove("hidden");
  }
  window.addEventListener("error", (e) => showCrash(e.message || e.error || "error"));
  window.addEventListener("unhandledrejection", (e) => showCrash((e.reason && e.reason.message) || e.reason || "error"));

  // ---------- Install helper (shown when opened in a browser instead of from the home-screen icon) ----------
  let deferredInstall = null;
  window.addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); deferredInstall = e; renderInstall(); });
  window.addEventListener("appinstalled", () => { deferredInstall = null; renderInstall(); });
  function installContext() {
    const ua = navigator.userAgent;
    const standalone = window.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
    const ios = /iPhone|iPad|iPod/.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
    const inApp = /FBAN|FBAV|Instagram|Line\/|LinkedInApp|GSA\/|Outlook|MicrosoftTeams|; wv\)/i.test(ua) || (ios && !/Safari\//.test(ua));
    const iosSafari = ios && !inApp && !/CriOS|FxiOS|EdgiOS|OPiOS/.test(ua);
    return { standalone, ios, iosSafari, inApp, android: /Android/.test(ua), mobile: ios || /Android/.test(ua) };
  }
  function renderInstall() {
    const card = document.getElementById("install-card"), body = document.getElementById("install-body"), arrow = document.getElementById("install-arrow");
    if (!card) return;
    const c = installContext();
    if (c.standalone || DEMO && !new URLSearchParams(location.search).has("install")) { card.classList.add("hidden"); arrow.classList.add("hidden"); return; }
    const share = `<svg class="share" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-label="Share"><path d="M12 3v12"/><path d="M8 7l4-4 4 4"/><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"/></svg>`;
    const link = location.origin + location.pathname;
    let html = "";
    arrow.classList.add("hidden");
    if (c.inApp) {
      html = `<p>This opened inside another app. To install it, open it in ${c.ios ? "<strong>Safari</strong>" : "<strong>Chrome</strong>"}:</p>
        <ol><li>Tap <strong>&#8943;</strong> or the compass icon › <strong>Open in ${c.ios ? "Safari" : "Chrome"}</strong></li><li>Or copy the link and paste it into ${c.ios ? "Safari" : "Chrome"}</li></ol>
        <button class="btn" id="btn-copy-link" style="width:100%;margin-top:10px">Copy link</button>`;
    } else if (c.ios) {
      html = `<ol><li>Tap ${share} <strong>Share</strong> ${c.iosSafari ? "at the bottom of Safari" : "(in the address bar)"}</li>
        <li>Scroll down and tap <strong>Add to Home Screen</strong></li><li>Tap <strong>Add</strong>, then open <strong>Muir Receipts</strong> from your home screen and sign in there</li></ol>`;
      if (c.iosSafari) arrow.classList.remove("hidden");
    } else if (c.android && deferredInstall) {
      html = `<p>One tap and it's on your home screen.</p><button class="btn primary" id="btn-install" style="width:100%">Install Muir Receipts</button>`;
    } else if (c.android) {
      html = `<ol><li>Tap <strong>&#8942;</strong> (top right of Chrome)</li><li>Tap <strong>Install app</strong> or <strong>Add to Home screen</strong></li><li>Open <strong>Muir Receipts</strong> from your home screen</li></ol>`;
    } else {
      html = `<p>Open this link on your phone: <strong>${link}</strong></p>`;
    }
    body.innerHTML = html; card.classList.remove("hidden");
    const bi = document.getElementById("btn-install");
    if (bi) bi.onclick = async () => { deferredInstall.prompt(); await deferredInstall.userChoice.catch(() => null); deferredInstall = null; renderInstall(); };
    const bc = document.getElementById("btn-copy-link");
    if (bc) bc.onclick = async () => { try { await navigator.clipboard.writeText(link); bc.textContent = "Copied — now paste it into " + (c.ios ? "Safari" : "Chrome"); } catch (e) { bc.textContent = link; } };
  }
  document.addEventListener("DOMContentLoaded", renderInstall);
  setTimeout(renderInstall, 0);

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const screens = { loading: $("s-loading"), signin: $("s-signin"), capture: $("s-capture"), details: $("s-details"), done: $("s-done"), admin: $("s-admin"), scan: $("s-scan") };
  const el = {
    who: $("who"), version: $("version"), btnSignin: $("btn-signin"), signinErr: $("signin-err"),
    btnCamera: $("btn-camera"), btnGallery: $("btn-gallery"), btnFile: $("btn-file"),
    inCamera: $("in-camera"), inGallery: $("in-gallery"), inFile: $("in-file"),
    queueBanner: $("queue-banner"), netStatus: $("net-status"), recent: $("recent"),
    thumbs: $("thumbs"), types: $("types"), reason: $("reason"), reasonPreview: $("reason-preview"),
    attendeesWrap: $("attendees-wrap"), attendees: $("attendees"), nMuir: $("n-muir"), nClients: $("n-clients"),
    currencies: $("currencies"),
    projectWrap: $("project-wrap"), project: $("project"), projectList: $("project-list"),
    swPersonal: $("sw-personal"), payCard: $("pay-card"), payCash: $("pay-cash"), belongs: $("belongs"), belongsCurrent: $("belongs-current"), belongsList: $("belongs-list"), belongsHint: $("belongs-hint"),
    sendErr: $("send-err"), btnBack: $("btn-back"), btnSend: $("btn-send"), btnLater: $("btn-later"),
    tofinishCard: $("tofinish-card"), tofinish: $("tofinish"),
    doneTitle: $("done-title"), doneText: $("done-text"), btnAnother: $("btn-another"), btnSignout: $("btn-signout"),
  };
  el.version.textContent = `Muir Receipts v${APP_VERSION}${DEMO ? " · demo" : ""}`;

  // ---------- State ----------
  const state = {
    account: null,                 // MSAL account
    user: null,                    // { name, upn }
    lists: { ...structuredClone(CFG.defaults || {}), ...APP_LISTS },
    files: [],                     // [{ name, type, blob, previewUrl, kind: 'image'|'pdf' }]
    expenseType: null, currency: null,
    personal: false, cash: false, belongs: "",
    sending: false,
    finishing: null,             // draft being completed: { prefix, files: [names], meta }
  };

  function show(name) {
    Object.values(screens).forEach((s) => s.classList.remove("active"));
    screens[name].classList.add("active");
    const hb = document.getElementById("btn-home2");
    if (hb) hb.classList.toggle("hidden", name === "capture" || name === "signin" || name === "loading");
    if (name !== "loading") window.__muirStarted = true;
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
      system: { iframeHashTimeout: 6000 },        // fail fast when iPhone blocks the hidden sign-in refresh
    });
    return msal;
  }
  async function initAuth() {
    if (DEMO) {
      if (localStorage.getItem("muir.demoSlow") === "1") await sleep(30000);   // test hook: a start-up that hangs
      state.user = { name: "Demo User", upn: "demo@muir.com.au" }; renderBelongs(null);
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
    state.user = { name: String(account.name || account.username).split("|")[0].trim(), upn: account.username };
    renderBelongs(null);
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
      // Microsoft only lets a web app refresh its sign-in quietly for about 24 hours; after that the quiet refresh needs
      // a hidden page that iPhone Safari blocks, so it fails with a timeout that looks like a network problem.
      // Treat every such failure as "sign in again" unless the phone really is offline.
      const code = String(e && (e.errorCode || e.name) || "");
      const offline = !navigator.onLine || /no_network_connectivity/.test(code);
      if (e instanceof window.msal.InteractionRequiredAuthError || (!offline && (e instanceof window.msal.BrowserAuthError || /timeout|login_required|consent_required|interaction_required|no_tokens_found|token_refresh_required|block_iframe|no_account/i.test(code)))) {
        const err = new Error("Sign-in expired (" + code + ")"); err.signin = true; throw err;
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
  // iPhone (Safari) bug: a photo saved in the app's storage (IndexedDB) as a Blob often can't be sent with fetch later —
  // it fails with "Load failed" even when online. So receipts are now saved as plain bytes, and anything older is
  // copied into a brand-new Blob just before upload.
  async function bytesOf(b) {
    if (!b) return null;
    if (b instanceof ArrayBuffer) return b;
    try { return await b.arrayBuffer(); }
    catch (e) { const err = new Error("photo-unreadable: " + (e && e.message ? e.message : e)); err.unreadable = true; throw err; }
  }
  async function freshBlob(f) {
    const bytes = f.bytes || await bytesOf(f.blob);
    return new Blob([bytes], { type: f.type || "image/jpeg" });
  }
  function asBlob(x, type) { return !x ? null : x instanceof Blob ? x : new Blob([x], { type: type || "image/jpeg" }); }
  async function uploadFile(name, blob, conflict = "rename") {
    if (DEMO) { await token(); await sleep(400); if (localStorage.getItem("muir.demoFail") === "403") throw new Error("Graph 403: Access denied"); console.log("[demo] uploaded", name, blob.size, conflict); return { name }; }
    const base = `/drives/${CFG.driveId}/items/${CFG.incomingFolderId}:/${encodeURIComponent(name)}:`;
    if (blob.size <= 4 * 1024 * 1024) {
      const r = await graph(`${base}/content?@microsoft.graph.conflictBehavior=${conflict}`, {
        method: "PUT", headers: { "Content-Type": blob.type || "application/octet-stream" }, body: blob,
      });
      return r.json();
    }
    const s = await (await graph(`${base}/createUploadSession`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": conflict, name } }),
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
  async function normaliseImage(file, edge) {
    const { maxEdge: cfgEdge = 1600, quality = 0.85 } = CFG.image || {};
    const maxEdge = edge || cfgEdge;
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
    const fresh = [];
    for (const f of files) {
      if (f.type === "application/pdf" || /\.pdf$/i.test(f.name)) {
        state.files.push({ name: f.name, type: "application/pdf", blob: f, kind: "pdf" });
      } else {
        state.files.push(await scanPage(f)); if (state.files[state.files.length - 1].orig) fresh.push(state.files.length - 1);
      }
    }
    $("scanning").classList.add("hidden");
    renderThumbs();
    // every new photo goes through the crop screen: the person sets the corners (nothing is cropped automatically)
    sc.queue = fresh.slice(1);
    if (fresh.length) { openScan(fresh[0], true); return; }
    show("details");
    if (!state.expenseType) el.reason.blur();
  }
  // ---------- Scanner: find the receipt, straighten it, sharpen the text (scan.js) ----------
  async function scanPage(f) {
    const name = f.name.replace(/\.[^.]+$/, "") + ".jpg";
    $("scanning").classList.remove("hidden");
    try {
      if (!window.MuirScan) throw new Error("scanner not loaded");
      const orig = await normaliseImage(f, 2200);               // kept so the corners can be set and changed later
      const img = await loadImage(orig);
      // best go: the scanner's guess at the receipt's corners, shown on the crop screen for the person to accept or fix.
      // Nothing is cropped until they tap "Use this". No confident guess → a frame a little in from the edges.
      const scan = { ...bestGuess(img), enhance: true, edited: false };
      const blob = await normaliseImage(f);                    // used as-is if they skip the crop
      return { name, type: "image/jpeg", blob, kind: "image", previewUrl: URL.createObjectURL(blob), orig, scan };
    } catch (e) {
      console.warn("Scan skipped:", e.message);
      let blob = f; try { blob = await normaliseImage(f); } catch (_) {}
      return { name, type: "image/jpeg", blob, kind: "image", previewUrl: URL.createObjectURL(blob) };
    }
  }
  function bestGuess(img) {
    const w = img.naturalWidth, h = img.naturalHeight, mx = w * 0.08, my = h * 0.08;
    const frame = { quad: [[mx, my], [w - mx, my], [w - mx, h - my], [mx, h - my]], rot: 0, method: "manual", confident: false };
    try {
      const a = window.MuirScan.analyse(img);
      return a.confident ? { quad: a.quad, rot: a.rot, method: a.method, confident: true } : frame;
    } catch (_) { return frame; }
  }
  const sc = { i: -1, img: null, quad: null, rot: 0, enhance: true, url: null, prevUrl: null, timer: null, queue: [], fresh: false };
  async function openScan(i, fresh = false) {
    const f = state.files[i]; if (!f || !f.orig) return;
    sc.i = i; sc.quad = (f.scan.quad).map((p) => p.slice()); sc.rot = f.scan.rot || 0; sc.enhance = f.scan.enhance !== false;
    if (sc.url) URL.revokeObjectURL(sc.url);
    sc.url = URL.createObjectURL(f.orig); sc.img = await loadImage(f.orig);
    const im = $("scan-img"); im.src = sc.url; await im.decode().catch(() => {});
    sc.fresh = fresh;
    const n = state.files.filter((x) => x.orig).length, k = state.files.slice(0, i + 1).filter((x) => x.orig).length;
    $("scan-hint").textContent = (n > 1 ? `Page ${k} of ${n}. ` : "") + (f.scan.confident && !f.scan.edited
      ? "Best guess shown — check the result below. If it's right, tap Use this; if not, drag the corners or tap Turn."
      : "Drag the four corners onto the corners of the receipt. Use Turn if it's sideways.");
    $("scan-cancel").textContent = fresh ? "Skip crop" : "Cancel";
    setSwitch($("scan-enhance"), sc.enhance);
    show("scan"); drawScan(); previewScan();
  }
  function drawScan() {
    const im = $("scan-img"), k = im.clientWidth / (sc.img.naturalWidth || sc.img.width);
    $("scan-svg").setAttribute("viewBox", `0 0 ${im.clientWidth} ${im.clientHeight}`);
    $("scan-poly").setAttribute("points", sc.quad.map((p) => `${p[0] * k},${p[1] * k}`).join(" "));
    document.querySelectorAll("#scan-stage .handle").forEach((h) => { const p = sc.quad[+h.dataset.i]; h.style.left = p[0] * k + "px"; h.style.top = p[1] * k + "px"; });
  }
  // magnifier above the finger so the corner can be placed exactly
  function loupe(ix, iy, px, py) {
    const c = $("scan-loupe"), x = c.getContext("2d"), Z = 110, zoom = 3 * ($("scan-img").clientWidth / (sc.img.naturalWidth || sc.img.width));
    c.width = c.height = Z * 2; c.classList.remove("hidden");
    x.clearRect(0, 0, Z * 2, Z * 2); x.save(); x.beginPath(); x.arc(Z, Z, Z, 0, Math.PI * 2); x.clip();
    const src = Z / zoom; x.drawImage(sc.img, ix - src, iy - src, src * 2, src * 2, 0, 0, Z * 2, Z * 2); x.restore();
    x.strokeStyle = "#e8a33d"; x.lineWidth = 3; x.beginPath(); x.moveTo(Z - 18, Z); x.lineTo(Z + 18, Z); x.moveTo(Z, Z - 18); x.lineTo(Z, Z + 18); x.stroke();
    const above = py > 140; c.style.left = Math.max(0, px - 55) + "px"; c.style.top = (above ? py - 135 : py + 25) + "px";
  }
  function previewScan() {
    clearTimeout(sc.timer);
    sc.timer = setTimeout(async () => {
      const b = await window.MuirScan.render(sc.img, { quad: sc.quad, rot: sc.rot, enhance: sc.enhance, maxEdge: 700, quality: 0.7 });
      if (sc.prevUrl) URL.revokeObjectURL(sc.prevUrl);
      sc.prevUrl = URL.createObjectURL(b); $("scan-preview").src = sc.prevUrl;
    }, 120);
  }
  function wireScan() {
    const stage = $("scan-stage");
    let drag = -1;
    stage.querySelectorAll(".handle").forEach((h) => {
      h.addEventListener("pointerdown", (e) => { drag = +h.dataset.i; h.setPointerCapture(e.pointerId); e.preventDefault(); });
      h.addEventListener("pointermove", (e) => {
        if (drag !== +h.dataset.i) return;
        const im = $("scan-img"), r = im.getBoundingClientRect(), k = im.clientWidth / (sc.img.naturalWidth || sc.img.width);
        const x = Math.max(0, Math.min(r.width, e.clientX - r.left)) / k, y = Math.max(0, Math.min(r.height, e.clientY - r.top)) / k;
        sc.quad[drag] = [x, y]; drawScan(); loupe(x, y, e.clientX - r.left, e.clientY - r.top);
      });
      const end = () => { if (drag >= 0) { drag = -1; $("scan-loupe").classList.add("hidden"); previewScan(); } };
      h.addEventListener("pointerup", end); h.addEventListener("pointercancel", end);
    });
    $("scan-rotate").onclick = () => { sc.rot = (sc.rot + 90) % 360; previewScan(); };
    $("scan-guess").onclick = () => { const g = bestGuess(sc.img); sc.quad = g.quad.map((p) => p.slice()); sc.rot = g.rot; drawScan(); previewScan();
      $("scan-hint").textContent = g.confident ? "Best guess shown — tap Use this if it's right, or drag the corners." : "Couldn't find the receipt on its own — drag the corners onto it."; };
    $("scan-whole").onclick = () => { const w = sc.img.naturalWidth, h = sc.img.naturalHeight; sc.quad = [[0, 0], [w, 0], [w, h], [0, h]]; drawScan(); previewScan(); };
    $("scan-enhance").onclick = () => { sc.enhance = !sc.enhance; setSwitch($("scan-enhance"), sc.enhance); previewScan(); };
    const next = () => {
      renderThumbs();
      if (sc.queue.length) { openScan(sc.queue.shift(), true); return; }
      show("details"); if (!state.expenseType) el.reason.blur();
    };
    $("scan-cancel").onclick = () => { if (!sc.fresh) sc.queue = []; next(); };
    $("scan-done").onclick = async () => {
      const f = state.files[sc.i]; if (!f) { show("details"); return; }
      $("scanning").classList.remove("hidden");
      try {
        const blob = await window.MuirScan.render(sc.img, { quad: sc.quad, rot: sc.rot, enhance: sc.enhance, maxEdge: (CFG.image || {}).maxEdge || 1600 });
        if (f.previewUrl) URL.revokeObjectURL(f.previewUrl);
        f.blob = blob; f.previewUrl = URL.createObjectURL(blob);
        f.scan = { ...f.scan, quad: sc.quad.map((p) => p.slice()), rot: sc.rot, enhance: sc.enhance, edited: true };
      } finally { $("scanning").classList.add("hidden"); }
      next();
    };
    window.addEventListener("resize", () => { if (screens.scan.classList.contains("active")) drawScan(); });
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
      d.innerHTML = f.kind === "image" ? `<img alt="receipt ${i + 1}" src="${f.previewUrl}">${f.orig ? '<span class="adj">Adjust</span>' : ""}` : `<div class="pdf">PDF</div>`;
      if (f.orig) { d.style.cursor = "pointer"; d.onclick = (e) => { if (e.target.tagName !== "BUTTON") openScan(i); }; }
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
      b.onclick = () => { state.expenseType = t; renderTypes(); updateConditional(); updateReasonPreview(); };
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
  // Reason = what the person typed + " – " + a short word for the type, e.g. "Taxi to airport – travel"
  function typeSuffix() { const m = state.lists.typeSuffix || {}; return state.expenseType ? (m[state.expenseType] || state.expenseType.toLowerCase()) : ""; }
  function fullReason() { const t = el.reason.value.trim(); const s = typeSuffix(); return t ? (s && !t.toLowerCase().endsWith(s.toLowerCase()) ? `${t} – ${s}` : t) : ""; }
  function updateReasonPreview() {
    const t = el.reason.value.trim(), s = typeSuffix();
    el.reasonPreview.textContent = s ? (t ? `Sends as: \u201c${fullReason()}\u201d` : `Type what it was for — \u201c– ${s}\u201d is added for you`) : "";
  }
  function updateConditional() {
    const needAttendees = (state.lists.attendeeTypes || []).includes(state.expenseType);
    el.attendeesWrap.classList.toggle("hidden", !needAttendees);
    // meals: start at 1 Muir person (the payer); a solo meal must be ticked as confirmed
    if (needAttendees && el.nMuir.value === "" && el.nClients.value === "") el.nMuir.value = "1";
    updateSolo();
  }
  function soloBox() { return document.getElementById("solo-confirm") || { checked: false }; }
  function isSolo() {
    return !el.attendeesWrap.classList.contains("hidden") && Number(el.nMuir.value || 0) + Number(el.nClients.value || 0) === 1;
  }
  function updateSolo() {
    const solo = isSolo();
    const sw = document.getElementById("solo-wrap"); if (sw) sw.classList.toggle("hidden", !solo);
    if (!solo) soloBox().checked = false;
  }
  function setSwitch(btn, on) { btn.classList.toggle("on", on); btn.setAttribute("aria-checked", String(on)); }
  // ---------- Cardholder list (shared, in SharePoint) ----------
  const lc = (x) => String(x || "").trim().toLowerCase();
  function applyCardholders(doc) {
    state.chDoc = doc;
    state.lists.cardholders = (doc.cardholders || []).filter((c) => c.active !== false && c.email)
      .map((c) => ({ name: c.name, upn: lc(c.email) })).sort((a, b) => a.name.localeCompare(b.name));
    state.isAdmin = !!state.user && (doc.admins || []).map(lc).includes(lc(state.user.upn));
    $("admin-link").classList.toggle("hidden", !state.isAdmin);
    renderBelongs(state.belongs || null);
  }
  function builtInDoc() {
    return { schema: "muir-cardholders/1", admins: DEMO ? ["demo@muir.com.au"] : [], builtIn: true,
      areas: ["Management & Admin", "Sales", "Sales International", "Operations"], states: ["TAS", "QLD", "NSW", "VIC", "SA", "WA", "NT", "ACT", "Overseas"],
      cardholders: APP_LISTS.cardholders.map((c) => ({ name: c.name, email: c.upn, card_last4: "", manager: "", area: "", state: "", other_emails: [], active: true })) };
  }
  async function loadCardholders() {
    if (DEMO) { applyCardholders(state.chDoc && !state.chDoc.builtIn ? state.chDoc : { ...builtInDoc(), builtIn: false }); state.chEtag = "demo"; return; }
    try {
      const meta = await (await graph(`/drives/${CFG.driveId}/items/${CARDHOLDERS_FILE_ID}?select=id,eTag,lastModifiedDateTime`)).json();
      const doc = await (await graph(`/drives/${CFG.driveId}/items/${CARDHOLDERS_FILE_ID}/content`, { cache: "no-store" })).json();
      if (doc.schema !== "muir-cardholders/1" || !Array.isArray(doc.cardholders)) throw new Error("unexpected file");
      state.chEtag = meta.eTag;
      try { localStorage.setItem("muir.cardholders", JSON.stringify(doc)); } catch (_) {}
      applyCardholders(doc);
    } catch (e) {
      console.warn("Cardholder list not loaded:", e.message);
      let cached = null; try { cached = JSON.parse(localStorage.getItem("muir.cardholders") || "null"); } catch (_) {}
      applyCardholders(cached || builtInDoc());
      state.chEtag = null;                              // can't save without a fresh copy
    }
  }
  async function saveCardholders(doc) {
    doc.updated_at = new Date().toISOString(); doc.updated_by = state.user.upn;
    if (DEMO) { await sleep(300); applyCardholders(doc); return; }
    if (!state.chEtag) throw new Error("The list couldn't be loaded from SharePoint, so it can't be saved. Check your connection and reopen the app.");
    const body = JSON.stringify(doc, null, 2);
    try {
      const r = await graph(`/drives/${CFG.driveId}/items/${CARDHOLDERS_FILE_ID}/content`, {
        method: "PUT", headers: { "Content-Type": "application/json", "If-Match": state.chEtag }, body });
      const j = await r.json(); state.chEtag = j.eTag || state.chEtag;
    } catch (e) {
      if (/Graph 412/.test(e.message)) { await loadCardholders(); throw new Error("Someone else changed the list just now. It has been reloaded — please make your change again."); }
      if (/Graph 403|Graph 401/.test(e.message)) throw new Error("SharePoint says you can't edit the cardholder list. Ask Max for edit access to Corporate Cards › Settings.");
      throw e;
    }
    try { localStorage.setItem("muir.cardholders", JSON.stringify(doc)); } catch (_) {}
    applyCardholders(doc);
  }

  // ---------- Cardholders screen (admins) ----------
  const ad = { editing: null, manager: "", area: "", st: "", active: true };
  function adminErr(id, msg) { const x = $(id); x.textContent = msg || ""; x.classList.toggle("hidden", !msg); }
  function renderAdmin() {
    const doc = state.chDoc || builtInDoc();
    const list = $("admin-list"); list.innerHTML = "";
    const people = (doc.cardholders || []).slice().sort((a, b) => (b.active !== false) - (a.active !== false) || a.name.localeCompare(b.name));
    people.forEach((c) => {
      const li = document.createElement("li"); if (c.active === false) li.className = "inactive";
      const left = document.createElement("div");
      const n = document.createElement("div"); n.textContent = c.name + (c.active === false ? " (inactive)" : ""); n.style.fontWeight = "600";
      const sub = document.createElement("div"); sub.className = "sub";
      sub.textContent = [c.card_last4 ? "Card ••" + c.card_last4 : "no card digits", c.manager ? "Mgr " + c.manager.split(" ")[0] : "", c.area, c.state].filter(Boolean).join(" · ");
      left.append(n, sub);
      const e = document.createElement("span"); e.className = "edit"; e.textContent = "Edit ›";
      li.append(left, e); li.onclick = () => openAdminEdit(c.name);
      list.appendChild(li);
    });
    const when = doc.updated_at ? new Date(doc.updated_at).toLocaleString("en-AU", { dateStyle: "medium", timeStyle: "short" }) : "";
    $("admin-meta").textContent = doc.builtIn ? "Showing the built-in list — the SharePoint file couldn't be read." :
      `${people.filter((c) => c.active !== false).length} active · last changed ${when}${doc.updated_by ? " by " + String(doc.updated_by).split(" ")[0] : ""}`;
    $("ad-admins").value = (doc.admins || []).join("\n");
    adminErr("admin-err", "");
    $("admin-main").classList.remove("hidden"); $("admin-edit").classList.add("hidden");
  }
  function chipRow(id, options, current, onPick) {
    const box = $(id); box.innerHTML = "";
    ["", ...options].forEach((o) => {
      const b = document.createElement("button"); b.type = "button"; b.className = "chip" + (current === o ? " on" : "");
      b.textContent = o || "—"; b.onclick = () => { onPick(o); chipRow(id, options, o, onPick); };
      box.appendChild(b);
    });
  }
  function openAdminEdit(name) {
    const doc = state.chDoc || builtInDoc();
    const c = name ? doc.cardholders.find((x) => x.name === name) : null;
    ad.editing = c ? c.name : null;
    ad.manager = c ? c.manager || "" : ""; ad.area = c ? c.area || "" : ""; ad.st = c ? c.state || "" : ""; ad.active = c ? c.active !== false : true;
    $("ad-title").textContent = c ? c.name : "Add a cardholder";
    $("ad-name").value = c ? c.name : ""; $("ad-name").readOnly = !!c; $("ad-name-hint").classList.toggle("hidden", !c);
    $("ad-email").value = c ? c.email || "" : ""; $("ad-last4").value = c ? c.card_last4 || "" : "";
    $("ad-other").value = c ? (c.other_emails || []).join("\n") : "";
    const others = doc.cardholders.filter((x) => x.active !== false && x.name !== (c && c.name)).map((x) => x.name);
    chipRow("ad-manager", others, ad.manager, (v) => (ad.manager = v));
    chipRow("ad-area", doc.areas || [], ad.area, (v) => (ad.area = v));
    chipRow("ad-state", doc.states || [], ad.st, (v) => (ad.st = v));
    setSwitch($("ad-active"), ad.active);
    adminErr("ad-err", "");
    $("admin-main").classList.add("hidden"); $("admin-edit").classList.remove("hidden");
    window.scrollTo(0, 0);
  }
  async function saveAdminEdit() {
    const doc = JSON.parse(JSON.stringify(state.chDoc || builtInDoc())); delete doc.builtIn;
    const name = $("ad-name").value.trim().replace(/\s+/g, " ");
    const email = lc($("ad-email").value);
    const last4 = $("ad-last4").value.trim();
    const other = $("ad-other").value.split(/[\s,;]+/).map(lc).filter(Boolean);
    const rest = doc.cardholders.filter((x) => x.name !== ad.editing);
    const emailRe = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
    let err = null;
    if (!/^[A-Za-zÀ-ÿ'’.-]+( [A-Za-zÀ-ÿ'’.-]+)+$/.test(name)) err = "Enter their full name (first and last).";
    else if (!ad.editing && rest.some((x) => lc(x.name) === lc(name))) err = `${name} is already on the list.`;
    else if (!/^[^@\s]+@muir\.com\.au$/.test(email)) err = "Their Muir email should end in @muir.com.au.";
    else if (rest.some((x) => lc(x.email) === email)) err = "Another cardholder already has that email.";
    else if (ad.active && !/^\d{4}$/.test(last4)) err = "Card last 4 digits: four numbers, as on the NAB statement.";
    else if (last4 && rest.some((x) => x.active !== false && x.card_last4 === last4)) err = `Card ••${last4} already belongs to ${rest.find((x) => x.active !== false && x.card_last4 === last4).name}.`;
    else if (other.some((o) => !emailRe.test(o))) err = "One of the other emails doesn't look like an email address.";
    else if (other.some((o) => o.endsWith("@muir.com.au"))) err = "Other emails are for non-Muir addresses (personal, partner).";
    else if (other.some((o) => rest.some((x) => (x.other_emails || []).map(lc).includes(o)))) err = "One of those other emails is already listed under someone else.";
    if (err) { adminErr("ad-err", err); return; }
    const rec = { name, email, card_last4: last4, manager: ad.manager, area: ad.area, state: ad.st, other_emails: other, active: ad.active };
    if (ad.editing) doc.cardholders = doc.cardholders.map((x) => (x.name === ad.editing ? { ...x, ...rec } : x));
    else doc.cardholders.push(rec);
    $("btn-ad-save").disabled = true;
    try { await saveCardholders(doc); renderAdmin(); }
    catch (e) { adminErr("ad-err", e.message); }
    finally { $("btn-ad-save").disabled = false; }
  }
  async function saveAdmins() {
    const doc = JSON.parse(JSON.stringify(state.chDoc || builtInDoc())); delete doc.builtIn;
    const list = $("ad-admins").value.split(/[\s,;]+/).map(lc).filter(Boolean);
    if (!list.includes(lc(state.user.upn))) { adminErr("admin-err", "Keep yourself on the list, or you'll lock yourself out of this screen."); return; }
    if (list.some((x) => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x))) { adminErr("admin-err", "One of those isn't an email address."); return; }
    doc.admins = list;
    try { await saveCardholders(doc); renderAdmin(); adminErr("admin-err", ""); }
    catch (e) { adminErr("admin-err", e.message); }
  }
  async function openAdmin() { show("admin"); $("admin-meta").textContent = "Loading…"; await loadCardholders(); renderAdmin(); }

  // ---------- Whose card or cash (default: me) ----------
  // Custom picker (not a native <select>: its pop-up list rendered blank in desktop Chrome on Windows).
  function renderBelongs(selectedUpn) {
    if (!el.belongs || !state.user) return;
    const me = (state.user.upn || "").toLowerCase();
    const others = (state.lists.cardholders || []).filter((c) => c.upn.toLowerCase() !== me);
    state.belongs = selectedUpn && others.some((c) => c.upn === selectedUpn) ? selectedUpn : "";
    el.belongsList.innerHTML = "";
    const add = (value, label) => {
      const b = document.createElement("button"); b.type = "button"; b.setAttribute("role", "option");
      const on = state.belongs === value;
      b.className = "chip" + (on ? " on" : "") + (on && value ? " behalf" : ""); b.setAttribute("aria-selected", String(on));
      b.textContent = label;
      b.onclick = () => { renderBelongs(value); toggleBelongsList(false); };
      el.belongsList.appendChild(b);
    };
    add("", `Me (${state.user.name})`);
    others.forEach((c) => add(c.upn, c.name));
    updateBelongsHint();
  }
  function toggleBelongsList(open) {
    const show = open === undefined ? el.belongsList.classList.contains("hidden") : open;
    el.belongsList.classList.toggle("hidden", !show); el.belongs.setAttribute("aria-expanded", String(show));
    el.belongs.querySelector(".picker-change").textContent = show ? "Close ▴" : "Change ▾";
  }
  function belongsTo() {
    const v = state.belongs || "";
    const c = v && (state.lists.cardholders || []).find((x) => x.upn === v);
    return c ? { name: c.name, upn: c.upn, onBehalf: true } : { name: state.user.name, upn: state.user.upn, onBehalf: false };
  }
  function updateBelongsHint() {
    const b = belongsTo();
    el.belongs.classList.toggle("behalf", b.onBehalf);
    el.belongsCurrent.textContent = b.onBehalf ? b.name : `Me (${state.user.name})`;
    el.belongsHint.textContent = b.onBehalf ? `Sent by you for ${b.name} — it goes on ${b.name.split(" ")[0]}'s list${state.cash ? " and any reimbursement is theirs" : ""}.`
                                            : "Leave as you unless you're sending this for someone else.";
  }
  function resetForm() {
    state.files.forEach((f) => f.previewUrl && URL.revokeObjectURL(f.previewUrl));
    state.files = []; state.expenseType = null; state.personal = false; state.cash = false; state.currency = null; state.finishing = null;
    el.reason.value = ""; el.attendees.value = ""; el.project.value = ""; el.nMuir.value = ""; el.nClients.value = ""; soloBox().checked = false;
    setSwitch(el.swPersonal, false); el.payCard.checked = true; el.payCash.checked = false; renderBelongs(null); toggleBelongsList(false);
    el.sendErr.classList.add("hidden");
    [el.inCamera, el.inGallery, el.inFile].forEach((i) => (i.value = ""));
    renderTypes(); renderCurrencies(); updateConditional(); updateReasonPreview();
  }
  function validate() {
    if (!state.files.length) return "Add at least one photo or PDF.";
    if (!state.expenseType) return "Pick an expense type.";
    if (el.reason.value.trim().length < 3) return "Say what it was for (e.g. \u201cTaxi to airport\u201d) — finance will chase you for it otherwise.";
    if (!state.currency) return "Pick the currency shown on the receipt.";
    if (!el.attendeesWrap.classList.contains("hidden") && !(Number(el.nMuir.value) + Number(el.nClients.value) > 0)) return "How many Muir employees and clients? (needed for FBT).";
    if (isSolo() && !soloBox().checked) return "Please confirm only 1 person present for meal (tick the box), or change the number of people.";
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
    const fin0 = state.finishing;
    const localDraft = !!(fin0 && fin0.local);       // saved on the phone only — nothing uploaded yet
    const fin = localDraft ? null : fin0;              // legacy (pre-0.4) drafts: photo already uploaded, send meta only
    const prefix = fin0 ? fin0.prefix : `${stamp(now)}_${local}`;
    const files = fin ? [] : state.files.map((f, i) => ({ name: `${prefix}_${i + 1}${f.kind === "pdf" ? ".pdf" : ".jpg"}`, type: f.type, blob: f.blob }));
    const fileNames = fin ? fin.files : files.map((f) => f.name);
    const meta = {
      schema: "muir-receipt/2", app: `muir-receipts-pwa ${APP_VERSION}`, status,
      submitted_by: { name: state.user.name, upn: state.user.upn },
      cardholder: (() => { const b = belongsTo(); return { name: b.name, upn: b.upn }; })(), on_behalf: belongsTo().onBehalf,
      captured_at: fin ? fin.meta.captured_at : localDraft ? fin0.captured_at : now.toISOString(), details_at: status === "complete" ? now.toISOString() : null,
      tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
      expense_type: state.expenseType, reason: fullReason() || null, reason_text: el.reason.value.trim() || null,
      currency_selected_by_user: !!state.currency,
      attendees: el.attendeesWrap.classList.contains("hidden") ? null : (el.attendees.value.trim() || null),
      n_muir_employees: el.attendeesWrap.classList.contains("hidden") ? null : Number(el.nMuir.value || 0),
      n_clients: el.attendeesWrap.classList.contains("hidden") ? null : Number(el.nClients.value || 0),
      solo_meal_confirmed: isSolo() ? soloBox().checked : null,
      project: el.project.value.trim() || null,
      currency: state.currency, payment_type: state.cash ? "Cash" : "Credit Card",
      to_be_reimbursed: state.personal, card: state.personal ? "personal" : "company",
      files: fileNames,
      scan: state.files.map((f) => (f.scan ? { guess: f.scan.method, guessed: !!f.scan.confident, cropped: !!f.scan.edited, sharpened: !!f.scan.edited && f.scan.enhance !== false } : null)),
    };
    return { id: prefix + (status === "complete" && fin ? "-details" : ""), prefix, meta, files, metaOnly: !!fin };
  }
  async function transmit(sub) {
    // "replace" so a retry after a half-finished upload overwrites rather than making duplicates (names are unique per receipt)
    for (const f of sub.files) await uploadFile(f.name, await freshBlob(f), "replace");
    const metaBlob = new Blob([JSON.stringify(sub.meta, null, 2)], { type: "application/json" });
    // Uploaded last so its presence marks the set complete; a details-later completion replaces the pending meta.
    await uploadFile(`${sub.prefix}_meta.json`, metaBlob, "replace");
  }
  // ---- "No time now": keep the photo ON THE PHONE (nothing is sent) until a reason is added ----
  async function sendLater() {
    if (!state.files.length) { el.sendErr.textContent = "Add at least one photo or PDF."; el.sendErr.classList.remove("hidden"); return; }
    const local = (state.user.upn.split("@")[0] || "user").replace(/[^a-z0-9._-]/gi, "").toLowerCase();
    const fin0 = state.finishing;
    const draft = {
      prefix: fin0 ? fin0.prefix : `${stamp()}_${local}`, local: true,
      captured_at: fin0 && fin0.captured_at ? fin0.captured_at : new Date().toISOString(),
      files: await Promise.all(state.files.map(async (f) => ({ type: f.type, bytes: await bytesOf(f.blob), kind: f.kind, orig: f.orig ? await bytesOf(f.orig) : null, scan: f.scan || null }))), pages: state.files.length,
      fields: { expenseType: state.expenseType, reason: el.reason.value, currency: state.currency, cash: state.cash, personal: state.personal,
                attendees: el.attendees.value, nMuir: el.nMuir.value, nClients: el.nClients.value, project: el.project.value, belongs: state.belongs || "", solo: soloBox().checked },
    };
    await draftPut(draft);
    askNotifyPermission();   // needed for the icon badge on iPhone; asked once, on this tap
    resetForm(); renderDrafts();
    showDone("Saved on your phone — not sent", "It's at the top of the home screen under \u201cNot sent yet\u201d, and the app icon shows a count until it's done. Add the reason and it sends.");
  }
  async function openDraft(prefix) {
    try { await openDraftInner(prefix); }
    catch (e) {
      console.error("openDraft", e);
      alert("Couldn't open this receipt (" + (e && e.message ? e.message : e) + "). If it keeps happening, Delete it and snap the receipt again.");
      show("capture");
    }
  }
  async function openDraftInner(prefix) {
    const d = (await draftsAll()).find((x) => x.prefix === prefix); if (!d) { renderDrafts(); return; }
    // iPhone sometimes drops photos saved inside the app; check the photo can still be read before opening
    const firstBlob = d.local ? (d.files && d.files[0] && (d.files[0].blob || d.files[0].bytes)) : d.thumb;
    if (firstBlob && firstBlob instanceof Blob) {
      try { await firstBlob.slice(0, 16).arrayBuffer(); }
      catch (_) { alert("The photo for this receipt was cleared from the phone, so it can't be opened. Please Delete it and snap the receipt again."); return; }
    }
    resetForm();
    if (d.local) {
      state.finishing = { prefix: d.prefix, captured_at: d.captured_at, local: true };
      state.files = d.files.map((f) => {
        const blob = asBlob(f.blob || f.bytes, f.type), orig = asBlob(f.orig, "image/jpeg");
        return { type: f.type, kind: f.kind, scan: f.scan || null, blob, orig, name: f.name, previewUrl: f.kind === "image" ? URL.createObjectURL(blob) : null };
      });
      const x = d.fields || {};
      state.expenseType = x.expenseType || null; state.cash = !!x.cash; state.personal = !!x.personal;
      el.reason.value = x.reason || ""; el.attendees.value = x.attendees || ""; el.nMuir.value = x.nMuir || ""; el.nClients.value = x.nClients || ""; el.project.value = x.project || "";
      soloBox().checked = !!x.solo;
      el.payCash.checked = state.cash; el.payCard.checked = !state.cash; setSwitch(el.swPersonal, state.personal);
      state.currency = x.currency || null;
      renderBelongs(x.belongs || null);
      renderThumbs(); renderTypes(); renderCurrencies(); updateConditional(); updateReasonPreview();
      el.btnSend.textContent = "Send to finance";
      show("details"); setTimeout(() => el.reason.focus(), 200);
      return;
    }
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
    const wasLocal = wasFinishing && !!state.finishing.local;
    localStorage.setItem("muir.lastType", state.expenseType);
    for (const f of sub.files) { if (!f.bytes) { f.bytes = await bytesOf(f.blob); delete f.blob; } }   // bytes, not Blob (iPhone bug)
    await queuePut(sub);              // saved on the phone before any network call
    inflight.add(sub.id);
    try {
      if (!navigator.onLine) throw new Error("offline");
      await transmit(sub);
      await queueDel(sub.id);
      addRecent(sub, "sent");
      state.needSignin = false; renderSigninBar();
      showDone(wasFinishing && !wasLocal ? "Details sent" : "Sent", `${sub.meta.expense_type} · ${sub.meta.reason}`);
    } catch (e) {
      addRecent(sub, "queued");
      const wf = navigator.onLine ? whyFailed(e) : { kind: "network", text: "You're offline — it will send when you're back online." };
      sub.lastError = wf; sub.lastTry = new Date().toISOString(); sub.tries = 1; await queuePut(sub).catch(() => {});
      const why = wf.text;
      showDone(wf.kind === "signin" ? "Saved — sign in to send" : "Saved — will send shortly", `${sub.meta.expense_type} · ${sub.meta.reason}. ${why}`);
      if (wf.kind === "signin") setTimeout(() => signInToSend(false), 1500);   // the receipt is safe on the phone; sends on return
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
    const at = (d) => d.captured_at || (d.meta && d.meta.captured_at) || "";
    const list = (await draftsAll().catch(() => [])).sort((p, q) => (at(p) < at(q) ? 1 : -1));
    el.tofinishCard.classList.toggle("hidden", list.length === 0);
    updateBadge(list.length);
    el.tofinish.innerHTML = "";
    list.forEach((d) => {
      const li = document.createElement("li"); li.style.cursor = "pointer";
      const when = new Date(at(d)).toLocaleString("en-AU", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
      const hrs = (Date.now() - new Date(at(d)).getTime()) / 36e5;
      const age = hrs >= 24 ? `${Math.floor(hrs / 24)} day${hrs >= 48 ? "s" : ""} old` : "";
      li.innerHTML = `<span class="grow">${when} · ${d.pages || (d.files || []).length || 1} page${(d.pages || (d.files || []).length) > 1 ? "s" : ""}${age ? ` · <strong style="color:var(--warn)">${age}</strong>` : ""}<br><span class="muted">${d.local ? "Not sent — tap to add the reason" : "Tap to add type and reason"}</span></span>`;
      const del = document.createElement("button"); del.type = "button"; del.className = "del"; del.textContent = "Delete";
      del.onclick = (e) => { e.stopPropagation(); deleteDraft(d, del); };
      const go = document.createElement("span"); go.textContent = "›";
      li.append(del, go);
      li.onclick = () => openDraft(d.prefix);
      el.tofinish.appendChild(li);
    });
  }
  // App-icon badge = receipts waiting for a reason (iPhone: home-screen app + notifications allowed; Android: installed app)
  function updateBadge(n) {
    try { if ("setAppBadge" in navigator) { n ? navigator.setAppBadge(n) : navigator.clearAppBadge(); } } catch (e) { /* unsupported */ }
  }
  function askNotifyPermission() {
    try { if ("Notification" in window && Notification.permission === "default") Notification.requestPermission(); } catch (e) { /* unsupported */ }
  }
  // On opening the app, a reminder if anything has been waiting more than 2 hours
  async function remindDrafts() {
    const list = (await draftsAll().catch(() => [])).filter((d) => d.local);
    const old = list.filter((d) => Date.now() - new Date(d.captured_at).getTime() > 2 * 36e5);
    if (!old.length || !("Notification" in window) || Notification.permission !== "granted") return;
    const last = Number(localStorage.getItem("muir.lastRemind") || 0);
    if (Date.now() - last < 6 * 36e5) return;
    localStorage.setItem("muir.lastRemind", String(Date.now()));
    const reg = await navigator.serviceWorker?.getRegistration?.();
    const body = `${old.length} receipt${old.length > 1 ? "s are" : " is"} saved on your phone but not sent — add the reason to send.`;
    try { reg ? reg.showNotification("Muir Receipts", { body, tag: "muir-drafts", icon: "icons/icon-192.png", badge: "icons/icon-192.png" }) : new Notification("Muir Receipts", { body }); } catch (e) { /* ignore */ }
  }

  let flushing = false;
  const inflight = new Set();
  async function flushQueue() {
    if (flushing || !navigator.onLine || !state.user) return;
    flushing = true;
    try {
      for (const sub of await queueAll()) {
        if (inflight.has(sub.id)) continue;
        try { await transmit(sub); await queueDel(sub.id); addRecent(sub, "sent"); state.needSignin = false; renderSigninBar(); }
        catch (e) {
          console.warn("Queue item still pending:", sub.id, e.message);
          sub.lastError = whyFailed(e); sub.lastTry = new Date().toISOString(); sub.tries = (sub.tries || 0) + 1;
          await queuePut(sub).catch(() => {});
          if (e.signin) { signInToSend(false); return; }
          if (e.signin || !navigator.onLine) break;     // same for every item: stop and wait
          // otherwise carry on with the next receipt — one bad item mustn't hold up the rest
        }
      }
    } finally { flushing = false; refreshQueueBanner(); }
  }
  // ---------- Sign-in only when it's needed, and never in the middle of a receipt ----------
  // The app opens straight to the camera using the account remembered on the phone. Talking to Microsoft happens in the
  // background. If Microsoft wants a fresh sign-in, we go there only when nothing is in progress (receipts are safe on
  // the phone) — or show a "Sign in to send" bar on the home screen.
  function busy() { return state.files.length > 0 || screens.details.classList.contains("active") || screens.scan.classList.contains("active") || screens.admin.classList.contains("active"); }
  function renderSigninBar() { const b = document.getElementById("signin-banner"); if (b) b.classList.toggle("hidden", !state.needSignin); }
  function signInToSend(userAsked) {
    state.needSignin = true; renderSigninBar();
    if (DEMO) return;
    if (!userAsked && (busy() || sessionStorage.getItem("muir.autoSignin"))) return;   // once per session automatically
    sessionStorage.setItem("muir.autoSignin", "1");
    msalInstance().acquireTokenRedirect({ scopes: CFG.scopes, account: state.account }).catch((e) => console.warn("sign-in redirect", e));
  }
  // Plain-English reason an upload failed, saved with the receipt and shown in the waiting list
  function whyFailed(e) {
    const m = String((e && e.message) || e || "");
    if (e && e.signin) return { kind: "signin", text: "Your Microsoft sign-in has expired — tap Sign in again." };
    if (/Graph 403|accessDenied/i.test(m)) return { kind: "access", text: "SharePoint won't accept it: you don't have permission to upload. Ask Max to check your access to Corporate Cards › Incoming." };
    if (/Graph 401/.test(m)) return { kind: "signin", text: "Microsoft didn't accept your sign-in — tap Sign in again." };
    if (/Graph 404/.test(m)) return { kind: "folder", text: "The receipts folder couldn't be found. Tell Max (the app may need an update)." };
    if ((e && e.unreadable) || /NotReadable|WebKitBlobResource|could not be read|NotFoundError/i.test(m)) return { kind: "photo", text: "The photo saved on this phone can't be read any more (the iPhone cleared it). Remove this one and snap the receipt again." };
    if (/Failed to fetch|NetworkError|Load failed|network|offline|timed? ?out/i.test(m)) return { kind: "network", text: (navigator.onLine ? "Couldn't send (" + m.slice(0, 60) + ") — tap Send now to try again." : "No connection — it will send when you're back online.") };
    return { kind: "other", text: "Upload failed: " + m.slice(0, 120) };
  }
  async function refreshQueueBanner() {
    const items = await queueAll().catch(() => []);
    const n = items.length;
    el.queueBanner.classList.toggle("hidden", n === 0);
    el.queueBanner.innerHTML = "";
    if (n) {
      const needSignin = items.some((x) => x.lastError && x.lastError.kind === "signin");
      const msg = document.createElement("div");
      msg.textContent = `${n} receipt${n > 1 ? "s" : ""} waiting to send` + (navigator.onLine ? (items.some((x) => x.lastError) ? " — see why below." : " — sending…") : " — will go automatically when you're online.");
      const row = document.createElement("div"); row.style.cssText = "display:flex;gap:8px;margin-top:8px;flex-wrap:wrap";
      const go = document.createElement("button"); go.type = "button"; go.className = "del"; go.style.color = "var(--ink)"; go.textContent = "Send now";
      go.onclick = async () => { go.disabled = true; go.textContent = "Sending…"; await flushQueue(); };
      row.appendChild(go);
      if (needSignin && !DEMO) {
        const si = document.createElement("button"); si.type = "button"; si.className = "del"; si.style.color = "var(--ink)"; si.textContent = "Sign in again";
        si.onclick = () => msalInstance().acquireTokenRedirect({ scopes: CFG.scopes, account: state.account });
        row.appendChild(si);
      }
      el.queueBanner.append(msg, row);
    }
    const ql = $("queue-list"); ql.innerHTML = ""; ql.classList.toggle("hidden", n === 0);
    items.forEach((sub) => {
      const li = document.createElement("li");
      const when = new Date(sub.meta.captured_at).toLocaleString("en-AU", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
      const err = sub.lastError ? `<br><span style="color:var(--warn);font-size:13px">Not sent: ${escapeHtml(sub.lastError.text)}</span>` : "";
      li.innerHTML = `<span class="grow">${when} · ${escapeHtml(sub.meta.expense_type || "Photo")}<br><span class="muted">${escapeHtml(sub.meta.reason || "")}</span>${err}</span>`;
      const del = document.createElement("button"); del.type = "button"; del.className = "del"; del.textContent = "Remove";
      del.onclick = async () => {
        if (!confirm("Remove this receipt? It hasn't been sent yet, so finance will never see it.")) return;
        if (inflight.has(sub.id)) { alert("It's sending right now — wait a moment, then use Undo under Recent if needed."); return; }
        await queueDel(sub.id); removeRecent(sub.id); refreshQueueBanner();
      };
      li.appendChild(del); ql.appendChild(li);
    });
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
    const item = { id: sub.id, prefix: sub.prefix, files: [`${sub.prefix}_meta.json`].concat(sub.meta.files || []),   // meta first: the review ignores a set without it
                   at: sub.meta.captured_at, type: sub.meta.expense_type, reason: sub.meta.reason, status };
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
      li.innerHTML = `<span class="grow">${when} · ${label}</span><span>${x.status === "sent" ? "✓" : x.status === "removed" ? "✕" : "⏳"}</span>`;
      if (x.status === "sent" && x.files && x.files.length) {
        const del = document.createElement("button"); del.type = "button"; del.className = "del"; del.textContent = "Undo";
        del.title = "Sent by mistake? Take it back before finance picks it up";
        del.onclick = () => undoSent(x, del);
        li.insertBefore(del, li.lastChild);
      }
      el.recent.appendChild(li);
    });
  }
  function removeRecent(id) {
    const list = JSON.parse(localStorage.getItem("muir.recent") || "[]").filter((x) => x.id !== id);
    localStorage.setItem("muir.recent", JSON.stringify(list)); renderRecent();
  }
  function markRecent(id, status) {
    const list = JSON.parse(localStorage.getItem("muir.recent") || "[]");
    list.forEach((x) => { if (x.id === id) x.status = status; });
    localStorage.setItem("muir.recent", JSON.stringify(list)); renderRecent();
  }
  // Delete files this phone uploaded to Incoming, if finance hasn't picked them up yet.
  // Returns { deleted, gone } — gone = files no longer in Incoming (already filed by the receipts review).
  async function deleteFromIncoming(names) {
    let deleted = 0, gone = 0;
    for (const n of names) {
      if (DEMO) { await sleep(150); deleted++; continue; }
      try { await graph(`/drives/${CFG.driveId}/items/${CFG.incomingFolderId}:/${encodeURIComponent(n)}:`, { method: "DELETE" }); deleted++; }
      catch (e) { if (/Graph 404/.test(e.message)) gone++; else throw e; }
    }
    return { deleted, gone };
  }
  const PICKED_UP = "Finance has already picked this one up, so it can't be taken back from here. Email receipts@muir.com.au saying which receipt to remove and why.";
  async function undoSent(x, btn) {
    if (!confirm("Take this receipt back? It will be removed before finance processes it.")) return;
    btn.disabled = true;
    try {
      const r = await deleteFromIncoming(x.files);
      if (r.deleted === 0 && r.gone > 0) { alert(PICKED_UP); btn.disabled = false; return; }
      markRecent(x.id, "removed");
    } catch (e) {
      btn.disabled = false;
      alert(/Graph 403/.test(e.message) ? "SharePoint won't let you remove it. Email receipts@muir.com.au saying which receipt to remove." : "Couldn't remove it: " + e.message);
    }
  }
  async function deleteDraft(d, btn) {
    const msg = d.local ? "Delete this receipt from your phone? It hasn't been sent, so finance will never see it."
                        : "Delete this receipt? The photo was already uploaded, so it will be taken back from finance too.";
    if (!confirm(msg)) return;
    btn.disabled = true;
    try {
      if (!d.local && d.files && d.files.length) {
        const r = await deleteFromIncoming([`${d.prefix}_meta.json`].concat(d.files));
        if (r.deleted === 0 && r.gone > 0) { alert(PICKED_UP); btn.disabled = false; return; }
      }
      await draftDel(d.prefix); renderDrafts();
    } catch (e) { btn.disabled = false; alert("Couldn't delete it: " + e.message); }
  }
  // Home from anywhere (header logo). Asks first if a new receipt is half-done.
  function goHome() {
    if (!state.user) return;
    const onDetails = screens.details.classList.contains("active");
    if (onDetails && state.files.length && !state.finishing && !confirm("Leave this receipt? It hasn't been sent or saved. (Use “No time now” to keep it.)")) return;
    if (onDetails) { resetForm(); el.btnLater.classList.remove("hidden"); el.btnSend.textContent = "Send to finance"; }
    show("capture"); renderDrafts(); refreshQueueBanner(); renderRecent();
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
  // Paid with: Muir credit card (default) or Cash. Cash was paid personally, so it switches "To be reimbursed" on (can be turned off).
  el.payCard.onchange = () => { if (el.payCard.checked) { state.cash = false; } updateBelongsHint(); };
  el.payCash.onchange = () => { if (el.payCash.checked) { state.cash = true; state.personal = true; setSwitch(el.swPersonal, true); } updateBelongsHint(); };
  el.belongs.onclick = () => toggleBelongsList();
  $("btn-admin").onclick = (e) => { e.preventDefault(); openAdmin(); };
  $("btn-admin-back").onclick = () => show("capture");
  $("btn-admin-add").onclick = () => openAdminEdit(null);
  $("btn-admin-save-admins").onclick = saveAdmins;
  $("btn-ad-cancel").onclick = renderAdmin;
  $("btn-ad-save").onclick = saveAdminEdit;
  $("ad-active").onclick = () => { ad.active = !ad.active; setSwitch($("ad-active"), ad.active); };
  ["ad-name", "ad-email", "ad-last4", "ad-other"].forEach((id) => $(id).addEventListener("input", () => adminErr("ad-err", "")));
  $("ad-last4").oninput = (e) => { e.target.value = e.target.value.replace(/\D/g, "").slice(0, 4); };
  el.btnBack.onclick = () => { if (state.finishing) { resetForm(); el.btnLater.classList.remove("hidden"); el.btnSend.textContent = "Send to finance"; } show("capture"); };
  el.btnSend.onclick = send;
  el.btnLater.onclick = sendLater;
  el.reason.addEventListener("input", updateReasonPreview);
  wireScan();
  $("btn-signin-send").onclick = () => signInToSend(true);
  [el.nMuir, el.nClients].forEach((i) => i.addEventListener("input", updateSolo));
  soloBox().addEventListener("change", () => el.sendErr.classList.add("hidden"));
  $("btn-home").onclick = goHome;
  $("btn-home2").onclick = goHome;
  $("btn-done-home").onclick = goHome;
  el.btnAnother.onclick = () => { show("capture"); setTimeout(() => el.inCamera.click(), 50); };
  window.addEventListener("online", setNet); window.addEventListener("offline", setNet);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) flushQueue(); });

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((e) => console.warn("SW registration failed", e));
  }

  // ---------- Start-up: loading screen, time limits, Refresh ----------
  // Each network step gets a time limit, so a slow or flaky connection can't leave the app stuck on a blank screen.
  const withTimeout = (p, ms, fallback) => Promise.race([p, new Promise((res) => setTimeout(() => res(fallback), ms))]);
  const step = (t) => { const x = $("loading-step"); if (x) x.textContent = t; };
  let started = false;
  const slowTimer = setTimeout(() => { if (!started && !window.__muirStarted) $("loading-slow").classList.remove("hidden"); }, 8000);
  async function resetApp() {
    if (!confirm("Reset the app? This clears the saved copy of the app and reloads it. Receipts not sent yet stay on the phone.")) return;
    try { const regs = (await navigator.serviceWorker?.getRegistrations?.()) || []; for (const r of regs) await r.unregister(); } catch (_) {}
    try { for (const k of await caches.keys()) await caches.delete(k); } catch (_) {}
    location.reload();
  }
  $("btn-reload").onclick = () => location.reload();
  $("btn-refresh").onclick = (e) => { e.preventDefault(); location.reload(); };
  $("btn-reset").onclick = (e) => { e.preventDefault(); resetApp(); };
  let skipWait = null;
  const skipped = new Promise((res) => { skipWait = res; });
  $("btn-skipwait").onclick = () => skipWait("skip");

  async function boot() {
    step("Starting…");
    // 1. Who is this? Read from the phone (no network unless we're coming back from the Microsoft sign-in page).
    const ok = await Promise.race([withTimeout(initAuth(), 12000, "slow"), skipped]);
    if (ok === "slow" || ok === "skip") {
      if (DEMO) state.user = { name: "Demo User", upn: "demo@muir.com.au" };
      else try { const m = msalInstance(); const a = m.getActiveAccount() || m.getAllAccounts()[0]; if (a) { state.account = a; state.user = { name: String(a.name || a.username).split("|")[0].trim(), upn: a.username }; } } catch (_) {}
      if (!state.user) { show("signin"); return; }
    } else if (!ok) { show("signin"); return; }
    sessionStorage.removeItem("muir.autoSignin");
    // 2. Open straight away with what's saved on the phone (lists, cardholders, unsent receipts).
    el.who.textContent = state.user.name;
    try { applyCardholders(JSON.parse(localStorage.getItem("muir.cardholders") || "null") || builtInDoc()); } catch (_) { applyCardholders(builtInDoc()); }
    renderTypes(); renderCurrencies(); updateConditional(); renderRecent(); setNet(); refreshQueueBanner(); renderDrafts(); remindDrafts(); renderSigninBar();
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") { renderDrafts(); remindDrafts(); flushQueue(); } });
    started = true; clearTimeout(slowTimer);
    show("capture");
    // 3. In the background: refresh the Microsoft sign-in, update the lists, send anything waiting.
    background();
  }
  async function background() {
    if (!DEMO) {
      const t = await withTimeout(token().then(() => "ok", (e) => (e.signin ? "signin" : "net")), 12000, "slow");
      if (t === "signin") {
        state.needSignin = true; renderSigninBar();
        // go to Microsoft now only if receipts are already waiting and nothing is in progress; otherwise when they next send
        if ((await queueAll().catch(() => [])).length && !busy()) signInToSend(false);
        return;
      }
      if (t === "ok") { state.needSignin = false; renderSigninBar(); }
    }
    await withTimeout(Promise.all([loadRemoteConfig(), loadCardholders()]), 10000);
    if (!busy()) { renderTypes(); renderCurrencies(); }
    flushQueue();
    checkForUpdate();
  }
  // Is a newer version on GitHub? (The phone's copy opens instantly; this tells people when to tap update.)
  async function checkForUpdate() {
    try {
      const html = await (await fetch("./index.html?nosw=" + Date.now(), { cache: "no-store" })).text();
      const m = html.match(/app\.js\?v=([0-9.]+)/);
      if (m && m[1] !== APP_VERSION) {
        await fetch("./index.html", { cache: "no-cache" }).catch(() => {});   // let the service worker store the new page
        document.getElementById("update-bar").classList.remove("hidden");
      }
    } catch (_) { /* offline */ }
  }
  boot().catch((e) => {
    console.error("Start-up failed", e);
    step("Something went wrong starting up (" + (e && e.message ? e.message : e) + ").");
    $("loading-slow").classList.remove("hidden");
  });

  // Exposed for tests / console
  window.__muir = { state, queuePut, loadCardholders, openAdmin, normaliseImage, buildSubmission, queueAll, flushQueue, draftsAll, APP_VERSION };
})();
