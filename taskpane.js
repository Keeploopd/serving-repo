const msalConfig = {
  auth: {
    clientId: "705cf97b-720b-4240-b6d0-02a6655300b2",
    authority: "https://login.microsoftonline.com/5904ae0b-47e9-4b06-843e-60769342a32b",
    redirectUri: "https://co-draft.keeploopd.com/taskpane.html"
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const apiRequest = {
  scopes: ["api://co-draft.keeploopd.com/705cf97b-720b-4240-b6d0-02a6655300b2/access_as_user"]
};

const API_BASE = "https://api.keeploopd.com";

// Gate noisy logs that can contain email content / thread analysis (PII).
// REMINDER: set to false before shipping — currently true for dev.
const DEBUG = true;
function dlog(...args) { if (DEBUG) console.log(...args); }

let currentToken = null;
let monitoringStarted = false;

async function hashEmail(email) {
  const encoded = new TextEncoder().encode(email.toLowerCase().trim());
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

// Best-effort client-side expiry check (with 60s clock-skew margin) so we
// never send a token we already know is dead. Signature validation stays
// server-side, obviously.
function tokenExpired(token) {
  try {
    const payloadB64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const { exp } = JSON.parse(atob(payloadB64));
    return !exp || Date.now() / 1000 > exp - 60;
  } catch (e) {
    return true;
  }
}

// ── Cross-runtime token bridge ───────────────────────────────────────────────
// The Sync Co-Drafters button and the onNewMessageCompose event run in a
// separate UI-less runtime (commands.html) where MSAL's silent flows are
// unreliable; localStorage is the channel Outlook shares between the two
// runtimes, so that runtime falls back to this bridge when Office SSO and
// ssoSilent both fail. Accepted trade-off, mitigated: only unexpired tokens
// are ever written, the reader deletes expired ones on sight, so at-rest
// exposure is bounded by the token's own ~1h lifetime (vs. the previous
// version, which stored tokens indefinitely with no expiry check). Delete
// this once the add-in moves to a SharedRuntime or Nested App Auth.
const BRIDGE_KEY = "keeploopd_token";

function storeBridgeToken(token) {
  try {
    if (token && !tokenExpired(token)) {
      localStorage.setItem(BRIDGE_KEY, token);
    }
  } catch (e) {}
}

// Timebox any auth call. In classic Outlook desktop, MSAL's hidden-iframe
// flows (ssoSilent / acquireTokenSilent) can hang for 60s+ inside the
// embedded webview before failing — which is what made the taskpane appear
// to "take forever to load". Nothing silent is allowed to block startup for
// longer than these budgets; on timeout we fall through to the next source
// or show the sign-in button.
function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms)
    )
  ]);
}

// Office SSO without any prompts — native and fast in classic Outlook, so it
// runs FIRST at startup; MSAL is the fallback (it shines in OWA).
async function tryOfficeSsoQuiet() {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.auth) return null;
    const token = await withTimeout(
      OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: false,
        allowConsentPrompt: false
      }),
      6000,
      "Office SSO"
    );
    return token || null;
  } catch (e) {
    console.warn("Quiet Office SSO unavailable:", e);
    return null;
  }
}

async function signInAndCallBackend() {

  Office.context.ui.displayDialogAsync(
    "https://co-draft.keeploopd.com/auth.html",
    { height: 60, width: 40 },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Dialog failed to open", result.error);
        document.getElementById("status").textContent = "Could not open authentication dialog.";
        return;
      }

      const dialog = result.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        try {
          const payload = JSON.parse(arg.message);
          dialog.close();

          if (!payload.success) {
            console.error("Auth failed", payload);
            document.getElementById("status").textContent =
              `Authentication failed: ${payload.errorCode || ""} ${payload.errorMessage || ""}`;
            return;
          }

          const response = await fetch(`${API_BASE}/api/auth/test`, {
            headers: {
              Authorization: `Bearer ${payload.accessToken}`
            }
          });

          dlog("Backend auth test status", response.status);

          if (!response.ok) {
            document.getElementById("status").textContent =
              `Backend auth failed: ${response.status}`;
            return;
          }

          // Token is held in memory for this runtime; the expiry-checked
          // bridge write below exists solely for the command runtime
          // (Sync Co-Drafters / compose event) — see storeBridgeToken().
          currentToken = payload.accessToken;
          storeBridgeToken(currentToken);
          document.getElementById("signin").style.display = "none";
          document.getElementById("status").textContent = "Authenticated successfully";
          await init();
          await loadMissionControl();

        } catch (err) {
          console.error("Taskpane token handling error", err);
          document.getElementById("status").textContent =
            `Taskpane error: ${err.message || err}`;
        }
      });
    }
  );
}

let activeConversationId = null;
let heartbeatInterval = null;
let lastNotificationCount = null;
let missionAnalysisInProgress = false;

function stopMonitoring() {
  if (heartbeatInterval) clearInterval(heartbeatInterval);

  heartbeatInterval = null;
  monitoringStarted = false;
  lastNotificationCount = null;

  try {
    Office.context.mailbox.item.notificationMessages.removeAsync("codraftStatus");
  } catch (err) {
    console.warn("Could not remove notification", err);
  }
}

function updateNotification(count) {
  if (count === lastNotificationCount) return;

  lastNotificationCount = count;

  Office.context.mailbox.item.notificationMessages.removeAsync("codraftStatus", () => {
    if (count >= 1) {
      Office.context.mailbox.item.notificationMessages.addAsync("codraftStatus", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `${count} ${count > 1 ? "people are" : "person is"} currently drafting a reply`,
        icon: "Icon.16x16",
        persistent: false
      });
    }
  });
}

// Single place that reacts to a drafter count, fed by heartbeat responses.
// (Previously a separate /active-drafters poll ran every 5s in parallel with
// the heartbeat — the count now rides back on the heartbeat itself.)
function applyDrafterCount(count) {
  const banner = document.getElementById("banner");

  if (count >= 1) {
    banner.style.display = "block";
    banner.textContent = `${count} active drafter(s) detected`;
  } else {
    banner.style.display = "none";
    banner.textContent = "";
  }

  updateNotification(count);
}

// Convo ID testing begin
function getSubjectText(item) {
  if (typeof item.subject === "string") {
    return item.subject;
  }
  return "";
}

async function getConversationKey() {
  const item = Office.context.mailbox.item;

  const rawConversationId = item.conversationId || "";

  const subjectText = getSubjectText(item);
  const subject = subjectText
    .toLowerCase()
    .replace(/^(re|fw|fwd):\s*/i, "")
    .trim();

  const markerIndex = rawConversationId.lastIndexOf("AQ");
  const sharedConversationPart =
    markerIndex >= 0 ? rawConversationId.slice(markerIndex) : rawConversationId.slice(-24);

  const keyMaterial = `${sharedConversationPart}|${subject}`;

  return await sha256(keyMaterial);
}

async function sha256(value) {
  const encoded = new TextEncoder().encode(value);
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  return Array.from(new Uint8Array(hashBuffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}
// Convo ID testing end

async function init() {
  if (!currentToken) {
    document.getElementById("status").textContent = "Please authenticate first.";
    return;
  }

  const conversationId = await getConversationKey();

  if (activeConversationId && activeConversationId !== conversationId) {
    stopMonitoring();
  }

  if (monitoringStarted && activeConversationId === conversationId) {
    return;
  }

  activeConversationId = conversationId;
  monitoringStarted = true;

  document.getElementById("status").textContent = "Monitoring active drafters...";

  async function sendHeartbeat() {
    try {
      const token = await getValidToken();
      if (!token) return;

      // NOTE: no userId in the body — the backend derives identity from the
      // validated bearer token, so presence can't be spoofed client-side.
      const hb = await fetch(`${API_BASE}/heartbeat`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${token}`
        },
        body: JSON.stringify({ conversationId })
      });

      if (!hb.ok) {
        console.error("heartbeat failed", hb.status);
        return;
      }

      const data = await hb.json();
      if (typeof data.count === "number") {
        applyDrafterCount(data.count);
      }
    } catch (err) {
      console.error("Heartbeat failed:", err);
    }
  }

  await sendHeartbeat();

  heartbeatInterval = setInterval(sendHeartbeat, 10000);
}


async function trySilentAuth() {
  try {
    await withTimeout(msalInstance.initialize(), 5000, "MSAL initialize");

    try {
      const silentResult = await withTimeout(
        msalInstance.ssoSilent({
          scopes: apiRequest.scopes,
          loginHint: Office.context.mailbox.userProfile.emailAddress
        }),
        8000,
        "ssoSilent"
      );

      return silentResult.accessToken;
    } catch (ssoErr) {
      console.warn("ssoSilent failed", ssoErr);
    }

    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) return null;

    const result = await withTimeout(
      msalInstance.acquireTokenSilent({
        ...apiRequest,
        account: accounts[0]
      }),
      8000,
      "acquireTokenSilent"
    );

    return result.accessToken;
  } catch (err) {
    console.warn("Silent auth unavailable", err);
    return null;
  }
}

Office.onReady(async () => {
  const signinButton = document.getElementById("signin");

  if (!signinButton) {
    console.error("signin button not found");
    return;
  }

  signinButton.addEventListener("click", signInAndCallBackend);

  const refreshButton = document.getElementById("refresh-analysis");
  if (refreshButton) {
    refreshButton.addEventListener("click", refreshAnalysis);
  }

  // Immediate feedback so the pane never looks frozen while auth resolves.
  document.getElementById("status").textContent = "Checking sign-in status...";

  // Classic Outlook fix: Office SSO first (native, fast in the desktop
  // client), MSAL second (best in OWA). Every source is timeboxed, so the
  // worst-case wait before showing the sign-in button is a few seconds
  // instead of the webview hanging on MSAL's hidden iframe.
  let silentToken = await tryOfficeSsoQuiet();
  if (!silentToken) {
    silentToken = await trySilentAuth();
  }

  if (silentToken) {
    currentToken = silentToken;
    storeBridgeToken(currentToken);

    signinButton.style.display = "none";

    await init();
    await loadMissionControl();
  } else {
    document.getElementById("status").textContent = "Sign in required.";
    signinButton.style.display = "block";
  }

  window.addEventListener("beforeunload", () => {
    try {
      Office.context.mailbox.item.notificationMessages.removeAsync("codraftStatus");
    } catch (err) {
      console.warn("Could not clear Co-Draft notification on unload", err);
    }
  });
});

async function getValidToken() {
  if (currentToken && !tokenExpired(currentToken)) return currentToken;

  // Refresh order mirrors startup: quiet Office SSO (fast, esp. classic
  // desktop) → MSAL silent (sessionStorage cache / refresh tokens) → Office
  // SSO with prompts allowed as the last resort.
  currentToken = await tryOfficeSsoQuiet();
  if (!currentToken) {
    currentToken = await trySilentAuth();
  }
  if (!currentToken) {
    try {
      currentToken = await getAuthToken();
    } catch (e) {
      console.warn("Office SSO fallback failed", e);
      currentToken = null;
    }
  }
  storeBridgeToken(currentToken);
  return currentToken;
}

async function getConversationContext() {
  const item = Office.context.mailbox.item;

  dlog("itemId:", item.itemId);
  dlog("conversationId:", item.conversationId);

  return {
    conversationId: await getConversationKey(),
    subject: getSubjectText(item)
  };
}

async function getAuthToken() {
  return await OfficeRuntime.auth.getAccessToken({
    allowSignInPrompt: true,
    allowConsentPrompt: true
  });
}



function getMyDomain() {
  return Office.context.mailbox.userProfile.emailAddress.split("@")[1]?.toLowerCase();
}

// Session-local knowledge: hash -> is_internal for every address THIS client
// has actually seen. Hashes are one-way, so this map (plus the display-name
// fallback in resolveIsInternal) is the only way to correct stale stored
// flags for participants visible in the current item.
const localInternalByHash = new Map();

async function buildParticipantPayload(recipients) {
  const myDomain = getMyDomain();
  dlog("myDomain:", myDomain, "recipients:", recipients);
  return Promise.all(
    recipients.filter(r => r.emailAddress).map(async r => {
      const participant_hash = await hashEmail(r.emailAddress);
      const is_internal = r.emailAddress.split("@")[1]?.toLowerCase() === myDomain;
      localInternalByHash.set(participant_hash, is_internal);
      return {
        participant_hash,
        display_name: r.displayName || r.emailAddress,
        is_internal
      };
    })
  );
}

// Best available truth for a server-returned participant, in order:
//   1. We saw their address this session (authoritative for this client)
//   2. Their stored display_name IS an email — derive from its domain
//   3. Whatever the server stored (may be stale FALSE for rows written
//      before the is_internal feature, until that person next appears in
//      an updated client's from/to/cc)
function resolveIsInternal(p) {
  if (localInternalByHash.has(p.participant_hash)) {
    return localInternalByHash.get(p.participant_hash);
  }
  const name = p.display_name || "";
  if (name.includes("@")) {
    return name.split("@")[1]?.toLowerCase().trim() === getMyDomain();
  }
  return !!p.is_internal;
}


async function loadMissionControl() {
  try {
    setMissionStatus("Loading Mission Control...", "amber");

    const context = await getConversationContext();
    const token = await getValidToken();

    const url = new URL("/api/thread/state", API_BASE);
    url.searchParams.set("conversationId", context.conversationId);

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) throw new Error(`Mission Control request failed: ${response.status}`);

    const data = await response.json();
    dlog("THREAD_STATE response:", data);

    const noParticipantsYet = !data.participants || data.participants.length === 0;

    if ((data.status === "missing" || noParticipantsYet) && !missionAnalysisInProgress) {
      await refreshAnalysis();
      return;
    }

    // POST /api/thread/participants now returns the refreshed participant
    // list, so the second GET /api/thread/state round trip is gone.
    const updatedParticipants = await updateParticipantsOnly(context.conversationId, token);

    if (updatedParticipants) {
      data.participants = updatedParticipants;
    }

    renderMissionControl(data);
    setMissionStatus("Mission Control ready", "green");

  } catch (err) {
    console.error("Mission Control load failed:", err);
    setMissionStatus("Mission Control unavailable", "red");
  }
}


async function updateParticipantsOnly(conversationId, token) {
  try {
    const recipients = await getRecipients();
    const participants = await buildParticipantPayload(recipients);

    const res = await fetch(`${API_BASE}/api/thread/participants`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`
      },
      body: JSON.stringify({ conversationId, participants })
    });

    if (!res.ok) return null;

    const data = await res.json();
    return data.participants || null;
  } catch (err) {
    console.warn("Background participant update failed:", err);
    return null;
  }
}

async function getRecipients() {
  const item = Office.context.mailbox.item;

  const getAsync = (prop) => new Promise(resolve => {
    prop.getAsync(r => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null);
    });
  });

  const from = await getAsync(item.from);
  const toList = await getAsync(item.to);
  const ccList = await getAsync(item.cc);

  dlog("from:", from, "to:", toList, "cc:", ccList);

  const recipients = [];

  if (from?.emailAddress) {
    recipients.push({
      emailAddress: from.emailAddress,
      displayName: from.displayName || from.emailAddress
    });
  }

  const toArr = Array.isArray(toList) ? toList : toList ? [toList] : [];
  const ccArr = Array.isArray(ccList) ? ccList : ccList ? [ccList] : [];

  [...toArr, ...ccArr]
    .filter(r => r?.emailAddress)
    .forEach(r => recipients.push({
      emailAddress: r.emailAddress,
      displayName: r.displayName || r.emailAddress
    }));

  dlog("getRecipients result:", recipients);
  return recipients;
}


async function refreshAnalysis() {
  if (missionAnalysisInProgress) return;
  missionAnalysisInProgress = true;

  try {
    setMissionStatus("Refreshing analysis...", "amber");

    const context = await getConversationContext();
    const recipients = await getRecipients();
    const participants = await buildParticipantPayload(recipients);
    const token = await getValidToken();
    const threadText = await getCurrentEmailText();

    const response = await fetch(`${API_BASE}/api/thread/analyse`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`
      },
      body: JSON.stringify({
        ...context,
        threadText,
        participants
      })
    });

    if (!response.ok) throw new Error("Analysis failed");

    const data = await response.json();
    renderMissionControl(data);
    setMissionStatus("Mission Control ready", "green");

  } catch (err) {
    console.error("Mission analysis failed:", err);
    setMissionStatus("Mission analysis failed", "red");
  } finally {
    missionAnalysisInProgress = false;
  }
}

async function getCurrentEmailText() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      }
    );
  });
}

async function pollUntilReady(conversationId, token) {
  const interval = setInterval(async () => {
    const url = new URL("/api/thread/state", API_BASE);
    url.searchParams.set("conversationId", conversationId);
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    const data = await response.json();

    dlog("pollUntilReady status:", data.status);

    renderMissionControl(data);

    if (data.status === "ready") {
      clearInterval(interval);
    }
  }, 3000);
}

function renderMissionControl(data) {
  const state = data.state || {};

  dlog("renderMissionControl data:", data);

  try { renderParticipants(data.participants || []); }
  catch(e) { console.error("renderParticipants failed:", e); }

  try { renderMissionItems(state.missionControl || []); }
  catch(e) { console.error("renderMissionItems failed:", e); }

  try { renderQuestions(state.openQuestions || []); }
  catch(e) { console.error("renderQuestions failed:", e); }

  try { renderReplyFocus(state.suggestedReplyFocus || {}); }
  catch(e) { console.error("renderReplyFocus failed:", e); }

  try { renderThreadHealth(state.threadHealth || {}); }
  catch(e) { console.error("renderThreadHealth failed:", e); }
}

function renderMissionItems(items) {
  const el = document.getElementById("mission-control-list");
  el.innerHTML = "";

  if (!items.length) {
    el.innerHTML = `<div class="empty-state">No mission items yet.</div>`;
    return;
  }

  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "task-row";

    row.innerHTML = `
      <div class="task-left">
        <div class="task-icon-wrap">•</div>
        <div class="task-text">${escapeHtml(item.text || item)}</div>
      </div>
      <div class="task-tag tag-blue">${escapeHtml(item.type || "item")}</div>
    `;

    el.appendChild(row);
  });
}

function renderQuestions(questions) {
  const el = document.getElementById("open-questions-list");
  el.innerHTML = "";

  const badge = document.getElementById("questions-badge");
  if (badge) badge.textContent = questions.length ? String(questions.length) : "—";

  if (!questions.length) {
    el.innerHTML = `<div class="empty-state">No open questions detected.</div>`;
    return;
  }

  questions.forEach((q) => {
    const item = document.createElement("div");
    item.className = "question-item";
    item.innerHTML = `
      <div class="question-bullet"></div>
      <div>${escapeHtml(q)}</div>
    `;
    el.appendChild(item);
  });
}

function renderReplyFocus(focus) {
  document.getElementById("reply-focus-reason").innerText =
    focus.reason || "No reply focus detected yet.";

  const el = document.getElementById("reply-targets-list");
  el.innerHTML = "";

  const recipients = focus.primaryRecipients || [];

  recipients.forEach((person) => {
    const div = document.createElement("div");
    div.className = "reply-target";

    // getInitials output is escaped AND the function itself is hardened —
    // this was previously an unescaped innerHTML sink fed by AI-influenced
    // recipient strings.
    div.innerHTML = `
      <div class="reply-avatar" style="background:#0f6cbd;">
        ${escapeHtml(getInitials(person))}
      </div>
      <div>
        <div class="reply-target-name">${escapeHtml(person)}</div>
        <div class="reply-target-role">Suggested recipient</div>
      </div>
    `;

    el.appendChild(div);
  });
}

function renderThreadHealth(health) {
  document.getElementById("health-unresolved").innerText =
    health.unresolvedCount ?? "0";

  document.getElementById("health-decisions").innerText =
    health.decisionsCount ?? "0";

  document.getElementById("health-blockers").innerText =
    health.blockersCount ?? "0";

  document.getElementById("health-messages").innerText =
    health.messageCount ?? "—";

  document.getElementById("health-unresolved-status").innerText =
    health.urgency || "Normal";
}

function formatParticipantName(name) {
  if (!name) return "";

  if (name.includes("@")) {
    const localPart = name.split("@")[0];
    const parts = localPart.split(/[._-]/).filter(Boolean);
    if (parts.length === 1) return parts[0];
    return `${parts[0]} ${parts[1][0].toUpperCase()}.`;
  }

  const parts = name.trim().split(" ").filter(Boolean);
  if (parts.length === 1) return parts[0];
  return `${parts[0]} ${parts[1][0].toUpperCase()}.`;
}

function renderParticipants(participants) {
  const el = document.getElementById("participants-list");
  el.innerHTML = "";

  // Overlay the best locally available is_internal before counting — server
  // values can be stale for rows written before the feature existed.
  participants = (participants || []).map(p => ({
    ...p,
    is_internal: resolveIsInternal(p)
  }));

  const badge = document.getElementById("participants-badge");
  if (badge) {
    const internal = participants.filter(p => p.is_internal).length;
    const external = participants.filter(p => !p.is_internal).length;
    const internalEl = document.getElementById("internal-count");
    const externalEl = document.getElementById("external-count");
    if (internalEl && externalEl) {
      internalEl.textContent = participants.length ? internal : "—";
      externalEl.textContent = participants.length ? external : "—";
    }
  }

  if (!participants?.length) {
    el.innerHTML = `<div class="empty-state">No participants yet.</div>`;
    return;
  }

  const statusConfig = {
    active: { label: "Active", className: "active" },
    dropped: { label: "Dropped", className: "dropped" },
    additional: { label: "Additional", className: "additional" },
    shared_out_of_thread: { label: "Shared out-of-thread", className: "shared" }
  };

  participants.forEach((p) => {
    const name = p.display_name || p.email || "Unknown";
    const config = statusConfig[p.display_status] || statusConfig.active;
    const isEmail = name.includes("@");

    const div = document.createElement("div");
    div.className = "participant";

    if (isEmail) {
      div.title = name;
    }

    div.innerHTML = `
      <div class="avatar-wrap">
        <div class="avatar" style="background:#0f6cbd;">
          ${escapeHtml(getInitials(name))}
        </div>
        <div class="avatar-status ${config.className}"></div>
      </div>
      <div class="participant-name">${escapeHtml(formatParticipantName(name))}</div>
      <div class="participant-role">${config.label}</div>
    `;

    el.appendChild(div);
  });
}

function getInitials(value) {
  // Hardened: previously returned raw first characters (e.g. "<" from a
  // display name like "<img ...") straight into innerHTML. Now only
  // alphanumeric initials can come out; anything else collapses to "?".
  const initials = String(value || "")
    .split(/[ .@_-]/)
    .filter(Boolean)
    .slice(0, 2)
    .map(x => x[0].toUpperCase())
    .join("");

  return /^[A-Z0-9]{1,2}$/.test(initials) ? initials : "?";
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderList(elementId, items) {
  const el = document.getElementById(elementId);
  el.innerHTML = "";

  if (!items.length) {
    el.innerHTML = "<li>No items found.</li>";
    return;
  }

  items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = typeof item === "string" ? item : item.text;
    el.appendChild(li);
  });
}

function setStatus(message) {
  document.getElementById("status").innerText = message;
}

function setMissionStatus(message, colour = "amber") {
  const status = document.getElementById("mission-status");
  const dot = document.getElementById("mission-dot");

  if (!status || !dot) return;

  status.innerText = message;

  dot.classList.remove(
    "status-green",
    "status-amber",
    "status-red"
  );

  dot.classList.add(`status-${colour}`);
}


// ── Dark mode toggle ──────────────────────────────────
// (localStorage is fine here — it's a UI preference, not a credential.)
(function () {
  const btn = document.getElementById('theme-toggle');
  const label = document.getElementById('theme-label');

  // Restore saved preference
  if (localStorage.getItem('kl-theme') === 'dark') {
    document.body.classList.add('dark');
    label.textContent = 'Dark';
  }

  btn.addEventListener('click', () => {
    const isDark = document.body.classList.toggle('dark');
    label.textContent = isDark ? 'Dark' : 'Light';
    localStorage.setItem('kl-theme', isDark ? 'dark' : 'light');
  });
})();
