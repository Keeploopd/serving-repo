Office.initialize = function () {};


const msalConfig = {
  auth: {
    clientId: "705cf97b-720b-4240-b6d0-02a6655300b2",
    authority: "https://login.microsoftonline.com/5904ae0b-47e9-4b06-843e-60769342a32b",
    redirectUri: "https://co-draft.keeploopd.com/commands.html"
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

const apiRequest = {
  scopes: ["api://co-draft.keeploopd.com/705cf97b-720b-4240-b6d0-02a6655300b2/access_as_user"]
};

const API_BASE = "https://api.keeploopd.com";

// Bump with each deploy — confirms which build the command runtime executes
// (classic Outlook caches this file too).
const BUILD = "2026-07-13.1";
console.log("Keeploopd launchevent build", BUILD);

const BRIDGE_KEY = "keeploopd_token";
const POLL_MS = 15000;         // check every 15s
// 4.5 min: deliberately inside the host's ~5 min budget for both event-based
// runtimes AND held-open function commands, so WE end the window (and show
// the re-sync message) before the host tears the runtime down under us.
const POLL_WINDOW_MS = 270000;


// ── Auth ─────────────────────────────────────────────────────────────────────
// Token sources, in order of preference:
//   1. In-memory cache (expiry-checked)
//   2. OfficeRuntime.auth.getAccessToken — Office SSO; native and fast,
//      works in command/event runtimes when WebApplicationInfo is in the
//      manifest. The backend accepts the api:// audience these tokens carry.
//   3. MSAL ssoSilent — works where the runtime supports hidden iframes (OWA).
//   4. localStorage bridge written by the taskpane — LAST RESORT, expiry-
//      checked on every read, expired tokens deleted on sight. Delete this
//      once the add-in moves to a SharedRuntime or Nested App Auth.

let msalInstancePromise = null;
let cachedToken = null;

function getMsal() {
  if (!msalInstancePromise) {
    msalInstancePromise = (async () => {
      const instance = new msal.PublicClientApplication(msalConfig);
      // Timeboxed — this was the one un-budgeted await in the token chain.
      await withTimeout(instance.initialize(), 4000, "MSAL initialize");
      return instance;
    })().catch((e) => {
      console.warn("MSAL unavailable in command runtime:", e);
      msalInstancePromise = null; // allow retry on a later tick
      return null;
    });
  }
  return msalInstancePromise;
}

function tokenExpired(token) {
  try {
    const payloadB64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const { exp } = JSON.parse(atob(payloadB64));
    return !exp || Date.now() / 1000 > exp - 60;
  } catch (e) {
    return true;
  }
}

function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`${label} timed out after ${ms}ms`)), ms)
    )
  ]);
}

async function tryOfficeSso() {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.auth) return null;
    const token = await withTimeout(
      OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: false,   // never prompt from a UI-less runtime
        allowConsentPrompt: false
      }),
      4000,
      "Office SSO"
    );
    return token || null;
  } catch (e) {
    console.warn("Office SSO unavailable in this runtime:", e);
    return null;
  }
}

async function tryMsalSilent() {
  try {
    const msalInstance = await getMsal();
    if (!msalInstance) return null;
    const result = await withTimeout(
      msalInstance.ssoSilent({
        ...apiRequest,
        loginHint: Office.context.mailbox.userProfile.emailAddress
      }),
      5000,
      "ssoSilent"
    );
    return result.accessToken;
  } catch (e) {
    console.warn("ssoSilent failed:", e);
    return null;
  }
}

function tryBridgeToken() {
  try {
    const stored = localStorage.getItem(BRIDGE_KEY);
    if (!stored) return null;
    if (tokenExpired(stored)) {
      localStorage.removeItem(BRIDGE_KEY); // never leave dead tokens at rest
      return null;
    }
    return stored;
  } catch (e) {
    return null;
  }
}

async function getToken() {
  if (cachedToken && !tokenExpired(cachedToken)) return cachedToken;

  cachedToken =
    (await tryOfficeSso()) ||
    (await tryMsalSilent()) ||
    tryBridgeToken();

  return cachedToken;
}


// ── Conversation key ─────────────────────────────────────────────────────────

async function sha256(value) {
  const encoded = new TextEncoder().encode(value);
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  return Array.from(new Uint8Array(hashBuffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

async function getConversationKey(item) {
  const rawConversationId = item.conversationId || "";
  const subjectText = typeof item.subject === "string" ? item.subject : "";
  const subject = subjectText
    .toLowerCase()
    .replace(/^(re|fw|fwd):\s*/i, "")
    .trim();

  const markerIndex = rawConversationId.lastIndexOf("AQ");
  const sharedConversationPart =
    markerIndex >= 0
      ? rawConversationId.slice(markerIndex)
      : rawConversationId.slice(-24);

  return sha256(`${sharedConversationPart}|${subject}`);
}


// ── Notifications ────────────────────────────────────────────────────────────

function showInfoNotification(message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.16x16",
      persistent: false
    });
  } catch (e) {
    console.warn("Could not show notification:", e);
  }
}


// ── Co-draft check ───────────────────────────────────────────────────────────

async function runCoDraftCheck(item, token) {
  const conversationId = await getConversationKey(item);

  // No userId in the body — the backend derives identity from the bearer
  // token, so presence can't be spoofed. The heartbeat response includes the
  // current drafter count, so no separate /active-drafters call is needed.
  const res = await fetch(`${API_BASE}/heartbeat`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify({ conversationId })
  });

  if (!res.ok) throw new Error(`heartbeat returned ${res.status}`);

  const data = await res.json();
  const count = data.count ?? 0;

  await item.notificationMessages.removeAsync("codraftStatus");

  if (count >= 1) {
    await item.notificationMessages.addAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `${count} ${count > 1 ? "people are" : "person is"} currently drafting a reply`,
      icon: "Icon.16x16",
      persistent: false
    });
  }
}


async function runCheck() {
  const item = Office.context.mailbox.item;
  const token = await getToken();

  if (!token) {
    // Every further tick would fail identically — end the window (no resume
    // message; the sign-in message below is the actionable one).
    endPollingWindow(false);
    showInfoNotification("Open Keeploopd Panel to sign in, then click Sync Co-Drafters.");
    return;
  }

  try {
    await runCoDraftCheck(item, token);
  } catch (err) {
    console.error("Co-draft check failed:", err);
    cachedToken = null; // possibly revoked/rejected — force re-acquire next tick
    showInfoNotification("Co-drafter check failed. Click Sync Co-Drafters to retry.");
  }
}


// ── Polling window ───────────────────────────────────────────────────────────
// One managed window shared by both entry points.
//
// Two lifecycle tricks make this behave the same from the Sync button as it
// does from the compose event:
//
//   1. HELD-OPEN EVENT: UI-less function-command runtimes are torn down
//      almost immediately after event.completed() — which is why the Sync
//      button previously fired only a single check. We now HOLD the event
//      open for the duration of the window and complete it in
//      endPollingWindow(). POLL_WINDOW_MS sits inside the host's ~5 min
//      command budget so the host never times us out first.
//
//   2. TICK-DRIVEN END: the "monitoring paused" message is shown from an
//      interval tick (elapsed-time check) rather than only a separate
//      setTimeout — throttled/suspended runtimes can skip a lone timeout,
//      which is why the re-sync message wasn't appearing reliably at the end
//      of the compose window. A backup timeout remains as a belt-and-braces.

let pollingInterval = null;
let pollingEndTimeout = null;
let pollingStartedAt = 0;
let pendingEvent = null;

function completePendingEvent() {
  if (pendingEvent) {
    try { pendingEvent.completed(); } catch (e) {}
    pendingEvent = null;
  }
}

function stopPolling() {
  if (pollingInterval) {
    clearInterval(pollingInterval);
    pollingInterval = null;
  }
  if (pollingEndTimeout) {
    clearTimeout(pollingEndTimeout);
    pollingEndTimeout = null;
  }
}

function endPollingWindow(showResumeMessage) {
  stopPolling();
  if (showResumeMessage) {
    showInfoNotification("Co-drafter monitoring paused. Click Sync Co-Drafters to resume.");
  }
  completePendingEvent();
}

function startPollingWindow(event = null) {
  stopPolling();
  completePendingEvent(); // release any previous held event before replacing it
  pendingEvent = event;
  pollingStartedAt = Date.now();

  pollingInterval = setInterval(() => {
    if (Date.now() - pollingStartedAt >= POLL_WINDOW_MS) {
      endPollingWindow(true); // message shown from a tick we know is firing
      return;
    }
    runCheck();
  }, POLL_MS);

  pollingEndTimeout = setTimeout(() => endPollingWindow(true), POLL_WINDOW_MS + POLL_MS);
}


// ── Entry points ─────────────────────────────────────────────────────────────

async function onNewMessageCompose(event) {
  // Launch events must complete promptly; the event-based runtime persists
  // on its own (host policy, ~5 min) so nothing needs holding open here.
  event.completed();
  startPollingWindow();
  await runCheck();
}


async function syncCoDrafters(event) {
  // Event is held open by the window manager — see lifecycle note above.
  startPollingWindow(event);
  await runCheck();
}


Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("syncCoDrafters", syncCoDrafters);
