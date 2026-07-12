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

const BRIDGE_KEY = "keeploopd_token";
const POLL_MS = 15000;        // check every 15s
const POLL_WINDOW_MS = 300000; // stop after 5 min and ask the user to re-sync


// ── Auth ─────────────────────────────────────────────────────────────────────
// Token sources, in order of preference:
//   1. In-memory cache (expiry-checked)
//   2. OfficeRuntime.auth.getAccessToken — Office SSO, works in command/event
//      runtimes when WebApplicationInfo is configured in the manifest; the
//      backend accepts the api:// audience these tokens carry. No storage,
//      no prompts.
//   3. MSAL ssoSilent — works where the runtime supports hidden iframes (OWA).
//   4. localStorage bridge written by the taskpane — LAST RESORT. This is the
//      only channel Outlook guarantees between the taskpane and this runtime,
//      which is why it existed originally. It is now expiry-checked on every
//      read and expired tokens are deleted immediately, so the worst case is
//      a short-lived (~1h) access token at rest, not a stale credential
//      forever. When you move to a SharedRuntime or Nested App Auth, delete
//      this bridge (source 4) and the taskpane's storeBridgeToken().

let msalInstancePromise = null;
let cachedToken = null;

function getMsal() {
  if (!msalInstancePromise) {
    msalInstancePromise = (async () => {
      const instance = new msal.PublicClientApplication(msalConfig);
      await instance.initialize();
      return instance;
    })();
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

async function tryOfficeSso() {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.auth) return null;
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false,   // never prompt from a UI-less runtime
      allowConsentPrompt: false
    });
    return token || null;
  } catch (e) {
    console.warn("Office SSO unavailable in this runtime:", e);
    return null;
  }
}

async function tryMsalSilent() {
  try {
    const msalInstance = await getMsal();
    const result = await msalInstance.ssoSilent({
      ...apiRequest,
      loginHint: Office.context.mailbox.userProfile.emailAddress
    });
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


async function runCheck(event = null) {
  const item = Office.context.mailbox.item;
  const token = await getToken();

  if (!token) {
    // Stop the polling window — every further tick would fail identically.
    stopPolling();
    showInfoNotification("Open Keeploopd Panel to sign in, then click Sync Co-Drafters.");
    if (event) event.completed();
    return;
  }

  try {
    await runCoDraftCheck(item, token);
  } catch (err) {
    console.error("Co-draft check failed:", err);
    cachedToken = null; // could be a revoked/rejected token — force re-acquire next tick
    showInfoNotification("Co-drafter check failed. Click Sync Co-Drafters to retry.");
  }

  if (event) event.completed();
}


// ── Polling window ───────────────────────────────────────────────────────────
// One managed window shared by both entry points. Fixes two issues in the
// original implementation:
//   1. Restarting (clicking Sync again) never cleared the previous window's
//      setTimeouts, so a stale timeout could kill a fresh window early.
//   2. The "paused" message was shown at 4.5 min while polling ran until
//      5 min — the remaining ticks overwrote the message. Now the interval is
//      stopped FIRST, then the re-sync message is shown.
// NOTE: hosts may tear this runtime down after event.completed(); where that
// happens polling ends early and the user still has the Sync button. For
// guaranteed continuous monitoring, move it into the taskpane/SharedRuntime.

let pollingInterval = null;
let pollingEndTimeout = null;

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

function startPollingWindow() {
  stopPolling();

  pollingInterval = setInterval(() => {
    runCheck();
  }, POLL_MS);

  pollingEndTimeout = setTimeout(() => {
    stopPolling();
    showInfoNotification("Co-drafter monitoring paused. Click Sync Co-Drafters to resume.");
  }, POLL_WINDOW_MS);
}


// ── Entry points ─────────────────────────────────────────────────────────────

async function onNewMessageCompose(event) {
  // Complete the launch event promptly, then poll for as long as the host
  // keeps this runtime alive (bounded by POLL_WINDOW_MS).
  event.completed();
  startPollingWindow();
  await runCheck();
}


async function syncCoDrafters(event) {
  startPollingWindow();
  await runCheck(event); // immediate check; completes the button event
}


Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("syncCoDrafters", syncCoDrafters);
