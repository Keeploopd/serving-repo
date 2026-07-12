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


// ── Auth ─────────────────────────────────────────────────────────────────────

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

async function getToken() {
  if (cachedToken && !tokenExpired(cachedToken)) return cachedToken;

  try {
    const msalInstance = await getMsal();
    const result = await msalInstance.ssoSilent({
      ...apiRequest,
      loginHint: Office.context.mailbox.userProfile.emailAddress
    });
    cachedToken = result.accessToken;
    return cachedToken;
  } catch (e) {
    console.warn("ssoSilent fallback failed:", e);
    return null;
  }
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


// ── Co-draft check ───────────────────────────────────────────────────────────

async function runCoDraftCheck(item, token) {
  const conversationId = await getConversationKey(item);

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
    item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Open Keeploopd Panel to sign in and enable co-drafter detection.",
      icon: "Icon.16x16",
      persistent: false
    }, () => { if (event) event.completed(); });
    return;
  }

  try {
    await runCoDraftCheck(item, token);
  } catch (err) {
    console.error("Co-draft check failed:", err);
    item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Co-drafter check failed. Please open Keeploopd Panel to sign in.",
      icon: "Icon.16x16",
      persistent: false
    });
  }

  if (event) event.completed();
}

let pollingInterval = null;

async function onNewMessageCompose(event) {
  try {
    await runCheck();
  } finally {
    event.completed();
  }
}


async function syncCoDrafters(event) {
  if (pollingInterval) clearInterval(pollingInterval);

  pollingInterval = setInterval(async () => {
    await runCheck();
  }, 15000);

  setTimeout(() => {
    if (pollingInterval) {
      clearInterval(pollingInterval);
      pollingInterval = null;
    }
  }, 300000);

  await runCheck(event);
}


Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("syncCoDrafters", syncCoDrafters);
