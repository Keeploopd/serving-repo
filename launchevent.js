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


async function getToken() {
  try {
    const stored = localStorage.getItem("keeploopd_token");
    if (stored) return stored;
  } catch (e) {}

  try {
    const msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    const result = await msalInstance.ssoSilent({
      ...apiRequest,
      loginHint: Office.context.mailbox.userProfile.emailAddress
    });
    return result.accessToken;
  } catch (e) {
    console.warn("ssoSilent fallback failed:", e);
    return null;
  }
}


async function sha256(value) {
  const encoded = new TextEncoder().encode(value);
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  return Array.from(new Uint8Array(hashBuffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

async function hashEmail(email) {
  return sha256(email.toLowerCase().trim());
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


async function runCoDraftCheck(item, token) {
  const conversationId = await getConversationKey(item);
  const userId = await hashEmail(Office.context.mailbox.userProfile.emailAddress);

  // Heartbeat
  await fetch("https://api.keeploopd.com/heartbeat", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify({ conversationId, userId, timestamp: Date.now() })
  });

  // Fetch count
  const res = await fetch(
    `https://api.keeploopd.com/active-drafters?conversationId=${encodeURIComponent(conversationId)}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!res.ok) throw new Error(`active-drafters returned ${res.status}`);

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
  event.completed();
  await runCheck();

  pollingInterval = setInterval(async () => {
    await runCheck();
  }, 15000);

  setTimeout(async () => {
    const item = Office.context.mailbox.item;
    try {
      await item.notificationMessages.removeAsync("codraftStatus");
      await item.notificationMessages.addAsync("codraftStatus", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Co-drafter monitoring paused. Click Sync Co-Drafters to resume.",
        icon: "Icon.16x16",
        persistent: false
      });
    } catch (e) {}
  }, 270000); 

  setTimeout(() => {
    if (pollingInterval) {
      clearInterval(pollingInterval);
      pollingInterval = null;
    }
  }, 300000); 
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
