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


async function getSilentToken() {
  const msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  const email = Office.context.mailbox.userProfile.emailAddress;

  const result = await msalInstance.ssoSilent({
    ...apiRequest,
    loginHint: email
  });

  return result.accessToken;
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

function getSubjectText(item) {
  return typeof item.subject === "string" ? item.subject : "";
}

async function getConversationKey(item) {
  const rawConversationId = item.conversationId || "";
  const subjectText = getSubjectText(item);
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

  // Show progress while in flight
  await item.notificationMessages.replaceAsync("codraftStatus", {
    type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
    message: "Checking for co-drafters...",
    persistent: false
  });


  await fetch("https://api.keeploopd.com/heartbeat", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify({ conversationId, userId, timestamp: Date.now() })
  });


  const res = await fetch(
    `https://api.keeploopd.com/active-drafters?conversationId=${encodeURIComponent(conversationId)}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!res.ok) {
    throw new Error(`active-drafters returned ${res.status}`);
  }

  const data = await res.json();
  const count = data.count ?? 0;

  if (count >= 1) {
    await item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `${count} person${count > 1 ? "s are" : " is"} currently drafting a reply`,
      icon: "Icon.16x16",
      persistent: true
    });
  } else {
    await item.notificationMessages.removeAsync("codraftStatus");
  }
}


async function onNewMessageCompose(event) {
  try {
    const token = await getSilentToken();
    const item = Office.context.mailbox.item;
    await runCoDraftCheck(item, token);
  } catch (err) {
    
    console.warn("onNewMessageCompose auth/check failed:", err);
    Office.context.mailbox.item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Open Co-Draft to sign in and enable co-drafter detection.",
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());
    return;
  }

  event.completed();
}


async function syncCoDrafters(event) {
  try {
    const token = await getSilentToken();
    const item = Office.context.mailbox.item;
    await runCoDraftCheck(item, token);
  } catch (err) {
    console.error("syncCoDrafters failed:", err);
    Office.context.mailbox.item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Co-drafter check failed. Please try again or open the Co-Draft taskpane.",
      persistent: false
    }, () => event.completed());
    return;
  }

  event.completed();
}


Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("syncCoDrafters", syncCoDrafters);
