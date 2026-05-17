// launchevent.js
// Handles:
//   1. OnNewMessageCompose — fires automatically when a compose window opens
//   2. syncCoDrafters      — fires when the "Sync Co-Drafters" ribbon button is clicked
//
// Token strategy: reads from localStorage written by taskpane.js after auth.
// ssoSilent is NOT used here — it requires an iframe which is blocked in the
// restricted event-based activation runtime.

Office.initialize = function () {};

// ---------------------------------------------------------------------------
// Token retrieval — reads what taskpane.js stored after sign-in
// ---------------------------------------------------------------------------

function getStoredToken() {
  try {
    return localStorage.getItem("keeploopd_token");
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Conversation key helpers (kept consistent with taskpane.js)
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Core detection logic — shared by both entry points
// ---------------------------------------------------------------------------

async function runCoDraftCheck(item, token) {
  const conversationId = await getConversationKey(item);
  const userId = await hashEmail(Office.context.mailbox.userProfile.emailAddress);

  // Show progress while in flight — no persistent property on ProgressIndicator
  await item.notificationMessages.replaceAsync("codraftStatus", {
    type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
    message: "Checking for co-drafters..."
  });

  // Send heartbeat so this user registers as an active drafter
  await fetch("https://api.keeploopd.com/heartbeat", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify({ conversationId, userId, timestamp: Date.now() })
  });

  // Fetch active drafter count
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

// ---------------------------------------------------------------------------
// Entry point 1: LaunchEvent — fires automatically when compose opens
// ---------------------------------------------------------------------------

async function onNewMessageCompose(event) {
  const item = Office.context.mailbox.item;
  const token = getStoredToken();

  if (!token) {
    item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Open Co-Draft to sign in and enable co-drafter detection.",
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());
    return;
  }

  try {
    await runCoDraftCheck(item, token);
  } catch (err) {
    console.error("onNewMessageCompose check failed:", err);
    item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Co-drafter check failed. Click Sync Co-Drafters to retry.",
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());
    return;
  }

  event.completed();
}

// ---------------------------------------------------------------------------
// Entry point 2: ExecuteFunction — fires when "Sync Co-Drafters" button clicked
// ---------------------------------------------------------------------------

async function syncCoDrafters(event) {
  const item = Office.context.mailbox.item;
  const token = getStoredToken();

  if (!token) {
    item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Open Co-Draft to sign in and enable co-drafter detection.",
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());
    return;
  }

  try {
    await runCoDraftCheck(item, token);
  } catch (err) {
    console.error("syncCoDrafters failed:", err);
    item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Co-drafter check failed. Please try again.",
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());
    return;
  }

  event.completed();
}

// ---------------------------------------------------------------------------
// Register both functions with Office
// ---------------------------------------------------------------------------

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("syncCoDrafters", syncCoDrafters);
