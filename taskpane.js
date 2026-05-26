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

let currentToken = null;
let monitoringStarted = false;

async function hashEmail(email) {
  const encoded = new TextEncoder().encode(email.toLowerCase().trim());
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
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

          const response = await fetch("https://api.keeploopd.com/api/auth/test", {
            headers: {
              Authorization: `Bearer ${payload.accessToken}`
            }
          });

          const text = await response.text();
          console.log("Backend response", response.status, text);

          if (!response.ok) {
            document.getElementById("status").textContent =
              `Backend auth failed: ${response.status}`;
            return;
          }

          currentToken = payload.accessToken;
          try { localStorage.setItem("keeploopd_token", currentToken); } catch (e) {}
          document.getElementById("signin").style.display = "none";
          document.getElementById("status").textContent = "Authenticated successfully";
          await init();

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
let bannerInterval = null;
let lastNotificationCount = null;

function stopMonitoring() {
  if (heartbeatInterval) clearInterval(heartbeatInterval);
  if (bannerInterval) clearInterval(bannerInterval);

  heartbeatInterval = null;
  bannerInterval = null;
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
        //message: `${count} people are currently drafting replies`,
        message: `${count} ${count > 1 ? "people are" : "person is"} currently drafting a reply`,
        icon: "Icon.16x16",
        persistent: false
      });
    }
  });
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

  const item = Office.context.mailbox.item;
  //const conversationId = item.conversationId.slice(-24);
  const conversationId = await getConversationKey();
  const rawEmail = Office.context.mailbox.userProfile.emailAddress;
  const userId = await hashEmail(rawEmail);

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
      const hb = await fetch("https://api.keeploopd.com/heartbeat", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${currentToken}`
        },
        body: JSON.stringify({
          conversationId,
          userId,
          timestamp: Date.now()
        })
      });

      console.log("heartbeat status", hb.status, await hb.text());
    } catch (err) {
      console.error("Heartbeat failed:", err);
    }
  }

  async function updateBanner() {
    try {
      const res = await fetch(
        `https://api.keeploopd.com/active-drafters?conversationId=${encodeURIComponent(conversationId)}`,
        {
          headers: {
            "Authorization": `Bearer ${currentToken}`
          }
        }
      );

      if (!res.ok) {
        console.error("active-drafters failed", res.status, await res.text());
        return;
      }

      const data = await res.json();
      const banner = document.getElementById("banner");

      if (data.count >= 1) {
        banner.style.display = "block";
        banner.textContent = `${data.count} active drafter(s) detected`;
      } else {
        banner.style.display = "none";
        banner.textContent = "";
      }

      updateNotification(data.count);

    } catch (err) {
      console.error("Banner update failed:", err);
    }
  }

  await sendHeartbeat();
  await updateBanner();

  heartbeatInterval = setInterval(sendHeartbeat, 10000);
  bannerInterval = setInterval(updateBanner, 5000);
}


async function trySilentAuth() {
  try {
    await msalInstance.initialize();

    try {
      const silentResult = await msalInstance.ssoSilent({
        scopes: apiRequest.scopes,
        loginHint: Office.context.mailbox.userProfile.emailAddress
      });

      return silentResult.accessToken;
    } catch (ssoErr) {
      console.warn("ssoSilent failed", ssoErr);
    }

    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) return null;

    const result = await msalInstance.acquireTokenSilent({
      ...apiRequest,
      account: accounts[0]
    });

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

  const silentToken = await trySilentAuth();

  if (silentToken) {
    currentToken = silentToken;
    try { localStorage.setItem("keeploopd_token", currentToken); } catch (e) {}
    signinButton.style.display = "none";
    await init();
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
    try { localStorage.removeItem("keeploopd_token"); } catch (e) {}
  });
  document
    .getElementById("refresh-analysis")
    .addEventListener("click", refreshAnalysis);

  await loadMissionControl();
});

async function getConversationContext() {
  const item = Office.context.mailbox.item;

  return {
    conversationId: item.conversationId,
    subject: item.subject,
    latestMessageSentAtUtc: new Date().toISOString()
  };
}

async function getAuthToken() {
  return await OfficeRuntime.auth.getAccessToken({
    allowSignInPrompt: true,
    allowConsentPrompt: true
  });
}

async function loadMissionControl() {
  setStatus("Loading Mission Control...");

  const context = await getConversationContext();
  const token = await getAuthToken();

  const url = new URL("/api/thread/state", window.location.origin);
  url.searchParams.set("conversationId", context.conversationId);
  url.searchParams.set("latestMessageSentAtUtc", context.latestMessageSentAtUtc);

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!response.ok) {
    setStatus("Unable to load Mission Control.");
    return;
  }

  const data = await response.json();

  renderMissionControl(data);

  if (data.status === "refreshing") {
    pollUntilReady(context.conversationId, token);
  }
}

async function refreshAnalysis() {
  setStatus("Refreshing analysis...");

  const context = await getConversationContext();
  const token = await getAuthToken();

  const threadText = await getCurrentEmailText();

  await fetch("/api/thread/analyse", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify({
      ...context,
      threadText
    })
  });

  await loadMissionControl();
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
    const url = new URL("/api/thread/state", window.location.origin);
    url.searchParams.set("conversationId", conversationId);

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const data = await response.json();
    renderMissionControl(data);

    if (data.status === "ready") {
      clearInterval(interval);
    }
  }, 3000);
}

function renderMissionControl(data) {
  const state = data.state || {};

  setStatus(data.status === "refreshing"
    ? "Refreshing Mission Control..."
    : "Mission Control ready");

  renderList("participants", state.activeParticipants || []);
  renderList("mission-control", state.missionControl || []);
  renderList("open-questions", state.openQuestions || []);

  document.getElementById("reply-focus").innerText =
    state.suggestedReplyFocus?.reason || "No reply focus detected.";

  document.getElementById("thread-health").innerText =
    state.threadHealth?.summary || "No thread health available.";
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
