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
let bannerInterval = null;
let lastNotificationCount = null;
let missionAnalysisInProgress = false;

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

  const refreshButton = document.getElementById("refresh-analysis");
  if (refreshButton) {
    refreshButton.addEventListener("click", refreshAnalysis);
  }

  const silentToken = await trySilentAuth();

  if (silentToken) {
    currentToken = silentToken;

    try {
      localStorage.setItem("keeploopd_token", currentToken);
    } catch (e) {}

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

    try {
      localStorage.removeItem("keeploopd_token");
    } catch (e) {}
  });
});

async function getValidToken() {
  if (currentToken) return currentToken;

  currentToken = await getAuthToken();
  return currentToken;
}

async function getConversationContext() {
  const item = Office.context.mailbox.item;

  console.log(
    "itemId:",
    item.itemId
  );

  console.log(
    "conversationId:",
    item.conversationId
  );

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

async function loadMissionControl() {
  try {
    setMissionStatus(
      "Loading Mission Control...",
      "amber"
    );

    const context = await getConversationContext();
    const token = await getValidToken();

    const url = new URL(
      "/api/thread/state", API_BASE
    );

    url.searchParams.set(
      "conversationId",
      context.conversationId
    );

    // debug in dev tools
    console.log("MC fetching:", url.toString());

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      throw new Error(
        `Mission Control request failed: ${response.status}`
      );
    }

    const data = await response.json();
    
    // debug
    console.log("THREAD_STATE response:", data);

    if (
      (data.status === "missing" || data.isStale === true) &&
      !missionAnalysisInProgress
    ) {
      await refreshAnalysis();
      return;
    }

    renderMissionControl(data);

    setMissionStatus(
      data.status === "refreshing"
        ? "Refreshing Mission Control..."
        : "Mission Control ready",
      data.status === "refreshing"
        ? "amber"
        : "green"
    );

    if (data.status === "refreshing") {
      pollUntilReady(
        context.conversationId,
        token
      );
    }

  } catch (err) {
    console.error(
      "Mission Control load failed:",
      err
    );

    setMissionStatus(
      "Mission Control unavailable",
      "red"
    );
  }
}

async function getRecipients() {
  const item = Office.context.mailbox.item;

  const getAsync = (prop) => new Promise(resolve => {
    prop.getAsync(r => {
      console.log("getAsync result:", r.status, r.value, r.error);
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null);
    });
  });

  const from = await getAsync(item.from);
  const toList = await getAsync(item.to);
  const ccList = await getAsync(item.cc);

  console.log("from:", from, "to:", toList, "cc:", ccList);

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

  console.log("getRecipients result:", recipients);
  return recipients;
}


async function refreshAnalysis() {
  console.trace("refreshAnalysis called");
  if (missionAnalysisInProgress) return;

  missionAnalysisInProgress = true;

  try {
    setMissionStatus("Refreshing analysis...", "amber");

    const context = await getConversationContext();

    const recipients = await getRecipients();
    //const recipients = getRecipients();
    const token = await getValidToken();
    const threadText = await getCurrentEmailText();

    console.log("recipients before hash:", recipients);

    const participants = await Promise.all(
      recipients
        .filter(r => r.emailAddress)
        .map(async (r) => ({
          participant_hash: await hashEmail(r.emailAddress),
          display_name: r.displayName || r.emailAddress
        }))
    );

    console.log("participants to send:", participants);

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

    if (!response.ok) {
      throw new Error("Analysis failed");
    }

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

  console.log("renderMissionControl data:", data);
  console.log("state:", state);

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

    div.innerHTML = `
      <div class="reply-avatar" style="background:#0f6cbd;">
        ${getInitials(person)}
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

// participant rendering 
function formatParticipantName(name) {
  if (!name) return "";

  const parts = name.trim().split(" ");
  if (parts.length === 1) return parts[0];

  return `${parts[0]} ${parts[1][0].toUpperCase()}`;
}

function renderParticipants(participants) {
  const el = document.getElementById("participants-list");
  el.innerHTML = "";

  if (!participants?.length) {
    el.innerHTML = `<div class="empty-state">No participants yet.</div>`;
    return;
  }

  participants.forEach((p) => {
    //const name = p.displayName || "Unknown";
    const name = p.display_name || p.displayName || p.email || "Unknown";

    const isDropped = p.status === "dropped";

    const div = document.createElement("div");
    div.className = "participant";

    div.innerHTML = `
      <div class="avatar-wrap">
        <div class="avatar" style="background:#0f6cbd;">
          ${getInitials(name)}
        </div>

        <div class="avatar-status ${isDropped ? "dropped" : "active"}"></div>
      </div>

      <div class="participant-name">${escapeHtml(formatParticipantName(name))}</div>

      <div class="participant-role">
        ${isDropped ? "Dropped" : "Active"}
      </div>
    `;

    el.appendChild(div);
  });
}

function getInitials(value) {
  if (!value) return "?";

  return value
    .split(/[ .@]/)
    .filter(Boolean)
    .slice(0, 2)
    .map(x => x[0].toUpperCase())
    .join("");
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
