// MSAL configuration
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

async function init() {
  if (monitoringStarted) return;

  if (!currentToken) {
    document.getElementById("status").textContent = "Please authenticate first.";
    return;
  }

  monitoringStarted = true;

  const item = Office.context.mailbox.item;
  const conversationId = item.conversationId.slice(-24);
  const rawEmail = Office.context.mailbox.userProfile.emailAddress;
  const userId = await hashEmail(rawEmail);

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

let lastNotificationCount = null;
  
function updateNotification(count) {
  if (count === lastNotificationCount) return;

  lastNotificationCount = count;

  Office.context.mailbox.item.notificationMessages.removeAsync("codraftStatus", () => {
    if (count >= 1) {
      Office.context.mailbox.item.notificationMessages.addAsync("codraftStatus", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `${count} people are currently drafting replies`,
        icon: "Icon.16x16",
        persistent: false
      });
    }
  });
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

  setInterval(sendHeartbeat, 10000);
  setInterval(updateBanner, 5000);
}

async function trySilentAuth() {
  try {
    await msalInstance.initialize();

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
  document.getElementById("signin").addEventListener("click", signInAndCallBackend);

  const silentToken = await trySilentAuth();

  if (silentToken) {
    currentToken = silentToken;
    document.getElementById("signin").style.display = "none";
    await init();
  } else {
    document.getElementById("status").textContent = "Sign in required.";
    document.getElementById("signin").style.display = "block";
  }
});
