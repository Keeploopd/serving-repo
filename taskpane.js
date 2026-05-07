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

          console.log("Access token received", payload.accessToken);

          const response = await fetch("https://api.keeploopd.com/api/auth/test", {
            headers: {
              Authorization: `Bearer ${payload.accessToken}`
            }
          });
          
          const text = await response.text();
          
          console.log("Backend response", response.status, text);
          
          if (response.ok) {
            document.getElementById("status").textContent = "Authenticated successfully";
            const banner = document.getElementById("banner");
            banner.style.display = "block";
            banner.textContent = "Backend token validation succeeded.";
          } else {
            document.getElementById("status").textContent =
              `Backend auth failed: ${response.status}`;
          }

        } catch (err) {
          console.error("Taskpane token handling error", err);
          document.getElementById("status").textContent =
            `Taskpane error: ${err.message || err}`;
        }
      });
    }
  );
}

async function getAccessToken() {
  await msalInstance.initialize();
  
  // Handle redirect response first
  await msalInstance.handleRedirectPromise();
  
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    try {
      // Try silent token acquisition first
      const result = await msalInstance.acquireTokenSilent({
        ...apiRequest,
        account: accounts[0]
      });
      return result.accessToken;
    } catch (err) {
      // Silent failed, fall through to popup
      console.warn("Silent token acquisition failed, trying popup:", err);
    }
  }

  // Popup login
  const result = await msalInstance.loginPopup(apiRequest);
  return result.accessToken;
}

async function init() {
  const item = Office.context.mailbox.item;
  const conversationId = item.conversationId;
  const rawEmail = Office.context.mailbox.userProfile.emailAddress;
  const userId = await hashEmail(rawEmail);

  document.getElementById("status").innerText = "Authenticating...";

  let token;
  try {
    token = await getAccessToken();
    document.getElementById("status").innerText = "Monitoring active drafters...";
  } catch (err) {
    document.getElementById("status").innerText = "Authentication failed. Please reload.";
    console.error("Auth error:", err);
    return;
  }

  async function sendHeartbeat() {
    try {
      // Refresh token silently on each heartbeat
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        const result = await msalInstance.acquireTokenSilent({
          ...apiRequest,
          account: accounts[0]
        });
        token = result.accessToken;
      }

      await fetch("https://api.keeploopd.com/heartbeat", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${token}`
        },
        body: JSON.stringify({
          conversationId,
          userId,
          timestamp: Date.now()
        })
      });
    } catch (err) {
      console.error("Heartbeat failed:", err);
    }
  }

  async function updateBanner() {
    try {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        const result = await msalInstance.acquireTokenSilent({
          ...apiRequest,
          account: accounts[0]
        });
        token = result.accessToken;
      }

      const res = await fetch(
        `https://api.keeploopd.com/active-drafters?conversationId=${conversationId}`,
        { headers: { "Authorization": `Bearer ${token}` } }
      );
      const data = await res.json();
      const banner = document.getElementById("banner");
      if (data.count > 1) {
        banner.style.display = "block";
        banner.innerText = `${data.count} people are currently drafting replies`;
      } else {
        banner.style.display = "none";
        banner.innerText = "";
      }
    } catch (err) {
      console.error("Banner update failed:", err);
    }
  }

  sendHeartbeat();
  updateBanner();
  setInterval(sendHeartbeat, 30000);
  setInterval(updateBanner, 30000);
}

//Office.onReady(() => init());
