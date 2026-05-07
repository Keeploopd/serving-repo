// MSAL configuration
const msalConfig = {
  auth: {
    clientId: "your-azure-client-id",
    authority: "https://login.microsoftonline.com/your-tenant-id",
    redirectUri: "https://gitsubdomain.mydomain.com/taskpane.html"
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
  scopes: ["openid", "profile", "User.Read"]
};

async function hashEmail(email) {
  const encoded = new TextEncoder().encode(email.toLowerCase().trim());
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
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
        ...loginRequest,
        account: accounts[0]
      });
      return result.accessToken;
    } catch (err) {
      // Silent failed, fall through to popup
      console.warn("Silent token acquisition failed, trying popup:", err);
    }
  }

  // Popup login
  const result = await msalInstance.loginPopup(loginRequest);
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
          ...loginRequest,
          account: accounts[0]
        });
        token = result.accessToken;
      }

      await fetch("https://api.your-domain.com/heartbeat", {
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
          ...loginRequest,
          account: accounts[0]
        });
        token = result.accessToken;
      }

      const res = await fetch(
        `https://api.your-domain.com/active-drafters?conversationId=${conversationId}`,
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

Office.onReady(() => init());
