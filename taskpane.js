async function hashEmail(email) {
  const encoded = new TextEncoder().encode(email.toLowerCase().trim());
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessToken(
      { allowSignInPrompt: true, allowConsentPrompt: true },
      (result) => {
        if (result.status === "succeeded") {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

async function init() {
  const item = Office.context.mailbox.item;
  const conversationId = item.conversationId;
  const rawEmail = Office.context.mailbox.userProfile.emailAddress;
  const userId = await hashEmail(rawEmail);

  async function sendHeartbeat() {
    try {
      const token = await getAccessToken();
      await fetch("https://api.keeploopd.com/heartbeat", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${token}`    // token sent in header
        },
        body: JSON.stringify({ conversationId, userId, timestamp: Date.now() })
      });
    } catch (err) {
      console.error("Heartbeat failed:", err);
    }
  }

  async function updateBanner() {
    try {
      const token = await getAccessToken();
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

Office.onReady(() => init());