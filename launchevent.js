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

async function onNewMessageCompose(event) {
  try {
    const token = await getSilentToken();

    const res = await fetch("https://api.keeploopd.com/health", {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    Office.context.mailbox.item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: res.ok
        ? "Co-Draft background auth succeeded."
        : `Co-Draft background auth failed: ${res.status}`,
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());

  } catch (err) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("codraftStatus", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Open Co-Draft to sign in.",
      icon: "Icon.16x16",
      persistent: false
    }, () => event.completed());
  }
}

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
