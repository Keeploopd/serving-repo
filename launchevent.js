function onNewMessageCompose(event) {
  Office.context.mailbox.item.notificationMessages.addAsync("codraftAutoStart", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Co-Draft is available. Open the add-in to start monitoring.",
    icon: "Icon.16x16",
    persistent: false
  });

  event.completed();
}

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);