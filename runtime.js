(function () {
  console.log("WWG runtime.js loaded");

  function onNewMessageComposeHandler(event) {
    console.log("WWG EVENT: OnNewMessageCompose fired");

    try {
      // Mensaje visible en la barra amarilla/azul de Outlook
      Office.context.mailbox.item.notificationMessages.add(
        "wwgTest",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "WWG: OnNewMessageCompose se ejecutó.",
          icon: "Icon.16x16", // opcional, definido en bt:Images
          persistent: false
        }
      );
    } catch (e) {
      console.error("Error en handler:", e);
    } finally {
      event.completed();
    }
  }

  Office.actions.associate(
    "onNewMessageComposeHandler",
    onNewMessageComposeHandler
  );
})();