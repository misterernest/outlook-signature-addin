(function () {
  console.log("WWG runtime.js loaded");

  function onNewMessageComposeHandler(event) {
    try {
      const testBody = `
        <p><strong>WWG TEST AUTO SIGNATURE</strong></p>
        <p>If you see this text, OnNewMessageCompose is working.</p>
      `;

      Office.context.mailbox.item.body.setAsync(
        testBody,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("setAsync failed:", asyncResult.error);
          }
          // Siempre se debe llamar, aunque falle
          event.completed();
        }
      );
    } catch (e) {
      console.error("Exception in handler:", e);
      event.completed();
    }
  }

  // Registra el handler con el nombre usado en el manifest
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
})();
