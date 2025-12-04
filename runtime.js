// ===========================================================
//  WWG Auto Signature - runtime.js
// ===========================================================

// Espera a que la API de Office esté lista
Office.onReady().then(function () {
  // Asocia la función que Outlook llamará automáticamente
  // El nombre *debe coincidir* con el manifest.xml
  Office.actions.associate(
    "onNewMessageComposeHandler",
    onNewMessageComposeHandler
  );
});

// ===========================================================
//  Función principal: se ejecuta AUTOMÁTICAMENTE
//  cada vez que se crea un nuevo correo, reply o forward
// ===========================================================

async function onNewMessageComposeHandler(event) {
  try {
    // 1. Obtener el cuerpo actual del mensaje
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Error reading message body:", result.error);
          event.completed();
          return;
        }

        const body = result.value || "";

        // 2. Evitar duplicar la firma
        //    (Usamos un marcador invisible que solo tú conoces)
        const SIGNATURE_MARKER = "<!--WWG_SIGNATURE-->";

        if (body.includes(SIGNATURE_MARKER)) {
          console.log("Firma ya insertada. No se duplica.");
          event.completed();
          return;
        }

        // 3. Construir la firma HTML

        // TODO: aquí insertas tu firma real en HTML.
        // Te doy una plantilla simple. Cambia TODO.
        const signatureHTML = `
          ${SIGNATURE_MARKER}
          <div style="font-family:Arial, sans-serif; font-size:14px; color:#333;">
            <strong>William Moreno</strong><br>
            CEO – Will World Global<br>
            <a href="mailto:it.support@willworldglobal.com">it.support@willworldglobal.com</a><br>
            <br>
            <img src="https://willworldglobal.com/wp-content/uploads/2025/11/logo.png" width="180">
          </div>
        `;

        // 4. Insertar firma al inicio del correo
        Office.context.mailbox.item.body.prependAsync(
          signatureHTML,
          { coercionType: Office.CoercionType.Html },
          function (result2) {
            if (result2.status !== Office.AsyncResultStatus.Succeeded) {
              console.error("Error inserting signature:", result2.error);
            } else {
              console.log("Firma WWG insertada automáticamente.");
            }

            // SIEMPRE terminar con event.completed()
            event.completed();
          }
        );
      }
    );
  } catch (err) {
    console.error("Runtime error:", err);
    event.completed();
  }
}
