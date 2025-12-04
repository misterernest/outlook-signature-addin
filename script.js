Office.onReady(() => {
  const btn = document.getElementById("btnInsert");
  if (btn) {
    btn.onclick = insertSignature;
  }
});

function insertSignature() {
  // Carga el HTML de tu firma desde el mismo sitio del add-in
  fetch("signature-test.html")
    .then(resp => {
      if (!resp.ok) {
        throw new Error("HTTP " + resp.status);
      }
      return resp.text();
    })
    .then(html => {
      Office.context.mailbox.item.body.setAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        result => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            alert("Failed: " + result.error.message);
          }
        }
      );
    })
    .catch(err => {
      console.error("Error fetching signature:", err);
      alert("Could not load signature-test.html");
    });
}
