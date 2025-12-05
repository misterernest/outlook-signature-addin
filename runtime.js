console.log("WWG runtime.js loaded");

function onNewMessageComposeHandler(event) {
  console.log("WWG EVENT: OnNewMessageCompose fired");

  event.completed();
}

Office.actions.associate(
  "onNewMessageComposeHandler",
  onNewMessageComposeHandler
);