Office.onReady(() => {
      console.log("🚀 Launch Event: onNewMessageCompose triggered");
  Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onNewMessageCompose(event) {
  console.log("🚀 Launch Event: onNewMessageCompose triggered");
  Office.addin.showAsTaskpane();
  event.completed();
}

function onMessageSendHandler(event) {
  console.log("📤 Launch Event: onMessageSendHandler triggered");
  event.completed({ allowEvent: true });
}
