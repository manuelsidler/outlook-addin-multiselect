Office.onReady()

function onMessageSendHandler(event) {
    console.log('Playground onMessageSendHandler')

    event.completed({ allowEvent: true });
  }
  
  // IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }