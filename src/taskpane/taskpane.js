Office.onReady(() => {
  document.getElementById("tokenButton").onclick = getAccessToken;
  document.getElementById("selectedItemsButton").onclick = getSelectedItems;
});

function getAccessToken() {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setOutput(result.error);
    } else {
      setOutput(result.value);
    }
  });
}

function getSelectedItems() {
  Office.context.mailbox.getSelectedItemsAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setOutput(result.error);
    } else {
      setOutput(result.value);
    }
  });
}

function setOutput(output) {
  document.getElementById("output").innerText = JSON.stringify(output);
}
