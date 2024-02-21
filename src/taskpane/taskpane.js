Office.onReady(() => {
  document.getElementById("tokenButton").onclick = getAccessToken;
  document.getElementById("selectedItemsButton").onclick = getSelectedItems;
  document.getElementById("itemIdButton").onclick = getItemId;
  document.getElementById("getMasterCategoriesButton").onclick = getMasterCategories;
  document.getElementById("saveSessionDataButton").onclick = saveSessionData;
  document.getElementById("getSessionDataButton").onclick = getSessionData;
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

function getItemId() {
  setOutput(Office.context.mailbox.item.itemId ?? "undefined");
}

function getMasterCategories() {
  Office.context.mailbox.masterCategories.getAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setOutput(result.error);
    } else {
      setOutput(result.value);
    }
  });
}

function saveSessionData() {
  Office.context.mailbox.item.sessionData.setAsync("test", "value", function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setOutput(result.error);
    }
  });
}

function getSessionData() {
  Office.context.mailbox.item.sessionData.getAsync("test", function (result) {
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
