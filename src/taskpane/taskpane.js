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

function getSubject() {
  setOutput(Office.context.mailbox.item.subject);
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

function setCategory() {
  Office.context.mailbox.item.categories.addAsync(["Blue category"], function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setOutput(result.error);
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

function getAttachments() {
  for (let index = 0; index < Office.context.mailbox.item.attachments.length; index++) {
    Office.context.mailbox.item.getAttachmentContentAsync(
      Office.context.mailbox.item.attachments[index].id,
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          appendOutput(result.error);
        } else {
          appendOutput(result.value);

          let content = [];
          switch (result.value.format) {
            case Office.MailboxEnums.AttachmentContentFormat.Base64:
              content = [Uint8Array.from(atob(result.value.content), (c) => c.charCodeAt(0))];
              // Handle file attachment.
              console.log("Attachment is a Base64-encoded string.");
              break;
            case Office.MailboxEnums.AttachmentContentFormat.Eml:
              content = [new Uint8Array(new TextEncoder().encode(result.value.content))];
              break;
            case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
              // Handle .icalender attachment.
              console.log("Attachment is a calendar item.");
              break;
            case Office.MailboxEnums.AttachmentContentFormat.Url:
              // Handle cloud attachment.
              console.log("Attachment is a cloud attachment.");
              break;
            default:
            // Handle attachment formats that aren't supported.
          }
        }
      }
    );
  }
}

function getItemClass() {
  setOutput(Office.context.mailbox.item.itemClass);
}

function appendOutput(output) {
  document.getElementById("output").innerText += JSON.stringify(output);
}

function setOutput(output) {
  document.getElementById("output").innerText = JSON.stringify(output);
}

Office.onReady(() => {
  document.getElementById("tokenButton").onclick = getAccessToken;
  document.getElementById("selectedItemsButton").onclick = getSelectedItems;
  document.getElementById("getSubjectButton").onclick = getSubject;
  document.getElementById("itemIdButton").onclick = getItemId;
  document.getElementById("getMasterCategoriesButton").onclick = getMasterCategories;
  document.getElementById("setCategoryButton").onclick = setCategory;
  document.getElementById("saveSessionDataButton").onclick = saveSessionData;
  document.getElementById("getSessionDataButton").onclick = getSessionData;
  document.getElementById("getAttachmentsButton").onclick = getAttachments;
  document.getElementById("itemClassButton").onclick = getItemClass;

  Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, function (args) {
    window.location.reload();
  });

  document.getElementById("subject").innerText = Office.context.mailbox.item.subject;
  document.getElementById("itemId").innerText = Office.context.mailbox.item.itemId;
});
