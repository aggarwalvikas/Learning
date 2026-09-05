/* global Office */
(function () {
  "use strict";

  function setText(elementId, value) {
    var element = document.getElementById(elementId);
    if (element) {
      element.textContent = value;
    }
  }

  function renderMessageDetails() {
    var item = Office.context.mailbox && Office.context.mailbox.item;

    if (!item) {
      setText("sender", "No message selected");
      setText("subject", "No message selected");
      setText("status", "Open a message in read mode to view its subject.");
      return;
    }

    var sender = item.from;
    var senderName = sender && sender.displayName;
    var senderAddress = sender && sender.emailAddress;

    if (senderName && senderAddress) {
      setText("sender", senderName + " <" + senderAddress + ">");
    } else if (senderName || senderAddress) {
      setText("sender", senderName || senderAddress);
    } else {
      setText("sender", "Unknown sender");
    }

    if (typeof item.subject !== "string" || item.subject.length === 0) {
      setText("subject", "No subject");
      setText("status", "Sender and subject loaded from the selected message.");
      return;
    }

    setText("subject", item.subject);
    setText("status", "Sender and subject loaded from the selected message.");
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      if (info.host === Office.HostType.Outlook) {
        renderMessageDetails();
      } else {
        setText("subject", "Outlook is required");
        setText("status", "This add-in must be opened from Outlook.");
      }
    });
  } else {
    Office.initialize = renderMessageDetails;
  }
})();
