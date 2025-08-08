"use strict";

const dialogPath = "https://www.noraneko.co.jp/outlookonsendaddinsample/dialog.html";
let mailItem;

Office.onReady().then(function () {
  mailItem = Office.context.mailbox.item;
});

function messageOnSent(event) {
  const thisEvent = event;

  if (Office.context.mailbox.diagnostics.hostName !== "OutlookWebApp") {
    thisEvent.completed({
      allowEvent: true,
    });
    return;
  }
  if (mailItem.itemType !== Office.MailboxEnums.ItemType.Message) {
    thisEvent.completed({
      allowEvent: true,
    });
    return;
  }

  Office.context.ui.displayDialogAsync(dialogPath, { height: 50, width: 50, promptBeforeOpen: false, displayInIframe: true }, function (asyncResult) {
    const dialog = asyncResult.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
      dialog.close();
      if (arg.message === "send") {
        thisEvent.completed({
          allowEvent: true,
        });
      } else {
        thisEvent.completed({
          allowEvent: false,
        });
      }
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
      if (arg.error >= 12002) {
        thisEvent.completed({
          allowEvent: false,
        });
      }
    });
  });
}
