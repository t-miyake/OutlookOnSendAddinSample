Office.onReady = function () {};

setTimeout(function () {
  Office.context.ui.messageParent("cancel");
}, 180000);

$("#sendButton").click(function () {
  Office.context.ui.messageParent("send");
});

$("#cancelButton").click(function () {
  Office.context.ui.messageParent("cancel");
});
