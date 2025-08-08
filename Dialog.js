Office.onReady = function () {};

setTimeout(function () {
  Office.context.ui.messageParent("cancel");
}, 180000);

document.getElementById("sendButton").addEventListener("click", function () {
  Office.context.ui.messageParent("send");
});

document.getElementById("cancelButton").addEventListener("click", function () {
  Office.context.ui.messageParent("cancel");
});
