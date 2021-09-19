const dialogPath = "https://www.noraneko.co.jp/outlookonsendaddinsample/Dialog.html";

Office.onReady = function () { }

Office.initialize = function () {
    $(document).ready(function () {
        try {
            if (Office.context.mailbox.diagnostics.hostName !== "OutlookWebApp") {
                return;
            }
        } catch (e) {
            return;
        }

        mailboxItem = Office.context.mailbox.item;
    });
}

function messageOnSent(event) {
    if (Office.context.mailbox.diagnostics.hostName !== "OutlookWebApp") {
        event.completed({
            allowEvent: true
        });
        return;
    }
    if (mailboxItem.itemType !== Office.MailboxEnums.ItemType.Message) {
        event.completed({
            allowEvent: true
        });
        return;
    }

    Office.context.ui.displayDialogAsync(dialogPath, { height: 50, width: 50, promptBeforeOpen: false, displayInIframe: true },
        function (asyncResult) {
            dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                dialog.close();
                if (arg.message === "send") {
                    event.completed({
                        allowEvent: true
                    });
                } else {
                    event.completed({
                        allowEvent: false
                    });
                }
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
                if (arg.error >= 12002) {
                    event.completed({
                        allowEvent: false
                    });
                }
            });
        }
    );
}