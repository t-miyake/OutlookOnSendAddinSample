# OutlookOnSendAddinSample
This is a small sample add-in that displays a Dialog when sending email by Outlook on the web.

This add-in works fine with Outlook on the web version 20210419002.11 or earlier.

In version 20210913004.06, the Dialog will not be displayed and email cannot be sent.
However, if i change displayInIframe to false, it works fine. 
(In this case, the dialog is displayed in a separate window as expected.)