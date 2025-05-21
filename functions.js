Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {}

});

function onMessageSendHandler(event) {
  const approvedDomains = ["fujimi.com", "fujimiinc.co.jp", "fujimiinc.com.tw", "fujimi.com.my", "fujimieurope.de"];

  Office.context.mailbox.item.getRecipientsAsync("all", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;
      const allEmails = [...recipients.to, ...recipients.cc, ...recipients.bcc].map(r => r.emailAddress.toLowerCase());

      const externalEmails = allEmails.filter(email => {
        const domain = email.split("@")[1];
        return !approvedDomains.includes(domain);
      });

      if (externalEmails.length > 0) {
        Office.context.ui.displayDialogAsync(
          "https://bsmith-FC.github.io/email-recipient-check/confirm.html",
          { height: 40, width: 30, displayInIframe: true },
          function (dialogResult) {
            const dialog = dialogResult.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
              if (arg.message === "Send") {
                dialog.close();
                event.completed({ allowEvent: true });
              } else {
                dialog.close();
                event.completed({ allowEvent: false });
              }
            });

            dialog.addEventHandler(Office.EventType.DialogReady, function () {
              dialog.messageParent(JSON.stringify(externalEmails));
            });

            dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
              dialog.close();
              event.completed({ allowEvent: false });
            });
          }
        );
      } else {
        event.completed({ allowEvent: true });
      }
    } else {
      console.error("Failed to get recipients.");
      event.completed({ allowEvent: false });
    }
  });
}
