Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler(event) {

    Office.context.roamingSettings.saveAsync(function(asyncResult) {
        if (asyncResult.status == AsyncResultStatus.Failed) {
            console.log(asyncResult.error);
            var status = asyncResult.error;
        }
        else {
            console.log('OK');
            status = 'OK';
        }
    });
    var signature = `<strong style='font-size: 25px;'> ${status.toString()} </strong>`;
    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);