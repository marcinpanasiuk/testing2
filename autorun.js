Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler(event) {

    Office.context.roamingSettings.saveAsync(function(asyncResult) {
        var status = 'OK';
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.error(asyncResult.error);
            status = 'Error occurred while saving roaming data, see console for details';
        }
        var signature = `<strong style='font-size: 20px;'> ${status} </strong>`;
        Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
    });
   
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);