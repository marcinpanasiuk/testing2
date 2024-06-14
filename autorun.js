Office.onReady();

function onNewMessageComposeHandler(event) {
    let result = "no result";
    Office.auth.getAccessToken()
        .then(function (t) { result = `Success: ${t}`; })
        .catch(function (e) { result = `Fail: ${e.message}`; })
        .finally(function () { Office.context.mailbox.item.body.prependAsync(result, { coercionType: "html" }, function () { event.completed(); }) });
}

Office.actions?.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
