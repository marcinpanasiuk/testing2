Office.onReady();

function onNewMessageComposeHandler(event) {
    var result = "no result";
    var startTime = performance.now();
    Office.auth.getAuthContext()
        .then(function (t) { result = `Success: ${t.userObjectId}`; })
        .catch(function (e) { result = `Fail: ${e.message}`; })
        .finally(function () { 
            var endTime = performance.now();
            Office.context.mailbox.item.body.setSignatureAsync(`${result}, it took ${Math.round(endTime - startTime)} ms`, { coercionType: "html" }, function () { event.completed(); }) });
}

Office.actions?.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
