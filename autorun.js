Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler(event) {
    var xmlhttp = new XMLHttpRequest();
    
    xmlhttp.onload = function() {
        var signature = `<strong style='font-size: 20px; font-color: greeen'>Success: ${xmlhttp.responseText} </strong>`;
        Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
    }
    
    xmlhttp.onerror = function() {
        var signature = `<strong style='font-size: 20px; font-color: red'>Error: ${xmlhttp.responseText} </strong>`;
        Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
    }
    
    xmlhttp.open('GET', 'https://marcinpanasiuk.github.io/testing2/api/get', true);    
    xmlhttp.send();
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
