Office.initialize = function (reason) { };

function onNewMessageComposeHandler(event) {
    
    Office.context.roamingSettings.set("test-key", "test-value");

    Office.context.roamingSettings.saveAsync(function(asyncResult) {      
        var xmlhttp = new XMLHttpRequest();

        xmlhttp.onload = () => {
            insertDebugSignature("CORS OK", event);
        };

        xmlhttp.onerror = () => {
            insertDebugSignature(xmlhttp.responseText, event)
        }

        xmlhttp.open("POST", 'https://esig365lane1ring2-public-api.esig365.net/api/v2/clientsignatures/default');
        xmlhttp.setRequestHeader("Content-Type", "application/json");
        xmlhttp.setRequestHeader("Authorization", "Bearer test-token");
        xmlhttp.send(JSON.stringify({test: "value"}));
    });
}

function insertDebugSignature(text, event) {
    var signature = `<strong style='font-size: 16px;'> ${text} </strong>`;
    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);