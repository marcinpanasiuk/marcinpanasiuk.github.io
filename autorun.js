Office.onReady();

function autoRunFunction(event) {
    Office.context.mailbox.item.internetHeaders.setAsync(
        { "x-test-header": "true" },
        function (asyncResult) {
            let status = "Successfully set headers";
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                status = "Error setting headers: " + JSON.stringify(asyncResult.error);
            }
            Office.context.mailbox.item.body.setSignatureAsync(
                status,
                { coercionType: "html" },
                function () {
                    event.completed();
                }
            );
        }
    );
}
