Office.onReady();

async function autoRunFunction(event) {

    Office.context.mailbox.item.body.setSignatureAsync(
        "<table><tr><td><span style='COLOR: #0000FF'>Test signature!</span> <span style='COLOR: #00FF00'>Message with this signature is always </span><span style='COLOR: #FF0000'>saved as draft.</span></td></tr></table>",
        {
            "coercionType": "html"
        },
        function (asyncResult) {
            event.completed();
        });
}
