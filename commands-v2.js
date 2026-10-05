Office.onReady();

function onItemSendHandler(event) {

    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function (getResult) {

            if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({
                    allowEvent: false,
                    errorMessage:
                        "TEST RESULT: GET BODY FAILED - " +
                        (getResult.error ? getResult.error.message : "unknown error")
                });
                return;
            }

            var marker =
                "<p><strong>[AUTOMATIC MODIFICATION TEST PASSED]</strong></p>";

            Office.context.mailbox.item.body.setAsync(
                getResult.value + marker,
                { coercionType: Office.CoercionType.Html },
                function (setResult) {

                    if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
                        event.completed({
                            allowEvent: false,
                            errorMessage:
                                "TEST RESULT: SET BODY FAILED - " +
                                (setResult.error ? setResult.error.message : "unknown error")
                        });
                        return;
                    }

                    event.completed({
                        allowEvent: true
                    });
                }
            );
        }
    );
}

Office.actions.associate(
    "onMessageSendHandler",
    onItemSendHandler
);
