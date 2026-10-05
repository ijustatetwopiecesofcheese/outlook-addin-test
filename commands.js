Office.onReady();

function onItemSendHandler(event) {
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }

            var marker =
                "<p><strong>[AUTOMATIC MODIFICATION TEST PASSED]</strong></p>";

            Office.context.mailbox.item.body.setAsync(
                result.value + marker,
                { coercionType: Office.CoercionType.Html },
                function () {
                    event.completed({ allowEvent: true });
                }
            );
        }
    );
}

Office.actions.associate(
    "onMessageSendHandler",
    onItemSendHandler
);
