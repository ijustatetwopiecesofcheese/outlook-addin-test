Office.onReady();

function ping(step) {
    fetch("./" + step + "?t=" + Date.now())
        .catch(function () {});
}

ping("GITHUB_JS_LOADED");

function onItemSendHandler(event) {
    ping("GITHUB_HANDLER_FIRED");

    event.completed({
        allowEvent: true
    });
}

Office.actions.associate(
    "onMessageSendHandler",
    onItemSendHandler
);

ping("GITHUB_HANDLER_ASSOCIATED");
