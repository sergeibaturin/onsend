function onMessageSendHandler(event) {
    event.completed({allowEvent: true});
}

function getGlobal() {
    return typeof self !== 'undefined'
        ? self
        : typeof window !== 'undefined'
            ? window
            : typeof global !== 'undefined'
                ? global
                : undefined;
}

const g = getGlobal();
g.onMessageSendHandler = onMessageSendHandler;

Office.onReady(() => {});