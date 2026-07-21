const ENGINE_BASE_URL = "http://127.0.0.1:48652";

const ENGINE_STATE_URL =
    `${ENGINE_BASE_URL}/video/state`;

const ENGINE_BROWSER_COMMAND_URL =
    `${ENGINE_BASE_URL}/command/browser`;


chrome.runtime.onMessage.addListener(
    (message, sender, sendResponse) => {

        if (message?.type === "VEOVIDEOCONTROL_VIDEO_STATE") {
            fetch(ENGINE_STATE_URL, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(message.payload)
            })
                .then(async (response) => {
                    if (!response.ok) {
                        throw new Error(
                            `Erreur HTTP ${response.status}`
                        );
                    }

                    return response.json();
                })
                .then((data) => {
                    sendResponse({
                        ok: true,
                        data
                    });
                })
                .catch((error) => {
                    sendResponse({
                        ok: false,
                        error: error.message
                    });
                });

            return true;
        }

        if (message?.type === "VEOVIDEOCONTROL_GET_COMMAND") {
            fetch(ENGINE_BROWSER_COMMAND_URL)
                .then(async (response) => {
                    if (!response.ok) {
                        throw new Error(
                            `Erreur HTTP ${response.status}`
                        );
                    }

                    return response.json();
                })
                .then((data) => {
                    sendResponse({
                        ok: true,
                        command: data.command
                    });
                })
                .catch((error) => {
                    sendResponse({
                        ok: false,
                        error: error.message
                    });
                });

            return true;
        }

        return false;
    }
);