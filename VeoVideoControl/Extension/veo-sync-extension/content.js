console.log(
    "VeoVideoControl : extension chargée sur Veo ✅"
);

let videoInitialisee = null;


function extraireMatchId() {
    const match = window.location.pathname.match(
        /^\/matches\/([^/]+)/
    );

    return match ? match[1] : "";
}


function envoyerEtatVideo(video) {
    if (!video) {
        return;
    }

    const payload = {
        currentTime: Number(video.currentTime) || 0,
        paused: Boolean(video.paused),
        url: window.location.href,
        matchId: extraireMatchId(),
        source: "browser"
    };

    chrome.runtime.sendMessage(
        {
            type: "VEOVIDEOCONTROL_VIDEO_STATE",
            payload
        },
        (response) => {
            if (chrome.runtime.lastError) {
                const message =
                    chrome.runtime.lastError.message || "";

                if (
                    message.includes(
                        "Extension context invalidated"
                    )
                ) {
                    return;
                }

                console.error(
                    "VeoVideoControl — envoi de l’état impossible :",
                    message
                );

                return;
            }

            if (!response?.ok) {
                console.error(
                    "VeoVideoControl — état vidéo refusé :",
                    response?.error
                );
            }
        }
    );
}


function initialiserVideo(video) {
    if (videoInitialisee === video) {
        return;
    }

    videoInitialisee = video;

    console.log(
        "VeoVideoControl : vidéo détectée"
    );

    envoyerEtatVideo(video);

    video.addEventListener(
        "play",
        () => envoyerEtatVideo(video)
    );

    video.addEventListener(
        "pause",
        () => envoyerEtatVideo(video)
    );

    video.addEventListener(
        "seeked",
        () => envoyerEtatVideo(video)
    );

    /*
     * On limite les envois pendant la lecture.
     * timeupdate est déjà déclenché périodiquement
     * par le navigateur.
     */
    video.addEventListener(
        "timeupdate",
        () => envoyerEtatVideo(video)
    );
}


function attendreVideo() {
    const video = document.querySelector("video");

    if (video) {
        initialiserVideo(video);
    }

    setTimeout(attendreVideo, 500);
}

function naviguerActionVeo(direction) {
    const selecteur =
        direction === "previous"
            ? "button.btn.jump.prev"
            : "button.btn.jump.next";

    const bouton = document.querySelector(selecteur);

    if (!bouton) {
        console.error(
            `VeoVideoControl : bouton d’action ${direction} introuvable`
        );
        return;
    }

    bouton.click();

    console.log(
        `VeoVideoControl : navigation ${direction} exécutée`
    );
}


async function executerCommande(command) {
    const video = document.querySelector("video");

    if (!video) {
        console.error(
            "VeoVideoControl : aucune vidéo disponible"
        );

        return;
    }

    switch (command) {
        case "playpause":
            if (video.paused) {
                try {
                    await video.play();
                } catch (error) {
                    console.error(
                        "VeoVideoControl : lecture impossible",
                        error
                    );
                }
            } else {
                video.pause();
            }
            break;

        case "seek_minus_5":
            video.currentTime = Math.max(
                0,
                video.currentTime - 5
            );
            break;

        case "seek_plus_5":
            video.currentTime = Math.min(
                Number.isFinite(video.duration)
                    ? video.duration
                    : Number.POSITIVE_INFINITY,
                video.currentTime + 5
            );
            break;

        case "previous_action":
            naviguerActionVeo("previous");
            break;

        case "next_action":
            naviguerActionVeo("next");
            break;

        case "reset":
            video.pause();
            video.currentTime = 0;
            break;

        default:
            console.warn(
                "VeoVideoControl : commande inconnue",
                command
            );

            return;
    }

    /*
     * seeked/play/pause renverront aussi l’état,
     * mais cet envoi immédiat accélère la réponse.
     */
    envoyerEtatVideo(video);
}


function verifierCommandes() {
    if (!chrome.runtime?.id) {
        return;
    }

    chrome.runtime.sendMessage(
        {
            type: "VEOVIDEOCONTROL_GET_COMMAND"
        },
        (response) => {
            if (chrome.runtime.lastError) {
                const message =
                    chrome.runtime.lastError.message || "";

                if (
                    message.includes(
                        "Extension context invalidated"
                    )
                ) {
                    return;
                }

                console.error(
                    "VeoVideoControl — communication impossible :",
                    message
                );

                return;
            }

            if (!response?.ok) {
                console.error(
                    "VeoVideoControl — récupération impossible :",
                    response?.error
                );

                return;
            }

            if (!response.command) {
                return;
            }

            console.log(
                "VeoVideoControl : commande reçue",
                response.command
            );

            executerCommande(response.command);
        }
    );
}


attendreVideo();

/*
 * Deux interrogations par seconde.
 * Plus tard, on pourra remplacer ce polling par WebSocket.
 */
setInterval(verifierCommandes, 500);

setInterval(
    () => {
        const video = document.querySelector("video");

        if (video) {
            envoyerEtatVideo(video);
        }
    },
    1000
);