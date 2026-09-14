/*
 * Le script peut etre injecte deux fois sur la meme page :
 * une fois automatiquement par Chrome, une fois par le
 * service worker pour les onglets deja ouverts. On evite
 * alors de dedoubler les minuteries.
 */
if (window.__veoVideoControlCharge) {

    console.log(
        "VeoVideoControl : extension deja active sur cet onglet"
    );

} else {

    window.__veoVideoControlCharge = true;

    demarrerVeoVideoControl();
}


function demarrerVeoVideoControl() {

console.log(
    "VeoVideoControl : extension chargée sur Veo ✅"
);

let videoInitialisee = null;

let urlCourante = window.location.href;

const minuteries = [];


/*
 * Lecture arriere.
 *
 * Aucun navigateur ne lit une video a l'envers : un
 * playbackRate negatif est ignore. On recule donc
 * currentTime par petits pas a intervalle regulier, ce qui
 * donne un defilement arriere continu.
 *
 * 0,2 s toutes les 100 ms = vitesse x2 en arriere.
 */
const REWIND_PAS_SECONDES = 0.2;

const REWIND_INTERVALLE_MS = 100;

/*
 * Chaque saut declenche un evenement seeked, donc un envoi
 * d'etat. A dix sauts par seconde, on inonderait le moteur :
 * pendant le rewind, on limite les envois a quatre par
 * seconde, ce qui suffit au chrono d'Excel.
 */
const REWIND_ETAT_INTERVALLE_MS = 250;

let rewindMinuterie = null;

let rewindDernierEtatEnvoye = 0;


/*
 * Quand l'extension est rechargee (page « Extensions » de
 * Chrome), les scripts deja en place deviennent orphelins :
 * plus aucun message ne passe. On arrete alors proprement et
 * on previent dans la console au lieu d'echouer en silence.
 */
function contexteValide() {
    return Boolean(chrome.runtime?.id);
}


function arreterToutesLesMinuteries() {
    while (minuteries.length > 0) {
        clearInterval(minuteries.pop());
    }
}


let moteurInjoignableSignale = false;


/*
 * Le moteur n'est lance qu'avec Excel : son absence est un
 * etat normal, pas une erreur. On l'annonce une seule fois,
 * et en console.log pour ne pas alimenter le gestionnaire
 * d'extensions de Chrome.
 */
function signalerMoteurInjoignable(message) {
    if (moteurInjoignableSignale) {
        return;
    }

    moteurInjoignableSignale = true;

    console.log(
        "VeoVideoControl : moteur injoignable " +
        "(Excel est-il ouvert ?) —",
        message
    );
}


function signalerMoteurJoignable() {
    if (!moteurInjoignableSignale) {
        return;
    }

    moteurInjoignableSignale = false;

    console.log(
        "VeoVideoControl : liaison avec le moteur rétablie"
    );
}


function signalerContexteInvalide() {
    arreterToutesLesMinuteries();

    console.log(
        "VeoVideoControl : l’extension a été rechargée. " +
        "Rechargez cette page (Cmd+R) pour rétablir la " +
        "connexion avec Excel."
    );
}


/*
 * sendMessage leve une exception synchrone lorsque
 * l'extension vient d'etre rechargee : sans cette
 * protection, l'erreur remonte dans le gestionnaire
 * d'extensions de Chrome.
 */
function envoyerMessage(message, rappel) {
    try {
        chrome.runtime.sendMessage(message, rappel);

    } catch (error) {
        signalerContexteInvalide();
    }
}


function extraireMatchId() {
    const match = window.location.pathname.match(
        /^\/matches\/([^/]+)/
    );

    return match ? match[1] : "";
}


/*
 * Les pages Veo peuvent contenir plusieurs balises video
 * (apercus, vignettes). On retient celle qui porte une vraie
 * duree de match, sinon la premiere trouvee.
 */
function trouverVideo() {
    const videos = Array.from(
        document.querySelectorAll("video")
    );

    if (videos.length === 0) {
        return null;
    }

    let meilleure = null;

    for (const video of videos) {
        if (!Number.isFinite(video.duration)) {
            continue;
        }

        if (
            meilleure === null
            || video.duration > meilleure.duration
        ) {
            meilleure = video;
        }
    }

    return meilleure || videos[0];
}


function envoyerEtatVideo(video) {
    if (!video) {
        return;
    }

    if (!contexteValide()) {
        signalerContexteInvalide();
        return;
    }

    if (rewindMinuterie !== null) {
        const maintenant = Date.now();

        if (
            maintenant - rewindDernierEtatEnvoye
            < REWIND_ETAT_INTERVALLE_MS
        ) {
            return;
        }

        rewindDernierEtatEnvoye = maintenant;
    }

    const payload = {
        currentTime: Number(video.currentTime) || 0,
        paused: Boolean(video.paused),
        url: window.location.href,
        matchId: extraireMatchId(),
        source: "browser"
    };

    envoyerMessage(
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
                    signalerContexteInvalide();
                    return;
                }

                signalerMoteurInjoignable(message);

                return;
            }

            if (!response?.ok) {
                signalerMoteurInjoignable(response?.error);
                return;
            }

            signalerMoteurJoignable();
        }
    );
}


function initialiserVideo(video) {
    if (videoInitialisee === video) {
        return;
    }

    videoInitialisee = video;

    console.log(
        "VeoVideoControl : vidéo détectée",
        extraireMatchId()
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


/*
 * Veo est une application monopage : passer d'un match a
 * l'autre ne recharge pas le document. Il faut donc relacher
 * la video precedente pour raccrocher la nouvelle.
 */
function surveillerChangementDeMatch() {
    if (window.location.href === urlCourante) {
        return;
    }

    urlCourante = window.location.href;

    videoInitialisee = null;

    console.log(
        "VeoVideoControl : changement de match détecté",
        extraireMatchId()
    );
}


function attendreVideo() {
    if (!contexteValide()) {
        signalerContexteInvalide();
        return;
    }

    surveillerChangementDeMatch();

    const video = trouverVideo();

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
        console.log(
            `VeoVideoControl : bouton d’action ${direction} introuvable`
        );
        return;
    }

    bouton.click();

    console.log(
        `VeoVideoControl : navigation ${direction} exécutée`
    );
}


function rewindActif() {
    return rewindMinuterie !== null;
}


function arreterRewind() {
    if (rewindMinuterie === null) {
        return;
    }

    clearInterval(rewindMinuterie);

    rewindMinuterie = null;

    console.log("VeoVideoControl : lecture arrière arrêtée");
}


function demarrerRewind(video) {
    if (rewindMinuterie !== null) {
        return;
    }

    /*
     * La lecture normale continuerait d'avancer entre deux
     * sauts : on met la video en pause pendant le rewind.
     */
    video.pause();

    rewindMinuterie = setInterval(
        () => {
            const videoCourante = trouverVideo();

            if (!videoCourante) {
                arreterRewind();
                return;
            }

            const nouveauTemps =
                videoCourante.currentTime
                - REWIND_PAS_SECONDES;

            if (nouveauTemps <= 0) {
                videoCourante.currentTime = 0;
                arreterRewind();
                return;
            }

            videoCourante.currentTime = nouveauTemps;
        },
        REWIND_INTERVALLE_MS
    );

    minuteries.push(rewindMinuterie);

    console.log("VeoVideoControl : lecture arrière démarrée");
}


async function executerCommande(command) {
    const video = trouverVideo();

    if (!video) {
        console.log(
            "VeoVideoControl : aucune vidéo disponible"
        );

        return;
    }

    /*
     * Toute commande de transport interrompt la lecture
     * arriere : play/pause, sauts de 5 s, reset, navigation
     * entre actions.
     */
    if (command !== "rewind_toggle") {
        arreterRewind();
    }

    switch (command) {
        case "rewind_toggle":
            if (rewindActif()) {
                arreterRewind();
            } else {
                demarrerRewind(video);
            }
            break;

        case "playpause":
            if (video.paused) {
                try {
                    await video.play();
                } catch (error) {
                    console.log(
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
    if (!contexteValide()) {
        signalerContexteInvalide();
        return;
    }

    // Une page Veo sans lecteur ne pilote rien.
    if (!trouverVideo()) {
        return;
    }

    envoyerMessage(
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
                    signalerContexteInvalide();
                    return;
                }

                signalerMoteurInjoignable(message);

                return;
            }

            if (!response?.ok) {
                signalerMoteurInjoignable(response?.error);

                return;
            }

            signalerMoteurJoignable();

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
minuteries.push(
    setInterval(verifierCommandes, 500)
);

minuteries.push(
    setInterval(
        () => {
            const video = trouverVideo();

            if (video) {
                envoyerEtatVideo(video);
            }
        },
        1000
    )
);

}
