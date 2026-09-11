const ENGINE_BASE_URL = "http://127.0.0.1:48652";

const ENGINE_STATE_URL =
    `${ENGINE_BASE_URL}/video/state`;

const ENGINE_BROWSER_COMMAND_URL =
    `${ENGINE_BASE_URL}/command/browser`;

const CIBLES_VEO = [
    "https://app.veo.co/*"
];

/*
 * Plusieurs onglets Veo peuvent etre ouverts en meme temps.
 * Sans arbitrage, ils envoient tous leur etat au moteur et
 * se volent mutuellement les commandes clavier : Excel se
 * retrouve alors connecte a la video d'un autre match.
 *
 * Un seul onglet est donc "pilote" : celui que l'utilisateur
 * a reellement sous les yeux (onglet actif de sa fenetre).
 */

let ongletPiloteId = null;

let dernierEtatPilote = 0;

let piloteEstActif = false;

let fenetreFocaliseeId = null;

const DELAI_PILOTE_SILENCIEUX = 3000;


chrome.windows.onFocusChanged.addListener(
    (windowId) => {
        if (windowId !== chrome.windows.WINDOW_ID_NONE) {
            fenetreFocaliseeId = windowId;
        }
    }
);

chrome.windows.getLastFocused(
    {},
    (fenetre) => {
        if (chrome.runtime.lastError) {
            return;
        }

        if (fenetre) {
            fenetreFocaliseeId = fenetre.id;
        }
    }
);

chrome.tabs.onRemoved.addListener(
    (tabId) => {
        if (tabId === ongletPiloteId) {
            ongletPiloteId = null;
            piloteEstActif = false;
        }
    }
);


function estOngletPilote(sender) {
    const onglet = sender?.tab;

    if (!onglet) {
        return false;
    }

    /*
     * Un onglet en arriere-plan ne prend jamais la main, mais
     * celui qui pilote deja la garde : consulter un autre
     * onglet ne doit pas couper la connexion avec Excel.
     */
    if (!onglet.active) {
        if (onglet.id === ongletPiloteId) {
            dernierEtatPilote = Date.now();
            piloteEstActif = false;
            return true;
        }

        return false;
    }

    const maintenant = Date.now();

    if (ongletPiloteId === onglet.id) {
        dernierEtatPilote = maintenant;
        piloteEstActif = true;
        return true;
    }

    const piloteSilencieux =
        maintenant - dernierEtatPilote
        > DELAI_PILOTE_SILENCIEUX;

    const dansFenetreFocalisee =
        fenetreFocaliseeId === null
        || onglet.windowId === fenetreFocaliseeId;

    /*
     * Un onglet que l'utilisateur regarde reprend toujours la
     * main sur un pilote passe en arriere-plan.
     */
    if (
        ongletPiloteId === null
        || piloteSilencieux
        || !piloteEstActif
        || dansFenetreFocalisee
    ) {
        const changementDePilote =
            ongletPiloteId !== null
            && ongletPiloteId !== onglet.id;

        ongletPiloteId = onglet.id;
        dernierEtatPilote = maintenant;
        piloteEstActif = true;

        if (changementDePilote) {
            viderCommandesEnAttente();
        }

        return true;
    }

    return false;
}


/*
 * Lors d'un changement d'onglet pilote, les commandes encore
 * en file etaient destinees a la video precedente : on les
 * jette pour ne pas les appliquer au mauvais match.
 */
async function viderCommandesEnAttente() {
    for (let essai = 0; essai < 10; essai += 1) {
        try {
            const reponse = await fetch(
                ENGINE_BROWSER_COMMAND_URL
            );

            if (!reponse.ok) {
                return;
            }

            const donnees = await reponse.json();

            if (!donnees.command) {
                return;
            }

        } catch (error) {
            return;
        }
    }
}


/*
 * Chrome n'injecte les content scripts que dans les pages
 * chargees APRES l'installation ou le rechargement de
 * l'extension. Les onglets Veo deja ouverts restaient donc
 * muets jusqu'a un rechargement manuel (F5).
 */
async function injecterDansOngletsExistants() {
    let onglets = [];

    try {
        onglets = await chrome.tabs.query({
            url: CIBLES_VEO
        });

    } catch (error) {
        console.log(
            "VeoVideoControl : onglets Veo introuvables —",
            error.message
        );

        return;
    }

    for (const onglet of onglets) {
        try {
            await chrome.scripting.executeScript({
                target: {
                    tabId: onglet.id
                },
                files: [
                    "content.js"
                ]
            });

            console.log(
                "VeoVideoControl : script injecte dans",
                onglet.url
            );

        } catch (error) {
            console.log(
                "VeoVideoControl : injection ignorée pour",
                onglet.url,
                "—",
                error.message
            );
        }
    }
}


chrome.runtime.onInstalled.addListener(
    injecterDansOngletsExistants
);

chrome.runtime.onStartup.addListener(
    injecterDansOngletsExistants
);


chrome.runtime.onMessage.addListener(
    (message, sender, sendResponse) => {

        if (message?.type === "VEOVIDEOCONTROL_VIDEO_STATE") {

            if (!estOngletPilote(sender)) {
                sendResponse({
                    ok: true,
                    ignore: true
                });

                return false;
            }

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

            // Seul l'onglet pilote execute les commandes.
            if (!estOngletPilote(sender)) {
                sendResponse({
                    ok: true,
                    command: null
                });

                return false;
            }

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
