import { useCallback, useEffect, useRef, useState } from "react";
import { Separateur, Video } from "./Video.jsx";
import { PaletteActions, PaletteJoueurs } from "./Palettes.jsx";
import { Journal, TempsDeJeu } from "./Tableaux.jsx";
import {
  DialogueFaute, DialoguePhase, DialogueRemplacement, DialogueRetourCarton,
} from "./dialogues.jsx";
import { codeCarton, codeRetourCarton, inverse } from "../etat/regles.js";

/**
 * La vue de saisie : vidéo et journal à gauche, palettes à
 * droite sur toute la hauteur. Trois poignées pour ajuster.
 */
export function Saisie({ donnees, etat, miTemps, possession, setPossession,
  position, setPosition, lecteur, enregistrerLecteur, surComposition }) {

  const { actions, feuille, referentiels, types,
          ajouterAction, supprimerAction, ajouterRemplacement } = donnees;

  const [tailles, setTailles] = useState({
    largeur: 56, hauteurVideo: 60, hauteurJoueurs: 26,
  });
  const [onglet, setOnglet] = useState("journal");
  const [attente, setAttente] = useState(null);
  const [popup, setPopup] = useState(null);
  const [version, setVersion] = useState(0);

  // Les statistiques du bas se rafraîchissent après chaque
  // écriture : elles viennent du serveur.
  const toucher = useCallback(() => setVersion((v) => v + 1), []);

  // Le temps exact au moment du clic : la position affichee est
  // arrondie a la seconde pour la fluidite, elle ne doit pas
  // servir d'horodatage.
  const temps = useCallback(
    () => lecteur.current?.currentTime ?? position,
    [lecteur, position]
  );

  const sauter = useCallback(
    (t) => {
      if (lecteur.current) {
        lecteur.current.currentTime = t;
        lecteur.current.pause();
      }
    },
    [lecteur]
  );

  // ------------------------------------------------------
  // Écriture d'une action
  // ------------------------------------------------------

  const ecrire = useCallback(
    async (champs, suites = {}) => {
      const creee = await ajouterAction({
        temps: temps(),
        mt: miTemps,
        possession,
        numero: null,
        groupe: null,
        motif_code: null,
        commentaire: "",
        ...champs,
      });
      if (!creee) return null;

      // L'arrêt automatique suit l'action qui le déclenche, au
      // même instant : le jeu s'arrête là, pas une seconde plus
      // tard.
      if (suites.arret) {
        await ajouterAction({
          temps: champs.temps ?? temps(), mt: miTemps,
          possession: suites.possessionApres ?? possession,
          type_code: "ARRET", numero: null, groupe: null,
          motif_code: null, commentaire: "",
        });
      }
      if (suites.possessionApres) setPossession(suites.possessionApres);
      toucher();
      return creee;
    },
    [ajouterAction, temps, miTemps, possession, setPossession, toucher]
  );

  // ------------------------------------------------------
  // Clic sur une action de la palette
  // ------------------------------------------------------

  function clicAction(item) {
    // Un geste à double libellé : le code dépend de qui tient
    // le ballon au moment du clic.
    if (item.rec) {
      const code = possession === "Adv" ? item.rec : item.perte;
      setAttente({
        libelle: item.libelle, type_code: code, temps: temps(),
        bascule: true,
      });
      return;
    }

    const t = item;
    if (t.famille === "PENALITE" || t.famille === "COUP_FRANC") {
      setPopup({ sorte: "faute", type: t, temps: temps() });
      return;
    }
    if (t.attend_joueur) {
      setAttente({
        libelle: t.libelle, type_code: t.code, temps: temps(),
        bascule: t.bascule, arret: t.arret_auto,
      });
      return;
    }
    ecrire(
      { type_code: t.code, temps: temps() },
      {
        arret: t.arret_auto,
        possessionApres: t.bascule ? inverse(possession) : null,
      }
    );
  }

  // Le carton attend toujours le joueur : le camp se lit à lui,
  // jamais à la possession.
  function clicCarton(couleur) {
    setAttente({
      libelle: `Carton ${couleur}`, carton: couleur, temps: temps(),
    });
  }

  async function clicJoueur({ numero, groupe }) {
    if (!attente) return;
    const a = attente;
    setAttente(null);

    const type_code = a.carton
      ? codeCarton(a.carton, true) // un joueur désigné vaut notre camp
      : a.type_code;

    await ecrire(
      {
        type_code,
        temps: a.temps,
        numero: numero ?? null,
        groupe: groupe ?? null,
      },
      {
        arret: a.arret,
        possessionApres: a.bascule ? inverse(possession) : null,
      }
    );
  }

  // ------------------------------------------------------
  // Raccourcis clavier
  // ------------------------------------------------------

  useEffect(() => {
    function touche(e) {
      if (popup) return;
      const cible = e.target.tagName;
      if (/^(INPUT|SELECT|TEXTAREA|BUTTON)$/.test(cible)) return;
      const v = lecteur.current;
      if (!v?.src) return;

      if (e.code === "Space") {
        e.preventDefault();
        v.paused ? v.play() : v.pause();
      } else if (e.code === "ArrowLeft") {
        e.preventDefault();
        v.currentTime = Math.max(0, v.currentTime - 5);
      } else if (e.code === "ArrowRight") {
        e.preventDefault();
        v.currentTime += 5;
      }
    }
    document.addEventListener("keydown", touche);
    return () => document.removeEventListener("keydown", touche);
  }, [popup, lecteur]);

  // ------------------------------------------------------

  return (
    <section className="vue saisie">
      <div className="pan" style={{ width: `${tailles.largeur}%` }}>
        <div className="zone-fixe" style={{ height: `${tailles.hauteurVideo}%` }}>
          <Video surPosition={setPosition} enregistrerLecteur={enregistrerLecteur} />
        </div>

        <Separateur
          orientation="horizontal"
          valeur={tailles.hauteurVideo}
          min={24}
          max={86}
          libelle="Hauteur de la vidéo"
          surChangement={(v) => setTailles((t) => ({ ...t, hauteurVideo: v }))}
        />

        <div className="carte" style={{ flex: 1, minHeight: 0 }}>
          <div className="onglets" role="tablist">
            <button
              type="button" role="tab"
              aria-selected={onglet === "journal"}
              onClick={() => setOnglet("journal")}
            >
              Journal d'actions
            </button>
            <button
              type="button" role="tab"
              aria-selected={onglet === "tdj"}
              onClick={() => setOnglet("tdj")}
            >
              Temps de jeu
            </button>
            <span className="compte">
              {actions.length} action{actions.length > 1 ? "s" : ""}
            </span>
          </div>
          <div className="panneau">
            {onglet === "journal" ? (
              <Journal
                actions={actions}
                etat={etat}
                surSaut={sauter}
                surSuppression={async (id) => {
                  await supprimerAction(id);
                  toucher();
                }}
              />
            ) : (
              <TempsDeJeu matchId={donnees.match.id} rafraichir={version} />
            )}
          </div>
        </div>
      </div>

      <Separateur
        orientation="vertical"
        valeur={tailles.largeur}
        min={26}
        max={76}
        libelle="Largeur des palettes"
        surChangement={(v) => setTailles((t) => ({ ...t, largeur: v }))}
      />

      <div className="pan" style={{ width: `${100 - tailles.largeur}%` }}>
        <div className="zone-fixe" style={{ height: `${tailles.hauteurJoueurs}%` }}>
          <PaletteJoueurs
            etat={etat}
            attente={attente}
            surJoueur={clicJoueur}
            surAnnuler={() => setAttente(null)}
            surRemplacement={() => setPopup({ sorte: "remplacement", temps: temps() })}
            surComposition={surComposition}
          />
        </div>

        <Separateur
          orientation="horizontal"
          valeur={tailles.hauteurJoueurs}
          min={12}
          max={70}
          libelle="Hauteur de la palette joueurs"
          surChangement={(v) => setTailles((t) => ({ ...t, hauteurJoueurs: v }))}
        />

        <PaletteActions
          referentiels={referentiels}
          etat={etat}
          surAction={clicAction}
          surPhase={(famille, issue) =>
            setPopup({ sorte: "phase", famille, issue, temps: temps() })
          }
          surCarton={clicCarton}
          surRetourCarton={() => setPopup({ sorte: "retour", temps: temps() })}
        />
      </div>

      {/* ---------------- popups ---------------- */}

      {popup?.sorte === "phase" && (
        <DialoguePhase
          famille={popup.famille}
          issue={popup.issue}
          etat={etat}
          possession={possession}
          temps={popup.temps}
          surFermer={() => setPopup(null)}
          surValider={async (champs) => {
            const { possessionApres, ...reste } = champs;
            await ecrire({ ...reste, temps: popup.temps }, { possessionApres });
            setPopup(null);
          }}
        />
      )}

      {popup?.sorte === "faute" && (
        <DialogueFaute
          type={popup.type}
          motifs={referentiels?.motifs ?? []}
          etat={etat}
          temps={popup.temps}
          surFermer={() => setPopup(null)}
          surValider={async (champs) => {
            const { possessionApres, ...reste } = champs;
            // Pénalité comme coup franc arrêtent le jeu.
            await ecrire(
              { ...reste, temps: popup.temps },
              { arret: true, possessionApres }
            );
            setPopup(null);
          }}
        />
      )}

      {popup?.sorte === "remplacement" && (
        <DialogueRemplacement
          etat={etat}
          feuille={feuille}
          temps={popup.temps}
          mt={miTemps}
          surFermer={() => setPopup(null)}
          surValider={async (r) => {
            await ajouterRemplacement(r);
            toucher();
            setPopup(null);
          }}
        />
      )}

      {popup?.sorte === "retour" && (
        <DialogueRetourCarton
          etat={etat}
          temps={popup.temps}
          surFermer={() => setPopup(null)}
          surValider={async ({ numero, couleur }) => {
            await ecrire({
              type_code: codeRetourCarton(couleur, true),
              temps: popup.temps,
              numero,
            });
            setPopup(null);
          }}
        />
      )}
    </section>
  );
}
