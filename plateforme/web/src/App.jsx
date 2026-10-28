import { useCallback, useRef, useState } from "react";
import { useEtatCourant, useMatch } from "./etat/useMatch.js";
import { court } from "./etat/regles.js";
import { Saisie } from "./composants/Saisie.jsx";
import { Composition } from "./composants/Composition.jsx";
import { ParJoueur, Possession, TouchesMelees } from "./composants/Tableaux.jsx";
import { ListeMatchs } from "./composants/ListeMatchs.jsx";

export default function App() {
  const [matchId, setMatchId] = useState(null);

  if (!matchId) return <ListeMatchs surOuvrir={setMatchId} />;
  return <Match matchId={matchId} surFermer={() => setMatchId(null)} />;
}

function Match({ matchId, surFermer }) {
  const [vue, setVue] = useState("saisie");
  const [miTemps, setMiTemps] = useState("MT1");
  const [possession, setPossession] = useState("Nous");
  // La position arrondie a la seconde suffit aux regles et au
  // chrono. Sans cet arrondi, chaque battement du lecteur
  // redessinerait les palettes et le journal — ce qui rendait
  // justement la saisie poussive dans le classeur.
  const [position, setPosition] = useState(0);
  const lecteur = useRef(null);

  const majPosition = useCallback((t) => {
    setPosition((ancienne) =>
      Math.floor(t) === Math.floor(ancienne) ? ancienne : t
    );
  }, []);

  const donnees = useMatch(matchId);
  const etat = useEtatCourant(donnees, miTemps, position);

  const enregistrerLecteur = useCallback((element) => {
    lecteur.current = element;
  }, []);

  if (donnees.chargement) {
    return <p className="vide" style={{ padding: 40 }}>Chargement du match…</p>;
  }
  if (!donnees.match) {
    return (
      <div className="page">
        <div className="erreur">{donnees.erreur ?? "Match introuvable"}</div>
        <button type="button" className="btn-vide" onClick={surFermer}>
          Retour aux matchs
        </button>
      </div>
    );
  }

  const m = donnees.match;

  return (
    <div className="appli">
      <header className="bandeau">
        <button type="button" className="retour" onClick={surFermer} aria-label="Retour aux matchs">
          ←
        </button>
        <h1 className="titre-match">
          {m.adversaire}
          <small>
            {m.date_match} · {m.categorie || "—"} · {m.lieu}
            {m.journee ? ` · J${m.journee}` : ""}
          </small>
        </h1>

        <div className="vues" role="tablist">
          {[
            ["saisie", "Saisie"],
            ["compo", "Composition"],
            ["stats", "Statistiques"],
          ].map(([code, libelle]) => (
            <button
              key={code} type="button" role="tab"
              aria-selected={vue === code}
              onClick={() => setVue(code)}
            >
              {libelle}
            </button>
          ))}
        </div>

        <div className="bandeau-droite">
          <div className="groupe">
            <span className="etiq">Possession</span>
            <div className="seg">
              {["Nous", "Adv"].map((p) => (
                <button
                  key={p} type="button"
                  className={p === "Adv" ? "adv" : ""}
                  aria-pressed={possession === p}
                  onClick={() => setPossession(p)}
                >
                  {p === "Adv" ? "Adversaire" : "Nous"}
                </button>
              ))}
            </div>
          </div>

          <div className="groupe">
            <span className="etiq">Mi-temps</span>
            <div className="seg">
              {["MT1", "MT2"].map((t) => (
                <button
                  key={t} type="button"
                  aria-pressed={miTemps === t}
                  onClick={() => setMiTemps(t)}
                >
                  {t}
                </button>
              ))}
            </div>
          </div>

          <span className={"etat-jeu " + (etat.ouvert ? "ouvert" : "ferme")}>
            {etat.ouvert ? "Jeu ouvert" : "Reprise attendue"}
          </span>
          <span className="chrono">{court(position)}</span>
        </div>
      </header>

      {donnees.erreur && (
        <div className="erreur" style={{ margin: "0 0 8px" }}>
          {donnees.erreur}
        </div>
      )}

      <main>
        {vue === "saisie" && (
          <Saisie
            donnees={donnees}
            etat={etat}
            miTemps={miTemps}
            possession={possession}
            setPossession={setPossession}
            position={position}
            setPosition={majPosition}
            lecteur={lecteur}
            enregistrerLecteur={enregistrerLecteur}
            surComposition={() => setVue("compo")}
          />
        )}

        {vue === "compo" && (
          <Composition
            feuille={donnees.feuille}
            joueurs={donnees.joueurs}
            erreur={donnees.erreur}
            surEnregistrer={donnees.poserFeuille}
          />
        )}

        {vue === "stats" && (
          <Statistiques
            matchId={matchId}
            actions={donnees.actions}
            etat={etat}
            referentiels={donnees.referentiels}
            surSaut={(t) => {
              setVue("saisie");
              if (lecteur.current) {
                lecteur.current.currentTime = t;
                lecteur.current.pause();
              }
            }}
          />
        )}
      </main>
    </div>
  );
}

function Statistiques({ matchId, actions, etat, referentiels, surSaut }) {
  const [onglet, setOnglet] = useState("poss");

  return (
    <section className="vue">
      <div className="carte" style={{ flex: 1 }}>
        <div className="onglets" role="tablist">
          {[
            ["poss", "Possession"],
            ["joueurs", "Par joueur"],
            ["phases", "Touches et mêlées"],
          ].map(([code, libelle]) => (
            <button
              key={code} type="button" role="tab"
              aria-selected={onglet === code}
              onClick={() => setOnglet(code)}
            >
              {libelle}
            </button>
          ))}
        </div>
        <div className="panneau">
          {onglet === "poss" && <Possession matchId={matchId} rafraichir={actions.length} />}
          {onglet === "joueurs" && (
            <ParJoueur actions={actions} etat={etat} referentiels={referentiels} />
          )}
          {onglet === "phases" && (
            <TouchesMelees actions={actions} etat={etat} surSaut={surSaut} />
          )}
        </div>
      </div>
    </section>
  );
}
