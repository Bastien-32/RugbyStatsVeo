import { useState } from "react";
import { Choix, Dialogue } from "./Dialogue.jsx";
import {
  ALIGNEMENT, BALLON, UTILISATION,
  ZONE_LARGEUR, ZONE_LONGUEUR, ZONE_SAUT,
} from "../etat/disposition.js";
import { codePhase, mmss, possessionApres } from "../etat/regles.js";

/**
 * Touche ou mêlée.
 *
 * Le détail est facultatif : en vidéo on ne distingue pas
 * toujours tout, et mieux vaut une information partielle qu'une
 * information inventée. Seule l'issue est décidée par le bouton
 * cliqué.
 */
export function DialoguePhase({ famille, issue, etat, possession, temps,
  surValider, surFermer }) {

  const estTouche = famille === "TOUCHE";
  const [sauteur, setSauteur] = useState(null);
  const [autreVisible, setAutreVisible] = useState(false);
  const [saut, setSaut] = useState(null);
  const [alignement, setAlignement] = useState(null);
  const [ballon, setBallon] = useState(null);
  const [longueur, setLongueur] = useState(null);
  const [largeur, setLargeur] = useState(null);
  const [utilisation, setUtilisation] = useState(null);

  // Les sauteurs proposés sont les avants présents. « Autre »
  // déplie les trois-quarts : des équipes en font sauter un,
  // sans que la popup ait à montrer tout le monde.
  const avants = etat.presents
    .filter((n) => etat.ligneDe(n) === "Avant")
    .map((n) => ({ code: String(n), libelle: etat.etiquette(n) }));
  const autres = etat.presents
    .filter((n) => etat.ligneDe(n) !== "Avant")
    .map((n) => ({ code: String(n), libelle: etat.etiquette(n) }));

  return (
    <Dialogue
      titre={estTouche ? "Touche" : "Mêlée"}
      sous={`Issue ${issue} · possession ${possession} · ${mmss(temps)}`}
      surFermeture={surFermer}
      surValidation={() => {
        surValider({
          type_code: codePhase(famille, possession, issue),
          numero: sauteur ? Number(sauteur) : null,
          detail: {
            issue,
            sauteur_numero: sauteur ? Number(sauteur) : null,
            zone_saut: saut ?? "",
            alignement: alignement ?? "",
            ballon: ballon ?? "",
            zone_longueur: longueur ?? "",
            zone_largeur: largeur ?? "",
            utilisation: utilisation ?? "",
            observation: "",
          },
          possessionApres: possessionApres(issue, possession),
        });
      }}
    >
      {estTouche && (
        <>
          <Choix
            libelle="Sauteur"
            options={[...avants, { code: "AUTRE", libelle: "Autre" }]}
            valeur={autreVisible ? "AUTRE" : sauteur}
            surChoix={(v) => {
              if (v === "AUTRE") {
                setAutreVisible(true);
                setSauteur(null);
              } else {
                setAutreVisible(false);
                setSauteur(v);
              }
            }}
          />
          {autreVisible && (
            <Choix
              libelle="Qui a sauté"
              options={autres}
              valeur={sauteur}
              surChoix={setSauteur}
            />
          )}
          <Choix libelle="Zone de saut" options={ZONE_SAUT}
            valeur={saut} surChoix={setSaut} />
          <Choix libelle="Alignement" options={ALIGNEMENT}
            valeur={alignement} surChoix={setAlignement} />
          <Choix libelle="Ballon" options={BALLON}
            valeur={ballon} surChoix={setBallon} />
        </>
      )}
      <Choix libelle="Zone (longueur)" options={ZONE_LONGUEUR}
        valeur={longueur} surChoix={setLongueur} />
      <Choix libelle="Zone (largeur)" options={ZONE_LARGEUR}
        valeur={largeur} surChoix={setLargeur} />
      <Choix libelle="Utilisation" options={UTILISATION}
        valeur={utilisation} surChoix={setUtilisation} />
    </Dialogue>
  );
}

/**
 * Pénalité ou coup franc.
 *
 * Les motifs viennent de la base, rangés par famille. Celui qui
 * les ajoute n'a pas à toucher au code.
 */
export function DialogueFaute({ type, motifs, etat, temps, surValider, surFermer }) {
  const [motif, setMotif] = useState(null);
  const [fautif, setFautif] = useState(null);

  const estCF = type.code.startsWith("CF_");
  const pour = type.code.endsWith("_ADV") ? "Nous" : "Adv";

  const retenus = motifs.filter((m) => (estCF ? m.coup_franc : m.penalite));
  const familles = {};
  for (const m of retenus) (familles[m.famille] ||= []).push(m);

  const fautifs = [
    ...etat.presents.map((n) => ({ code: String(n), libelle: etat.etiquette(n) })),
    { code: "Avant", libelle: "Avant" },
    { code: "Trois-quarts", libelle: "Trois-quarts" },
  ];

  return (
    <Dialogue
      titre={estCF ? "Coup franc" : "Pénalité"}
      sous={`${type.libelle} · ${mmss(temps)}`}
      surFermeture={surFermer}
      surValidation={() => {
        const estGroupe = fautif === "Avant" || fautif === "Trois-quarts";
        surValider({
          type_code: type.code,
          motif_code: motif,
          // Le fautif n'est demandé que pour une faute de notre
          // camp : nous ne nommons pas les joueurs adverses.
          numero: pour === "Adv" && fautif && !estGroupe ? Number(fautif) : null,
          groupe: pour === "Adv" && estGroupe ? fautif : null,
          possessionApres: pour,
        });
      }}
    >
      <div className="champ">
        <label>Motif</label>
        {Object.entries(familles).map(([nom, liste]) => (
          <div key={nom} className="motifs-famille">
            <h4>{nom}</h4>
            <div className="choix">
              {liste.map((m) => (
                <button
                  key={m.code}
                  type="button"
                  aria-pressed={motif === m.code}
                  title={m.libelle}
                  onClick={() => setMotif(motif === m.code ? null : m.code)}
                >
                  {m.libelle.includes(" - ")
                    ? m.libelle.slice(m.libelle.indexOf(" - ") + 3)
                    : m.libelle}
                </button>
              ))}
            </div>
          </div>
        ))}
      </div>

      {pour === "Adv" && (
        <Choix libelle="Fautif (notre camp)" options={fautifs}
          valeur={fautif} surChoix={setFautif} />
      )}
    </Dialogue>
  );
}

/**
 * Remplacement.
 *
 * Entrent tous ceux de la feuille qui ne sont pas sur le
 * terrain : le banc, mais aussi un joueur déjà sorti, qui peut
 * revenir.
 */
export function DialogueRemplacement({ etat, feuille, temps, mt, surValider, surFermer }) {
  const [sortant, setSortant] = useState(null);
  const [entrant, setEntrant] = useState(null);

  const sortants = etat.presents.map((n) => ({
    code: String(n), libelle: etat.etiquette(n),
  }));
  const entrants = feuille
    .filter((l) => !etat.presents.includes(l.numero))
    .map((l) => ({ code: String(l.numero), libelle: `${l.numero} ${l.nom_complet}` }));

  if (entrants.length === 0) {
    return (
      <Dialogue
        titre="Remplacement"
        sous="Personne à faire entrer"
        surFermeture={surFermer}
        surValidation={() => true}
      >
        <p className="note">
          Tous les joueurs de la feuille sont sur le terrain. Complétez le
          banc depuis la vue Composition.
        </p>
      </Dialogue>
    );
  }

  return (
    <Dialogue
      titre="Remplacement"
      sous={`Le remplaçant entre avec son propre numéro · ${mmss(temps)}`}
      surFermeture={surFermer}
      surValidation={() => {
        if (!sortant || !entrant) return false;
        surValider({
          temps, mt,
          sortant: Number(sortant),
          entrant: Number(entrant),
          temporaire: false,
        });
      }}
    >
      <Choix libelle="Sort" options={sortants} valeur={sortant} surChoix={setSortant} />
      <Choix libelle="Entre" options={entrants} valeur={entrant} surChoix={setEntrant} />
    </Dialogue>
  );
}

/** Retour d'un joueur sous carton. */
export function DialogueRetourCarton({ etat, temps, surValider, surFermer }) {
  const [qui, setQui] = useState(null);

  const options = Object.entries(etat.cartons).map(([n, couleur]) => ({
    code: n,
    libelle: `${etat.etiquette(Number(n))} — ${couleur}`,
  }));

  return (
    <Dialogue
      titre="Retour de carton"
      sous={`Le joueur revient sur le terrain · ${mmss(temps)}`}
      surFermeture={surFermer}
      surValidation={() => {
        if (!qui) return false;
        surValider({ numero: Number(qui), couleur: etat.cartons[qui] });
      }}
    >
      <Choix libelle="Joueur" options={options} valeur={qui} surChoix={setQui} />
    </Dialogue>
  );
}
