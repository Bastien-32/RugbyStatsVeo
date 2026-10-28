import {
  ABREGE, COLONNES, COULEURS_CARTON, CODES_EN_DUO, DUOS,
  ISSUES_MELEE, ISSUES_TOUCHE,
} from "../etat/disposition.js";
import { GROUPES } from "../etat/regles.js";

/**
 * Les joueurs sur le terrain.
 *
 * Le bouton porte le numéro du maillot, pas celui du poste : un
 * remplaçant entre avec le sien. Un joueur sous carton garde sa
 * couleur et ne répond plus, tant que son retour n'est pas
 * déclaré.
 */
export function PaletteJoueurs({ etat, attente, surJoueur, surRemplacement,
  surAnnuler, surComposition }) {

  // Sans feuille de match, la palette n'a personne à proposer :
  // mieux vaut le dire et y conduire que montrer un cadre vide.
  if (etat.presents.length === 0) {
    return (
      <div className="carte">
        <h2>Sur le terrain</h2>
        <div className="corps-carte">
          <p className="vide">
            La feuille de match n'est pas encore posée.
            <br />
            <button
              type="button"
              className="btn-vide"
              style={{ marginTop: 10 }}
              onClick={surComposition}
            >
              Composer l'équipe
            </button>
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="carte">
      <h2>
        Sur le terrain
        <button type="button" className="remp" onClick={surRemplacement}>
          Remplacement
        </button>
      </h2>
      <div className="corps-carte">
        {attente && (
          <div className="attente-barre">
            <span>{attente.libelle} — désignez le joueur</span>
            <button type="button" onClick={surAnnuler}>
              annuler
            </button>
          </div>
        )}
        <div className="joueurs">
          {etat.presents.map((n) => {
            const carton = etat.cartons[n];
            const bloque = carton && !attente?.carton;
            return (
              <button
                key={n}
                type="button"
                className={
                  "joueur" +
                  (carton ? ` carton-${carton}` : "") +
                  (attente ? " attente" : "")
                }
                disabled={bloque}
                title={
                  bloque
                    ? `Sous carton ${carton} — déclarer son retour d'abord`
                    : undefined
                }
                onClick={() => surJoueur({ numero: n })}
              >
                <span className="num">{n}</span>
                <span className="nom">{etat.nomDe(n).split(" ")[0]}</span>
              </button>
            );
          })}
          {GROUPES.map((g) => (
            <button
              key={g}
              type="button"
              className={"joueur special" + (attente ? " attente" : "")}
              onClick={() => surJoueur({ groupe: g })}
            >
              <span className="num">{g}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/**
 * La palette d'actions.
 *
 * Elle se construit depuis la base : un club qui ajoute un motif
 * ou une action la voit apparaître sans qu'on recompile. Seule
 * la disposition en colonnes relève de l'interface.
 */
export function PaletteActions({ referentiels, etat, surAction, surPhase, surCarton, surRetourCarton }) {
  if (!referentiels) return <div className="carte" />;

  const { familles, types } = referentiels;
  const parFamille = {};
  for (const t of types) (parFamille[t.famille] ||= []).push(t);

  const libelleFamille = (code) =>
    familles.find((f) => f.code === code)?.libelle ?? code;

  const aucunCarton = Object.keys(etat.cartons).length === 0;

  return (
    <div className="carte" style={{ flex: 1 }}>
      <h2>Palette actions</h2>
      <div className="corps-carte">
        <div className="colonnes">
          {COLONNES.map((colonne, i) => (
            <div key={i}>
              {colonne.map((famille) => (
                <section key={famille} className="famille">
                  <h3>{libelleFamille(famille)}</h3>
                  <div className="actions">
                    {contenuFamille(famille, parFamille[famille] ?? [], {
                      etat, surAction, surPhase, surCarton,
                      surRetourCarton, aucunCarton,
                    })}
                  </div>
                </section>
              ))}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function contenuFamille(famille, types, ctx) {
  const { etat, surAction, surPhase, surCarton, surRetourCarton, aucunCarton } = ctx;

  // Touches et mêlées passent par une popup : un seul bouton par
  // issue, le camp se déduisant de la possession.
  if (famille === "TOUCHE" || famille === "MELEE") {
    const issues = famille === "TOUCHE" ? ISSUES_TOUCHE : ISSUES_MELEE;
    return issues.map((i) => (
      <button
        key={i.code}
        type="button"
        className="act"
        onClick={() => surPhase(famille, i.code)}
      >
        {i.libelle}
      </button>
    ));
  }

  // Les cartons sont quatre pastilles de couleur, plus le retour.
  if (famille === "CARTON") {
    return [
      ...COULEURS_CARTON.map((c) => (
        <button
          key={c}
          type="button"
          className={`act carton ${c[0]}`}
          title={`Carton ${c}`}
          aria-label={`Carton ${c}`}
          onClick={() => surCarton(c)}
        />
      )),
      <button
        key="retour"
        type="button"
        className="act"
        disabled={aucunCarton}
        title={aucunCarton ? "Aucun joueur sous carton" : undefined}
        onClick={surRetourCarton}
      >
        retour
      </button>,
    ];
  }

  const boutons = [];

  // Les gestes à double libellé : un bouton, le code choisi au
  // moment du clic selon qui tient le ballon.
  if (famille === "POSSESSION") {
    for (const duo of DUOS) {
      boutons.push(
        <button
          key={duo.libelle}
          type="button"
          className="act"
          disabled={!etat.ouvert}
          title={!etat.ouvert ? "Le jeu n'est pas ouvert" : undefined}
          onClick={() => surAction(duo)}
        >
          {duo.libelle}
        </button>
      );
    }
  }

  for (const t of types) {
    if (CODES_EN_DUO.has(t.code)) continue;

    // Tant que le jeu n'a pas repris, seules la reprise et
    // l'arrêt répondent.
    const bloque = !etat.ouvert && !t.ouvre_jeu && !t.ferme_jeu;

    boutons.push(
      <button
        key={t.code}
        type="button"
        className={
          "act" +
          (famille === "POINTS" ? " pts" : "") +
          (t.ferme_jeu ? " arret" : "")
        }
        disabled={bloque}
        title={
          bloque
            ? "Le jeu n'est pas ouvert : saisir d'abord une reprise"
            : t.libelle
        }
        onClick={() => surAction(t)}
      >
        {ABREGE[t.code] ?? t.libelle}
      </button>
    );
  }
  return boutons;
}
