import { useEffect, useState } from "react";

const TITULAIRES = Array.from({ length: 15 }, (_, i) => i + 1);
const BANC = Array.from({ length: 8 }, (_, i) => i + 16);

/**
 * La feuille de match.
 *
 * Les noms viennent de l'effectif du club : on choisit, on ne
 * retape pas. La ligne d'un titulaire découle de son poste ;
 * celle d'un remplaçant se déclare, un banc pouvant compter
 * trois avants et quatre trois-quarts comme l'inverse.
 */
export function Composition({ feuille, joueurs, surEnregistrer, erreur }) {
  const [lignes, setLignes] = useState({});
  const [enregistre, setEnregistre] = useState(false);

  useEffect(() => {
    const d = {};
    for (const l of feuille) {
      d[l.numero] = { joueur_id: l.joueur_id, ligne: l.ligne };
    }
    setLignes(d);
  }, [feuille]);

  const parNom = {};
  for (const j of joueurs) parNom[nomComplet(j)] = j;

  function changerJoueur(numero, nom) {
    const j = parNom[nom];
    setLignes((d) => ({
      ...d,
      [numero]: {
        joueur_id: j?.id ?? null,
        // Un joueur connu apporte sa ligne habituelle ; elle
        // reste modifiable si ce jour-là il dépanne ailleurs.
        ligne: d[numero]?.ligne ?? j?.ligne ?? ligneDePoste(numero),
      },
    }));
    setEnregistre(false);
  }

  function changerLigne(numero, ligne) {
    setLignes((d) => ({ ...d, [numero]: { ...d[numero], ligne } }));
    setEnregistre(false);
  }

  async function enregistrer() {
    const payload = Object.entries(lignes)
      .filter(([, v]) => v?.joueur_id)
      .map(([numero, v]) => ({
        numero: Number(numero),
        joueur_id: v.joueur_id,
        nom_complet: "",
        ligne: Number(numero) <= 15 ? ligneDePoste(Number(numero)) : v.ligne,
        titulaire: Number(numero) <= 15,
      }));
    await surEnregistrer(payload);
    setEnregistre(true);
  }

  // Un nom tapé qui ne correspond à aucun joueur du club ne
  // peut pas être enregistré : on le dit avant de valider.
  const retenues = Object.values(lignes).filter((v) => v?.joueur_id).length;

  const compteBanc = BANC.reduce(
    (acc, n) => {
      const l = lignes[n];
      if (!l?.joueur_id) return acc;
      l.ligne === "Avant" ? acc.avants++ : acc.troisQuarts++;
      return acc;
    },
    { avants: 0, troisQuarts: 0 }
  );

  return (
    <section className="vue">
      <div className="carte" style={{ flex: 1 }}>
        <h2>
          Feuille de match
          <span className="compte-feuille">
            {retenues} joueur{retenues > 1 ? "s" : ""} retenu
            {retenues > 1 ? "s" : ""}
          </span>
          <button type="button" className="remp" onClick={enregistrer}>
            {enregistre ? "Enregistré" : "Enregistrer"}
          </button>
        </h2>

        {erreur && <div className="erreur" style={{ margin: "10px 12px 0" }}>{erreur}</div>}

        <div className="compo-grille">
          <div className="compo-bloc">
            <h3>Avants</h3>
            <p className="note">Postes 1 à 8. La ligne découle du poste.</p>
            {TITULAIRES.slice(0, 8).map((n) => (
              <LigneCompo
                key={n} numero={n} valeur={lignes[n]} joueurs={joueurs}
                parNom={parNom} surJoueur={changerJoueur} modifiable={false}
              />
            ))}
          </div>

          <div className="compo-bloc">
            <h3>Trois-quarts</h3>
            <p className="note">Postes 9 à 15. La ligne découle du poste.</p>
            {TITULAIRES.slice(8).map((n) => (
              <LigneCompo
                key={n} numero={n} valeur={lignes[n]} joueurs={joueurs}
                parNom={parNom} surJoueur={changerJoueur} modifiable={false}
              />
            ))}
          </div>

          <div className="compo-bloc">
            <h3>Remplaçants</h3>
            <p className="note">
              Chacun garde son numéro en entrant. Sa ligne se choisit ici.
            </p>
            {BANC.map((n) => (
              <LigneCompo
                key={n} numero={n} valeur={lignes[n]} joueurs={joueurs}
                parNom={parNom} surJoueur={changerJoueur}
                surLigne={changerLigne} modifiable
              />
            ))}
            <p className="resume-banc">
              Banc : {compteBanc.avants} avant{compteBanc.avants > 1 ? "s" : ""} ·{" "}
              {compteBanc.troisQuarts} trois-quart
              {compteBanc.troisQuarts > 1 ? "s" : ""}
            </p>
          </div>
        </div>
      </div>

      <datalist id="effectif-club">
        {joueurs.map((j) => (
          <option key={j.id} value={nomComplet(j)} />
        ))}
      </datalist>
    </section>
  );
}

function LigneCompo({ numero, valeur, joueurs, parNom, surJoueur, surLigne, modifiable }) {
  const joueur = joueurs.find((j) => j.id === valeur?.joueur_id);
  const ligne = modifiable
    ? valeur?.ligne ?? "Avant"
    : ligneDePoste(numero);

  return (
    <div className="compo-ligne">
      <span className="poste">{numero}</span>
      <input
        id={`compo-${numero}`}
        type="text"
        list="effectif-club"
        defaultValue={joueur ? nomComplet(joueur) : ""}
        placeholder="nom ou prénom…"
        aria-label={`Joueur numéro ${numero}`}
        onChange={(e) => surJoueur(numero, e.target.value)}
      />
      {modifiable ? (
        <select
          id={`ligne-${numero}`}
          value={ligne}
          aria-label={`Ligne du joueur ${numero}`}
          onChange={(e) => surLigne(numero, e.target.value)}
        >
          <option>Avant</option>
          <option>Trois-quarts</option>
        </select>
      ) : (
        <span className="ligne-fixe">{ligne}</span>
      )}
    </div>
  );
}

function ligneDePoste(n) {
  return n <= 8 ? "Avant" : "Trois-quarts";
}

function nomComplet(j) {
  return `${j.nom} ${j.prenom}`.trim();
}
