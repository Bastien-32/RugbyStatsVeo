import { useEffect, useState } from "react";
import { api } from "./api.js";

// Cette page vérifie la chaîne complète : le navigateur parle
// à nginx, nginx à l'API, l'API à PostgreSQL. Elle montre ce
// que la base contient réellement — le vocabulaire du jeu,
// l'effectif, les matchs — plutôt qu'un « Hello world » qui
// ne prouverait rien.
export default function App() {
  const [etat, setEtat] = useState("attente");
  const [erreur, setErreur] = useState(null);
  const [ref, setRef] = useState(null);
  const [joueurs, setJoueurs] = useState([]);
  const [matchs, setMatchs] = useState([]);

  async function recharger() {
    try {
      await api.sante();
      setEtat("ok");
      const [r, j, m] = await Promise.all([
        api.referentiels(),
        api.joueurs(),
        api.matchs(),
      ]);
      setRef(r);
      setJoueurs(j ?? []);
      setMatchs(m ?? []);
      setErreur(null);
    } catch (e) {
      setEtat("panne");
      setErreur(e.message);
    }
  }

  useEffect(() => {
    recharger();
  }, []);

  return (
    <div className="page">
      <header className="appli">
        <h1>RugbyStats</h1>
        <p>Analyse vidéo de match</p>
        <span className={`etat ${etat}`}>
          <span className="pastille" />
          {etat === "ok" ? "API connectée" :
           etat === "panne" ? "API injoignable" : "connexion…"}
        </span>
      </header>

      {erreur && <div className="erreur">{erreur}</div>}

      <div className="cartes">
        <Vocabulaire ref={ref} />
        <Effectif joueurs={joueurs} apresAjout={recharger} />
        <Matchs matchs={matchs} apresAjout={recharger} />
      </div>

      <p className="pied">
        Base PostgreSQL · API Go · interface React — trois conteneurs.
      </p>
    </div>
  );
}

// ---------------------------------------------------------

function Vocabulaire({ ref: referentiels }) {
  if (!referentiels) {
    return (
      <section className="carte">
        <h2>Vocabulaire du jeu</h2>
        <p className="vide">Chargement…</p>
      </section>
    );
  }

  const { familles, types, motifs } = referentiels;
  const parFamille = (code) => types.filter((t) => t.famille === code).length;

  return (
    <section className="carte">
      <h2>Vocabulaire du jeu</h2>
      <p className="sous">
        Chargé depuis la base, pas écrit dans le code : un club peut
        avoir ses propres motifs de faute.
      </p>

      <div className="compteur">{types.length}</div>
      <div className="libelle">types d'actions</div>

      <div className="familles" style={{ marginTop: 14 }}>
        {familles.map((f) => (
          <span key={f.code}>
            {f.libelle} <b>{parFamille(f.code)}</b>
          </span>
        ))}
      </div>

      <p className="sous" style={{ marginTop: 16, marginBottom: 4 }}>
        {motifs.filter((m) => m.penalite).length} motifs de pénalité ·{" "}
        {motifs.filter((m) => m.coup_franc).length} de coup franc
      </p>
    </section>
  );
}

// ---------------------------------------------------------

function Effectif({ joueurs, apresAjout }) {
  const [nom, setNom] = useState("");
  const [prenom, setPrenom] = useState("");
  const [ligne, setLigne] = useState("Avant");
  const [erreur, setErreur] = useState(null);

  async function ajouter(e) {
    e.preventDefault();
    if (!nom.trim()) return;
    try {
      await api.creerJoueur({
        nom: nom.trim().toUpperCase(),
        prenom: prenom.trim(),
        ligne,
        actif: true,
      });
      setNom("");
      setPrenom("");
      setErreur(null);
      apresAjout();
    } catch (e) {
      setErreur(e.message);
    }
  }

  return (
    <section className="carte">
      <h2>Effectif du club</h2>
      <p className="sous">
        Commun à toutes les équipes. Une feuille de match y puise.
      </p>

      {joueurs.length === 0 ? (
        <p className="vide">Aucun joueur. Ajoutez-en un ci-dessous.</p>
      ) : (
        <ul className="liste">
          {joueurs.slice(0, 8).map((j) => (
            <li key={j.id}>
              <span className="code">{j.ligne ?? "—"}</span>
              <span className="nom">
                {j.nom} {j.prenom}
              </span>
            </li>
          ))}
          {joueurs.length > 8 && (
            <li>
              <span className="code" />
              <span className="nom" style={{ opacity: 0.7 }}>
                et {joueurs.length - 8} autres
              </span>
            </li>
          )}
        </ul>
      )}

      {erreur && <div className="erreur" style={{ marginTop: 12 }}>{erreur}</div>}

      <form className="ligne" onSubmit={ajouter}>
        <input
          id="joueur-nom"
          value={nom}
          onChange={(e) => setNom(e.target.value)}
          placeholder="NOM"
          aria-label="Nom du joueur"
        />
        <input
          id="joueur-prenom"
          value={prenom}
          onChange={(e) => setPrenom(e.target.value)}
          placeholder="Prénom"
          aria-label="Prénom du joueur"
        />
        <select
          id="joueur-ligne"
          value={ligne}
          onChange={(e) => setLigne(e.target.value)}
          aria-label="Ligne de jeu"
        >
          <option>Avant</option>
          <option>Trois-quarts</option>
        </select>
        <button type="submit">Ajouter</button>
      </form>
    </section>
  );
}

// ---------------------------------------------------------

function Matchs({ matchs, apresAjout }) {
  const [adversaire, setAdversaire] = useState("");
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [erreur, setErreur] = useState(null);

  async function ajouter(e) {
    e.preventDefault();
    if (!adversaire.trim()) return;
    try {
      await api.creerMatch({
        saison: saisonDe(date),
        date_match: date,
        adversaire: adversaire.trim(),
        categorie: "Première",
        lieu: "Domicile",
        phase: "Aller",
        sens_mt1: "droite",
        video_ref: "",
      });
      setAdversaire("");
      setErreur(null);
      apresAjout();
    } catch (e) {
      setErreur(e.message);
    }
  }

  return (
    <section className="carte">
      <h2>Matchs</h2>
      <p className="sous">
        Chacun porte son journal d'actions, horodaté sur la vidéo.
      </p>

      {matchs.length === 0 ? (
        <p className="vide">Aucun match enregistré.</p>
      ) : (
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th>Adversaire</th>
              <th>Lieu</th>
            </tr>
          </thead>
          <tbody>
            {matchs.map((m) => (
              <tr key={m.id}>
                <td>{m.date_match}</td>
                <td>{m.adversaire}</td>
                <td>{m.lieu}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {erreur && <div className="erreur" style={{ marginTop: 12 }}>{erreur}</div>}

      <form className="ligne" onSubmit={ajouter}>
        <input
          id="match-date"
          type="date"
          value={date}
          onChange={(e) => setDate(e.target.value)}
          aria-label="Date du match"
        />
        <input
          id="match-adversaire"
          value={adversaire}
          onChange={(e) => setAdversaire(e.target.value)}
          placeholder="Adversaire"
          aria-label="Adversaire"
        />
        <button type="submit">Créer</button>
      </form>
    </section>
  );
}

// La saison bascule en juillet : un match de septembre 2026
// appartient à la saison 26-27.
function saisonDe(date) {
  const d = new Date(date);
  const annee = d.getMonth() >= 6 ? d.getFullYear() : d.getFullYear() - 1;
  return `${String(annee).slice(2)}-${String(annee + 1).slice(2)}`;
}
