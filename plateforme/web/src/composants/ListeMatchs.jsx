import { useEffect, useState } from "react";
import { api } from "../api.js";

/**
 * L'accueil : les matchs du club, et de quoi en créer un.
 *
 * Un match ouvert devient la vue de saisie ; la vidéo se charge
 * à ce moment-là, depuis le poste.
 */
export function ListeMatchs({ surOuvrir }) {
  const [matchs, setMatchs] = useState([]);
  const [joueurs, setJoueurs] = useState([]);
  const [erreur, setErreur] = useState(null);
  const [chargement, setChargement] = useState(true);

  const [adversaire, setAdversaire] = useState("");
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [categorie, setCategorie] = useState("Première");
  const [lieu, setLieu] = useState("Domicile");

  async function recharger() {
    try {
      const [m, j] = await Promise.all([api.matchs(), api.joueurs()]);
      setMatchs(m ?? []);
      setJoueurs(j ?? []);
      setErreur(null);
    } catch (e) {
      setErreur(e.message);
    } finally {
      setChargement(false);
    }
  }

  useEffect(() => {
    recharger();
  }, []);

  async function creer(e) {
    e.preventDefault();
    if (!adversaire.trim()) return;
    try {
      const m = await api.creerMatch({
        saison: saisonDe(date),
        date_match: date,
        adversaire: adversaire.trim(),
        categorie,
        lieu,
        phase: "Aller",
        sens_mt1: "droite",
        video_ref: "",
      });
      setAdversaire("");
      await recharger();
      surOuvrir(m.id);
    } catch (e) {
      setErreur(e.message);
    }
  }

  return (
    <div className="page">
      <header className="accueil-tete">
        <h1>RugbyStats</h1>
        <p>
          Analyse vidéo de match — {joueurs.length} joueurs à l'effectif
        </p>
      </header>

      {erreur && <div className="erreur">{erreur}</div>}

      <section className="carte" style={{ marginBottom: 16 }}>
        <h2>Nouveau match</h2>
        <form className="ligne" onSubmit={creer}>
          <input
            id="m-date" type="date" value={date}
            onChange={(e) => setDate(e.target.value)}
            aria-label="Date du match"
          />
          <input
            id="m-adversaire" value={adversaire}
            onChange={(e) => setAdversaire(e.target.value)}
            placeholder="Adversaire" aria-label="Adversaire"
          />
          <select
            id="m-categorie" value={categorie}
            onChange={(e) => setCategorie(e.target.value)}
            aria-label="Catégorie"
          >
            <option>Première</option>
            <option>Réserve</option>
          </select>
          <select
            id="m-lieu" value={lieu}
            onChange={(e) => setLieu(e.target.value)}
            aria-label="Lieu"
          >
            <option>Domicile</option>
            <option>Extérieur</option>
            <option>Neutre</option>
          </select>
          <button type="submit">Créer et ouvrir</button>
        </form>
      </section>

      <section className="carte">
        <h2>Matchs</h2>
        {chargement ? (
          <p className="vide">Chargement…</p>
        ) : matchs.length === 0 ? (
          <p className="vide">
            Aucun match. Créez-en un ci-dessus pour commencer la saisie.
          </p>
        ) : (
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Adversaire</th>
                <th>Catégorie</th>
                <th>Lieu</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {matchs.map((m) => (
                <tr key={m.id}>
                  <td>{m.date_match}</td>
                  <td>{m.adversaire}</td>
                  <td>{m.categorie || "—"}</td>
                  <td>{m.lieu}</td>
                  <td style={{ textAlign: "right" }}>
                    <button
                      type="button" className="btn-vide"
                      onClick={() => surOuvrir(m.id)}
                    >
                      Ouvrir
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      <p className="pied">
        Base PostgreSQL · API Go · interface React — trois conteneurs.
      </p>
    </div>
  );
}

// La saison bascule en juillet : un match de septembre 2026
// appartient à la saison 26-27.
function saisonDe(date) {
  const d = new Date(date);
  const annee = d.getMonth() >= 6 ? d.getFullYear() : d.getFullYear() - 1;
  return `${String(annee).slice(2)}-${String(annee + 1).slice(2)}`;
}
