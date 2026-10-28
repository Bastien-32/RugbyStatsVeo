// Accès à l'API.
//
// Tout passe par /api : nginx en production, le proxy de Vite
// en développement. Le front ne connaît jamais l'adresse du
// serveur, donc rien à reconfigurer d'un environnement à
// l'autre.

const BASE = "/api";

async function appeler(chemin, options = {}) {
  const reponse = await fetch(BASE + chemin, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });

  if (reponse.status === 204) return null;

  const corps = await reponse.json().catch(() => null);

  if (!reponse.ok) {
    // L'API renvoie toujours { erreur: "..." } : on remonte ce
    // message plutôt qu'un code HTTP, qui ne dirait rien à
    // celui qui saisit un match.
    throw new Error(corps?.erreur || `Erreur ${reponse.status}`);
  }
  return corps;
}

export const api = {
  sante: () => appeler("/sante"),

  referentiels: () => appeler("/referentiels"),

  joueurs: () => appeler("/joueurs"),
  creerJoueur: (joueur) =>
    appeler("/joueurs", { method: "POST", body: JSON.stringify(joueur) }),

  matchs: () => appeler("/matchs"),
  match: (id) => appeler(`/matchs/${id}`),
  creerMatch: (match) =>
    appeler("/matchs", { method: "POST", body: JSON.stringify(match) }),

  feuille: (id) => appeler(`/matchs/${id}/feuille`),
  poserFeuille: (id, lignes) =>
    appeler(`/matchs/${id}/feuille`, {
      method: "PUT",
      body: JSON.stringify(lignes),
    }),

  actions: (id) => appeler(`/matchs/${id}/actions`),
  ajouterAction: (id, action) =>
    appeler(`/matchs/${id}/actions`, {
      method: "POST",
      body: JSON.stringify(action),
    }),
  supprimerAction: (actionId) =>
    appeler(`/actions/${actionId}`, { method: "DELETE" }),

  remplacements: (id) => appeler(`/matchs/${id}/remplacements`),
  ajouterRemplacement: (id, remplacement) =>
    appeler(`/matchs/${id}/remplacements`, {
      method: "POST",
      body: JSON.stringify(remplacement),
    }),

  statistiques: (id) => appeler(`/matchs/${id}/statistiques`),
};
