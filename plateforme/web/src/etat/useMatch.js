import { useCallback, useEffect, useMemo, useState } from "react";
import { api } from "../api.js";
import {
  cartonsEnCours, jeuOuvert, surLeTerrain, inverse,
} from "./regles.js";

/**
 * Tout ce qu'un match met en jeu : sa feuille, son journal, ses
 * remplacements, et les dérivés qu'on recalcule plutôt que de
 * les stocker.
 *
 * Le serveur reste la source : chaque écriture part à l'API et
 * la réponse nourrit l'état local. Une saisie perdue au
 * rafraîchissement serait pire qu'une saisie lente.
 */
export function useMatch(matchId) {
  const [match, setMatch] = useState(null);
  const [feuille, setFeuille] = useState([]);
  const [actions, setActions] = useState([]);
  const [remplacements, setRemplacements] = useState([]);
  const [referentiels, setReferentiels] = useState(null);
  const [joueurs, setJoueurs] = useState([]);
  const [erreur, setErreur] = useState(null);
  const [chargement, setChargement] = useState(true);

  const recharger = useCallback(async () => {
    try {
      const [m, f, a, r, ref, j] = await Promise.all([
        api.match(matchId),
        api.feuille(matchId),
        api.actions(matchId),
        api.remplacements(matchId),
        api.referentiels(),
        api.joueurs(),
      ]);
      setMatch(m);
      setFeuille(f ?? []);
      setActions(a ?? []);
      setRemplacements(r ?? []);
      setReferentiels(ref);
      setJoueurs(j ?? []);
      setErreur(null);
    } catch (e) {
      setErreur(e.message);
    } finally {
      setChargement(false);
    }
  }, [matchId]);

  useEffect(() => {
    setChargement(true);
    recharger();
  }, [recharger]);

  // Les types par code : les règles ont besoin de savoir ce que
  // chaque action fait au jeu.
  const types = useMemo(() => {
    const m = {};
    for (const t of referentiels?.types ?? []) m[t.code] = t;
    return m;
  }, [referentiels]);

  const ajouterAction = useCallback(
    async (action) => {
      try {
        const creee = await api.ajouterAction(matchId, action);
        setActions((liste) =>
          [...liste, creee].sort((a, b) => b.temps - a.temps)
        );
        setErreur(null);
        return creee;
      } catch (e) {
        setErreur(e.message);
        return null;
      }
    },
    [matchId]
  );

  const supprimerAction = useCallback(async (id) => {
    try {
      await api.supprimerAction(id);
      setActions((liste) => liste.filter((a) => a.id !== id));
    } catch (e) {
      setErreur(e.message);
    }
  }, []);

  const ajouterRemplacement = useCallback(
    async (r) => {
      try {
        const cree = await api.ajouterRemplacement(matchId, r);
        setRemplacements((liste) =>
          [...liste, cree].sort((a, b) => a.temps - b.temps)
        );
        return cree;
      } catch (e) {
        setErreur(e.message);
        return null;
      }
    },
    [matchId]
  );

  const poserFeuille = useCallback(
    async (lignes) => {
      try {
        await api.poserFeuille(matchId, lignes);
        setFeuille(await api.feuille(matchId));
        setErreur(null);
      } catch (e) {
        setErreur(e.message);
      }
    },
    [matchId]
  );

  return {
    match, feuille, actions, remplacements, referentiels, joueurs,
    types, erreur, chargement,
    recharger, ajouterAction, supprimerAction, ajouterRemplacement,
    poserFeuille, setErreur,
  };
}

/**
 * L'état du jeu à la position courante de la vidéo : qui est sur
 * le terrain, qui porte un carton, si la palette répond.
 *
 * Tout se déduit du journal. Rien n'est tenu à jour en parallèle,
 * donc rien ne peut diverger après une correction.
 */
export function useEtatCourant({ actions, feuille, remplacements, types },
  miTemps, position) {

  return useMemo(() => {
    const presents = surLeTerrain(feuille, remplacements, position);
    const cartons = cartonsEnCours(actions, miTemps, position);
    const ouvert = jeuOuvert(actions, types, miTemps, position);

    const parNumero = {};
    for (const l of feuille) parNumero[l.numero] = l;

    return {
      presents,
      cartons,
      ouvert,
      ligneDe: (numero) => parNumero[numero]?.ligne ?? "Avant",
      nomDe: (numero) => parNumero[numero]?.nom_complet ?? "",
      etiquette: (numero) =>
        `${numero} ${parNumero[numero]?.nom_complet ?? ""}`.trim(),
      banc: feuille
        .filter((l) => !presents.includes(l.numero))
        .map((l) => l.numero),
    };
  }, [actions, feuille, remplacements, types, miTemps, position]);
}

export { inverse };
