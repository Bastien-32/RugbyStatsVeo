// Les règles du jeu, côté interface.
//
// L'API les applique aussi, pour ses calculs. Ici elles servent
// à décider ce que la palette propose : griser les actions tant
// que le jeu n'a pas repris, choisir le libellé d'un geste selon
// qui tient le ballon.

export const ARRET = "ARRET";

export const GROUPES = ["Avant", "Trois-quarts", "Collectif", "?"];

// Une demi-seconde de marge : deux lectures du lecteur vidéo ne
// rendent jamais exactement le même nombre.
const MARGE = 0.5;

/**
 * Le jeu est-il ouvert à cet instant de la vidéo ?
 *
 * Seule compte la dernière reprise ou le dernier arrêt qui
 * précède la position courante. Revenir en arrière pour
 * compléter une phase doit rendre l'état de cet instant-là,
 * pas celui de la fin du match.
 */
export function jeuOuvert(actions, types, mt, temps) {
  let meilleur = -1;
  let ouvert = false;

  for (const a of actions) {
    if (a.mt !== mt || a.temps > temps + MARGE) continue;
    const t = types[a.type_code];
    if (!t || (!t.ouvre_jeu && !t.ferme_jeu)) continue;
    if (a.temps > meilleur) {
      meilleur = a.temps;
      ouvert = t.ouvre_jeu;
    }
  }
  return ouvert;
}

/**
 * Les joueurs sous carton à cet instant, par numéro.
 *
 * L'état se lit au journal, comme dans le classeur : un carton
 * posé, puis un retour déclaré. Rien n'est stocké à part.
 */
export function cartonsEnCours(actions, mt, temps) {
  const etat = {};
  const triees = [...actions]
    .filter((a) => a.temps <= temps + MARGE)
    .sort((a, b) => a.temps - b.temps);

  for (const a of triees) {
    if (a.numero == null) continue;
    const couleur = couleurCarton(a.type_code);
    if (couleur && a.type_code.startsWith("CAR_")) {
      etat[a.numero] = couleur;
    } else if (couleur && a.type_code.startsWith("RET_")) {
      delete etat[a.numero];
    }
  }
  return etat;
}

function couleurCarton(code) {
  if (code.includes("ROUGE")) return "rouge";
  if (code.includes("JAUNE")) return "jaune";
  if (code.includes("BLANC")) return "blanc";
  if (code.includes("BLEU")) return "bleu";
  return null;
}

/**
 * Qui est sur le terrain à cet instant, par numéro de maillot.
 *
 * Le numéro suit le joueur : un remplaçant entre avec le sien,
 * il ne reprend pas celui du sortant.
 */
export function surLeTerrain(feuille, remplacements, temps) {
  const presents = feuille
    .filter((l) => l.titulaire)
    .map((l) => l.numero)
    .sort((a, b) => a - b);

  const faits = [...remplacements]
    .filter((r) => r.temps <= temps + MARGE)
    .sort((a, b) => a.temps - b.temps);

  for (const r of faits) {
    const place = presents.indexOf(r.sortant);
    if (place !== -1) presents[place] = r.entrant;
  }
  return presents;
}

/**
 * L'issue d'une phase décide à qui revient le ballon.
 *
 *   G  il revient à l'équipe qui lance ;
 *   P  il reste à l'autre ;
 *   ND la phase est rejouée par l'adversaire.
 */
export function possessionApres(issue, lanceur) {
  if (issue === "G") return lanceur;
  return lanceur === "Nous" ? "Adv" : "Nous";
}

export function inverse(possession) {
  return possession === "Nous" ? "Adv" : "Nous";
}

/**
 * Le code d'une phase, selon le camp qui la joue et son issue.
 * Les quatre combinaisons existent en base ; on les compose
 * plutôt que de les énumérer dans l'interface.
 */
export function codePhase(famille, possession, issue) {
  const camp = possession === "Nous" ? "NS" : "ADV";
  return famille === "TOUCHE" ? `TO_${camp}_${issue}` : `ME_${camp}_${issue}`;
}

/**
 * Le code d'un carton, selon qu'un joueur de chez nous est
 * désigné ou non.
 *
 * Le camp se lit au joueur, jamais à la possession : après un
 * essai adverse, le ballon est déjà revenu à nous alors que le
 * carton est pour notre camp.
 */
export function codeCarton(couleur, pourNous) {
  const c = couleur.toUpperCase();
  return `CAR_${c}_${pourNous ? "NS" : "ADV"}`;
}

export function codeRetourCarton(couleur, pourNous) {
  const c = couleur.toUpperCase();
  return `RET_${c}_${pourNous ? "NS" : "ADV"}`;
}

/** Mise en forme des temps vidéo. */
export function mmss(secondes) {
  const s = Math.max(0, secondes || 0);
  const m = Math.floor(s / 60);
  const r = s % 60;
  return `${m}:${r < 10 ? "0" : ""}${r.toFixed(1).replace(".", ",")}`;
}

export function court(secondes) {
  const s = Math.max(0, Math.floor(secondes || 0));
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
}
