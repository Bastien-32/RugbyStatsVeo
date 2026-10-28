// La disposition de la palette.
//
// Ce que les boutons font vient de la base ; où ils se trouvent
// regarde l'interface seule. On range les familles en colonnes
// pour que tout tienne sans ascenseur, l'œil restant sur la
// vidéo pendant la saisie.

export const COLONNES = [
  ["POINTS", "REPRISE"],
  ["CARTON", "TOUCHE", "MELEE"],
  ["PLAQUAGE", "ATTITUDE"],
  ["POSSESSION"],
  ["PENALITE"],
  ["COUP_FRANC"],
];

// Un libellé de bouton plus court que celui du journal : la
// famille dit déjà de quoi il s'agit, « Plaquage normal » peut
// se lire « normal » sous le titre Plaquages.
export const ABREGE = {
  PTS_ESSAI: "essai",
  PTS_ESSAI_PEN: "essai de pénalité",
  PTS_TRANSFO: "transfo",
  PTS_PENALITE: "pénalité",
  PTS_DROP: "drop",

  REP_TOUCHE: "touche",
  REP_MELEE: "mêlée",
  REP_RENVOI: "renvoi",
  REP_MAIN: "à la main",
  ARRET: "arrêt du jeu",

  PL_OFF: "offensif",
  PL_NORMAL: "normal",
  PL_RATE: "raté",
  PL_A2: "à 2",
  PL_HAUT: "haut",

  ATT_FRANCHISSEMENT: "franchissement",
  ATT_AVANCEE: "avancée",
  ATT_NULLE: "stoppé / recul",

  EN_AVANT: "en-avant",
  JEU_AU_PIED: "jeu au pied",
  RECEPTION_NS: "réception nous",
  RECEPTION_ADV: "réception Adv",

  PEN_CONTRE_ADV: "contre adversaire",
  PEN_CONTRE_NS: "contre nous",
  CF_CONTRE_ADV: "contre adversaire",
  CF_CONTRE_NS: "contre nous",
};

// Les gestes qui changent de nom selon qui tenait le ballon.
// Un seul bouton, deux libellés : gratter, c'est récupérer
// quand ils l'avaient, perdre quand nous l'avions.
export const DUOS = [
  { libelle: "grattage", rec: "GRATTAGE_REC", perte: "GRATTAGE_PERTE" },
  { libelle: "contre-ruck", rec: "CRUCK_REC", perte: "CRUCK_PERTE" },
  { libelle: "arrachage", rec: "ARRACHAGE_REC", perte: "ARRACHAGE_PERTE" },
  { libelle: "turn-over", rec: "TO_REC", perte: "TO_PERTE" },
];

// Les codes repris par un duo n'apparaissent pas seuls dans la
// palette : leur bouton est déjà là.
export const CODES_EN_DUO = new Set(
  DUOS.flatMap((d) => [d.rec, d.perte])
);

export const COULEURS_CARTON = ["rouge", "jaune", "blanc", "bleu"];

// Les listes des popups de touche et de mêlée, reprises du
// classeur. Elles rejoindront la base le jour où un club
// voudra les siennes.
export const ZONE_SAUT = ["Z0", "Z1", "Z1,5", "Z2", "Z3"];
export const ALIGNEMENT = ["2", "3", "4", "5", "6", "Complet"];
export const BALLON = ["C", "F"];
export const UTILISATION = ["Avants", "Trois-quart", "Pied"];
export const ZONE_LONGUEUR = [
  "N en-but", "N 15m", "N 22m", "N 30m", "N 40m", "50m",
  "E 40m", "E 30m", "E 22m", "E 15m", "E en-but",
];
export const ZONE_LARGEUR = ["Gauche", "Milieu", "Droit"];

export const ISSUES_TOUCHE = [
  { code: "G", libelle: "gagnée" },
  { code: "P", libelle: "perdue" },
  { code: "ND", libelle: "pas droit" },
];
export const ISSUES_MELEE = [
  { code: "G", libelle: "gagnée" },
  { code: "P", libelle: "perdue" },
];
