-- =========================================================
-- Schéma de base
--
-- Le vocabulaire du rugby vit en base, pas dans le code :
-- une fédération, un club ou une catégorie d'âge peut avoir
-- ses propres motifs de faute sans qu'on recompile quoi que
-- ce soit.
-- =========================================================

CREATE TYPE possession AS ENUM ('Nous', 'Adv');

CREATE TYPE ligne_jeu AS ENUM ('Avant', 'Trois-quarts');

-- Une touche se gagne, se perd, ou est rejouée pour non-droite.
-- Une mêlée n'a que les deux premières.
CREATE TYPE issue AS ENUM ('G', 'P', 'ND');

CREATE TYPE lieu_match AS ENUM ('Domicile', 'Extérieur', 'Neutre');

CREATE TYPE mi_temps AS ENUM ('MT1', 'MT2');


-- ---------------------------------------------------------
-- Le club et ses joueurs
--
-- L'effectif ne dépend pas d'un match : il est commun à
-- toutes les équipes du club, et une feuille de match y
-- puise.
-- ---------------------------------------------------------

CREATE TABLE club (
    id          integer GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nom         text NOT NULL,
    cree_le     timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE joueur (
    id          integer GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    club_id     integer NOT NULL REFERENCES club(id) ON DELETE CASCADE,
    nom         text NOT NULL,
    prenom      text NOT NULL DEFAULT '',
    -- La ligne habituelle : elle se redéfinit match par match
    -- sur la feuille, un avant pouvant dépanner derrière.
    ligne       ligne_jeu,
    actif       boolean NOT NULL DEFAULT true,
    cree_le     timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX joueur_club_idx ON joueur (club_id) WHERE actif;

-- Le nom affiché partout : « NOM Prénom », comme sur la
-- feuille de match.
CREATE VIEW joueur_nomme AS
SELECT j.*, trim(j.nom || ' ' || j.prenom) AS nom_complet
FROM joueur j;


-- ---------------------------------------------------------
-- Les référentiels d'actions
--
-- Chaque action du journal pointe vers un type. Le libellé
-- peut changer sans toucher aux matchs déjà saisis.
-- ---------------------------------------------------------

CREATE TABLE famille_action (
    code        text PRIMARY KEY,
    libelle     text NOT NULL,
    rang        integer NOT NULL
);

CREATE TABLE type_action (
    code        text PRIMARY KEY,
    famille     text NOT NULL REFERENCES famille_action(code),
    libelle     text NOT NULL,
    rang        integer NOT NULL DEFAULT 0,

    -- Ce que l'action fait au jeu, lu par le moteur de saisie.
    ouvre_jeu   boolean NOT NULL DEFAULT false,  -- une reprise
    ferme_jeu   boolean NOT NULL DEFAULT false,  -- un arrêt
    arret_auto  boolean NOT NULL DEFAULT false,  -- entraîne un arrêt
    bascule     boolean NOT NULL DEFAULT false,  -- change la possession

    -- Qui porte l'action : notre joueur, personne, l'adversaire.
    attend_joueur boolean NOT NULL DEFAULT false
);

CREATE INDEX type_action_famille_idx ON type_action (famille, rang);

CREATE TABLE motif_faute (
    code        text PRIMARY KEY,
    libelle     text NOT NULL,
    famille     text NOT NULL DEFAULT 'Autres',
    -- Un motif sert à la pénalité, au coup franc, ou aux deux.
    penalite    boolean NOT NULL DEFAULT true,
    coup_franc  boolean NOT NULL DEFAULT false,
    rang        integer NOT NULL DEFAULT 0
);


-- ---------------------------------------------------------
-- Les matchs
-- ---------------------------------------------------------

CREATE TABLE match (
    id          integer GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    club_id     integer NOT NULL REFERENCES club(id) ON DELETE CASCADE,
    saison      text NOT NULL,
    date_match  date NOT NULL,
    adversaire  text NOT NULL,
    categorie   text NOT NULL DEFAULT '',
    lieu        lieu_match NOT NULL DEFAULT 'Domicile',
    phase       text NOT NULL DEFAULT '',
    journee     integer,
    -- Le sens de jeu de notre équipe en première mi-temps.
    sens_mt1    text NOT NULL DEFAULT 'droite',
    -- L'identifiant de la vidéo chez l'hébergeur, ou un
    -- chemin local tant qu'on travaille hors ligne.
    video_ref   text NOT NULL DEFAULT '',
    cree_le     timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX match_club_date_idx ON match (club_id, date_match DESC);


-- La feuille de match : qui porte quel numéro, et dans quelle
-- ligne il joue ce jour-là.
CREATE TABLE feuille_match (
    match_id    integer NOT NULL REFERENCES match(id) ON DELETE CASCADE,
    numero      integer NOT NULL CHECK (numero BETWEEN 1 AND 30),
    joueur_id   integer NOT NULL REFERENCES joueur(id),
    ligne       ligne_jeu NOT NULL,
    titulaire   boolean NOT NULL,
    PRIMARY KEY (match_id, numero)
);

-- Un joueur ne figure qu'une fois sur une feuille.
CREATE UNIQUE INDEX feuille_joueur_unique
    ON feuille_match (match_id, joueur_id);


-- ---------------------------------------------------------
-- Le journal d'actions
--
-- Le cœur : une ligne par fait de jeu, horodatée sur la
-- vidéo. Tout le reste s'en déduit.
-- ---------------------------------------------------------

CREATE TABLE action (
    id          bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    match_id    integer NOT NULL REFERENCES match(id) ON DELETE CASCADE,

    -- En secondes depuis le début de la vidéo. Le temps du
    -- lecteur, pas le temps de jeu : c'est lui qui permet de
    -- revenir à l'image.
    temps       numeric(9,3) NOT NULL CHECK (temps >= 0),
    mt          mi_temps NOT NULL,

    type_code   text NOT NULL REFERENCES type_action(code),
    possession  possession NOT NULL,

    -- Le joueur, quand il est identifié. Sinon la ligne de jeu
    -- ou le collectif portent l'action, et numero reste nul.
    numero      integer,
    groupe      text,          -- 'Avant', 'Trois-quarts', 'Collectif', '?'

    motif_code  text REFERENCES motif_faute(code),
    commentaire text NOT NULL DEFAULT '',

    saisi_le    timestamptz NOT NULL DEFAULT now(),

    -- Un joueur nommé ou un groupe, jamais les deux.
    CONSTRAINT action_porteur CHECK (numero IS NULL OR groupe IS NULL)
);

CREATE INDEX action_match_temps_idx ON action (match_id, temps);
CREATE INDEX action_match_type_idx ON action (match_id, type_code);
CREATE INDEX action_joueur_idx ON action (match_id, numero)
    WHERE numero IS NOT NULL;


-- Le détail d'une touche ou d'une mêlée, qui ne tient pas
-- dans la ligne du journal.
CREATE TABLE detail_phase (
    action_id       bigint PRIMARY KEY REFERENCES action(id) ON DELETE CASCADE,
    issue           issue NOT NULL,
    sauteur_numero  integer,
    zone_saut       text,
    alignement      text,
    ballon          text,          -- 'C' chaud, 'F' froid
    zone_longueur   text,
    zone_largeur    text,
    utilisation     text,
    observation     text NOT NULL DEFAULT ''
);


-- Les entrées et sorties. Le temps de jeu s'en déduit, sans
-- être stocké : il resterait à recalculer à la moindre
-- correction.
CREATE TABLE remplacement (
    id          bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    match_id    integer NOT NULL REFERENCES match(id) ON DELETE CASCADE,
    temps       numeric(9,3) NOT NULL CHECK (temps >= 0),
    mt          mi_temps NOT NULL,
    sortant     integer NOT NULL,
    entrant     integer NOT NULL,
    temporaire  boolean NOT NULL DEFAULT false,
    CONSTRAINT remplacement_distincts CHECK (sortant <> entrant)
);

CREATE INDEX remplacement_match_idx ON remplacement (match_id, temps);


-- Les cartons, avec leur retour éventuel.
CREATE TABLE carton (
    id          bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    match_id    integer NOT NULL REFERENCES match(id) ON DELETE CASCADE,
    temps       numeric(9,3) NOT NULL,
    mt          mi_temps NOT NULL,
    couleur     text NOT NULL CHECK (couleur IN ('rouge','jaune','blanc','bleu')),
    -- Nul quand le carton revient à l'adversaire : nous ne
    -- nommons pas ses joueurs.
    numero      integer,
    retour_temps numeric(9,3)
);

CREATE INDEX carton_match_idx ON carton (match_id, temps);
