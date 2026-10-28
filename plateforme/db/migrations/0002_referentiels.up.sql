-- =========================================================
-- Le vocabulaire du jeu
--
-- Repris du classeur « Createur de match.xlsm », au mot près.
-- C'est une saison d'analyse qui l'a façonné : il décrit ce
-- qu'un entraîneur regarde vraiment, pas ce qu'un modèle de
-- données suggère.
-- =========================================================

INSERT INTO famille_action (code, libelle, rang) VALUES
    ('POINTS',      'Points',                   1),
    ('REPRISE',     'Reprise du jeu, ou arrêt', 2),
    ('CARTON',      'Cartons',                  3),
    ('TOUCHE',      'Touches',                  4),
    ('MELEE',       'Mêlées',                   5),
    ('PLAQUAGE',    'Plaquages',                6),
    ('ATTITUDE',    'Attitudes au contact',     7),
    ('POSSESSION',  'Changement de possession', 8),
    ('PENALITE',    'Pénalité',                 9),
    ('COUP_FRANC',  'Coup franc',              10);


-- ---------------------------------------------------------
-- Points
--
-- L'essai et le drop sont suivis d'un renvoi : la possession
-- bascule, et le jeu s'arrête le temps de se replacer.
-- ---------------------------------------------------------
INSERT INTO type_action
    (code, famille, libelle, rang, bascule, arret_auto, attend_joueur) VALUES
    ('PTS_ESSAI',      'POINTS', 'Essai',             1, true,  true,  true),
    ('PTS_ESSAI_PEN',  'POINTS', 'Essai de pénalité', 2, true,  true,  false),
    ('PTS_TRANSFO',    'POINTS', 'Transformation',    3, false, false, true),
    ('PTS_PENALITE',   'POINTS', 'Pénalité',          4, false, false, true),
    ('PTS_DROP',       'POINTS', 'Drop',              5, true,  true,  true);


-- ---------------------------------------------------------
-- Reprise du jeu
--
-- Rien ne se saisit tant que le jeu n'a pas repris : c'est
-- ce qui garantit que les séquences de possession se ferment
-- proprement.
-- ---------------------------------------------------------
INSERT INTO type_action
    (code, famille, libelle, rang, ouvre_jeu, ferme_jeu) VALUES
    ('REP_TOUCHE',   'REPRISE', 'Reprise du jeu sur touche',  1, true,  false),
    ('REP_MELEE',    'REPRISE', 'Reprise du jeu sur melée',   2, true,  false),
    ('REP_RENVOI',   'REPRISE', 'Reprise du jeu sur renvoi',  3, true,  false),
    ('REP_MAIN',     'REPRISE', 'Reprise du jeu à la main',   4, true,  false),
    ('ARRET',        'REPRISE', 'Arret du jeu',               5, false, true);


-- ---------------------------------------------------------
-- Cartons
--
-- Le camp se lit au joueur désigné, jamais à la possession :
-- après un essai adverse, le ballon est déjà revenu à nous
-- alors que le carton est pour notre camp.
-- ---------------------------------------------------------
INSERT INTO type_action (code, famille, libelle, rang, attend_joueur) VALUES
    ('CAR_ROUGE_NS',  'CARTON', 'Carton rouge contre nous',        1, true),
    ('CAR_ROUGE_ADV', 'CARTON', 'Carton rouge contre adversaire',  2, false),
    ('CAR_JAUNE_NS',  'CARTON', 'Carton jaune contre nous',        3, true),
    ('CAR_JAUNE_ADV', 'CARTON', 'Carton jaune contre adversaire',  4, false),
    ('CAR_BLANC_NS',  'CARTON', 'Carton blanc contre nous',        5, true),
    ('CAR_BLANC_ADV', 'CARTON', 'Carton blanc contre adversaire',  6, false),
    ('CAR_BLEU_NS',   'CARTON', 'Carton bleu contre nous',         7, true),
    ('CAR_BLEU_ADV',  'CARTON', 'Carton bleu contre adversaire',   8, false),
    ('RET_ROUGE_NS',  'CARTON', 'Retour de carton rouge - nous',   9, true),
    ('RET_ROUGE_ADV', 'CARTON', 'Retour de carton rouge - adversaire', 10, false),
    ('RET_JAUNE_NS',  'CARTON', 'Retour de carton jaune - nous',  11, true),
    ('RET_JAUNE_ADV', 'CARTON', 'Retour de carton jaune - adversaire', 12, false),
    ('RET_BLANC_NS',  'CARTON', 'Retour de carton blanc - nous',  13, true),
    ('RET_BLANC_ADV', 'CARTON', 'Retour de carton blanc - adversaire', 14, false),
    ('RET_BLEU_NS',   'CARTON', 'Retour de carton bleu - nous',   15, true),
    ('RET_BLEU_ADV',  'CARTON', 'Retour de carton bleu - adversaire',  16, false);


-- ---------------------------------------------------------
-- Touches et mêlées
--
-- Deux jeux de libellés selon qui la joue, et l'issue dit ce
-- que devient le ballon : G le rend à l'équipe qui lance, P
-- le laisse à l'autre, ND fait rejouer la touche.
-- ---------------------------------------------------------
INSERT INTO type_action (code, famille, libelle, rang, attend_joueur) VALUES
    ('TO_NS_G',   'TOUCHE', 'Touche à nous conservée',  1, true),
    ('TO_NS_P',   'TOUCHE', 'Touche à nous perdue',     2, true),
    ('TO_NS_ND',  'TOUCHE', 'Touche à nous pas droite', 3, true),
    ('TO_ADV_G',  'TOUCHE', 'Touche à eux volée',       4, false),
    ('TO_ADV_P',  'TOUCHE', 'Touche à eux perdue',      5, false),
    ('TO_ADV_ND', 'TOUCHE', 'Touche à eux pas droite',  6, false);

INSERT INTO type_action (code, famille, libelle, rang) VALUES
    ('ME_NS_G',  'MELEE', 'Mêlée à nous gagnée', 1),
    ('ME_NS_P',  'MELEE', 'Mêlée à nous perdue', 2),
    ('ME_ADV_G', 'MELEE', 'Mêlée à eux volée',   3),
    ('ME_ADV_P', 'MELEE', 'Mêlée à eux perdue',  4);


-- ---------------------------------------------------------
-- Plaquages
--
-- Ils se saisissent quand l'adversaire a le ballon : c'est
-- nous qui plaquons, et le joueur est attendu.
-- ---------------------------------------------------------
INSERT INTO type_action (code, famille, libelle, rang, attend_joueur) VALUES
    ('PL_OFF',    'PLAQUAGE', 'Plaquage offensif', 1, true),
    ('PL_NORMAL', 'PLAQUAGE', 'Plaquage normal',   2, true),
    ('PL_RATE',   'PLAQUAGE', 'Plaquage raté',     3, true),
    ('PL_A2',     'PLAQUAGE', 'Plaquage à 2',      4, true),
    ('PL_HAUT',   'PLAQUAGE', 'Plaquage haut',     5, true);


-- ---------------------------------------------------------
-- Attitudes au contact
--
-- Ce que devient le ballon porté : il avance, il est stoppé,
-- il franchit.
-- ---------------------------------------------------------
INSERT INTO type_action (code, famille, libelle, rang, attend_joueur) VALUES
    ('ATT_FRANCHISSEMENT', 'ATTITUDE', 'Franchissement',            1, true),
    ('ATT_AVANCEE',        'ATTITUDE', 'Avancée au contact',        2, true),
    ('ATT_NULLE',          'ATTITUDE', 'Recul ou sans progression', 3, true);


-- ---------------------------------------------------------
-- Changement de possession
--
-- Le même geste porte deux noms selon qui tenait le ballon :
-- gratter, c'est récupérer quand ils l'avaient, le perdre
-- quand nous l'avions. Le journal doit dire lequel.
-- ---------------------------------------------------------
INSERT INTO type_action
    (code, famille, libelle, rang, bascule, attend_joueur) VALUES
    ('EN_AVANT',        'POSSESSION', 'En avant',                            1, true,  true),
    ('JEU_AU_PIED',     'POSSESSION', 'Jeu au pied',                         2, false, true),
    ('RECEPTION_NS',    'POSSESSION', 'Réception nous',                      3, false, true),
    ('RECEPTION_ADV',   'POSSESSION', 'Réception adversaire',                4, false, false),
    ('GRATTAGE_REC',    'POSSESSION', 'Grattage - récupération du ballon',   5, true,  true),
    ('GRATTAGE_PERTE',  'POSSESSION', 'Grattage - perte du ballon',          6, true,  true),
    ('CRUCK_REC',       'POSSESSION', 'Contre-ruck - récupération du ballon',7, true,  true),
    ('CRUCK_PERTE',     'POSSESSION', 'Contre-ruck - perte du ballon',       8, true,  true),
    ('ARRACHAGE_REC',   'POSSESSION', 'Arrachage - récupération du ballon',  9, true,  true),
    ('ARRACHAGE_PERTE', 'POSSESSION', 'Arrachage - perte du ballon',        10, true,  true),
    ('TO_REC',          'POSSESSION', 'Turn-over - récupération du ballon', 11, true,  true),
    ('TO_PERTE',        'POSSESSION', 'Turn-over - perte du ballon',        12, true,  true);


-- ---------------------------------------------------------
-- Fautes
--
-- Pénalité comme coup franc arrêtent le jeu : la ligne
-- d'arrêt suit, et le temps de jeu effectif reste juste.
-- ---------------------------------------------------------
INSERT INTO type_action
    (code, famille, libelle, rang, arret_auto, attend_joueur) VALUES
    ('PEN_CONTRE_ADV', 'PENALITE',   'Pénalité contre adversaire',   1, true, false),
    ('PEN_CONTRE_NS',  'PENALITE',   'Pénalité contre nous',         2, true, true),
    ('CF_CONTRE_ADV',  'COUP_FRANC', 'Coup-franc contre adversaire', 1, true, false),
    ('CF_CONTRE_NS',   'COUP_FRANC', 'Coup-franc contre nous',       2, true, true);


-- =========================================================
-- Motifs de faute
--
-- Rangés par famille, comme la popup les présente. Un motif
-- peut servir aux deux types de faute : « Passage à vide »
-- et « 10m de plus - contestation » en sont.
-- =========================================================

INSERT INTO motif_faute (code, libelle, famille, penalite, coup_franc, rang) VALUES
    ('EA_VOLONTAIRE',   'En-avant volontaire',        'En-avant',  true,  false,  1),
    ('EA_REPRIS',       'En-avant repris devant',     'En-avant',  true,  false,  2),

    ('MAUL_COTE',       'Maul - Entrer sur le côté',  'Maul',      true,  false,  1),
    ('MAUL_ECROUL',     'Maul - Ecroulement',         'Maul',      true,  false,  2),

    ('ME_POUSSEE_AV',   'Melee - Poussee avant introduction',  'Melee', true,  false, 1),
    ('ME_MISE_PLACE',   'Melee - Mise en place trop longue',   'Melee', false, true,  2),
    ('ME_ENGAGEMENT',   'Melee - Engagement anticipé',         'Melee', false, true,  3),
    ('ME_POUSSEE_ANT',  'Melee - Poussée anticipée',           'Melee', false, true,  4),
    ('ME_INTRO',        'Melee - Introduction irreguliere',    'Melee', false, true,  5),
    ('ME_TALONNAGE',    'Melee - Talonnage irregulier',        'Melee', false, true,  6),
    ('ME_POUSSEE_TAL',  'Melee - Poussée après talonnage',     'Melee', false, true,  7),

    ('TO_SAISIE_BRAS',  'Touche - Saisie bras sauteur',              'Touche', true,  false, 1),
    ('TO_PLAQ_SAUTEUR', 'Touche - Plaquage sauteur avant retombée',  'Touche', true,  false, 2),
    ('TO_POUSSEE',      'Touche - Poussée avant retombée sauteur',   'Touche', true,  false, 3),
    ('TO_MISE_PLACE',   'Touche - Mise en place trop longue',        'Touche', false, true,  4),
    ('TO_ALIGNEMENT',   'Touche - Alignement irregulier',            'Touche', false, true,  5),
    ('TO_LANCER',       'Touche - Lancer non conforme',              'Touche', false, true,  6),
    ('TO_SAUT_PREM',    'Touche - Saut ou soutien premature',        'Touche', false, true,  7),
    ('TO_REDESCENTE',   'Touche - Sauteur mal redescendu',           'Touche', false, true,  8),
    ('TO_SOUTIEN',      'Touche - Soutient devant sauteur retombée', 'Touche', false, true,  9),
    ('TO_CINQ_M',       'Touche - Lancer dans les 5 m',              'Touche', false, true, 10),

    ('RUCK_AU_DELA',    'Ruck - Va au-delà',                   'Ruck', true,  false,  1),
    ('RUCK_COUCHE',     'Ruck - Soutient couché sur porteur',  'Ruck', true,  false,  2),
    ('RUCK_RETARD',     'Ruck - Retard soutient',              'Ruck', true,  false,  3),
    ('RUCK_GARDE',      'Ruck - Garde le ballon au sol',       'Ruck', true,  false,  4),
    ('RUCK_DEBLAYAGE',  'Ruck - Déblayage sur le coté',        'Ruck', true,  false,  5),
    ('RUCK_RELAYEUR',   'Ruck - Saisie relayeur',              'Ruck', true,  false,  6),
    ('RUCK_TALONNAGE',  'Ruck - Talonnage à la main',          'Ruck', true,  false,  7),
    ('RUCK_PLAQUEUR',   'Ruck - Plaqueur qui ne sort pas',     'Ruck', true,  false,  8),
    ('RUCK_4_APPUIS',   'Ruck - 4 appuis',                     'Ruck', true,  false,  9),
    ('RUCK_GRATTEUR',   'Ruck - Gratteur ne relâche pas',      'Ruck', true,  false, 10),
    ('RUCK_FAUSSE_SOR', 'Ruck - Fausse sortie',                'Ruck', false, true,  11),

    ('PLAQ_A2',         'Plaquage - à 2',             'Plaquage', true, false, 1),
    ('PLAQ_HAUT',       'Plaquage - haut',            'Plaquage', true, false, 2),
    ('PLAQ_SANS_BAL',   'Plaquage - sans ballon',     'Plaquage', true, false, 3),
    ('PLAQ_RETARD',     'Plaquage - à retardement',   'Plaquage', true, false, 4),
    ('PLAQ_AIR',        'Plaquage - en l''air',       'Plaquage', true, false, 5),

    ('BRUTALITE',       'Brutalité',                  'Autres', true,  false,  1),
    ('ROULE_SOL',       'Porteur roule au sol',       'Autres', true,  false,  2),
    ('PASSAGE_VIDE',    'Passage à vide',             'Autres', true,  true,   3),
    ('HORS_JEU',        'Hors-jeu',                   'Autres', true,  false,  4),
    ('DIX_M',           '10m de plus',                'Autres', true,  false,  5),
    ('DIX_M_CONTEST',   '10m de plus - contestation', 'Autres', true,  true,   6),
    ('PARLE_ARBITRE',   'Parle arbitre',              'Autres', true,  false,  7),
    ('PERTE_TEMPS',     'Perte de temps',             'Autres', false, true,   8),
    ('JEU_MAIN',        'Jeu à la main - mauvaise remise en jeu', 'Autres', false, true, 9),
    ('ECRAN',           'Ecran',                      'Autres', false, true,  10),
    ('JEU_RAPIDE',      'Jeu-rapide 10m de plus',     'Autres', false, true,  11),
    ('AUTRE',           'Autre',                      'Autres', true,  true,  98),
    ('INCONNU',         '?',                          'Autres', true,  true,  99);
