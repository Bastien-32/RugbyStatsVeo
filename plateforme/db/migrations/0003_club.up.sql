-- =========================================================
-- Le club et son effectif
--
-- Un seul club pour l'instant, celui qui a servi à bâtir le
-- modèle. Les requêtes sont déjà découpées par club_id : en
-- ajouter d'autres ne demandera pas de reprendre le schéma.
-- =========================================================

INSERT INTO club (nom) VALUES ('Masseube XV');

INSERT INTO joueur (club_id, nom, prenom, ligne) VALUES
    (1, 'ABADIE',            'Quentin',      'Avant'),
    (1, 'ABADIE',            'Emeric',       'Avant'),
    (1, 'BARATGIN',          'Julian',       'Avant'),
    (1, 'BAUP',              'Nicolas',      'Avant'),
    (1, 'BIFFI',             'Pierre',       'Avant'),
    (1, 'BOURREC',           'Anthony',      'Trois-quarts'),
    (1, 'CESCA',             'Sacha',        'Avant'),
    (1, 'CESCA',             'Simon',        'Trois-quarts'),
    (1, 'COURREGES',         'Romain',       'Trois-quarts'),
    (1, 'DALLA ROSA',        'Jonah',        'Avant'),
    (1, 'DUTREY',            'Thomas',       'Avant'),
    (1, 'GALLINARO',         'Pierrick',     'Trois-quarts'),
    (1, 'JOUNAUX',           'Joris',        'Trois-quarts'),
    (1, 'KAPFER',            'Bastien',      'Trois-quarts'),
    (1, 'LAUDET',            'Jean Baptiste','Trois-quarts'),
    (1, 'LOUREIRO LABAZUY',  'Antonio',      'Avant'),
    (1, 'MOUDENS',           'Alexis',       'Trois-quarts'),
    (1, 'MULLIER',           'Martin',       'Avant'),
    (1, 'MULLIER',           'Victor',       'Avant'),
    (1, 'POUJOLS',           'Dylan',        'Trois-quarts'),
    (1, 'RICHIARDONNE',      'Gregory',      'Avant'),
    (1, 'ROBERT',            'Jimmy',        'Trois-quarts'),
    (1, 'SALLES',            'Dorian',       'Trois-quarts'),
    (1, 'SALLES',            'Fabien',       'Trois-quarts'),
    (1, 'SOUVILLE',          'Quentin',      'Avant'),
    (1, 'STEFANI',           'Titouan',      'Avant');
