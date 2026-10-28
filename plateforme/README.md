# RugbyStats — plateforme

Analyse vidéo de match : le modèle du classeur `Createur de match.xlsm`,
porté sur une base de données, une API et une interface web.

Trois conteneurs, aucune installation sur le poste.

## Démarrer

```sh
cp .env.exemple .env
docker compose up --build
```

Puis <http://localhost:5173>.

La base garde ses données dans un volume nommé : arrêter les conteneurs
ne les efface pas. Pour repartir de zéro, `docker compose down -v`.

Si le port 5434 est déjà pris, changer `DB_PORT` dans `.env`.

## Les trois services

| service | rôle | image finale |
|---|---|---|
| `db` | PostgreSQL 16, les données | officielle |
| `api` | Go, les règles et les calculs | ~15 Mo, binaire statique |
| `web` | React servi par nginx | ~50 Mo |

Un quatrième conteneur, `migrations`, applique le schéma puis s'arrête.
L'API ne démarre pas avant qu'il ait fini.

nginx sert aussi de passerelle : le front appelle `/api`, jamais une
adresse de serveur. Rien à reconfigurer d'un environnement à l'autre.

## Ce que la base contient

Le vocabulaire du jeu vit en base, pas dans le code :

- `famille_action`, `type_action` — les actions saisissables, et ce que
  chacune fait au jeu : ouvrir, fermer, faire basculer la possession,
  attendre un joueur ;
- `motif_faute` — les motifs de pénalité et de coup franc, par famille ;
- `joueur` — l'effectif du club, commun à toutes les équipes ;
- `match`, `feuille_match` — qui porte quel numéro, dans quelle ligne,
  ce jour-là ;
- `action`, `detail_phase` — le journal, horodaté sur la vidéo ;
- `remplacement`, `carton`.

Un club peut donc ajouter ses propres motifs sans qu'on recompile quoi
que ce soit.

## Ce que l'API calcule

Rien n'est stocké qui puisse se déduire : le temps de jeu et les durées
de possession se recalculent à la demande, et restent justes après la
moindre correction du journal.

Les règles sont dans `api/internal/modele/regles.go` :

- **l'état du jeu** à un instant donné de la vidéo — seule compte la
  dernière reprise ou le dernier arrêt qui le précède, pour qu'un
  retour en arrière rende l'état de cet instant-là ;
- **les séquences de possession**, d'une reprise à l'arrêt ou au
  changement de camp, rangées dans les cinq tranches du classeur ;
- **le temps de jeu**, reconstitué depuis les passages sur le terrain :
  un joueur sorti puis rentré cumule ses deux périodes.

Le typage ferme le reste : une possession ne peut valoir que `Nous` ou
`Adv`, une issue de touche que `G`, `P` ou `ND`. Le compilateur refuse
ce qui n'appartient pas au vocabulaire.

## Points d'entrée

```
GET    /sante
GET    /referentiels                  tout le vocabulaire, en un appel

GET    /joueurs                       POST /joueurs
GET    /matchs                        POST /matchs
GET    /matchs/{id}
GET    /matchs/{id}/feuille           PUT  /matchs/{id}/feuille
GET    /matchs/{id}/actions           POST /matchs/{id}/actions
DELETE /actions/{id}
GET    /matchs/{id}/remplacements     POST /matchs/{id}/remplacements
GET    /matchs/{id}/statistiques
```

## Développer

Le front en rechargement à chaud, l'API et la base dans Docker :

```sh
docker compose up -d db api
cd web && npm install && npm run dev
```

Vite fait lui-même la passerelle vers `localhost:8080`, le front appelle
donc `/api` comme en production.

Pour l'API, `cd api && go run .` avec `BASE_URL` pointant sur la base
exposée en 5434.

## L'interface

Trois vues, reliées entre elles.

**Saisie** — la vidéo et le journal à gauche, les palettes à droite sur
toute la hauteur. Trois poignées pour ajuster les zones : largeur des
palettes, hauteur de la vidéo, hauteur de la palette joueurs. Au
clavier, `Espace` pour lire, `←` `→` pour ±5 s.

La palette joueurs montre qui est sur le terrain à cet instant de la
vidéo, remplaçants compris, et le numéro porté est celui du maillot :
un remplaçant entre avec le sien. Un joueur sous carton garde sa
couleur et ne répond plus jusqu'à son retour.

La palette d'actions se construit depuis la base. Tant que le jeu n'a
pas repris, seules les reprises et l'arrêt répondent.

Cinq popups : touche, mêlée, pénalité, coup franc, remplacement, retour
de carton. Rien n'y est présélectionné — une valeur ne part au journal
que si elle a été choisie.

**Composition** — la feuille de match, puisée dans l'effectif du club
par saisie assistée. La ligne d'un titulaire découle de son poste ;
celle d'un remplaçant se déclare, un banc pouvant compter trois avants
et quatre trois-quarts comme l'inverse. C'est cette ligne qui décide
qui la popup touche propose comme sauteur.

**Statistiques** — possession par tranches, production par joueur,
détail des touches et mêlées. Un clic sur l'heure d'une ligne ramène la
vidéo à cet instant.

## Ce qui reste à faire

- Le récapitulatif multi-matchs, qui existe dans le classeur.
- Le sens de jeu, qui oriente le terrain des popups.
- L'authentification. Les requêtes sont déjà découpées par club ; il
  manque de savoir qui est connecté.
- La vidéo est chargée depuis le poste. Chez un hébergeur, elle
  viendrait du lecteur de la plateforme, et `match.video_ref` est là
  pour ça.
