# Saisie vidéo rugby — démonstration web

Maquette fonctionnelle d'une saisie statistique de match sur vidéo,
reprenant le modèle métier du classeur `Createur de match.xlsm`.

Elle existe pour répondre à une question : ce que le classeur fait
aujourd'hui dans Excel, avec un moteur local de 55 Mo et une extension
de navigateur, peut-il se faire dans une page web, sans rien installer ?

## Ouvrir

Double-cliquer sur `saisie-rugby.html`. C'est un fichier unique, sans
serveur ni dépendance : style, code et vocabulaire métier sont dedans.

Déposer ensuite une vidéo de match, ou la choisir avec le bouton. Elle
reste sur le poste ; rien n'est envoyé nulle part.

## Ce que la page démontre

Le vocabulaire et les règles viennent du classeur, au mot près :

- les libellés d'actions, des plaquages aux cartons ;
- les motifs de pénalité et de coup franc, rangés par famille ;
- les issues de touche `G` / `P` / `ND` et de mêlée `G` / `P`, avec
  zone de saut, alignement, ballon et utilisation ;
- la bascule de possession : `G` rend le ballon, `P` le laisse, `ND` le
  rejoue à l'autre équipe ;
- l'état du jeu, qui n'ouvre la palette qu'après une reprise, et se
  juge à la position courante dans la vidéo — revenir en arrière pour
  compléter une phase ne bloque pas la saisie ;
- le camp d'un carton lu au joueur désigné, non à la possession ;
- les cinq tranches de durée de possession.

Trois vues reliées : la saisie, la feuille de match, les statistiques.
Le temps de jeu se calcule depuis les remplacements, et chaque ligne du
journal renvoie à son instant dans la vidéo.

## Ce qu'elle ne fait pas

- Rien n'est enregistré entre deux sessions : tout vit dans l'onglet.
- Pas de récapitulatif multi-matchs.
- L'effectif du club est écrit en tête du script. Il vivrait ailleurs
  dans une vraie application : une page commune à toutes les équipes,
  où la feuille de match viendrait puiser.

## Pourquoi un seul fichier

Pour qu'il s'ouvre partout, se lise d'un bout à l'autre et se reprenne
sans outillage. Une application réelle découperait le tout ; ici, la
contrainte sert la démonstration.
