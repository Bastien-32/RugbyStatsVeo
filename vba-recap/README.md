# Récapitulatif de saison

Classeur qui rassemble les statistiques de tous les matchs, pour établir un
classement des joueurs et suivre leur temps de jeu, en séparant la Première de
la Réserve.

Le classeur garde chez lui une copie du journal d'actions et des compositions.
**Un match déjà intégré n'est relu que si son fichier a changé** :
l'actualisation dure quelques secondes, même en fin de saison.

Les fichiers de match ne sont jamais modifiés : ils sont ouverts en lecture
seule, puis refermés sans enregistrer.

## Le périmètre : l'emplacement, et rien d'autre

Seuls les fichiers posés **à la racine** du dossier `Matchs` sont regardés. Le
balayage ne descend dans aucun sous-dossier.

C'est le seul critère. Un match rangé dans un sous-dossier n'entre pas dans la
base ; un match posé à la racine y entre, amical compris. Pour comparer un
amical au reste de la saison, il suffit donc de le déposer à la racine : le
filtre **Phase** de la feuille Filtres permettra ensuite de l'isoler ou de
l'écarter.

**Un match = un fichier.** C'est le fichier fusionné, qui porte le match
entier, qu'on dépose à la racine — pas les deux fichiers de mi-temps. Le
classement ne distingue pas MT1 de MT2 : les statistiques d'un match sont le
total des deux périodes.

Si deux fichiers **différents** portent le même `ID match`, le second est
ignoré et signalé, pour que les statistiques ne soient pas comptées deux fois.

## Comment un fichier modifié est repéré

À chaque actualisation, le classeur compare, pour chaque fichier de la racine :

| Situation | Ce qui se passe |
|---|---|
| Fichier inconnu | Il est lu et ajouté. |
| Fichier connu, date de modification inchangée | Il n'est **pas** ouvert. |
| Fichier connu, modifié depuis | Il est relu, et **toutes** ses lignes sont remplacées. |

La date de modification se lit **sans ouvrir le fichier** : c'est ce qui permet
de passer une saison entière en revue en un instant. Elle est notée dans la
colonne *Fichier modifié le* de la feuille Matchs, au moment de la lecture.

**C'est un remplacement, jamais une fusion.** Le classeur ne cherche pas quelles
lignes ont été ajoutées : il efface les anciennes lignes du match et réécrit
tout. C'est ce qui garantit qu'aucune ligne n'est doublée — et cela reste juste
même si, en reprenant la saisie, vous avez corrigé ou supprimé des lignes
existantes plutôt que d'en ajouter à la fin.

Une tolérance de deux secondes évite les relectures inutiles dues aux arrondis
d'horloge. Une synchronisation iCloud qui retoucherait la date sans changer le
contenu provoque au pire une relecture pour rien, sans aucun effet sur les
chiffres.

## Installation

1. Dans Excel, créer un classeur vierge et l'enregistrer au format
   **classeur Excel prenant en charge les macros (.xlsm)** sous
   `Matchs/Recapitulatif saison.xlsm`.

   Le classeur doit être **dans** le dossier `Matchs` : c'est son propre
   dossier qu'il parcourt.

2. Ouvrir l'éditeur VBA (`Outils → Macro → Éditeur Visual Basic`).

3. Pour chacun des quatre fichiers de ce dossier, dans l'ordre :
   `modRecapBase.bas`, `modRecapConstruction.bas`, `modRecapImport.bas`,
   `modRecapCalcul.bas` — faire `Fichier → Importer un fichier…` et le
   sélectionner.

   (Ou, si l'import de fichier n'est pas disponible : `Insertion → Module`,
   puis coller le contenu du fichier **sans** sa première ligne
   `Attribute VB_Name = "…"`, et renommer le module dans la fenêtre
   Propriétés.)

4. Revenir dans Excel, `Outils → Macro → Macros`, choisir
   **ConstruireClasseurRecapitulatif** et exécuter.

   Les sept feuilles sont créées avec leurs boutons.

5. Enregistrer le classeur.

## Utilisation

### Après un match

Feuille **Journal actions** (ou **Matchs**) → bouton
**Ajouter les nouveaux matchs**.

Le même bouton sert dans les deux cas : nouveaux matchs et matchs dont la
saisie a été complétée. macOS demande une fois l'autorisation d'accès :
l'accepter.

Penser à **enregistrer le classeur** ensuite : c'est lui qui porte la base.

### Les deux boutons de rattrapage

Feuille **Matchs** → cliquer sur la ligne du match, puis :

- **Réimporter ce match** — force la relecture, même si la date du fichier n'a
  pas bougé. Utile si vous avez corrigé des minutes à la main et voulez
  repartir de ce que dit le fichier.
- **Retirer ce match** — le sort de la base sans le relire. Le fichier de
  match, lui, n'est pas touché ; il reviendra à la prochaine actualisation
  s'il est toujours à la racine.

### Filtrer

Feuille **Filtres**, puis bouton **Appliquer les filtres**.

- **Saison, Catégorie, Phase, Lieu, Résultat, Adversaire** : menus alimentés
  par ce qui existe réellement dans la base.
- **Matchs joués à partir du / jusqu'au** : bornes de dates.
- **N derniers matchs** : `3` pour les trois derniers. S'applique **après** les
  autres critères, donc « catégorie = Réserve » + « 3 derniers » donne bien les
  trois derniers matchs de réserve.
- **Matchs ciblés uniquement** : sur `Oui`, seuls les matchs cochés *Cibler*
  dans la feuille Matchs comptent, et les autres critères sont ignorés.
- **Minutes jouées minimum** : écarte du classement les joueurs sous le seuil.

Le rappel des filtres s'affiche au-dessus de chaque tableau.

### Les trois lectures du classement

Feuille **Classement**, bouton **Mode : …** — un clic passe à la suivante.

| Mode | Lecture |
|---|---|
| **Totaux** | Nombre d'actions sur la période. |
| **Par match** | Divisé par le nombre de matchs joués. |
| **Par 80 minutes** | Ramené au temps de jeu, pour comparer un titulaire et un finisseur. |

Les colonnes d'identité (feuilles, matchs joués, minutes) restent toujours en
valeurs réelles : ce sont elles qui expliquent les ratios.

Les colonnes sont triables : cliquer sur la flèche d'un en-tête.

### Choisir les statistiques affichées

Feuille **Parametres**, tableau `RECAP_STATS`.

- Décocher **Afficher** retire la colonne du classement.
- L'ordre des lignes est l'ordre des colonnes.
- **Action** doit reprendre exactement le libellé du journal
  (voir `Paramètres!Q:R` d'un fichier de match).
- **Motif** vide compte toutes les occurrences de l'action ; renseigné, il ne
  compte que ce motif. C'est ainsi que « Pénalités concédées » totalise ce que
  « Pén. ruck » détaille.

Cinq lignes de détail des pénalités sont livrées décochées.

## Le temps de jeu

Il vient de la colonne **temps de jeu** de la feuille **Compo** de chaque
fichier de match, saisie par l'entraîneur. Rien ne la calcule automatiquement :
le journal d'actions ne contient ni entrée ni sortie de joueur.

Deux écritures sont acceptées : un nombre de minutes (`80`) ou une durée
(`1:20`).

Un oubli peut se rattraper **directement dans la feuille Compositions** du
récapitulatif, en écrivant la valeur dans la colonne *Minutes* : inutile de
rouvrir le fichier de match.

Attention : cette correction est perdue si le match est relu — soit parce que
vous avez cliqué sur *Réimporter ce match*, soit parce que le fichier de match
a été rouvert et réenregistré entre-temps. Pour une correction durable, la
saisir dans la feuille Compo du fichier de match.

Une ligne de composition sans temps de jeu compte comme une **feuille de
match** mais pas comme un **match joué** : c'est un remplaçant resté sur le
banc. Un joueur sans minute saisie apparaît donc au classement en mode
*Totaux*, mais ses cases restent vides en mode *Par 80 minutes*.

## Les sept feuilles

| Feuille | Contenu |
|---|---|
| **Classement** | Une ligne par joueur, une colonne par statistique. |
| **Temps de jeu** | Matchs, minutes et moyenne, par catégorie. |
| **Filtres** | Les critères, et le mode d'affichage. |
| **Matchs** | Les matchs de la base : inventaire, ciblage, réimport, retrait. |
| **Journal actions** | La base : une ligne par action, tous matchs confondus. |
| **Compositions** | La base : une ligne par joueur et par match, avec les minutes. |
| **Parametres** | Statistiques affichées, et listes des menus. |

**Journal actions** et **Compositions** sont les deux feuilles qui portent les
données. Elles restent visibles : on peut y brancher un tableau croisé
dynamique monté à la main. La colonne *Mi-temps* du journal y est conservée
pour cet usage, même si le classement ne s'en sert pas.
