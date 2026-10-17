# VeoVideoControl

> **À propos**
>
> VeoVideoControl permet au classeur **Createur de match.xlsm** de piloter automatiquement une vidéo ouverte dans **Google Chrome** (site Veo) ou dans **VLC**.
>
> Il synchronise les commandes vidéo (lecture, pause, avance, retour…) ainsi que le temps réel affiché dans le classeur Excel.

## Installation

### Prérequis

Avant la première utilisation, installer les logiciels suivants.

- Microsoft Excel 2019, 2021 ou Microsoft 365 avec prise en charge des macros VBA.
- VLC media player (dernière version stable), à télécharger : voir **logiciels à installer/LIENS.md**.
- Google Chrome (uniquement pour utiliser les vidéos Veo).
- Lors de la première ouverture, macOS peut demander une autorisation d'accès au dossier **VeoVideoControl** : il faut l'accepter.

L'ensemble du projet doit être copié sans modifier son arborescence. Tous les fichiers et dossiers doivent rester ensemble afin que VeoVideoControl fonctionne correctement.

L'arborescence du projet doit être la suivante :

```text
DOSSIER_A_UTILISER
│
├── Createur de match.xlsm
├── Matchs
│
└── VeoVideoControl
    ├── Documentation
    ├── Engine
    ├── Extension
    │   └── veo-sync-extension
    ├── VeoVideoControlEngine   (moteur Windows)
    ├── VeoVideoControl.app
    └── Arreter VeoVideoControl.app
```

Le dossier qui contient l'ensemble (ici `DOSSIER_A_UTILISER`) peut être placé où vous le souhaitez sur votre machine. En revanche, le dossier **VeoVideoControl** doit rester **à côté du classeur**, sans que son contenu interne soit modifié : le classeur va le chercher à cet emplacement.

## Installation de l'extension Chrome

1. Ouvrir Google Chrome.
2. Aller sur `chrome://extensions`.
3. Activer le **Mode développeur**.
4. Cliquer sur **Charger l'extension non empaquetée**.
5. Sélectionner le dossier :

```text
VeoVideoControl/Extension/veo-sync-extension
```

## Premier lancement sur macOS

Les deux applications du dossier `VeoVideoControl` ne sont pas signées par un développeur identifié par Apple. Au premier lancement, macOS les bloque. Comme c'est Excel qui les démarre, le blocage passe inaperçu : le lecteur vidéo ne se connecte jamais, sans message d'erreur explicite.

Il faut donc les ouvrir **une fois à la main**, avant d'ouvrir le classeur :

1. Ouvrir le dossier `VeoVideoControl` dans le Finder.
2. Clic droit sur **VeoVideoControl.app**, puis **Ouvrir**.
3. Dans la fenêtre d'avertissement, cliquer de nouveau sur **Ouvrir**.
4. Répéter l'opération sur **Arreter VeoVideoControl.app**, qu'Excel utilise à la fermeture du classeur.

> Si la fenêtre d'avertissement ne propose pas d'ouvrir l'application, aller dans **Réglages Système → Confidentialité et sécurité**, puis cliquer sur **Ouvrir quand même** en face du message qui mentionne l'application bloquée.

Cette manipulation n'est à faire qu'une seule fois.

## Autorisation d'accès macOS

Lors de la première ouverture du classeur après avoir déplacé le projet ou changé son emplacement, macOS peut afficher une fenêtre demandant à Microsoft Excel d'accéder au dossier `VeoVideoControl.app`.

Dans ce cas :

1. Cliquer sur **Sélectionner…**.
2. Dans la fenêtre du Finder, ne pas essayer de sélectionner directement `VeoVideoControl.app` : l'application apparaît grisée car macOS demande l'accès au dossier qui la contient.
3. Sélectionner le dossier **VeoVideoControl**.
4. Cliquer sur **Accorder l'accès**.

Cette autorisation ne devrait être demandée qu'une seule fois pour cet emplacement. Si le projet est déplacé, macOS peut redemander l'autorisation.

## Utilisation

1. Ouvrir la vidéo :
   - sur le site Veo dans Google Chrome ;
   - ou dans VLC.
2. Vérifier que la vidéo est prête à être lue.
3. Ouvrir `Createur de match.xlsm`.
4. Le moteur VeoVideoControl démarre automatiquement et se connecte au lecteur détecté.
5. Si aucun lecteur n'est détecté automatiquement, cliquer sur **Connecter lecteur vidéo**.


## Tutoriel de prise en main

### 1. Créer les listes du club

À faire une seule fois, en début de saison, dans **Createur de match.xlsm**.

Dans l'onglet **Listes** :

- ajouter chaque joueur de l'effectif avec le bouton **Ajouter un joueur** (nom et prénom) ;
- indiquer `oui` dans la colonne **1ère ligne** pour les joueurs concernés ;
- saisir les équipes de la poule dans la colonne **EQUIPES_A_AFFRONTER**.

> Ces deux listes alimentent les listes déroulantes de l'onglet **Compo** : l'équipe adverse et les noms de la composition. Elles sont recopiées dans chaque fichier de match créé ensuite, elles doivent donc être renseignées **avant** de créer un match.

### 2. Créer le fichier du match

1. Ouvrir **Createur de match.xlsm**.
2. Renseigner les informations du match dans l'onglet **Compo**.
3. Cliquer sur **Créer le fichier de match**.
4. Un nouveau fichier est créé dans le dossier **Matchs**.

> Une fois le fichier créé, travaillez uniquement sur ce fichier de match. Le classeur **Createur de match.xlsm** ne sert qu'à générer de nouveaux matchs.

### 3. Préparer le match

Dans l'onglet **Compo**, pour chacun des 23 postes :

- choisir le joueur dans la colonne **Nom du joueur** ;
- indiquer **Avant** ou **Trois-quarts** dans la colonne **Avant ou 3/4**.

> La colonne **temps de jeu** se remplit toute seule : elle est calculée à partir des remplacements saisis pendant la vidéo.

### 4. Lancer la saisie vidéo

1. Ouvrir l'onglet **Saisie vidéo**.
2. Vérifier que le lecteur vidéo est connecté.
3. Si le voyant n'est pas vert, cliquer sur **Connecter lecteur vidéo**.
4. Dès que le voyant devient vert, la saisie peut commencer.

> **Important (macOS)** : ne lancez pas VLC manuellement avant d'ouvrir Excel. Utilisez toujours **Connecter lecteur vidéo** pour que VeoVideoControl lance VLC avec la bonne configuration.

### 5. Afficher le Journal d'action

Pour suivre la saisie en direct :

1. **Affichage → Nouvelle fenêtre**
2. **Affichage → Réorganiser tout**
3. Afficher **Saisie vidéo** dans une fenêtre.
4. Afficher **Journal d'action** dans l'autre.

### 6. Consulter les statistiques

L'onglet **Stat** est entièrement automatique. Aucune saisie n'y est nécessaire.

### Bonnes pratiques

- **N'ouvrez qu'un seul fichier de match à la fois.** VeoVideoControl pilote un seul lecteur vidéo. Ouvrir plusieurs fichiers de match simultanément peut provoquer des conflits de pilotage ou de synchronisation.
- À la fermeture du fichier de match, **VeoVideoControl s'arrête automatiquement**. Aucune manipulation supplémentaire n'est nécessaire.
