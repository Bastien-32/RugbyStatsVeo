# Construire le moteur Windows

> Écrit pour la personne qui lance la construction sur le PC Windows.

Le moteur `VeoVideoControl` est du Python compilé en exécutable autonome par
PyInstaller. Chaque plateforme a le sien, et **chaque plateforme doit être
construite sur elle-même** : PyInstaller ne sait pas produire un exécutable
Windows depuis un Mac.

Tant que le moteur Windows n'est pas reconstruit, il fait tourner l'ancienne
version du code, même si `main.py` a changé.

## Ce qu'il faut

- Un **vrai PC Windows Intel ou AMD 64 bits**.
- **Python 3** installé depuis [python.org](https://www.python.org/downloads/),
  en cochant *Add python.exe to PATH* à l'installation.

### Sur un poste déjà installé

Inutile de retélécharger le paquet complet : il pèse près de 200 Mo, dont
128 d'installeurs VLC et 28 de moteur macOS, qui ne changent pas.

Deux choses suffisent, copiées depuis le Mac :

| À copier | Où le déposer |
|---|---|
| `Createur de match.xlsm` | remplace celui à la racine de l'installation |
| le dossier `Engine` | dans `VeoVideoControl\`, à côté de `VeoVideoControlEngine` |

Le moteur compilé, lui, ne voyage pas : il est refabriqué sur place par le
script ci-dessous, qui le dépose au bon endroit précisément parce que `Engine`
est à cet emplacement.

Pour transmettre aussi des fichiers de match, les glisser **un par un** dans
le dossier `Matchs` du poste — ne pas remplacer le dossier entier, ce qui
effacerait les matchs saisis sur place.

### Pourquoi pas une machine virtuelle sur le Mac

Un Mac Apple Silicon fait tourner un Windows **ARM64**. PyInstaller y produit
un exécutable ARM64, qui ne démarre pas sur les postes du club, tous en x64.

Le script refuse de tourner s'il ne détecte pas un Python `AMD64`, pour que
l'erreur se voie tout de suite plutôt qu'au premier match.

## La construction

Double-cliquer sur :

```
VeoVideoControl\Engine\construire-moteur-windows.bat
```

Le script vérifie Python et son architecture, crée un environnement isolé,
installe les dépendances, lance PyInstaller, puis met le résultat en place.

Compter une à deux minutes.

À la fin :

| Dossier | Contenu |
|---|---|
| `VeoVideoControl\VeoVideoControlEngine` | le moteur neuf |
| `VeoVideoControl\VeoVideoControlEngine-precedent` | celui d'avant, conservé |

L'ancien moteur n'est jamais supprimé : en cas de problème, il suffit
d'échanger les deux dossiers.

## Vérifier avant de repartir

1. Ouvrir un fichier de match du dossier `Matchs`.
2. Cliquer sur **Connecter lecteur vidéo** — le voyant doit passer au vert.
3. Tester la **lecture arrière continue** : c'est la commande `rewind_toggle`,
   ajoutée le 14/09/2026, et absente des moteurs construits avant cette date.
   C'est elle qui motive cette reconstruction.

## Puis

Rapporter le dossier `VeoVideoControlEngine` sur le Mac pour le versionner :
c'est le dépôt qui distribue le moteur aux autres postes.

## Si ça échoue

Le script s'arrête sur un message commençant par `[ERREUR]` et laisse la
fenêtre ouverte. Les cas courants :

| Message | Cause |
|---|---|
| `Python est introuvable` | Python absent, ou installé sans *Add to PATH*. |
| `Ce Python n'est pas en 64 bits Intel/AMD` | Windows ARM, ou Python 32 bits. |
| `L'installation des dependances a echoue` | Pas d'accès réseau, ou proxy d'entreprise. |

Pour tout autre message, garder le texte affiché : il dit à quelle étape la
construction s'est arrêtée.
