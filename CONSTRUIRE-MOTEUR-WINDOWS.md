# Construire le moteur Windows, sans internet

Tout ce qu'il faut est déjà dans ce dossier. Seul Python est à installer,
et son installeur est fourni. Compter 10 à 15 minutes.

## 1. Vérifier le PC

**Paramètres › Système › Informations › Type du système** doit afficher
« Système d'exploitation 64 bits, processeur x64 ».

S'il est écrit « ARM », ce PC ne convient pas : arrêter là.

## 2. Dézipper STAT_VEO

Dézipper l'archive **directement dans `C:\`**, pour obtenir `C:\STAT_VEO`.

Pas sur le Bureau : les chemins y deviennent trop longs et Windows refuse
d'extraire certains fichiers.

## 3. Installer Python

Double-cliquer sur `logiciels à installer\python-3.13.17-amd64.exe`.

- Cocher en bas **« Add python.exe to PATH »**. C'est indispensable.
- Sans droits administrateur, décocher aussi « Use admin privileges when
  installing py.exe ».
- Cliquer sur **Install Now**, puis sur **Close** à la fin. Ignorer
  « Disable path length limit ».

## 4. Construire le moteur

Double-cliquer sur
`VeoVideoControl\Engine\construire-moteur-windows.bat`.

Si Windows affiche « Windows a protégé votre ordinateur » : cliquer sur
**Informations complémentaires**, puis sur **Exécuter quand même**.

La fenêtre doit afficher :

```
Architecture detectee : AMD64
Mode hors ligne : dependances lues dans le dossier wheels.
```

Une à deux minutes plus tard, elle affiche « Moteur Windows construit ».
Appuyer sur une touche pour la fermer.

Si l'antivirus signale `VeoVideoControl.exe`, c'est un faux positif
courant avec PyInstaller. Autoriser le fichier, puis relancer le script.

## 5. Rapporter le résultat

Copier sur une clé USB le dossier
`VeoVideoControl\VeoVideoControlEngine`. Il doit contenir
`VeoVideoControl.exe` et un dossier `_internal`.

C'est la seule chose à rapporter sur le Mac.

## 6. Ranger le PC (facultatif)

- **Paramètres › Applications** : désinstaller « Python 3.13.17 » et
  « Python Launcher ».
- Supprimer `C:\STAT_VEO`.

## En cas d'erreur

Le script s'arrête sur un message `[ERREUR]` et laisse la fenêtre ouverte.
Prendre l'écran en photo.

| Message | Que faire |
|---|---|
| Python est introuvable | La case « Add python.exe to PATH » n'était pas cochée. Relancer l'installeur, choisir **Modify**, puis cocher la case. |
| Ce Python n'est pas en 64 bits Intel/AMD | PC ARM, ou un ancien Python 32 bits déjà présent. |
| L'installation des dependances a echoue | Le dossier `wheels` manque dans `VeoVideoControl\Engine`. |
