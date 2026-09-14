#!/usr/bin/env python3
"""Prepare le dossier a remettre aux utilisateurs de l'outil.

Le depot contient de quoi developper : sources du moteur video,
export du code VBA, scripts, historique. Rien de tout cela ne sert a
quelqu'un qui veut seulement saisir un match. Ce script recopie donc
la liste des elements utiles dans un dossier a part, pret a etre
transmis ou compresse.

Ce qui est volontairement laisse de cote :

    Engine/          sources Python du moteur, deja compilees
    Documentation/   protocole interne Excel <-> moteur
    old/             anciennes versions
    outils/, vba/    outillage de developpement
    office.txt       lien vers une copie piratee d'Office
    .vscode, .DS_Store, .venv, .git

Usage :
    python3 outils/preparer_distribution.py
    python3 outils/preparer_distribution.py --zip
    python3 outils/preparer_distribution.py --destination /tmp/paquet

Le dossier de destination est efface puis reconstruit a chaque appel.
"""

import argparse
import os
import shutil
import sys
import zipfile


# Elements recopies tels quels, dans l'ordre d'affichage.
ELEMENTS = (
    "Createur de match.xlsm",
    "INSTALLATION.md",
    "INSTALLATION.txt",
    "VeoVideoControl/VeoVideoControl.app",
    "VeoVideoControl/Arreter VeoVideoControl.app",
    "VeoVideoControl/VeoVideoControlEngine",
    "VeoVideoControl/Extension",
)

# Fichiers repris du dossier des installeurs, extension par extension.
INSTALLEURS = ("logiciels à installer", (".dmg", ".exe"))

# Dossier vide a creer : le classeur y depose les fichiers de match.
DOSSIERS_VIDES = ("Matchs",)

PARASITES = (".DS_Store", ".vscode", "__pycache__")

DESTINATION_DEFAUT = os.path.join("dist", "Stats Rugby")


class ErreurPreparation(Exception):
    pass


def racine_projet():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def nettoyer(dossier):
    """Supprime les fichiers parasites recopies avec les dossiers."""

    for chemin, sous_dossiers, fichiers in os.walk(dossier, topdown=True):

        for nom in list(sous_dossiers):

            if nom in PARASITES:
                shutil.rmtree(os.path.join(chemin, nom), ignore_errors=True)
                sous_dossiers.remove(nom)

        for nom in fichiers:

            if nom in PARASITES:
                os.remove(os.path.join(chemin, nom))


def copier(racine, destination):
    """Recopie la liste blanche et renvoie ce qui a ete copie."""

    copies = []

    for element in ELEMENTS:

        source = os.path.join(racine, element)

        if not os.path.exists(source):
            raise ErreurPreparation(f"Element introuvable : {element}")

        cible = os.path.join(destination, element)

        os.makedirs(os.path.dirname(cible), exist_ok=True)

        if os.path.isdir(source):
            shutil.copytree(source, cible, symlinks=True)
        else:
            shutil.copy2(source, cible)

        copies.append(element)

    dossier_installeurs, extensions = INSTALLEURS

    source_installeurs = os.path.join(racine, dossier_installeurs)

    if os.path.isdir(source_installeurs):

        cible_installeurs = os.path.join(destination, dossier_installeurs)

        os.makedirs(cible_installeurs, exist_ok=True)

        for nom in sorted(os.listdir(source_installeurs)):

            if nom.lower().endswith(extensions):

                shutil.copy2(
                    os.path.join(source_installeurs, nom),
                    os.path.join(cible_installeurs, nom),
                )

                copies.append(os.path.join(dossier_installeurs, nom))

    for dossier in DOSSIERS_VIDES:

        os.makedirs(os.path.join(destination, dossier), exist_ok=True)

        copies.append(f"{dossier}/ (vide)")

    return copies


def taille(dossier):
    """Taille totale, en octets."""

    total = 0

    for chemin, _, fichiers in os.walk(dossier):

        for nom in fichiers:

            fichier = os.path.join(chemin, nom)

            if not os.path.islink(fichier):
                total += os.path.getsize(fichier)

    return total


def formater(octets):

    for unite in ("o", "Ko", "Mo", "Go"):

        if octets < 1024:
            return f"{octets:.0f} {unite}"

        octets /= 1024

    return f"{octets:.0f} To"


def compresser(destination):
    """Cree une archive zip a cote du dossier prepare."""

    archive = destination.rstrip(os.sep) + ".zip"

    if os.path.exists(archive):
        os.remove(archive)

    base = os.path.basename(destination.rstrip(os.sep))

    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zip_sortie:

        for chemin, _, fichiers in os.walk(destination):

            for nom in fichiers:

                fichier = os.path.join(chemin, nom)

                interne = os.path.join(
                    base,
                    os.path.relpath(fichier, destination),
                )

                zip_sortie.write(fichier, interne)

    return archive


def preparer(destination, avec_zip=False):

    racine = racine_projet()

    destination = os.path.abspath(destination)

    if os.path.exists(destination):
        shutil.rmtree(destination)

    os.makedirs(destination)

    copies = copier(racine, destination)

    nettoyer(destination)

    print(f"Paquet prepare dans {destination} :")

    for element in copies:
        print(f"  {element}")

    print(f"\nTaille : {formater(taille(destination))}")

    if avec_zip:

        archive = compresser(destination)

        print(
            f"Archive : {archive} "
            f"({formater(os.path.getsize(archive))})"
        )

    return 0


def analyser_arguments(arguments):

    analyseur = argparse.ArgumentParser(
        description=(
            "Prepare le dossier a remettre aux utilisateurs, "
            "sans les fichiers de developpement."
        )
    )

    analyseur.add_argument(
        "--destination",
        default=DESTINATION_DEFAUT,
        help=f"dossier de sortie (defaut : {DESTINATION_DEFAUT})",
    )

    analyseur.add_argument(
        "--zip",
        action="store_true",
        dest="avec_zip",
        help="cree aussi une archive zip du paquet",
    )

    return analyseur.parse_args(arguments)


def main(arguments=None):

    options = analyser_arguments(arguments)

    try:

        return preparer(options.destination, avec_zip=options.avec_zip)

    except ErreurPreparation as erreur:

        print(f"Erreur : {erreur}", file=sys.stderr)

        return 2


if __name__ == "__main__":
    raise SystemExit(main())
