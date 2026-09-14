#!/usr/bin/env python3
"""Exporte les modules VBA d'un classeur .xlsm en fichiers texte.

Le code VBA vit dans un flux binaire (vbaProject.bin) : git ne sait
donc pas comparer deux versions d'un classeur, et l'historique ne
montre qu'un "Bin 380873 -> 418206 bytes". Cet outil ecrit chaque
module dans un fichier texte a part, ce qui rend les evolutions
lisibles dans l'historique et permet de reporter une correction d'un
classeur a l'autre par un simple import de module.

L'operation est a sens unique : Excel pour Mac n'autorise pas
l'ecriture du projet VBA depuis l'exterieur. Pour reinjecter un
module dans un classeur, passer par l'editeur VBA,
Fichier > Importer un fichier.

Le classeur n'est jamais modifie : il est ouvert en lecture seule.

Usage :
    python3 outils/export_vba.py "Createur de match.xlsm"
    python3 outils/export_vba.py "Createur de match.xlsm" --dossier vba
    python3 outils/export_vba.py "Createur de match.xlsm" --verifier

Avec --verifier, rien n'est ecrit : la commande signale les ecarts
entre le classeur et l'export, et rend un code de sortie non nul.
Pratique pour controler avant un commit que l'export est a jour.

Necessite oletools :
    python3 -m pip install oletools
"""

import argparse
import os
import sys


EXTENSIONS_MODULES = (".bas", ".cls", ".frm")

DOSSIER_DEFAUT = "vba"


class ErreurExport(Exception):
    pass


def charger_extracteur():
    """Importe oletools en expliquant comment l'installer au besoin."""

    try:

        from oletools.olevba import VBA_Parser

    except ImportError:

        raise ErreurExport(
            "oletools est necessaire pour lire le code VBA.\n"
            "Installation : python3 -m pip install oletools"
        )

    return VBA_Parser


def normaliser(code):
    """Uniformise les fins de ligne pour que les diffs restent lisibles."""

    code = code.replace("\r\n", "\n").replace("\r", "\n")

    if code and not code.endswith("\n"):
        code += "\n"

    return code


def lire_modules(chemin_classeur):
    """Renvoie {nom de fichier du module: code source}."""

    if not os.path.isfile(chemin_classeur):
        raise ErreurExport(f"Classeur introuvable : {chemin_classeur}")

    VBA_Parser = charger_extracteur()

    parseur = VBA_Parser(chemin_classeur)

    modules = {}

    try:

        if not parseur.detect_vba_macros():
            raise ErreurExport(
                f"Aucun code VBA dans {os.path.basename(chemin_classeur)}"
            )

        for _, _, nom_module, code in parseur.extract_macros():

            if not nom_module:
                continue

            modules[nom_module] = normaliser(code)

    finally:

        parseur.close()

    return modules


def lire_export_existant(dossier):
    """Renvoie les modules deja presents dans le dossier d'export."""

    existants = {}

    if not os.path.isdir(dossier):
        return existants

    for nom in sorted(os.listdir(dossier)):

        if not nom.endswith(EXTENSIONS_MODULES):
            continue

        chemin = os.path.join(dossier, nom)

        with open(chemin, encoding="utf-8") as fichier:
            existants[nom] = fichier.read()

    return existants


def comparer(modules, existants):
    """Classe les ecarts entre le classeur et l'export."""

    ajoutes = sorted(set(modules) - set(existants))

    supprimes = sorted(set(existants) - set(modules))

    modifies = sorted(
        nom
        for nom in set(modules) & set(existants)
        if modules[nom] != existants[nom]
    )

    return ajoutes, supprimes, modifies


def decrire(ajoutes, supprimes, modifies):

    for nom in ajoutes:
        print(f"  + {nom}")

    for nom in modifies:
        print(f"  ~ {nom}")

    for nom in supprimes:
        print(f"  - {nom}")


def exporter(chemin_classeur, dossier, verifier=False):
    """Ecrit les modules, ou signale les ecarts si verifier est vrai."""

    modules = lire_modules(chemin_classeur)

    existants = lire_export_existant(dossier)

    ajoutes, supprimes, modifies = comparer(modules, existants)

    if not (ajoutes or supprimes or modifies):

        print(
            f"{len(modules)} modules : l'export de {dossier} "
            "est deja a jour."
        )

        return 0

    if verifier:

        print(
            f"L'export de {dossier} ne correspond plus au classeur :"
        )

        decrire(ajoutes, supprimes, modifies)

        print("\nRelancer sans --verifier pour le mettre a jour.")

        return 1

    os.makedirs(dossier, exist_ok=True)

    for nom in ajoutes + modifies:

        chemin = os.path.join(dossier, nom)

        with open(chemin, "w", encoding="utf-8") as fichier:
            fichier.write(modules[nom])

    for nom in supprimes:
        os.remove(os.path.join(dossier, nom))

    print(f"{len(modules)} modules exportes dans {dossier} :")

    decrire(ajoutes, supprimes, modifies)

    return 0


def analyser_arguments(arguments):

    analyseur = argparse.ArgumentParser(
        description=(
            "Exporte les modules VBA d'un classeur .xlsm "
            "en fichiers texte versionnables."
        )
    )

    analyseur.add_argument(
        "classeur",
        help="chemin du classeur .xlsm a lire",
    )

    analyseur.add_argument(
        "--dossier",
        default=DOSSIER_DEFAUT,
        help=f"dossier d'export (defaut : {DOSSIER_DEFAUT})",
    )

    analyseur.add_argument(
        "--verifier",
        action="store_true",
        help="n'ecrit rien, signale seulement les ecarts",
    )

    return analyseur.parse_args(arguments)


def main(arguments=None):

    options = analyser_arguments(arguments)

    try:

        return exporter(
            options.classeur,
            options.dossier,
            verifier=options.verifier,
        )

    except ErreurExport as erreur:

        print(f"Erreur : {erreur}", file=sys.stderr)

        return 2


if __name__ == "__main__":
    raise SystemExit(main())
