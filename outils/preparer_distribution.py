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

Le classeur ne doit plus contenir de donnees appartenant a un club :
le script refuse de preparer le paquet tant qu'il y trouve un
effectif, une composition, des equipes de poule ou un journal. La
macro ReinitialiserPourDistribution, dans le classeur, fait ce
menage ; elle s'execute sur une copie.

Usage :
    python3 outils/preparer_distribution.py
    python3 outils/preparer_distribution.py --zip
    python3 outils/preparer_distribution.py --classeur "copie.xlsm"
    python3 outils/preparer_distribution.py --controler --classeur "copie.xlsm"
    python3 outils/preparer_distribution.py --destination /tmp/paquet

Le dossier de destination est efface puis reconstruit a chaque appel.
"""

import argparse
import os
import re
import shutil
import sys
import zipfile


CLASSEUR = "Createur de match.xlsm"

# Tableaux qui ne doivent plus rien contenir dans un classeur
# destine a un autre club. Une colonne precise peut etre visee :
# la composition garde ses postes, mais pas ses joueurs.
TABLEAUX_A_CONTROLER = (
    ("LstEffectif", None, "l'effectif"),
    ("LstEquipesPoule", None, "les equipes de la poule"),
    ("JournalActions", None, "le journal d'actions"),
    ("COMPO", "Nom du joueur", "la composition"),
)

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


# =========================================================
# Controle des donnees restantes dans le classeur.
#
# Le classeur est lu tel quel, sans Excel ni bibliotheque
# tierce : un .xlsm est une archive zip de fichiers XML.
# =========================================================


def colonne_en_index(lettres):
    """Convertit "AA" en 27."""

    index = 0

    for caractere in lettres:
        index = index * 26 + (ord(caractere) - ord("A") + 1)

    return index


def decouper_reference(reference):
    """Renvoie (colonne1, ligne1, colonne2, ligne2) pour "A1:E68"."""

    debut, _, fin = reference.partition(":")

    if not fin:
        fin = debut

    motif = re.compile(r"([A-Z]+)(\d+)")

    colonne1, ligne1 = motif.match(debut).groups()
    colonne2, ligne2 = motif.match(fin).groups()

    return (
        colonne_en_index(colonne1),
        int(ligne1),
        colonne_en_index(colonne2),
        int(ligne2),
    )


def feuille_de_chaque_tableau(archive):
    """Renvoie {nom du tableau: (chemin de la feuille, ref, colonnes)}."""

    tableaux = {}

    chemins_tables = {}

    for nom in archive.namelist():

        if not nom.startswith("xl/tables/table"):
            continue

        contenu = archive.read(nom).decode("utf-8")

        chemins_tables[nom] = (
            re.search(r'name="([^"]+)"', contenu).group(1),
            re.search(r'ref="([^"]+)"', contenu).group(1),
            re.findall(r'<tableColumn[^>]*name="([^"]+)"', contenu),
        )

    for nom in archive.namelist():

        if not re.match(r"xl/worksheets/_rels/sheet\d+\.xml\.rels$", nom):
            continue

        feuille = nom.replace("_rels/", "").replace(".rels", "")

        rels = archive.read(nom).decode("utf-8")

        for cible in re.findall(r'Target="([^"]+)"', rels):

            chemin = "xl/" + cible.replace("../", "")

            if chemin in chemins_tables:

                nom_tableau, ref, colonnes = chemins_tables[chemin]

                tableaux[nom_tableau] = (feuille, ref, colonnes)

    return tableaux


def compter_cellules_remplies(archive, feuille, colonne1, colonne2, ligne1, ligne2):
    """Compte les cellules porteuses d'une valeur dans la zone."""

    contenu = archive.read(feuille).decode("utf-8")

    total = 0

    for cellule in re.finditer(r"<c\s[^>]*r=\"([A-Z]+)(\d+)\"[^>]*?(/>|>(.*?)</c>)", contenu, re.S):

        colonne = colonne_en_index(cellule.group(1))

        ligne = int(cellule.group(2))

        if not (colonne1 <= colonne <= colonne2):
            continue

        if not (ligne1 <= ligne <= ligne2):
            continue

        interieur = cellule.group(4) or ""

        if "<v>" in interieur or "<is>" in interieur:
            total += 1

    return total


def donnees_restantes(chemin_classeur):
    """Renvoie la liste des donnees de club encore presentes."""

    restes = []

    with zipfile.ZipFile(chemin_classeur) as archive:

        tableaux = feuille_de_chaque_tableau(archive)

        for nom_tableau, colonne_visee, description in TABLEAUX_A_CONTROLER:

            if nom_tableau not in tableaux:
                continue

            feuille, ref, colonnes = tableaux[nom_tableau]

            colonne1, ligne1, colonne2, ligne2 = decouper_reference(ref)

            # La premiere ligne porte les en-tetes.
            ligne1 += 1

            if ligne1 > ligne2:
                continue

            if colonne_visee:

                if colonne_visee not in colonnes:
                    continue

                colonne1 += colonnes.index(colonne_visee)
                colonne2 = colonne1

            remplies = compter_cellules_remplies(
                archive,
                feuille,
                colonne1,
                colonne2,
                ligne1,
                ligne2,
            )

            if remplies:
                restes.append(f"{description} ({remplies} cellules)")

    return restes


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


def copier(racine, destination, chemin_classeur):
    """Recopie la liste blanche et renvoie ce qui a ete copie."""

    copies = []

    for element in ELEMENTS:

        if element == CLASSEUR:
            source = chemin_classeur
        else:
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


def preparer(
    destination,
    avec_zip=False,
    classeur=None,
    forcer=False,
    controler_seulement=False,
):

    racine = racine_projet()

    chemin_classeur = classeur or os.path.join(racine, CLASSEUR)

    if not os.path.isfile(chemin_classeur):
        raise ErreurPreparation(f"Classeur introuvable : {chemin_classeur}")

    restes = donnees_restantes(chemin_classeur)

    if restes:

        print(
            f"{os.path.basename(chemin_classeur)} contient encore "
            "des donnees de club :"
        )

        for reste in restes:
            print(f"  - {reste}")

        if not forcer:

            raise ErreurPreparation(
                "Ouvrir une copie du classeur, y lancer la macro "
                "ReinitialiserPourDistribution, l'enregistrer, puis "
                "relancer avec --classeur sur cette copie.\n"
                "Pour passer outre : --forcer."
            )

        print("  (--forcer : le paquet est prepare quand meme)\n")

    elif controler_seulement:

        print(
            f"{os.path.basename(chemin_classeur)} ne contient plus "
            "de donnees de club."
        )

    if controler_seulement:
        return 0

    destination = os.path.abspath(destination)

    if os.path.exists(destination):
        shutil.rmtree(destination)

    os.makedirs(destination)

    copies = copier(racine, destination, chemin_classeur)

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
        "--classeur",
        help=(
            "classeur a placer dans le paquet "
            "(defaut : celui du projet)"
        ),
    )

    analyseur.add_argument(
        "--controler",
        action="store_true",
        dest="controler_seulement",
        help=(
            "verifie seulement que le classeur est vide de donnees, "
            "sans fabriquer le paquet"
        ),
    )

    analyseur.add_argument(
        "--forcer",
        action="store_true",
        help="prepare le paquet malgre des donnees de club restantes",
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

        return preparer(
            options.destination,
            avec_zip=options.avec_zip,
            classeur=options.classeur,
            forcer=options.forcer,
            controler_seulement=options.controler_seulement,
        )

    except ErreurPreparation as erreur:

        print(f"Erreur : {erreur}", file=sys.stderr)

        return 2


if __name__ == "__main__":
    raise SystemExit(main())
