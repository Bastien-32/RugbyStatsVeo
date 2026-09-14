#!/usr/bin/env python3
"""Lit ou ecrit le numero de version dans un classeur .xlsm.

La version est stockee dans la plage nommee APP_VERSION (creee au besoin
sur la feuille Parametres, cellule T2). Le classeur est modifie
chirurgicalement : seules les parties XML concernees sont reecrites, tout
le reste (macros VBA, graphiques, tableaux, mises en forme) est recopie
a l'octet pres.

Usage :
    python3 outils/version_xlsm.py lire  "Createur de match.xlsm"
    python3 outils/version_xlsm.py ecrire "Createur de match.xlsm" 1.0.2
"""

import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

FEUILLE_DEFAUT = "Paramètres"
CELLULE_DEFAUT = "T2"
NOM_PLAGE = "APP_VERSION"


class ErreurClasseur(Exception):
    pass


def _colonne_en_index(colonne):
    """Convertit "T" en 20, "AA" en 27."""
    index = 0
    for caractere in colonne:
        index = index * 26 + (ord(caractere) - ord("A") + 1)
    return index


def _decouper_reference(reference):
    """Decoupe "T2" en ("T", 2)."""
    correspondance = re.fullmatch(r"([A-Z]+)(\d+)", reference)
    if not correspondance:
        raise ErreurClasseur(f"Reference de cellule invalide : {reference}")
    return correspondance.group(1), int(correspondance.group(2))


def _cible(workbook_xml):
    """Retourne (feuille, cellule) d'apres la plage nommee, ou les valeurs par defaut."""
    motif = rf'<definedName name="{NOM_PLAGE}"[^>]*>([^<]+)</definedName>'
    correspondance = re.search(motif, workbook_xml)
    if not correspondance:
        return FEUILLE_DEFAUT, CELLULE_DEFAUT

    refers_to = correspondance.group(1)
    correspondance = re.fullmatch(r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)", refers_to)
    if not correspondance:
        raise ErreurClasseur(
            f"La plage {NOM_PLAGE} pointe sur '{refers_to}', que je ne sais pas "
            "interpreter. Elle doit designer une cellule unique."
        )
    return correspondance.group(1), f"{correspondance.group(2)}{correspondance.group(3)}"


def _chemin_feuille(archive, nom_feuille):
    """Retrouve le fichier XML correspondant a une feuille, via son r:id."""
    workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")

    motif = rf'<sheet name="{re.escape(nom_feuille)}"[^>]*r:id="([^"]+)"'
    correspondance = re.search(motif, workbook_xml)
    if not correspondance:
        raise ErreurClasseur(f"Feuille introuvable dans le classeur : {nom_feuille}")
    rid = correspondance.group(1)

    rels_xml = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    correspondance = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels_xml)
    if not correspondance:
        raise ErreurClasseur(f"Relation {rid} introuvable pour la feuille {nom_feuille}")

    return "xl/" + correspondance.group(1).lstrip("/")


def _valeur_cellule(archive, chemin_feuille, cellule):
    """Lit la valeur texte d'une cellule, qu'elle soit inline ou en table partagee."""
    feuille_xml = archive.read(chemin_feuille).decode("utf-8")

    correspondance = re.search(rf'<c r="{cellule}"[^>]*?(?:/>|>(.*?)</c>)', feuille_xml, re.S)
    if not correspondance or not correspondance.group(1):
        return None

    balise = correspondance.group(0)
    contenu = correspondance.group(1)

    if 't="inlineStr"' in balise:
        valeur = re.search(r"<t[^>]*>(.*?)</t>", contenu, re.S)
        return valeur.group(1) if valeur else None

    valeur = re.search(r"<v>(.*?)</v>", contenu, re.S)
    if not valeur:
        return None

    if 't="s"' in balise:
        strings_xml = archive.read("xl/sharedStrings.xml").decode("utf-8")
        entrees = re.findall(r"<si>(.*?)</si>", strings_xml, re.S)
        index = int(valeur.group(1))
        if index >= len(entrees):
            return None
        return "".join(re.findall(r"<t[^>]*>(.*?)</t>", entrees[index], re.S))

    return valeur.group(1)


def lire(chemin_classeur):
    with zipfile.ZipFile(chemin_classeur) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        nom_feuille, cellule = _cible(workbook_xml)
        return _valeur_cellule(archive, _chemin_feuille(archive, nom_feuille), cellule)


def _injecter_cellule(feuille_xml, cellule, version):
    """Insere ou remplace la cellule, en respectant l'ordre des colonnes."""
    colonne, ligne = _decouper_reference(cellule)
    nouvelle = f'<c r="{cellule}" t="inlineStr"><is><t>{escape(version)}</t></is></c>'

    existante = re.search(rf'<c r="{cellule}"[^>]*?(?:/>|>.*?</c>)', feuille_xml, re.S)
    if existante:
        return feuille_xml[: existante.start()] + nouvelle + feuille_xml[existante.end() :]

    bloc_ligne = re.search(rf'<row r="{ligne}"[^>]*?(?:/>|>.*?</row>)', feuille_xml, re.S)
    if not bloc_ligne:
        raise ErreurClasseur(
            f"La ligne {ligne} n'existe pas dans la feuille. Choisis une cellule "
            "situee sur une ligne deja utilisee."
        )

    texte_ligne = bloc_ligne.group(0)
    if texte_ligne.endswith("/>"):
        raise ErreurClasseur(f"La ligne {ligne} est vide, choisis une autre cellule.")

    # Insere la cellule a sa place, les cellules d'une ligne devant rester triees.
    position = len(texte_ligne) - len("</row>")
    for suivante in re.finditer(r'<c r="([A-Z]+)\d+"', texte_ligne):
        if _colonne_en_index(suivante.group(1)) > _colonne_en_index(colonne):
            position = suivante.start()
            break
    ligne_modifiee = texte_ligne[:position] + nouvelle + texte_ligne[position:]

    # Elargit l'etendue declaree de la ligne si besoin.
    spans = re.search(r'spans="(\d+):(\d+)"', ligne_modifiee)
    if spans:
        debut, fin = int(spans.group(1)), int(spans.group(2))
        index = _colonne_en_index(colonne)
        ligne_modifiee = ligne_modifiee.replace(
            spans.group(0), f'spans="{min(debut, index)}:{max(fin, index)}"', 1
        )

    return feuille_xml[: bloc_ligne.start()] + ligne_modifiee + feuille_xml[bloc_ligne.end() :]


def _elargir_dimension(feuille_xml, cellule):
    """Etend <dimension> pour englober la cellule ajoutee."""
    dimension = re.search(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/>', feuille_xml)
    if not dimension:
        return feuille_xml

    colonne, ligne = _decouper_reference(cellule)
    col_fin = dimension.group(3)
    if _colonne_en_index(colonne) > _colonne_en_index(col_fin):
        col_fin = colonne
    lig_fin = max(int(dimension.group(4)), ligne)

    remplacement = (
        f'<dimension ref="{dimension.group(1)}{dimension.group(2)}:{col_fin}{lig_fin}"/>'
    )
    return feuille_xml.replace(dimension.group(0), remplacement, 1)


def _declarer_plage(workbook_xml, nom_feuille, cellule):
    """Ajoute la plage nommee si elle n'existe pas, a sa place alphabetique."""
    if f'<definedName name="{NOM_PLAGE}"' in workbook_xml:
        return workbook_xml

    colonne, ligne = _decouper_reference(cellule)
    feuille = nom_feuille if re.fullmatch(r"[\w\.]+", nom_feuille) else f"'{nom_feuille}'"
    entree = f'<definedName name="{NOM_PLAGE}">{feuille}!${colonne}${ligne}</definedName>'

    if "<definedNames>" not in workbook_xml:
        raise ErreurClasseur("Le classeur ne contient aucun bloc <definedNames>.")

    for suivante in re.finditer(r'<definedName name="([^"]+)"', workbook_xml):
        if suivante.group(1) > NOM_PLAGE:
            return workbook_xml[: suivante.start()] + entree + workbook_xml[suivante.start() :]

    return workbook_xml.replace("</definedNames>", entree + "</definedNames>", 1)


def ecrire(chemin_classeur, version):
    with zipfile.ZipFile(chemin_classeur) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        nom_feuille, cellule = _cible(workbook_xml)
        chemin_feuille = _chemin_feuille(archive, nom_feuille)

        feuille_xml = archive.read(chemin_feuille).decode("utf-8")
        feuille_xml = _injecter_cellule(feuille_xml, cellule, version)
        feuille_xml = _elargir_dimension(feuille_xml, cellule)
        workbook_xml = _declarer_plage(workbook_xml, nom_feuille, cellule)

        remplacements = {
            chemin_feuille: feuille_xml.encode("utf-8"),
            "xl/workbook.xml": workbook_xml.encode("utf-8"),
        }

        # Reecrit l'archive entiere pour ne pas laisser d'entree orpheline,
        # en recopiant tel quel tout ce qui n'est pas concerne.
        descripteur, temporaire = tempfile.mkstemp(
            suffix=".xlsm", dir=os.path.dirname(os.path.abspath(chemin_classeur))
        )
        os.close(descripteur)
        try:
            with zipfile.ZipFile(temporaire, "w", zipfile.ZIP_DEFLATED) as sortie:
                for membre in archive.infolist():
                    donnees = remplacements.get(membre.filename)
                    if donnees is None:
                        donnees = archive.read(membre.filename)
                    info = zipfile.ZipInfo(membre.filename, date_time=membre.date_time)
                    info.compress_type = membre.compress_type
                    info.external_attr = membre.external_attr
                    sortie.writestr(info, donnees)
        except Exception:
            os.unlink(temporaire)
            raise

    shutil.move(temporaire, chemin_classeur)
    return nom_feuille, cellule


def main(arguments):
    if len(arguments) < 2:
        print(__doc__.strip(), file=sys.stderr)
        return 2

    action, chemin_classeur = arguments[0], arguments[1]

    if not os.path.exists(chemin_classeur):
        print(f"Classeur introuvable : {chemin_classeur}", file=sys.stderr)
        return 1

    verrou = os.path.join(
        os.path.dirname(os.path.abspath(chemin_classeur)),
        "~$" + os.path.basename(chemin_classeur),
    )
    if action == "ecrire" and os.path.exists(verrou):
        print(
            f"'{os.path.basename(chemin_classeur)}' est ouvert dans Excel. "
            "Ferme-le avant de relancer.",
            file=sys.stderr,
        )
        return 1

    try:
        if action == "lire":
            version = lire(chemin_classeur)
            if version is None:
                print("(aucune version enregistree)", file=sys.stderr)
                return 1
            print(version)
            return 0

        if action == "ecrire":
            if len(arguments) < 3:
                print("Version manquante.", file=sys.stderr)
                return 2
            version = arguments[2].lstrip("v")
            nom_feuille, cellule = ecrire(chemin_classeur, version)
            print(f"Version {version} ecrite dans {nom_feuille}!{cellule}")
            return 0
    except (ErreurClasseur, zipfile.BadZipFile) as erreur:
        print(f"Erreur : {erreur}", file=sys.stderr)
        return 1

    print(f"Action inconnue : {action}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
