Attribute VB_Name = "modRecapConstruction"
Option Explicit

' =========================================================
' RECAPITULATIF DE SAISON - FABRICATION DU CLASSEUR
'
' A executer une seule fois, dans un classeur neuf
' enregistre sous Matchs/Recapitulatif saison.xlsm.
'
' La macro fabrique les sept feuilles, leurs tableaux,
' leurs listes deroulantes et leurs boutons. Elle est
' idempotente : relancee, elle remet en place ce qui
' manque sans toucher aux matchs deja ajoutes.
' =========================================================

Private Const NB_LIGNES_DEPART As Long = 1

' Hauteur du tableau qui recoit la liste de statistiques
' livree avec le classeur, avant son ecriture dans la
' feuille Parametres.
Private Const MAX_STATS_PAR_DEFAUT As Long = 60


Public Sub ConstruireClasseurRecapitulatif()

    Dim EtatAffichage As Boolean

    EtatAffichage = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    ' Parametres en premier : c'est elle qui porte les
    ' plages nommees auxquelles les listes deroulantes des
    ' criteres se referent.
    ConstruireFeuilleParametres

    ConstruireFeuilleJournal
    ConstruireFeuilleCompositions
    ConstruireFeuilleMatchs
    ConstruireFeuilleFiltres
    ConstruireFeuilleClassement
    ConstruireFeuilleTemps

    RangerFeuilles

    Application.ScreenUpdating = EtatAffichage

    FeuilleRecap(FEUILLE_JOURNAL).Activate

    MsgBox _
        "Classeur r" & ChrW(233) & "capitulatif pr" & ChrW(234) & "t." & _
        vbCrLf & vbCrLf & _
        "Enregistre-le dans le dossier Matchs, puis " & _
        "clique sur " & ChrW(171) & " Ajouter les nouveaux " & _
        "matchs " & ChrW(187) & ".", _
        vbInformation, _
        "R" & ChrW(233) & "capitulatif de saison"

    Exit Sub

GestionErreur:

    Application.ScreenUpdating = EtatAffichage

    MsgBox _
        "La construction a " & ChrW(233) & "chou" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Erreur"

End Sub


' =========================================================
' FEUILLE JOURNAL ACTIONS
'
' La base cumulative : une ligne par action, tous matchs
' de championnat confondus. C'est elle qui rend inutile la
' relecture des fichiers de match.
'
' La colonne Mi-temps y figure mais ne sert pas au
' classement, qui totalise le match. Elle reste la pour qui
' voudrait monter un tableau croise dessus.
' =========================================================
Private Sub ConstruireFeuilleJournal()

    Dim ws As Worksheet

    Set ws = FeuilleRecapOuCreee(FEUILLE_JOURNAL)

    EcrireTitre _
        ws, _
        "Journal d'actions de la saison", _
        "Base cumulative. Le bouton n'ouvre que les " & _
        "fichiers de championnat pos" & ChrW(233) & "s " & ChrW(224) & _
        " la racine du dossier et pas encore int" & _
        ChrW(233) & "gr" & ChrW(233) & "s."

    CreerTableau _
        ws, _
        TBL_JOURNAL, _
        Array( _
            "ID match", "Fichier", "Saison", "Date", _
            "Cat" & ChrW(233) & "gorie", "Adversaire", "Lieu", _
            "Phase", "Journ" & ChrW(233) & "e", "R" & ChrW(233) & "sultat", _
            "Mi-temps", "Joueur", "Groupe fautif", _
            "Action", "Motif", "Possession", "Temps vid" & ChrW(233) & "o" _
        )

    PoserBouton _
        ws, "Ajouter les nouveaux matchs", _
        "AjouterNouveauxMatchs", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 210

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 34
    ws.Columns(PREMIERE_COLONNE + 1).ColumnWidth = 34
    ws.Columns(PREMIERE_COLONNE + 2).ColumnWidth = 9
    ws.Columns(PREMIERE_COLONNE + 3).ColumnWidth = 11
    ws.Columns(PREMIERE_COLONNE + 4).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 5).ColumnWidth = 24
    ws.Columns(PREMIERE_COLONNE + 6).Resize(, 5).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 11).ColumnWidth = 26
    ws.Columns(PREMIERE_COLONNE + 12).Resize(, 5).ColumnWidth = 18

    ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + 3) _
        .EntireColumn.NumberFormat = "dd/mm/yyyy"

    FigerVolets ws, LIGNE_ENTETE, PREMIERE_COLONNE + 1

End Sub


' =========================================================
' FEUILLE COMPOSITIONS
'
' L'autre moitie de la base : une ligne par joueur et par
' match, avec le temps de jeu saisi dans le fichier de
' match.
'
' La colonne Minutes peut etre corrigee ici a la main quand
' un temps de jeu a ete oublie : cela evite de rouvrir le
' fichier de match. La correction tient tant que le match
' n'est pas reimporte.
' =========================================================
Private Sub ConstruireFeuilleCompositions()

    Dim ws As Worksheet

    Set ws = FeuilleRecapOuCreee(FEUILLE_COMPOSITIONS)

    EcrireTitre _
        ws, _
        "Compositions et temps de jeu", _
        "Minutes lues dans la colonne " & ChrW(171) & _
        " temps de jeu " & ChrW(187) & " de la feuille Compo de " & _
        "chaque match. Elles peuvent " & ChrW(234) & "tre " & _
        "corrig" & ChrW(233) & "es directement ici."

    CreerTableau _
        ws, _
        TBL_COMPOSITIONS, _
        Array( _
            "ID match", "Fichier", "Saison", "Date", _
            "Cat" & ChrW(233) & "gorie", "Adversaire", "Lieu", _
            "Phase", "Journ" & ChrW(233) & "e", "R" & ChrW(233) & "sultat", _
            "Joueur", "Poste", "Ligne", "Minutes" _
        )

    PoserBouton _
        ws, "Recalculer", "AppliquerFiltres", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 130

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 34
    ws.Columns(PREMIERE_COLONNE + 1).ColumnWidth = 34
    ws.Columns(PREMIERE_COLONNE + 2).ColumnWidth = 9
    ws.Columns(PREMIERE_COLONNE + 3).ColumnWidth = 11
    ws.Columns(PREMIERE_COLONNE + 4).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 5).ColumnWidth = 24
    ws.Columns(PREMIERE_COLONNE + 6).Resize(, 4).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 10).ColumnWidth = 26
    ws.Columns(PREMIERE_COLONNE + 11).Resize(, 3).ColumnWidth = 12

    ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + 3) _
        .EntireColumn.NumberFormat = "dd/mm/yyyy"

    FigerVolets ws, LIGNE_ENTETE, PREMIERE_COLONNE + 1

End Sub


' =========================================================
' FEUILLE MATCHS
'
' L'inventaire de ce que contient la base : une ligne par
' match ajoute. C'est aussi la memoire qui dit quels
' fichiers ont deja ete lus.
'
' Une seule colonne se coche a la main :
'
'   Cibler : le match fait partie d'une selection
'            manuelle, utilisee quand le filtre
'            "Matchs cibles uniquement" est sur Oui.
'
' La colonne Retenu est ecrite par le calcul : elle montre
' quels matchs ont servi au dernier classement.
' =========================================================
Private Sub ConstruireFeuilleMatchs()

    Dim ws As Worksheet

    Set ws = FeuilleRecapOuCreee(FEUILLE_MATCHS)

    EcrireTitre _
        ws, _
        "Matchs de la base", _
        "Coche Cibler pour retenir un match dans une " & _
        "s" & ChrW(233) & "lection manuelle. Pour reprendre un " & _
        "match dont la saisie a chang" & ChrW(233) & ", clique sur " & _
        "sa ligne puis sur " & ChrW(171) & " R" & ChrW(233) & _
        "importer ce match " & ChrW(187) & "."

    CreerTableau _
        ws, _
        TBL_MATCHS, _
        Array( _
            "Cibler", "Retenu", "ID match", "Fichier", _
            "Saison", "Date", "Cat" & ChrW(233) & "gorie", _
            "Adversaire", "Lieu", "Phase", _
            "Journ" & ChrW(233) & "e", "R" & ChrW(233) & "sultat", _
            "Actions", "Joueurs", "Minutes saisies", _
            "Ajout" & ChrW(233) & " le", _
            "Fichier modifi" & ChrW(233) & " le" _
        )

    PoserBouton _
        ws, "Ajouter les nouveaux matchs", _
        "AjouterNouveauxMatchs", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 210

    PoserBouton _
        ws, "R" & ChrW(233) & "importer ce match", _
        "ReimporterMatch", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 150, 220

    PoserBouton _
        ws, "Retirer ce match", "RetirerMatch", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 140, 380

    PoserBouton _
        ws, "Tout cibler", "ToutCibler", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 100, 530

    PoserBouton _
        ws, "Ne rien cibler", "RienCibler", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 110, 640

    With ws.Columns(PREMIERE_COLONNE).Resize(, 2)
        .ColumnWidth = 9
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(PREMIERE_COLONNE + 2).ColumnWidth = 34
    ws.Columns(PREMIERE_COLONNE + 3).ColumnWidth = 40
    ws.Columns(PREMIERE_COLONNE + 4).ColumnWidth = 9
    ws.Columns(PREMIERE_COLONNE + 5).ColumnWidth = 11
    ws.Columns(PREMIERE_COLONNE + 6).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 7).ColumnWidth = 24
    ws.Columns(PREMIERE_COLONNE + 8).Resize(, 4).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 12).Resize(, 3).ColumnWidth = 12
    ws.Columns(PREMIERE_COLONNE + 15).Resize(, 2).ColumnWidth = 18

    PoserListeCoche ws, PREMIERE_COLONNE

    ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + MTC_DATE - 1) _
        .EntireColumn.NumberFormat = "dd/mm/yyyy"

    ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + MTC_AJOUTE - 1) _
        .Resize(1, 2).EntireColumn.NumberFormat = "dd/mm/yyyy hh:mm"

    FigerVolets ws, LIGNE_ENTETE, PREMIERE_COLONNE + 3

End Sub


' =========================================================
' FEUILLE FILTRES
'
' Un critere par ligne : le libelle en colonne B, la valeur
' choisie en colonne C. Chaque valeur porte un nom de
' cellule, seul point d'entree du module de calcul.
'
' Les listes deroulantes pointent sur des plages nommees
' reconstruites a chaque ajout, pour ne proposer que ce qui
' existe reellement dans la base.
' =========================================================
Private Sub ConstruireFeuilleFiltres()

    Dim ws As Worksheet

    Set ws = FeuilleRecapOuCreee(FEUILLE_FILTRES)

    EcrireTitre _
        ws, _
        "Filtres d'analyse", _
        "Choisis les crit" & ChrW(232) & "res, puis clique sur " & _
        ChrW(171) & " Appliquer les filtres " & ChrW(187) & ". " & _
        "Le classement et les temps de jeu se " & _
        "recalculent ensemble."

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 30
    ws.Columns(PREMIERE_COLONNE + 1).ColumnWidth = 26

    EcrireCritere ws, 6, "Saison", "FILTRE_SAISON", VALEUR_TOUS
    EcrireCritere ws, 7, "Cat" & ChrW(233) & "gorie", "FILTRE_CATEGORIE", VALEUR_TOUS
    EcrireCritere ws, 8, "Phase", "FILTRE_PHASE", VALEUR_TOUS
    EcrireCritere ws, 9, "Lieu", "FILTRE_LIEU", VALEUR_TOUS
    EcrireCritere ws, 10, "R" & ChrW(233) & "sultat", "FILTRE_RESULTAT", VALEUR_TOUS
    EcrireCritere ws, 11, "Adversaire", "FILTRE_ADVERSAIRE", VALEUR_TOUS

    EcrireCritere ws, 13, "Matchs jou" & ChrW(233) & "s " & ChrW(224) & " partir du", "FILTRE_DATE_DEBUT", ""
    EcrireCritere ws, 14, "Matchs jou" & ChrW(233) & "s jusqu'au", "FILTRE_DATE_FIN", ""
    EcrireCritere ws, 15, "N derniers matchs (0 = tous)", "FILTRE_DERNIERS", 0

    EcrireCritere ws, 17, "Matchs cibl" & ChrW(233) & "s uniquement", "FILTRE_CIBLES", "Non"
    EcrireCritere ws, 18, "Minutes jou" & ChrW(233) & "es minimum", "FILTRE_MINUTES_MIN", 0

    EcrireCritere ws, 20, "Mode d'affichage", "RECAP_MODE", MODE_TOTAUX

    ws.Range("C13:C14").NumberFormat = "dd/mm/yyyy"
    ws.Range("C15").NumberFormat = "0"
    ws.Range("C18").NumberFormat = "0"

    PoserListe ws.Range("C6"), "=LST_SAISONS"
    PoserListe ws.Range("C7"), "=LST_CATEGORIES"
    PoserListe ws.Range("C8"), "=LST_PHASES"
    PoserListe ws.Range("C9"), "=LST_LIEUX"
    PoserListe ws.Range("C10"), "=LST_RESULTATS"
    PoserListe ws.Range("C11"), "=LST_ADVERSAIRES"
    PoserListe ws.Range("C17"), "Oui,Non"

    PoserListe _
        ws.Range("C20"), _
        MODE_TOTAUX & "," & MODE_PAR_MATCH & "," & MODE_PAR_80

    ws.Range("B22").Value = _
        "Le mode d'affichage change la lecture du " & _
        "classement : totaux de la p" & ChrW(233) & "riode, " & _
        "moyenne par match jou" & ChrW(233) & ", ou ramen" & ChrW(233) & _
        " " & ChrW(224) & " 80 minutes de jeu. Le bouton de la " & _
        "feuille Classement fait la m" & ChrW(234) & "me chose."

    ws.Range("B22").Font.Italic = True

    PoserBouton _
        ws, "Appliquer les filtres", "AppliquerFiltres", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 170

    PoserBouton _
        ws, "R" & ChrW(233) & "initialiser", "ReinitialiserFiltres", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 110, 180

End Sub


' =========================================================
' FEUILLE CLASSEMENT
'
' Le tableau est entierement reecrit par le calcul : ses
' colonnes dependent des statistiques cochees dans
' Parametres et des categories rencontrees. La
' construction ne pose donc ici que le titre et les
' boutons.
' =========================================================
Private Sub ConstruireFeuilleClassement()

    Dim ws As Worksheet

    Set ws = FeuilleRecapOuCreee(FEUILLE_CLASSEMENT)

    EcrireTitre _
        ws, _
        "Statistiques par joueur", _
        "Totaux du match, mi-temps confondues. Clique " & _
        "sur l'en-t" & ChrW(234) & "te d'une colonne pour trier. " & _
        "Le bandeau ci-dessous rappelle les filtres " & _
        "appliqu" & ChrW(233) & "s."

    PoserBouton _
        ws, "Mode d'affichage", "BasculerModeAffichage", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 190

    PoserBouton _
        ws, "Recalculer", "AppliquerFiltres", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 110, 200

End Sub


' =========================================================
' FEUILLE TEMPS DE JEU
'
' Meme principe : une colonne de minutes par categorie
' rencontree, donc un tableau reecrit a chaque calcul.
' =========================================================
Private Sub ConstruireFeuilleTemps()

    Dim ws As Worksheet

    Set ws = FeuilleRecapOuCreee(FEUILLE_TEMPS)

    EcrireTitre _
        ws, _
        "Temps de jeu par cat" & ChrW(233) & "gorie", _
        "Matchs, minutes et moyenne, s" & ChrW(233) & "par" & _
        ChrW(233) & "s selon l'" & ChrW(233) & "quipe dans laquelle " & _
        "le joueur a jou" & ChrW(233) & "."

    PoserBouton _
        ws, "Recalculer", "AppliquerFiltres", _
        LIGNE_ENTETE - 2, PREMIERE_COLONNE, 130

End Sub


' =========================================================
' FEUILLE PARAMETRES
'
' Le tableau des statistiques dit quelles colonnes le
' classement affiche, dans quel ordre, et sur quelle
' action du journal chacune compte.
'
' Une ligne sans motif compte toutes les occurrences de
' l'action. Une ligne avec motif ne compte que les lignes
' du journal portant ce motif : c'est ainsi que les
' penalites concedees se detaillent par faute.
'
' A droite, les listes qui alimentent les menus de la
' feuille Filtres. Elles sont reecrites a chaque ajout.
' =========================================================
Private Sub ConstruireFeuilleParametres()

    Dim ws As Worksheet
    Dim Tableau As ListObject

    Set ws = FeuilleRecapOuCreee(FEUILLE_PARAMETRES)

    EcrireTitre _
        ws, _
        "Statistiques du classement", _
        "D" & ChrW(233) & "coche une ligne pour retirer la " & _
        "colonne du classement, ou r" & ChrW(233) & "ordonne les " & _
        "lignes pour changer l'ordre des colonnes."

    CreerTableau _
        ws, _
        TBL_STATS, _
        Array( _
            "Afficher", "Colonne", "Action", "Motif", _
            "Groupe" _
        )

    Set Tableau = ws.ListObjects(TBL_STATS)

    ' Une reconstruction ne doit pas effacer la liste que
    ' l'entraineur a pu adapter : elle n'est remplie que
    ' lorsque le tableau est encore vierge.
    If Application.CountA(Tableau.DataBodyRange) = 0 Then
        RemplirStatistiquesParDefaut Tableau
    End If

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 9
    ws.Columns(PREMIERE_COLONNE).HorizontalAlignment = xlCenter
    ws.Columns(PREMIERE_COLONNE + 1).ColumnWidth = 26
    ws.Columns(PREMIERE_COLONNE + 2).ColumnWidth = 30
    ws.Columns(PREMIERE_COLONNE + 3).ColumnWidth = 18
    ws.Columns(PREMIERE_COLONNE + 4).ColumnWidth = 16

    PoserListeCoche ws, PREMIERE_COLONNE

    ws.Range("I5").Value = "Listes des menus de la feuille Filtres"
    ws.Range("I5").Font.Bold = True

    ws.Range("I6").Value = "Saisons"
    ws.Range("J6").Value = "Cat" & ChrW(233) & "gories"
    ws.Range("K6").Value = "Phases"
    ws.Range("L6").Value = "Lieux"
    ws.Range("M6").Value = "R" & ChrW(233) & "sultats"
    ws.Range("N6").Value = "Adversaires"

    ws.Range("I6:N6").Font.Italic = True
    ws.Columns("I:N").ColumnWidth = 22

    ' Les plages nommees doivent exister des la
    ' construction : sans elles, les listes deroulantes de
    ' la feuille Filtres refuseraient d'etre posees.
    NommerListeVide "LST_SAISONS", ws, "I"
    NommerListeVide "LST_CATEGORIES", ws, "J"
    NommerListeVide "LST_PHASES", ws, "K"
    NommerListeVide "LST_LIEUX", ws, "L"
    NommerListeVide "LST_RESULTATS", ws, "M"
    NommerListeVide "LST_ADVERSAIRES", ws, "N"

End Sub


' =========================================================
' Liste des statistiques livree avec le classeur. Elle
' reprend les actions du journal qui sont imputees a un
' joueur, groupees comme dans la feuille Stats match d'un
' fichier de match.
'
' Les penalites concedees apparaissent deux fois : en
' total, puis detaillees par motif. Le detail est decoche
' au depart pour que le classement reste lisible.
'
' Une instruction par statistique, et non un seul tableau
' litteral : VBA n'accepte que vingt-cinq lignes de
' continuation par instruction, et la liste en demanderait
' davantage.
' =========================================================
Private Sub RemplirStatistiquesParDefaut( _
    ByVal Tableau As ListObject _
)

    Dim Lignes() As Variant
    Dim NbStats As Long
    Dim PremiereLigne As Long

    ReDim Lignes(1 To MAX_STATS_PAR_DEFAUT, 1 To 5)

    AjouterStat Lignes, NbStats, True, _
        "Plaquages r" & ChrW(233) & "ussis", _
        "Plaquage normal", "", _
        "D" & ChrW(233) & "fense"

    AjouterStat Lignes, NbStats, True, _
        "Plaquages offensifs", _
        "Plaquage offensif", "", _
        "D" & ChrW(233) & "fense"

    AjouterStat Lignes, NbStats, True, _
        "Plaquages " & ChrW(224) & " 2", _
        "Plaquage " & ChrW(224) & " 2", "", _
        "D" & ChrW(233) & "fense"

    AjouterStat Lignes, NbStats, True, _
        "Plaquages rat" & ChrW(233) & "s", _
        "Plaquage rat" & ChrW(233), "", _
        "D" & ChrW(233) & "fense"

    AjouterStat Lignes, NbStats, True, _
        "Plaquages hauts", _
        "Plaquage haut", "", _
        "D" & ChrW(233) & "fense"

    AjouterStat Lignes, NbStats, True, _
        "Grattages", "Grattage", "", _
        "Conqu" & ChrW(234) & "te"

    AjouterStat Lignes, NbStats, True, _
        "Contre-rucks", "Contre-ruck", "", _
        "Conqu" & ChrW(234) & "te"

    AjouterStat Lignes, NbStats, True, _
        "Arrachages", "Arrachage", "", _
        "Conqu" & ChrW(234) & "te"

    AjouterStat Lignes, NbStats, True, _
        "Franchissements", "Franchissement", "", _
        "Attaque"

    AjouterStat Lignes, NbStats, True, _
        "En-avants", "En avant", "", _
        "Attaque"

    AjouterStat Lignes, NbStats, True, _
        "Jeux au pied", "Jeu au pied", "", _
        "Attaque"

    AjouterStat Lignes, NbStats, True, _
        "R" & ChrW(233) & "ceptions", _
        "R" & ChrW(233) & "ception nous", "", _
        "Attaque"

    AjouterStat Lignes, NbStats, True, _
        "Essais", "Essai", "", "Points"

    AjouterStat Lignes, NbStats, True, _
        "Transformations", "Transformation", "", "Points"

    AjouterStat Lignes, NbStats, True, _
        "P" & ChrW(233) & "nalit" & ChrW(233) & "s pass" & ChrW(233) & "es", _
        "P" & ChrW(233) & "nalit" & ChrW(233), "", _
        "Points"

    AjouterStat Lignes, NbStats, True, _
        "Drops", "Drop", "", "Points"

    AjouterStat Lignes, NbStats, True, _
        "Touches conserv" & ChrW(233) & "es", _
        "Touche " & ChrW(224) & " nous conserv" & ChrW(233) & "e", "", _
        "Touches"

    AjouterStat Lignes, NbStats, True, _
        "Touches perdues", _
        "Touche " & ChrW(224) & " nous perdue", "", _
        "Touches"

    AjouterStat Lignes, NbStats, True, _
        "Touches pas droites", _
        "Touche " & ChrW(224) & " nous pas droite", "", _
        "Touches"

    AjouterStat Lignes, NbStats, True, _
        "Touches vol" & ChrW(233) & "es", _
        "Touche " & ChrW(224) & " eux vol" & ChrW(233) & "e", "", _
        "Touches"

    AjouterStat Lignes, NbStats, True, _
        "P" & ChrW(233) & "nalit" & ChrW(233) & "s conc" & ChrW(233) & "d" & ChrW(233) & "es", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre nous", "", _
        "Discipline"

    AjouterStat Lignes, NbStats, True, _
        "P" & ChrW(233) & "nalit" & ChrW(233) & "s obtenues", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre adversaire", "", _
        "Discipline"

    AjouterStat Lignes, NbStats, True, _
        "Cartons jaunes", _
        "Carton jaune contre nous", "", _
        "Discipline"

    AjouterStat Lignes, NbStats, True, _
        "Cartons rouges", _
        "Carton rouge contre nous", "", _
        "Discipline"

    ' Detail des penalites concedees : decoche au depart.
    AjouterStat Lignes, NbStats, False, _
        "P" & ChrW(233) & "n. ruck", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre nous", "Ruck", _
        "Discipline"

    AjouterStat Lignes, NbStats, False, _
        "P" & ChrW(233) & "n. maul", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre nous", "Maul", _
        "Discipline"

    AjouterStat Lignes, NbStats, False, _
        "P" & ChrW(233) & "n. hors-jeu", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre nous", "Hors-jeu", _
        "Discipline"

    AjouterStat Lignes, NbStats, False, _
        "P" & ChrW(233) & "n. plaquage " & ChrW(224) & " 2", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre nous", _
        "Plaquage " & ChrW(224) & " 2", _
        "Discipline"

    AjouterStat Lignes, NbStats, False, _
        "P" & ChrW(233) & "n. plaquage haut", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & " contre nous", _
        "Plaquage haut", _
        "Discipline"

    If NbStats = 0 Then
        Exit Sub
    End If

    ViderTableauRecap Tableau

    PremiereLigne = PreparerLignesTableau(Tableau, NbStats)

    ' Le tableau est dimensionne large : ecrire dans une
    ' plage plus courte n'en prend que le debut.
    Tableau.Parent.Cells( _
        PremiereLigne, _
        Tableau.Range.Column _
    ).Resize(NbStats, 5).Value = Lignes

End Sub


' =========================================================
' Pose une statistique dans le tableau en cours de
' constitution.
' =========================================================
Private Sub AjouterStat( _
    ByRef Lignes() As Variant, _
    ByRef NbStats As Long, _
    ByVal Afficher As Boolean, _
    ByVal Colonne As String, _
    ByVal Action As String, _
    ByVal Motif As String, _
    ByVal Groupe As String _
)

    If NbStats >= MAX_STATS_PAR_DEFAUT Then
        Exit Sub
    End If

    NbStats = NbStats + 1

    If Afficher Then
        Lignes(NbStats, 1) = MARQUE_COCHE
    End If

    Lignes(NbStats, 2) = Colonne
    Lignes(NbStats, 3) = Action
    Lignes(NbStats, 4) = Motif
    Lignes(NbStats, 5) = Groupe

End Sub


' =========================================================
' OUTILS DE MISE EN PAGE
' =========================================================

Private Sub EcrireTitre( _
    ByVal ws As Worksheet, _
    ByVal Titre As String, _
    ByVal SousTitre As String _
)

    With ws.Cells(2, PREMIERE_COLONNE)
        .Value = Titre
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(150, 30, 30)
    End With

    With ws.Cells(3, PREMIERE_COLONNE)
        .Value = SousTitre
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = RGB(110, 110, 110)
    End With

    ws.Rows(2).RowHeight = 26

    ' La ligne des boutons doit les contenir en entier.
    ws.Rows(LIGNE_ENTETE - 2).RowHeight = 34

End Sub


' =========================================================
' Cree un tableau structure sous la ligne d'en-tete, ou le
' laisse en place s'il existe deja.
' =========================================================
Private Sub CreerTableau( _
    ByVal ws As Worksheet, _
    ByVal NomTableau As String, _
    ByVal Entetes As Variant _
)

    Dim Tableau As ListObject
    Dim NbColonnes As Long
    Dim Plage As Range

    On Error Resume Next
    Set Tableau = ws.ListObjects(NomTableau)
    On Error GoTo 0

    If Not Tableau Is Nothing Then
        Exit Sub
    End If

    NbColonnes = UBound(Entetes) - LBound(Entetes) + 1

    ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE) _
        .Resize(1, NbColonnes).Value = Entetes

    Set Plage = _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE) _
            .Resize(NB_LIGNES_DEPART + 1, NbColonnes)

    Set Tableau = _
        ws.ListObjects.Add( _
            xlSrcRange, _
            Plage, _
            , _
            xlYes _
        )

    Tableau.Name = NomTableau
    Tableau.TableStyle = "TableStyleMedium3"

    With Tableau.HeaderRowRange
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlBottom
    End With

End Sub


' =========================================================
' Bouton dessine dans la feuille. Decalage en points
' depuis le bord gauche de la premiere colonne, pour
' aligner plusieurs boutons sur une meme ligne.
' =========================================================
Private Sub PoserBouton( _
    ByVal ws As Worksheet, _
    ByVal Libelle As String, _
    ByVal Macro As String, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long, _
    ByVal Largeur As Single, _
    Optional ByVal Decalage As Single = 0 _
)

    Dim Forme As Shape
    Dim NomForme As String
    Dim i As Long

    NomForme = "BTN_" & Macro

    ' Une seconde construction ne doit pas empiler les
    ' boutons les uns sur les autres.
    For i = ws.Shapes.Count To 1 Step -1

        If ws.Shapes(i).Name = NomForme Then
            ws.Shapes(i).Delete
        End If

    Next i

    Set Forme = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            ws.Cells(Ligne, Colonne).Left + Decalage, _
            ws.Cells(Ligne, Colonne).Top, _
            Largeur, _
            26 _
        )

    Forme.Name = NomForme
    Forme.OnAction = Macro

    With Forme.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(150, 30, 30)
        .Solid
    End With

    Forme.Line.Visible = msoFalse

    With Forme.TextFrame2.TextRange
        .Text = Libelle
        .Font.Size = 11
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Forme.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Forme.TextFrame2.HorizontalAnchor = msoAnchorCenter

End Sub


Private Sub PoserListe( _
    ByVal Cellule As Range, _
    ByVal Source As String _
)

    With Cellule.Validation

        .Delete

        On Error Resume Next

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=Source

        On Error GoTo 0

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

End Sub


' =========================================================
' Menu a deux choix sur toute une colonne d'un tableau :
' la marque de coche, ou rien.
' =========================================================
Private Sub PoserListeCoche( _
    ByVal ws As Worksheet, _
    ByVal Colonne As Long _
)

    PoserListe _
        ws.Columns(Colonne).Cells( _
            LIGNE_ENTETE + 1 _
        ).Resize(5000), _
        MARQUE_COCHE

End Sub


Private Sub EcrireCritere( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String, _
    ByVal NomCellule As String, _
    ByVal ValeurDepart As Variant _
)

    Dim Valeur As Range

    ws.Cells(Ligne, PREMIERE_COLONNE).Value = Libelle
    ws.Cells(Ligne, PREMIERE_COLONNE).Font.Bold = True

    Set Valeur = ws.Cells(Ligne, PREMIERE_COLONNE + 1)

    With Valeur
        .Interior.Color = RGB(255, 242, 224)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With

    ' Une reconstruction ne doit pas ecraser les criteres
    ' deja choisis par l'entraineur.
    If Len(TexteCellule(Valeur.Value)) = 0 Then
        Valeur.Value = ValeurDepart
    End If

    ThisWorkbook.Names.Add _
        Name:=NomCellule, _
        RefersTo:=Valeur

End Sub


' =========================================================
' Cree une plage nommee d'une seule cellule, que l'ajout de
' matchs remplacera par la liste reelle.
'
' Sans elle, la liste deroulante qui s'y refere ne peut pas
' etre posee. La cellule porte des maintenant "(Tous)" :
' une validation dont la source est vide est refusee.
' =========================================================
Private Sub NommerListeVide( _
    ByVal NomPlage As String, _
    ByVal ws As Worksheet, _
    ByVal Colonne As String _
)

    Dim Existante As Name

    On Error Resume Next
    Set Existante = ThisWorkbook.Names(NomPlage)
    On Error GoTo 0

    If Not Existante Is Nothing Then
        Exit Sub
    End If

    If Len(TexteCellule(ws.Range(Colonne & "7").Value)) = 0 Then
        ws.Range(Colonne & "7").Value = VALEUR_TOUS
    End If

    ThisWorkbook.Names.Add _
        Name:=NomPlage, _
        RefersTo:=ws.Range(Colonne & "7")

End Sub


Private Sub FigerVolets( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long _
)

    Dim FeuilleActive As Object

    On Error Resume Next

    Set FeuilleActive = ActiveSheet

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(Ligne + 1, Colonne + 1).Select
    ActiveWindow.FreezePanes = True

    If Not FeuilleActive Is Nothing Then
        FeuilleActive.Activate
    End If

    On Error GoTo 0

End Sub


' =========================================================
' Remet les feuilles dans l'ordre de lecture : ce que
' l'entraineur consulte d'abord, puis ce qu'il alimente,
' puis ce qu'il parametre.
' =========================================================
Private Sub RangerFeuilles()

    Dim Ordre As Variant
    Dim i As Long
    Dim ws As Worksheet

    Ordre = Array( _
        FEUILLE_CLASSEMENT, FEUILLE_TEMPS, FEUILLE_FILTRES, _
        FEUILLE_MATCHS, FEUILLE_JOURNAL, _
        FEUILLE_COMPOSITIONS, FEUILLE_PARAMETRES _
    )

    For i = 0 To UBound(Ordre)

        Set ws = FeuilleRecap(CStr(Ordre(i)))

        If Not ws Is Nothing Then

            If i = 0 Then
                ws.Move Before:=ThisWorkbook.Worksheets(1)
            Else
                ws.Move After:=ThisWorkbook.Worksheets(i)
            End If

        End If

    Next i

    ' Le classeur neuf arrive avec une feuille vide que
    ' rien ne remplace.
    SupprimerFeuilleVideParDefaut

End Sub


Private Sub SupprimerFeuilleVideParDefaut()

    Dim ws As Worksheet
    Dim Connues As String
    Dim EtatAlertes As Boolean
    Dim i As Long

    Connues = _
        "|" & FEUILLE_CLASSEMENT & "|" & FEUILLE_TEMPS & _
        "|" & FEUILLE_FILTRES & "|" & FEUILLE_MATCHS & _
        "|" & FEUILLE_JOURNAL & "|" & FEUILLE_COMPOSITIONS & _
        "|" & FEUILLE_PARAMETRES & "|"

    EtatAlertes = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ' Parcours descendant par index : supprimer une feuille
    ' pendant un For Each sauterait la suivante.
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1

        If ThisWorkbook.Worksheets.Count > 1 Then

            Set ws = ThisWorkbook.Worksheets(i)

            If InStr(1, Connues, "|" & ws.Name & "|", vbTextCompare) = 0 Then

                If Application.CountA(ws.Cells) = 0 Then
                    ws.Delete
                End If

            End If

            Set ws = Nothing

        End If

    Next i

    Application.DisplayAlerts = EtatAlertes

End Sub


' =========================================================
' BOUTONS DE CIBLAGE DE LA FEUILLE MATCHS
' =========================================================

Public Sub ToutCibler()

    CocherColonneCibler True

End Sub


Public Sub RienCibler()

    CocherColonneCibler False

End Sub


Private Sub CocherColonneCibler( _
    ByVal Cocher As Boolean _
)

    Dim Tableau As ListObject
    Dim Colonne As Range

    Set Tableau = TableauRecap(FEUILLE_MATCHS, TBL_MATCHS)

    If Tableau Is Nothing Then
        Exit Sub
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Sub
    End If

    Set Colonne = Tableau.ListColumns(MTC_CIBLER).DataBodyRange

    If Cocher Then
        Colonne.Value = MARQUE_COCHE
    Else
        Colonne.ClearContents
    End If

End Sub
