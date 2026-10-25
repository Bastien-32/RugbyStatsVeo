Attribute VB_Name = "modRecapCalcul"
Option Explicit

' =========================================================
' RECAPITULATIF DE SAISON - FILTRES ET CLASSEMENT
'
' Un seul point d'entree, AppliquerFiltres, qui enchaine
' quatre etapes :
'
'   1. lire les criteres de la feuille Filtres ;
'   2. en deduire les matchs retenus, et les marquer dans
'      la feuille Matchs ;
'   3. parcourir le journal et les compositions en ne
'      comptant que ces matchs ;
'   4. reecrire le classement et les temps de jeu.
'
' Les statistiques sont celles du match entier : la
' colonne Mi-temps du journal n'entre pas dans le comptage.
'
' Les deux tableaux de sortie sont reconstruits a chaque
' fois : leurs colonnes dependent des statistiques cochees
' dans Parametres et des categories rencontrees, qui
' changent au fil de la saison.
'
' Le comptage se fait en memoire plutot qu'avec des
' formules NB.SI.ENS : sur une saison entiere, un tableau
' de formules croisant quarante joueurs et trente
' statistiques recalculerait des millions de comparaisons
' a chaque frappe.
' =========================================================

Private Type Criteres

    Saison As String
    Categorie As String
    Phase As String
    Lieu As String
    Resultat As String
    Adversaire As String

    DateDebut As Double
    DateFin As Double

    Derniers As Long
    CiblesSeulement As Boolean
    MinutesMini As Double

    Mode As String

End Type

' Colonnes d'identite du classement, avant les minutes :
' joueur, ligne de jeu, feuilles de match, matchs joues.
' Viennent ensuite une colonne de minutes par categorie,
' le total des minutes, puis les statistiques.
' Le nom du joueur, et rien d'autre : le classement ne
' montre que les statistiques choisies. Le temps de jeu a
' sa feuille, et c'est la qu'on va le chercher.
' =========================================================
' Les couleurs des feuilles de match, relevees dans
' "Stats match" et "Compo" : le recapitulatif doit se lire
' comme elles.
'
' Le corps garde le vert pale et le blanc des feuilles de
' match. Les trois en-tetes, eux, tranchent franchement :
' c'est la seule facon de ne jamais lire un chiffre dans
' la mauvaise colonne.
' =========================================================
Private Const VERT_ENTETE As Long = 3057485   ' RGB(77, 167, 46)
Private Const VERT_FONCE As Long = 2392636    ' RGB(60, 130, 36)

' Les trois groupes, en teintes franches : bleu pour la
' premiere, orange pour la reserve, et le vert du club
' pour le cumul, qui couronne les deux.
Private Const TETE_PREMIERE As Long = 12611584 ' RGB(0, 112, 192)
Private Const TETE_RESERVE As Long = 36095     ' RGB(255, 140, 0)
Private Const TETE_TOTAL As Long = 3057485     ' RGB(77, 167, 46)

Private Const GRIS_FILET As Long = 5855577     ' RGB(89, 89, 89)

' Bandes vertes et blanches, au plus pres des feuilles de
' match parmi les styles qu'Excel propose.
Private Const STYLE_TABLEAU As String = "TableStyleLight21"
Private Const VERT_PALE As Long = 13693658    ' RGB(218, 242, 208)
Private Const BLANC As Long = 16777215


Private Const COLONNES_FIXES As Long = 1

' Duree de reference des ratios.
Private Const MINUTES_REFERENCE As Double = 80


' =========================================================
' LE TABLEAU DES ACTIONS NON ATTRIBUEES
'
' Pose sous le classement, il reprend les memes colonnes de
' statistiques et compte les actions dont la colonne Joueur
' ne porte pas un nom, mais l'un de ces quatre labels. Le
' classement les ecarte, faute de savoir a qui les mettre ;
' elles disparaissaient donc du recapitulatif alors qu'elles
' renseignent sur le jeu.
'
' Memes libelles et meme ordre que le tableau de la feuille
' Stats match d'un fichier de match.
' =========================================================
Private Const NA_AVANT As Long = 1
Private Const NA_TROIS_QUARTS As Long = 2
Private Const NA_COLLECTIF As Long = 3
Private Const NA_NON_IDENTIFIE As Long = 4

Private Const NB_NON_ATTRIBUE As Long = 4

' Lignes laissees entre le bas du classement et ce tableau.
Private Const ECART_NON_ATTRIBUE As Long = 3

' Lignes laissees entre le bas de ce tableau et la ligne de
' total : deux vides, la troisieme porte le total.
Private Const ECART_TOTAL As Long = 3

Private Const TBL_NON_ATTRIBUE As String = "RECAP_NON_ATTRIBUE"


' =========================================================
' Point d'entree, appele par les boutons des feuilles
' Filtres, Classement, Temps de jeu et Compositions, et a
' la fin de chaque ajout de matchs.
' =========================================================
Public Sub AppliquerFiltres()

    Dim Filtres As Criteres

    Dim Matchs As Collection
    Dim NbMatchsRetenus As Long

    Dim Joueurs As Collection
    Dim NomsJoueurs() As String
    Dim LignesJoueurs() As String
    Dim NbJoueurs As Long

    Dim Categories As Collection
    Dim NomsCategories() As String
    Dim NbCategories As Long

    Dim IndexAction As Collection
    Dim IndexActionMotif As Collection
    Dim NomsStats() As String
    Dim NbStats As Long

    Dim Compteurs() As Double
    Dim CompteursNA() As Double
    Dim MinutesCat() As Double
    Dim MatchsCat() As Long
    Dim Feuilles() As Long
    Dim Joues() As Long

    Dim EtatAffichage As Boolean

    EtatAffichage = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    LireCriteres Filtres

    NbMatchsRetenus = RetenirMatchs(Filtres, Matchs)

    NbStats = _
        ChargerStatistiques( _
            IndexAction, IndexActionMotif, NomsStats _
        )

    NbJoueurs = _
        RecenserJoueurs( _
            Matchs, Joueurs, NomsJoueurs, LignesJoueurs _
        )

    NbCategories = _
        RecenserCategories(Matchs, Categories, NomsCategories)

    If NbJoueurs > 0 Then

        ReDim Compteurs(1 To NbJoueurs, 1 To NbStats)
        ReDim MinutesCat(1 To NbJoueurs, 1 To NbCategories)
        ReDim MatchsCat(1 To NbJoueurs, 1 To NbCategories)
        ReDim Feuilles(1 To NbJoueurs)
        ReDim Joues(1 To NbJoueurs)

        CompterActions _
            Matchs, Joueurs, _
            IndexAction, IndexActionMotif, _
            Compteurs

        CompterTempsJeu _
            Matchs, Joueurs, Categories, _
            MinutesCat, MatchsCat, Feuilles, Joues

    End If

    ' Hors du test sur NbJoueurs : un journal peut n'avoir
    ' ete saisi qu'avec ces labels, sans qu'aucun joueur n'y
    ' soit jamais nomme.
    ReDim CompteursNA(1 To NB_NON_ATTRIBUE, 1 To Maxi(NbStats, 1))

    CompterActionsNonAttribuees _
        Matchs, IndexAction, IndexActionMotif, CompteursNA

    EcrireClassement _
        Filtres, NbMatchsRetenus, _
        NomsJoueurs, LignesJoueurs, NbJoueurs, _
        NomsCategories, NbCategories, _
        NomsStats, NbStats, _
        Compteurs, MinutesCat, Feuilles, Joues, _
        CompteursNA

    EcrireTempsDeJeu _
        Filtres, NbMatchsRetenus, _
        NomsJoueurs, LignesJoueurs, NbJoueurs, _
        NomsCategories, NbCategories, _
        MinutesCat, MatchsCat, Feuilles, Joues

    ActualiserLibelleMode Filtres.Mode

    Application.ScreenUpdating = EtatAffichage

    Exit Sub

GestionErreur:

    Dim DescriptionErreur As String

    DescriptionErreur = Err.Description

    Application.ScreenUpdating = EtatAffichage

    MsgBox _
        "Le calcul a " & ChrW(233) & "chou" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & DescriptionErreur, _
        vbCritical, _
        "Erreur"

End Sub


' =========================================================
' LECTURE DES CRITERES
' =========================================================
Private Sub LireCriteres( _
    ByRef Filtres As Criteres _
)

    Filtres.Saison = ValeurFiltre("FILTRE_SAISON")
    Filtres.Categorie = ValeurFiltre("FILTRE_CATEGORIE")
    Filtres.Phase = ValeurFiltre("FILTRE_PHASE")
    Filtres.Lieu = ValeurFiltre("FILTRE_LIEU")
    Filtres.Resultat = ValeurFiltre("FILTRE_RESULTAT")
    Filtres.Adversaire = ValeurFiltre("FILTRE_ADVERSAIRE")

    Filtres.DateDebut = NombreFiltre("FILTRE_DATE_DEBUT")
    Filtres.DateFin = NombreFiltre("FILTRE_DATE_FIN")
    Filtres.Derniers = CLng(NombreFiltre("FILTRE_DERNIERS"))
    Filtres.MinutesMini = NombreFiltre("FILTRE_MINUTES_MIN")

    Filtres.CiblesSeulement = _
        (Normaliser(ValeurFiltre("FILTRE_CIBLES")) = "oui")

    Filtres.Mode = ValeurFiltre("RECAP_MODE")

    If Len(Filtres.Mode) = 0 Then
        Filtres.Mode = MODE_TOTAUX
    End If

End Sub


Private Function ValeurFiltre( _
    ByVal NomCellule As String _
) As String

    On Error Resume Next

    ValeurFiltre = _
        TexteCellule( _
            ThisWorkbook.Names(NomCellule).RefersToRange.Value _
        )

    On Error GoTo 0

End Function


Private Function NombreFiltre( _
    ByVal NomCellule As String _
) As Double

    Dim Valeur As Variant

    On Error Resume Next

    Valeur = _
        ThisWorkbook.Names(NomCellule).RefersToRange.Value

    On Error GoTo 0

    NombreFiltre = NombreCellule(Valeur)

End Function


' =========================================================
' MATCHS RETENUS
'
' Renvoie le nombre de matchs retenus et remplit une
' collection indexee sur le nom du fichier : c'est elle que
' consultent ensuite les boucles de comptage.
'
' La limite aux N derniers matchs s'applique apres les
' autres criteres, pour que "les trois derniers matchs de
' reserve" designe bien les trois derniers de reserve.
' =========================================================
Private Function RetenirMatchs( _
    ByRef Filtres As Criteres, _
    ByRef Matchs As Collection _
) As Long

    Dim Tableau As ListObject
    Dim Donnees As Variant
    Dim Retenu() As Boolean
    Dim Dates() As Double
    Dim NbLignes As Long
    Dim Ligne As Long
    Dim Seuil As Double
    Dim Cle As String
    Dim NbRetenus As Long

    Set Matchs = New Collection

    Set Tableau = TableauRecap(FEUILLE_MATCHS, TBL_MATCHS)

    If Tableau Is Nothing Then
        Exit Function
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Function
    End If

    Donnees = Tableau.DataBodyRange.Value

    NbLignes = UBound(Donnees, 1)

    ReDim Retenu(1 To NbLignes)
    ReDim Dates(1 To NbLignes)

    For Ligne = 1 To NbLignes

        Dates(Ligne) = NombreCellule(Donnees(Ligne, MTC_DATE))

        Retenu(Ligne) = _
            MatchPasseLesFiltres(Filtres, Donnees, Ligne)

    Next Ligne

    If Filtres.Derniers > 0 Then

        Seuil = _
            DateDuNiemeMatch( _
                Retenu, Dates, NbLignes, Filtres.Derniers _
            )

        For Ligne = 1 To NbLignes

            If Retenu(Ligne) Then

                If Dates(Ligne) < Seuil Then
                    Retenu(Ligne) = False
                End If

            End If

        Next Ligne

    End If

    For Ligne = 1 To NbLignes

        If Retenu(Ligne) Then

            NbRetenus = NbRetenus + 1

            Cle = _
                Normaliser(TexteCellule(Donnees(Ligne, MTC_FICHIER)))

            If Len(Cle) > 0 Then

                If Not CleExiste(Matchs, Cle) Then
                    Matchs.Add Ligne, Cle
                End If

            End If

        End If

    Next Ligne


    RetenirMatchs = NbRetenus

End Function


Private Function MatchPasseLesFiltres( _
    ByRef Filtres As Criteres, _
    ByRef Donnees As Variant, _
    ByVal Ligne As Long _
) As Boolean

    Dim DateMatch As Double

    ' La selection manuelle court-circuite les criteres :
    ' l'entraineur qui coche des matchs precis n'a pas a
    ' remettre les menus sur (Tous).
    If Filtres.CiblesSeulement Then

        MatchPasseLesFiltres = EstCoche(Donnees(Ligne, MTC_ANALYSER))
        Exit Function

    End If

    If Not CritereRespecte( _
        Filtres.Saison, Donnees(Ligne, MTC_SAISON) _
    ) Then
        Exit Function
    End If

    If Not CritereRespecte( _
        Filtres.Categorie, Donnees(Ligne, MTC_CATEGORIE) _
    ) Then
        Exit Function
    End If

    If Not CritereRespecte( _
        Filtres.Phase, Donnees(Ligne, MTC_PHASE) _
    ) Then
        Exit Function
    End If

    If Not CritereRespecte( _
        Filtres.Lieu, Donnees(Ligne, MTC_LIEU) _
    ) Then
        Exit Function
    End If

    If Not CritereRespecte( _
        Filtres.Resultat, Donnees(Ligne, MTC_RESULTAT) _
    ) Then
        Exit Function
    End If

    If Not CritereRespecte( _
        Filtres.Adversaire, Donnees(Ligne, MTC_ADVERSAIRE) _
    ) Then
        Exit Function
    End If

    DateMatch = NombreCellule(Donnees(Ligne, MTC_DATE))

    If Filtres.DateDebut > 0 Then

        If DateMatch < Filtres.DateDebut Then
            Exit Function
        End If

    End If

    If Filtres.DateFin > 0 Then

        If DateMatch > Filtres.DateFin Then
            Exit Function
        End If

    End If

    MatchPasseLesFiltres = True

End Function


Private Function CritereRespecte( _
    ByVal Critere As String, _
    ByVal Valeur As Variant _
) As Boolean

    If Len(Critere) = 0 Then
        CritereRespecte = True
        Exit Function
    End If

    If Normaliser(Critere) = Normaliser(VALEUR_TOUS) Then
        CritereRespecte = True
        Exit Function
    End If

    CritereRespecte = _
        (Normaliser(Critere) = Normaliser(TexteCellule(Valeur)))

End Function


' =========================================================
' Date du Nieme match le plus recent parmi ceux deja
' retenus. Les matchs joues le meme jour comptent ensemble :
' garder l'un et pas l'autre n'aurait pas de sens.
' =========================================================
Private Function DateDuNiemeMatch( _
    ByRef Retenu() As Boolean, _
    ByRef Dates() As Double, _
    ByVal NbLignes As Long, _
    ByVal Combien As Long _
) As Double

    Dim Triees() As Double
    Dim NbDates As Long
    Dim Tampon As Double
    Dim i As Long
    Dim j As Long

    ReDim Triees(1 To NbLignes)

    For i = 1 To NbLignes

        If Retenu(i) Then

            NbDates = NbDates + 1
            Triees(NbDates) = Dates(i)

        End If

    Next i

    If NbDates = 0 Then
        Exit Function
    End If

    ' Tri decroissant par insertion : le nombre de matchs
    ' d'une saison ne justifie pas mieux.
    For i = 2 To NbDates

        Tampon = Triees(i)
        j = i - 1

        Do While j >= 1
            If Triees(j) >= Tampon Then
                Exit Do
            End If
            Triees(j + 1) = Triees(j)
            j = j - 1
        Loop

        Triees(j + 1) = Tampon

    Next i

    If Combien >= NbDates Then
        DateDuNiemeMatch = Triees(NbDates)
    Else
        DateDuNiemeMatch = Triees(Combien)
    End If

End Function


' =========================================================
' STATISTIQUES DU CLASSEMENT
'
' Deux index sont construits : l'un sur l'action seule,
' l'autre sur le couple action et motif. Une action dont
' le motif est renseigne alimente les deux, de sorte que
' "Penalites concedees" totalise les fautes que
' "Pen. ruck" detaille.
' =========================================================
Private Function ChargerStatistiques( _
    ByRef IndexAction As Collection, _
    ByRef IndexActionMotif As Collection, _
    ByRef NomsStats() As String _
) As Long

    Dim Donnees As Variant
    Dim Ligne As Long
    Dim NbStats As Long
    Dim Libelle As String
    Dim Action As String
    Dim Motif As String
    Dim Cle As String

    Set IndexAction = New Collection
    Set IndexActionMotif = New Collection

    Donnees = DonneesTableau(FEUILLE_PARAMETRES, TBL_STATS)

    If IsEmpty(Donnees) Then
        ChargerStatistiques = AucuneStatistique(NomsStats)
        Exit Function
    End If

    ReDim NomsStats(1 To Maxi(UBound(Donnees, 1), 1))

    For Ligne = 1 To UBound(Donnees, 1)

        If EstCoche(Donnees(Ligne, 1)) Then

            Libelle = TexteCellule(Donnees(Ligne, 2))
            Action = TexteCellule(Donnees(Ligne, 3))
            Motif = TexteCellule(Donnees(Ligne, 4))

            If Len(Libelle) > 0 And Len(Action) > 0 Then

                NbStats = NbStats + 1
                NomsStats(NbStats) = Libelle

                If Len(Motif) = 0 Then

                    Cle = Normaliser(Action)

                    If Not CleExiste(IndexAction, Cle) Then
                        IndexAction.Add NbStats, Cle
                    End If

                Else

                    Cle = Normaliser(Action) & "|" & Normaliser(Motif)

                    If Not CleExiste(IndexActionMotif, Cle) Then
                        IndexActionMotif.Add NbStats, Cle
                    End If

                End If

            End If

        End If

    Next Ligne

    If NbStats = 0 Then
        NbStats = AucuneStatistique(NomsStats)
    End If

    ChargerStatistiques = NbStats

End Function


' =========================================================
' Le classement garde toujours au moins une colonne de
' statistique : sans elle, les tableaux de comptage
' seraient dimensionnes a zero.
' =========================================================
Private Function AucuneStatistique( _
    ByRef NomsStats() As String _
) As Long

    ReDim NomsStats(1 To 1)

    NomsStats(1) = "Aucune statistique coch" & ChrW(233) & "e"

    AucuneStatistique = 1

End Function


' =========================================================
' RECENSEMENT DES JOUEURS
'
' Les joueurs viennent d'abord des compositions, pour que
' celui qui n'a fait aucune action figure quand meme au
' classement. Le journal complete la liste, au cas ou une
' action serait imputee a un joueur absent de la compo.
' =========================================================
Private Function RecenserJoueurs( _
    ByVal Matchs As Collection, _
    ByRef Joueurs As Collection, _
    ByRef NomsJoueurs() As String, _
    ByRef LignesJoueurs() As String _
) As Long

    Dim Compositions As Variant
    Dim Actions As Variant
    Dim NbJoueurs As Long
    Dim Capacite As Long
    Dim Ligne As Long
    Dim Nom As String
    Dim Cle As String

    Set Joueurs = New Collection

    Capacite = 200

    ReDim NomsJoueurs(1 To Capacite)
    ReDim LignesJoueurs(1 To Capacite)

    Compositions = _
        DonneesTableau(FEUILLE_COMPOSITIONS, TBL_COMPOSITIONS)

    If Not IsEmpty(Compositions) Then

        For Ligne = 1 To UBound(Compositions, 1)

            If MatchRetenu( _
                Matchs, Compositions(Ligne, CMP_FICHIER) _
            ) Then

                Nom = TexteCellule(Compositions(Ligne, CMP_JOUEUR))
                Cle = Normaliser(Nom)

                If Len(Cle) > 0 Then

                    If Not CleExiste(Joueurs, Cle) Then

                        NbJoueurs = NbJoueurs + 1

                        AgrandirListeJoueurs _
                            NomsJoueurs, LignesJoueurs, _
                            Capacite, NbJoueurs

                        Joueurs.Add NbJoueurs, Cle
                        NomsJoueurs(NbJoueurs) = Nom

                        LignesJoueurs(NbJoueurs) = _
                            TexteCellule(Compositions(Ligne, CMP_LIGNE))

                    ElseIf Len(LignesJoueurs(CLng(Joueurs.Item(Cle)))) = 0 Then

                        ' La ligne de jeu n'est pas toujours
                        ' saisie : on garde la premiere
                        ' renseignee.
                        LignesJoueurs(CLng(Joueurs.Item(Cle))) = _
                            TexteCellule(Compositions(Ligne, CMP_LIGNE))

                    End If

                End If

            End If

        Next Ligne

    End If

    Actions = DonneesTableau(FEUILLE_JOURNAL, TBL_JOURNAL)

    If Not IsEmpty(Actions) Then

        For Ligne = 1 To UBound(Actions, 1)

            If MatchRetenu(Matchs, Actions(Ligne, JRN_FICHIER)) Then

                Nom = TexteCellule(Actions(Ligne, JRN_JOUEUR))

                If EstJoueur(Nom) Then

                    Cle = Normaliser(Nom)

                    If Not CleExiste(Joueurs, Cle) Then

                        NbJoueurs = NbJoueurs + 1

                        AgrandirListeJoueurs _
                            NomsJoueurs, LignesJoueurs, _
                            Capacite, NbJoueurs

                        Joueurs.Add NbJoueurs, Cle
                        NomsJoueurs(NbJoueurs) = Nom

                    End If

                End If

            End If

        Next Ligne

    End If

    RecenserJoueurs = NbJoueurs

End Function


Private Sub AgrandirListeJoueurs( _
    ByRef NomsJoueurs() As String, _
    ByRef LignesJoueurs() As String, _
    ByRef Capacite As Long, _
    ByVal Necessaire As Long _
)

    If Necessaire <= Capacite Then
        Exit Sub
    End If

    Capacite = Capacite * 2

    ReDim Preserve NomsJoueurs(1 To Capacite)
    ReDim Preserve LignesJoueurs(1 To Capacite)

End Sub


' =========================================================
' Categories rencontrees dans les matchs retenus. Premiere
' et Reserve viennent en tete quand elles sont presentes :
' c'est la distinction que l'entraineur regarde en premier.
' =========================================================
Private Function RecenserCategories( _
    ByVal Matchs As Collection, _
    ByRef Categories As Collection, _
    ByRef NomsCategories() As String _
) As Long

    Dim Compositions As Variant
    Dim Ordre As Collection
    Dim Trouvees As Collection
    Dim NbCategories As Long
    Dim Ligne As Long
    Dim Nom As String
    Dim Cle As String
    Dim i As Long

    Set Categories = New Collection
    Set Trouvees = New Collection
    Set Ordre = New Collection

    Compositions = _
        DonneesTableau(FEUILLE_COMPOSITIONS, TBL_COMPOSITIONS)

    If Not IsEmpty(Compositions) Then

        For Ligne = 1 To UBound(Compositions, 1)

            If MatchRetenu( _
                Matchs, Compositions(Ligne, CMP_FICHIER) _
            ) Then

                Nom = TexteCellule(Compositions(Ligne, CMP_CATEGORIE))

                If Len(Nom) = 0 Then
                    Nom = "Non pr" & ChrW(233) & "cis" & ChrW(233) & "e"
                End If

                Cle = Normaliser(Nom)

                If Not CleExiste(Trouvees, Cle) Then
                    Trouvees.Add Nom, Cle
                    Ordre.Add Nom
                End If

            End If

        Next Ligne

    End If

    ReDim NomsCategories(1 To Maxi(Ordre.Count, 1))

    If CleExiste(Trouvees, "premiere") Then
        NbCategories = NbCategories + 1
        NomsCategories(NbCategories) = CStr(Trouvees.Item("premiere"))
        Categories.Add NbCategories, "premiere"
    End If

    If CleExiste(Trouvees, "reserve") Then
        NbCategories = NbCategories + 1
        NomsCategories(NbCategories) = CStr(Trouvees.Item("reserve"))
        Categories.Add NbCategories, "reserve"
    End If

    For i = 1 To Ordre.Count

        Nom = CStr(Ordre.Item(i))
        Cle = Normaliser(Nom)

        If Not CleExiste(Categories, Cle) Then

            NbCategories = NbCategories + 1
            NomsCategories(NbCategories) = Nom
            Categories.Add NbCategories, Cle

        End If

    Next i

    If NbCategories = 0 Then
        NbCategories = 1
        NomsCategories(1) = "Cat" & ChrW(233) & "gorie"
    End If

    RecenserCategories = NbCategories

End Function


' =========================================================
' COMPTAGE
'
' Toutes les lignes du match comptent, quelle que soit leur
' mi-temps : la statistique d'un match est le total des
' deux periodes.
' =========================================================
Private Sub CompterActions( _
    ByVal Matchs As Collection, _
    ByVal Joueurs As Collection, _
    ByVal IndexAction As Collection, _
    ByVal IndexActionMotif As Collection, _
    ByRef Compteurs() As Double _
)

    Dim Actions As Variant
    Dim Ligne As Long
    Dim Joueur As Long
    Dim Stat As Long
    Dim Nom As String
    Dim CleJoueur As String
    Dim CleAction As String
    Dim CleMotif As String

    Actions = DonneesTableau(FEUILLE_JOURNAL, TBL_JOURNAL)

    If IsEmpty(Actions) Then
        Exit Sub
    End If

    For Ligne = 1 To UBound(Actions, 1)

        If MatchRetenu(Matchs, Actions(Ligne, JRN_FICHIER)) Then

            Nom = TexteCellule(Actions(Ligne, JRN_JOUEUR))

            If EstJoueur(Nom) Then

                CleJoueur = Normaliser(Nom)

                If CleExiste(Joueurs, CleJoueur) Then

                    Joueur = CLng(Joueurs.Item(CleJoueur))

                    CleAction = _
                        Normaliser( _
                            TexteCellule(Actions(Ligne, JRN_ACTION)) _
                        )

                    If CleExiste(IndexAction, CleAction) Then

                        Stat = CLng(IndexAction.Item(CleAction))

                        Compteurs(Joueur, Stat) = _
                            Compteurs(Joueur, Stat) + 1

                    End If

                    CleMotif = _
                        Normaliser( _
                            TexteCellule(Actions(Ligne, JRN_MOTIF)) _
                        )

                    If Len(CleMotif) > 0 Then

                        CleMotif = CleAction & "|" & CleMotif

                        If CleExiste(IndexActionMotif, CleMotif) Then

                            Stat = CLng(IndexActionMotif.Item(CleMotif))

                            Compteurs(Joueur, Stat) = _
                                Compteurs(Joueur, Stat) + 1

                        End If

                    End If

                End If

            End If

        End If

    Next Ligne

End Sub


' =========================================================
' Le meme comptage que CompterActions, pour les seules
' actions que personne n'a pu endosser : celles dont la
' colonne Joueur porte Avant, Trois-quarts, Collectif ou le
' point d'interrogation.
'
' Les joueurs nommes n'entrent pas ici : ils ont deja leur
' ligne au classement, et les compter une seconde fois
' doublerait les totaux.
'
' Une case Joueur vide n'est comptee nulle part : c'est une
' action portee a l'adversaire.
' =========================================================
Private Sub CompterActionsNonAttribuees( _
    ByVal Matchs As Collection, _
    ByVal IndexAction As Collection, _
    ByVal IndexActionMotif As Collection, _
    ByRef CompteursNA() As Double _
)

    Dim Actions As Variant
    Dim Ligne As Long
    Dim Bloc As Long
    Dim Stat As Long
    Dim CleAction As String
    Dim CleMotif As String

    Actions = DonneesTableau(FEUILLE_JOURNAL, TBL_JOURNAL)

    If IsEmpty(Actions) Then
        Exit Sub
    End If

    For Ligne = 1 To UBound(Actions, 1)

        If MatchRetenu(Matchs, Actions(Ligne, JRN_FICHIER)) Then

            Bloc = _
                BlocNonAttribue( _
                    TexteCellule(Actions(Ligne, JRN_JOUEUR)) _
                )

            If Bloc > 0 Then

                CleAction = _
                    Normaliser( _
                        TexteCellule(Actions(Ligne, JRN_ACTION)) _
                    )

                If CleExiste(IndexAction, CleAction) Then

                    Stat = CLng(IndexAction.Item(CleAction))

                    CompteursNA(Bloc, Stat) = _
                        CompteursNA(Bloc, Stat) + 1

                End If

                CleMotif = _
                    Normaliser( _
                        TexteCellule(Actions(Ligne, JRN_MOTIF)) _
                    )

                If Len(CleMotif) > 0 Then

                    CleMotif = CleAction & "|" & CleMotif

                    If CleExiste(IndexActionMotif, CleMotif) Then

                        Stat = _
                            CLng(IndexActionMotif.Item(CleMotif))

                        CompteursNA(Bloc, Stat) = _
                            CompteursNA(Bloc, Stat) + 1

                    End If

                End If

            End If

        End If

    Next Ligne

End Sub


' =========================================================
' A quel bloc du tableau revient une action, d'apres le
' contenu de sa colonne Joueur. Rend zero des qu'un nom de
' joueur y figure : le classement s'en charge.
' =========================================================
Private Function BlocNonAttribue( _
    ByVal Nom As String _
) As Long

    ' Avant toute normalisation : celle-ci ne garde que les
    ' lettres et les chiffres, le point d'interrogation s'y
    ' reduirait a rien et passerait pour une case vide.
    If Trim(Nom) = "?" Then
        BlocNonAttribue = NA_NON_IDENTIFIE
        Exit Function
    End If

    Select Case Normaliser(Nom)

        Case "avant"
            BlocNonAttribue = NA_AVANT

        Case "troisquarts"
            ' Le trait d'union de "Trois-quarts" a disparu
            ' a la normalisation.
            BlocNonAttribue = NA_TROIS_QUARTS

        Case "collectif"
            BlocNonAttribue = NA_COLLECTIF

    End Select

End Function


' =========================================================
' Une ligne de composition vaut une feuille de match. Elle
' ne compte comme match joue que si un temps de jeu y a ete
' saisi : un remplacant reste sur le banc.
' =========================================================
Private Sub CompterTempsJeu( _
    ByVal Matchs As Collection, _
    ByVal Joueurs As Collection, _
    ByVal Categories As Collection, _
    ByRef MinutesCat() As Double, _
    ByRef MatchsCat() As Long, _
    ByRef Feuilles() As Long, _
    ByRef Joues() As Long _
)

    Dim Compositions As Variant
    Dim Ligne As Long
    Dim Joueur As Long
    Dim Categorie As Long
    Dim Nom As String
    Dim Cle As String
    Dim Minutes As Double

    Compositions = _
        DonneesTableau(FEUILLE_COMPOSITIONS, TBL_COMPOSITIONS)

    If IsEmpty(Compositions) Then
        Exit Sub
    End If

    For Ligne = 1 To UBound(Compositions, 1)

        If MatchRetenu( _
            Matchs, Compositions(Ligne, CMP_FICHIER) _
        ) Then

            Nom = TexteCellule(Compositions(Ligne, CMP_JOUEUR))
            Cle = Normaliser(Nom)

            If CleExiste(Joueurs, Cle) Then

                Joueur = CLng(Joueurs.Item(Cle))

                Categorie = _
                    IndexCategorie( _
                        Categories, _
                        Compositions(Ligne, CMP_CATEGORIE) _
                    )

                Feuilles(Joueur) = Feuilles(Joueur) + 1

                Minutes = _
                    NombreCellule(Compositions(Ligne, CMP_MINUTES))

                If Minutes > 0 Then

                    Joues(Joueur) = Joues(Joueur) + 1

                    If Categorie > 0 Then

                        MinutesCat(Joueur, Categorie) = _
                            MinutesCat(Joueur, Categorie) + Minutes

                        MatchsCat(Joueur, Categorie) = _
                            MatchsCat(Joueur, Categorie) + 1

                    End If

                End If

            End If

        End If

    Next Ligne

End Sub


Private Function IndexCategorie( _
    ByVal Categories As Collection, _
    ByVal Valeur As Variant _
) As Long

    Dim Cle As String

    Cle = Normaliser(TexteCellule(Valeur))

    If Len(Cle) = 0 Then
        Cle = Normaliser("Non pr" & ChrW(233) & "cis" & ChrW(233) & "e")
    End If

    If CleExiste(Categories, Cle) Then
        IndexCategorie = CLng(Categories.Item(Cle))
    End If

End Function


' =========================================================
' Les lignes de la base sont groupees par match : garder la
' reponse precedente evite de relancer une recherche par
' cle, et surtout l'erreur VBA qu'une cle absente declenche,
' sur chacune des milliers de lignes d'une saison.
' =========================================================
Private Function MatchRetenu( _
    ByVal Matchs As Collection, _
    ByVal Fichier As Variant _
) As Boolean

    Static DerniereListe As Collection
    Static DerniereCle As String
    Static DerniereReponse As Boolean

    Dim Cle As String

    Cle = Normaliser(TexteCellule(Fichier))

    If Not DerniereListe Is Nothing Then

        If DerniereListe Is Matchs And Cle = DerniereCle Then
            MatchRetenu = DerniereReponse
            Exit Function
        End If

    End If

    Set DerniereListe = Matchs
    DerniereCle = Cle
    DerniereReponse = CleExiste(Matchs, Cle)

    MatchRetenu = DerniereReponse

End Function


' =========================================================
' ECRITURE DU CLASSEMENT
' =========================================================
Private Sub EcrireClassement( _
    ByRef Filtres As Criteres, _
    ByVal NbMatchsRetenus As Long, _
    ByRef NomsJoueurs() As String, _
    ByRef LignesJoueurs() As String, _
    ByVal NbJoueurs As Long, _
    ByRef NomsCategories() As String, _
    ByVal NbCategories As Long, _
    ByRef NomsStats() As String, _
    ByVal NbStats As Long, _
    ByRef Compteurs() As Double, _
    ByRef MinutesCat() As Double, _
    ByRef Feuilles() As Long, _
    ByRef Joues() As Long, _
    ByRef CompteursNA() As Double _
)

    Dim ws As Worksheet
    Dim Entetes As Variant
    Dim Sortie As Variant
    Dim NbColonnes As Long
    Dim NbRetenus As Long
    Dim Retenu() As Boolean
    Dim MinutesTotal() As Double

    Dim TotauxClassement() As Double

    Dim i As Long
    Dim c As Long
    Dim s As Long
    Dim r As Long
    Dim Ligne As Long
    Dim Colonne As Long

    Set ws = FeuilleRecap(FEUILLE_CLASSEMENT)

    If ws Is Nothing Then
        Exit Sub
    End If

    ' Le tableau des actions non attribuees vit sous le
    ' classement, dont la pose efface tout ce qui suit. On le
    ' delie d'abord, sinon il subsisterait comme tableau vide
    ' au milieu de la zone nettoyee.
    DelierTableau ws, TBL_NON_ATTRIBUE

    NbColonnes = COLONNES_FIXES + NbStats

    ' Les totaux de la ligne finale se prennent sur ce qui
    ' est affiche, et non sur les totaux bruts : ainsi la
    ' somme se verifie a la main, colonne par colonne, quel
    ' que soit le mode d'affichage.
    ReDim TotauxClassement(1 To Maxi(NbStats, 1))

    ReDim Entetes(1 To 1, 1 To NbColonnes)

    Entetes(1, 1) = "Joueur"

    For s = 1 To NbStats
        Entetes(1, COLONNES_FIXES + s) = NomsStats(s)
    Next s

    ' Joueurs retenus : ceux qui atteignent le seuil de
    ' minutes demande.
    ReDim Retenu(1 To Maxi(NbJoueurs, 1))
    ReDim MinutesTotal(1 To Maxi(NbJoueurs, 1))

    For i = 1 To NbJoueurs

        For c = 1 To NbCategories
            MinutesTotal(i) = MinutesTotal(i) + MinutesCat(i, c)
        Next c

        Retenu(i) = (MinutesTotal(i) >= Filtres.MinutesMini)

        If Retenu(i) Then
            NbRetenus = NbRetenus + 1
        End If

    Next i

    EcrireBandeau ws, Filtres, NbMatchsRetenus, NbRetenus

    If NbRetenus = 0 Then

        ' Sortie n'a pas ete dimensionne : il vaut Empty, ce
        ' que PoserTableauSortie accepte puisqu'il n'ecrit
        ' aucune ligne.
        PoserTableauSortie _
            ws, "RECAP_CLASSEMENT", Entetes, Sortie, 0, NbColonnes

        EcrireTableauNonAttribue _
            ws, LIGNE_ENTETE + 1 + ECART_NON_ATTRIBUE, _
            NomsStats, NbStats, NbColonnes, _
            CompteursNA, NbMatchsRetenus, Filtres.Mode, _
            TotauxClassement

        Exit Sub

    End If

    ReDim Sortie(1 To NbRetenus, 1 To NbColonnes)

    For i = 1 To NbJoueurs

        If Retenu(i) Then

            Ligne = Ligne + 1

            Sortie(Ligne, 1) = NomsJoueurs(i)

            ' Les minutes servent encore au calcul des
            ' ratios, mais ne sont plus affichees ici.
            For s = 1 To NbStats

                Colonne = COLONNES_FIXES + s

                Sortie(Ligne, Colonne) = _
                    ValeurAffichee( _
                        Compteurs(i, s), _
                        Joues(i), _
                        MinutesTotal(i), _
                        Filtres.Mode _
                    )

            Next s

        End If

    Next i

    For r = 1 To NbRetenus

        For s = 1 To NbStats

            Colonne = COLONNES_FIXES + s

            If IsNumeric(Sortie(r, Colonne)) Then

                TotauxClassement(s) = _
                    TotauxClassement(s) + CDbl(Sortie(r, Colonne))

            End If

        Next s

    Next r

    PoserTableauSortie _
        ws, "RECAP_CLASSEMENT", Entetes, Sortie, NbRetenus, NbColonnes

    MettreEnFormeClassement _
        ws, NbRetenus, NbColonnes, NbCategories, Filtres.Mode

    EcrireTableauNonAttribue _
        ws, LIGNE_ENTETE + NbRetenus + ECART_NON_ATTRIBUE, _
        NomsStats, NbStats, NbColonnes, _
        CompteursNA, NbMatchsRetenus, Filtres.Mode, _
        TotauxClassement

End Sub


' =========================================================
' Le tableau des actions non attribuees, sous le classement.
'
' Quatre lignes toujours presentes, meme a zero : leur
' absence se lirait comme un oubli alors qu'un zero est une
' information.
'
' Le total ne figure pas dans le tableau : il porte sur les
' deux a la fois, et se pose a part, plus bas.
' =========================================================
Private Sub EcrireTableauNonAttribue( _
    ByVal ws As Worksheet, _
    ByVal Entete As Long, _
    ByRef NomsStats() As String, _
    ByVal NbStats As Long, _
    ByVal NbColonnes As Long, _
    ByRef CompteursNA() As Double, _
    ByVal NbMatchsRetenus As Long, _
    ByVal Mode As String, _
    ByRef TotauxClassement() As Double _
)

    Dim Libelles As Variant
    Dim Entetes As Variant
    Dim Sortie As Variant
    Dim Totaux() As Double

    Dim Minutes As Double
    Dim Valeur As Variant
    Dim i As Long
    Dim s As Long

    If NbStats <= 0 Then
        Exit Sub
    End If

    Libelles = Array( _
        "Avant", _
        "Trois-quarts", _
        "Collectif", _
        "?" _
    )

    ReDim Entetes(1 To 1, 1 To NbColonnes)

    Entetes(1, 1) = "Sans joueur nomm" & ChrW(233)

    For s = 1 To NbStats
        Entetes(1, COLONNES_FIXES + s) = NomsStats(s)
    Next s

    ' Les totaux du classement, auxquels s'ajoutent ceux de
    ' ce tableau : la ligne finale couvre les deux.
    ReDim Totaux(1 To NbStats)

    For s = 1 To NbStats
        Totaux(s) = TotauxClassement(s)
    Next s

    ReDim Sortie(1 To NB_NON_ATTRIBUE, 1 To NbColonnes)

    ' Ces actions n'appartiennent a personne : leur
    ' denominateur n'est pas un temps de jeu individuel mais
    ' le nombre de matchs retenus, et les quatre-vingts
    ' minutes qui vont avec.
    Minutes = NbMatchsRetenus * MINUTES_REFERENCE

    For i = 1 To NB_NON_ATTRIBUE

        Sortie(i, 1) = Libelles(i - 1)

        For s = 1 To NbStats

            Valeur = _
                ValeurAffichee( _
                    CompteursNA(i, s), _
                    NbMatchsRetenus, _
                    Minutes, _
                    Mode _
                )

            Sortie(i, COLONNES_FIXES + s) = Valeur

            If IsNumeric(Valeur) Then
                Totaux(s) = Totaux(s) + CDbl(Valeur)
            End If

        Next s

    Next i

    PoserTableauSortie _
        ws, TBL_NON_ATTRIBUE, Entetes, Sortie, _
        NB_NON_ATTRIBUE, NbColonnes, Entete

    MettreEnFormeNonAttribue _
        ws, Entete, NB_NON_ATTRIBUE, NbColonnes, Mode

    EcrireTotalGeneral _
        ws, Entete + NB_NON_ATTRIBUE + ECART_TOTAL, _
        NbStats, NbColonnes, Totaux, Mode

End Sub


' =========================================================
' La ligne de total, seule sous les deux tableaux.
'
' Elle n'appartient a aucun d'eux : elle additionne la
' colonne du classement et celle des actions non
' attribuees, pour donner enfin le compte complet de la
' saison. Hors tableau Excel, donc, sinon un tri la
' deplacerait parmi les donnees.
' =========================================================
Private Sub EcrireTotalGeneral( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal NbStats As Long, _
    ByVal NbColonnes As Long, _
    ByRef Totaux() As Double, _
    ByVal Mode As String _
)

    Dim Valeurs As Variant
    Dim Corps As Range
    Dim s As Long

    ReDim Valeurs(1 To 1, 1 To NbColonnes)

    Valeurs(1, 1) = "Total"

    For s = 1 To NbStats
        Valeurs(1, COLONNES_FIXES + s) = Totaux(s)
    Next s

    Set Corps = _
        ws.Cells(Ligne, PREMIERE_COLONNE).Resize(1, NbColonnes)

    Corps.Value = Valeurs

    With Corps

        .Font.Bold = True
        .Font.Color = BLANC
        .Interior.Color = VERT_ENTETE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin

    End With

    ws.Cells(Ligne, PREMIERE_COLONNE).HorizontalAlignment = xlLeft

    ' Meme regle que les deux tableaux pour les decimales,
    ' mais le zero reste affiche : sur une ligne unique, une
    ' case vide ne se lit plus comme un zero.
    If Mode = MODE_TOTAUX Then
        Corps.Offset(0, COLONNES_FIXES) _
            .Resize(1, NbColonnes - COLONNES_FIXES) _
            .NumberFormat = "0"
    Else
        Corps.Offset(0, COLONNES_FIXES) _
            .Resize(1, NbColonnes - COLONNES_FIXES) _
            .NumberFormat = "0.00"
    End If

End Sub


' =========================================================
' Delie un tableau sans toucher a son contenu. Rien a
' signaler s'il n'existe pas.
' =========================================================
Private Sub DelierTableau( _
    ByVal ws As Worksheet, _
    ByVal NomTableau As String _
)

    Dim Tableau As ListObject

    On Error Resume Next
    Set Tableau = ws.ListObjects(NomTableau)
    On Error GoTo 0

    If Not Tableau Is Nothing Then
        Tableau.Unlist
    End If

End Sub


Private Sub MettreEnFormeNonAttribue( _
    ByVal ws As Worksheet, _
    ByVal Entete As Long, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long, _
    ByVal Mode As String _
)

    Dim PremiereStat As Long
    Dim FormatNombre As String

    PremiereStat = PREMIERE_COLONNE + COLONNES_FIXES

    ' Meme regle que le classement : les zeros restent
    ' masques pour que le tableau se lise d'un coup d'oeil.
    If Mode = MODE_TOTAUX Then
        FormatNombre = "0;;"
    Else
        FormatNombre = "0.00;;"
    End If

    If NbLignes > 0 Then

        ws.Range( _
            ws.Cells(Entete + 1, PremiereStat), _
            ws.Cells( _
                Entete + NbLignes, _
                PREMIERE_COLONNE + NbColonnes - 1 _
            ) _
        ).NumberFormat = FormatNombre

    End If

    ws.Rows(Entete).RowHeight = 46

End Sub


' =========================================================
' Traduit un total brut dans le mode d'affichage choisi.
'
' Une case reste vide plutot que d'afficher zero quand le
' denominateur manque : un joueur sans minute saisie ne
' peut pas etre ramene a 80 minutes.
' =========================================================
Private Function ValeurAffichee( _
    ByVal Total As Double, _
    ByVal MatchsJoues As Long, _
    ByVal Minutes As Double, _
    ByVal Mode As String _
) As Variant

    Select Case Mode

        Case MODE_PAR_MATCH

            If MatchsJoues <= 0 Then
                Exit Function
            End If

            ValeurAffichee = Total / MatchsJoues

        Case MODE_PAR_80

            If Minutes <= 0 Then
                Exit Function
            End If

            ValeurAffichee = Total * MINUTES_REFERENCE / Minutes

        Case Else

            ValeurAffichee = Total

    End Select

End Function


' =========================================================
' ECRITURE DES TEMPS DE JEU
'
' Trois colonnes par categorie : matchs joues, minutes
' cumulees, moyenne par match. C'est la lecture qui separe
' la premiere de la reserve.
' =========================================================
Private Sub EcrireTempsDeJeu( _
    ByRef Filtres As Criteres, _
    ByVal NbMatchsRetenus As Long, _
    ByRef NomsJoueurs() As String, _
    ByRef LignesJoueurs() As String, _
    ByVal NbJoueurs As Long, _
    ByRef NomsCategories() As String, _
    ByVal NbCategories As Long, _
    ByRef MinutesCat() As Double, _
    ByRef MatchsCat() As Long, _
    ByRef Feuilles() As Long, _
    ByRef Joues() As Long _
)

    Dim ws As Worksheet
    Dim Entetes As Variant
    Dim Sortie As Variant
    Dim NbColonnes As Long
    Dim TotalMinutes As Double

    Dim i As Long
    Dim c As Long
    Dim Colonne As Long

    Set ws = FeuilleRecap(FEUILLE_TEMPS)

    If ws Is Nothing Then
        Exit Sub
    End If

    NbColonnes = 1 + NbCategories * 2 + 1

    ReDim Entetes(1 To 1, 1 To NbColonnes)

    Entetes(1, 1) = "Joueur"

    ' La categorie est portee par le regroupement pose au
    ' dessus : les en-tetes du tableau ne disent plus que
    ' la nature du chiffre.
    '
    ' Excel refuse deux colonnes de meme nom dans un
    ' tableau. Des espaces de fin, invisibles a l'ecran,
    ' les distinguent sans changer ce qu'on lit.
    For c = 1 To NbCategories

        Colonne = 1 + (c - 1) * 2

        Entetes(1, Colonne + 1) = "Matchs" & Space$(c)
        Entetes(1, Colonne + 2) = "Minutes" & Space$(c)

    Next c

    ' Le cumul ne garde que les minutes : le nombre de
    ' matchs se lit deja categorie par categorie.
    Colonne = 1 + NbCategories * 2

    Entetes(1, Colonne + 1) = "Minutes"

    EcrireBandeau ws, Filtres, NbMatchsRetenus, NbJoueurs

    If NbJoueurs = 0 Then

        PoserTableauSortie _
            ws, "RECAP_TEMPS_JOUEURS", Entetes, Sortie, 0, _
            NbColonnes, LIGNE_ENTETE + 1

        EcrireRegroupementTemps _
            ws, LIGNE_ENTETE, NomsCategories, NbCategories

        Exit Sub

    End If

    ReDim Sortie(1 To NbJoueurs, 1 To NbColonnes)

    For i = 1 To NbJoueurs

        Sortie(i, 1) = NomsJoueurs(i)

        TotalMinutes = 0

        For c = 1 To NbCategories

            Colonne = 1 + (c - 1) * 2

            Sortie(i, Colonne + 1) = MatchsCat(i, c)
            Sortie(i, Colonne + 2) = MinutesCat(i, c)

            TotalMinutes = TotalMinutes + MinutesCat(i, c)

        Next c

        Colonne = 1 + NbCategories * 2

        Sortie(i, Colonne + 1) = TotalMinutes

    Next i

    PoserTableauSortie _
        ws, "RECAP_TEMPS_JOUEURS", Entetes, Sortie, _
        NbJoueurs, NbColonnes, LIGNE_ENTETE + 1

    EcrireRegroupementTemps _
        ws, LIGNE_ENTETE, NomsCategories, NbCategories

    MettreEnFormeTemps ws, NbJoueurs, NbColonnes, NbCategories

End Sub


' =========================================================
' Pose un tableau de sortie : efface l'ancien, ecrit les
' en-tetes et les valeurs, recree le tableau structure
' pour retrouver le tri et les filtres automatiques.
' =========================================================
Private Sub PoserTableauSortie( _
    ByVal ws As Worksheet, _
    ByVal NomTableau As String, _
    ByRef Entetes As Variant, _
    ByRef Sortie As Variant, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long, _
    Optional ByVal Entete As Long = 0 _
)

    Dim Tableau As ListObject
    Dim Plage As Range
    Dim DerniereLigne As Long

    ' La feuille des temps de jeu reserve la ligne
    ' au-dessus a son regroupement par categorie : elle
    ' pose donc son tableau un rang plus bas.
    If Entete = 0 Then Entete = LIGNE_ENTETE

    On Error Resume Next
    Set Tableau = ws.ListObjects(NomTableau)
    On Error GoTo 0

    If Not Tableau Is Nothing Then
        Tableau.Unlist
        Set Tableau = Nothing
    End If

    DerniereLigne = _
        ws.Cells(ws.Rows.Count, PREMIERE_COLONNE).End(xlUp).Row

    If DerniereLigne >= Entete Then

        ws.Range( _
            ws.Cells(Entete, PREMIERE_COLONNE), _
            ws.Cells( _
                DerniereLigne + 5, _
                PREMIERE_COLONNE + 200 _
            ) _
        ).Clear

    End If

    ws.Cells(Entete, PREMIERE_COLONNE) _
        .Resize(1, NbColonnes).Value = Entetes

    If NbLignes > 0 Then

        ws.Cells(Entete + 1, PREMIERE_COLONNE) _
            .Resize(NbLignes, NbColonnes).Value = Sortie

    End If

    Set Plage = _
        ws.Cells(Entete, PREMIERE_COLONNE) _
            .Resize(Maxi(NbLignes, 1) + 1, NbColonnes)

    Set Tableau = _
        ws.ListObjects.Add(xlSrcRange, Plage, , xlYes)

    Tableau.Name = NomTableau

    ' Les bandes d'un tableau Excel suivent les lignes
    ' visibles : elles se recalculent au tri comme au
    ' filtre, ce qu'une couleur posee sur la cellule ne
    ' sait pas faire, et ce que la mise en forme
    ' conditionnelle refuse sur cette version.
    On Error Resume Next
    Tableau.TableStyle = STYLE_TABLEAU
    If Err.Number <> 0 Then
        Err.Clear
        Tableau.TableStyle = "TableStyleMedium7"
    End If
    On Error GoTo 0

    Tableau.ShowTableStyleRowStripes = True

    With Tableau.HeaderRowRange
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Interior.Color = VERT_ENTETE
        .Font.Color = BLANC
    End With

    PeindreCorps ws, Entete, NbLignes, NbColonnes

End Sub


' =========================================================
' Le corps d'un tableau, aux couleurs des feuilles de
' match : une ligne sur deux en vert pale, les autres en
' blanc, et un filet clair pour poser la grille.
' =========================================================
Private Sub PeindreCorps( _
    ByVal ws As Worksheet, _
    ByVal Entete As Long, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long _
)

    Dim Corps As Range

    If NbLignes <= 0 Then
        Exit Sub
    End If

    Set Corps = ws.Range( _
        ws.Cells(Entete + 1, PREMIERE_COLONNE), _
        ws.Cells(Entete + NbLignes, _
            PREMIERE_COLONNE + NbColonnes - 1))

    Corps.Font.Color = RGB(0, 0, 0)

    ' Aucun remplissage pose a la main : les bandes du
    ' style doivent rester visibles dessous.
    Corps.Interior.Pattern = xlNone

End Sub


Private Sub MettreEnFormeClassement( _
    ByVal ws As Worksheet, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long, _
    ByVal NbCategories As Long, _
    ByVal Mode As String _
)

    Dim PremiereStat As Long
    Dim FormatNombre As String

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 30

    ws.Range( _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + 1), _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + NbColonnes - 1) _
    ).EntireColumn.ColumnWidth = 10

    PremiereStat = PREMIERE_COLONNE + COLONNES_FIXES

    ' Les zeros sont masques, comme dans la feuille Stats
    ' match d'un fichier de match : un tableau de quarante
    ' joueurs sur trente colonnes devient illisible quand
    ' chaque case vide affiche un zero.
    If Mode = MODE_TOTAUX Then
        FormatNombre = "0;;"
    Else
        FormatNombre = "0.00;;"
    End If

    If NbLignes > 0 Then

        ws.Range( _
            ws.Cells(LIGNE_ENTETE + 1, PremiereStat), _
            ws.Cells( _
                LIGNE_ENTETE + NbLignes, _
                PREMIERE_COLONNE + NbColonnes - 1 _
            ) _
        ).NumberFormat = FormatNombre

    End If

    ws.Rows(LIGNE_ENTETE).RowHeight = 46

End Sub


Private Sub MettreEnFormeTemps( _
    ByVal ws As Worksheet, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long, _
    ByVal NbCategories As Long _
)

    Dim Entete As Long
    Dim c As Long
    Dim Colonne As Long

    Entete = LIGNE_ENTETE + 1

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 30

    ws.Range( _
        ws.Cells(Entete, PREMIERE_COLONNE + 1), _
        ws.Cells(Entete, PREMIERE_COLONNE + NbColonnes - 1) _
    ).EntireColumn.ColumnWidth = 12

    ' Les valeurs sont entieres : un format a decimale
    ' afficherait une virgule seule derriere chaque chiffre.
    ' Les deux sections vides laissent les zeros en blanc.
    If NbLignes > 0 Then

        ws.Range( _
            ws.Cells(Entete + 1, PREMIERE_COLONNE + 1), _
            ws.Cells( _
                Entete + NbLignes, _
                PREMIERE_COLONNE + NbColonnes - 1 _
            ) _
        ).NumberFormat = "0;;"

    End If

    ' Chaque groupe porte sa couleur, du bandeau jusqu'aux
    ' sous-titres, et un trait le separe du precedent.
    For c = 1 To NbCategories + 1

        Colonne = PREMIERE_COLONNE + 1 + (c - 1) * 2

        HabillerGroupeTemps _
            ws, Entete, Colonne, NbLignes, _
            CouleurGroupe(c, NbCategories), _
            LargeurGroupe(c, NbCategories)

    Next c

    ' Le tableau se ferme de part et d'autre, du vert du
    ' cumul : un cadre, et non deux traits qui pendent.
    EncadrerTableauTemps ws, Entete, NbLignes, NbColonnes

    ws.Rows(LIGNE_ENTETE).RowHeight = 22
    ws.Rows(Entete).RowHeight = 20

End Sub


Private Sub EncadrerTableauTemps( _
    ByVal ws As Worksheet, _
    ByVal Entete As Long, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long _
)

    Dim Haut As Long
    Dim Bas As Long

    Haut = Entete + 1
    Bas = Entete + Maxi(NbLignes, 1)

    With ws.Range( _
        ws.Cells(Haut, PREMIERE_COLONNE), _
        ws.Cells(Bas, PREMIERE_COLONNE) _
    ).Borders(xlEdgeLeft)

        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = TETE_TOTAL

    End With

    With ws.Range( _
        ws.Cells(Haut, PREMIERE_COLONNE + NbColonnes - 1), _
        ws.Cells(Bas, PREMIERE_COLONNE + NbColonnes - 1) _
    ).Borders(xlEdgeRight)

        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = TETE_TOTAL

    End With

    With ws.Range( _
        ws.Cells(Bas, PREMIERE_COLONNE), _
        ws.Cells(Bas, PREMIERE_COLONNE + NbColonnes - 1) _
    ).Borders(xlEdgeBottom)

        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = TETE_TOTAL

    End With

End Sub


' La premiere equipe en bleu, la reserve en orange, le
' cumul en vert : trois teintes franches, pour qu'un
' chiffre ne soit jamais lu dans la mauvaise colonne.
' Deux colonnes par equipe, une seule pour le cumul.
Private Function LargeurGroupe( _
    ByVal Rang As Long, _
    ByVal NbCategories As Long _
) As Long

    LargeurGroupe = IIf(Rang > NbCategories, 1, 2)

End Function


Private Function CouleurGroupe( _
    ByVal Rang As Long, _
    ByVal NbCategories As Long _
) As Long

    If Rang > NbCategories Then
        CouleurGroupe = TETE_TOTAL
    ElseIf Rang Mod 2 = 1 Then
        CouleurGroupe = TETE_PREMIERE
    Else
        CouleurGroupe = TETE_RESERVE
    End If

End Function


' Les deux colonnes d'un groupe : sous-titres a la couleur
' du bandeau, et un trait vertical a gauche qui descend sur
' toute la hauteur du tableau.
' =========================================================
' REGROUPEMENT DES COLONNES PAR CATEGORIE
'
' Un tableau structure n'a qu'une ligne d'en-tete : la
' categorie se pose au-dessus de lui, en cellules
' fusionnees, hors du tableau. Chaque equipe coiffe ses
' deux colonnes, le cumul les siennes.
' =========================================================
Private Sub EcrireRegroupementTemps( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByRef NomsCategories() As String, _
    ByVal NbCategories As Long _
)

    Dim c As Long
    Dim Colonne As Long
    Dim Zone As Range

    Set Zone = ws.Range( _
        ws.Cells(Ligne, PREMIERE_COLONNE), _
        ws.Cells(Ligne, PREMIERE_COLONNE + 200))

    Zone.UnMerge
    Zone.Clear

    For c = 1 To NbCategories

        Colonne = PREMIERE_COLONNE + 1 + (c - 1) * 2

        PoserGroupeTemps ws, Ligne, Colonne, _
            UCase$(NomsCategories(c)), _
            CouleurGroupe(c, NbCategories), _
            LargeurGroupe(c, NbCategories)

    Next c

    Colonne = PREMIERE_COLONNE + 1 + NbCategories * 2

    PoserGroupeTemps ws, Ligne, Colonne, "TOTAL", _
        CouleurGroupe(NbCategories + 1, NbCategories), _
        LargeurGroupe(NbCategories + 1, NbCategories)

End Sub


Private Sub PoserGroupeTemps( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long, _
    ByVal Titre As String, _
    ByVal Couleur As Long, _
    ByVal Largeur As Long _
)

    With ws.Cells(Ligne, Colonne).Resize(1, Largeur)

        .Merge
        .Value = Titre
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = Couleur
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)

    End With

End Sub


Private Sub HabillerGroupeTemps( _
    ByVal ws As Worksheet, _
    ByVal Entete As Long, _
    ByVal Colonne As Long, _
    ByVal NbLignes As Long, _
    ByVal Couleur As Long, _
    ByVal Largeur As Long _
)

    With ws.Cells(Entete, Colonne).Resize(1, Largeur)
        .Interior.Color = Couleur
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' Chaque groupe est precede d'un trait de sa propre
    ' couleur, sur les seules lignes de donnees : les deux
    ' rangees d'en-tete portent deja leur fond, un trait de
    ' plus les alourdirait.
    With ws.Range( _
        ws.Cells(Entete + 1, Colonne), _
        ws.Cells(Entete + Maxi(NbLignes, 1), Colonne) _
    ).Borders(xlEdgeLeft)

        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = Couleur

    End With

End Sub


' =========================================================
' Rappelle en clair, au-dessus du tableau, ce qui a servi
' au calcul. Sans ce rappel, un classement imprime ne dit
' plus sur quels matchs il porte.
' =========================================================
Private Sub EcrireBandeau( _
    ByVal ws As Worksheet, _
    ByRef Filtres As Criteres, _
    ByVal NbMatchs As Long, _
    ByVal NbJoueurs As Long _
)

    Dim Texte As String

    Texte = _
        NbMatchs & " match(s) retenu(s), " & _
        NbJoueurs & " joueur(s) - " & Filtres.Mode

    If Filtres.CiblesSeulement Then

        Texte = Texte & " - s" & ChrW(233) & "lection manuelle " & _
            "de la feuille Matchs"

    Else

        Texte = Texte & AjoutCritere("saison", Filtres.Saison)
        Texte = Texte & AjoutCritere("cat" & ChrW(233) & "gorie", Filtres.Categorie)
        Texte = Texte & AjoutCritere("phase", Filtres.Phase)
        Texte = Texte & AjoutCritere("lieu", Filtres.Lieu)
        Texte = Texte & AjoutCritere("r" & ChrW(233) & "sultat", Filtres.Resultat)
        Texte = Texte & AjoutCritere("adversaire", Filtres.Adversaire)

        If Filtres.Derniers > 0 Then
            Texte = Texte & " - " & Filtres.Derniers & " derniers matchs"
        End If

        If Filtres.DateDebut > 0 Then
            Texte = Texte & " - " & ChrW(224) & " partir du " & _
                Format$(Filtres.DateDebut, "dd/mm/yyyy")
        End If

        If Filtres.DateFin > 0 Then
            Texte = Texte & " - jusqu'au " & _
                Format$(Filtres.DateFin, "dd/mm/yyyy")
        End If

    End If

    With ws.Cells(LIGNE_ENTETE - 1, PREMIERE_COLONNE)
        .Value = Texte
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(150, 30, 30)
        .VerticalAlignment = xlBottom
    End With

End Sub


Private Function AjoutCritere( _
    ByVal Libelle As String, _
    ByVal Valeur As String _
) As String

    If Len(Valeur) = 0 Then
        Exit Function
    End If

    If Normaliser(Valeur) = Normaliser(VALEUR_TOUS) Then
        Exit Function
    End If

    AjoutCritere = " - " & Libelle & " : " & Valeur

End Function


' =========================================================
' BOUTON DE BASCULE DU MODE D'AFFICHAGE
'
' Un clic fait defiler les trois lectures possibles du
' classement, puis relance le calcul.
' =========================================================
Public Sub BasculerModeAffichage()

    Dim Cellule As Range
    Dim Mode As String

    On Error Resume Next
    Set Cellule = ThisWorkbook.Names("RECAP_MODE").RefersToRange
    On Error GoTo 0

    If Cellule Is Nothing Then
        Exit Sub
    End If

    Mode = TexteCellule(Cellule.Value)

    Select Case Mode
        Case MODE_TOTAUX: Mode = MODE_PAR_MATCH
        Case MODE_PAR_MATCH: Mode = MODE_PAR_80
        Case Else: Mode = MODE_TOTAUX
    End Select

    Cellule.Value = Mode

    AppliquerFiltres

End Sub


' =========================================================
' Le bouton porte le mode en cours : l'entraineur voit
' d'un coup d'oeil ce qu'il lit, et ce qu'un clic va
' changer.
' =========================================================
Private Sub ActualiserLibelleMode( _
    ByVal Mode As String _
)

    Dim ws As Worksheet
    Dim Forme As Shape

    Set ws = FeuilleRecap(FEUILLE_CLASSEMENT)

    If ws Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next
    Set Forme = ws.Shapes("BTN_BasculerModeAffichage")
    On Error GoTo 0

    If Forme Is Nothing Then
        Exit Sub
    End If

    Forme.TextFrame2.TextRange.Text = "Mode : " & Mode

End Sub


' =========================================================
' Remet tous les criteres a leur valeur de depart.
' =========================================================
Public Sub ReinitialiserFiltres()

    EcrireFiltre "FILTRE_SAISON", VALEUR_TOUS
    EcrireFiltre "FILTRE_CATEGORIE", VALEUR_TOUS
    EcrireFiltre "FILTRE_PHASE", VALEUR_TOUS
    EcrireFiltre "FILTRE_LIEU", VALEUR_TOUS
    EcrireFiltre "FILTRE_RESULTAT", VALEUR_TOUS
    EcrireFiltre "FILTRE_ADVERSAIRE", VALEUR_TOUS
    EcrireFiltre "FILTRE_DATE_DEBUT", ""
    EcrireFiltre "FILTRE_DATE_FIN", ""
    EcrireFiltre "FILTRE_DERNIERS", 0
    EcrireFiltre "FILTRE_CIBLES", "Non"
    EcrireFiltre "FILTRE_MINUTES_MIN", 0

    AppliquerFiltres

End Sub


Private Sub EcrireFiltre( _
    ByVal NomCellule As String, _
    ByVal Valeur As Variant _
)

    On Error Resume Next

    ThisWorkbook.Names(NomCellule).RefersToRange.Value = Valeur

    On Error GoTo 0

End Sub
