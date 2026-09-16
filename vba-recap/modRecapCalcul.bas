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
Private Const COLONNES_FIXES As Long = 4

' Duree de reference des ratios.
Private Const MINUTES_REFERENCE As Double = 80


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

    EcrireClassement _
        Filtres, NbMatchsRetenus, _
        NomsJoueurs, LignesJoueurs, NbJoueurs, _
        NomsCategories, NbCategories, _
        NomsStats, NbStats, _
        Compteurs, MinutesCat, Feuilles, Joues

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

    MarquerColonneRetenu Tableau, Retenu, NbLignes

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

        MatchPasseLesFiltres = EstCoche(Donnees(Ligne, MTC_CIBLER))
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


Private Sub MarquerColonneRetenu( _
    ByVal Tableau As ListObject, _
    ByRef Retenu() As Boolean, _
    ByVal NbLignes As Long _
)

    Dim Marques As Variant
    Dim Ligne As Long

    ReDim Marques(1 To NbLignes, 1 To 1)

    For Ligne = 1 To NbLignes

        If Retenu(Ligne) Then
            Marques(Ligne, 1) = MARQUE_COCHE
        End If

    Next Ligne

    Tableau.ListColumns(MTC_RETENU).DataBodyRange.Value = Marques

End Sub


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
    ByRef Joues() As Long _
)

    Dim ws As Worksheet
    Dim Entetes As Variant
    Dim Sortie As Variant
    Dim NbColonnes As Long
    Dim NbRetenus As Long
    Dim Retenu() As Boolean
    Dim MinutesTotal() As Double

    Dim i As Long
    Dim c As Long
    Dim s As Long
    Dim Ligne As Long
    Dim Colonne As Long

    Set ws = FeuilleRecap(FEUILLE_CLASSEMENT)

    If ws Is Nothing Then
        Exit Sub
    End If

    NbColonnes = COLONNES_FIXES + NbCategories + 1 + NbStats

    ReDim Entetes(1 To 1, 1 To NbColonnes)

    Entetes(1, 1) = "Joueur"
    Entetes(1, 2) = "Ligne"
    Entetes(1, 3) = "Feuilles"
    Entetes(1, 4) = "Matchs jou" & ChrW(233) & "s"

    For c = 1 To NbCategories
        Entetes(1, COLONNES_FIXES + c) = _
            "Min. " & NomsCategories(c)
    Next c

    Entetes(1, COLONNES_FIXES + NbCategories + 1) = _
        "Min. total"

    For s = 1 To NbStats
        Entetes(1, COLONNES_FIXES + NbCategories + 1 + s) = _
            NomsStats(s)
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

        Exit Sub

    End If

    ReDim Sortie(1 To NbRetenus, 1 To NbColonnes)

    For i = 1 To NbJoueurs

        If Retenu(i) Then

            Ligne = Ligne + 1

            Sortie(Ligne, 1) = NomsJoueurs(i)
            Sortie(Ligne, 2) = LignesJoueurs(i)
            Sortie(Ligne, 3) = Feuilles(i)
            Sortie(Ligne, 4) = Joues(i)

            For c = 1 To NbCategories
                Sortie(Ligne, COLONNES_FIXES + c) = MinutesCat(i, c)
            Next c

            Sortie(Ligne, COLONNES_FIXES + NbCategories + 1) = _
                MinutesTotal(i)

            For s = 1 To NbStats

                Colonne = COLONNES_FIXES + NbCategories + 1 + s

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

    PoserTableauSortie _
        ws, "RECAP_CLASSEMENT", Entetes, Sortie, NbRetenus, NbColonnes

    MettreEnFormeClassement _
        ws, NbRetenus, NbColonnes, NbCategories, Filtres.Mode

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

    NbColonnes = 3 + NbCategories * 3 + 3

    ReDim Entetes(1 To 1, 1 To NbColonnes)

    Entetes(1, 1) = "Joueur"
    Entetes(1, 2) = "Ligne"
    Entetes(1, 3) = "Feuilles"

    For c = 1 To NbCategories

        Colonne = 3 + (c - 1) * 3

        Entetes(1, Colonne + 1) = NomsCategories(c) & " - matchs"
        Entetes(1, Colonne + 2) = NomsCategories(c) & " - minutes"
        Entetes(1, Colonne + 3) = NomsCategories(c) & " - moyenne"

    Next c

    Colonne = 3 + NbCategories * 3

    Entetes(1, Colonne + 1) = "Total matchs jou" & ChrW(233) & "s"
    Entetes(1, Colonne + 2) = "Total minutes"
    Entetes(1, Colonne + 3) = "Moyenne par match"

    EcrireBandeau ws, Filtres, NbMatchsRetenus, NbJoueurs

    If NbJoueurs = 0 Then

        PoserTableauSortie _
            ws, "RECAP_TEMPS_JOUEURS", Entetes, Sortie, 0, NbColonnes

        Exit Sub

    End If

    ReDim Sortie(1 To NbJoueurs, 1 To NbColonnes)

    For i = 1 To NbJoueurs

        Sortie(i, 1) = NomsJoueurs(i)
        Sortie(i, 2) = LignesJoueurs(i)
        Sortie(i, 3) = Feuilles(i)

        TotalMinutes = 0

        For c = 1 To NbCategories

            Colonne = 3 + (c - 1) * 3

            Sortie(i, Colonne + 1) = MatchsCat(i, c)
            Sortie(i, Colonne + 2) = MinutesCat(i, c)

            If MatchsCat(i, c) > 0 Then

                Sortie(i, Colonne + 3) = _
                    MinutesCat(i, c) / MatchsCat(i, c)

            End If

            TotalMinutes = TotalMinutes + MinutesCat(i, c)

        Next c

        Colonne = 3 + NbCategories * 3

        Sortie(i, Colonne + 1) = Joues(i)
        Sortie(i, Colonne + 2) = TotalMinutes

        If Joues(i) > 0 Then
            Sortie(i, Colonne + 3) = TotalMinutes / Joues(i)
        End If

    Next i

    PoserTableauSortie _
        ws, "RECAP_TEMPS_JOUEURS", Entetes, Sortie, _
        NbJoueurs, NbColonnes

    MettreEnFormeTemps ws, NbJoueurs, NbColonnes

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
    ByVal NbColonnes As Long _
)

    Dim Tableau As ListObject
    Dim Plage As Range
    Dim DerniereLigne As Long

    On Error Resume Next
    Set Tableau = ws.ListObjects(NomTableau)
    On Error GoTo 0

    If Not Tableau Is Nothing Then
        Tableau.Unlist
        Set Tableau = Nothing
    End If

    DerniereLigne = _
        ws.Cells(ws.Rows.Count, PREMIERE_COLONNE).End(xlUp).Row

    If DerniereLigne >= LIGNE_ENTETE Then

        ws.Range( _
            ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE), _
            ws.Cells( _
                DerniereLigne + 5, _
                PREMIERE_COLONNE + 200 _
            ) _
        ).Clear

    End If

    ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE) _
        .Resize(1, NbColonnes).Value = Entetes

    If NbLignes > 0 Then

        ws.Cells(LIGNE_ENTETE + 1, PREMIERE_COLONNE) _
            .Resize(NbLignes, NbColonnes).Value = Sortie

    End If

    Set Plage = _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE) _
            .Resize(Maxi(NbLignes, 1) + 1, NbColonnes)

    Set Tableau = _
        ws.ListObjects.Add(xlSrcRange, Plage, , xlYes)

    Tableau.Name = NomTableau
    Tableau.TableStyle = "TableStyleMedium3"

    With Tableau.HeaderRowRange
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlBottom
    End With

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

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 28
    ws.Columns(PREMIERE_COLONNE + 1).ColumnWidth = 14

    ws.Range( _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + 2), _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + NbColonnes - 1) _
    ).EntireColumn.ColumnWidth = 10

    ' Les minutes restent des minutes quel que soit le
    ' mode : ce sont elles qui expliquent les ratios.
    If NbLignes > 0 Then

        ws.Range( _
            ws.Cells(LIGNE_ENTETE + 1, PREMIERE_COLONNE + 2), _
            ws.Cells( _
                LIGNE_ENTETE + NbLignes, _
                PREMIERE_COLONNE + COLONNES_FIXES + NbCategories _
            ) _
        ).NumberFormat = "0"

    End If

    PremiereStat = _
        PREMIERE_COLONNE + COLONNES_FIXES + NbCategories + 1

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
    ByVal NbColonnes As Long _
)

    ws.Columns(PREMIERE_COLONNE).ColumnWidth = 28
    ws.Columns(PREMIERE_COLONNE + 1).ColumnWidth = 14

    ws.Range( _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + 2), _
        ws.Cells(LIGNE_ENTETE, PREMIERE_COLONNE + NbColonnes - 1) _
    ).EntireColumn.ColumnWidth = 12

    If NbLignes > 0 Then

        ws.Range( _
            ws.Cells(LIGNE_ENTETE + 1, PREMIERE_COLONNE + 2), _
            ws.Cells( _
                LIGNE_ENTETE + NbLignes, _
                PREMIERE_COLONNE + NbColonnes - 1 _
            ) _
        ).NumberFormat = "0.#;;"

    End If

    ws.Rows(LIGNE_ENTETE).RowHeight = 46

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
