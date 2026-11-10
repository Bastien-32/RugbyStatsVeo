Attribute VB_Name = "modRecapPalmares"
Option Explicit

' =========================================================
' RECAPITULATIF DE SAISON - TOPS ET FLOPS
'
' Deux feuilles, Tops et Flops, ou chaque statistique
' choisie a son petit classement : rang, joueur, valeur.
' Elles remplacent l'extraction faite a la main, colonne
' par colonne, depuis la feuille Classement.
'
' La liste des classements vit dans la feuille Parametres,
' a droite des menus : une ligne par classement, avec son
' sens (Top ou Flop), le libelle d'une colonne du tableau
' des statistiques, et le nombre de places. Le tableau est
' cree, rempli des classements de depart, au premier
' calcul qui ne le trouve pas.
'
' Une statistique n'a pas besoin d'etre affichee au
' classement pour figurer ici : seul son libelle compte.
'
' Les deux feuilles suivent les memes filtres et le meme
' mode d'affichage que le classement : elles sont
' reecrites par AppliquerFiltres, a la suite des temps de
' jeu.
'
' Les ex aequo partagent le meme rang, et le rang suivant
' saute d'autant : 1, 2, 2, 4. Tous les joueurs a egalite
' sur la derniere place retenue sont montres, de sorte
' qu'un Top 10 peut compter douze noms.
'
' Une troisieme feuille, Penalites, classe les penalites
' concedees par motif, un tableau par categorie. C'est un
' bilan d'equipe : toutes les penalites comptent, qu'un
' joueur soit nomme ou non.
' =========================================================

Public Const FEUILLE_TOPS As String = "Tops"

Public Const FEUILLE_FLOPS As String = "Flops"

Public Const TBL_PALMARES As String = "RECAP_PALMARES"

' Le tableau commence en colonne P de Parametres, apres les
' listes des menus qui occupent I a N.
Private Const COLONNE_PALMARES As Long = 16

Private Const TAILLE_PAR_DEFAUT As Long = 10

Private Const NB_PALMARES_DEFAUT As Long = 9

' Mise en page : trois classements cote a cote, chacun sur
' trois colonnes (rang, joueur, valeur) suivies d'une
' colonne d'intervalle.
Private Const BLOCS_PAR_RANGEE As Long = 3

Private Const LARGEUR_BLOC As Long = 4

Private Const ECART_RANGEES As Long = 2

' Memes teintes que la feuille Classement : le vert du club
' pour les tops, le rouge des titres pour les flops.
Private Const VERT_ENTETE As Long = 3057485   ' RGB(77, 167, 46)
Private Const ROUGE_ENTETE As Long = 1973910  ' RGB(150, 30, 30)
Private Const VERT_PALE As Long = 13693658    ' RGB(218, 242, 208)
Private Const BLANC As Long = 16777215

' Les teintes des categories, reprises de la feuille Temps
' de jeu : bleu pour la premiere, orange pour la reserve.
Private Const TETE_PREMIERE As Long = 12611584 ' RGB(0, 112, 192)
Private Const TETE_RESERVE As Long = 36095     ' RGB(255, 140, 0)

' La categorie du match, dans le journal : la Premiere et
' la Reserve ont chacune leur colonne de comptage.
Private Const JRN_CATEGORIE As Long = 5

' L'action comptee par la feuille Penalites, normalisee.
Private Const ACTION_PENALITE As String = "penalitecontrenous"


' =========================================================
' Point d'entree, appele par AppliquerFiltres avec ce qu'il
' vient de recenser.
' =========================================================
Public Sub EcrirePalmares( _
    ByVal Matchs As Collection, _
    ByVal Joueurs As Collection, _
    ByRef NomsJoueurs() As String, _
    ByVal NbJoueurs As Long, _
    ByRef MinutesCat() As Double, _
    ByVal NbCategories As Long, _
    ByRef Joues() As Long, _
    ByVal Mode As String, _
    ByVal MinutesMini As Double _
)

    Dim Sens() As String
    Dim Libelles() As String
    Dim Tailles() As Long
    Dim StatDuPalmares() As Long
    Dim NbPalmares As Long

    Dim IndexAction As Collection
    Dim IndexActionMotif As Collection
    Dim NbStats As Long

    Dim Compteurs() As Double
    Dim MinutesTotal() As Double
    Dim Retenu() As Boolean

    Dim i As Long
    Dim c As Long

    PreparerTableauPalmares

    NbPalmares = _
        ChargerPalmares( _
            Sens, Libelles, Tailles, StatDuPalmares, _
            IndexAction, IndexActionMotif, NbStats _
        )

    ReDim Retenu(1 To Maxi(NbJoueurs, 1))
    ReDim MinutesTotal(1 To Maxi(NbJoueurs, 1))
    ReDim Compteurs(1 To Maxi(NbJoueurs, 1), 1 To Maxi(NbStats, 1))

    If NbJoueurs > 0 And NbStats > 0 Then

        CompterActions _
            Matchs, Joueurs, _
            IndexAction, IndexActionMotif, _
            Compteurs

    End If

    ' Meme seuil de minutes que le classement : un joueur
    ' qui n'y figure pas ne doit pas apparaitre ici.
    For i = 1 To NbJoueurs

        For c = 1 To NbCategories
            MinutesTotal(i) = MinutesTotal(i) + MinutesCat(i, c)
        Next c

        Retenu(i) = (MinutesTotal(i) >= MinutesMini)

    Next i

    EcrireFeuillePalmares _
        FEUILLE_TOPS, FEUILLE_TEMPS, "Tops", VERT_ENTETE, _
        Sens, Libelles, Tailles, StatDuPalmares, NbPalmares, _
        NomsJoueurs, NbJoueurs, Compteurs, Joues, MinutesTotal, _
        Retenu, Mode

    EcrireFeuillePalmares _
        FEUILLE_FLOPS, FEUILLE_TOPS, "Flops", ROUGE_ENTETE, _
        Sens, Libelles, Tailles, StatDuPalmares, NbPalmares, _
        NomsJoueurs, NbJoueurs, Compteurs, Joues, MinutesTotal, _
        Retenu, Mode

End Sub


' =========================================================
' LECTURE DES CLASSEMENTS DEMANDES
'
' Chaque libelle est cherche dans le tableau des
' statistiques, coche ou non, pour en connaitre l'action et
' le motif. Les index sont construits comme ceux du
' classement, afin que CompterActions s'en serve tel quel.
'
' Une statistique demandee deux fois, en top et en flop,
' n'est comptee qu'une fois. Un libelle introuvable garde sa
' place, avec zero pour statistique : son bloc le signale.
' =========================================================
Private Function ChargerPalmares( _
    ByRef Sens() As String, _
    ByRef Libelles() As String, _
    ByRef Tailles() As Long, _
    ByRef StatDuPalmares() As Long, _
    ByRef IndexAction As Collection, _
    ByRef IndexActionMotif As Collection, _
    ByRef NbStats As Long _
) As Long

    Dim Palmares As Variant
    Dim Stats As Variant
    Dim LigneDeStat As Collection
    Dim StatsVues As Collection
    Dim NbPalmares As Long
    Dim Ligne As Long
    Dim LigneStat As Long
    Dim Libelle As String
    Dim Action As String
    Dim Motif As String
    Dim Cle As String

    Set IndexAction = New Collection
    Set IndexActionMotif = New Collection
    Set LigneDeStat = New Collection
    Set StatsVues = New Collection

    NbStats = 0

    Palmares = DonneesTableau(FEUILLE_PARAMETRES, TBL_PALMARES)

    If IsEmpty(Palmares) Then
        Exit Function
    End If

    Stats = DonneesTableau(FEUILLE_PARAMETRES, TBL_STATS)

    If Not IsEmpty(Stats) Then

        For Ligne = 1 To UBound(Stats, 1)

            Cle = Normaliser(TexteCellule(Stats(Ligne, 2)))

            If Len(Cle) > 0 Then

                If Not CleExiste(LigneDeStat, Cle) Then
                    LigneDeStat.Add Ligne, Cle
                End If

            End If

        Next Ligne

    End If

    ReDim Sens(1 To UBound(Palmares, 1))
    ReDim Libelles(1 To UBound(Palmares, 1))
    ReDim Tailles(1 To UBound(Palmares, 1))
    ReDim StatDuPalmares(1 To UBound(Palmares, 1))

    For Ligne = 1 To UBound(Palmares, 1)

        Libelle = TexteCellule(Palmares(Ligne, 2))

        If Len(Libelle) > 0 Then

            NbPalmares = NbPalmares + 1

            Libelles(NbPalmares) = Libelle

            ' Tout ce qui ne commence pas par Flop est un
            ' top : une faute de frappe ne fait pas
            ' disparaitre le classement.
            If Left$(Normaliser(TexteCellule(Palmares(Ligne, 1))), 4) = "flop" Then
                Sens(NbPalmares) = FEUILLE_FLOPS
            Else
                Sens(NbPalmares) = FEUILLE_TOPS
            End If

            Tailles(NbPalmares) = CLng(NombreCellule(Palmares(Ligne, 3)))

            If Tailles(NbPalmares) <= 0 Then
                Tailles(NbPalmares) = TAILLE_PAR_DEFAUT
            End If

            Cle = Normaliser(Libelle)

            If CleExiste(StatsVues, Cle) Then

                StatDuPalmares(NbPalmares) = CLng(StatsVues.Item(Cle))

            ElseIf CleExiste(LigneDeStat, Cle) Then

                LigneStat = CLng(LigneDeStat.Item(Cle))

                Action = TexteCellule(Stats(LigneStat, 3))
                Motif = TexteCellule(Stats(LigneStat, 4))

                If Len(Action) > 0 Then

                    NbStats = NbStats + 1

                    StatsVues.Add NbStats, Cle
                    StatDuPalmares(NbPalmares) = NbStats

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

        End If

    Next Ligne

    ChargerPalmares = NbPalmares

End Function


' =========================================================
' ECRITURE D'UNE FEUILLE
'
' La feuille est videe sous son titre, puis les classements
' de son sens sont poses par rangees de trois. Une rangee
' prend la hauteur de son plus long classement, ex aequo
' compris.
' =========================================================
Private Sub EcrireFeuillePalmares( _
    ByVal NomFeuille As String, _
    ByVal FeuillePrecedente As String, _
    ByVal Titre As String, _
    ByVal CouleurEntete As Long, _
    ByRef Sens() As String, _
    ByRef Libelles() As String, _
    ByRef Tailles() As Long, _
    ByRef StatDuPalmares() As Long, _
    ByVal NbPalmares As Long, _
    ByRef NomsJoueurs() As String, _
    ByVal NbJoueurs As Long, _
    ByRef Compteurs() As Double, _
    ByRef Joues() As Long, _
    ByRef MinutesTotal() As Double, _
    ByRef Retenu() As Boolean, _
    ByVal Mode As String _
)

    Dim ws As Worksheet
    Dim FormatNombre As String
    Dim Rangs() As Long
    Dim Noms() As String
    Dim Valeurs() As Double
    Dim NbLignes As Long
    Dim NbBlocs As Long
    Dim LigneRangee As Long
    Dim HauteurRangee As Long
    Dim Hauteur As Long
    Dim Colonne As Long
    Dim p As Long
    Dim b As Long

    Set ws = _
        FeuillePalmares( _
            NomFeuille, FeuillePrecedente, Titre, _
            "Un classement par statistique, avec les filtres " & _
            "et le mode d'affichage du classement. Les ex " & _
            "aequo partagent le m" & ChrW(234) & "me rang. La liste " & _
            "se r" & ChrW(232) & "gle dans Parametres." _
        )

    ViderFeuillePalmares ws

    For b = 0 To BLOCS_PAR_RANGEE - 1
        ws.Columns(PREMIERE_COLONNE + b * LARGEUR_BLOC).ColumnWidth = 6
        ws.Columns(PREMIERE_COLONNE + b * LARGEUR_BLOC + 1).ColumnWidth = 26
        ws.Columns(PREMIERE_COLONNE + b * LARGEUR_BLOC + 2).ColumnWidth = 13
        ws.Columns(PREMIERE_COLONNE + b * LARGEUR_BLOC + 3).ColumnWidth = 3
    Next b

    RecopierBandeau ws

    If Mode = MODE_TOTAUX Then
        FormatNombre = "0"
    Else
        FormatNombre = "0.00"
    End If

    LigneRangee = LIGNE_ENTETE + 1

    For p = 1 To NbPalmares

        If Sens(p) = NomFeuille Then

            If NbBlocs > 0 And (NbBlocs Mod BLOCS_PAR_RANGEE) = 0 Then
                LigneRangee = LigneRangee + HauteurRangee + ECART_RANGEES
                HauteurRangee = 0
            End If

            Colonne = _
                PREMIERE_COLONNE + _
                (NbBlocs Mod BLOCS_PAR_RANGEE) * LARGEUR_BLOC

            NbLignes = _
                ClasserJoueurs( _
                    StatDuPalmares(p), Tailles(p), _
                    NomsJoueurs, NbJoueurs, Compteurs, _
                    Joues, MinutesTotal, Retenu, Mode, _
                    Rangs, Noms, Valeurs _
                )

            Hauteur = _
                EcrireBloc( _
                    ws, LigneRangee, Colonne, _
                    Libelles(p), StatDuPalmares(p) > 0, _
                    CouleurEntete, FormatNombre, _
                    Rangs, Noms, Valeurs, NbLignes _
                )

            HauteurRangee = Maxi(HauteurRangee, Hauteur)

            NbBlocs = NbBlocs + 1

        End If

    Next p

End Sub


' =========================================================
' LE CLASSEMENT D'UNE STATISTIQUE
'
' Les joueurs sont tries par valeur decroissante, puis par
' nom a egalite. Le rang d'un joueur est un de plus que le
' nombre de joueurs qui font strictement mieux : les ex
' aequo le partagent. On s'arrete au premier rang qui
' depasse le nombre de places.
'
' Les valeurs sont arrondies a deux decimales avant d'etre
' comparees : en moyenne par match, deux joueurs affiches a
' 1,50 doivent etre ex aequo, meme si le calcul les separe
' au millionieme.
'
' Un joueur a zero n'est pas classe : un top des
' franchissements n'a pas a lister ceux qui n'en ont fait
' aucun.
' =========================================================
Private Function ClasserJoueurs( _
    ByVal Stat As Long, _
    ByVal Taille As Long, _
    ByRef NomsJoueurs() As String, _
    ByVal NbJoueurs As Long, _
    ByRef Compteurs() As Double, _
    ByRef Joues() As Long, _
    ByRef MinutesTotal() As Double, _
    ByRef Retenu() As Boolean, _
    ByVal Mode As String, _
    ByRef Rangs() As Long, _
    ByRef Noms() As String, _
    ByRef Valeurs() As Double _
) As Long

    Dim Valeur() As Double
    Dim Candidats() As Long
    Dim NbCandidats As Long
    Dim Affichee As Variant
    Dim Courant As Long
    Dim Rang As Long
    Dim NbLignes As Long
    Dim i As Long
    Dim j As Long

    ReDim Rangs(1 To 1)
    ReDim Noms(1 To 1)
    ReDim Valeurs(1 To 1)

    If Stat <= 0 Or NbJoueurs = 0 Then
        Exit Function
    End If

    ReDim Valeur(1 To NbJoueurs)
    ReDim Candidats(1 To NbJoueurs)

    For i = 1 To NbJoueurs

        If Retenu(i) Then

            Affichee = _
                ValeurAffichee( _
                    Compteurs(i, Stat), Joues(i), _
                    MinutesTotal(i), Mode _
                )

            If IsNumeric(Affichee) And Not IsEmpty(Affichee) Then

                Valeur(i) = Round(CDbl(Affichee), 2)

                If Valeur(i) > 0 Then
                    NbCandidats = NbCandidats + 1
                    Candidats(NbCandidats) = i
                End If

            End If

        End If

    Next i

    If NbCandidats = 0 Then
        Exit Function
    End If

    ' Tri par insertion : une quarantaine de joueurs au
    ' plus, inutile de sortir l'artillerie.
    For i = 2 To NbCandidats

        Courant = Candidats(i)
        j = i - 1

        Do While j >= 1

            If Not PasseDevant( _
                Courant, Candidats(j), Valeur, NomsJoueurs _
            ) Then
                Exit Do
            End If

            Candidats(j + 1) = Candidats(j)
            j = j - 1

        Loop

        Candidats(j + 1) = Courant

    Next i

    ReDim Rangs(1 To NbCandidats)
    ReDim Noms(1 To NbCandidats)
    ReDim Valeurs(1 To NbCandidats)

    For i = 1 To NbCandidats

        If i = 1 Then
            Rang = 1
        ElseIf Valeur(Candidats(i)) < Valeur(Candidats(i - 1)) Then
            Rang = i
        End If

        If Rang > Taille Then
            Exit For
        End If

        NbLignes = NbLignes + 1

        Rangs(NbLignes) = Rang
        Noms(NbLignes) = NomsJoueurs(Candidats(i))
        Valeurs(NbLignes) = Valeur(Candidats(i))

    Next i

    ClasserJoueurs = NbLignes

End Function


Private Function PasseDevant( _
    ByVal Joueur As Long, _
    ByVal Autre As Long, _
    ByRef Valeur() As Double, _
    ByRef NomsJoueurs() As String _
) As Boolean

    If Valeur(Joueur) <> Valeur(Autre) Then
        PasseDevant = (Valeur(Joueur) > Valeur(Autre))
    Else
        PasseDevant = _
            (StrComp(NomsJoueurs(Joueur), NomsJoueurs(Autre), vbTextCompare) < 0)
    End If

End Function


' =========================================================
' Pose un classement et renvoie le nombre de lignes qu'il
' occupe, en-tete compris.
'
' Une plage mise en forme et non un tableau structure : un
' tri par clic ruinerait la colonne des rangs.
' =========================================================
Private Function EcrireBloc( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long, _
    ByVal Libelle As String, _
    ByVal StatConnue As Boolean, _
    ByVal CouleurEntete As Long, _
    ByVal FormatNombre As String, _
    ByRef Rangs() As Long, _
    ByRef Noms() As String, _
    ByRef Valeurs() As Double, _
    ByVal NbLignes As Long _
) As Long

    Dim Sortie() As Variant
    Dim Corps As Range
    Dim r As Long

    With ws.Cells(Ligne, Colonne).Resize(1, 3)
        .Value = Array("Rang", "Joueur", Libelle)
        .Font.Bold = True
        .Font.Color = BLANC
        .Interior.Color = CouleurEntete
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With

    ws.Rows(Ligne).RowHeight = 32

    If NbLignes = 0 Then

        With ws.Cells(Ligne + 1, Colonne + 1)

            If StatConnue Then
                .Value = "Aucune donn" & ChrW(233) & "e"
            Else
                .Value = "Statistique introuvable dans Parametres"
            End If

            .Font.Italic = True
            .Font.Color = RGB(110, 110, 110)

        End With

        EcrireBloc = 2

        Exit Function

    End If

    ReDim Sortie(1 To NbLignes, 1 To 3)

    For r = 1 To NbLignes
        Sortie(r, 1) = Rangs(r)
        Sortie(r, 2) = Noms(r)
        Sortie(r, 3) = Valeurs(r)
    Next r

    Set Corps = ws.Cells(Ligne + 1, Colonne).Resize(NbLignes, 3)

    Corps.Value = Sortie

    ' Le remplissage peut etre pose a la main : rien ne
    ' trie ces lignes, les bandes ne se decaleront pas.
    For r = 1 To NbLignes

        If (r Mod 2) = 1 Then
            Corps.Rows(r).Interior.Color = VERT_PALE
        Else
            Corps.Rows(r).Interior.Color = BLANC
        End If

    Next r

    Corps.Columns(1).HorizontalAlignment = xlCenter
    Corps.Columns(3).HorizontalAlignment = xlCenter
    Corps.Columns(3).NumberFormat = FormatNombre

    With ws.Cells(Ligne, Colonne).Resize(NbLignes + 1, 3)
        .Borders(xlEdgeLeft).Color = CouleurEntete
        .Borders(xlEdgeRight).Color = CouleurEntete
        .Borders(xlEdgeBottom).Color = CouleurEntete
    End With

    EcrireBloc = NbLignes + 1

End Function


' =========================================================
' MISE EN PLACE DES FEUILLES
' =========================================================

' La feuille est creee au premier calcul, rangee derriere
' celle qui la precede dans l'ordre de lecture.
Private Function FeuillePalmares( _
    ByVal NomFeuille As String, _
    ByVal FeuillePrecedente As String, _
    ByVal Titre As String, _
    ByVal SousTitre As String _
) As Worksheet

    Dim ws As Worksheet
    Dim Precedente As Worksheet

    Set ws = FeuilleRecap(NomFeuille)

    If ws Is Nothing Then

        Set ws = FeuilleRecapOuCreee(NomFeuille)

        Set Precedente = FeuilleRecap(FeuillePrecedente)

        If Not Precedente Is Nothing Then
            ws.Move After:=Precedente
        End If

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
        ws.Rows(LIGNE_ENTETE - 2).RowHeight = 34

        PoserBouton _
            ws, "Recalculer", "AppliquerFiltres", _
            LIGNE_ENTETE - 2, PREMIERE_COLONNE, 130

    End If

    Set FeuillePalmares = ws

End Function


' Tout ce qui suit le titre et le bouton : le bandeau et
' les classements du calcul precedent.
Private Sub ViderFeuillePalmares( _
    ByVal ws As Worksheet _
)

    Dim Derniere As Long
    Dim DerniereColonne As Long

    With ws.UsedRange
        Derniere = .Row + .Rows.Count
        DerniereColonne = .Column + .Columns.Count
    End With

    If Derniere < LIGNE_ENTETE - 1 Then
        Exit Sub
    End If

    With ws.Range( _
        ws.Cells(LIGNE_ENTETE - 1, 1), _
        ws.Cells(Derniere, Maxi(DerniereColonne, 2)) _
    )
        .Clear
        .EntireRow.RowHeight = ws.StandardHeight
    End With

End Sub


' Le bandeau des filtres est celui que le classement vient
' d'ecrire : les deux feuilles disent la meme chose.
Private Sub RecopierBandeau( _
    ByVal ws As Worksheet _
)

    Dim Classement As Worksheet

    Set Classement = FeuilleRecap(FEUILLE_CLASSEMENT)

    If Classement Is Nothing Then
        Exit Sub
    End If

    With ws.Cells(LIGNE_ENTETE - 1, PREMIERE_COLONNE)
        .Value = Classement.Cells(LIGNE_ENTETE - 1, PREMIERE_COLONNE).Value
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(150, 30, 30)
        .VerticalAlignment = xlBottom
    End With

End Sub


' =========================================================
' LE TABLEAU DES CLASSEMENTS, DANS PARAMETRES
'
' Cree une seule fois : ensuite, c'est l'entraineur qui
' ajoute, retire ou reordonne les lignes. L'ordre des lignes
' est l'ordre des blocs dans la feuille.
' =========================================================
Private Sub PreparerTableauPalmares()

    Dim ws As Worksheet
    Dim Tableau As ListObject
    Dim Lignes() As Variant
    Dim NbLignes As Long

    Set ws = FeuilleRecap(FEUILLE_PARAMETRES)

    If ws Is Nothing Then
        Exit Sub
    End If

    If Not TableauRecap(FEUILLE_PARAMETRES, TBL_PALMARES) Is Nothing Then
        Exit Sub
    End If

    ReDim Lignes(1 To NB_PALMARES_DEFAUT, 1 To 3)

    ' Les classements du bilan des blocs 1 et 2.
    AjouterPalmares Lignes, NbLignes, "Top", "Franchissements"
    AjouterPalmares Lignes, NbLignes, "Top", "Plaquages r" & ChrW(233) & "ussis"
    AjouterPalmares Lignes, NbLignes, "Top", "Plaquages offensifs"
    AjouterPalmares Lignes, NbLignes, "Top", "Touches conserv" & ChrW(233) & "es"
    AjouterPalmares Lignes, NbLignes, "Flop", "Plaquages rat" & ChrW(233) & "s"
    AjouterPalmares Lignes, NbLignes, "Flop", "Plaquages hauts"
    AjouterPalmares Lignes, NbLignes, "Flop", "En-avants"
    AjouterPalmares Lignes, NbLignes, "Flop", _
        "P" & ChrW(233) & "nalit" & ChrW(233) & "s conc" & ChrW(233) & "d" & ChrW(233) & "es"
    AjouterPalmares Lignes, NbLignes, "Flop", "Plaquages " & ChrW(224) & " 2"

    With ws.Cells(LIGNE_ENTETE - 1, COLONNE_PALMARES)
        .Value = "Tops et flops"
        .Font.Bold = True
    End With

    ws.Cells(LIGNE_ENTETE, COLONNE_PALMARES).Resize(1, 3).Value = _
        Array("Sens", "Statistique", "Places")

    ws.Cells(LIGNE_ENTETE + 1, COLONNE_PALMARES) _
        .Resize(NbLignes, 3).Value = Lignes

    Set Tableau = _
        ws.ListObjects.Add( _
            xlSrcRange, _
            ws.Cells(LIGNE_ENTETE, COLONNE_PALMARES).Resize(NbLignes + 1, 3), _
            , _
            xlYes _
        )

    Tableau.Name = TBL_PALMARES
    Tableau.TableStyle = "TableStyleMedium3"
    Tableau.HeaderRowRange.Font.Bold = True

    ws.Columns(COLONNE_PALMARES).ColumnWidth = 9
    ws.Columns(COLONNE_PALMARES + 1).ColumnWidth = 26
    ws.Columns(COLONNE_PALMARES + 2).ColumnWidth = 8

    ' Un menu sur la colonne Statistique, alimente par les
    ' libelles du tableau de gauche : c'est sur eux que se
    ' fait la correspondance.
    On Error Resume Next

    With Tableau.ListColumns(2).DataBodyRange.Validation
        .Delete
        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertWarning, _
            Operator:=xlBetween, _
            Formula1:="=INDIRECT(""" & TBL_STATS & "[Colonne]"")"
    End With

    On Error GoTo 0

End Sub


Private Sub AjouterPalmares( _
    ByRef Lignes() As Variant, _
    ByRef NbLignes As Long, _
    ByVal Sens As String, _
    ByVal Statistique As String _
)

    If NbLignes >= NB_PALMARES_DEFAUT Then
        Exit Sub
    End If

    NbLignes = NbLignes + 1

    Lignes(NbLignes, 1) = Sens
    Lignes(NbLignes, 2) = Statistique
    Lignes(NbLignes, 3) = TAILLE_PAR_DEFAUT

End Sub


' =========================================================
' FEUILLE PENALITES
'
' Un tableau par categorie, cote a cote : la Premiere et la
' Reserve ne commettent ni les memes fautes ni autant, et
' la part de chaque motif se lit sur le total de sa propre
' equipe.
'
' Deux libelles qui ne different que par la ponctuation,
' "Plaquage a 2" et "Plaquage - a 2", sont un seul motif :
' la saisie a change d'ecriture en cours de saison. C'est
' la forme avec tiret, celle de la saisie actuelle, qui est
' affichee.
'
' Les penalites sans motif ne sont pas classees : elles
' ferment le tableau, avant la ligne de total, et comptent
' dans le total sur lequel se calculent les parts.
'
' Le mode d'affichage et le seuil de minutes ne s'y
' appliquent pas : ce sont des reglages de classement
' individuel.
' =========================================================
Public Function FeuillePenalites() As String

    FeuillePenalites = "P" & ChrW(233) & "nalit" & ChrW(233) & "s"

End Function


Public Sub EcrirePenalites( _
    ByVal Matchs As Collection, _
    ByVal Categories As Collection, _
    ByRef NomsCategories() As String, _
    ByVal NbCategories As Long _
)

    Dim ws As Worksheet
    Dim Actions As Variant

    Dim Motifs As Collection
    Dim Libelles() As String
    Dim Comptes() As Double
    Dim NbMotifs As Long

    Dim SansMotif() As Double
    Dim Capacite As Long
    Dim Ligne As Long
    Dim Categorie As Long
    Dim CleCategorie As String
    Dim Motif As String
    Dim Colonne As Long
    Dim c As Long

    Set ws = _
        FeuillePalmares( _
            FeuillePenalites(), FEUILLE_FLOPS, _
            "P" & ChrW(233) & "nalit" & ChrW(233) & "s contre nous", _
            "Class" & ChrW(233) & "es par motif, une " & ChrW(233) & _
            "quipe par tableau, avec les filtres du classement. " & _
            "La part se lit sur le total de l'" & ChrW(233) & "quipe. " & _
            "Les ex aequo partagent le m" & ChrW(234) & "me rang." _
        )

    ViderFeuillePalmares ws

    RecopierBandeau ws

    Set Motifs = New Collection

    ReDim SansMotif(1 To NbCategories)

    Actions = DonneesTableau(FEUILLE_JOURNAL, TBL_JOURNAL)

    ' Une premiere passe compte les penalites : il ne peut
    ' y avoir plus de motifs differents que de penalites.
    If Not IsEmpty(Actions) Then

        For Ligne = 1 To UBound(Actions, 1)

            If EstPenaliteRetenue(Matchs, Actions, Ligne) Then
                Capacite = Capacite + 1
            End If

        Next Ligne

    End If

    If Capacite = 0 Then

        With ws.Cells(LIGNE_ENTETE + 1, PREMIERE_COLONNE)
            .Value = "Aucune p" & ChrW(233) & "nalit" & ChrW(233) & _
                " contre nous dans les matchs retenus."
            .Font.Italic = True
            .Font.Color = RGB(110, 110, 110)
        End With

        Exit Sub

    End If

    ' La categorie en premier indice : c'est le dernier,
    ' le motif, qui depend du contenu.
    ReDim Libelles(1 To Capacite)
    ReDim Comptes(1 To NbCategories, 1 To Capacite)

    For Ligne = 1 To UBound(Actions, 1)

        If EstPenaliteRetenue(Matchs, Actions, Ligne) Then

            CleCategorie = _
                Normaliser(TexteCellule(Actions(Ligne, JRN_CATEGORIE)))

            ' Un match sans categorie reconnue n'a pas de
            ' tableau ou se ranger.
            If CleExiste(Categories, CleCategorie) Then

                Categorie = CLng(Categories.Item(CleCategorie))

                Motif = TexteCellule(Actions(Ligne, JRN_MOTIF))

                ' Normaliser ne garde que lettres et
                ' chiffres : le point d'interrogation devient
                ' vide, comme une case laissee blanche.
                If Len(Normaliser(Motif)) = 0 Then

                    SansMotif(Categorie) = SansMotif(Categorie) + 1

                Else

                    CompterMotif _
                        Motif, Categorie, Motifs, _
                        Libelles, Comptes, NbMotifs

                End If

            End If

        End If

    Next Ligne

    Colonne = PREMIERE_COLONNE

    For c = 1 To NbCategories

        Colonne = _
            Colonne + 1 + _
            EcrireTableauPenalites( _
                ws, Colonne, c, NomsCategories(c), _
                Libelles, Comptes, NbMotifs, SansMotif(c) _
            )

    Next c

End Sub


' Vrai pour une penalite contre nous d'un match retenu.
' Le fichier precedent est garde en memoire : les lignes
' d'un meme match se suivent dans la base.
Private Function EstPenaliteRetenue( _
    ByVal Matchs As Collection, _
    ByRef Actions As Variant, _
    ByVal Ligne As Long _
) As Boolean

    Static DerniereListe As Collection
    Static DerniereCle As String
    Static DerniereReponse As Boolean

    Dim Cle As String

    If Normaliser(TexteCellule(Actions(Ligne, JRN_ACTION))) <> ACTION_PENALITE Then
        Exit Function
    End If

    Cle = Normaliser(TexteCellule(Actions(Ligne, JRN_FICHIER)))

    If Not DerniereListe Is Matchs Or Cle <> DerniereCle Then
        Set DerniereListe = Matchs
        DerniereCle = Cle
        DerniereReponse = CleExiste(Matchs, Cle)
    End If

    EstPenaliteRetenue = DerniereReponse

End Function


' Ajoute une penalite au motif, en le creant au besoin.
Private Sub CompterMotif( _
    ByVal Libelle As String, _
    ByVal Categorie As Long, _
    ByVal Index As Collection, _
    ByRef Libelles() As String, _
    ByRef Comptes() As Double, _
    ByRef NbLibelles As Long _
)

    Dim Cle As String
    Dim i As Long

    Cle = Normaliser(Libelle)

    If CleExiste(Index, Cle) Then

        i = CLng(Index.Item(Cle))

        If InStr(1, Libelles(i), " - ") = 0 And InStr(1, Libelle, " - ") > 0 Then
            Libelles(i) = Libelle
        End If

    Else

        NbLibelles = NbLibelles + 1
        i = NbLibelles

        Index.Add i, Cle
        Libelles(i) = Libelle

    End If

    Comptes(Categorie, i) = Comptes(Categorie, i) + 1

End Sub


Private Function CouleurCategorie( _
    ByVal NomCategorie As String _
) As Long

    Select Case Normaliser(NomCategorie)
        Case "premiere"
            CouleurCategorie = TETE_PREMIERE
        Case "reserve"
            CouleurCategorie = TETE_RESERVE
        Case Else
            CouleurCategorie = ROUGE_ENTETE
    End Select

End Function


' =========================================================
' Pose le classement des motifs d'une categorie et renvoie
' sa largeur en colonnes : rang, motif, nombre, part.
'
' Une ligne de titre au-dessus porte le nom de l'equipe.
' Seuls les motifs que cette equipe a commis y figurent.
' =========================================================
Private Function EcrireTableauPenalites( _
    ByVal ws As Worksheet, _
    ByVal Colonne As Long, _
    ByVal Categorie As Long, _
    ByVal NomCategorie As String, _
    ByRef Libelles() As String, _
    ByRef Comptes() As Double, _
    ByVal NbMotifs As Long, _
    ByVal SansMotif As Double _
) As Long

    Const NB_COLONNES As Long = 4

    Dim Ordre() As Long
    Dim NbRetenus As Long
    Dim Sortie() As Variant
    Dim Total As Double
    Dim Couleur As Long
    Dim NbLignes As Long
    Dim Ligne As Long
    Dim Courant As Long
    Dim Rang As Long
    Dim Corps As Range
    Dim i As Long
    Dim j As Long

    EcrireTableauPenalites = NB_COLONNES

    Ligne = LIGNE_ENTETE + 1
    Couleur = CouleurCategorie(NomCategorie)

    ws.Columns(Colonne).ColumnWidth = 6
    ws.Columns(Colonne + 1).ColumnWidth = 34
    ws.Columns(Colonne + 2).ColumnWidth = 11
    ws.Columns(Colonne + 3).ColumnWidth = 9
    ws.Columns(Colonne + NB_COLONNES).ColumnWidth = 3

    With ws.Cells(Ligne, Colonne)
        .Value = NomCategorie
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = Couleur
    End With

    Ligne = Ligne + 1

    ' Les motifs de l'equipe, par nombre decroissant puis
    ' par ordre alphabetique.
    ReDim Ordre(1 To Maxi(NbMotifs, 1))

    For i = 1 To NbMotifs

        If Comptes(Categorie, i) > 0 Then

            Total = Total + Comptes(Categorie, i)

            Courant = i
            j = NbRetenus

            Do While j >= 1

                If Comptes(Categorie, Courant) < Comptes(Categorie, Ordre(j)) Then Exit Do

                If Comptes(Categorie, Courant) = Comptes(Categorie, Ordre(j)) Then
                    If StrComp(Libelles(Courant), Libelles(Ordre(j)), vbTextCompare) >= 0 Then Exit Do
                End If

                Ordre(j + 1) = Ordre(j)
                j = j - 1

            Loop

            Ordre(j + 1) = Courant
            NbRetenus = NbRetenus + 1

        End If

    Next i

    Total = Total + SansMotif

    With ws.Cells(Ligne, Colonne).Resize(1, NB_COLONNES)
        .Value = Array("Rang", "Motif", "P" & ChrW(233) & "nalit" & ChrW(233) & "s", "Part")
        .Font.Bold = True
        .Font.Color = BLANC
        .Interior.Color = Couleur
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With

    ws.Rows(Ligne).RowHeight = 24

    If Total = 0 Then

        With ws.Cells(Ligne + 1, Colonne + 1)
            .Value = "Aucune p" & ChrW(233) & "nalit" & ChrW(233)
            .Font.Italic = True
            .Font.Color = RGB(110, 110, 110)
        End With

        Exit Function

    End If

    ' Les motifs, puis la ligne sans motif s'il y en a,
    ' puis le total.
    NbLignes = NbRetenus + 1

    If SansMotif > 0 Then
        NbLignes = NbLignes + 1
    End If

    ReDim Sortie(1 To NbLignes, 1 To NB_COLONNES)

    For i = 1 To NbRetenus

        If i = 1 Then
            Rang = 1
        ElseIf Comptes(Categorie, Ordre(i)) < Comptes(Categorie, Ordre(i - 1)) Then
            Rang = i
        End If

        Sortie(i, 1) = Rang
        Sortie(i, 2) = Libelles(Ordre(i))
        Sortie(i, 3) = Comptes(Categorie, Ordre(i))
        Sortie(i, 4) = Comptes(Categorie, Ordre(i)) / Total

    Next i

    If SansMotif > 0 Then
        Sortie(NbRetenus + 1, 2) = "Motif non renseign" & ChrW(233)
        Sortie(NbRetenus + 1, 3) = SansMotif
        Sortie(NbRetenus + 1, 4) = SansMotif / Total
    End If

    Sortie(NbLignes, 2) = "Total"
    Sortie(NbLignes, 3) = Total
    Sortie(NbLignes, 4) = 1

    Set Corps = ws.Cells(Ligne + 1, Colonne).Resize(NbLignes, NB_COLONNES)

    Corps.Value = Sortie

    For i = 1 To NbLignes - 1

        If (i Mod 2) = 1 Then
            Corps.Rows(i).Interior.Color = VERT_PALE
        Else
            Corps.Rows(i).Interior.Color = BLANC
        End If

    Next i

    If SansMotif > 0 Then
        Corps.Rows(NbLignes - 1).Font.Italic = True
    End If

    With Corps.Rows(NbLignes)
        .Font.Bold = True
        .Borders(xlEdgeTop).Color = Couleur
    End With

    Corps.Columns(1).HorizontalAlignment = xlCenter
    Corps.Columns(3).HorizontalAlignment = xlCenter
    Corps.Columns(4).HorizontalAlignment = xlCenter
    Corps.Columns(3).NumberFormat = "0"
    Corps.Columns(4).NumberFormat = "0%"

    With ws.Cells(Ligne, Colonne).Resize(NbLignes + 1, NB_COLONNES)
        .Borders(xlEdgeLeft).Color = Couleur
        .Borders(xlEdgeRight).Color = Couleur
        .Borders(xlEdgeBottom).Color = Couleur
    End With

End Function
