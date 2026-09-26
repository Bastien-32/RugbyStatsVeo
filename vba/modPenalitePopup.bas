Attribute VB_Name = "modPenalitePopup"
Option Explicit

' =========================================================
' PENALITES - POPUP DE SAISIE
'
' Au clic sur "penalite contre nous" ou "contre
' adversaire", une popup s'ouvre avec les motifs a gauche
' et la composition a droite. On choisit l'un, l'autre,
' les deux ou rien, puis on valide.
'
' Ce qui n'est pas renseigne vaut "?", comme le faisait
' l'attente de la palette. Le groupe fautif n'est pas
' demande : GetPlayerGroup le deduit du joueur.
'
' La disposition reprend la maquette dessinee dans le
' classeur : motifs groupes par famille, joueurs places
' comme sur le terrain, l'attaque vers le bas.
'
' Ce que la case affiche est le nom du joueur, mais c'est
' sa POSITION qui l'identifie : un nom ne permettrait pas
' de remonter au poste, et la compo change d'un match a
' l'autre.
'
' Aucun caractere accentue en clair : le VBE de macOS
' importe les .bas en Mac Roman. Les accents passent donc
' par ChrW, comme ailleurs dans ce classeur.
' =========================================================

Public Const FEUILLE_POPUP_PEN As String = "Popup penalite"

' Colonnes des trois familles de motifs, fusionnees par
' deux : B:C, E:F, H:I.
Private Const COL_MOTIF_1 As Long = 2
Private Const COL_MOTIF_2 As Long = 5
Private Const COL_MOTIF_3 As Long = 8
Private Const LARGEUR_MOTIF As Long = 2

' Les valeurs du contexte se collent a leurs libelles.
Private Const COL_CONTEXTE As Long = 3

' Les boutons joueurs tiennent sur trois colonnes.
Private Const LARGEUR_JOUEUR As Long = 3

Private Const PREMIERE_LIGNE As Long = 7
Private Const LIGNE_BOUTONS As Long = 29

Private Const GRIS As Long = 15132390     ' RGB(230,230,230)
Private Const BLEU As Long = 12611584     ' RGB(0,112,192)
Private Const VERT As Long = 32768        ' RGB(0,128,0)
Private Const ROUGE As Long = 192         ' RGB(192,0,0)
Private Const SABLE As Long = 14083324    ' RGB(252,228,214)
Private Const CIEL As Long = 15983321     ' RGB(217,226,243)

' Contexte de la penalite en cours de saisie.
Private PenaliteAction As String
Private PenaliteTemps As Double
Private PenaliteEquipe As String
Private PenaliteOuverte As Boolean


' ---------------------------------------------------------
' Disposition
'
' Chaque entree vaut "ligne:colonne:contenu". Lignes et
' colonnes sont celles de la feuille, pour que la maquette
' reste lisible telle qu'elle a ete dessinee.
' ---------------------------------------------------------

Private Function MotifsDisposes() As Variant

    MotifsDisposes = Array( _
        "11:2:Maul", _
        "11:5:Ruck", _
        "11:8:Plaquage " & ChrW(224) & " 2", _
        "13:2:Entrer sur le c" & ChrW(244) & "t" & _
            ChrW(233) & " maul", _
        "13:5:Talonnage " & ChrW(224) & " la main", _
        "13:8:Plaquage haut", _
        "15:2:Maul " & ChrW(233) & "croul" & ChrW(233), _
        "15:5:Retard soutient", _
        "15:8:Plaquage sans ballon", _
        "17:2:Melee poussee avant introduction", _
        "17:5:Soutient va au-del" & ChrW(224), _
        "17:8:Plaquage " & ChrW(224) & " retardement", _
        "19:5:Soutient couch" & ChrW(233) & " sur porteur", _
        "21:2:En-avant volontaire", _
        "21:5:4 appuis", _
        "23:2:Parle arbitre", _
        "23:5:Plaqueur qui ne sort pas", _
        "23:8:Hors-jeu", _
        "25:2:Brutalit" & ChrW(233), _
        "25:5:Garde le ballon au sol", _
        "25:8:Hors-jeu - d" & ChrW(233) & "part devant botteur")

End Function


' L'equipe vue de dessus, l'attaque vers le bas.
Private Function JoueursDisposes() As Variant

    JoueursDisposes = Array( _
        "7:11:1", "7:15:2", "7:19:3", _
        "9:13:4", "9:17:5", _
        "11:11:6", "11:15:8", "11:19:7", _
        "13:13:9", _
        "15:16:10", _
        "17:13:12", "17:17:13", _
        "19:11:11", "19:15:15", "19:19:14", _
        "21:11:16", "21:15:17", "21:19:18", _
        "23:11:19", "23:15:20", "23:19:21", _
        "25:15:22", _
        "27:11:Collectif", "27:19:?")

End Function


' Element d'une entree : 0 la ligne, 1 la colonne, 2 le
' contenu. Le contenu peut lui-meme contenir un deux-
' points, d'ou le recollage.
Private Function Champ( _
    ByVal Entree As String, _
    ByVal Rang As Long) As String

    Dim Morceaux() As String
    Dim i As Long

    Morceaux = Split(Entree, ":")

    If Rang > UBound(Morceaux) Then Exit Function

    If Rang < 2 Then
        Champ = Morceaux(Rang)
        Exit Function
    End If

    Champ = Morceaux(2)

    For i = 3 To UBound(Morceaux)
        Champ = Champ & ":" & Morceaux(i)
    Next i

End Function


' ---------------------------------------------------------
' Construction
' ---------------------------------------------------------

Public Sub ConstruirePopupPenalite()

    Dim ws As Worksheet

    On Error GoTo GestionErreur

    Set ws = FeuillePopup

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws
    EcrireContexte ws
    EcrireMotifs ws
    EcrireJoueurs ws
    EcrireBoutons ws
    MettreEnForme ws

    MsgBox _
        "La popup des penalites est prete.", _
        vbInformation, _
        "Penalites"

    Exit Sub

GestionErreur:

    MsgBox _
        "La construction a echoue." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Penalites"

End Sub


Public Sub SupprimerPopupPenalite()

    Dim ws As Worksheet
    Dim EtatAlertes As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    EtatAlertes = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ws.Visible = xlSheetVisible
    ws.Delete

    Application.DisplayAlerts = EtatAlertes

End Sub


Private Sub EcrireTitre(ByVal ws As Worksheet)

    With ws.Cells(2, COL_MOTIF_1)
        .Value = "PENALITE"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = BLEU
    End With

End Sub


Private Sub EcrireContexte(ByVal ws As Worksheet)

    Dim Libelles As Variant
    Dim i As Long

    Libelles = Array("Temps video", "Mi-temps", "Contre")

    For i = 0 To 2

        With ws.Cells(4 + i, COL_MOTIF_1)
            .Value = Libelles(i)
            .Font.Italic = True
        End With

        With ws.Cells(4 + i, COL_CONTEXTE).Resize(1, 2)
            .Merge
            .Interior.Color = GRIS
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(4, COL_CONTEXTE), _
            ws.Cells(6, COL_CONTEXTE)).Address, _
        "PEN_CONTEXTE"

End Sub


Private Sub EcrireMotifs(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    Entrees = MotifsDisposes

    For i = 0 To UBound(Entrees)

        With ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))) _
            .Resize(1, LARGEUR_MOTIF)

            .Merge
            .Value = Champ(CStr(Entrees(i)), 2)
            .Interior.Color = SABLE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)

        End With

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(PREMIERE_LIGNE, COL_MOTIF_1), _
            ws.Cells(LIGNE_BOUTONS - 1, _
                COL_MOTIF_3 + LARGEUR_MOTIF - 1)).Address, _
        "PEN_MOTIFS"

End Sub


Private Sub EcrireJoueurs(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        With ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))) _
            .Resize(1, LARGEUR_JOUEUR)

            .Merge
            .Interior.Color = CIEL
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)

        End With

        EcrireCaseJoueur _
            ws.Cells( _
                CLng(Champ(CStr(Entrees(i)), 0)), _
                CLng(Champ(CStr(Entrees(i)), 1))), _
            Champ(CStr(Entrees(i)), 2)

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(PREMIERE_LIGNE, 11), _
            ws.Cells(LIGNE_BOUTONS - 1, 21)).Address, _
        "PEN_JOUEURS"

End Sub


Private Sub EcrireBoutons(ByVal ws As Worksheet)

    With ws.Cells(LIGNE_BOUTONS, COL_MOTIF_1).Resize(1, 2)
        .Merge
        .Value = "ANNULER"
        .Interior.Color = ROUGE
    End With

    With ws.Cells(LIGNE_BOUTONS, COL_MOTIF_2).Resize(1, 4)
        .Merge
        .Value = "VALIDER  (Entree)"
        .Interior.Color = VERT
    End With

    With ws.Range( _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_1), _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_3 + 1))

        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ws.Rows(LIGNE_BOUTONS).RowHeight = 26

    NommerPlage ws, _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_1).Address, _
        "PEN_ANNULER"

    NommerPlage ws, _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_2).Address, _
        "PEN_VALIDER"

End Sub


Private Sub MettreEnForme(ByVal ws As Worksheet)

    ' Largeurs de la maquette : une colonne etroite separe
    ' les familles de motifs, et la composition tient sur
    ' des colonnes fines.
    ws.Columns("B:C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 3.5
    ws.Columns("E:F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 4
    ws.Columns("H:I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 4
    ws.Columns("K:U").ColumnWidth = 4.3

    ws.Cells.Font.Name = "Calibri"

    AjusterHauteurs ws

End Sub


' Les rangs qui portent quelque chose font 20, les rangs
' intercalaires 10 : la grille respire sans s'etaler.
Private Sub AjusterHauteurs(ByVal ws As Worksheet)

    Dim Ligne As Long

    For Ligne = PREMIERE_LIGNE To LIGNE_BOUTONS - 2

        If Application.WorksheetFunction.CountA( _
            ws.Range( _
                ws.Cells(Ligne, 2), _
                ws.Cells(Ligne, 21))) > 0 Then

            ws.Rows(Ligne).RowHeight = 20

        Else

            ws.Rows(Ligne).RowHeight = 10

        End If

    Next Ligne

End Sub


' ---------------------------------------------------------
' Libelles
' ---------------------------------------------------------

' Nom de famille du joueur, ou le numero si la compo ne le
' donne pas. Meme regle que la palette de saisie : les
' mots en majuscules forment le nom de famille.
Private Function LibelleJoueur( _
    ByVal Poste As String) As String

    Dim Complet As String
    Dim Mots() As String
    Dim Mot As Variant
    Dim Famille As String

    LibelleJoueur = Poste

    If Not IsNumeric(Poste) Then Exit Function

    Complet = Trim(GetPlayerName(Poste))

    If Complet = "" Then Exit Function

    Mots = Split(Complet, " ")

    For Each Mot In Mots

        If CStr(Mot) = UCase(CStr(Mot)) Then

            If Famille = "" Then
                Famille = CStr(Mot)
            Else
                Famille = Famille & " " & CStr(Mot)
            End If

        Else

            Exit For

        End If

    Next Mot

    If Famille <> "" Then LibelleJoueur = Famille

End Function


' Le numero en gras, le nom dessous : meme presentation
' que les boutons de la palette de saisie.
Private Sub EcrireCaseJoueur( _
    ByVal Cellule As Range, _
    ByVal Poste As String)

    Dim Nom As String

    Nom = LibelleJoueur(Poste)

    With Cellule

        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        If Nom = Poste Then

            ' Sans compo, ou pour Collectif et "?", il n'y
            ' a que l'etiquette.
            .Value = Poste
            .Font.Size = 14
            .Font.Bold = True

            Exit Sub

        End If

        .Value = Poste & vbLf & Nom

        .Characters(1, Len(Poste)).Font.Size = 14
        .Characters(1, Len(Poste)).Font.Bold = True

        .Characters(Len(Poste) + 2, Len(Nom)).Font.Size = 9
        .Characters(Len(Poste) + 2, Len(Nom)).Font.Bold = False

    End With

End Sub


' Les libelles suivent la compo, qui change d'un match a
' l'autre : ils sont refaits a chaque ouverture.
Private Sub RafraichirLibelles(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    Application.EnableEvents = False

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        EcrireCaseJoueur _
            ws.Cells( _
                CLng(Champ(CStr(Entrees(i)), 0)), _
                CLng(Champ(CStr(Entrees(i)), 1))), _
            Champ(CStr(Entrees(i)), 2)

    Next i

    Application.EnableEvents = True

End Sub


' ---------------------------------------------------------
' Ouverture et fermeture
' ---------------------------------------------------------

Public Sub OuvrirPopupPenalite( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal Equipe As String)

    Dim ws As Worksheet

    PenaliteAction = ActionTexte
    PenaliteTemps = TempsVideo
    PenaliteEquipe = Equipe
    PenaliteOuverte = True

    ' Une action restee en attente doit etre close avant
    ' que la penalite prenne la main, comme le faisait
    ' InitialiserPenalite.
    FermerActionEnAttente

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    ws.Visible = xlSheetVisible
    ws.Activate

    EffacerSelections ws
    RafraichirLibelles ws

    With ws.Range("PEN_CONTEXTE")
        .Cells(1, 1).Value = FormaterTemps(TempsVideo)
        .Cells(2, 1).Value = CurrentHalf
        .Cells(3, 1).Value = _
            IIf(Equipe = "Nous", "Nous", "Adversaire")
    End With

    ' La touche Entree vaut validation, comme le fait deja
    ' la popup d'ajout de joueur pour son bouton.
    Application.OnKey "~", "ValiderPopupPenalite"
    Application.OnKey "{ENTER}", "ValiderPopupPenalite"

End Sub


Public Sub ValiderPopupPenalite()

    Dim Motif As String
    Dim Joueur As String

    If Not PenaliteOuverte Then Exit Sub

    Motif = SelectionCourante("PEN_MOTIFS")
    Joueur = PosteSelectionne

    If Motif = "" Then Motif = "?"

    ' GetPlayerName accepte un numero, "Collectif" ou "?"
    ' et rend le nom porte par la compo.
    If Joueur = "" Then
        Joueur = "?"
    Else
        Joueur = GetPlayerName(Joueur)
    End If

    If Trim(Joueur) = "" Then Joueur = "?"

    SetCurrentTeam PenaliteEquipe

    AjouterAction _
        Joueur, _
        PenaliteAction, _
        PenaliteTemps, _
        Motif, _
        GetPlayerGroup(Joueur)

    FermerPopupPenalite

End Sub


Public Sub AnnulerPopupPenalite()

    If Not PenaliteOuverte Then Exit Sub

    ' Comme la palette avant elle : la penalite est
    ' enregistree, motif et joueur inconnus.
    SetCurrentTeam PenaliteEquipe

    AjouterAction _
        "?", _
        PenaliteAction, _
        PenaliteTemps, _
        "?", _
        "?"

    FermerPopupPenalite

End Sub


Private Sub FermerPopupPenalite()

    Dim ws As Worksheet

    PenaliteOuverte = False

    Application.OnKey "~"
    Application.OnKey "{ENTER}"

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    shSaisieVideo.Activate
    ws.Visible = xlSheetHidden

End Sub


' ---------------------------------------------------------
' Selection
'
' Un clic met la case en evidence et eteint les autres :
' un seul motif et un seul joueur.
' ---------------------------------------------------------

Public Sub ClicPopupPenalite(ByVal Target As Range)

    Dim ws As Worksheet

    If Target.Areas.Count > 1 Then Exit Sub

    ' Un clic sur une cellule fusionnee livre toute la
    ' plage : on ne garde que son ancre.
    If Target.Cells.Count > 1 Then

        If Not Target.MergeCells Then Exit Sub
        Set Target = Target.Cells(1, 1)

    End If

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    If Not Intersect(Target, ws.Range("PEN_VALIDER")) _
        Is Nothing Then

        ValiderPopupPenalite
        Exit Sub

    End If

    If Not Intersect(Target, ws.Range("PEN_ANNULER")) _
        Is Nothing Then

        AnnulerPopupPenalite
        Exit Sub

    End If

    BasculerSelection ws, Target, "PEN_MOTIFS", SABLE
    BasculerSelection ws, Target, "PEN_JOUEURS", CIEL

End Sub


Private Sub BasculerSelection( _
    ByVal ws As Worksheet, _
    ByVal Target As Range, _
    ByVal NomGrille As String, _
    ByVal CouleurNormale As Long)

    Dim Grille As Range
    Dim Cellule As Range
    Dim DejaChoisie As Boolean

    Set Grille = ws.Range(NomGrille)

    If Intersect(Target, Grille) Is Nothing Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub

    DejaChoisie = (Target.Interior.Color = BLEU)

    Application.EnableEvents = False

    For Each Cellule In Grille

        If Trim(CStr(Cellule.Value)) <> "" Then
            Cellule.Interior.Color = CouleurNormale
            Cellule.Font.Color = RGB(0, 0, 0)
        End If

    Next Cellule

    If Not DejaChoisie Then
        Target.Interior.Color = BLEU
        Target.Font.Color = RGB(255, 255, 255)
    End If

    Application.EnableEvents = True

End Sub


' Poste choisi dans la composition, rendu par sa position
' et non par ce que la case affiche.
Private Function PosteSelectionne() As String

    Dim ws As Worksheet
    Dim Entrees As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        If ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))) _
            .Interior.Color = BLEU Then

            PosteSelectionne = Champ(CStr(Entrees(i)), 2)
            Exit Function

        End If

    Next i

End Function


Private Function SelectionCourante( _
    ByVal NomGrille As String) As String

    Dim Cellule As Range

    For Each Cellule In ThisWorkbook _
        .Sheets(FEUILLE_POPUP_PEN).Range(NomGrille)

        If Cellule.Interior.Color = BLEU Then
            SelectionCourante = Trim(CStr(Cellule.Value))
            Exit Function
        End If

    Next Cellule

End Function


Private Sub EffacerSelections(ByVal ws As Worksheet)

    Application.EnableEvents = False

    RendreGrille ws, "PEN_MOTIFS", SABLE
    RendreGrille ws, "PEN_JOUEURS", CIEL

    Application.EnableEvents = True

End Sub


Private Sub RendreGrille( _
    ByVal ws As Worksheet, _
    ByVal NomGrille As String, _
    ByVal Couleur As Long)

    Dim Cellule As Range

    For Each Cellule In ws.Range(NomGrille)

        If Trim(CStr(Cellule.Value)) <> "" Then
            Cellule.Interior.Color = Couleur
            Cellule.Font.Color = RGB(0, 0, 0)
        End If

    Next Cellule

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function FormaterTemps( _
    ByVal Secondes As Double) As String

    Dim Total As Long

    Total = CLng(Int(Secondes))

    FormaterTemps = Format(Total \ 60, "00") & ":" & _
        Format(Total Mod 60, "00")

End Function


Private Sub NommerPlage( _
    ByVal ws As Worksheet, _
    ByVal Adresse As String, _
    ByVal NomPlage As String)

    On Error Resume Next
    ThisWorkbook.Names(NomPlage).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=NomPlage, _
        RefersTo:=ws.Range(Adresse)

End Sub


Private Function FeuillePopup() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = FEUILLE_POPUP_PEN

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePopup = ws

End Function
