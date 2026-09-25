Attribute VB_Name = "modTouchesMeleesPopup"
Option Explicit

' =========================================================
' TOUCHES ET MELEES - FABRICATION DES DEUX POPUPS
'
' A executer une fois depuis "Createur de match.xlsm",
' apres ConstruireFeuilleTouchesMelees : les listes
' deroulantes pointent sur les colonnes masquees de la
' feuille "Touches Melees".
'
' Deux feuilles separees plutot qu'une seule : une touche
' et une melee ne demandent pas les memes informations,
' et la feuille "Popup" qui sert a ajouter un joueur reste
' intacte.
'
' Aucun caractere accentue, pas meme dans les chaines :
' le VBE de macOS importe les .bas en Mac Roman.
'
' Les quatre premieres lignes affichent ce que la saisie
' video a deja determine. Elles sont grisees pour signaler
' qu'on n'a pas a y toucher : une correction se fait dans
' le tableau, ou la liste deroulante est la pour ca.
' =========================================================

Public Const FEUILLE_POPUP_TO As String = "Popup touche"
Public Const FEUILLE_POPUP_ME As String = "Popup melee"

Private Const FEUILLE_TABLEAUX As String = "Touches Melees"

' Colonnes masquees de la feuille des tableaux, ou vivent
' les listes. Voir modTouchesMelees.EcrireListes.
Private Const COL_ZONE_SAUT As String = "$Z$2:$Z$6"
Private Const COL_BALLON As String = "$AF$2:$AF$3"
Private Const COL_ZONE_LONGUEUR As String = "$AB$2:$AB$12"
Private Const COL_ALIGNEMENT As String = "$AC$2:$AC$7"
Private Const COL_ZONE_LARGEUR As String = "$AD$2:$AD$4"

' Listes en toutes lettres, propres aux popups. Les
' tableaux gardent leurs abreviations : la conversion se
' fait a l'ecriture, dans modTouchesMeleesSaisie.
Private Const COL_UTILISATION As String = "$AE$2:$AE$4"

Private Const GRIS As Long = 15132390     ' RGB(230, 230, 230)
Private Const BLEU As Long = 12611584     ' RGB(0, 112, 192)
Private Const VERT As Long = 32768        ' RGB(0, 128, 0)
Private Const ROUGE As Long = 192         ' RGB(192, 0, 0)

Private EtapeEnCours As String


Public Sub ConstruirePopupsTouchesMelees()

    Dim EtatAffichage As Boolean

    EtatAffichage = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    EtapeEnCours = "popup touche"
    ConstruirePopupTouche

    EtapeEnCours = "popup melee"
    ConstruirePopupMelee

    Application.ScreenUpdating = EtatAffichage

    MsgBox _
        "Les deux popups sont pretes." & vbCrLf & vbCrLf & _
        "Elles sont visibles pour que tu en verifies le " & _
        "rendu. Le VBA de liaison les masquera.", _
        vbInformation, _
        "Touches et melees"

    Exit Sub

GestionErreur:

    Application.ScreenUpdating = EtatAffichage

    MsgBox _
        "La construction a echoue pendant : " & _
        EtapeEnCours & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


Public Sub SupprimerPopupsTouchesMelees()

    Dim EtatAlertes As Boolean

    If MsgBox( _
        "Supprimer les deux feuilles popup ?", _
        vbYesNo + vbExclamation, _
        "Touches et melees") <> vbYes Then Exit Sub

    EtatAlertes = Application.DisplayAlerts
    Application.DisplayAlerts = False

    SupprimerFeuille FEUILLE_POPUP_TO
    SupprimerFeuille FEUILLE_POPUP_ME

    Application.DisplayAlerts = EtatAlertes

End Sub


' ---------------------------------------------------------
' Popup touche
' ---------------------------------------------------------

Private Sub ConstruirePopupTouche()

    Dim ws As Worksheet

    Set ws = FeuillePopup(FEUILLE_POPUP_TO)

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws, "TOUCHE"
    EcrireContexte ws, "Lance pour"

    EcrireChamp ws, 9, "Sauteur", "POPUP_TO_SAUTEUR", ""
    EcrireChamp ws, 10, "Zone de saut", _
        "POPUP_TO_ZONE_SAUT", COL_ZONE_SAUT
    EcrireChamp ws, 11, "Ballon", _
        "POPUP_TO_BALLON", COL_BALLON
    EcrireChamp ws, 12, "Zone terrain", _
        "POPUP_TO_ZONE_TERRAIN", COL_ZONE_LONGUEUR
    EcrireChamp ws, 13, "Alignement", _
        "POPUP_TO_ALIGNEMENT", COL_ALIGNEMENT
    EcrireChamp ws, 14, "Observations", _
        "POPUP_TO_OBS", ""

    EcrireBoutons ws, 16, "POPUP_TO_VALIDER", _
        "POPUP_TO_ANNULER"

    NommerPlage ws, "C4:C7", "POPUP_TO_CONTEXTE"
    NommerPlage ws, "C9:C14", "POPUP_TO_CHAMPS"

    MettreEnFormePopup ws, 16

End Sub


' ---------------------------------------------------------
' Popup melee
'
' Une seule liste pour l'utilisation : le tableau garde
' ses trois colonnes, la saisie n'en demande qu'une.
' ---------------------------------------------------------

Private Sub ConstruirePopupMelee()

    Dim ws As Worksheet

    Set ws = FeuillePopup(FEUILLE_POPUP_ME)

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws, "MELEE"
    EcrireContexte ws, "Introduction pour"

    EcrireChamp ws, 9, "Zone longueur", _
        "POPUP_ME_ZONE_LONGUEUR", COL_ZONE_LONGUEUR
    EcrireChamp ws, 10, "Zone largeur", _
        "POPUP_ME_ZONE_LARGEUR", COL_ZONE_LARGEUR
    EcrireChamp ws, 11, "Utilisation", _
        "POPUP_ME_UTILISATION", COL_UTILISATION
    EcrireChamp ws, 12, "Observations", _
        "POPUP_ME_OBS", ""

    EcrireBoutons ws, 14, "POPUP_ME_VALIDER", _
        "POPUP_ME_ANNULER"

    NommerPlage ws, "C4:C7", "POPUP_ME_CONTEXTE"
    NommerPlage ws, "C9:C12", "POPUP_ME_CHAMPS"

    MettreEnFormePopup ws, 14

End Sub


' ---------------------------------------------------------
' Briques communes
' ---------------------------------------------------------

Private Sub EcrireTitre( _
    ByVal ws As Worksheet, _
    ByVal Titre As String)

    With ws.Range("B2")
        .Value = Titre
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = BLEU
    End With

End Sub


' Les quatre valeurs que la saisie video connait deja.
Private Sub EcrireContexte( _
    ByVal ws As Worksheet, _
    ByVal LibelleEquipe As String)

    Dim Libelles As Variant
    Dim i As Long

    Libelles = Array( _
        "Temps video", "Mi-temps", LibelleEquipe, "Issue")

    For i = 0 To 3

        ws.Cells(4 + i, 2).Value = Libelles(i)
        ws.Cells(4 + i, 2).Font.Italic = True

        With ws.Cells(4 + i, 3)
            .Interior.Color = GRIS
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

    Next i

End Sub


Private Sub EcrireChamp( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String, _
    ByVal NomPlage As String, _
    ByVal Source As String)

    ws.Cells(Ligne, 2).Value = Libelle

    With ws.Cells(Ligne, 3)

        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter

    End With

    NommerPlage ws, ws.Cells(Ligne, 3).Address, NomPlage

    If Source <> "" Then PoserValidation ws, Ligne, Source

End Sub


' ---------------------------------------------------------
' Listes deroulantes
'
' La source est une plage d'une autre feuille, citee par
' adresse. modPlayers.ActualiserJoueursJournal procede
' ainsi depuis Compo vers le Journal : c'est accepte.
'
' Alerte non bloquante ici : la popup doit pouvoir etre
' validee sans qu'un champ soit renseigne.
' ---------------------------------------------------------

Private Sub PoserValidation( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Source As String)

    With ws.Cells(Ligne, 3).Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="='" & FEUILLE_TABLEAUX & "'!" & Source

    End With

End Sub


' ---------------------------------------------------------
' Boutons
'
' Des cellules nommees, pas des formes : la palette de
' saisie fonctionne deja ainsi, et Worksheet_SelectionChange
' suffit a les detecter.
' ---------------------------------------------------------

Private Sub EcrireBoutons( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal NomValider As String, _
    ByVal NomAnnuler As String)

    ' ANNULER sous les libelles, dans la colonne etroite ;
    ' VALIDER sous les champs, ou la place ne manque pas.
    With ws.Cells(Ligne, 2)
        .Value = "ANNULER"
        .Interior.Color = ROUGE
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(Ligne, 3)
        .Value = "VALIDER"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    NommerPlage ws, ws.Cells(Ligne, 2).Address, NomAnnuler
    NommerPlage ws, ws.Cells(Ligne, 3).Address, NomValider

End Sub


Private Sub MettreEnFormePopup( _
    ByVal ws As Worksheet, _
    ByVal LigneBoutons As Long)

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 22

    ws.Rows(LigneBoutons).RowHeight = 24

    ws.Cells.Font.Name = "Calibri"

    ' Laissee visible : c'est la seule facon d'en verifier
    ' le rendu. Le VBA de liaison la masquera, la saisie
    ' video l'affichant alors au moment utile.

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function FeuillePopup( _
    ByVal Nom As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Nom)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = Nom

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePopup = ws

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


Private Sub SupprimerFeuille(ByVal Nom As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Nom)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ws.Visible = xlSheetVisible
    ws.Delete

End Sub
