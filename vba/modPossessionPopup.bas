Attribute VB_Name = "modPossessionPopup"
Option Explicit

' =========================================================
' POSSESSION SUR LA REPRISE EN TOUCHE
'
' Aucune regle ne permet de deviner qui jouera la touche :
' un 50:22 la rend a l'equipe qui a tape, une touche
' rapide peut revenir a celle qui vient de la conceder.
' La question est donc posee a chaque reprise.
'
' La popup n'est pas modale - une feuille ne bloque pas le
' code - alors l'action est mise de cote, puis ecrite au
' clic sur l'un des deux boutons. Meme principe que les
' popups touche et melee.
' =========================================================

Public Const FEUILLE_POPUP_POSS As String = "Popup possession"

Private Const VERT As Long = 32768        ' RGB(0, 128, 0)
Private Const BLEU_SOMBRE As Long = 7884319
Private Const GRIS As Long = 15132390     ' RGB(230, 230, 230)

Private Const NOM_GARDER As String = "POPUP_POSS_GARDER"
Private Const NOM_CHANGER As String = "POPUP_POSS_CHANGER"
Private Const NOM_QUESTION As String = "POPUP_POSS_QUESTION"

' L'action attend que la possession soit tranchee.
Private PossessionAction As String
Private PossessionTemps As Double
Private PossessionEnAttente As Boolean


' ---------------------------------------------------------
' Construction
' ---------------------------------------------------------

Public Sub ConstruirePopupPossession()

    Dim ws As Worksheet

    Set ws = FeuillePossession()

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ws.Columns("A:A").ColumnWidth = 2
    ws.Columns("B:I").ColumnWidth = 11

    With ws.Range("B2")
        .Value = "REPRISE SUR TOUCHE"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = BLEU_SOMBRE
    End With

    ' La phrase est reecrite a chaque ouverture : elle
    ' nomme l'equipe qui detient le ballon a cet instant.
    With ws.Range("B4:I5")
        .Merge
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 11
    End With

    NommerPlagePossession ws, "B4:I5", NOM_QUESTION

    With ws.Range("B7:D8")
        .Merge
        .Value = "Ne change pas de main" & vbLf & "(Echap)"
        .Interior.Color = GRIS
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
    End With

    With ws.Range("F7:I8")
        .Merge
        .Value = "Change de main" & vbLf & "(Entree)"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    NommerPlagePossession ws, "B7:D8", NOM_GARDER
    NommerPlagePossession ws, "F7:I8", NOM_CHANGER

    ws.Rows("7:8").RowHeight = 22

    ws.Visible = xlSheetHidden

    MsgBox _
        "La feuille """ & FEUILLE_POPUP_POSS & """ est prete." & _
        vbCrLf & vbCrLf & _
        "Colle dans son module :" & vbCrLf & vbCrLf & _
        "Private Sub Worksheet_SelectionChange(" & _
        "ByVal Target As Range)" & vbCrLf & _
        "    ClicPopupPossession Target" & vbCrLf & _
        "End Sub", _
        vbInformation, _
        "Popup possession"

End Sub


' ---------------------------------------------------------
' Ouverture
' ---------------------------------------------------------

Public Sub OuvrirPopupPossession( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double)

    Dim ws As Worksheet
    Dim Actuelle As String
    Dim Suivante As String

    PossessionAction = ActionTexte
    PossessionTemps = TempsVideo
    PossessionEnAttente = True

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_POSS)

    Actuelle = EquipeLisible(CurrentTeam)
    Suivante = EquipeLisible(EquipeOpposee(CurrentTeam))

    ws.Range(NOM_QUESTION).Cells(1, 1).Value = _
        "La possession est actuellement sur " & _
        Actuelle & "." & vbLf & vbLf & _
        "Quelle " & ChrW(233) & "quipe aura la possession " & _
        "sur cette phase ?" & vbLf & vbLf & _
        "Ne change pas de main : " & Actuelle & _
        "     -     Change de main : " & Suivante

    SuspendreVideoPourPopup TempsVideo

    ws.Visible = xlSheetVisible
    ws.Activate

    ActiverEspacePopup
    ActiverTouchesPossession

    Application.EnableEvents = False
    ws.Range("A1").Select
    Application.EnableEvents = True

End Sub


' ---------------------------------------------------------
' Reponses
' ---------------------------------------------------------

Public Sub PossessionInchangee()

    FinaliserReprise False

End Sub


Public Sub PossessionChangee()

    FinaliserReprise True

End Sub


Private Sub FinaliserReprise(ByVal Basculer As Boolean)

    Dim ws As Worksheet

    If Not PossessionEnAttente Then Exit Sub

    PossessionEnAttente = False

    DesactiverTouchesPossession

    ' La bascule precede l'ecriture : la ligne du journal
    ' porte l'equipe qui joue la touche, pas celle qui
    ' vient de perdre le ballon.
    If Basculer Then
        ToggleCurrentTeam
    End If

    AjouterAction "", PossessionAction, PossessionTemps

    ReprendreVideoApresPopup

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_POSS)

    shSaisieVideo.Activate
    ws.Visible = xlSheetHidden

    ReleaseButton

End Sub


' ---------------------------------------------------------
' Clic
' ---------------------------------------------------------

Public Sub ClicPopupPossession(ByVal Target As Range)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_POSS)

    If Not Intersect( _
        Target, ws.Range(NOM_GARDER)) Is Nothing Then

        PossessionInchangee
        Exit Sub

    End If

    If Not Intersect( _
        Target, ws.Range(NOM_CHANGER)) Is Nothing Then

        PossessionChangee
        Exit Sub

    End If

End Sub


' ---------------------------------------------------------
' Clavier
' ---------------------------------------------------------

Private Sub ActiverTouchesPossession()

    On Error Resume Next
    Application.OnKey "~", "PossessionChangee"
    Application.OnKey "{ENTER}", "PossessionChangee"
    Application.OnKey "{ESC}", "PossessionInchangee"
    On Error GoTo 0

End Sub


Private Sub DesactiverTouchesPossession()

    On Error Resume Next
    Application.OnKey "~"
    Application.OnKey "{ENTER}"
    Application.OnKey "{ESC}"
    On Error GoTo 0

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function FeuillePossession() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_POSS)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = FEUILLE_POPUP_POSS

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePossession = ws

End Function


Private Sub NommerPlagePossession( _
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


Private Function EquipeLisible( _
    ByVal Equipe As String) As String

    If StrComp(Equipe, "Nous", vbTextCompare) = 0 Then
        EquipeLisible = "Nous"
    Else
        EquipeLisible = "l'adversaire"
    End If

End Function


Private Function EquipeOpposee( _
    ByVal Equipe As String) As String

    If StrComp(Equipe, "Nous", vbTextCompare) = 0 Then
        EquipeOpposee = "Adv"
    Else
        EquipeOpposee = "Nous"
    End If

End Function
