Attribute VB_Name = "modRemplacements"
Option Explicit

' =========================================================
' REMPLACEMENTS
'
' Le bouton "Remplacement" de la palette met la video en
' pause, releve sa position, et ouvre une popup ou l'on
' saisit la minute de jeu puis les mouvements.
'
' Deux echelles de temps cohabitent, et les deux servent :
'
'   la MINUTE DE JEU, saisie a la main, sert au temps de
'   jeu des joueurs ; elle ne se deduit pas de la video,
'   qui ignore les arrets ;
'
'   le TEMPS VIDEO, releve automatiquement, sert a savoir
'   qui etait sur le terrain a un instant de la video. En
'   revenant en arriere, la palette retrouve l'equipe de
'   ce moment-la.
'
' Chaque mouvement occupe une colonne a droite du tableau
' COMPO :
'
'   ligne 6  : la minute de jeu, en tete de colonne
'   lignes 7 a 28 : "S" pour le sortant, "E" pour l'entrant
'   ligne 29 : le temps video en secondes, ligne masquee
'
' Aucun caractere accentue en clair : le VBE de macOS
' importe les .bas en Mac Roman.
' =========================================================

Public Const FEUILLE_POPUP_REMP As String = "Popup remplacement"

' Tableau COMPO : E6:H28, en-tete ligne 6, postes 7 a 28.
Private Const LIGNE_ENTETE As Long = 6
Private Const PREMIER_POSTE As Long = 7
Private Const DERNIER_POSTE As Long = 28
Private Const LIGNE_TEMPS_VIDEO As Long = 29
Private Const COL_POSTE As Long = 5
Private Const COL_NOM As Long = 6
Private Const PREMIERE_COL_MOUVEMENT As Long = 9

' Nombre de titulaires : les postes 1 a 15 commencent le
' match sur le terrain.
Private Const NB_TITULAIRES As Long = 15

Private Const NB_LIGNES_SAISIE As Long = 5

' Duree de reference d'un match, pour le temps de jeu de
' qui termine la rencontre sur le terrain.
Private Const DUREE_MATCH As Long = 80

' Derniere colonne du tableau COMPO : son format sert de
' modele aux colonnes de mouvement.
Private Const COL_MODELE As Long = 8

' Colonnes masquees de la popup, ou vivent les listes.
Private Const COL_LISTE_SORTANTS As Long = 10
Private Const COL_LISTE_ENTRANTS As Long = 12

Private Const GRIS As Long = 15132390     ' RGB(230,230,230)
Private Const BLEU As Long = 12611584     ' RGB(0,112,192)
Private Const VERT As Long = 32768        ' RGB(0,128,0)
Private Const ROUGE As Long = 192         ' RGB(192,0,0)

Private RempTempsVideo As Double
Private RempOuverte As Boolean


' ---------------------------------------------------------
' Le bouton de la palette
' ---------------------------------------------------------

Public Sub CreerBoutonRemplacement()

    Dim ws As Worksheet

    On Error GoTo GestionErreur

    Set ws = shSaisieVideo

    ws.Range("AF8").Value = "Remplacement"

    ApplyFormat _
        ws.Range("AF8"), _
        shParametres.Range("STYLE_BTN_POSSESSION")

    On Error Resume Next
    ThisWorkbook.Names("BTN_REMPLACEMENT").Delete
    On Error GoTo GestionErreur

    ThisWorkbook.Names.Add _
        Name:="BTN_REMPLACEMENT", _
        RefersTo:=ws.Range("AF8")

    MsgBox _
        "Le bouton Remplacement est en place en AF8.", _
        vbInformation, _
        "Remplacements"

    Exit Sub

GestionErreur:

    MsgBox _
        "La creation du bouton a echoue." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Remplacements"

End Sub


' ---------------------------------------------------------
' Construction de la popup
' ---------------------------------------------------------

Public Sub ConstruirePopupRemplacement()

    Dim ws As Worksheet
    Dim i As Long

    On Error GoTo GestionErreur

    Set ws = FeuillePopup

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    With ws.Range("B2")
        .Value = "REMPLACEMENT"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = BLEU
    End With

    ws.Range("B4").Value = "Temps video"
    ws.Range("B4").Font.Italic = True

    With ws.Range("D4")
        .Interior.Color = GRIS
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("B5").Value = "Minute de jeu"
    ws.Range("B5").Font.Bold = True

    With ws.Range("D5")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
    End With

    ws.Range("B7").Value = "Joueur sortant"
    ws.Range("D7").Value = "Joueur entrant"
    ws.Range("B7,D7").Font.Bold = True

    ' Aucune cellule fusionnee ici : une liste deroulante
    ' ne peut pas se poser sur une fusion.
    For i = 0 To NB_LIGNES_SAISIE - 1

        MettreEnFormeCase ws.Cells(8 + i, 2)
        MettreEnFormeCase ws.Cells(8 + i, 4)

    Next i

    With ws.Cells(8 + NB_LIGNES_SAISIE + 1, 2)
        .Value = "ANNULER"
        .Interior.Color = ROUGE
    End With

    With ws.Cells(8 + NB_LIGNES_SAISIE + 1, 4)
        .Value = "VALIDER"
        .Interior.Color = VERT
    End With

    With ws.Range( _
        ws.Cells(8 + NB_LIGNES_SAISIE + 1, 2), _
        ws.Cells(8 + NB_LIGNES_SAISIE + 1, 4))

        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ws.Rows(8 + NB_LIGNES_SAISIE + 1).RowHeight = 26

    NommerPlage ws, "D4", "REMP_TEMPS"
    NommerPlage ws, "D5", "REMP_MINUTE"

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(8, 2), _
            ws.Cells(7 + NB_LIGNES_SAISIE, 2)).Address, _
        "REMP_SORTANTS"

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(8, 4), _
            ws.Cells(7 + NB_LIGNES_SAISIE, 4)).Address, _
        "REMP_ENTRANTS"

    NommerPlage ws, _
        ws.Cells(8 + NB_LIGNES_SAISIE + 1, 2).Address, _
        "REMP_ANNULER"

    NommerPlage ws, _
        ws.Cells(8 + NB_LIGNES_SAISIE + 1, 4).Address, _
        "REMP_VALIDER"

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 26

    ws.Range( _
        ws.Columns(COL_LISTE_SORTANTS), _
        ws.Columns(COL_LISTE_ENTRANTS + 1)).Hidden = True

    ws.Cells.Font.Name = "Calibri"

    MsgBox _
        "La popup des remplacements est prete.", _
        vbInformation, _
        "Remplacements"

    Exit Sub

GestionErreur:

    MsgBox _
        "La construction a echoue." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Remplacements"

End Sub


Public Sub SupprimerPopupRemplacement()

    Dim ws As Worksheet
    Dim EtatAlertes As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    EtatAlertes = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ws.Visible = xlSheetVisible
    ws.Delete

    Application.DisplayAlerts = EtatAlertes

End Sub


Private Sub MettreEnFormeCase(ByVal Zone As Range)

    With Zone
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
    End With

End Sub


' ---------------------------------------------------------
' Qui est sur le terrain
'
' On part des quinze titulaires, puis on applique les
' mouvements dont le temps video precede celui demande.
' ---------------------------------------------------------

Public Function PostesSurLeTerrain( _
    ByVal TempsVideo As Double) As Collection

    Dim Presents As New Collection
    Dim ws As Worksheet
    Dim Colonne As Long
    Dim Ligne As Long
    Dim Poste As String
    Dim Marque As String
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Compo")

    For i = 1 To NB_TITULAIRES
        Presents.Add CStr(i), CStr(i)
    Next i

    Colonne = PREMIERE_COL_MOUVEMENT

    Do While Trim(CStr( _
        ws.Cells(LIGNE_ENTETE, Colonne).Value)) <> ""

        If TempsMouvement(ws, Colonne) <= TempsVideo Then

            For Ligne = PREMIER_POSTE To DERNIER_POSTE

                Marque = UCase(Trim(CStr( _
                    ws.Cells(Ligne, Colonne).Value)))

                If Marque <> "" Then

                    Poste = Trim(CStr( _
                        ws.Cells(Ligne, COL_POSTE).Value))

                    On Error Resume Next

                    If Marque = "S" Then
                        Presents.Remove Poste
                    ElseIf Marque = "E" Then
                        Presents.Add Poste, Poste
                    End If

                    On Error GoTo 0

                End If

            Next Ligne

        End If

        Colonne = Colonne + 1

    Loop

    Set PostesSurLeTerrain = Presents

End Function


Private Function TempsMouvement( _
    ByVal ws As Worksheet, _
    ByVal Colonne As Long) As Double

    Dim Valeur As Variant

    Valeur = ws.Cells(LIGNE_TEMPS_VIDEO, Colonne).Value

    If IsNumeric(Valeur) Then TempsMouvement = CDbl(Valeur)

End Function


' Etiquette lisible d'un poste : "12 - LAUDET".
Private Function Etiquette(ByVal Poste As String) As String

    Dim Nom As String

    Nom = Trim(GetPlayerName(Poste))

    If Nom = "" Or Nom = Poste Then
        Etiquette = Poste
    Else
        Etiquette = Poste & " - " & Nom
    End If

End Function


Private Function PosteDeLEtiquette( _
    ByVal Texte As String) As String

    Dim p As Long

    p = InStr(Texte, " - ")

    If p = 0 Then
        PosteDeLEtiquette = Trim(Texte)
    Else
        PosteDeLEtiquette = Trim(Left(Texte, p - 1))
    End If

End Function


' ---------------------------------------------------------
' Ouverture et fermeture
' ---------------------------------------------------------

Public Sub OuvrirPopupRemplacement(ByVal TempsVideo As Double)

    Dim ws As Worksheet

    RempTempsVideo = TempsVideo
    RempOuverte = True

    FermerActionEnAttente
    MettreVideoEnPause

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)

    ws.Visible = xlSheetVisible
    ws.Activate

    ' Les evenements sont rendus quoi qu'il arrive : une
    ' erreur ici bloquerait toute la palette ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    ws.Range("REMP_TEMPS").Value = FormaterTemps(TempsVideo)
    ws.Range("REMP_MINUTE").ClearContents
    ws.Range("REMP_SORTANTS").ClearContents
    ws.Range("REMP_ENTRANTS").ClearContents

    RemplirListes ws, TempsVideo

Sortie:

    Application.EnableEvents = True

    If Err.Number <> 0 Then

        MsgBox _
            "Les listes n'ont pas pu etre construites." & _
            vbCrLf & vbCrLf & _
            "Erreur " & Err.Number & " : " & _
            Err.Description, _
            vbExclamation, _
            "Remplacements"

    End If

    ws.Range("REMP_MINUTE").Select

End Sub


' Le moteur ne dit pas s'il lit : on regarde si le temps
' avance entre deux releves.
Private Sub MettreVideoEnPause()

    Dim Apres As Double
    Dim Fin As Single

    On Error Resume Next

    If RempTempsVideo <= 0 Then Exit Sub

    Fin = Timer + 0.3

    Do While Timer < Fin
        DoEvents
    Loop

    Apres = GetTimeVideo()

    If Apres < 0 Then Exit Sub

    If Abs(Apres - RempTempsVideo) > 0.05 Then
        PlayPauseChronoVideo
    End If

    On Error GoTo 0

End Sub


' Les deux listes sont refaites a chaque ouverture : elles
' dependent du moment de la video.
Private Sub RemplirListes( _
    ByVal ws As Worksheet, _
    ByVal TempsVideo As Double)

    Dim Presents As Collection
    Dim wsCompo As Worksheet
    Dim Ligne As Long
    Dim Poste As String
    Dim SurLeTerrain As Boolean
    Dim nSortants As Long
    Dim nEntrants As Long
    Dim Element As Variant

    Set Presents = PostesSurLeTerrain(TempsVideo)
    Set wsCompo = ThisWorkbook.Sheets("Compo")

    ws.Columns(COL_LISTE_SORTANTS).ClearContents
    ws.Columns(COL_LISTE_ENTRANTS).ClearContents

    nSortants = 0
    nEntrants = 0

    For Ligne = PREMIER_POSTE To DERNIER_POSTE

        Poste = Trim(CStr( _
            wsCompo.Cells(Ligne, COL_POSTE).Value))

        If Poste <> "" Then

            SurLeTerrain = False

            For Each Element In Presents
                If CStr(Element) = Poste Then
                    SurLeTerrain = True
                    Exit For
                End If
            Next Element

            If SurLeTerrain Then
                nSortants = nSortants + 1
                ws.Cells(nSortants, COL_LISTE_SORTANTS).Value = _
                    Etiquette(Poste)
            Else
                nEntrants = nEntrants + 1
                ws.Cells(nEntrants, COL_LISTE_ENTRANTS).Value = _
                    Etiquette(Poste)
            End If

        End If

    Next Ligne

    PoserListe ws, "REMP_SORTANTS", _
        COL_LISTE_SORTANTS, nSortants

    PoserListe ws, "REMP_ENTRANTS", _
        COL_LISTE_ENTRANTS, nEntrants

End Sub


Private Sub PoserListe( _
    ByVal ws As Worksheet, _
    ByVal NomPlage As String, _
    ByVal Colonne As Long, _
    ByVal Nombre As Long)

    Dim Source As String

    Dim Cellule As Range

    If Nombre < 1 Then Exit Sub

    Source = "=" & ws.Range( _
        ws.Cells(1, Colonne), _
        ws.Cells(Nombre, Colonne)).Address(True, True)

    For Each Cellule In ws.Range(NomPlage)

        With Cellule.Validation

            .Delete

            .Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:=Source

        End With

    Next Cellule

End Sub


Public Sub ValiderPopupRemplacement()

    Dim ws As Worksheet
    Dim Minute As Variant
    Dim i As Long
    Dim Sortant As String
    Dim Entrant As String
    Dim Ecrits As Long

    If Not RempOuverte Then Exit Sub

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)

    Minute = ws.Range("REMP_MINUTE").Value

    If Not IsNumeric(Minute) Then

        MsgBox _
            "Renseigne la minute de jeu avant de valider.", _
            vbExclamation, _
            "Minute manquante"

        Exit Sub

    End If

    For i = 1 To NB_LIGNES_SAISIE

        Sortant = PosteDeLEtiquette(CStr( _
            ws.Range("REMP_SORTANTS").Cells(i, 1).Value))

        Entrant = PosteDeLEtiquette(CStr( _
            ws.Range("REMP_ENTRANTS").Cells(i, 1).Value))

        If Sortant <> "" And Entrant <> "" Then

            EcrireMouvement CLng(Minute), Sortant, Entrant
            Ecrits = Ecrits + 1

        End If

    Next i

    If Ecrits > 0 Then RecalculerTempsDeJeu

    If Ecrits = 0 Then

        MsgBox _
            "Aucun mouvement complet : chaque ligne " & _
            "demande un sortant et un entrant.", _
            vbExclamation, _
            "Remplacements"

        Exit Sub

    End If

    FermerPopupRemplacement

End Sub


Public Sub AnnulerPopupRemplacement()

    If Not RempOuverte Then Exit Sub

    FermerPopupRemplacement

End Sub


' Un mouvement, une colonne : la minute en tete, le temps
' video en pied, "S" et "E" sur les lignes concernees.
Private Sub EcrireMouvement( _
    ByVal Minute As Long, _
    ByVal Sortant As String, _
    ByVal Entrant As String)

    Dim ws As Worksheet
    Dim Colonne As Long
    Dim Ligne As Long
    Dim Poste As String

    Set ws = ThisWorkbook.Sheets("Compo")

    Colonne = PREMIERE_COL_MOUVEMENT

    Do While Trim(CStr( _
        ws.Cells(LIGNE_ENTETE, Colonne).Value)) <> ""
        Colonne = Colonne + 1
    Loop

    ' La colonne neuve reprend le format du tableau :
    ' sans cela elle garde le fond de la feuille et
    ' tranche avec les bandes.
    ws.Range( _
        ws.Cells(LIGNE_ENTETE, COL_MODELE), _
        ws.Cells(DERNIER_POSTE, COL_MODELE)).Copy

    ws.Range( _
        ws.Cells(LIGNE_ENTETE, Colonne), _
        ws.Cells(DERNIER_POSTE, Colonne)) _
        .PasteSpecial xlPasteFormats

    Application.CutCopyMode = False

    ws.Columns(Colonne).ColumnWidth = 6

    ws.Cells(LIGNE_ENTETE, Colonne).Value = Minute
    ws.Cells(LIGNE_ENTETE, Colonne).Font.Bold = True
    ws.Cells(LIGNE_ENTETE, Colonne).HorizontalAlignment = _
        xlCenter

    ws.Cells(LIGNE_TEMPS_VIDEO, Colonne).Value = _
        RempTempsVideo

    For Ligne = PREMIER_POSTE To DERNIER_POSTE

        Poste = Trim(CStr( _
            ws.Cells(Ligne, COL_POSTE).Value))

        If Poste = Sortant Then
            ws.Cells(Ligne, Colonne).Value = "S"
        ElseIf Poste = Entrant Then
            ws.Cells(Ligne, Colonne).Value = "E"
        End If

        ws.Cells(Ligne, Colonne).HorizontalAlignment = _
            xlCenter

    Next Ligne

    ws.Rows(LIGNE_TEMPS_VIDEO).Hidden = True

End Sub


' ---------------------------------------------------------
' Temps de jeu
'
' Recalcule apres chaque remplacement : les minutes sont
' toutes connues a ce moment-la, et la colonne reste juste
' sans qu'on ait a y penser.
'
' Les mouvements sont lus par minute croissante, et non
' dans l'ordre des colonnes : une saisie faite apres coup
' ne doit pas fausser le compte.
' ---------------------------------------------------------

Public Sub RecalculerTempsDeJeu()

    Dim ws As Worksheet
    Dim Ligne As Long
    Dim Poste As Long

    Set ws = ThisWorkbook.Sheets("Compo")

    For Ligne = PREMIER_POSTE To DERNIER_POSTE

        If IsNumeric(ws.Cells(Ligne, COL_POSTE).Value) Then

            Poste = CLng(ws.Cells(Ligne, COL_POSTE).Value)

            ws.Cells(Ligne, COL_MODELE).Value = _
                TempsDuPoste(ws, Ligne, Poste)

        End If

    Next Ligne

End Sub


Private Function TempsDuPoste( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Poste As Long) As Long

    Dim Colonnes As Variant
    Dim i As Long
    Dim Marque As String
    Dim Minute As Long
    Dim SurLeTerrain As Boolean
    Dim Entree As Long
    Dim Total As Long

    ' Les quinze premiers commencent la rencontre.
    SurLeTerrain = (Poste <= NB_TITULAIRES)
    Entree = 0

    Colonnes = ColonnesParMinute(ws)

    For i = 0 To UBound(Colonnes)

        If Colonnes(i) = "" Then Exit For

        Minute = CLng(Split(CStr(Colonnes(i)), ":")(0))

        Marque = UCase(Trim(CStr(ws.Cells( _
            Ligne, _
            CLng(Split(CStr(Colonnes(i)), ":")(1))).Value)))

        If Marque = "S" And SurLeTerrain Then

            Total = Total + (Minute - Entree)
            SurLeTerrain = False

        ElseIf Marque = "E" And Not SurLeTerrain Then

            Entree = Minute
            SurLeTerrain = True

        End If

    Next i

    If SurLeTerrain Then
        Total = Total + (DUREE_MATCH - Entree)
    End If

    TempsDuPoste = Total

End Function


' Les colonnes de mouvement, rendues "minute:colonne" et
' triees par minute croissante.
Private Function ColonnesParMinute( _
    ByVal ws As Worksheet) As Variant

    Dim Liste() As String
    Dim n As Long
    Dim Colonne As Long
    Dim Valeur As Variant
    Dim i As Long
    Dim j As Long
    Dim Tampon As String

    ReDim Liste(0 To 200)
    n = -1

    Colonne = PREMIERE_COL_MOUVEMENT

    Do While Trim(CStr( _
        ws.Cells(LIGNE_ENTETE, Colonne).Value)) <> ""

        Valeur = ws.Cells(LIGNE_ENTETE, Colonne).Value

        If IsNumeric(Valeur) Then
            n = n + 1
            Liste(n) = CLng(Valeur) & ":" & Colonne
        End If

        Colonne = Colonne + 1

    Loop

    If n < 0 Then
        ColonnesParMinute = Array("")
        Exit Function
    End If

    ReDim Preserve Liste(0 To n)

    ' Tri a bulles : une poignee de mouvements par match.
    For i = 0 To n - 1
        For j = i + 1 To n

            If CLng(Split(Liste(i), ":")(0)) > _
                CLng(Split(Liste(j), ":")(0)) Then

                Tampon = Liste(i)
                Liste(i) = Liste(j)
                Liste(j) = Tampon

            End If

        Next j
    Next i

    ColonnesParMinute = Liste

End Function


Private Sub FermerPopupRemplacement()

    Dim ws As Worksheet

    RempOuverte = False

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)

    shSaisieVideo.Activate
    ws.Visible = xlSheetHidden

End Sub


' ---------------------------------------------------------
' Clics
' ---------------------------------------------------------

Public Sub ClicPopupRemplacement(ByVal Target As Range)

    Dim ws As Worksheet

    If Target.Areas.Count > 1 Then Exit Sub

    If Target.Cells.Count > 1 Then

        If Not Target.MergeCells Then Exit Sub
        Set Target = Target.Cells(1, 1)

    End If

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)

    If Not Intersect(Target, ws.Range("REMP_VALIDER")) _
        Is Nothing Then

        ValiderPopupRemplacement
        Exit Sub

    End If

    If Not Intersect(Target, ws.Range("REMP_ANNULER")) _
        Is Nothing Then

        AnnulerPopupRemplacement

    End If

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function FormaterTemps( _
    ByVal Secondes As Double) As String

    Dim Total As Long

    If Secondes < 0 Then
        FormaterTemps = "--:--"
        Exit Function
    End If

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
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = FEUILLE_POPUP_REMP

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePopup = ws

End Function
