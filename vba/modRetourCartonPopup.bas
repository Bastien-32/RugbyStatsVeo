Attribute VB_Name = "modRetourCartonPopup"
Option Explicit

' =========================================================
' RETOUR DE CARTON - POPUP DE SAISIE
'
' Le bouton "Retour" de la ligne des cartons ouvre cette
' popup. Elle ne propose rien d'autre que ce qui est
' reellement en cours a cet instant de la video : un carton
' pose plus tot et dont le retour n'a pas encore ete
' journalise. La liste vient donc du journal, pas d'une
' memoire tenue a part.
'
' Plusieurs lignes peuvent etre cochees : deux exclusions
' qui s'achevent au meme moment se declarent en une fois.
'
' Chaque retour s'ecrit au journal comme une action a part
' entiere. C'est ce qui permettra plus tard de compter
' ce qui s'est passe pendant l'exclusion : les points
' encaisses pendant un carton jaune, par exemple, se lisent
' entre la ligne du carton et celle de son retour.
'
' Rouge et bleu y figurent aussi. Personne n'en declare le
' retour en vrai, mais la popup sert alors a rattraper une
' saisie faite par erreur.
'
' Aucun caractere accentue en clair : le VBE de macOS
' importe les .bas en Mac Roman.
' =========================================================

Public Const FEUILLE_POPUP_RETOUR As String = _
    "Popup retour carton"

' La popup tient dans les colonnes B a H.
Private Const COL_PREMIERE As Long = 2
Private Const COL_DERNIERE As Long = 8

' Les valeurs du contexte se collent a leurs libelles.
Private Const COL_VALEUR As Long = 3

Private Const COL_PLAY_PAUSE As Long = 7

Private Const LIGNE_TITRE_LISTE As Long = 7
Private Const PREMIERE_LIGNE As Long = 9

' Une ligne sur deux : les rangs intercalaires font
' respirer la liste. Dix cartons simultanes ne se verront
' jamais, le compte est large.
Private Const NB_LIGNES_LISTE As Long = 10

Private Const LIGNE_BOUTONS As Long = 29

Private Const GRIS As Long = 15132390     ' RGB(230,230,230)
Private Const BLEU As Long = 12611584     ' RGB(0,112,192)
Private Const VERT As Long = 32768        ' RGB(0,128,0)
Private Const ROUGE As Long = 192         ' RGB(192,0,0)
Private Const CIEL As Long = 15983321     ' RGB(217,226,243)
Private Const MARINE As Long = 6567967    ' RGB(31,56,100)

' Contexte du retour en cours de saisie.
Private RetourTemps As Double
Private RetourOuverte As Boolean

' Le carton porte par chaque ligne de la liste. La case
' affiche une phrase, mais c'est l'evenement du journal qui
' sert a ecrire le retour.
Private CartonsListe(1 To NB_LIGNES_LISTE) As String


' ---------------------------------------------------------
' Construction
' ---------------------------------------------------------

Public Sub ConstruirePopupRetourCarton()

    Dim ws As Worksheet

    On Error GoTo GestionErreur

    Set ws = FeuillePopup

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws
    EcrireContexte ws
    EcrireBandeauListe ws
    EcrireListe ws
    EcrireBoutons ws
    MettreEnForme ws

    If ModeSilencieux Then Exit Sub

    MsgBox _
        "La popup des retours de carton est pr" & _
        ChrW(234) & "te.", _
        vbInformation, _
        "Retour de carton"

    Exit Sub

GestionErreur:

    MsgBox _
        "La construction a " & ChrW(233) & "chou" & _
        ChrW(233) & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Retour de carton"

End Sub


Private Sub EcrireTitre(ByVal ws As Worksheet)

    With ws.Cells(2, COL_PREMIERE)
        .Value = "RETOUR DE CARTON"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = BLEU
    End With

End Sub


Private Sub EcrireContexte(ByVal ws As Worksheet)

    Dim Libelles As Variant
    Dim i As Long

    Libelles = Array("Temps vid" & ChrW(233) & "o", "Mi-temps")

    For i = 0 To UBound(Libelles)

        With ws.Cells(4 + i, COL_PREMIERE)
            .Value = Libelles(i)
            .Font.Italic = True
        End With

        With ws.Cells(4 + i, COL_VALEUR).Resize(1, 2)
            .Merge
            .Interior.Color = GRIS
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(4, COL_VALEUR), _
            ws.Cells(5, COL_VALEUR)).Address, _
        "RET_CONTEXTE"

    ' Lecture et pause sans quitter la popup : on peut
    ' avancer la video pour retrouver le moment du retour.
    With ws.Cells(4, COL_PLAY_PAUSE).Resize(1, 2)

        .Merge
        .Value = "PLAY / PAUSE"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    NommerPlage ws, _
        ws.Cells(4, COL_PLAY_PAUSE).Address, _
        "RET_PLAY_PAUSE"

End Sub


Private Sub EcrireBandeauListe(ByVal ws As Worksheet)

    With ws.Range( _
        ws.Cells(LIGNE_TITRE_LISTE, COL_PREMIERE), _
        ws.Cells(LIGNE_TITRE_LISTE, COL_DERNIERE))

        .Merge
        .Cells(1, 1).Value = _
            "CARTONS EN COURS  -  plusieurs choix possibles"
        .Interior.Color = MARINE
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub


' Les dix cases de la liste sont fusionnees une fois pour
' toutes. Leur contenu, lui, est refait a chaque ouverture.
Private Sub EcrireListe(ByVal ws As Worksheet)

    Dim i As Long

    For i = 1 To NB_LIGNES_LISTE

        With ZoneLigne(ws, LigneDeLIndex(i))

            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

        End With

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(PREMIERE_LIGNE, COL_PREMIERE), _
            ws.Cells(LigneDeLIndex(NB_LIGNES_LISTE), _
                COL_DERNIERE)).Address, _
        "RET_LISTE"

End Sub


Private Sub EcrireBoutons(ByVal ws As Worksheet)

    With ws.Cells(LIGNE_BOUTONS, COL_PREMIERE).Resize(1, 2)
        .Merge
        .Value = "ANNULER"
        .Interior.Color = ROUGE
    End With

    With ws.Cells(LIGNE_BOUTONS, COL_PREMIERE + 3).Resize(1, 4)
        .Merge
        .Value = "VALIDER  (Entree)"
        .Interior.Color = VERT
    End With

    With ws.Range( _
        ws.Cells(LIGNE_BOUTONS, COL_PREMIERE), _
        ws.Cells(LIGNE_BOUTONS, COL_DERNIERE))

        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ws.Rows(LIGNE_BOUTONS).RowHeight = 26

    NommerPlage ws, _
        ws.Cells(LIGNE_BOUTONS, COL_PREMIERE).Address, _
        "RET_ANNULER"

    NommerPlage ws, _
        ws.Cells(LIGNE_BOUTONS, COL_PREMIERE + 3).Address, _
        "RET_VALIDER"

End Sub


Private Sub MettreEnForme(ByVal ws As Worksheet)

    Dim i As Long

    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C:E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("G:H").ColumnWidth = 12

    ws.Cells.Font.Name = "Calibri"

    ws.Rows(LIGNE_TITRE_LISTE).RowHeight = 22

    For i = 1 To NB_LIGNES_LISTE
        ws.Rows(LigneDeLIndex(i)).RowHeight = 22
        ws.Rows(LigneDeLIndex(i) + 1).RowHeight = 8
    Next i

End Sub


' ---------------------------------------------------------
' Ouverture et fermeture
' ---------------------------------------------------------

Public Sub OuvrirPopupRetourCarton(ByVal TempsVideo As Double)

    Dim ws As Worksheet
    Dim Cartons As Collection

    Set Cartons = CartonsEnCours(TempsVideo)

    If Cartons.Count = 0 Then

        MsgBox _
            "Aucun carton n'est en cours " & ChrW(224) & _
            " cet instant de la vid" & ChrW(233) & "o.", _
            vbInformation, _
            "Retour de carton"

        Exit Sub

    End If

    RetourTemps = TempsVideo
    RetourOuverte = True

    ' Une action restee en attente doit etre close avant que
    ' la popup prenne la main.
    FermerActionEnAttente
    SuspendreVideoPourPopup TempsVideo

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_RETOUR)

    ws.Visible = xlSheetVisible
    ws.Activate

    ' Les evenements sont rendus quoi qu'il arrive : coupes,
    ' plus aucun clic ne passerait ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    With ws.Range("RET_CONTEXTE")
        ' Format texte impose : sans lui, Excel lit "64:48"
        ' comme 64 heures 48 et affiche 16:48.
        .Cells(1, 1).NumberFormat = "@"
        .Cells(1, 1).Value = FormaterTemps(TempsVideo)
        .Cells(2, 1).Value = CurrentHalf
    End With

    RemplirListe ws, Cartons

Sortie:

    Application.EnableEvents = True

    If Err.Number <> 0 Then

        MsgBox _
            "La liste des cartons n'a pas pu " & ChrW(234) & _
            "tre construite." & vbCrLf & vbCrLf & _
            "Erreur " & Err.Number & " : " & Err.Description, _
            vbExclamation, _
            "Retour de carton"

    End If

    ' Le gestionnaire est desarme : branche sur Sortie, une
    ' erreur ici y reviendrait sans fin.
    On Error Resume Next

    ' La touche Entree vaut validation, comme dans les
    ' autres popups.
    ActiverEspacePopup

    Application.OnKey "~", "ValiderPopupRetourCarton"
    Application.OnKey "{ENTER}", "ValiderPopupRetourCarton"

End Sub


Private Sub RemplirListe( _
    ByVal ws As Worksheet, _
    ByVal Cartons As Collection)

    Dim i As Long

    For i = 1 To NB_LIGNES_LISTE

        CartonsListe(i) = ""

        If i <= Cartons.Count Then

            CartonsListe(i) = CStr(Cartons(i))

            With ZoneLigne(ws, LigneDeLIndex(i))
                ' La valeur sur l'ancre : le reste de la
                ' zone n'est que de la fusion.
                .Cells(1, 1).Value = LibelleLigne(CartonsListe(i))
                .Interior.Color = CIEL
                .Font.Color = RGB(0, 0, 0)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(150, 150, 150)
            End With

        Else

            With ZoneLigne(ws, LigneDeLIndex(i))
                .ClearContents
                .Interior.Color = RGB(255, 255, 255)
                .Borders.LineStyle = xlNone
            End With

        End If

    Next i

End Sub


' Ce que la case montre : le camp, la couleur, le joueur
' quand il y en a un, et la minute du carton.
Private Function LibelleLigne( _
    ByVal Evenement As String) As String

    Dim Texte As String

    If CampCarton(Evenement) = CAMP_NOUS Then
        Texte = "Nous"
    Else
        Texte = "Adversaire"
    End If

    Texte = Texte & "  -  carton " & _
        CouleurLisible(CouleurCarton(Evenement))

    If JoueurCarton(Evenement) <> "" Then
        Texte = Texte & "  -  " & JoueurCarton(Evenement)
    End If

    LibelleLigne = Texte & "  -  " & _
        FormaterTemps(TempsCarton(Evenement))

End Function


' Une ligne de journal par carton choisi, toutes au meme
' temps video.
Public Sub ValiderPopupRetourCarton()

    Dim Choisis As Collection
    Dim Element As Variant

    If Not RetourOuverte Then Exit Sub

    Set Choisis = CartonsSelectionnes

    If Choisis.Count = 0 Then

        MsgBox _
            "Choisis le ou les cartons qui reviennent.", _
            vbExclamation, _
            "Retour de carton"

        Exit Sub

    End If

    ' Les confirmations sont toutes demandees avant la
    ' premiere ecriture : un refus laisse le journal intact
    ' plutot qu'a moitie rempli.
    For Each Element In Choisis

        If Not RetourConfirme(CStr(Element)) Then Exit Sub

    Next Element

    For Each Element In Choisis

        ' Le meme joueur que le carton : c'est ce qui apparie
        ' les deux lignes du journal.
        AjouterAction _
            JoueurCarton(CStr(Element)), _
            LibelleRetour( _
                CouleurCarton(CStr(Element)), _
                CampCarton(CStr(Element))), _
            RetourTemps

    Next Element

    InvaliderCacheCartons

    FermerPopupRetourCarton

End Sub


' Rouge et bleu ne reviennent pas : le premier est une
' exclusion definitive, le second ecarte le joueur pour le
' reste de la rencontre. La popup les propose quand meme,
' pour rattraper une saisie erronee, mais elle demande
' confirmation avant d'ecrire un retour qui n'existe pas
' dans les regles.
'
' Le bouton par defaut est Non : une validation faite trop
' vite ne doit pas suffire.
Private Function RetourConfirme( _
    ByVal Evenement As String) As Boolean

    Dim Motif As String

    Select Case CouleurCarton(Evenement)

        Case "ROUGE"

            Motif = "Un carton rouge est une exclusion " & _
                "d" & ChrW(233) & "finitive : le joueur ne " & _
                "revient normalement pas sur le terrain."

        Case "BLEU"

            Motif = "Un carton bleu " & ChrW(233) & "carte " & _
                "le joueur pour le reste de la rencontre."

        Case Else

            RetourConfirme = True
            Exit Function

    End Select

    ' La ligne concernee est rappelee : plusieurs cartons
    ' peuvent etre coches, il faut savoir lequel on confirme.
    RetourConfirme = ( _
        MsgBox( _
            Motif & vbCrLf & vbCrLf & _
            LibelleLigne(Evenement) & vbCrLf & vbCrLf & _
            "Enregistrer tout de m" & ChrW(234) & _
            "me son retour en jeu ?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "Retour inhabituel") = vbYes)

End Function


' Rien a journaliser : contrairement a une penalite, un
' retour annule n'a pas eu lieu.
Public Sub AnnulerPopupRetourCarton()

    If Not RetourOuverte Then Exit Sub

    FermerPopupRetourCarton

End Sub


Private Sub FermerPopupRetourCarton()

    Dim ws As Worksheet

    RetourOuverte = False

    ReprendreVideoApresPopup

    Application.OnKey "~"
    Application.OnKey "{ENTER}"

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_RETOUR)

    shSaisieVideo.Activate
    ws.Visible = xlSheetHidden

End Sub


' ---------------------------------------------------------
' Clics
' ---------------------------------------------------------

Public Sub ClicPopupRetourCarton(ByVal Target As Range)

    Dim ws As Worksheet

    If Target.Areas.Count > 1 Then Exit Sub

    ' Un clic sur une cellule fusionnee livre toute la
    ' plage : on ne garde que son ancre.
    If Target.Cells.Count > 1 Then

        If Not Target.MergeCells Then Exit Sub
        Set Target = Target.Cells(1, 1)

    End If

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_RETOUR)

    If Not Intersect(Target, ws.Range("RET_PLAY_PAUSE")) _
        Is Nothing Then

        BasculerLectureDepuisPopup
        Exit Sub

    End If

    If Not Intersect(Target, ws.Range("RET_VALIDER")) _
        Is Nothing Then

        ValiderPopupRetourCarton
        Exit Sub

    End If

    If Not Intersect(Target, ws.Range("RET_ANNULER")) _
        Is Nothing Then

        AnnulerPopupRetourCarton
        Exit Sub

    End If

    BasculerSelection ws, Target

End Sub


' Un clic allume ou eteint la ligne cliquee, sans toucher
' aux autres : plusieurs cartons peuvent revenir au meme
' moment, deux exclusions temporaires qui s'achevent
' ensemble par exemple.
Private Sub BasculerSelection( _
    ByVal ws As Worksheet, _
    ByVal Target As Range)

    Dim DejaChoisie As Boolean

    If Intersect(Target, ws.Range("RET_LISTE")) Is Nothing Then
        Exit Sub
    End If

    If Trim(CStr(Target.Value)) = "" Then Exit Sub

    DejaChoisie = (Target.Interior.Color = BLEU)

    On Error GoTo Sortie

    Application.EnableEvents = False

    With ZoneLigne(ws, Target.Row)

        If DejaChoisie Then

            .Interior.Color = CIEL
            .Font.Color = RGB(0, 0, 0)

        Else

            .Interior.Color = BLEU
            .Font.Color = RGB(255, 255, 255)

        End If

    End With

Sortie:

    Application.EnableEvents = True

    If Err.Number <> 0 Then

        MsgBox _
            "La s" & ChrW(233) & "lection n'a pas pu " & _
            ChrW(234) & "tre appliqu" & ChrW(233) & "e." & _
            vbCrLf & vbCrLf & _
            "Erreur " & Err.Number & " : " & Err.Description, _
            vbExclamation, _
            "Retour de carton"

    End If

End Sub


' Les cartons des lignes mises en evidence, dans l'ordre ou
' la liste les montre.
Private Function CartonsSelectionnes() As Collection

    Dim ws As Worksheet
    Dim Choisis As Collection
    Dim i As Long

    Set Choisis = New Collection
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_RETOUR)

    For i = 1 To NB_LIGNES_LISTE

        If CartonsListe(i) <> "" Then

            If ws.Cells(LigneDeLIndex(i), COL_PREMIERE) _
                .Interior.Color = BLEU Then

                Choisis.Add CartonsListe(i)

            End If

        End If

    Next i

    Set CartonsSelectionnes = Choisis

End Function


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function LigneDeLIndex(ByVal Index As Long) As Long

    LigneDeLIndex = PREMIERE_LIGNE + (Index - 1) * 2

End Function


Private Function ZoneLigne( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long) As Range

    Set ZoneLigne = ws.Range( _
        ws.Cells(Ligne, COL_PREMIERE), _
        ws.Cells(Ligne, COL_DERNIERE))

End Function


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
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_RETOUR)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = FEUILLE_POPUP_RETOUR

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePopup = ws

End Function
