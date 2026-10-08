Attribute VB_Name = "modTouchesMeleesPopup"
Option Explicit

' =========================================================
' TOUCHES ET MELEES - FABRICATION DES DEUX POPUPS
'
' A executer une fois depuis "Createur de match.xlsm".
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
' le tableau.
'
' Les choix se font au clic, sur des boutons poses en
' ligne : derouler une liste pour chaque champ coutait
' trop de temps pendant le depouillement. Seul le sauteur
' garde sa liste, la composition etant trop longue pour
' tenir en boutons.
'
' La valeur choisie ne s'ecrit pas dans le bouton mais
' dans une cellule a part, en colonne masquee : la saisie
' continue de la lire par son nom de plage, comme du temps
' des listes deroulantes.
' =========================================================

Public Const FEUILLE_POPUP_TO As String = "Popup touche"
Public Const FEUILLE_POPUP_ME As String = "Popup melee"

Private Const FEUILLE_TABLEAUX As String = "Touches Melees"

Private Const GRIS As Long = 15132390     ' RGB(230, 230, 230)
Private Const BLEU As Long = 12611584     ' RGB(0, 112, 192)
Private Const VERT As Long = 32768        ' RGB(0, 128, 0)
Private Const ROUGE As Long = 192         ' RGB(192, 0, 0)
Private Const SABLE As Long = 14083324    ' RGB(252, 228, 214)

' Disposition commune aux deux popups.
Private Const COL_LIBELLE As Long = 2
Private Const COL_PREMIER_BOUTON As Long = 3
Private Const LARGEUR_BOUTON As Long = 2

' Colonne masquee ou atterrissent les valeurs choisies.
Private Const COL_VALEURS As Long = 30

Private EtapeEnCours As String


' ---------------------------------------------------------
' Disposition des boutons
'
' Une entree par bouton, lue aussi bien pour dessiner la
' popup que pour reconnaitre le bouton clique :
'
'     nom de plage | ligne | rang | libelle | valeur
'
' Le rang est la position du bouton sur sa ligne, a partir
' de zero. Le libelle est ce qui s'affiche, la valeur ce
' qui part dans le tableau : les zones de terrain disent
' "en-but" des deux cotes, mais ecrivent "N en-but" ou
' "E en-but" selon le camp.
' ---------------------------------------------------------

Public Function OptionsTouche() As Variant

    Dim T As String

    T = _
        "POPUP_TO_ZONE_SAUT|11|0|Z0|Z0" & vbTab & _
        "POPUP_TO_ZONE_SAUT|11|1|Z1|Z1" & vbTab & _
        "POPUP_TO_ZONE_SAUT|11|2|Z1,5|Z1,5" & vbTab & _
        "POPUP_TO_ZONE_SAUT|11|3|Z2|Z2" & vbTab & _
        "POPUP_TO_ZONE_SAUT|11|4|Z3|Z3"

    T = T & vbTab & _
        "POPUP_TO_BALLON|13|0|Chaud|Chaud" & vbTab & _
        "POPUP_TO_BALLON|13|1|Froid|Froid"

    T = T & vbTab & ZonesTerrain("POPUP_TO_ZONE_TERRAIN", 16)

    T = T & vbTab & _
        "POPUP_TO_ALIGNEMENT|22|0|2|2" & vbTab & _
        "POPUP_TO_ALIGNEMENT|22|1|3|3" & vbTab & _
        "POPUP_TO_ALIGNEMENT|22|2|4|4" & vbTab & _
        "POPUP_TO_ALIGNEMENT|22|3|5|5" & vbTab & _
        "POPUP_TO_ALIGNEMENT|22|4|6|6" & vbTab & _
        "POPUP_TO_ALIGNEMENT|22|5|Complet|Complet"

    OptionsTouche = Split(T, vbTab)

End Function


Public Function OptionsMelee() As Variant

    Dim T As String

    T = ZonesTerrain("POPUP_ME_ZONE_LONGUEUR", 10)

    T = T & vbTab & _
        "POPUP_ME_ZONE_LARGEUR|16|0|Gauche|Gauche" & vbTab & _
        "POPUP_ME_ZONE_LARGEUR|16|1|Milieu|Milieu" & vbTab & _
        "POPUP_ME_ZONE_LARGEUR|16|2|Droit|Droit"

    T = T & vbTab & _
        "POPUP_ME_UTILISATION|18|0|Avants|Avants" & vbTab & _
        "POPUP_ME_UTILISATION|18|1|Trois-quart|Trois-quart" & vbTab & _
        "POPUP_ME_UTILISATION|18|2|Pied|Pied"

    OptionsMelee = Split(T, vbTab)

End Function


' Les onze zones du terrain, sur trois rangs : le notre,
' le milieu, celui d'en face. Le bouton ne porte que la
' distance, le camp etant donne par le libelle du rang ;
' la valeur ecrite, elle, reste celle du tableau.
Private Function ZonesTerrain( _
    ByVal NomPlage As String, _
    ByVal PremiereLigne As Long) As String

    Dim T As String
    Dim Distances As Variant
    Dim i As Long

    Distances = Array("en-but", "15m", "22m", "30m", "40m")

    For i = 0 To UBound(Distances)

        T = T & NomPlage & "|" & PremiereLigne & "|" & i & _
            "|" & Distances(i) & "|N " & Distances(i) & vbTab

    Next i

    T = T & NomPlage & "|" & (PremiereLigne + 2) & _
        "|0|50m|50m" & vbTab

    For i = 0 To UBound(Distances)

        T = T & NomPlage & "|" & (PremiereLigne + 4) & "|" & i & _
            "|" & Distances(i) & "|E " & Distances(i) & vbTab

    Next i

    ' Le dernier vbTab est de trop : il ouvrirait une
    ' entree vide au moment du Split.
    ZonesTerrain = Left$(T, Len(T) - 1)

End Function


' ---------------------------------------------------------
' Construction
' ---------------------------------------------------------

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

    If ModeSilencieux Then Exit Sub

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


Private Sub ConstruirePopupTouche()

    Dim ws As Worksheet

    Set ws = FeuillePopup(FEUILLE_POPUP_TO)

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws, "TOUCHE"
    EcrireContexte ws, "Lance pour"

    ' Le sauteur garde sa liste : la composition ne tient
    ' pas en boutons.
    EcrireChampListe ws, 9, "Sauteur", "POPUP_TO_SAUTEUR"

    EcrireLibelle ws, 11, "Zone de saut"
    EcrireLibelle ws, 13, "Ballon"

    EcrireSousTitre ws, 15, "ZONE TERRAIN"
    EcrireLibelle ws, 16, "Notre camp"
    EcrireLibelle ws, 18, "Milieu"
    EcrireLibelle ws, 20, "Camp adverse"

    EcrireLibelle ws, 22, "Alignement"

    DessinerOptions ws, OptionsTouche

    PoserValeurs ws, OptionsTouche

    EcrireChampTexte ws, 24, "Observations", "POPUP_TO_OBS"

    ' La saisie la repose a chaque ouverture, en suivant
    ' la composition : la poser ici donne une popup
    ' complete des sa construction.
    PoserListeSauteur

    EcrireBoutons ws, 26, "POPUP_TO_VALIDER", _
        "POPUP_TO_ANNULER"

    NommerPlage ws, "E4", "POPUP_TO_PLAY_PAUSE"
    NommerPlage ws, "C4:C7", "POPUP_TO_CONTEXTE"

    ' Ce que l'ouverture remet a blanc : les deux champs
    ' libres et les valeurs des boutons.
    NommerPlage ws, _
        AdresseNommee("POPUP_TO_SAUTEUR") & "," & _
        AdresseValeurs(ws, OptionsTouche) & "," & _
        AdresseNommee("POPUP_TO_OBS"), _
        "POPUP_TO_CHAMPS"

    MettreEnFormePopup ws, 26

End Sub


Private Sub ConstruirePopupMelee()

    Dim ws As Worksheet

    Set ws = FeuillePopup(FEUILLE_POPUP_ME)

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws, "MELEE"
    EcrireContexte ws, "Introduction pour"

    EcrireSousTitre ws, 9, "ZONE LONGUEUR"
    EcrireLibelle ws, 10, "Notre camp"
    EcrireLibelle ws, 12, "Milieu"
    EcrireLibelle ws, 14, "Camp adverse"

    EcrireLibelle ws, 16, "Zone largeur"
    EcrireLibelle ws, 18, "Utilisation"

    DessinerOptions ws, OptionsMelee

    PoserValeurs ws, OptionsMelee

    EcrireChampTexte ws, 20, "Observations", "POPUP_ME_OBS"

    EcrireBoutons ws, 22, "POPUP_ME_VALIDER", _
        "POPUP_ME_ANNULER"

    NommerPlage ws, "E4", "POPUP_ME_PLAY_PAUSE"
    NommerPlage ws, "C4:C7", "POPUP_ME_CONTEXTE"

    NommerPlage ws, _
        AdresseValeurs(ws, OptionsMelee) & "," & _
        AdresseNommee("POPUP_ME_OBS"), _
        "POPUP_ME_CHAMPS"

    MettreEnFormePopup ws, 22

End Sub


' ---------------------------------------------------------
' Boutons d'options
' ---------------------------------------------------------

Private Sub DessinerOptions( _
    ByVal ws As Worksheet, _
    ByVal Entrees As Variant)

    Dim i As Long

    For i = 0 To UBound(Entrees)

        ' Le fond suit l'ancre d'une cellule fusionnee,
        ' mais pas les bordures : tout se pose sur la zone
        ' entiere, sans quoi le bouton n'a de contour que
        ' sur sa premiere colonne.
        With ZoneBouton(ws, CStr(Entrees(i)))

            .Merge
            .Value = Champ(CStr(Entrees(i)), 3)
            .Interior.Color = SABLE
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)

        End With

    Next i

End Sub


' Une cellule par champ, dans la colonne masquee : c'est
' elle que la saisie lit, et c'est elle que le clic sur un
' bouton remplit.
Private Sub PoserValeurs( _
    ByVal ws As Worksheet, _
    ByVal Entrees As Variant)

    Dim i As Long
    Dim Vus As String
    Dim NomPlage As String

    For i = 0 To UBound(Entrees)

        NomPlage = Champ(CStr(Entrees(i)), 0)

        If InStr(1, Vus, "|" & NomPlage & "|") = 0 Then

            Vus = Vus & "|" & NomPlage & "|"

            NommerPlage ws, _
                ws.Cells( _
                    CLng(Champ(CStr(Entrees(i)), 1)), _
                    COL_VALEURS).Address, _
                NomPlage

        End If

    Next i

    ws.Columns(COL_VALEURS).Hidden = True

End Sub


' Les memes cellules, citees d'un trait : l'ouverture de
' la popup les vide toutes en une fois.
Private Function AdresseValeurs( _
    ByVal ws As Worksheet, _
    ByVal Entrees As Variant) As String

    Dim i As Long
    Dim Vus As String
    Dim NomPlage As String
    Dim Adresses As String

    For i = 0 To UBound(Entrees)

        NomPlage = Champ(CStr(Entrees(i)), 0)

        If InStr(1, Vus, "|" & NomPlage & "|") = 0 Then

            Vus = Vus & "|" & NomPlage & "|"

            If Adresses <> "" Then Adresses = Adresses & ","

            Adresses = Adresses & ws.Cells( _
                CLng(Champ(CStr(Entrees(i)), 1)), _
                COL_VALEURS).Address

        End If

    Next i

    AdresseValeurs = Adresses

End Function


Private Function ZoneBouton( _
    ByVal ws As Worksheet, _
    ByVal Entree As String) As Range

    Set ZoneBouton = CelluleOption(ws, Entree) _
        .Resize(1, LARGEUR_BOUTON)

End Function


Private Function CelluleOption( _
    ByVal ws As Worksheet, _
    ByVal Entree As String) As Range

    Set CelluleOption = ws.Cells( _
        CLng(Champ(Entree, 1)), _
        COL_PREMIER_BOUTON + _
            CLng(Champ(Entree, 2)) * LARGEUR_BOUTON)

End Function


' Un champ d'une entree, decoupee sur ses barres.
Private Function Champ( _
    ByVal Entree As String, _
    ByVal Rang As Long) As String

    Dim Morceaux As Variant

    Morceaux = Split(Entree, "|")

    If Rang > UBound(Morceaux) Then Exit Function

    Champ = Morceaux(Rang)

End Function


' ---------------------------------------------------------
' Clics
'
' Des cellules nommees, pas des formes : la palette de
' saisie fonctionne deja ainsi, et Worksheet_SelectionChange
' suffit a les detecter.
' ---------------------------------------------------------

Public Sub ClicPopupTouche(ByVal Target As Range)

    Dim ws As Worksheet
    Dim Cellule As Range

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_TO)

    Set Cellule = AncreCliquee(Target)
    If Cellule Is Nothing Then Exit Sub

    If Touche(ws, Cellule, "POPUP_TO_PLAY_PAUSE") Then
        BasculerLectureDepuisPopup
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_TO_VALIDER") Then
        ValiderPopupTouche
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_TO_ANNULER") Then
        AnnulerPopupTouche
        Exit Sub
    End If

    AppliquerClicOption ws, Cellule, OptionsTouche

End Sub


Public Sub ClicPopupMelee(ByVal Target As Range)

    Dim ws As Worksheet
    Dim Cellule As Range

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_ME)

    Set Cellule = AncreCliquee(Target)
    If Cellule Is Nothing Then Exit Sub

    If Touche(ws, Cellule, "POPUP_ME_PLAY_PAUSE") Then
        BasculerLectureDepuisPopup
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_ME_VALIDER") Then
        ValiderPopupMelee
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_ME_ANNULER") Then
        AnnulerPopupMelee
        Exit Sub
    End If

    AppliquerClicOption ws, Cellule, OptionsMelee

End Sub


' Un clic sur une cellule fusionnee livre toute la plage :
' on ne garde que son ancre, la seule que la disposition
' connaisse.
Private Function AncreCliquee( _
    ByVal Target As Range) As Range

    If Target.Areas.Count > 1 Then Exit Function

    If Target.Cells.Count > 1 Then

        If Not Target.MergeCells Then Exit Function
        Set AncreCliquee = Target.Cells(1, 1)
        Exit Function

    End If

    Set AncreCliquee = Target

End Function


Private Function Touche( _
    ByVal ws As Worksheet, _
    ByVal Cellule As Range, _
    ByVal NomPlage As String) As Boolean

    On Error Resume Next

    Touche = Not Intersect( _
        Cellule, ws.Range(NomPlage)) Is Nothing

    On Error GoTo 0

End Function


' Le bouton clique prend la couleur, ses voisins de champ
' la perdent, et la valeur part dans la cellule masquee.
' Reclique, il se rend : le champ repart vide.
Private Sub AppliquerClicOption( _
    ByVal ws As Worksheet, _
    ByVal Cellule As Range, _
    ByVal Entrees As Variant)

    Dim i As Long
    Dim Choisie As Long
    Dim NomPlage As String
    Dim DejaChoisi As Boolean

    Choisie = -1

    For i = 0 To UBound(Entrees)

        If CelluleOption(ws, CStr(Entrees(i))).Address = _
            Cellule.Address Then

            Choisie = i
            Exit For

        End If

    Next i

    If Choisie < 0 Then Exit Sub

    NomPlage = Champ(CStr(Entrees(Choisie)), 0)
    DejaChoisi = (Cellule.Interior.Color = BLEU)

    ' Les evenements sont rendus meme en cas d'erreur :
    ' coupes, plus aucun clic ne passerait ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    For i = 0 To UBound(Entrees)

        If Champ(CStr(Entrees(i)), 0) = NomPlage Then

            With ZoneBouton(ws, CStr(Entrees(i)))
                .Interior.Color = SABLE
                .Font.Color = RGB(0, 0, 0)
            End With

        End If

    Next i

    If DejaChoisi Then

        ws.Range(NomPlage).ClearContents

    Else

        With Cellule.Resize(1, LARGEUR_BOUTON)
            .Interior.Color = BLEU
            .Font.Color = RGB(255, 255, 255)
        End With

        ws.Range(NomPlage).Value = _
            Champ(CStr(Entrees(Choisie)), 4)

    End If

Sortie:

    Application.EnableEvents = True

End Sub


' Appelee a chaque ouverture : sans elle, les boutons de
' la touche precedente resteraient allumes.
Public Sub EffacerSelectionsOptions(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    If ws.Name = FEUILLE_POPUP_TO Then
        Entrees = OptionsTouche
    ElseIf ws.Name = FEUILLE_POPUP_ME Then
        Entrees = OptionsMelee
    Else
        Exit Sub
    End If

    On Error Resume Next

    For i = 0 To UBound(Entrees)

        With ZoneBouton(ws, CStr(Entrees(i)))
            .Interior.Color = SABLE
            .Font.Color = RGB(0, 0, 0)
        End With

    Next i

    On Error GoTo 0

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

        With ws.Cells(4 + i, 3).Resize(1, 2)
            .Merge
            .Interior.Color = GRIS
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

    Next i

    ' Lecture et pause sans quitter la popup : la video
    ' reste pilotable pendant la saisie.
    With ws.Cells(4, 5).Resize(1, 3)

        .Merge
        .Value = "PLAY / PAUSE"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub


Private Sub EcrireLibelle( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String)

    With ws.Cells(Ligne, COL_LIBELLE)
        .Value = Libelle
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

End Sub


' Les trois rangs du terrain vont ensemble : un intertitre
' les rassemble sous un meme nom.
Private Sub EcrireSousTitre( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Titre As String)

    With ws.Cells(Ligne, COL_LIBELLE)
        .Value = Titre
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = BLEU
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With

End Sub


' Le champ du sauteur : une cellule, sa liste posee plus
' tard par PoserListeSauteur, qui suit la composition.
Private Sub EcrireChampListe( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String, _
    ByVal NomPlage As String)

    EcrireLibelle ws, Ligne, Libelle

    With ws.Cells(Ligne, COL_PREMIER_BOUTON).Resize(1, 4)

        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter

    End With

    ' La zone entiere, pas son ancre : Excel refuse de
    ' vider une partie d'une cellule fusionnee, et
    ' l'ouverture de la popup s'arreterait la.
    NommerPlage ws, _
        ws.Cells(Ligne, COL_PREMIER_BOUTON) _
            .Resize(1, 4).Address, NomPlage

End Sub


' Le champ libre des observations, large comme la popup.
Private Sub EcrireChampTexte( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String, _
    ByVal NomPlage As String)

    EcrireLibelle ws, Ligne, Libelle

    With ws.Cells(Ligne, COL_PREMIER_BOUTON).Resize(1, 12)

        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlLeft

    End With

    ' Meme raison : la zone entiere, pas son ancre.
    NommerPlage ws, _
        ws.Cells(Ligne, COL_PREMIER_BOUTON) _
            .Resize(1, 12).Address, NomPlage

End Sub


Private Sub EcrireBoutons( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal NomValider As String, _
    ByVal NomAnnuler As String)

    With ws.Cells(Ligne, COL_LIBELLE).Resize(1, 2)
        .Merge
        .Value = "ANNULER"
        .Interior.Color = ROUGE
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(Ligne, COL_LIBELLE + 3).Resize(1, 4)
        .Merge
        .Value = "VALIDER  (Entree)"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    NommerPlage ws, _
        ws.Cells(Ligne, COL_LIBELLE).Address, NomAnnuler

    NommerPlage ws, _
        ws.Cells(Ligne, COL_LIBELLE + 3).Address, NomValider

End Sub


Private Sub MettreEnFormePopup( _
    ByVal ws As Worksheet, _
    ByVal LigneBoutons As Long)

    Dim Ligne As Long

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 16

    ' Douze colonnes de boutons : six options par ligne au
    ' plus large, l'alignement complet des touches.
    ws.Range(ws.Columns(COL_PREMIER_BOUTON), _
        ws.Columns(COL_PREMIER_BOUTON + 11)).ColumnWidth = 6.5

    ' Les rangs qui portent quelque chose respirent, les
    ' rangs intercalaires restent fins.
    For Ligne = 4 To LigneBoutons

        If Application.WorksheetFunction.CountA( _
            ws.Range( _
                ws.Cells(Ligne, 2), _
                ws.Cells(Ligne, COL_PREMIER_BOUTON + 11))) > 0 Then

            ws.Rows(Ligne).RowHeight = 20

        Else

            ws.Rows(Ligne).RowHeight = 8

        End If

    Next Ligne

    ws.Rows(LigneBoutons).RowHeight = 26

    ws.Cells.Font.Name = "Calibri"

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


' L'adresse d'une plage deja nommee : les champs libres
' sont fusionnes, et c'est leur etendue qui compte.
Private Function AdresseNommee( _
    ByVal NomPlage As String) As String

    AdresseNommee = _
        ThisWorkbook.Names(NomPlage).RefersToRange.Address

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
