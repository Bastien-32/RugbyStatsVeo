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

' La ligne 29 porte des secondes brutes, que le code relit
' pour savoir qui etait sur le terrain : elle reste
' masquee. La 30 en donne la lecture en minutes.
Private Const LIGNE_TEMPS_LISIBLE As Long = 30
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

' Minute la plus tardive acceptee : au-dela des 80, pour
' couvrir les arrets de jeu et une eventuelle prolongation.
Private Const MINUTE_MAX As Long = 110

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

    ws.Range("B5").Value = "Minute de remplacement"
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

    ' Lecture et pause sans quitter la popup : la video
    ' reste pilotable pendant la saisie.
    With ws.Range("E4")

        .Value = "PLAY / PAUSE"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    NommerPlage ws, "E4", "REMP_PLAY_PAUSE"

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
    ws.Columns("E").ColumnWidth = 16

    ws.Range( _
        ws.Columns(COL_LISTE_SORTANTS), _
        ws.Columns(COL_LISTE_ENTRANTS + 1)).Hidden = True

    ws.Cells.Font.Name = "Calibri"

    If ModeSilencieux Then Exit Sub

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

                    ' La marque porte son rang : seule la
                    ' premiere lettre nous interesse ici.
                    If Left(Marque, 1) = "S" Then
                        Presents.Remove Poste
                    ElseIf Left(Marque, 1) = "E" Then
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


' Poste occupant un emplacement du terrain a un instant
' donne. L'emplacement porte le numero du titulaire ; le
' remplacant qui entre en prend la place, et ainsi de
' suite si lui-meme est remplace.
Public Function OccupantEmplacement( _
    ByVal Emplacement As Long, _
    ByVal TempsVideo As Double) As String

    Dim ws As Worksheet
    Dim Colonnes As Variant
    Dim i As Long
    Dim Colonne As Long
    Dim Ligne As Long
    Dim Courant As String
    Dim Marque As String
    Dim Rang As String

    Set ws = ThisWorkbook.Sheets("Compo")

    Courant = CStr(Emplacement)

    Colonnes = ColonnesParMinute(ws)

    For i = 0 To UBound(Colonnes)

        If Colonnes(i) = "" Then Exit For

        Colonne = CLng(Split(CStr(Colonnes(i)), ":")(1))

        If TempsMouvement(ws, Colonne) <= TempsVideo Then

            ' On cherche le couple dont le sortant occupe
            ' l'emplacement suivi, puis on prend son
            ' entrant de meme rang.
            Rang = ""

            For Ligne = PREMIER_POSTE To DERNIER_POSTE

                Marque = UCase(Trim(CStr( _
                    ws.Cells(Ligne, Colonne).Value)))

                If Left(Marque, 1) = "S" Then

                    If Trim(CStr(ws.Cells(Ligne, _
                        COL_POSTE).Value)) = Courant Then

                        Rang = Mid(Marque, 2)
                        Exit For

                    End If

                End If

            Next Ligne

            If Rang <> "" Then

                For Ligne = PREMIER_POSTE To DERNIER_POSTE

                    Marque = UCase(Trim(CStr( _
                        ws.Cells(Ligne, Colonne).Value)))

                    If Marque = "E" & Rang Then

                        Courant = Trim(CStr(ws.Cells( _
                            Ligne, COL_POSTE).Value))

                        Exit For

                    End If

                Next Ligne

            End If

        End If

    Next i

    OccupantEmplacement = Courant

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
    SuspendreVideoPourPopup TempsVideo

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)

    ws.Visible = xlSheetVisible
    ws.Activate

    ' Les evenements sont rendus quoi qu'il arrive : une
    ' erreur ici bloquerait toute la palette ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    ' Format texte impose : sans lui, Excel lit "64:48"
    ' comme 64 heures 48 et affiche 16:48.
    ActiverEspacePopup

    ws.Range("REMP_TEMPS").NumberFormat = "@"
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
    Dim Sortants() As String
    Dim Entrants() As String

    If Not RempOuverte Then Exit Sub

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_REMP)

    ' Une cellule vide rend Empty, et IsNumeric(Empty)
    ' vaut True : sans ce passage par le texte, une minute
    ' oubliee s'ecrirait en zero.
    Minute = Trim(CStr(ws.Range("REMP_MINUTE").Value))

    If Minute = "" Or Not IsNumeric(Minute) Then

        MsgBox _
            "Renseigne la minute de remplacement avant " & _
            "de valider.", _
            vbExclamation, _
            "Minute manquante"

        ws.Range("REMP_MINUTE").Select
        Exit Sub

    End If

    If CDbl(Minute) < 0 Or CDbl(Minute) > MINUTE_MAX Then

        MsgBox _
            "La minute doit etre comprise entre 0 et " & _
            MINUTE_MAX & "." & vbCrLf & vbCrLf & _
            "Valeur lue : " & Minute, _
            vbExclamation, _
            "Minute hors limites"

        ws.Range("REMP_MINUTE").Select
        Exit Sub

    End If

    ' Une ligne a moitie remplie est une erreur de saisie,
    ' pas une ligne vide : il vaut mieux le dire que
    ' l'ignorer en silence.
    If Not LignesCompletes(ws) Then Exit Sub

    ReDim Sortants(1 To NB_LIGNES_SAISIE)
    ReDim Entrants(1 To NB_LIGNES_SAISIE)

    For i = 1 To NB_LIGNES_SAISIE

        Sortant = PosteDeLEtiquette(CStr( _
            ws.Range("REMP_SORTANTS").Cells(i, 1).Value))

        Entrant = PosteDeLEtiquette(CStr( _
            ws.Range("REMP_ENTRANTS").Cells(i, 1).Value))

        If Sortant <> "" And Entrant <> "" Then

            Ecrits = Ecrits + 1
            Sortants(Ecrits) = Sortant
            Entrants(Ecrits) = Entrant

        End If

    Next i

    If Ecrits > 0 Then
        EcrireMouvements CLng(Minute), Sortants, Entrants, Ecrits
    End If

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


' Chaque ligne veut ses deux joueurs, ou aucun.
Private Function LignesCompletes( _
    ByVal ws As Worksheet) As Boolean

    Dim i As Long
    Dim Sortant As String
    Dim Entrant As String

    For i = 1 To NB_LIGNES_SAISIE

        Sortant = Trim(CStr( _
            ws.Range("REMP_SORTANTS").Cells(i, 1).Value))

        Entrant = Trim(CStr( _
            ws.Range("REMP_ENTRANTS").Cells(i, 1).Value))

        If (Sortant = "") <> (Entrant = "") Then

            MsgBox _
                "La ligne " & i & " est incomplete." & _
                vbCrLf & vbCrLf & _
                "Un remplacement demande un sortant et " & _
                "un entrant. Laisse la ligne entierement " & _
                "vide si elle ne sert pas.", _
                vbExclamation, _
                "Remplacement incomplet"

            ws.Range( _
                IIf(Sortant = "", "REMP_SORTANTS", _
                    "REMP_ENTRANTS")).Cells(i, 1).Select

            Exit Function

        End If

    Next i

    LignesCompletes = True

End Function


Public Sub AnnulerPopupRemplacement()

    If Not RempOuverte Then Exit Sub

    FermerPopupRemplacement

End Sub


' La colonne neuve imite le tableau voisin.
'
' Copier le format de la colonne modele ne suffit pas :
' les bandes viennent du style du tableau, applique par
' le ListObject et non porte par les cellules. On lit
' donc la couleur reellement affichee, ligne par ligne.
Private Sub HabillerColonne( _
    ByVal ws As Worksheet, _
    ByVal Colonne As Long)

    Dim Ligne As Long
    Dim Modele As Range
    Dim Cible As Range

    For Ligne = LIGNE_ENTETE To DERNIER_POSTE

        Set Modele = ws.Cells(Ligne, COL_MODELE)
        Set Cible = ws.Cells(Ligne, Colonne)

        On Error Resume Next

        Cible.Interior.Color = _
            Modele.DisplayFormat.Interior.Color

        Cible.Font.Color = _
            Modele.DisplayFormat.Font.Color

        Cible.Font.Bold = Modele.DisplayFormat.Font.Bold

        On Error GoTo 0

        With Cible.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(150, 150, 150)
        End With

    Next Ligne

End Sub


' Tous les mouvements d'une meme validation tiennent dans
' une seule colonne : ils ont lieu a la meme minute.
'
' Le tableau COMPO s'arrete a la colonne du temps de jeu.
' Sans couper l'extension automatique, Excel l'etendrait
' jusqu'ici et renommerait les minutes en doublon.
Private Sub EcrireMouvements( _
    ByVal Minute As Long, _
    ByRef Sortants() As String, _
    ByRef Entrants() As String, _
    ByVal Nombre As Long)

    Dim ws As Worksheet
    Dim Colonne As Long
    Dim Ligne As Long
    Dim Poste As String
    Dim i As Long
    Dim EtatExtension As Boolean

    Set ws = ThisWorkbook.Sheets("Compo")

    EtatExtension = _
        Application.AutoCorrect.AutoExpandListRange

    On Error GoTo Sortie

    Application.AutoCorrect.AutoExpandListRange = False

    Colonne = PREMIERE_COL_MOUVEMENT

    Do While Trim(CStr( _
        ws.Cells(LIGNE_ENTETE, Colonne).Value)) <> ""
        Colonne = Colonne + 1
    Loop

    HabillerColonne ws, Colonne

    ws.Columns(Colonne).ColumnWidth = 6

    ws.Cells(LIGNE_ENTETE, Colonne).Value = Minute
    ws.Cells(LIGNE_ENTETE, Colonne).Font.Bold = True
    ws.Cells(LIGNE_ENTETE, Colonne).HorizontalAlignment = _
        xlCenter

    ws.Cells(LIGNE_TEMPS_VIDEO, Colonne).Value = _
        RempTempsVideo

    EcrireTempsLisible ws, Colonne, RempTempsVideo

    For Ligne = PREMIER_POSTE To DERNIER_POSTE

        Poste = Trim(CStr( _
            ws.Cells(Ligne, COL_POSTE).Value))

        For i = 1 To Nombre

            ' Le rang du couple accompagne la marque :
            ' une colonne peut porter trois sorties et
            ' trois entrees, et rien d'autre ne dirait
            ' qui remplace qui.
            If Poste = Sortants(i) Then
                ws.Cells(Ligne, Colonne).Value = "S" & i
            ElseIf Poste = Entrants(i) Then
                ws.Cells(Ligne, Colonne).Value = "E" & i
            End If

        Next i

        ws.Cells(Ligne, Colonne).HorizontalAlignment = _
            xlCenter

    Next Ligne

    ws.Rows(LIGNE_TEMPS_VIDEO).Hidden = True

Sortie:

    Application.AutoCorrect.AutoExpandListRange = _
        EtatExtension

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

        If Left(Marque, 1) = "S" And SurLeTerrain Then

            Total = Total + (Minute - Entree)
            SurLeTerrain = False

        ElseIf Left(Marque, 1) = "E" And _
            Not SurLeTerrain Then

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

    ReprendreVideoApresPopup

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

    If Not Intersect(Target, ws.Range("REMP_PLAY_PAUSE")) _
        Is Nothing Then

        BasculerLectureDepuisPopup
        Exit Sub

    End If

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


' =========================================================
' Le temps du mouvement, en minutes et secondes, sous la
' colonne qui le porte.
'
' La valeur est une fraction de jour : le format [m]:ss
' cumule alors les minutes au lieu de les reporter en
' heures, et la 64e minute ne devient pas la 4e.
' =========================================================

Private Sub EcrireTempsLisible( _
    ByVal ws As Worksheet, _
    ByVal Colonne As Long, _
    ByVal Secondes As Double)

    With ws.Cells(LIGNE_TEMPS_LISIBLE, Colonne)

        .Value = Secondes / 86400
        .NumberFormat = "[m]:ss"
        .HorizontalAlignment = xlCenter
        .Font.Bold = False

    End With

    ws.Rows(LIGNE_TEMPS_LISIBLE).Hidden = False

End Sub

' =========================================================
' PALETTE DES QUINZE PRESENTS
'
' Les boutons ne portent plus les postes 1 a 22 mais les
' quinze emplacements du terrain. Apres le remplacement du
' 10 par le 22, l'emplacement du 10 affiche 22 et le nom
' du remplacant.
'
' Appelee a chaque battement du chrono : elle ne reecrit
' que ce qui a change, sinon la palette clignoterait une
' fois par seconde.
' =========================================================

Public Sub RafraichirPalettePresents()

    Dim Temps As Double

    Temps = GetTimeVideo()

    If Temps < 0 Then Temps = 0

    EcrirePalettePresents Temps

End Sub


Public Sub EcrirePalettePresents(ByVal TempsVideo As Double)

    Dim ws As Worksheet
    Dim i As Long
    Dim Poste As String
    Dim Nom As String
    Dim Texte As String
    Dim Cellule As Range
    Dim EtatEvenements As Boolean

    Set ws = shSaisieVideo

    EtatEvenements = Application.EnableEvents

    On Error GoTo Sortie

    Application.EnableEvents = False

    For i = 1 To NB_TITULAIRES

        Set Cellule = Nothing

        On Error Resume Next
        Set Cellule = ws.Range("BTN_JO_" & i)
        On Error GoTo Sortie

        If Not Cellule Is Nothing Then

            Poste = OccupantEmplacement(i, TempsVideo)
            Nom = NomDeFamille(GetPlayerName(Poste))

            If Nom = "" Then
                Texte = Poste
            Else
                Texte = Poste & vbLf & Nom
            End If

            ' Rien n'est ecrit si rien n'a change : le
            ' rafraichissement passe chaque seconde.
            If CStr(Cellule.Value) <> Texte Then

                Cellule.Value = Texte
                HabillerBoutonJoueur Cellule, Poste, Nom

            End If

        End If

    Next i

Sortie:

    Application.EnableEvents = EtatEvenements

End Sub


' Le numero en gros, le nom en petit dessous.
Private Sub HabillerBoutonJoueur( _
    ByVal Cellule As Range, _
    ByVal Poste As String, _
    ByVal Nom As String)

    On Error Resume Next

    With Cellule

        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        If Nom = "" Then

            ' Sans nom, la cellule ne porte qu'un nombre :
            ' Characters n'aurait pas de prise dessus.
            .Font.Size = 18
            .Font.Bold = True

        Else

            .Characters(1, Len(Poste)).Font.Size = 18
            .Characters(1, Len(Poste)).Font.Bold = True

            .Characters(Len(Poste) + 2, Len(Nom)).Font.Size = 9
            .Characters(Len(Poste) + 2, Len(Nom)).Font.Bold = False

        End If

    End With

End Sub


' Les mots en majuscules du debut : "ABADIE Quentin"
' donne "ABADIE", "LOUREIRO LABAZUY Antonio" donne les
' deux premiers.
Private Function NomDeFamille( _
    ByVal NomComplet As String) As String

    Dim Mots() As String
    Dim Mot As Variant
    Dim Resultat As String

    If Trim(NomComplet) = "" Then Exit Function

    Mots = Split(Trim(NomComplet), " ")

    For Each Mot In Mots

        If CStr(Mot) <> UCase(CStr(Mot)) Then Exit For

        If Resultat = "" Then
            Resultat = CStr(Mot)
        Else
            Resultat = Resultat & " " & CStr(Mot)
        End If

    Next Mot

    NomDeFamille = Resultat

End Function
