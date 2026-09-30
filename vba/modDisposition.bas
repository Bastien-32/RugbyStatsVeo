Attribute VB_Name = "modDisposition"
Option Explicit

' =========================================================
' DISPOSITION DE LA FEUILLE DE SAISIE VIDEO
'
' La feuille n'est plus remaniee composant par composant :
' elle est redessinee. Chaque mode a sa table, qui donne
' pour chaque element son adresse, sa fusion et son
' libelle ; le style vient de Parametres, comme partout
' ailleurs.
'
' Redessiner plutot que deplacer evite la derive qui
' guettait l'ancienne methode : un couper-coller de
' l'utilisateur emportait une plage nommee, et le nom ne
' designait plus le bouton qu'il etait cense designer.
' Ici les noms sont recrees a chaque bascule.
'
' La palette des joueurs ne porte plus que les quinze
' presents sur le terrain : voir modRemplacements, qui
' sait qui occupe chaque poste a un instant donne.
' =========================================================

Private Const MODE_BANDEAU As String = "BANDEAU"
Private Const MODE_PLEIN_ECRAN As String = "PLEIN_ECRAN"

' Couleurs relevees sur les maquettes, pour les elements
' qui n'ont pas de style nomme dans Parametres.
Private Const MARINE As Long = 7884319      ' RGB(31,78,120)
Private Const VERT_TITRE As Long = 3057485  ' RGB(77,167,46)
Private Const BLANC As Long = 16777215

' Largeur des colonnes etroites, qui forment la trame.
Private Const ETROITE As Double = 1.83


' =========================================================
' POINTS D'ENTREE
' =========================================================

Public Sub BasculerDispositionSaisieVideo()

    Select Case DetecterDispositionSaisieVideo(shSaisieVideo)

        Case MODE_BANDEAU
            PasserEnModePleinEcran

        Case Else
            PasserEnModeBandeau

    End Select

End Sub


Public Sub PasserEnModePleinEcran()

    ConstruireDisposition MODE_PLEIN_ECRAN

End Sub


Public Sub PasserEnModeBandeau()

    ConstruireDisposition MODE_BANDEAU

End Sub


' Le mode se lit a la position du titre de la palette des
' actions : a droite en plein ecran, en haut en bandeau.
Private Function DetecterDispositionSaisieVideo( _
    ByVal ws As Worksheet) As String

    Dim Cible As Range

    On Error Resume Next
    Set Cible = ws.Range("TITRE_PALETTE_ACTIONS")
    On Error GoTo 0

    If Cible Is Nothing Then
        DetecterDispositionSaisieVideo = ""
    ElseIf Cible.Column > 20 Then
        DetecterDispositionSaisieVideo = MODE_PLEIN_ECRAN
    Else
        DetecterDispositionSaisieVideo = MODE_BANDEAU
    End If

End Function


' =========================================================
' CONSTRUCTION
' =========================================================

Private Sub ConstruireDisposition(ByVal Mode As String)

    Dim ws As Worksheet
    Dim Entrees As Variant
    Dim i As Long
    Dim EtatEvenements As Boolean
    Dim EtatAffichage As Boolean

    Set ws = shSaisieVideo

    EtatEvenements = Application.EnableEvents
    EtatAffichage = Application.ScreenUpdating

    On Error GoTo Sortie

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' La zone de travail repart vierge : les fusions d'un
    ' mode n'ont rien a faire dans l'autre.
    With ws.Range("A1:CJ60")
        .UnMerge
        .Clear
        .Interior.Color = BLANC
    End With

    If Mode = MODE_PLEIN_ECRAN Then
        Entrees = ComposantsPleinEcran
    Else
        Entrees = ComposantsBandeau
    End If

    For i = 0 To UBound(Entrees)
        EcrireComposant ws, CStr(Entrees(i))
    Next i

    EcrireTitreMatch ws, Mode
    NommerZonesTechniques ws, Mode
    PlacerBoutonSwitch ws, Mode

    AppliquerDimensions ws, Mode
    DessinerCadres ws, Mode

    ' Apres les dimensions : les formes se centrent sur la
    ' hauteur des lignes, qui vient seulement d'etre fixee.
    PositionnerFormesVideo ws, Mode

    ' Les quinze presents, et le titre du match.
    RafraichirPalettePresents

    ' Select exige une feuille active : la bascule peut
    ' etre lancee depuis un autre onglet.
    ws.Activate
    ws.Range("A1").Select

Sortie:

    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    If Err.Number <> 0 Then

        MsgBox _
            "La disposition n'a pas pu " & ChrW(234) & _
            "tre construite." & vbCrLf & vbCrLf & _
            "Erreur " & Err.Number & " : " & _
            Err.Description, _
            vbExclamation, _
            "Saisie vid" & ChrW(233) & "o"

    End If

End Sub


' ---------------------------------------------------------
' Le bouton de connexion, la pastille et le texte d'etat
' sont aussi des formes flottantes.
'
' Elles se placent par rapport au chrono : les mesures
' valent donc pour les deux dispositions, puisque le bloc
' de gestion video garde la meme forme dans chacune.
' ---------------------------------------------------------

Private Sub PositionnerFormesVideo( _
    ByVal ws As Worksheet, _
    ByVal Mode As String)

    Dim Ligne As Long
    Dim Hauteur As Double

    On Error GoTo Sortie

    ' Ligne du bloc de gestion video, relevee sur les
    ' maquettes : les trois formes s'y alignent.
    If Mode = MODE_PLEIN_ECRAN Then
        Ligne = 29
    Else
        Ligne = 51
    End If

    Hauteur = ws.Rows(Ligne).Height

    CentrerForme ws, "BTN_CONNECTER_VIDEO", _
        Ligne, 3, 139, 55, Hauteur

    ' La zone d'etat se place par rapport au bouton, non
    ' par rapport a une colonne : la trame change d'un
    ' mode a l'autre, le bouton non.
    PlacerEtatConnexion ws, Hauteur

    ' La pastille s'aligne sur la premiere ligne de texte
    ' de la zone d'etat, qui ecrit en haut de son cadre :
    ' centree sur la ligne, elle tomberait plus bas que
    ' le mot qu'elle accompagne.
    AlignerPastilleSurEtat ws

Sortie:

End Sub


Private Sub PlacerEtatConnexion( _
    ByVal ws As Worksheet, _
    ByVal HauteurLigne As Double)

    Dim Bouton As Shape
    Dim Etat As Shape

    On Error GoTo Sortie

    Set Bouton = ws.Shapes("BTN_CONNECTER_VIDEO")
    Set Etat = ws.Shapes("ETAT_CONNEXION_VIDEO")

    With Etat

        .LockAspectRatio = msoFalse
        .Width = 230
        .Height = 63
        .Left = Bouton.Left + Bouton.Width + 34
        .Top = Bouton.Top + (Bouton.Height - .Height) / 2

    End With

Sortie:

End Sub


Private Sub AlignerPastilleSurEtat(ByVal ws As Worksheet)

    Dim Etat As Shape
    Dim Pastille As Shape

    On Error GoTo Sortie

    Set Etat = ws.Shapes("ETAT_CONNEXION_VIDEO")
    Set Pastille = ws.Shapes("PASTILLE_CONNEXION_VIDEO")

    With Pastille

        .LockAspectRatio = msoTrue
        .Width = 13.75
        .Height = 13.68
        .Left = Etat.Left - 14
        .Top = Etat.Top + 8

    End With

Sortie:

End Sub


' La forme est centree sur la hauteur de la ligne : posee
' sur son bord haut, elle en deborderait vers le bas.
Private Sub CentrerForme( _
    ByVal ws As Worksheet, _
    ByVal NomForme As String, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long, _
    ByVal Largeur As Double, _
    ByVal HauteurForme As Double, _
    ByVal HauteurLigne As Double)

    Dim shp As Shape
    Dim Cellule As Range

    On Error GoTo Sortie

    Set shp = ws.Shapes(NomForme)
    Set Cellule = ws.Cells(Ligne, Colonne)

    With shp

        .LockAspectRatio = msoFalse
        .Width = Largeur
        .Height = HauteurForme
        .Left = Cellule.Left
        .Top = Cellule.Top + (HauteurLigne - HauteurForme) / 2

    End With

Sortie:

End Sub


' ---------------------------------------------------------
' Le bouton de bascule est une forme flottante : rien ne
' la deplace quand la feuille est redessinee, il faut donc
' la reposer sur la cellule voulue.
'
' Il annonce la disposition vers laquelle il mene, non
' celle ou l'on se trouve.
' ---------------------------------------------------------

Private Sub PlacerBoutonSwitch( _
    ByVal ws As Worksheet, _
    ByVal Mode As String)

    Dim shp As Shape
    Dim Ancre As Range

    On Error Resume Next
    Set shp = ws.Shapes("BTN_SWITCH_MODE")
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    If Mode = MODE_PLEIN_ECRAN Then
        Set Ancre = ws.Range("BA1:BD1")
        shp.TextFrame2.TextRange.Text = "Mode bandeau"
    Else
        Set Ancre = ws.Range("AP1")
        shp.TextFrame2.TextRange.Text = "Mode plein ecran"
    End If

    shp.Left = Ancre.Left
    shp.Top = Ancre.Top

    ' Le bouton tient dans sa plage : sans cela il garde
    ' la largeur qu'il avait dans l'autre mode.
    If Mode = MODE_PLEIN_ECRAN Then
        shp.LockAspectRatio = msoFalse
        shp.Width = Ancre.Width
    End If

End Sub


' ---------------------------------------------------------
' Le bandeau de titre, qui nomme la rencontre.
'
' La formule lit les plages MATCH_* de la compo : elle
' suit donc l'en-tete du match sans qu'on ait a la
' reecrire.
' ---------------------------------------------------------

Private Sub EcrireTitreMatch( _
    ByVal ws As Worksheet, _
    ByVal Mode As String)

    Dim Zone As Range
    Dim Formule As String

    If Mode = MODE_PLEIN_ECRAN Then
        Set Zone = ws.Range("B1:AY1")
    Else
        Set Zone = ws.Range("B1:AM1")
    End If

    Zone.Merge

    Formule = "=~Match ~&" & _
        "IF(MATCH_PHASE=~Aller~,~A~," & _
        "IF(MATCH_PHASE=~Retour~,~R~," & _
        "IF(MATCH_PHASE=~Amical~,~AMI~,~PF~)))" & _
        "&~ ~&MATCH_JOURNEE&" & _
        "IF(MATCH_LIEU=~Ext@rieur~," & _
        "~ : ~&MATCH_ADV&~ / ~&MATCH_NOUS," & _
        "~ : ~&MATCH_NOUS&~ / ~&MATCH_ADV)"

    Formule = Replace(Formule, "~", Chr(34))
    Formule = Replace(Formule, "@", ChrW(233))

    With Zone.Cells(1, 1)

        .Formula = Formule
        .Interior.Color = VERT_TITRE
        .Font.Color = BLANC
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub


' ---------------------------------------------------------
' Deux plages sans bouton : la case ou la selection vient
' se garer, et la zone de la palette des joueurs, que la
' feuille interroge a chaque clic.
' ---------------------------------------------------------

Private Sub NommerZonesTechniques( _
    ByVal ws As Worksheet, _
    ByVal Mode As String)

    NommerComposant ws, "BTN_PARKING", "A1"

    If Mode = MODE_PLEIN_ECRAN Then
        NommerComposant ws, "PALETTE_JOUEURS", "C13:AX21"
    Else
        NommerComposant ws, "PALETTE_JOUEURS", "C35:AL43"
    End If

End Sub


' ---------------------------------------------------------
' Une entree : NOM|adresse|fusion|libelle
' ---------------------------------------------------------

Private Sub EcrireComposant( _
    ByVal ws As Worksheet, _
    ByVal Entree As String)

    Dim Champs() As String
    Dim Nom As String
    Dim Adresse As String
    Dim Fusion As String
    Dim Libelle As String
    Dim Cible As Range
    Dim Style As String

    Champs = Split(Entree, "|")

    If UBound(Champs) < 2 Then Exit Sub

    Nom = Champs(0)
    Adresse = Champs(1)
    Fusion = Champs(2)
    Libelle = ""

    If UBound(Champs) >= 3 Then Libelle = Champs(3)

    If Fusion <> "" Then
        Set Cible = ws.Range(Fusion)
        Cible.Merge
    Else
        Set Cible = ws.Range(Adresse)
    End If

    If Libelle <> "" Then Cible.Cells(1, 1).Value = Libelle

    Style = StyleDe(Nom)

    If Style <> "" Then

        ApplyFormat Cible.Cells(1, 1), _
            shParametres.Range(Style)

    Else

        HabillerSansStyle Cible.Cells(1, 1), Nom

    End If

    With Cible.Cells(1, 1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    NommerComposant ws, Nom, Adresse

End Sub


' Les libelles de famille et les titres de palette n'ont
' pas de style dans Parametres : leurs couleurs sont
' celles relevees sur la maquette.
Private Sub HabillerSansStyle( _
    ByVal Cellule As Range, _
    ByVal Nom As String)

    If Left(Nom, 4) = "LBL_" Then

        If Nom = "LBL_POSSESSION" Or Nom = "LBL_MI_TEMPS" _
            Or Nom = "LBL_POSITION_VIDEO" Then

            Cellule.Font.Bold = True
            Exit Sub

        End If

        Cellule.Interior.Color = MARINE
        Cellule.Font.Color = BLANC
        Cellule.Font.Bold = True
        Cellule.Font.Size = 10

    ElseIf Left(Nom, 6) = "TITRE_" Then

        Cellule.Interior.Color = VERT_TITRE
        Cellule.Font.Color = BLANC
        Cellule.Font.Bold = True
        Cellule.Font.Size = 10

    End If

End Sub


Private Sub NommerComposant( _
    ByVal ws As Worksheet, _
    ByVal Nom As String, _
    ByVal Adresse As String)

    On Error Resume Next
    ThisWorkbook.Names(Nom).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=Nom, _
        RefersTo:=ws.Range(Adresse)

End Sub


Private Function StyleDe(ByVal Nom As String) As String

    Select Case Nom

        Case "BTN_ADV"
            StyleDe = "STYLE_BTN_POSSESSION"

        Case "BTN_ARRACHAGE"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_ARRET_DU_JEU"
            StyleDe = "STYLE_BTN_ARRET_JEU"

        Case "BTN_ATT_NULLE"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_AVANCEE"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_CARTON_BLANC"
            StyleDe = "STYLE_CARTON_BLANC"

        Case "BTN_CARTON_BLEU"
            StyleDe = "STYLE_CARTON_BLEU"

        Case "BTN_CARTON_JAUNE"
            StyleDe = "STYLE_CARTON_JAUNE"

        Case "BTN_CARTON_ROUGE"
            StyleDe = "STYLE_BTN_CHRONO_RESET"

        Case "BTN_CF_CONTRE_ADV"
            StyleDe = "STYLE_BTN_PEN"

        Case "BTN_CF_CONTRE_NOUS"
            StyleDe = "STYLE_BTN_PEN"

        Case "BTN_CHRONO_MINUS_5_SEC"
            StyleDe = "STYLE_BTN_CHRONO_TMP"

        Case "BTN_CHRONO_PLAY_PAUSE"
            StyleDe = "STYLE_BTN_CHRONO_TMP"

        Case "BTN_CHRONO_PLUS_5_SEC"
            StyleDe = "STYLE_BTN_CHRONO_TMP"

        Case "BTN_CHRONO_RESET"
            StyleDe = "STYLE_BTN_CHRONO_RESET"

        Case "BTN_CONTRE_RUCK"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_EN_AVANT"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_FRANCHISSEMENT"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_GRATTAGE"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_JEU_AU_PIED"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_JO_1"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_10"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_11"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_12"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_13"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_14"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_15"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_2"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_3"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_4"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_5"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_6"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_7"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_8"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_9"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_COLLECTIF"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_JO_NON_IDENTIFIE"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

        Case "BTN_ME_G"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_ME_P"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_MT1"
            StyleDe = "STYLE_BTN_MT"

        Case "BTN_MT2"
            StyleDe = "STYLE_BTN_MT"

        Case "BTN_NOUS"
            StyleDe = "STYLE_BTN_POSSESSION"

        Case "BTN_PEN_CONTRE_ADV"
            StyleDe = "STYLE_BTN_PEN"

        Case "BTN_PEN_CONTRE_NOUS"
            StyleDe = "STYLE_BTN_PEN"

        Case "BTN_PL_A2"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_PL_HAUT"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_PL_NORMAL"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_PL_OFF"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_PL_RATE"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_PTS_DROP"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_PTS_ESSAI"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_PTS_ESSAI_PEN"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_PTS_PENALITE"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_PTS_TRANSFO"
            StyleDe = "STYLE_BTN_PTS"

        Case "BTN_RECEPTION_ADV"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_RECEPTION_NOUS"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_REMPLACEMENT"
            StyleDe = "STYLE_BTN_POSSESSION"

        Case "BTN_REPRISE_A_LA_MAIN"
            StyleDe = "STYLE_BTN_REPRISE"

        Case "BTN_REPRISE_MELEE"
            StyleDe = "STYLE_BTN_REPRISE"

        Case "BTN_REPRISE_RENVOI"
            StyleDe = "STYLE_BTN_REPRISE"

        Case "BTN_REPRISE_TOUCHE"
            StyleDe = "STYLE_BTN_REPRISE"

        Case "BTN_TO_G"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_TO_P"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_TO_PAS_DROIT"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "BTN_TURNOVER"
            StyleDe = "STYLE_BTN_PL_TOUCHES"

        Case "CELL_CHRONO_VIDEO"
            StyleDe = "STYLE_BTN_JOUEUR_INACTIF"

    End Select

End Function


' =========================================================
' DIMENSIONS
' =========================================================

Private Sub AppliquerDimensions( _
    ByVal ws As Worksheet, _
    ByVal Mode As String)

    Dim i As Long

    ws.Cells.RowHeight = 13

    If Mode = MODE_PLEIN_ECRAN Then

        ' Trame etroite a gauche, colonnes larges a droite
        ' pour la palette des actions.
        ' Colonnes B a AY : la trame la plus fine, celle
        ' qui donne sa souplesse au placement.
        ws.Range(ws.Columns(1), ws.Columns(52)) _
            .ColumnWidth = 1

        ws.Range(ws.Columns(2), ws.Columns(51)) _
            .ColumnWidth = 0.8

        ws.Columns(53).ColumnWidth = ETROITE
        ws.Columns(54).ColumnWidth = 11.66
        ws.Columns(68).ColumnWidth = 12

        ' Palette des actions : une colonne de separation
        ' etroite entre deux colonnes de boutons.
        For i = 55 To 71 Step 2
            ws.Columns(i).ColumnWidth = 1
        Next i

        For i = 56 To 70 Step 2
            If i <> 68 Then ws.Columns(i).ColumnWidth = 10
        Next i

    ws.Rows(12).RowHeight = 15
    ws.Rows(24).RowHeight = 15
    ws.Rows(26).RowHeight = 15
    ws.Range(ws.Rows(33), ws.Rows(53)).RowHeight = 15
    ws.Range(ws.Rows(2), ws.Rows(3)).RowHeight = 10
    ws.Range(ws.Rows(5), ws.Rows(7)).RowHeight = 10
    ws.Range(ws.Rows(9), ws.Rows(11)).RowHeight = 10
    ws.Rows(15).RowHeight = 10
    ws.Rows(17).RowHeight = 10
    ws.Rows(19).RowHeight = 10
    ws.Range(ws.Rows(21), ws.Rows(23)).RowHeight = 10
    ws.Rows(25).RowHeight = 10
    ws.Rows(27).RowHeight = 10
    ws.Range(ws.Rows(31), ws.Rows(32)).RowHeight = 10
    ws.Rows(4).RowHeight = 42
    ws.Rows(14).RowHeight = 42
    ws.Rows(16).RowHeight = 42
    ws.Rows(18).RowHeight = 42
    ws.Rows(20).RowHeight = 42
    ws.Rows(28).RowHeight = 42
    ws.Rows(1).RowHeight = 39
    ws.Rows(8).RowHeight = 26
    ws.Rows(29).RowHeight = 90
    ws.Rows(30).RowHeight = 106

    Else

        ' Trame a 1, et les quatre colonnes elargies pour
        ' caler les blocs.
        ws.Range(ws.Columns(1), ws.Columns(59)) _
            .ColumnWidth = 1

        ws.Columns(8).ColumnWidth = 2.17
        ws.Columns(14).ColumnWidth = 1.83
        ws.Columns(26).ColumnWidth = 1.5
        ws.Columns(32).ColumnWidth = 2

    ws.Range(ws.Rows(2), ws.Rows(3)).RowHeight = 10
    ws.Range(ws.Rows(5), ws.Rows(7)).RowHeight = 10
    ws.Range(ws.Rows(9), ws.Rows(11)).RowHeight = 10
    ws.Rows(15).RowHeight = 10
    ws.Rows(17).RowHeight = 10
    ws.Rows(19).RowHeight = 10
    ws.Rows(21).RowHeight = 10
    ws.Rows(23).RowHeight = 10
    ws.Rows(25).RowHeight = 10
    ws.Rows(27).RowHeight = 10
    ws.Rows(29).RowHeight = 10
    ws.Range(ws.Rows(31), ws.Rows(33)).RowHeight = 10
    ws.Rows(37).RowHeight = 10
    ws.Rows(39).RowHeight = 10
    ws.Rows(41).RowHeight = 10
    ws.Range(ws.Rows(43), ws.Rows(45)).RowHeight = 10
    ws.Rows(4).RowHeight = 42
    ws.Rows(14).RowHeight = 42
    ws.Rows(16).RowHeight = 42
    ws.Rows(18).RowHeight = 42
    ws.Rows(20).RowHeight = 42
    ws.Rows(22).RowHeight = 42
    ws.Rows(24).RowHeight = 42
    ws.Rows(26).RowHeight = 42
    ws.Rows(28).RowHeight = 42
    ws.Rows(30).RowHeight = 42
    ws.Rows(36).RowHeight = 42
    ws.Rows(38).RowHeight = 42
    ws.Rows(40).RowHeight = 42
    ws.Rows(42).RowHeight = 42
    ws.Rows(50).RowHeight = 42
    ws.Rows(12).RowHeight = 15
    ws.Rows(34).RowHeight = 15
    ws.Rows(46).RowHeight = 15
    ws.Rows(48).RowHeight = 15
    ws.Range(ws.Rows(52), ws.Rows(53)).RowHeight = 15
    ws.Rows(1).RowHeight = 39
    ws.Rows(8).RowHeight = 26
    ws.Rows(35).RowHeight = 13
    ws.Rows(47).RowHeight = 11
    ws.Rows(49).RowHeight = 6
    ws.Rows(51).RowHeight = 93

    End If

End Sub


' =========================================================
' CADRES
' =========================================================

Private Sub DessinerCadres( _
    ByVal ws As Worksheet, _
    ByVal Mode As String)

    Dim Zones As Variant
    Dim i As Long

    If Mode = MODE_PLEIN_ECRAN Then

        Zones = Array( _
            "B3:AY5", "BA3:BM5", "B7:AY9", "BA7:BM9", _
            "B11:AY21", "BA11:BS21", "B23:AY29")

    Else

        Zones = Array( _
            "B3:AM5", "B7:AM9", "B11:AM31", _
            "B33:AM43", "B45:AM52")

    End If

    For i = 0 To UBound(Zones)
        AppliquerCadreVert ws, ws.Range(CStr(Zones(i)))
    Next i

End Sub


Private Sub AppliquerCadreVert( _
    ByVal ws As Worksheet, _
    ByVal Zone As Range)

    Dim Cotes As Variant
    Dim i As Long

    Cotes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, _
        xlEdgeBottom)

    For i = 0 To UBound(Cotes)

        With Zone.Borders(Cotes(i))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = VERT_TITRE
        End With

    Next i

End Sub


' =========================================================
' BOUTON DE BASCULE ET PLAGE DES JOUEURS
' =========================================================

Public Sub MettreAJourBoutonSwitchMode()

    Dim shp As Shape

    On Error Resume Next
    Set shp = shSaisieVideo.Shapes("BTN_SWITCH_MODE")
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    Select Case DetecterDispositionSaisieVideo(shSaisieVideo)

        Case MODE_BANDEAU
            shp.TextFrame2.TextRange.Text = "Mode plein ecran"

        Case MODE_PLEIN_ECRAN
            shp.TextFrame2.TextRange.Text = "Mode bandeau"

        Case Else
            shp.TextFrame2.TextRange.Text = "Changer de mode"

    End Select

End Sub


' Les boutons joueurs, reunis en une seule plage.
'
' Quinze emplacements seulement : les remplacants n'ont
' plus de bouton, ils apparaissent a la place du joueur
' qu'ils remplacent.
Public Function PlageBoutonsJoueurs( _
    Optional ByVal ws As Worksheet _
) As Range

    Dim Cible As Worksheet
    Dim i As Long
    Dim Bouton As Range
    Dim Total As Range
    Dim Noms As Variant
    Dim Element As Variant

    If ws Is Nothing Then
        Set Cible = shSaisieVideo
    Else
        Set Cible = ws
    End If

    Noms = Array("BTN_JO_COLLECTIF", "BTN_JO_NON_IDENTIFIE")

    For i = 1 To 15

        Set Bouton = Nothing

        On Error Resume Next
        Set Bouton = Cible.Range("BTN_JO_" & i)
        On Error GoTo 0

        If Not Bouton Is Nothing Then

            If Total Is Nothing Then
                Set Total = Bouton
            Else
                Set Total = Union(Total, Bouton)
            End If

        End If

    Next i

    For Each Element In Noms

        Set Bouton = Nothing

        On Error Resume Next
        Set Bouton = Cible.Range(CStr(Element))
        On Error GoTo 0

        If Not Bouton Is Nothing Then

            If Total Is Nothing Then
                Set Total = Bouton
            Else
                Set Total = Union(Total, Bouton)
            End If

        End If

    Next Element

    Set PlageBoutonsJoueurs = Total

End Function


' =========================================================
' TABLES DE DISPOSITION
' =========================================================

Private Function ComposantsPleinEcran() As Variant

    Dim T As String

    T = "TITRE_PALETTE_ACTIONS|BB12|BB12:BR12|" & "PALETTE ACTIONS" & vbTab & _
        "LBL_REPRISE_JEU|BB8||" & "Reprise du jeu sur / ou arret" & vbTab & _
        "BTN_REPRISE_TOUCHE|BD8||" & "Touche " & vbTab & _
        "BTN_REPRISE_MELEE|BF8||" & "Mel" & ChrW(233) & "e " & vbTab & _
        "BTN_REPRISE_RENVOI|BH8||" & "renvoi" & vbTab & _
        "BTN_REPRISE_A_LA_MAIN|BJ8||" & ChrW(224) & " la main" & vbTab & _
        "BTN_ARRET_DU_JEU|BL8||" & "Arret du jeu" & vbTab & _
        "LBL_PLAQUAGES|BB14||" & "Plaquages"

    T = T & vbTab & _
        "BTN_PL_OFF|BD14||" & "Offensif" & vbTab & _
        "BTN_PL_NORMAL|BF14||" & "normal" & vbTab & _
        "BTN_PL_RATE|BH14||" & "rat" & ChrW(233) & vbTab & _
        "BTN_PL_A2|BJ14||" & ChrW(224) & " 2" & vbTab & _
        "BTN_PL_HAUT|BL14||" & "Haut" & vbTab & _
        "LBL_MELEES|BN14||" & "M" & ChrW(233) & "l" & ChrW(233) & "es" & vbTab & _
        "BTN_ME_G|BP14||" & "gagn" & ChrW(233) & "e" & vbTab & _
        "BTN_ME_P|BR14||" & "perdue"

    T = T & vbTab & _
        "LBL_TOUCHES|BB16||" & "Touches" & vbTab & _
        "BTN_TO_G|BD16||" & "gagn" & ChrW(233) & "e" & vbTab & _
        "BTN_TO_P|BF16||" & "perdue" & vbTab & _
        "BTN_TO_PAS_DROIT|BH16||" & "pas droit" & vbTab & _
        "LBL_ATTITUDES|BL16||" & "Attitudes au contact" & vbTab & _
        "BTN_FRANCHISSEMENT|BN16||" & "franchissement" & vbTab & _
        "BTN_AVANCEE|BP16||" & "avanc" & ChrW(233) & "e" & vbTab & _
        "BTN_ATT_NULLE|BR16||" & "stopp" & ChrW(233) & " / recul"

    T = T & vbTab & _
        "LBL_CHGT_POSS_1|BB18||" & "Changement possession" & vbTab & _
        "BTN_EN_AVANT|BD18||" & "en-avant" & vbTab & _
        "BTN_JEU_AU_PIED|BF18||" & "jeu au pied" & vbTab & _
        "BTN_RECEPTION_NOUS|BH18||" & "reception nous" & vbTab & _
        "BTN_RECEPTION_ADV|BJ18||" & "r" & ChrW(233) & "ception Adv" & vbTab & _
        "BTN_TURNOVER|BR18||" & "Turn-over" & vbTab & _
        "BTN_GRATTAGE|BL18||" & "Grattage" & vbTab & _
        "BTN_CONTRE_RUCK|BN18||" & "contre-ruck"

    T = T & vbTab & _
        "BTN_ARRACHAGE|BP18||" & "Arrachage" & vbTab & _
        "LBL_PENALITE|BB20||" & "P" & ChrW(233) & "nalit" & ChrW(233) & vbTab & _
        "BTN_PEN_CONTRE_ADV|BD20||" & "contre adversaire" & vbTab & _
        "BTN_PEN_CONTRE_NOUS|BF20||" & "contre" & vbLf & "nous" & vbTab & _
        "LBL_COUP_FRANC|BJ20||" & "Coup franc" & vbTab & _
        "BTN_CF_CONTRE_ADV|BL20||" & "contre adversaire" & vbTab & _
        "BTN_CF_CONTRE_NOUS|BN20||" & "contre" & vbLf & "nous" & vbTab & _
        "LBL_POINTS|BB4||" & "Points"

    T = T & vbTab & _
        "BTN_PTS_ESSAI|BD4||" & "essai" & vbTab & _
        "BTN_PTS_ESSAI_PEN|BF4||" & "essai de penalit" & ChrW(233) & vbTab & _
        "BTN_PTS_TRANSFO|BH4||" & "transfo" & vbTab & _
        "BTN_PTS_PENALITE|BJ4||" & "penalit" & ChrW(233) & vbTab & _
        "BTN_PTS_DROP|BL4||" & "drop" & vbTab & _
        "BTN_REMPLACEMENT|BO7|BO7:BQ9|" & "Remplacement" & vbTab & _
        "LBL_POSSESSION|C4|C4:I4|" & "Possession en faveur de : " & vbTab & _
        "BTN_NOUS|K4|K4:Q4|" & "Nous "

    T = T & vbTab & _
        "BTN_ADV|S4|S4:Y4|" & "Adversaire" & vbTab & _
        "LBL_MI_TEMPS|AA4|AA4:AG4|" & "Mi-temps :" & vbTab & _
        "BTN_MT1|AJ4|AJ4:AP4|" & "MT 1" & vbTab & _
        "BTN_MT2|AR4|AR4:AX4|" & "MT2" & vbTab & _
        "LBL_CARTONS|C8|C8:I8|" & "Cartons" & vbTab & _
        "BTN_CARTON_ROUGE|K8|K8:Q8|" & vbTab & _
        "BTN_CARTON_JAUNE|S8|S8:Y8|" & vbTab & _
        "BTN_CARTON_BLANC|AA8|AA8:AG8|"

    T = T & vbTab & _
        "BTN_CARTON_BLEU|AI8|AI8:AO8|" & vbTab & _
        "TITRE_PALETTE_JOUEURS|C12|C12:AX12|" & "PALETTE JOUEURS" & vbTab & _
        "BTN_JO_1|C14|C14:I14|" & vbTab & _
        "BTN_JO_2|K14|K14:Q14|" & vbTab & _
        "BTN_JO_3|S14|S14:Y14|" & vbTab & _
        "BTN_JO_4|G16|G16:M16|" & vbTab & _
        "BTN_JO_5|O16|O16:U16|" & vbTab & _
        "BTN_JO_6|C18|C18:I18|"

    T = T & vbTab & _
        "BTN_JO_7|S18|S18:Y18|" & vbTab & _
        "BTN_JO_8|K18|K18:Q18|" & vbTab & _
        "BTN_JO_9|AD14|AD14:AJ14|" & vbTab & _
        "BTN_JO_10|AP14|AP14:AV14|" & vbTab & _
        "BTN_JO_11|AB18|AB18:AH18|" & vbTab & _
        "BTN_JO_12|AF16|AF16:AL16|" & vbTab & _
        "BTN_JO_13|AN16|AN16:AT16|" & vbTab & _
        "BTN_JO_14|AR18|AR18:AX18|"

    T = T & vbTab & _
        "BTN_JO_15|AJ18|AJ18:AP18|" & vbTab & _
        "BTN_JO_COLLECTIF|K20|K20:Q20|Collectif" & vbTab & _
        "BTN_JO_NON_IDENTIFIE|AJ20|AJ20:AP20|?" & vbTab & _
        "TITRE_CHRONO_VIDEO|C24|C24:AX24|" & "Gestion chrono  + vid" & ChrW(233) & "o" & vbTab & _
        "LBL_POSITION_VIDEO|C26|C26:K26|" & "Position vid" & ChrW(233) & "o" & vbTab & _
        "CELL_CHRONO_VIDEO|C28|C28:P28|" & "0:00:00" & vbTab & _
        "BTN_CHRONO_MINUS_5_SEC|R28|R28:X28|" & "-5 sec" & vbTab & _
        "BTN_CHRONO_PLAY_PAUSE|Z28|Z28:AF28|" & "PLAY / PAUSE"

    T = T & vbTab & _
        "BTN_CHRONO_PLUS_5_SEC|AH28|AH28:AN28|" & " +5 sec" & vbTab & _
        "BTN_CHRONO_RESET|AP28|AP28:AV28|" & "RESET"

    ComposantsPleinEcran = Split(T, vbTab)

End Function


Private Function ComposantsBandeau() As Variant

    Dim T As String

    T = "TITRE_PALETTE_ACTIONS|C12|C12:AL12|" & "PALETTE ACTIONS" & vbTab & _
        "LBL_REPRISE_JEU|C14|C14:H14|" & "Reprise du jeu sur / ou arret" & vbTab & _
        "BTN_REPRISE_TOUCHE|J14|J14:N14|" & "Touche " & vbTab & _
        "BTN_REPRISE_MELEE|P14|P14:T14|" & "Mel" & ChrW(233) & "e " & vbTab & _
        "BTN_REPRISE_RENVOI|V14|V14:Z14|" & "renvoi" & vbTab & _
        "BTN_REPRISE_A_LA_MAIN|AB14|AB14:AF14|" & ChrW(224) & " la main" & vbTab & _
        "BTN_ARRET_DU_JEU|AH14|AH14:AL14|" & "Arret du jeu" & vbTab & _
        "LBL_PLAQUAGES|C16|C16:H16|" & "Plaquages"

    T = T & vbTab & _
        "BTN_PL_OFF|J16|J16:N16|" & "Offensif" & vbTab & _
        "BTN_PL_NORMAL|P16|P16:T16|" & "normal" & vbTab & _
        "BTN_PL_RATE|V16|V16:Z16|" & "rat" & ChrW(233) & vbTab & _
        "BTN_PL_A2|AB16|AB16:AF16|" & ChrW(224) & " 2" & vbTab & _
        "BTN_PL_HAUT|AH16|AH16:AL16|" & "Haut" & vbTab & _
        "LBL_MELEES|C20|C20:H20|" & "M" & ChrW(233) & "l" & ChrW(233) & "es" & vbTab & _
        "BTN_ME_G|J20|J20:N20|" & "gagn" & ChrW(233) & "e" & vbTab & _
        "BTN_ME_P|P20|P20:T20|" & "perdue"

    T = T & vbTab & _
        "LBL_TOUCHES|C18|C18:H18|" & "Touches" & vbTab & _
        "BTN_TO_G|J18|J18:N18|" & "gagn" & ChrW(233) & "e" & vbTab & _
        "BTN_TO_P|P18|P18:T18|" & "perdue" & vbTab & _
        "BTN_TO_PAS_DROIT|V18|V18:Z18|" & "pas droit" & vbTab & _
        "LBL_ATTITUDES|C22|C22:H22|" & "Attitudes au contact" & vbTab & _
        "BTN_FRANCHISSEMENT|J22|J22:N22|" & "franchissement" & vbTab & _
        "BTN_AVANCEE|P22|P22:T22|" & "avanc" & ChrW(233) & "e" & vbTab & _
        "BTN_ATT_NULLE|V22|V22:Z22|" & "stopp" & ChrW(233) & " / recul"

    T = T & vbTab & _
        "LBL_CHGT_POSS_1|C24|C24:H24|" & "Changement possession" & vbTab & _
        "BTN_EN_AVANT|J24|J24:N24|" & "en-avant" & vbTab & _
        "BTN_JEU_AU_PIED|P24|P24:T24|" & "jeu au pied" & vbTab & _
        "BTN_RECEPTION_NOUS|V24|V24:Z24|" & "reception nous" & vbTab & _
        "BTN_RECEPTION_ADV|AB24|AB24:AF24|" & "r" & ChrW(233) & "ception Adv" & vbTab & _
        "BTN_TURNOVER|AH24|AH24:AL24|" & "Turn-over" & vbTab & _
        "LBL_CHGT_POSS_2|C26|C26:H26|" & "Changement possession" & vbTab & _
        "BTN_GRATTAGE|J26|J26:N26|" & "Grattage"

    T = T & vbTab & _
        "BTN_CONTRE_RUCK|P26|P26:T26|" & "contre-ruck" & vbTab & _
        "BTN_ARRACHAGE|V26|V26:Z26|" & "Arrachage" & vbTab & _
        "LBL_PENALITE|C28|C28:H28|" & "P" & ChrW(233) & "nalit" & ChrW(233) & vbTab & _
        "BTN_PEN_CONTRE_ADV|J28|J28:N28|" & "contre adversaire" & vbTab & _
        "BTN_PEN_CONTRE_NOUS|P28|P28:T28|" & "contre" & vbLf & "nous" & vbTab & _
        "LBL_COUP_FRANC|V28|V28:Z28|" & "Coup franc" & vbTab & _
        "BTN_CF_CONTRE_ADV|AB28|AB28:AF28|" & "contre adversaire" & vbTab & _
        "BTN_CF_CONTRE_NOUS|AH28|AH28:AL28|" & "contre" & vbLf & "nous"

    T = T & vbTab & _
        "LBL_POINTS|C30|C30:H30|" & "Points" & vbTab & _
        "BTN_PTS_ESSAI|J30|J30:N30|" & "essai" & vbTab & _
        "BTN_PTS_ESSAI_PEN|P30|P30:T30|" & "essai de penalit" & ChrW(233) & vbTab & _
        "BTN_PTS_TRANSFO|V30|V30:Z30|" & "transfo" & vbTab & _
        "BTN_PTS_PENALITE|AB30|AB30:AF30|" & "penalit" & ChrW(233) & vbTab & _
        "BTN_PTS_DROP|AH30|AH30:AL30|" & "drop" & vbTab & _
        "BTN_REMPLACEMENT|AB20|AB20:AL20|" & "Remplacement" & vbTab & _
        "LBL_POSSESSION|C4|C4:H4|" & "Possession en faveur de : "

    T = T & vbTab & _
        "BTN_NOUS|J4|J4:N4|" & "Nous " & vbTab & _
        "BTN_ADV|P4|P4:T4|" & "Adversaire" & vbTab & _
        "BTN_MT1|AB4|AB4:AF4|" & "MT 1" & vbTab & _
        "BTN_MT2|AH4|AH4:AL4|" & "MT2" & vbTab & _
        "LBL_CARTONS|C8|C8:H8|" & "Cartons" & vbTab & _
        "BTN_CARTON_ROUGE|J8|J8:N8|" & vbTab & _
        "BTN_CARTON_JAUNE|P8|P8:T8|" & vbTab & _
        "BTN_CARTON_BLANC|V8|V8:Z8|"

    T = T & vbTab & _
        "BTN_CARTON_BLEU|AB8|AB8:AF8|" & vbTab & _
        "TITRE_PALETTE_JOUEURS|C34|C34:AL34|" & "PALETTE JOUEURS" & vbTab & _
        "BTN_JO_1|C36|C36:G36|" & vbTab & _
        "BTN_JO_2|I36|I36:M36|" & vbTab & _
        "BTN_JO_3|O36|O36:S36|" & vbTab & _
        "BTN_JO_4|F38|F38:J38|" & vbTab & _
        "BTN_JO_5|L38|L38:P38|" & vbTab & _
        "BTN_JO_6|C40|C40:G40|"

    T = T & vbTab & _
        "BTN_JO_7|O40|O40:S40|" & vbTab & _
        "BTN_JO_8|I40|I40:M40|" & vbTab & _
        "BTN_JO_9|W36|W36:AA36|" & vbTab & _
        "BTN_JO_10|AC36|AC36:AG36|" & vbTab & _
        "BTN_JO_11|V40|V40:Z40|" & vbTab & _
        "BTN_JO_12|Y38|Y38:AC38|" & vbTab & _
        "BTN_JO_13|AE38|AE38:AI38|" & vbTab & _
        "BTN_JO_14|AH40|AH40:AL40|"

    T = T & vbTab & _
        "BTN_JO_15|AB40|AB40:AF40|" & vbTab & _
        "BTN_JO_COLLECTIF|I42|I42:M42|Collectif" & vbTab & _
        "BTN_JO_NON_IDENTIFIE|AB42|AB42:AF42|?" & vbTab & _
        "TITRE_CHRONO_VIDEO|C46|C46:AL46|" & "Gestion chrono  + vid" & ChrW(233) & "o" & vbTab & _
        "LBL_POSITION_VIDEO|C48|C48:J48|" & "Position vid" & ChrW(233) & "o" & vbTab & _
        "CELL_CHRONO_VIDEO|C50|C50:N50|" & "0:00:00" & vbTab & _
        "BTN_CHRONO_MINUS_5_SEC|P50|P50:T50|" & "-5 sec" & vbTab & _
        "BTN_CHRONO_PLAY_PAUSE|V50|V50:Z50|" & "PLAY / PAUSE"

    T = T & vbTab & _
        "BTN_CHRONO_PLUS_5_SEC|AB50|AB50:AF50|" & " +5 sec" & vbTab & _
        "BTN_CHRONO_RESET|AH50|AH50:AL50|" & "RESET"

    ComposantsBandeau = Split(T, vbTab)

End Function
