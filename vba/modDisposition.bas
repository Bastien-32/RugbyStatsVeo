Attribute VB_Name = "modDisposition"
Option Explicit

Private Const MODE_BANDEAU As String = "BANDEAU"
Private Const MODE_PLEIN_ECRAN As String = "PLEIN_ECRAN"

' =========================================================
' BOUTON UNIQUE
' =========================================================

Public Sub BasculerDispositionSaisieVideo()

    Dim ws As Worksheet
    Dim ModeActuel As String

    Set ws = shSaisieVideo
    ModeActuel = DetecterDispositionSaisieVideo(ws)

    Select Case ModeActuel

        Case MODE_BANDEAU
            PasserEnModePleinEcran

        Case MODE_PLEIN_ECRAN
            PasserEnModeBandeau

        Case Else
            MsgBox _
                "Impossible de determiner la disposition actuelle.", _
                vbExclamation, _
                "Disposition saisie video"

    End Select

End Sub

' =========================================================
' DETECTION
' =========================================================

Private Function DetecterDispositionSaisieVideo(ByVal ws As Worksheet) As String

    Dim R As Range

    Set R = ws.Range("BTN_PL_OFF")

    If R.Row = 16 And R.Column = 6 Then

        DetecterDispositionSaisieVideo = MODE_BANDEAU

    ElseIf R.Row = 14 And R.Column = 20 Then

        DetecterDispositionSaisieVideo = MODE_PLEIN_ECRAN

    Else

        DetecterDispositionSaisieVideo = ""

    End If

End Function

' =========================================================
' DEPLACEMENT D'UN COMPOSANT
'
' Pas de Cut.
' Pas de Copy/Paste.
' Pas de zone parking.
'
' La valeur et le style sont reproduits a destination,
' puis l'ancienne zone est effacee.
' =========================================================

Private Sub DeplacerComposant( _
    ByVal ws As Worksheet, _
    ByVal NomComposant As String, _
    ByVal AdresseDestination As String _
)

    Dim SourceNommee As Range
    Dim Source As Range
    Dim Destination As Range
    Dim ValeurSource As Variant
    Dim FormuleSource As Variant
    Dim SourceContientFormule As Boolean

    Set SourceNommee = ws.Range(NomComposant)

    If SourceNommee.MergeCells Then
        Set Source = SourceNommee.MergeArea
    Else
        Set Source = SourceNommee
    End If

    Set Destination = ws.Range(AdresseDestination)

    ' Une destination ne doit jamais recouvrir sa source.
    If Not Intersect(Source, Destination) Is Nothing Then

        If Source.Address <> Destination.Address Then

            Err.Raise _
                vbObjectError + 4000, _
                "DeplacerComposant", _
                "Chevauchement source/destination pour " & NomComposant

        Else

            Exit Sub

        End If

    End If

    SourceContientFormule = Source.Cells(1, 1).HasFormula

    If SourceContientFormule Then
        FormuleSource = Source.Cells(1, 1).Formula
    Else
        ValeurSource = Source.Cells(1, 1).Value
    End If

    ' Prepare la destination.
    If Destination.MergeCells Then
        Destination.MergeArea.UnMerge
    End If

    Destination.Clear

    ' Recopie le style avant la fusion.
    CopierStyleComposant Source.Cells(1, 1), Destination

    ' Fusion eventuelle de la nouvelle zone.
    If Destination.Cells.CountLarge > 1 Then
        Destination.Merge
    End If

    ' Valeur ou formule.
    If SourceContientFormule Then
        Destination.Cells(1, 1).Formula = FormuleSource
    Else
        Destination.Cells(1, 1).Value = ValeurSource
    End If

    ' Le nom suit le composant.
    ReaffecterNom NomComposant, Destination.Cells(1, 1)

    ' Seulement maintenant, on supprime l'ancienne zone.
    If Source.MergeCells Then
        Source.UnMerge
    End If

    Source.Clear

End Sub

' =========================================================
' STYLE DU COMPOSANT
'
' Copie explicite des proprietes :
' comportement identique Mac et Windows.
' =========================================================

Private Sub CopierStyleComposant( _
    ByVal Modele As Range, _
    ByVal Destination As Range _
)

    Dim Cellule As Range

    For Each Cellule In Destination.Cells

        With Cellule

            ' Police
            .Font.Name = Modele.Font.Name
            .Font.Size = Modele.Font.Size
            .Font.Bold = Modele.Font.Bold
            .Font.Italic = Modele.Font.Italic
            .Font.Underline = Modele.Font.Underline
            .Font.Color = Modele.Font.Color

            ' Fond
            .Interior.Pattern = Modele.Interior.Pattern
            .Interior.Color = Modele.Interior.Color

            ' Alignement
            .HorizontalAlignment = Modele.HorizontalAlignment
            .VerticalAlignment = Modele.VerticalAlignment
            .WrapText = Modele.WrapText
            .ShrinkToFit = Modele.ShrinkToFit

            ' Format
            .NumberFormat = Modele.NumberFormat

            ' Protection
            .Locked = Modele.Locked

        End With

        CopierBorduresCellule Modele, Cellule

    Next Cellule

End Sub

Private Sub CopierBorduresCellule( _
    ByVal Modele As Range, _
    ByVal Destination As Range _
)

    Dim TypeBordure As Variant

    For Each TypeBordure In Array( _
        xlEdgeLeft, _
        xlEdgeTop, _
        xlEdgeBottom, _
        xlEdgeRight _
    )

        With Destination.Borders(TypeBordure)

            .LineStyle = Modele.Borders(TypeBordure).LineStyle

            If .LineStyle <> xlLineStyleNone Then
                .Weight = Modele.Borders(TypeBordure).Weight
                .Color = Modele.Borders(TypeBordure).Color
            End If

        End With

    Next TypeBordure

End Sub

' =========================================================
' REAFFECTATION D'UN NOM
' =========================================================

Private Sub ReaffecterNom( _
    ByVal NomPlage As String, _
    ByVal Cellule As Range _
)

    Dim NomFeuille As String
    Dim ReferenceR1C1 As String

    NomFeuille = Replace(Cellule.Worksheet.Name, "'", "''")

    ReferenceR1C1 = _
        "='" & NomFeuille & "'!R" & _
        Cellule.Row & "C" & Cellule.Column

    ThisWorkbook.Names(NomPlage).RefersToR1C1 = ReferenceR1C1

End Sub

' =========================================================
' ACTIONS -> PLEIN ECRAN
' =========================================================

Private Sub PlacerActionsPleinEcran(ByVal ws As Worksheet)

    ' POINTS
    DeplacerComposant ws, "LBL_POINTS", "R4"
    DeplacerComposant ws, "BTN_PTS_ESSAI", "T4"
    DeplacerComposant ws, "BTN_PTS_ESSAI_PEN", "V4"
    DeplacerComposant ws, "BTN_PTS_TRANSFO", "X4"
    DeplacerComposant ws, "BTN_PTS_PENALITE", "Z4"
    DeplacerComposant ws, "BTN_PTS_DROP", "AB4"

    ' REPRISE DU JEU
    DeplacerComposant ws, "LBL_REPRISE_JEU", "R8"
    DeplacerComposant ws, "BTN_REPRISE_TOUCHE", "T8"
    DeplacerComposant ws, "BTN_REPRISE_MELEE", "V8"
    DeplacerComposant ws, "BTN_REPRISE_RENVOI", "X8"
    DeplacerComposant ws, "BTN_REPRISE_A_LA_MAIN", "Z8"
    DeplacerComposant ws, "BTN_ARRET_DU_JEU", "AB8"

    ' TITRE
    DeplacerComposant ws, "TITRE_PALETTE_ACTIONS", "R12:AH12"

    ' PLAQUAGES
    DeplacerComposant ws, "LBL_PLAQUAGES", "R14"
    DeplacerComposant ws, "BTN_PL_OFF", "T14"
    DeplacerComposant ws, "BTN_PL_NORMAL", "V14"
    DeplacerComposant ws, "BTN_PL_RATE", "X14"
    DeplacerComposant ws, "BTN_PL_A2", "Z14"
    DeplacerComposant ws, "BTN_PL_HAUT", "AB14"
    DeplacerComposant ws, "BTN_TURNOVER", "AH18"


    ' TOUCHES
    DeplacerComposant ws, "LBL_TOUCHES", "R16"
    DeplacerComposant ws, "BTN_TO_G", "T16"
    DeplacerComposant ws, "BTN_TO_P", "V16"
    DeplacerComposant ws, "BTN_TO_PAS_DROIT", "X16"

    ' MELEES
    DeplacerComposant ws, "LBL_MELEES", "AD14"
    DeplacerComposant ws, "BTN_ME_G", "AF14"
    DeplacerComposant ws, "BTN_ME_P", "AH14"

    ' ATTITUDE AU CONTACT
    DeplacerComposant ws, "LBL_ATTITUDES", "AB16"
    DeplacerComposant ws, "BTN_FRANCHISSEMENT", "AD16"
    DeplacerComposant ws, "BTN_AVANCEE", "AF16"
    DeplacerComposant ws, "BTN_ATT_NULLE", "AH16"


    ' CHANGEMENT DE POSSESSION
    DeplacerComposant ws, "LBL_CHGT_POSS_1", "R18"

    ' Non affiche en plein ecran mais conserve.
    DeplacerComposant ws, "LBL_CHGT_POSS_2", "AZ1"

    DeplacerComposant ws, "BTN_EN_AVANT", "T18"
    DeplacerComposant ws, "BTN_JEU_AU_PIED", "V18"
    DeplacerComposant ws, "BTN_RECEPTION_NOUS", "X18"
    DeplacerComposant ws, "BTN_RECEPTION_ADV", "Z18"
    DeplacerComposant ws, "BTN_GRATTAGE", "AB18"
    DeplacerComposant ws, "BTN_CONTRE_RUCK", "AD18"
    DeplacerComposant ws, "BTN_ARRACHAGE", "AF18"

    ' PENALITES
    DeplacerComposant ws, "LBL_PENALITE", "R20"
    DeplacerComposant ws, "BTN_PEN_CONTRE_ADV", "T20"
    DeplacerComposant ws, "BTN_PEN_CONTRE_NOUS", "V20"
    DeplacerComposant ws, "LBL_MOTIF_PENALITE", "X20"
    DeplacerComposant ws, "BTN_PEN_MAUL", "Z20"
    DeplacerComposant ws, "BTN_PEN_RUCK", "AB20"
    DeplacerComposant ws, "BTN_PEN_HORS_JEU", "AD20"
    DeplacerComposant ws, "BTN_PEN_PL_A_2", "AF20"
    DeplacerComposant ws, "BTN_PEN_PL_HAUT", "AH20"

End Sub

' =========================================================
' ACTIONS -> BANDEAU
' =========================================================

Private Sub PlacerActionsBandeau(ByVal ws As Worksheet)

    DeplacerComposant ws, "TITRE_PALETTE_ACTIONS", "D12:N12"

    DeplacerComposant ws, "LBL_REPRISE_JEU", "D14"
    DeplacerComposant ws, "BTN_REPRISE_TOUCHE", "F14"
    DeplacerComposant ws, "BTN_REPRISE_MELEE", "H14"
    DeplacerComposant ws, "BTN_REPRISE_RENVOI", "J14"
    DeplacerComposant ws, "BTN_REPRISE_A_LA_MAIN", "L14"
    DeplacerComposant ws, "BTN_ARRET_DU_JEU", "N14"

    DeplacerComposant ws, "LBL_PLAQUAGES", "D16"
    DeplacerComposant ws, "BTN_PL_OFF", "F16"
    DeplacerComposant ws, "BTN_PL_NORMAL", "H16"
    DeplacerComposant ws, "BTN_PL_RATE", "J16"
    DeplacerComposant ws, "BTN_PL_A2", "L16"
    DeplacerComposant ws, "BTN_PL_HAUT", "N16"

    DeplacerComposant ws, "LBL_TOUCHES", "D18"
    DeplacerComposant ws, "BTN_TO_G", "F18"
    DeplacerComposant ws, "BTN_TO_P", "H18"
    DeplacerComposant ws, "BTN_TO_PAS_DROIT", "J18"

    DeplacerComposant ws, "LBL_MELEES", "D20"
    DeplacerComposant ws, "BTN_ME_G", "F20"
    DeplacerComposant ws, "BTN_ME_P", "H20"

    ' ATTITUDE AU CONTACT
    DeplacerComposant ws, "LBL_ATTITUDES", "D22"
    DeplacerComposant ws, "BTN_FRANCHISSEMENT", "F22"
    DeplacerComposant ws, "BTN_AVANCEE", "H22"
    DeplacerComposant ws, "BTN_ATT_NULLE", "J22"

    DeplacerComposant ws, "LBL_CHGT_POSS_1", "D24"
    DeplacerComposant ws, "BTN_EN_AVANT", "F24"
    DeplacerComposant ws, "BTN_JEU_AU_PIED", "H24"
    DeplacerComposant ws, "BTN_RECEPTION_NOUS", "J24"
    DeplacerComposant ws, "BTN_RECEPTION_ADV", "L24"
    DeplacerComposant ws, "BTN_TURNOVER", "N24"

    DeplacerComposant ws, "LBL_CHGT_POSS_2", "D26"
    DeplacerComposant ws, "BTN_GRATTAGE", "F26"
    DeplacerComposant ws, "BTN_CONTRE_RUCK", "H26"
    DeplacerComposant ws, "BTN_ARRACHAGE", "J26"

    DeplacerComposant ws, "LBL_PENALITE", "D28"
    DeplacerComposant ws, "BTN_PEN_CONTRE_ADV", "F28"
    DeplacerComposant ws, "BTN_PEN_CONTRE_NOUS", "H28"

    DeplacerComposant ws, "LBL_MOTIF_PENALITE", "D30"
    DeplacerComposant ws, "BTN_PEN_MAUL", "F30"
    DeplacerComposant ws, "BTN_PEN_RUCK", "H30"
    DeplacerComposant ws, "BTN_PEN_HORS_JEU", "J30"
    DeplacerComposant ws, "BTN_PEN_PL_A_2", "L30"
    DeplacerComposant ws, "BTN_PEN_PL_HAUT", "N30"

    DeplacerComposant ws, "LBL_POINTS", "D32"
    DeplacerComposant ws, "BTN_PTS_ESSAI", "F32"
    DeplacerComposant ws, "BTN_PTS_ESSAI_PEN", "H32"
    DeplacerComposant ws, "BTN_PTS_TRANSFO", "J32"
    DeplacerComposant ws, "BTN_PTS_PENALITE", "L32"
    DeplacerComposant ws, "BTN_PTS_DROP", "N32"

End Sub

' =========================================================
' JOUEURS -> PLEIN ECRAN
' =========================================================

Private Sub PlacerJoueursPleinEcran(ByVal ws As Worksheet)

    DeplacerComposant ws, "TITRE_PALETTE_JOUEURS", "D12:N12"

    DeplacerComposant ws, "BTN_JO_1", "D14"
    DeplacerComposant ws, "BTN_JO_2", "F14"
    DeplacerComposant ws, "BTN_JO_3", "H14"
    DeplacerComposant ws, "BTN_JO_4", "J14"
    DeplacerComposant ws, "BTN_JO_5", "L14"
    DeplacerComposant ws, "BTN_JO_6", "N14"

    DeplacerComposant ws, "BTN_JO_7", "D16"
    DeplacerComposant ws, "BTN_JO_8", "F16"
    DeplacerComposant ws, "BTN_JO_9", "H16"
    DeplacerComposant ws, "BTN_JO_10", "J16"
    DeplacerComposant ws, "BTN_JO_11", "L16"
    DeplacerComposant ws, "BTN_JO_12", "N16"

    DeplacerComposant ws, "BTN_JO_13", "D18"
    DeplacerComposant ws, "BTN_JO_14", "F18"
    DeplacerComposant ws, "BTN_JO_15", "H18"
    DeplacerComposant ws, "BTN_JO_16", "J18"
    DeplacerComposant ws, "BTN_JO_17", "L18"
    DeplacerComposant ws, "BTN_JO_18", "N18"

    DeplacerComposant ws, "BTN_JO_19", "D20"
    DeplacerComposant ws, "BTN_JO_20", "F20"
    DeplacerComposant ws, "BTN_JO_21", "H20"
    DeplacerComposant ws, "BTN_JO_22", "J20"
    DeplacerComposant ws, "BTN_JO_COLLECTIF", "L20"
    DeplacerComposant ws, "BTN_JO_NON_IDENTIFIE", "N20"

    ReaffecterPlagesJoueurs ws, MODE_PLEIN_ECRAN

End Sub

' =========================================================
' JOUEURS -> BANDEAU
' =========================================================

Private Sub PlacerJoueursBandeau(ByVal ws As Worksheet)

    DeplacerComposant ws, "TITRE_PALETTE_JOUEURS", "D36:N36"

    DeplacerComposant ws, "BTN_JO_1", "D38"
    DeplacerComposant ws, "BTN_JO_2", "F38"
    DeplacerComposant ws, "BTN_JO_3", "H38"
    DeplacerComposant ws, "BTN_JO_4", "J38"
    DeplacerComposant ws, "BTN_JO_5", "L38"
    DeplacerComposant ws, "BTN_JO_6", "N38"

    DeplacerComposant ws, "BTN_JO_7", "D40"
    DeplacerComposant ws, "BTN_JO_8", "F40"
    DeplacerComposant ws, "BTN_JO_9", "H40"
    DeplacerComposant ws, "BTN_JO_10", "J40"
    DeplacerComposant ws, "BTN_JO_11", "L40"
    DeplacerComposant ws, "BTN_JO_12", "N40"

    DeplacerComposant ws, "BTN_JO_13", "D42"
    DeplacerComposant ws, "BTN_JO_14", "F42"
    DeplacerComposant ws, "BTN_JO_15", "H42"
    DeplacerComposant ws, "BTN_JO_16", "J42"
    DeplacerComposant ws, "BTN_JO_17", "L42"
    DeplacerComposant ws, "BTN_JO_18", "N42"

    DeplacerComposant ws, "BTN_JO_19", "D44"
    DeplacerComposant ws, "BTN_JO_20", "F44"
    DeplacerComposant ws, "BTN_JO_21", "H44"
    DeplacerComposant ws, "BTN_JO_22", "J44"
    DeplacerComposant ws, "BTN_JO_COLLECTIF", "L44"
    DeplacerComposant ws, "BTN_JO_NON_IDENTIFIE", "N44"

    ReaffecterPlagesJoueurs ws, MODE_BANDEAU

End Sub

Private Sub ReaffecterPlagesJoueurs( _
    ByVal ws As Worksheet, _
    ByVal Mode As String _
)

    If Mode = MODE_PLEIN_ECRAN Then

        ThisWorkbook.Names("PALETTE_JOUEURS").RefersTo = _
            "='" & ws.Name & "'!$C$13:$O$21"

    Else

        ThisWorkbook.Names("PALETTE_JOUEURS").RefersTo = _
            "='" & ws.Name & "'!$C$35:$O$43"

    End If

End Sub

' =========================================================
' CHRONO -> PLEIN ECRAN
' =========================================================

Private Sub PlacerChronoPleinEcran(ByVal ws As Worksheet)

    DeplacerComposant ws, "TITRE_CHRONO_VIDEO", "D24:N24"
    DeplacerComposant ws, "LBL_POSITION_VIDEO", "D26"
    DeplacerComposant ws, "CELL_CHRONO_VIDEO", "D28:F28"
    DeplacerComposant ws, "BTN_CHRONO_MINUS_5_SEC", "H28"
    DeplacerComposant ws, "BTN_CHRONO_PLAY_PAUSE", "J28"
    DeplacerComposant ws, "BTN_CHRONO_PLUS_5_SEC", "L28"
    DeplacerComposant ws, "BTN_CHRONO_RESET", "N28"

End Sub

' =========================================================
' CHRONO -> BANDEAU
' =========================================================

Private Sub PlacerChronoBandeau(ByVal ws As Worksheet)

    DeplacerComposant ws, "TITRE_CHRONO_VIDEO", "D48:N48"
    DeplacerComposant ws, "LBL_POSITION_VIDEO", "D50"
    DeplacerComposant ws, "CELL_CHRONO_VIDEO", "D52:F52"
    DeplacerComposant ws, "BTN_CHRONO_MINUS_5_SEC", "H52"
    DeplacerComposant ws, "BTN_CHRONO_PLAY_PAUSE", "J52"
    DeplacerComposant ws, "BTN_CHRONO_PLUS_5_SEC", "L52"
    DeplacerComposant ws, "BTN_CHRONO_RESET", "N52"

End Sub


' =========================================================
' DIMENSIONS PLEIN ECRAN
' =========================================================

Private Sub AppliquerDimensionsPleinEcran(ByVal ws As Worksheet)

    ws.Rows(22).RowHeight = 10
    ws.Rows(24).RowHeight = 15
    ws.Rows(26).RowHeight = 15
    ws.Rows(29).RowHeight = 90
    ws.Rows(30).RowHeight = 106

    ws.Columns("Q").ColumnWidth = 1
    ws.Columns("R").ColumnWidth = 10
    ws.Columns("S").ColumnWidth = 1
    ws.Columns("T").ColumnWidth = 10
    ws.Columns("U").ColumnWidth = 1
    ws.Columns("V").ColumnWidth = 10
    ws.Columns("W").ColumnWidth = 1
    ws.Columns("X").ColumnWidth = 10
    ws.Columns("Y").ColumnWidth = 1
    ws.Columns("Z").ColumnWidth = 10
    ws.Columns("AA").ColumnWidth = 1
    ws.Columns("AB").ColumnWidth = 10
    ws.Columns("AC").ColumnWidth = 1
    ws.Columns("AD").ColumnWidth = 10
    ws.Columns("AE").ColumnWidth = 1
    ws.Columns("AF").ColumnWidth = 10
    ws.Columns("AG").ColumnWidth = 1
    ws.Columns("AH").ColumnWidth = 10
    ws.Columns("AI").ColumnWidth = 1
    
    ' Les lignes utilisées par le bandeau reprennent une
    ' hauteur neutre : elles sont vides en plein écran.
    ws.Rows("33:53").RowHeight = 15


End Sub

' =========================================================
' DIMENSIONS BANDEAU
' =========================================================

Private Sub AppliquerDimensionsBandeau(ByVal ws As Worksheet)

    ' Palette actions
    ws.Rows(22).RowHeight = 42
    ws.Rows(24).RowHeight = 42
    ws.Rows(26).RowHeight = 42
    ws.Rows(28).RowHeight = 42
    ws.Rows(29).RowHeight = 10
    ws.Rows(30).RowHeight = 42
    ws.Rows(31).RowHeight = 10
    ws.Rows(32).RowHeight = 42
    ws.Rows(33).RowHeight = 10

    ' Palette joueurs
    ws.Rows(35).RowHeight = 10
    ws.Rows(36).RowHeight = 15
    ws.Rows(37).RowHeight = 13
    ws.Rows(38).RowHeight = 42
    ws.Rows(39).RowHeight = 10
    ws.Rows(40).RowHeight = 42
    ws.Rows(41).RowHeight = 10
    ws.Rows(42).RowHeight = 42
    ws.Rows(43).RowHeight = 10
    ws.Rows(44).RowHeight = 42
    ws.Rows(45).RowHeight = 10
    ws.Rows(46).RowHeight = 10
    ws.Rows(47).RowHeight = 10

    ' Gestion chrono et video
    ws.Rows(48).RowHeight = 15
    ws.Rows(49).RowHeight = 11
    ws.Rows(50).RowHeight = 15
    ws.Rows(51).RowHeight = 6
    ws.Rows(52).RowHeight = 42
    ws.Rows(53).RowHeight = 93

    ws.Columns("P").ColumnWidth = 1
    ws.Columns("Q").ColumnWidth = 1
    ws.Columns("R").ColumnWidth = 10
    ws.Columns("S").ColumnWidth = 1
    ws.Columns("T").ColumnWidth = 10
    ws.Columns("U").ColumnWidth = 1
    ws.Columns("V").ColumnWidth = 10
    ws.Columns("W").ColumnWidth = 1
    ws.Columns("X").ColumnWidth = 10
    ws.Columns("Y").ColumnWidth = 1
    ws.Columns("Z").ColumnWidth = 10
    ws.Columns("AA").ColumnWidth = 1
    ws.Columns("AB").ColumnWidth = 10
    ws.Columns("AC").ColumnWidth = 1
    ws.Columns("AD").ColumnWidth = 10
    ws.Columns("AE").ColumnWidth = 1
    ws.Columns("AF").ColumnWidth = 10
    ws.Columns("AG").ColumnWidth = 1
    ws.Columns("AH").ColumnWidth = 10
    ws.Columns("AI").ColumnWidth = 1

End Sub

' =========================================================
' BORDURES
' =========================================================

Private Sub EffacerBorduresZone(ByVal Zone As Range)

    Zone.Borders.LineStyle = xlLineStyleNone

End Sub

Private Sub AppliquerCadreVert( _
    ByVal ws As Worksheet, _
    ByVal Zone As Range _
)

    Dim Modele As Range
    Dim B As Variant
    Dim Couleur As Long
    Dim Epaisseur As Long
    Dim Style As Long

    ' Cadre Possession servant de reference.
    Set Modele = ws.Range("C3:O6")

    With Modele.Borders(xlEdgeLeft)
        Couleur = .Color
        Epaisseur = .Weight
        Style = .LineStyle
    End With

    For Each B In Array( _
        xlEdgeLeft, _
        xlEdgeTop, _
        xlEdgeBottom, _
        xlEdgeRight _
    )

        With Zone.Borders(B)
            .LineStyle = Style
            .Weight = Epaisseur
            .Color = Couleur
        End With

    Next B

End Sub

Private Sub DessinerCadresPleinEcran(ByVal ws As Worksheet)

    ' =====================================================
    ' GAUCHE
    ' =====================================================

    ' Palette joueurs
    AppliquerCadreVert ws, ws.Range("C11:O21")

    ' Gestion chrono et video
    AppliquerCadreVert ws, ws.Range("C23:O29")

    ' =====================================================
    ' DROITE
    ' =====================================================

    ' Points
    AppliquerCadreVert ws, ws.Range("Q3:AC5")

    ' Reprise du jeu et arret
    AppliquerCadreVert ws, ws.Range("Q7:AC9")

    ' Palette actions principale
    AppliquerCadreVert ws, ws.Range("Q11:AI21")

    AppliquerBorduresBlanchesZone ws.Range("C3:O29")
    AppliquerBorduresBlanchesZone ws.Range("Q3:AI21")

End Sub

Private Sub DessinerCadresBandeau(ByVal ws As Worksheet)

    ' Palette Actions
    AppliquerCadreVert ws, ws.Range("C11:O33")

    ' Palette Joueurs
    AppliquerCadreVert ws, ws.Range("C35:O45")

    ' Gestion chrono et video
    AppliquerCadreVert ws, ws.Range("C47:O53")

    AppliquerBorduresBlanchesZone ws.Range("C3:O53")
    AppliquerBorduresBlanchesZone ws.Range("P3:AI21")


End Sub

' =========================================================
' BANDEAU -> PLEIN ECRAN
' =========================================================

Public Sub PasserEnModePleinEcran()

    Dim ws As Worksheet
    Dim AncienEvents As Boolean
    Dim AncienScreenUpdating As Boolean

    On Error GoTo GestionErreur

    Set ws = shSaisieVideo

    If DetecterDispositionSaisieVideo(ws) <> MODE_BANDEAU Then

        MsgBox _
            "La feuille n'est pas en disposition Bandeau.", _
            vbExclamation

        Exit Sub

    End If

    AncienEvents = Application.EnableEvents
    AncienScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' =====================================================
    ' 1. LES ACTIONS PARTENT A DROITE
    ' =====================================================

    PlacerActionsPleinEcran ws
    CorrigerBordureFranchissement ws

    ' =====================================================
    ' 2. NETTOYAGE COMPLET DE L'ANCIENNE PALETTE ACTIONS
    '
    ' IMPORTANT :
    ' on le fait avant de remonter Joueurs et Chrono.
    ' Cela supprime notamment les restes de
    ' BTN_FRANCHISSEMENT en L24:N24.
    ' =====================================================

    NettoyerAncienneZone ws.Range("C11:O33")

    ' =====================================================
    ' 3. REMONTEE DES JOUEURS
    ' =====================================================

    PlacerJoueursPleinEcran ws

    ' =====================================================
    ' 4. REMONTEE DU CHRONO
    ' =====================================================

    PlacerChronoPleinEcran ws

    ' =====================================================
    ' 5. DIMENSIONS
    ' =====================================================

    AppliquerDimensionsPleinEcran ws

    ' =====================================================
    ' 6. SUPPRESSION DES ANCIENS BLOCS
    ' =====================================================

    NettoyerAncienneZone ws.Range("C33:O45")
    NettoyerAncienneZone ws.Range("C45:O53")

    ' =====================================================
    ' 7. REDESSIN DES CADRES
    ' =====================================================

    DessinerCadresPleinEcran ws

    If ws.Range("BTN_PL_OFF").Row <> 14 Or _
        ws.Range("BTN_PL_OFF").Column <> 20 Then

        Err.Raise _
            vbObjectError + 5000, _
            "PasserEnModePleinEcran", _
            "BTN_PL_OFF n'est pas arrive en T14."

    End If

    If ws.Range("BTN_JO_1").Row <> 14 Then

        Err.Raise _
            vbObjectError + 5001, _
            "PasserEnModePleinEcran", _
            "La palette joueurs n'est pas correctement positionnee."

    End If

    If ws.Range("CELL_CHRONO_VIDEO").Row <> 28 Then

        Err.Raise _
            vbObjectError + 5002, _
            "PasserEnModePleinEcran", _
            "Le chrono n'est pas correctement positionne."

    End If

    AppliquerBorduresBlanchesZone ws.Range("C3:O53")

    PositionnerFormesVideo
    MettreAJourBoutonSwitchMode

Sortie:

    Application.EnableEvents = AncienEvents
    Application.ScreenUpdating = AncienScreenUpdating

    Exit Sub

GestionErreur:

    Application.EnableEvents = AncienEvents
    Application.ScreenUpdating = AncienScreenUpdating

    MsgBox _
        "La disposition Plein ecran n'a pas pu etre appliquee." & _
        vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description & _
        vbCrLf & vbCrLf & _
        "N'enregistre pas ce fichier apres cette erreur.", _
        vbExclamation, _
        "Disposition saisie video"

End Sub

' =========================================================
' PLEIN ECRAN -> BANDEAU
' =========================================================

Public Sub PasserEnModeBandeau()

    Dim ws As Worksheet
    Dim AncienEvents As Boolean
    Dim AncienScreenUpdating As Boolean

    On Error GoTo GestionErreur

    Set ws = shSaisieVideo

    If DetecterDispositionSaisieVideo(ws) <> MODE_PLEIN_ECRAN Then

        MsgBox _
            "La feuille n'est pas en disposition Plein ecran.", _
            vbExclamation

        Exit Sub

    End If

    AncienEvents = Application.EnableEvents
    AncienScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Dans ce sens, il faut liberer la future palette Actions
    ' avant d'y remettre les actions.
    PlacerJoueursBandeau ws
    PlacerChronoBandeau ws
    AppliquerDimensionsBandeau ws

    NettoyerAncienneZone ws.Range("C11:O33")

    PlacerActionsBandeau ws
    CorrigerBordureFranchissement ws

    ' Nettoyage des anciennes zones Plein ecran.
    NettoyerAncienneZone ws.Range("Q3:AI21")

    DessinerCadresBandeau ws

    If ws.Range("BTN_PL_OFF").Row <> 16 Or _
        ws.Range("BTN_PL_OFF").Column <> 6 Then

        Err.Raise _
            vbObjectError + 5010, _
            "PasserEnModeBandeau", _
            "BTN_PL_OFF n'est pas revenu en F16."

    End If

    If ws.Range("BTN_JO_1").Row <> 38 Then

        Err.Raise _
            vbObjectError + 5011, _
            "PasserEnModeBandeau", _
            "La palette joueurs n'est pas correctement revenue."

    End If

    If ws.Range("CELL_CHRONO_VIDEO").Row <> 52 Then

        Err.Raise _
            vbObjectError + 5012, _
            "PasserEnModeBandeau", _
            "Le chrono n'est pas correctement revenu."

    End If

    PositionnerFormesVideo
    MettreAJourBoutonSwitchMode

Sortie:

    Application.EnableEvents = AncienEvents
    Application.ScreenUpdating = AncienScreenUpdating

    Exit Sub

GestionErreur:

    Application.EnableEvents = AncienEvents
    Application.ScreenUpdating = AncienScreenUpdating

    MsgBox _
        "La disposition Bandeau n'a pas pu etre appliquee." & _
        vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description & _
        vbCrLf & vbCrLf & _
        "N'enregistre pas ce fichier apres cette erreur.", _
        vbExclamation, _
        "Disposition saisie video"

End Sub

Private Sub NettoyerAncienneZone(ByVal Zone As Range)

    On Error Resume Next

    Zone.UnMerge
    Zone.ClearContents
    Zone.ClearFormats
    Zone.Interior.Pattern = xlNone
    Zone.Borders.LineStyle = xlLineStyleNone

    On Error GoTo 0

End Sub

Private Sub AppliquerBorduresBlanchesZone(ByVal Zone As Range)

    Dim Cellule As Range
    Dim Bordure As Variant

    For Each Cellule In Zone.Cells

        For Each Bordure In Array( _
            xlEdgeLeft, _
            xlEdgeTop, _
            xlEdgeBottom, _
            xlEdgeRight _
        )

            If Cellule.Borders(Bordure).LineStyle = xlLineStyleNone Then

                With Cellule.Borders(Bordure)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(255, 255, 255)
                End With

            End If

        Next Bordure

    Next Cellule

End Sub

Private Sub CorrigerBordureFranchissement(ByVal ws As Worksheet)

    Dim Zone As Range
    Dim B As Variant

    Set Zone = ws.Range("BTN_FRANCHISSEMENT").MergeArea

    For Each B In Array( _
        xlEdgeLeft, _
        xlEdgeTop, _
        xlEdgeBottom, _
        xlEdgeRight _
    )

        With Zone.Borders(B)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With

    Next B

End Sub

Private Sub PositionnerFormesVideo()

    Dim ws As Worksheet
    Dim z As Range

    Set ws = shSaisieVideo
    Set z = ws.Range("CELL_CHRONO_VIDEO").MergeArea

    ' Bouton Connecter lecteur video
    With ws.Shapes("BTN_CONNECTER_VIDEO")
        .LockAspectRatio = msoFalse
        .Width = 139
        .Height = 55
        .Left = z.Left - 2.75
        .Top = z.Top + 58
    End With

    ' Texte Etat connexion
    With ws.Shapes("ETAT_CONNEXION_VIDEO")
        .LockAspectRatio = msoFalse
        .Width = 230
        .Height = 63
        .Left = z.Left + 166
        .Top = z.Top + 63
    End With

    ' Pastille
    With ws.Shapes("PASTILLE_CONNEXION_VIDEO")
        .LockAspectRatio = msoTrue
        .Width = 13.75
        .Height = 13.68
        .Left = z.Left + 152.5
        .Top = z.Top + 70
    End With

End Sub

Public Sub MettreAJourBoutonSwitchMode()

    Dim ws As Worksheet
    Dim ModeActuel As String
    Dim shp As Shape

    Set ws = shSaisieVideo
    Set shp = ws.Shapes("BTN_SWITCH_MODE")

    ModeActuel = DetecterDispositionSaisieVideo(ws)

    Select Case ModeActuel

        Case MODE_BANDEAU
            shp.TextFrame2.TextRange.Text = "Mode plein ecran"

        Case MODE_PLEIN_ECRAN
            shp.TextFrame2.TextRange.Text = "Mode bandeau"

        Case Else
            shp.TextFrame2.TextRange.Text = "Changer de mode"

    End Select

End Sub

Public Function PlageBoutonsJoueurs( _
    Optional ByVal ws As Worksheet _
) As Range

    Dim Cible As Worksheet
    Dim NomsBoutons As Variant
    Dim Element As Variant
    Dim Bouton As Range
    Dim Total As Range

    If ws Is Nothing Then
        Set Cible = ThisWorkbook.Worksheets("Saisie vid" & ChrW(233) & "o")
    Else
        Set Cible = ws
    End If

    NomsBoutons = Array( _
        "BTN_JO_1", "BTN_JO_2", "BTN_JO_3", _
        "BTN_JO_4", "BTN_JO_5", "BTN_JO_6", _
        "BTN_JO_7", "BTN_JO_8", "BTN_JO_9", _
        "BTN_JO_10", "BTN_JO_11", "BTN_JO_12", _
        "BTN_JO_13", "BTN_JO_14", "BTN_JO_15", _
        "BTN_JO_16", "BTN_JO_17", "BTN_JO_18", _
        "BTN_JO_19", "BTN_JO_20", "BTN_JO_21", _
        "BTN_JO_22", _
        "BTN_JO_COLLECTIF", "BTN_JO_NON_IDENTIFIE" _
    )

    For Each Element In NomsBoutons

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

