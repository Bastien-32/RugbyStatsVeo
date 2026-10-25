Attribute VB_Name = "modStatsPossession"
Option Explicit

' Temps de possession cumule par mi-temps, en fractions de
' jour comme les temps du journal. Indice 1 = MT1, 2 = MT2.
Private PossessionNous(1 To 2) As Double

Private PossessionAdv(1 To 2) As Double

' Tranches du tableau des durees de possession :
' 0-20, 20-40, 40-60, 60-90, >90. Celles des durees
' d'actions sont plus fines et vivent dans modDureesJeu.
Private Const NB_TRANCHES_POSSESSION As Long = 5

' Le graphique se place sous le tableau, a cette distance
' de son en-tete.
Private Const DECALAGE_GRAPHIQUE As Long = 18

Public Sub RecalculerStatsMatch()

    RecalculerTempsJeuEffectif
    RecalculerStatsSequencesPossession
    RecalculerStatsDureesJeu
    ActualiserGraphiqueSequencesPossession

End Sub


Public Sub RecalculerStatsSequencesPossession()

    Dim wsJournal As Worksheet
    Dim wsStats As Worksheet
    Dim loJournal As ListObject

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long
    Dim ColPossession As Long

    Dim i As Long
    Dim NbEvenements As Long

    Dim Temps() As Double
    Dim MiTemps() As String
    Dim Actions() As String
    Dim Possessions() As String
    Dim LignesSource() As Long

    Dim CompteursNous(1 To 5) As Long
    Dim CompteursAdv(1 To 5) As Long

    PossessionNous(1) = 0
    PossessionNous(2) = 0
    PossessionAdv(1) = 0
    PossessionAdv(2) = 0

    Set wsJournal = _
        ThisWorkbook.Worksheets("Journal actions")

    Set wsStats = _
        ThisWorkbook.Worksheets("Stats match")

    Set loJournal = _
        wsJournal.ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then

        EcrireTableauSequencesPossession _
            wsStats, _
            CompteursNous, _
            CompteursAdv
            
        EcrireTableauPossession wsStats

        Exit Sub

    End If

    ColTemps = _
        loJournal.ListColumns("temps video").Index

    ColMiTemps = _
        loJournal.ListColumns("Mi-temps").Index

    ColAction = _
        loJournal.ListColumns("Action").Index

    ColPossession = _
        loJournal.ListColumns("Possession").Index

    ' =====================================================
    ' RECUPERATION DES EVENEMENTS
    ' =====================================================

    For i = 1 To loJournal.ListRows.Count

        If IsNumeric( _
            loJournal.DataBodyRange.Cells( _
                i, _
                ColTemps _
            ).Value _
        ) Then

            NbEvenements = NbEvenements + 1

            ReDim Preserve _
                Temps(1 To NbEvenements)

            ReDim Preserve _
                MiTemps(1 To NbEvenements)

            ReDim Preserve _
                Actions(1 To NbEvenements)

            ReDim Preserve _
                Possessions(1 To NbEvenements)

            ReDim Preserve _
                LignesSource(1 To NbEvenements)

            Temps(NbEvenements) = _
                CDbl( _
                    loJournal.DataBodyRange.Cells( _
                        i, _
                        ColTemps _
                    ).Value _
                )

            MiTemps(NbEvenements) = _
                Trim(CStr( _
                    loJournal.DataBodyRange.Cells( _
                        i, _
                        ColMiTemps _
                    ).Value _
                ))

            Actions(NbEvenements) = _
                Trim(CStr( _
                    loJournal.DataBodyRange.Cells( _
                        i, _
                        ColAction _
                    ).Value _
                ))

            Possessions(NbEvenements) = _
                Trim(CStr( _
                    loJournal.DataBodyRange.Cells( _
                        i, _
                        ColPossession _
                    ).Value _
                ))

            LignesSource(NbEvenements) = i

        End If

    Next i

    If NbEvenements = 0 Then

        EcrireTableauSequencesPossession _
            wsStats, _
            CompteursNous, _
            CompteursAdv

        Exit Sub

    End If

    ' =====================================================
    ' TRI CHRONOLOGIQUE
    ' =====================================================

    TrierEvenementsPossession _
        Temps, _
        MiTemps, _
        Actions, _
        Possessions, _
        LignesSource, _
        NbEvenements

    ' =====================================================
    ' ANALYSE DE MT1
    ' =====================================================

    AnalyserSequencesMiTemps _
        "MT1", _
        Temps, _
        MiTemps, _
        Actions, _
        Possessions, _
        NbEvenements, _
        CompteursNous, _
        CompteursAdv

    ' =====================================================
    ' ANALYSE DE MT2
    ' =====================================================

    AnalyserSequencesMiTemps _
        "MT2", _
        Temps, _
        MiTemps, _
        Actions, _
        Possessions, _
        NbEvenements, _
        CompteursNous, _
        CompteursAdv

    ' =====================================================
    ' ECRITURE DES RESULTATS
    ' =====================================================

    EcrireTableauSequencesPossession _
        wsStats, _
        CompteursNous, _
        CompteursAdv
        
    EcrireTableauPossession wsStats

End Sub


Private Sub AnalyserSequencesMiTemps( _
    ByVal MiTempsRecherchee As String, _
    ByRef Temps() As Double, _
    ByRef MiTemps() As String, _
    ByRef Actions() As String, _
    ByRef Possessions() As String, _
    ByVal NbEvenements As Long, _
    ByRef CompteursNous() As Long, _
    ByRef CompteursAdv() As Long _
)

    Dim i As Long

    Dim SequenceOuverte As Boolean
    Dim EquipeSequence As String
    Dim DebutSequence As Double
    Dim TempsAction As Double

    Dim ActionTexte As String
    Dim PossessionAction As String
    
    Dim IndiceMiTemps As Long
    Dim JeuArrete As Boolean

    JeuArrete = True
    IndiceMiTemps = IIf( _
        UCase(MiTempsRecherchee) = "MT2", 2, 1)


    For i = 1 To NbEvenements

        If StrComp( _
            MiTemps(i), _
            MiTempsRecherchee, _
            vbTextCompare _
        ) = 0 Then

            TempsAction = Temps(i)
            ActionTexte = Actions(i)
            PossessionAction = Possessions(i)

            ' =================================================
            ' ARRET DU JEU
            ' =================================================

            If EstArretJeuPossession( _
                ActionTexte _
            ) Then

                If SequenceOuverte Then

                    ComptabiliserSequence _
                        EquipeSequence, _
                        DebutSequence, _
                        TempsAction, _
                        IndiceMiTemps, _
                        CompteursNous, _
                        CompteursAdv

                    SequenceOuverte = False
                    EquipeSequence = ""

                End If
                JeuArrete = True

            ' =================================================
            ' REPRISE DU JEU
            ' =================================================

            ElseIf EstRepriseJeuPossession( _
                ActionTexte _
            ) Then

                ' Par securite, ferme une eventuelle sequence
                ' encore ouverte avant la reprise.
                If SequenceOuverte Then

                    ComptabiliserSequence _
                        EquipeSequence, _
                        DebutSequence, _
                        TempsAction, _
                        IndiceMiTemps, _
                        CompteursNous, _
                        CompteursAdv

                End If

                SequenceOuverte = False
                EquipeSequence = ""
                JeuArrete = False

                If PossessionValide( _
                    PossessionAction _
                ) Then

                    SequenceOuverte = True
                    EquipeSequence = PossessionAction
                    DebutSequence = TempsAction

                End If

            ' =================================================
            ' ACTION NORMALE
            ' =================================================

            ElseIf PossessionValide( _
                PossessionAction _
            ) And Not JeuArrete Then

                If Not SequenceOuverte Then

                    ' Permet de recuperer proprement une sequence
                    ' meme si une reprise manque dans le journal.
                    SequenceOuverte = True
                    EquipeSequence = PossessionAction
                    DebutSequence = TempsAction

                ElseIf StrComp( _
                    PossessionAction, _
                    EquipeSequence, _
                    vbTextCompare _
                ) <> 0 Then

                    ' Changement de possession :
                    ' fin de l'ancienne sequence et debut
                    ' immediat de la nouvelle.

                    ComptabiliserSequence _
                        EquipeSequence, _
                        DebutSequence, _
                        TempsAction, _
                        IndiceMiTemps, _
                        CompteursNous, _
                        CompteursAdv

                    EquipeSequence = PossessionAction
                    DebutSequence = TempsAction

                End If

            End If

        End If

    Next i

    ' Une sequence non fermee en fin de journal
    ' n'est volontairement pas comptabilisee.
    ' Elle peut simplement correspondre a une saisie en cours.

End Sub


Private Sub ComptabiliserSequence( _
    ByVal Equipe As String, _
    ByVal TempsDebut As Double, _
    ByVal TempsFin As Double, _
    ByVal IndiceMiTemps As Long, _
    ByRef CompteursNous() As Long, _
    ByRef CompteursAdv() As Long _
)

    Dim DureeSecondes As Double
    Dim Classe As Long

    If TempsFin < TempsDebut Then
        Exit Sub
    End If

    ' Les temps Excel sont stockes en fractions de jour.
    DureeSecondes = _
        (TempsFin - TempsDebut) * 86400#

    If DureeSecondes < 0 Then
        Exit Sub
    End If

    Classe = ClasseDureePossession( _
        DureeSecondes _
    )

    If StrComp( _
        Equipe, _
        "Nous", _
        vbTextCompare _
    ) = 0 Then

        CompteursNous(Classe) = _
            CompteursNous(Classe) + 1
            
        PossessionNous(IndiceMiTemps) = _
            PossessionNous(IndiceMiTemps) + _
            (TempsFin - TempsDebut)


    ElseIf StrComp( _
        Equipe, _
        "Adv", _
        vbTextCompare _
    ) = 0 Then

        CompteursAdv(Classe) = _
            CompteursAdv(Classe) + 1
            
        PossessionAdv(IndiceMiTemps) = _
            PossessionAdv(IndiceMiTemps) + _
            (TempsFin - TempsDebut)


    End If

End Sub


Private Function ClasseDureePossession( _
    ByVal DureeSecondes As Double _
) As Long

    If DureeSecondes < 20 Then

        ClasseDureePossession = 1

    ElseIf DureeSecondes < 40 Then

        ClasseDureePossession = 2

    ElseIf DureeSecondes < 60 Then

        ClasseDureePossession = 3

    ElseIf DureeSecondes < 90 Then

        ClasseDureePossession = 4

    Else

        ClasseDureePossession = 5

    End If

End Function


Private Function PossessionValide( _
    ByVal Equipe As String _
) As Boolean

    If StrComp( _
        Equipe, _
        "Nous", _
        vbTextCompare _
    ) = 0 Then

        PossessionValide = True
        Exit Function

    End If

    If StrComp( _
        Equipe, _
        "Adv", _
        vbTextCompare _
    ) = 0 Then

        PossessionValide = True

    End If

End Function


Private Function EstArretJeuPossession( _
    ByVal ActionTexte As String _
) As Boolean

    EstArretJeuPossession = _
        StrComp( _
            ActionTexte, _
            CStr(Range("ACT_ARRET_DU_JEU").Value), _
            vbTextCompare _
        ) = 0

End Function


Private Function EstRepriseJeuPossession( _
    ByVal ActionTexte As String _
) As Boolean

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_TOUCHE").Value), _
        vbTextCompare _
    ) = 0 Then

        EstRepriseJeuPossession = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_MELEE").Value), _
        vbTextCompare _
    ) = 0 Then

        EstRepriseJeuPossession = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_RENVOI").Value), _
        vbTextCompare _
    ) = 0 Then

        EstRepriseJeuPossession = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_A_LA_MAIN").Value), _
        vbTextCompare _
    ) = 0 Then

        EstRepriseJeuPossession = True

    End If

End Function


Private Sub TrierEvenementsPossession( _
    ByRef Temps() As Double, _
    ByRef MiTemps() As String, _
    ByRef Actions() As String, _
    ByRef Possessions() As String, _
    ByRef LignesSource() As Long, _
    ByVal NbEvenements As Long _
)

    Dim i As Long
    Dim j As Long
    Dim DoitPermuter As Boolean

    Dim TempTemps As Double
    Dim TempTexte As String
    Dim TempLigne As Long

    For i = 1 To NbEvenements - 1

        For j = i + 1 To NbEvenements

            DoitPermuter = False

            If Temps(j) < Temps(i) Then

                DoitPermuter = True

            ElseIf Abs( _
                Temps(j) - Temps(i) _
            ) < 1E-09 Then

                ' Les nouvelles actions etant inserees en haut,
                ' une ligne plus basse est plus ancienne.

                If LignesSource(j) > _
                    LignesSource(i) Then

                    DoitPermuter = True

                End If

            End If

            If DoitPermuter Then

                TempTemps = Temps(i)
                Temps(i) = Temps(j)
                Temps(j) = TempTemps

                TempTexte = MiTemps(i)
                MiTemps(i) = MiTemps(j)
                MiTemps(j) = TempTexte

                TempTexte = Actions(i)
                Actions(i) = Actions(j)
                Actions(j) = TempTexte

                TempTexte = Possessions(i)
                Possessions(i) = Possessions(j)
                Possessions(j) = TempTexte

                TempLigne = LignesSource(i)
                LignesSource(i) = LignesSource(j)
                LignesSource(j) = TempLigne

            End If

        Next j

    Next i

End Sub


Private Sub EcrireTableauSequencesPossession( _
    ByVal wsStats As Worksheet, _
    ByRef CompteursNous() As Long, _
    ByRef CompteursAdv() As Long)

    Dim L As Long
    Dim i As Long

    ' Le tableau se place sous une ancre nommee plutot qu'a
    ' une ligne ecrite en dur : ajouter une ligne dans les
    ' tableaux de joueurs, plus haut dans la feuille, le
    ' decalait sans que rien ne le sache. modDureesJeu se
    ' repere sur cet en-tete : les deux tableaux glissaient
    ' ensemble et laissaient derriere eux leurs anciennes
    ' lignes, reecrites deux lignes plus haut.
    L = LigneDurees(wsStats)

    With wsStats

        .Cells(L, "B").Value = _
            "Dur" & ChrW(233) & "es possessions"
        .Cells(L, "C").Value = "Nous"
        .Cells(L, "D").Value = "Adv"

        .Cells(L + 1, "B").Value = "0-20 s"
        .Cells(L + 2, "B").Value = "20-40 s"
        .Cells(L + 3, "B").Value = "40-60 s"
        .Cells(L + 4, "B").Value = "60-90 s"
        .Cells(L + 5, "B").Value = ">90 s"

        For i = 1 To NB_TRANCHES_POSSESSION

            .Cells(L + i, "C").Value = CompteursNous(i)
            .Cells(L + i, "D").Value = CompteursAdv(i)

        Next i

        ' =====================================================
        ' MISE EN FORME
        ' =====================================================

        With .Range( _
            .Cells(L, "B"), _
            .Cells(L + NB_TRANCHES_POSSESSION, "D"))

            .Font.Name = "Calibri"
            .Font.Size = 12
            .Font.Bold = True

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin

        End With

        ' En-tetes un peu plus gros
        .Range(.Cells(L, "B"), .Cells(L, "D")).Font.Size = 13

        ' Colonne Duree
        With .Range( _
            .Cells(L, "B"), _
            .Cells(L + NB_TRANCHES_POSSESSION, "B"))

            .Interior.Color = RGB(78, 167, 46)
            .Font.Color = RGB(255, 255, 255)

        End With

        ' Colonne Nous
        With .Range( _
            .Cells(L, "C"), _
            .Cells(L + NB_TRANCHES_POSSESSION, "C"))

            .Interior.Color = RGB(226, 239, 218)
            .Font.Color = RGB(0, 0, 0)

        End With

        ' Colonne Adv
        With .Range( _
            .Cells(L, "D"), _
            .Cells(L + NB_TRANCHES_POSSESSION, "D"))

            .Interior.Color = RGB(255, 255, 255)
            .Font.Color = RGB(0, 0, 0)

        End With

        ' Bordure exterieure epaisse
        With .Range( _
            .Cells(L, "B"), _
            .Cells(L + NB_TRANCHES_POSSESSION, "D"))

            .BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlMedium

        End With

    End With

End Sub


Public Sub ActualiserGraphiqueSequencesPossession()

    Dim wsStats As Worksheet
    Dim Graphique As ChartObject
    Dim LigneAncre As Long
    Dim LigneGraphique As Long

    Set wsStats = _
        ThisWorkbook.Worksheets("Stats match")

    LigneAncre = LigneDurees(wsStats)
    LigneGraphique = LigneAncre + DECALAGE_GRAPHIQUE

    On Error Resume Next

    Set Graphique = _
        wsStats.ChartObjects( _
            "GRAPHIQUE_SEQUENCES_POSSESSION" _
        )

    On Error GoTo 0

    If Graphique Is Nothing Then

        Set Graphique = _
            wsStats.ChartObjects.Add( _
                Left:=wsStats.Cells(LigneGraphique, "B").Left, _
                Top:=wsStats.Cells(LigneGraphique, "B").Top, _
                Width:=520, _
                Height:=300 _
            )

        Graphique.Name = _
            "GRAPHIQUE_SEQUENCES_POSSESSION"

    End If

    With Graphique.Chart

        .ChartType = xlColumnClustered

        .SetSourceData _
            Source:=wsStats.Range( _
                wsStats.Cells(LigneAncre, "B"), _
                wsStats.Cells( _
                    LigneAncre + NB_TRANCHES_POSSESSION, "D"))

        .HasTitle = True

        .ChartTitle.Text = _
            "Duree des sequences de possession"

        .HasLegend = True

        .Legend.Position = xlLegendPositionBottom

        ' Permet au graphique de continuer a utiliser
        ' les cellules source lorsqu'on masque Z:AB.
        .PlotVisibleOnly = False

        If .SeriesCollection.Count >= 2 Then

            .SeriesCollection(1).Name = "Nous"

            .SeriesCollection(1) _
                .Format.Fill.ForeColor.RGB = _
                RGB(0, 128, 0)

            .SeriesCollection(1) _
                .Format.Line.ForeColor.RGB = _
                RGB(0, 128, 0)

            .SeriesCollection(2).Name = "Adversaire"

            .SeriesCollection(2) _
                .Format.Fill.ForeColor.RGB = _
                RGB(0, 0, 0)

            .SeriesCollection(2) _
                .Format.Line.ForeColor.RGB = _
                RGB(0, 0, 0)

        End If

        On Error Resume Next

        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MajorUnit = 5
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = _
            "Nombre de sequences"

        On Error GoTo 0

    End With

End Sub
' Temps de possession par mi-temps, a cote du temps de jeu
' effectif : H pour nous, I pour eux, J restant au total
' deja calcule par modTempsJeu.
Private Sub EcrireTableauPossession( _
    ByVal wsStats As Worksheet _
)

    Dim i As Long
    Dim L As Long

    ' Le tableau se place sous une ancre nommee plutot
    ' qu'a une ligne ecrite en dur : ajouter une ligne
    ' plus haut dans la feuille le decalait sans que rien
    ' ne le sache, et les durees s'ecrivaient par-dessus
    ' les melees.
    L = LignePossession(wsStats)

    With wsStats

        .Cells(L, "H").Value = "Nous"
        .Cells(L, "I").Value = "Eux"
        .Cells(L, "J").Value = "Total"

        For i = 1 To 2

            .Cells(L + i, "H").Value = PossessionNous(i)
            .Cells(L + i, "I").Value = PossessionAdv(i)

        Next i

        .Cells(L + 3, "H").Value = _
            PossessionNous(1) + PossessionNous(2)

        .Cells(L + 3, "I").Value = _
            PossessionAdv(1) + PossessionAdv(2)

        ' Les formats sont repris des colonnes voisines :
        ' en-tete vert, lignes de donnees, ligne de total.
        .Cells(L, "G").Copy
        .Range(.Cells(L, "H"), .Cells(L, "J")) _
            .PasteSpecial Paste:=xlPasteFormats

        For i = 1 To 3

            .Cells(L + i, "J").Copy
            .Range(.Cells(L + i, "H"), .Cells(L + i, "I")) _
                .PasteSpecial Paste:=xlPasteFormats

        Next i

        Application.CutCopyMode = False

        ' Trait de separation entre le jeu effectif et la
        ' possession.
        With .Range(.Cells(L, "H"), .Cells(L + 3, "H")) _
            .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With

        .Range(.Cells(L, "G"), .Cells(L, "J")) _
            .HorizontalAlignment = xlCenter

        .Cells(L + 3, "G").HorizontalAlignment = xlRight

        ' Le format d'en-tete repris de G n'a pas de bord
        ' droit : le cadre resterait ouvert sur le titre.
        With .Cells(L, "J").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With

    End With

End Sub


' =========================================================
' Ou commence le tableau des durees de possession.
'
' L'ancre est une plage nommee : elle suit les lignes
' ajoutees ou retirees au-dessus d'elle, ce qu'une adresse
' ecrite dans le code ne sait pas faire. Absente, elle est
' creee a sa place d'origine.
' =========================================================
Private Function LignePossession( _
    ByVal wsStats As Worksheet) As Long

    Const NOM_ANCRE As String = "STAT_POSSESSION_ANCRE"
    Const LIGNE_ORIGINE As Long = 44

    On Error GoTo Defaut

    LignePossession = ThisWorkbook.Names(NOM_ANCRE) _
        .RefersToRange.Row

    If LignePossession > 0 Then
        Exit Function
    End If

Defaut:

    On Error Resume Next

    ThisWorkbook.Names.Add _
        Name:=NOM_ANCRE, _
        RefersTo:=wsStats.Cells(LIGNE_ORIGINE, "H")

    On Error GoTo 0

    LignePossession = LIGNE_ORIGINE

End Function


' =========================================================
' Ou commence le tableau des durees de possession.
'
' Meme principe que LignePossession : une plage nommee suit
' les lignes ajoutees au-dessus d'elle, une adresse ecrite
' dans le code ne le sait pas. Absente, elle est creee a sa
' place d'origine.
'
' modDureesJeu cherche cet en-tete pour poser le tableau des
' durees d'actions a cote : les deux restent alignes.
' =========================================================
Private Function LigneDurees( _
    ByVal wsStats As Worksheet) As Long

    Const NOM_ANCRE As String = "STAT_DUREES_ANCRE"
    Const LIGNE_ORIGINE As Long = 94

    On Error GoTo Defaut

    LigneDurees = ThisWorkbook.Names(NOM_ANCRE) _
        .RefersToRange.Row

    If LigneDurees > 0 Then
        Exit Function
    End If

Defaut:

    On Error Resume Next

    ThisWorkbook.Names.Add _
        Name:=NOM_ANCRE, _
        RefersTo:=wsStats.Cells(LIGNE_ORIGINE, "B")

    On Error GoTo 0

    LigneDurees = LIGNE_ORIGINE

End Function
