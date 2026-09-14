Attribute VB_Name = "modStatsPossession"
Option Explicit

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
                        CompteursNous, _
                        CompteursAdv

                    SequenceOuverte = False
                    EquipeSequence = ""

                End If

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
                        CompteursNous, _
                        CompteursAdv

                End If

                SequenceOuverte = False
                EquipeSequence = ""

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
            ) Then

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

    ElseIf StrComp( _
        Equipe, _
        "Adv", _
        vbTextCompare _
    ) = 0 Then

        CompteursAdv(Classe) = _
            CompteursAdv(Classe) + 1

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

    With wsStats

        .Range("B92").Value = "Dur" & ChrW(233) & "es possessions"
        .Range("C92").Value = "Nous"
        .Range("D92").Value = "Adv"

        .Range("B93").Value = "0-20 s"
        .Range("B94").Value = "20-40 s"
        .Range("B95").Value = "40-60 s"
        .Range("B96").Value = "60-90 s"
        .Range("B97").Value = ">90 s"

        .Range("C93").Value = CompteursNous(1)
        .Range("C94").Value = CompteursNous(2)
        .Range("C95").Value = CompteursNous(3)
        .Range("C96").Value = CompteursNous(4)
        .Range("C97").Value = CompteursNous(5)

        .Range("D93").Value = CompteursAdv(1)
        .Range("D94").Value = CompteursAdv(2)
        .Range("D95").Value = CompteursAdv(3)
        .Range("D96").Value = CompteursAdv(4)
        .Range("D97").Value = CompteursAdv(5)

        ' =====================================================
        ' MISE EN FORME
        ' =====================================================

        With .Range("B92:D97")

            .Font.Name = "Calibri"
            .Font.Size = 12
            .Font.Bold = True

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin

        End With

        ' En-ttes un peu plus gros
        .Range("B92:D92").Font.Size = 13

        ' Colonne Dur_e
        With .Range("B92:B97")

            .Interior.Color = RGB(78, 167, 46)
            .Font.Color = RGB(255, 255, 255)

        End With

        ' Colonne Nous
        With .Range("C92:C97")

            .Interior.Color = RGB(226, 239, 218)
            .Font.Color = RGB(0, 0, 0)

        End With

        ' Colonne Adv
        With .Range("D92:D97")

            .Interior.Color = RGB(255, 255, 255)
            .Font.Color = RGB(0, 0, 0)

        End With

        ' Bordure ext_rieure _paisse
        With .Range("B92:D97")

            .BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThick

        End With

    End With

End Sub


Public Sub ActualiserGraphiqueSequencesPossession()

    Dim wsStats As Worksheet
    Dim Graphique As ChartObject

    Set wsStats = _
        ThisWorkbook.Worksheets("Stats match")

    On Error Resume Next

    Set Graphique = _
        wsStats.ChartObjects( _
            "GRAPHIQUE_SEQUENCES_POSSESSION" _
        )

    On Error GoTo 0

    If Graphique Is Nothing Then

        Set Graphique = _
            wsStats.ChartObjects.Add( _
                Left:=wsStats.Range("B110").Left, _
                Top:=wsStats.Range("B110").Top, _
                Width:=520, _
                Height:=300 _
            )

        Graphique.Name = _
            "GRAPHIQUE_SEQUENCES_POSSESSION"

    End If

    With Graphique.Chart

        .ChartType = xlColumnClustered

        .SetSourceData _
            Source:=wsStats.Range("B92:D97")

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


