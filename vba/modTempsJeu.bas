Attribute VB_Name = "modTempsJeu"
Option Explicit

Public Function CalculerTempsJeuEffectif( _
    ByVal MiTempsRecherchee As String _
) As Double

    Dim wsJournal As Worksheet
    Dim loJournal As ListObject

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long

    Dim i As Long

    Dim TempsAction As Double
    Dim ActionTexte As String
    Dim MiTempsLigne As String

    Dim JeuDemarre As Boolean
    Dim JeuEnCours As Boolean

    Dim DebutPeriodeJeu As Double
    Dim TempsJeuCumule As Double

    Dim ActionArret As String
    Dim ActionRepriseTouche As String
    Dim ActionRepriseMelee As String
    Dim ActionRepriseRenvoi As String
    Dim ActionRepriseMain As String

    Application.Volatile True

    On Error GoTo GestionErreur

    Set wsJournal = ThisWorkbook.Worksheets("Journal actions")
    Set loJournal = wsJournal.ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then
        CalculerTempsJeuEffectif = 0
        Exit Function
    End If

    ColTemps = _
        loJournal.ListColumns("Temps vid_o").Index

    ColMiTemps = _
        loJournal.ListColumns("Mi-temps").Index

    ColAction = _
        loJournal.ListColumns("Action").Index

    ActionArret = _
        CStr(ThisWorkbook.Names("ACT_ARRET_DU_JEU") _
            .RefersToRange.Value)

    ActionRepriseTouche = _
        CStr(ThisWorkbook.Names("ACT_REPRISE_TOUCHE") _
            .RefersToRange.Value)

    ActionRepriseMelee = _
        CStr(ThisWorkbook.Names("ACT_REPRISE_MELEE") _
            .RefersToRange.Value)

    ActionRepriseRenvoi = _
        CStr(ThisWorkbook.Names("ACT_REPRISE_RENVOI") _
            .RefersToRange.Value)

    ActionRepriseMain = _
        CStr(ThisWorkbook.Names("ACT_REPRISE_A_LA_MAIN") _
            .RefersToRange.Value)

    ' Les actions les plus r_centes sont ajout_es en haut
    ' du journal. On parcourt donc le tableau de bas en haut
    ' afin de traiter les _v_nements dans l'ordre chronologique.
    For i = loJournal.ListRows.Count To 1 Step -1

        MiTempsLigne = _
            Trim(CStr( _
                loJournal.DataBodyRange.Cells( _
                    i, _
                    ColMiTemps _
                ).Value _
            ))

        If StrComp( _
            MiTempsLigne, _
            MiTempsRecherchee, _
            vbTextCompare _
        ) = 0 Then

            ActionTexte = _
                Trim(CStr( _
                    loJournal.DataBodyRange.Cells( _
                        i, _
                        ColAction _
                    ).Value _
                ))

            If IsNumeric( _
                loJournal.DataBodyRange.Cells( _
                    i, _
                    ColTemps _
                ).Value _
            ) Then

                TempsAction = _
                    CDbl( _
                        loJournal.DataBodyRange.Cells( _
                            i, _
                            ColTemps _
                        ).Value _
                    )

                ' Le premier renvoi d_marre la mi-temps.
                If Not JeuDemarre Then

                    If StrComp( _
                        ActionTexte, _
                        ActionRepriseRenvoi, _
                        vbTextCompare _
                    ) = 0 Then

                        JeuDemarre = True
                        JeuEnCours = True
                        DebutPeriodeJeu = TempsAction

                    End If

                ElseIf JeuEnCours Then

                    ' Un arr_t clªt la p_riode de jeu en cours.
                    If StrComp( _
                        ActionTexte, _
                        ActionArret, _
                        vbTextCompare _
                    ) = 0 Then

                        If TempsAction >= DebutPeriodeJeu Then

                            TempsJeuCumule = _
                                TempsJeuCumule + _
                                TempsAction - DebutPeriodeJeu

                        End If

                        JeuEnCours = False

                    End If

                Else

                    ' Le jeu _tait arr_t_ :
                    ' une action de reprise lance une nouvelle p_riode.
                    If EstActionReprise( _
                        ActionTexte, _
                        ActionRepriseTouche, _
                        ActionRepriseMelee, _
                        ActionRepriseRenvoi, _
                        ActionRepriseMain _
                    ) Then

                        DebutPeriodeJeu = TempsAction
                        JeuEnCours = True

                    End If

                End If

            End If

        End If

    Next i

    CalculerTempsJeuEffectif = TempsJeuCumule
    Exit Function

GestionErreur:

    CalculerTempsJeuEffectif = 0

End Function


Private Function EstActionReprise( _
    ByVal ActionTexte As String, _
    ByVal ActionTouche As String, _
    ByVal ActionMelee As String, _
    ByVal ActionRenvoi As String, _
    ByVal ActionMain As String _
) As Boolean

    EstActionReprise = _
        StrComp( _
            ActionTexte, _
            ActionTouche, _
            vbTextCompare _
        ) = 0 _
        Or _
        StrComp( _
            ActionTexte, _
            ActionMelee, _
            vbTextCompare _
        ) = 0 _
        Or _
        StrComp( _
            ActionTexte, _
            ActionRenvoi, _
            vbTextCompare _
        ) = 0 _
        Or _
        StrComp( _
            ActionTexte, _
            ActionMain, _
            vbTextCompare _
        ) = 0

End Function

Public Sub RecalculerTempsJeuEffectif()

    Dim TempsMT1 As Double
    Dim TempsMT2 As Double

    On Error GoTo 0

    TempsMT1 = CalculerTempsJeuEffectifMiTemps("MT1")
    TempsMT2 = CalculerTempsJeuEffectifMiTemps("MT2")

    With ThisWorkbook

        .Names("STAT_TEMPS_EFFECTIF_MT1") _
            .RefersToRange.Value = TempsMT1

        .Names("STAT_TEMPS_EFFECTIF_MT2") _
            .RefersToRange.Value = TempsMT2

        .Names("STAT_TEMPS_EFFECTIF_TOTAL") _
            .RefersToRange.Value = TempsMT1 + TempsMT2

    End With

    Exit Sub

GestionErreur:

    MsgBox _
        "Le temps de jeu effectif n'a pas pu etre recalcule." & _
        vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Calcul du temps effectif"

End Sub


Private Function CalculerTempsJeuEffectifMiTemps( _
    ByVal MiTempsRecherchee As String _
) As Double

    Dim wsJournal As Worksheet
    Dim loJournal As ListObject

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long

    Dim i As Long
    Dim NbEvenements As Long

    Dim TempsEvenements() As Double
    Dim ActionsEvenements() As String
    Dim LignesSource() As Long

    Dim TempsAction As Double
    Dim ActionTexte As String
    Dim MiTempsLigne As String

    Dim JeuDemarre As Boolean
    Dim JeuEnCours As Boolean

    Dim DebutPeriodeJeu As Double
    Dim TempsJeuCumule As Double

    Dim ActionArret As String
    Dim ActionRepriseTouche As String
    Dim ActionRepriseMelee As String
    Dim ActionRepriseRenvoi As String
    Dim ActionRepriseMain As String

    Set wsJournal = _
        ThisWorkbook.Worksheets("Journal actions")

    Set loJournal = _
        wsJournal.ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then

        CalculerTempsJeuEffectifMiTemps = 0
        Exit Function

    End If

    ColTemps = _
        loJournal.ListColumns("temps video").Index

    ColMiTemps = _
        loJournal.ListColumns("Mi-temps").Index

    ColAction = _
        loJournal.ListColumns("Action").Index

    ActionArret = _
        CStr(Range("ACT_ARRET_DU_JEU").Value)

    ActionRepriseTouche = _
        CStr(Range("ACT_REPRISE_TOUCHE").Value)

    ActionRepriseMelee = _
        CStr(Range("ACT_REPRISE_MELEE").Value)

    ActionRepriseRenvoi = _
        CStr(Range("ACT_REPRISE_RENVOI").Value)

    ActionRepriseMain = _
        CStr(Range("ACT_REPRISE_A_LA_MAIN").Value)

    ' =====================================================
    ' 1. RECUPERATION DES EVENEMENTS DE LA MI-TEMPS
    ' =====================================================

    For i = 1 To loJournal.ListRows.Count

        MiTempsLigne = _
            Trim(CStr( _
                loJournal.DataBodyRange.Cells( _
                    i, _
                    ColMiTemps _
                ).Value _
            ))

        If StrComp( _
            MiTempsLigne, _
            MiTempsRecherchee, _
            vbTextCompare _
        ) = 0 Then

            If IsNumeric( _
                loJournal.DataBodyRange.Cells( _
                    i, _
                    ColTemps _
                ).Value _
            ) Then

                ActionTexte = _
                    Trim(CStr( _
                        loJournal.DataBodyRange.Cells( _
                            i, _
                            ColAction _
                        ).Value _
                    ))

                If ActionTexte <> "" Then

                    NbEvenements = NbEvenements + 1

                    ReDim Preserve _
                        TempsEvenements(1 To NbEvenements)

                    ReDim Preserve _
                        ActionsEvenements(1 To NbEvenements)

                    ReDim Preserve _
                        LignesSource(1 To NbEvenements)

                    TempsEvenements(NbEvenements) = _
                        CDbl( _
                            loJournal.DataBodyRange.Cells( _
                                i, _
                                ColTemps _
                            ).Value _
                        )

                    ActionsEvenements(NbEvenements) = _
                        ActionTexte

                    LignesSource(NbEvenements) = i

                End If

            End If

        End If

    Next i

    If NbEvenements = 0 Then

        CalculerTempsJeuEffectifMiTemps = 0
        Exit Function

    End If

    ' =====================================================
    ' 2. TRI PAR TEMPS VIDEO CROISSANT
    ' =====================================================

    TrierEvenementsChronologiquement _
        TempsEvenements, _
        ActionsEvenements, _
        LignesSource, _
        NbEvenements

    ' =====================================================
    ' 3. CALCUL DU TEMPS DE JEU EFFECTIF
    ' =====================================================

    For i = 1 To NbEvenements

        TempsAction = _
            TempsEvenements(i)

        ActionTexte = _
            ActionsEvenements(i)

        If Not JeuDemarre Then

            ' Le premier renvoi demarre la mi-temps.

            If StrComp( _
                ActionTexte, _
                ActionRepriseRenvoi, _
                vbTextCompare _
            ) = 0 Then

                JeuDemarre = True
                JeuEnCours = True
                DebutPeriodeJeu = TempsAction

            End If

        ElseIf JeuEnCours Then

            ' L'arret du jeu termine la periode active.

            If StrComp( _
                ActionTexte, _
                ActionArret, _
                vbTextCompare _
            ) = 0 Then

                If TempsAction >= DebutPeriodeJeu Then

                    TempsJeuCumule = _
                        TempsJeuCumule + _
                        TempsAction - DebutPeriodeJeu

                End If

                JeuEnCours = False

            End If

        Else

            ' Une reprise demarre une nouvelle periode active.

            If EstRepriseJeu( _
                ActionTexte, _
                ActionRepriseTouche, _
                ActionRepriseMelee, _
                ActionRepriseRenvoi, _
                ActionRepriseMain _
            ) Then

                DebutPeriodeJeu = TempsAction
                JeuEnCours = True

            End If

        End If

    Next i

    CalculerTempsJeuEffectifMiTemps = _
        TempsJeuCumule

End Function

Private Sub TrierEvenementsChronologiquement( _
    ByRef TempsEvenements() As Double, _
    ByRef ActionsEvenements() As String, _
    ByRef LignesSource() As Long, _
    ByVal NbEvenements As Long _
)

    Dim i As Long
    Dim j As Long

    Dim TempsTemporaire As Double
    Dim ActionTemporaire As String
    Dim LigneTemporaire As Long

    Dim DoitPermuter As Boolean

    For i = 1 To NbEvenements - 1

        For j = i + 1 To NbEvenements

            DoitPermuter = False

            If TempsEvenements(j) < _
                TempsEvenements(i) Then

                DoitPermuter = True

            ElseIf Abs( _
                TempsEvenements(j) - _
                TempsEvenements(i) _
            ) < 1E-09 Then

                ' Lorsque deux actions ont exactement le meme temps,
                ' la ligne la plus basse du journal est la plus ancienne,
                ' car les nouvelles actions sont inserees en haut.

                If LignesSource(j) > _
                    LignesSource(i) Then

                    DoitPermuter = True

                End If

            End If

            If DoitPermuter Then

                TempsTemporaire = _
                    TempsEvenements(i)

                TempsEvenements(i) = _
                    TempsEvenements(j)

                TempsEvenements(j) = _
                    TempsTemporaire

                ActionTemporaire = _
                    ActionsEvenements(i)

                ActionsEvenements(i) = _
                    ActionsEvenements(j)

                ActionsEvenements(j) = _
                    ActionTemporaire

                LigneTemporaire = _
                    LignesSource(i)

                LignesSource(i) = _
                    LignesSource(j)

                LignesSource(j) = _
                    LigneTemporaire

            End If

        Next j

    Next i

End Sub

Private Function EstRepriseJeu( _
    ByVal ActionTexte As String, _
    ByVal ActionTouche As String, _
    ByVal ActionMelee As String, _
    ByVal ActionRenvoi As String, _
    ByVal ActionMain As String _
) As Boolean

    EstRepriseJeu = _
        StrComp(ActionTexte, ActionTouche, vbTextCompare) = 0 _
        Or _
        StrComp(ActionTexte, ActionMelee, vbTextCompare) = 0 _
        Or _
        StrComp(ActionTexte, ActionRenvoi, vbTextCompare) = 0 _
        Or _
        StrComp(ActionTexte, ActionMain, vbTextCompare) = 0

End Function

