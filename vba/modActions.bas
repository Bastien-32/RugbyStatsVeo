Attribute VB_Name = "modActions"
Option Explicit

Public Function GetAction( _
    ByVal Target As Range) As String

    If Not Intersect(Target, Range("BTN_PL_OFF")) Is Nothing Then
        GetAction = Range("ACT_PL_OFF").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_NORMAL")) Is Nothing Then
        GetAction = Range("ACT_PL_NORMAL").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_RATE")) Is Nothing Then
        GetAction = Range("ACT_PL_RATE").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_A2")) Is Nothing Then
        GetAction = Range("ACT_PL_A2").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_HAUT")) Is Nothing Then
        GetAction = Range("ACT_PL_HAUT").Value
        Exit Function
    End If
    
    If Not Intersect(Target, Range("BTN_TURNOVER")) Is Nothing Then

        Select Case CurrentTeam

            Case "Nous"

                ' Nous avions le ballon : nous le perdons.
                GetAction = Range("ACT_TURNOVER_ADV").Value

            Case "Adv"

                ' Ils avaient le ballon : nous le recuperons.
                GetAction = Range("ACT_TURNOVER_NS").Value

            Case Else

                GetAction = ""

        End Select

        Exit Function

    End If

    If Not Intersect(Target, Range("BTN_TO_G")) Is Nothing Then
    
        Select Case CurrentTeam
    
            Case "Nous"
    
                GetAction = Range("ACT_TO_NS_G").Value
    
            Case "Adv"
    
                GetAction = Range("ACT_TO_ADV_VOLEE").Value
    
            Case Else
    
                GetAction = ""
    
        End Select
    
        Exit Function
    
    End If
    
    
    If Not Intersect(Target, Range("BTN_TO_P")) Is Nothing Then
    
        Select Case CurrentTeam
    
            Case "Nous"
    
                GetAction = Range("ACT_TO_NS_P").Value
    
            Case "Adv"
    
                GetAction = Range("ACT_TO_ADV_PERDUE").Value
    
            Case Else
    
                GetAction = ""
    
        End Select
    
        Exit Function
    
    End If
    
    
    If Not Intersect(Target, Range("BTN_TO_PAS_DROIT")) Is Nothing Then
    
        Select Case CurrentTeam
    
            Case "Nous"
    
                GetAction = Range("ACT_TO_NS_PAS_DROIT").Value
    
            Case "Adv"
    
                GetAction = Range("ACT_TO_ADV_PAS_DROIT").Value
    
            Case Else
    
                GetAction = ""
    
        End Select
    
        Exit Function
    
    End If

    If Not Intersect(Target, Range("BTN_ME_G")) Is Nothing Then
    
        Select Case CurrentTeam
    
            Case "Nous"
    
                GetAction = Range("ACT_ME_NS_G").Value
    
            Case "Adv"
    
                GetAction = Range("ACT_ME_ADV_G").Value
    
            Case Else
    
                GetAction = ""
    
        End Select
    
        Exit Function
    
    End If
    
    
    If Not Intersect(Target, Range("BTN_ME_P")) Is Nothing Then
    
        Select Case CurrentTeam
    
            Case "Nous"
    
                GetAction = Range("ACT_ME_NS_P").Value
    
            Case "Adv"
    
                GetAction = Range("ACT_ME_ADV_P").Value
    
            Case Else
    
                GetAction = ""
    
        End Select
    
        Exit Function
    
    End If

    If Not Intersect(Target, Range("BTN_PTS_ESSAI")) Is Nothing Then
        GetAction = Range("ACT_PTS_ESSAI").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PTS_ESSAI_PEN")) Is Nothing Then
        GetAction = Range("ACT_PTS_ESSAI_PEN").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PTS_PENALITE")) Is Nothing Then
        GetAction = Range("ACT_PTS_PENALITE").Value
        Exit Function
    End If
    
    If Not Intersect(Target, Range("BTN_PTS_TRANSFO")) Is Nothing Then
        GetAction = Range("ACT_PTS_TRANSFO").Value
        Exit Function
    End If


    If Not Intersect(Target, Range("BTN_PTS_DROP")) Is Nothing Then
        GetAction = Range("ACT_PTS_DROP").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_TOUCHE")) Is Nothing Then
        GetAction = Range("ACT_REPRISE_TOUCHE").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_MELEE")) Is Nothing Then
        GetAction = Range("ACT_REPRISE_MELEE").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_RENVOI")) Is Nothing Then
        GetAction = Range("ACT_REPRISE_RENVOI").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_A_LA_MAIN")) Is Nothing Then
        GetAction = Range("ACT_REPRISE_A_LA_MAIN").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_ARRET_DU_JEU")) Is Nothing Then
        GetAction = Range("ACT_ARRET_DU_JEU").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_EN_AVANT")) Is Nothing Then
        GetAction = Range("ACT_EN_AVANT").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_JEU_AU_PIED")) Is Nothing Then
        GetAction = Range("ACT_JEU_AU_PIED").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_RECEPTION_NOUS")) Is Nothing Then
        GetAction = Range("ACT_RECEPTION_NOUS").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_RECEPTION_ADV")) Is Nothing Then
        GetAction = Range("ACT_RECEPTION_ADV").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_GRATTAGE")) Is Nothing Then
        GetAction = Range("ACT_GRATTAGE").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_CONTRE_RUCK")) Is Nothing Then
        GetAction = Range("ACT_CONTRE_RUCK").Value
        Exit Function
    End If
    
    If Not Intersect(Target, Range("BTN_PEN_CONTRE_ADV")) Is Nothing Then
        GetAction = Range("ACT_PEN_CONTRE_ADV").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_CONTRE_NOUS")) Is Nothing Then
        GetAction = Range("ACT_PEN_CONTRE_NOUS").Value
        Exit Function
    End If
    
    If Not Intersect(Target, Range("BTN_ARRACHAGE")) Is Nothing Then
        GetAction = Range("ACT_ARRACHAGE").Value
        Exit Function
    End If

    ' =====================================================
    ' FRANCHISSEMENT
    ' =====================================================

    If Not Intersect( _
        Target, _
        Range("BTN_FRANCHISSEMENT") _
    ) Is Nothing Then

        GetAction = _
            Range("ACT_FRANCHISSEMENT").Value

        Exit Function

    End If

    ' =====================================================
    ' CARTON ROUGE
    ' =====================================================

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_ROUGE") _
    ) Is Nothing Then

        Select Case CurrentTeam

            Case "Nous"

                GetAction = _
                    Range("ACT_CARTON_ROUGE_ADV").Value

            Case "Adv"

                GetAction = _
                    Range("ACT_CARTON_ROUGE_NS").Value

            Case Else

                GetAction = ""

        End Select

        Exit Function

    End If

    ' =====================================================
    ' CARTON JAUNE
    ' =====================================================

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_JAUNE") _
    ) Is Nothing Then

        Select Case CurrentTeam

            Case "Nous"

                GetAction = _
                    Range("ACT_CARTON_JAUNE_ADV").Value

            Case "Adv"

                GetAction = _
                    Range("ACT_CARTON_JAUNE_NS").Value

            Case Else

                GetAction = ""

        End Select

        Exit Function

    End If

    ' =====================================================
    ' CARTON BLANC
    ' =====================================================

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_BLANC") _
    ) Is Nothing Then

        Select Case CurrentTeam

            Case "Nous"

                GetAction = _
                    Range("ACT_CARTON_BLANC_ADV").Value

            Case "Adv"

                GetAction = _
                    Range("ACT_CARTON_BLANC_NS").Value

            Case Else

                GetAction = ""

        End Select

        Exit Function

    End If

    ' =====================================================
    ' CARTON BLEU
    ' =====================================================

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_BLEU") _
    ) Is Nothing Then

        Select Case CurrentTeam

            Case "Nous"

                GetAction = _
                    Range("ACT_CARTON_BLEU_ADV").Value

            Case "Adv"

                GetAction = _
                    Range("ACT_CARTON_BLEU_NS").Value

            Case Else

                GetAction = ""

        End Select

        Exit Function

    End If
    
    GetAction = ""

End Function

Public Function IsImmediateAction(ByVal Target As Range) As Boolean

    If Not Intersect(Target, Range("BTN_PTS_ESSAI")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If
    
    If Not Intersect(Target, Range("BTN_PTS_TRANSFO")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    
    If Not Intersect(Target, Range("BTN_PTS_DROP")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PTS_ESSAI_PEN")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_TOUCHE")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_MELEE")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_RENVOI")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_REPRISE_A_LA_MAIN")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_EN_AVANT")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_JEU_AU_PIED")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_RECEPTION_NOUS")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_RECEPTION_ADV")) Is Nothing Then
        IsImmediateAction = True
        Exit Function
    End If
    
    IsImmediateAction = False

End Function

Public Function IsPlaquageAction(ByVal Target As Range) As Boolean

    If Not Intersect(Target, Range("BTN_PL_OFF")) Is Nothing Then
        IsPlaquageAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_NORMAL")) Is Nothing Then
        IsPlaquageAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_RATE")) Is Nothing Then
        IsPlaquageAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_A2")) Is Nothing Then
        IsPlaquageAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PL_HAUT")) Is Nothing Then
        IsPlaquageAction = True
        Exit Function
    End If

    IsPlaquageAction = False

End Function

Public Function IsTurnoverAction(ByVal Target As Range) As Boolean

    If Not Intersect(Target, Range("BTN_GRATTAGE")) Is Nothing Then
        IsTurnoverAction = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_CONTRE_RUCK")) Is Nothing Then
        IsTurnoverAction = True
        Exit Function
    End If

    IsTurnoverAction = False

End Function

Public Function IsPenaltyReasonButton( _
    ByVal Target As Range) As Boolean

    If Not Intersect(Target, Range("BTN_PEN_MAUL")) Is Nothing Then
        IsPenaltyReasonButton = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_RUCK")) Is Nothing Then
        IsPenaltyReasonButton = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_HORS_JEU")) Is Nothing Then
        IsPenaltyReasonButton = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_PL_A_2")) Is Nothing Then
        IsPenaltyReasonButton = True
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_PL_HAUT")) Is Nothing Then
        IsPenaltyReasonButton = True
        Exit Function
    End If

    IsPenaltyReasonButton = False

End Function


Public Function GetPenaltyReason( _
    ByVal Target As Range) As String

    If Not Intersect(Target, Range("BTN_PEN_MAUL")) Is Nothing Then
        GetPenaltyReason = Range("ACT_PEN_MAUL").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_RUCK")) Is Nothing Then
        GetPenaltyReason = Range("ACT_PEN_RUCK").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_HORS_JEU")) Is Nothing Then
        GetPenaltyReason = Range("ACT_PEN_HORS_JEU").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_PL_A_2")) Is Nothing Then
        GetPenaltyReason = Range("ACT_PEN_PL_A_2").Value
        Exit Function
    End If

    If Not Intersect(Target, Range("BTN_PEN_PL_HAUT")) Is Nothing Then
        GetPenaltyReason = Range("ACT_PEN_PL_HAUT").Value
        Exit Function
    End If

    GetPenaltyReason = ""

End Function

Public Sub InitialiserPenalite( _
    ByVal ActionBase As String, _
    ByVal TempsVideo As Double, _
    ByVal EquipeImposee As String, _
    ByVal AttendJoueur As Boolean)

    FermerActionEnAttente

    PendingPenaltyActive = True
    PendingPenaltyBaseAction = ActionBase
    PendingPenaltyTime = TempsVideo
    PendingPenaltyForcedTeam = EquipeImposee
    PendingPenaltyNeedsPlayer = AttendJoueur

End Sub


Public Sub FermerPenaliteEnAttente()

    PendingPenaltyActive = False
    PendingPenaltyBaseAction = ""
    PendingPenaltyTime = 0
    PendingPenaltyForcedTeam = ""
    PendingPenaltyNeedsPlayer = False

End Sub

Public Sub TransformerPenaliteEnAttenteJoueur( _
    ByVal Motif As String)

    Dim ActionPenalite As String
    Dim TempsPenalite As Double
    Dim EquipePenalite As String

    ActionPenalite = PendingPenaltyBaseAction
    TempsPenalite = PendingPenaltyTime
    EquipePenalite = PendingPenaltyForcedTeam

    FermerPenaliteEnAttente

    InitialiserActionEnAttente _
        ActionPenalite, _
        TempsPenalite, _
        False, _
        EquipePenalite, _
        False

    PendingMotifPenalite = Motif
    PendingGroupeFautifRequired = True

End Sub

Public Sub FinaliserPenaliteOubliee()

    Dim TempsPenalite As Double
    Dim ActionPenalite As String

    If Not PendingPenaltyActive Then
        Exit Sub
    End If

    TempsPenalite = PendingPenaltyTime
    ActionPenalite = PendingPenaltyBaseAction

    SetCurrentTeam PendingPenaltyForcedTeam

    If PendingPenaltyNeedsPlayer Then

        AjouterAction _
            "?", _
            ActionPenalite, _
            TempsPenalite, _
            "?", _
            "?"

    Else

        AjouterAction _
            "", _
            ActionPenalite, _
            TempsPenalite, _
            "?", _
            ""

    End If

    FermerPenaliteEnAttente

    AjouterArretJeuAutomatique _
        TempsPenalite

End Sub

Public Sub InitialiserActionEnAttente( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal ChangePossession As Boolean, _
    ByVal EquipeImposee As String, _
    ByVal AutoriserPlusieursJoueurs As Boolean)

    PendingAction = ActionTexte
    WaitingForPlayer = True

    PendingActionTime = TempsVideo
    PendingActionTimeAvailable = True

    PendingActionChangesPossession = ChangePossession
    PendingForcedTeam = EquipeImposee
    PendingAllowMultiplePlayers = AutoriserPlusieursJoueurs

    PendingPlayerCount = 0
    
    PendingMotifPenalite = ""
    PendingGroupeFautifRequired = False

End Sub


Public Sub FermerActionEnAttente()

    PendingAction = ""
    WaitingForPlayer = False

    PendingActionTime = 0
    PendingActionTimeAvailable = False

    PendingActionChangesPossession = False
    PendingForcedTeam = ""
    PendingAllowMultiplePlayers = False

    PendingPlayerCount = 0
    
    PendingMotifPenalite = ""
    PendingGroupeFautifRequired = False

End Sub


Public Sub FinaliserActionEnAttente()

    Dim TempsVideo As Double
    Dim GroupeFautif As String
    Dim ActionFinalisee As String
    Dim DoitAjouterArret As Boolean

    If Not WaitingForPlayer Then
        Exit Sub
    End If

    If PendingAction = "" Then
        FermerActionEnAttente
        Exit Sub
    End If

    ActionFinalisee = PendingAction

    DoitAjouterArret = _
        ActionDeclencheArretAutomatique( _
            ActionFinalisee _
        )

    If PendingPlayerCount = 0 Then

        If PendingActionTimeAvailable Then

            TempsVideo = PendingActionTime

        Else

            TempsVideo = GetTimeVideo()

        End If

        If TempsVideo >= 0 Then

            If PendingForcedTeam <> "" Then
                SetCurrentTeam PendingForcedTeam
            End If

            GroupeFautif = ""

            If PendingGroupeFautifRequired Then
                GroupeFautif = "?"
            End If

            AjouterAction _
                "?", _
                ActionFinalisee, _
                TempsVideo, _
                PendingMotifPenalite, _
                GroupeFautif

            If PendingActionChangesPossession Then
                ToggleCurrentTeam
            End If

        End If

    End If

    FermerActionEnAttente

    If DoitAjouterArret Then

        AjouterArretJeuAutomatique _
            TempsVideo

    Else

        ResetActionButtons

    End If

End Sub

Public Function EstActionPenaliteContreNousOuAdv( _
    ByVal ActionTexte As String _
) As Boolean

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_PEN_CONTRE_ADV").Value), _
        vbTextCompare _
    ) = 0 Then

        EstActionPenaliteContreNousOuAdv = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_PEN_CONTRE_NOUS").Value), _
        vbTextCompare _
    ) = 0 Then

        EstActionPenaliteContreNousOuAdv = True

    End If

End Function

Public Function ActionDeclencheArretAutomatique( _
    ByVal ActionTexte As String _
) As Boolean

    If EstActionPenaliteContreNousOuAdv(ActionTexte) Then
        ActionDeclencheArretAutomatique = True
        Exit Function
    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_PTS_ESSAI").Value), _
        vbTextCompare _
    ) = 0 Then
        ActionDeclencheArretAutomatique = True
        Exit Function
    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_PTS_DROP").Value), _
        vbTextCompare _
    ) = 0 Then
        ActionDeclencheArretAutomatique = True
    End If

End Function

Public Function EstBoutonRepriseJeu( _
    ByVal Target As Range _
) As Boolean

    If Not Intersect( _
        Target, _
        Range("BTN_REPRISE_TOUCHE") _
    ) Is Nothing Then

        EstBoutonRepriseJeu = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_REPRISE_MELEE") _
    ) Is Nothing Then

        EstBoutonRepriseJeu = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_REPRISE_RENVOI") _
    ) Is Nothing Then

        EstBoutonRepriseJeu = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_REPRISE_A_LA_MAIN") _
    ) Is Nothing Then

        EstBoutonRepriseJeu = True

    End If

End Function

Private Function EstTexteRepriseOuArret( _
    ByVal ActionTexte As String _
) As Boolean

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_ARRET_DU_JEU").Value), _
        vbTextCompare _
    ) = 0 Then

        EstTexteRepriseOuArret = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_TOUCHE").Value), _
        vbTextCompare _
    ) = 0 Then

        EstTexteRepriseOuArret = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_MELEE").Value), _
        vbTextCompare _
    ) = 0 Then

        EstTexteRepriseOuArret = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_RENVOI").Value), _
        vbTextCompare _
    ) = 0 Then

        EstTexteRepriseOuArret = True
        Exit Function

    End If

    If StrComp( _
        ActionTexte, _
        CStr(Range("ACT_REPRISE_A_LA_MAIN").Value), _
        vbTextCompare _
    ) = 0 Then

        EstTexteRepriseOuArret = True

    End If

End Function

Public Sub ActiverBoutonPenaliteEnAttente()

    ResetActionButtons

    If StrComp( _
        PendingPenaltyBaseAction, _
        CStr(Range("ACT_PEN_CONTRE_ADV").Value), _
        vbTextCompare _
    ) = 0 Then

        SetActiveButton _
            shSaisieVideo.Range("BTN_PEN_CONTRE_ADV")

        Exit Sub

    End If

    If StrComp( _
        PendingPenaltyBaseAction, _
        CStr(Range("ACT_PEN_CONTRE_NOUS").Value), _
        vbTextCompare _
    ) = 0 Then

        SetActiveButton _
            shSaisieVideo.Range("BTN_PEN_CONTRE_NOUS")

    End If

End Sub

Public Sub ActiverBoutonArretDuJeu()

    ResetActionButtons

    SetActiveButton _
        shSaisieVideo.Range("BTN_ARRET_DU_JEU")

End Sub

Public Sub AjouterArretJeuAutomatique( _
    ByVal TempsArret As Double _
)

    If Not JeuActuellementArrete Then

        AjouterAction _
            "", _
            CStr(Range("ACT_ARRET_DU_JEU").Value), _
            TempsArret

    End If

    ActiverBoutonArretDuJeu
    RecalculerTempsJeuEffectif

End Sub

Private Function ArretExisteDejaAuMemeTemps( _
    ByVal TempsRecherche As Double _
) As Boolean

    Dim loJournal As ListObject

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long

    Dim i As Long

    Dim TempsLigne As Variant
    Dim MiTempsLigne As String
    Dim ActionLigne As String
    Dim ActionArret As String

    Set loJournal = _
        ThisWorkbook _
            .Worksheets("Journal actions") _
            .ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then
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

    For i = 1 To loJournal.ListRows.Count

        TempsLigne = _
            loJournal.DataBodyRange.Cells( _
                i, _
                ColTemps _
            ).Value

        If IsNumeric(TempsLigne) Then

            If Abs( _
                CDbl(TempsLigne) - TempsRecherche _
            ) < 1E-09 Then

                MiTempsLigne = _
                    Trim(CStr( _
                        loJournal.DataBodyRange.Cells( _
                            i, _
                            ColMiTemps _
                        ).Value _
                    ))

                If StrComp( _
                    MiTempsLigne, _
                    CurrentHalf, _
                    vbTextCompare _
                ) = 0 Then

                    ActionLigne = _
                        Trim(CStr( _
                            loJournal.DataBodyRange.Cells( _
                                i, _
                                ColAction _
                            ).Value _
                        ))

                    If StrComp( _
                        ActionLigne, _
                        ActionArret, _
                        vbTextCompare _
                    ) = 0 Then

                        ArretExisteDejaAuMemeTemps = True
                        Exit Function

                    End If

                End If

            End If

        End If

    Next i

End Function

Public Function JeuActuellementArrete() As Boolean

    Dim loJournal As ListObject

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long

    Dim i As Long

    Dim TempsLigne As Variant
    Dim DernierTemps As Double

    Dim ActionLigne As String
    Dim DerniereActionEtat As String
    Dim MiTempsLigne As String
    Dim ActionArret As String

    Set loJournal = _
        ThisWorkbook _
            .Worksheets("Journal actions") _
            .ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then
        Exit Function
    End If

    If CurrentHalf = "" Then
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

    DernierTemps = -1

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
            CurrentHalf, _
            vbTextCompare _
        ) = 0 Then

            TempsLigne = _
                loJournal.DataBodyRange.Cells( _
                    i, _
                    ColTemps _
                ).Value

            If IsNumeric(TempsLigne) Then

                ActionLigne = _
                    Trim(CStr( _
                        loJournal.DataBodyRange.Cells( _
                            i, _
                            ColAction _
                        ).Value _
                    ))

                If EstTexteRepriseOuArret( _
                    ActionLigne _
                ) Then

                    If CDbl(TempsLigne) > DernierTemps Then

                        DernierTemps = CDbl(TempsLigne)

                        DerniereActionEtat = ActionLigne

                    End If

                End If

            End If

        End If

    Next i

    If StrComp( _
        DerniereActionEtat, _
        ActionArret, _
        vbTextCompare _
    ) = 0 Then

        JeuActuellementArrete = True

    End If

End Function


Public Function EstPenaliteEnAttenteJoueur() As Boolean

    EstPenaliteEnAttenteJoueur = _
        WaitingForPlayer _
        And EstActionPenaliteContreNousOuAdv(PendingAction)

End Function


Public Sub AfficherAlertePenaliteEnAttente()

    MsgBox _
        "Une penalite est en attente." & _
        vbCrLf & vbCrLf & _
        "Tu dois d'abord :" & _
        vbCrLf & _
        "- choisir un motif ;" & _
        vbCrLf & _
        "- choisir un joueur si necessaire ;" & _
        vbCrLf & _
        "- cliquer sur Arret du jeu ;" & _
        vbCrLf & _
        "- ou cliquer sur une reprise du jeu.", _
        vbExclamation, _
        "Penalite en attente"

End Sub

Public Function EstActionAutoriseePendantArret( _
    ByVal Target As Range _
) As Boolean

    If EstBoutonRepriseJeu(Target) Then
        EstActionAutoriseePendantArret = True
        Exit Function
    End If

    If Not Intersect( _
        Target, _
        Range("BTN_PTS_TRANSFO") _
    ) Is Nothing Then

        EstActionAutoriseePendantArret = True
        Exit Function

    End If
    
    If Not Intersect( _
        Target, _
        Range("BTN_PTS_ESSAI_PEN") _
    ) Is Nothing Then

        EstActionAutoriseePendantArret = True
        Exit Function

    End If


    If Not Intersect( _
        Target, _
        Range("BTN_PEN_CONTRE_ADV") _
    ) Is Nothing Then

        EstActionAutoriseePendantArret = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_PEN_CONTRE_NOUS") _
    ) Is Nothing Then

        EstActionAutoriseePendantArret = True

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_PTS_PENALITE") _
    ) Is Nothing Then
    
        EstActionAutoriseePendantArret = True
        Exit Function
    
    End If
    
    If IsCartonAction(Target) Then

        EstActionAutoriseePendantArret = True
        Exit Function

    End If
    
End Function

Public Function IsCartonAction( _
    ByVal Target As Range _
) As Boolean

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_ROUGE") _
    ) Is Nothing Then

        IsCartonAction = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_JAUNE") _
    ) Is Nothing Then

        IsCartonAction = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_BLANC") _
    ) Is Nothing Then

        IsCartonAction = True
        Exit Function

    End If

    If Not Intersect( _
        Target, _
        Range("BTN_CARTON_BLEU") _
    ) Is Nothing Then

        IsCartonAction = True

    End If

End Function


