Attribute VB_Name = "modUI"
Option Explicit

Public Sub ApplyFormat( _
    ByVal TargetCell As Range, _
    ByVal ModelCell As Range)

    Dim EtatEvenements As Boolean
    Dim ZoneCible As Range
    Dim Cellule As Range

    Dim ValeurCible As Variant
    Dim FormuleCible As Variant
    Dim ContientFormule As Boolean

    EtatEvenements = Application.EnableEvents

    On Error GoTo Fin

    Application.EnableEvents = False

    ' =====================================================
    ' CAS D'UN BOUTON FUSIONNE
    ' =====================================================

    If TargetCell.MergeCells Then

        Set ZoneCible = TargetCell.MergeArea

        ContientFormule = _
            ZoneCible.Cells(1, 1).HasFormula

        If ContientFormule Then

            FormuleCible = _
                ZoneCible.Cells(1, 1).Formula

        Else

            ValeurCible = _
                ZoneCible.Cells(1, 1).Value

        End If

        ' On d_fusionne temporairement.
        ZoneCible.UnMerge

        ' On copie le modle une seule fois.
        ModelCell.Copy

        ' Puis on applique r_ellement le format
        ' ˆ CHAQUE cellule de l'ancienne fusion.
        For Each Cellule In ZoneCible.Cells

            Cellule.PasteSpecial _
                Paste:=xlPasteFormats

        Next Cellule

        ' On recr_e la fusion.
        ZoneCible.Merge

        ' On restaure le contenu.
        If ContientFormule Then

            ZoneCible.Cells(1, 1).Formula = _
                FormuleCible

        Else

            ZoneCible.Cells(1, 1).Value = _
                ValeurCible

        End If

    ' =====================================================
    ' CAS NORMAL
    ' =====================================================

    Else

        ModelCell.Copy

        TargetCell.PasteSpecial _
            Paste:=xlPasteFormats

    End If

Fin:

    Application.CutCopyMode = False
    Application.EnableEvents = EtatEvenements

End Sub

Public Sub SetActiveButton( _
    ByVal Btn As Range)

    ApplyFormat Btn, shParametres.Range("STYLE_BTN_ACTIF")

End Sub

Public Sub ResetTeamButtons()

    Dim ws As Worksheet

    Set ws = shSaisieVideo

    ApplyFormat ws.Range("BTN_NOUS"), _
        shParametres.Range("STYLE_BTN_POSSESSION")

    ApplyFormat ws.Range("BTN_ADV"), _
        shParametres.Range("STYLE_BTN_POSSESSION")

End Sub

Public Sub SetCurrentTeam(ByVal Equipe As String)

    Dim ws As Worksheet

    Set ws = shSaisieVideo

    ResetTeamButtons

    Select Case Equipe

        Case "Nous"
            CurrentTeam = "Nous"
            SetActiveButton ws.Range("BTN_NOUS")

        Case "Adv"
            CurrentTeam = "Adv"
            SetActiveButton ws.Range("BTN_ADV")

        Case Else
            CurrentTeam = ""

    End Select

End Sub

Public Sub ToggleCurrentTeam()

    Select Case CurrentTeam

        Case "Nous"
            SetCurrentTeam "Adv"

        Case "Adv"
            SetCurrentTeam "Nous"

    End Select

End Sub

Public Sub ResetHalfButtons()

    Dim ws As Worksheet

    Set ws = shSaisieVideo

    ApplyFormat ws.Range("BTN_MT1"), _
        shParametres.Range("STYLE_BTN_MT")

    ApplyFormat ws.Range("BTN_MT2"), _
        shParametres.Range("STYLE_BTN_MT")

End Sub

Public Sub ResetActionButtons()

    Dim ws As Worksheet
    Dim EtatAffichage As Boolean
    Dim EtatEvenements As Boolean

    On Error GoTo SortiePropre

    EtatAffichage = Application.ScreenUpdating
    EtatEvenements = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = shSaisieVideo

    ApplyFormat ws.Range("BTN_PL_OFF"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_PL_NORMAL"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_PL_RATE"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_PL_A2"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_PL_HAUT"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_TO_G"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")
    
    ApplyFormat ws.Range("BTN_TO_P"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")
    
    ApplyFormat ws.Range("BTN_TO_PAS_DROIT"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_ME_G"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")
    
    ApplyFormat ws.Range("BTN_ME_P"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_PTS_ESSAI"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_PTS_ESSAI_PEN"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_PTS_TRANSFO"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_PTS_PENALITE"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_PTS_DROP"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_REPRISE_TOUCHE"), _
        shParametres.Range("STYLE_BTN_REPRISE")

    ApplyFormat ws.Range("BTN_REPRISE_MELEE"), _
        shParametres.Range("STYLE_BTN_REPRISE")

    ApplyFormat ws.Range("BTN_REPRISE_RENVOI"), _
        shParametres.Range("STYLE_BTN_REPRISE")

    ApplyFormat ws.Range("BTN_REPRISE_A_LA_MAIN"), _
        shParametres.Range("STYLE_BTN_REPRISE")

    ApplyFormat ws.Range("BTN_ARRET_DU_JEU"), _
        shParametres.Range("STYLE_BTN_ARRET_JEU")

    ApplyFormat ws.Range("BTN_EN_AVANT"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_JEU_AU_PIED"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_RECEPTION_NOUS"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_RECEPTION_ADV"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_GRATTAGE"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_CONTRE_RUCK"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")
        
    ApplyFormat ws.Range("BTN_ARRACHAGE"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_TURNOVER"), _
        shParametres.Range("STYLE_BTN_PL_TOUCHES")

    ApplyFormat ws.Range("BTN_PEN_CONTRE_ADV"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_CONTRE_NOUS"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_MAUL"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_RUCK"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_HORS_JEU"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_PL_A_2"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_PL_HAUT"), _
        shParametres.Range("STYLE_BTN_PEN")
        
    ApplyFormat ws.Range("BTN_FRANCHISSEMENT"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_CARTON_ROUGE"), _
        shParametres.Range("STYLE_BTN_CHRONO_RESET")

    ApplyFormat ws.Range("BTN_CARTON_JAUNE"), _
        shParametres.Range("STYLE_CARTON_JAUNE")

    ApplyFormat ws.Range("BTN_CARTON_BLANC"), _
        shParametres.Range("STYLE_CARTON_BLANC")

    ApplyFormat ws.Range("BTN_CARTON_BLEU"), _
        shParametres.Range("STYLE_CARTON_BLEU")

SortiePropre:

    Application.CutCopyMode = False
    Application.EnableEvents = EtatEvenements
    Application.ScreenUpdating = EtatAffichage

End Sub

Public Sub ReleaseButton()

    Dim EtatEvenements As Boolean

    EtatEvenements = Application.EnableEvents

    On Error GoTo SortiePropre

    Application.EnableEvents = False

    If ActiveWorkbook Is ThisWorkbook Then
        If ActiveSheet Is shSaisieVideo Then
            shSaisieVideo.Range("BTN_PARKING").Select
        End If
    End If

SortiePropre:

    Application.EnableEvents = EtatEvenements

End Sub

Public Sub ReinitialiserSelectionSaisieVideo()

    Dim EtatAffichage As Boolean
    Dim EtatEvenements As Boolean

    On Error GoTo SortiePropre

    EtatAffichage = Application.ScreenUpdating
    EtatEvenements = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ResetTeamButtons
    ResetHalfButtons
    ResetActionButtons

    CurrentTeam = ""
    CurrentHalf = ""

    PendingAction = ""
    WaitingForPlayer = False

    PendingActionTime = 0
    PendingActionTimeAvailable = False

    PendingPlayerCount = 0
    PendingAllowMultiplePlayers = False
    PendingForcedTeam = ""
    PendingActionChangesPossession = False

    PendingPenaltyActive = False
    PendingPenaltyBaseAction = ""
    PendingPenaltyTime = 0
    PendingPenaltyForcedTeam = ""
    PendingPenaltyNeedsPlayer = False

    PendingMotifPenalite = ""
    PendingGroupeFautifRequired = False

SortiePropre:

    Application.CutCopyMode = False
    Application.EnableEvents = EtatEvenements
    Application.ScreenUpdating = EtatAffichage

End Sub

