Attribute VB_Name = "modUI"
Option Explicit

Public Sub ApplyFormat( _
    ByVal TargetCell As Range, _
    ByVal ModelCell As Range)

    ' =====================================================
    ' Le format est recopie propriete par propriete, et
    ' seulement la ou il differe du modele.
    '
    ' La premiere version defusionnait le bouton, collait
    ' le modele par le presse-papiers sur chaque cellule,
    ' refusionnait, puis remettait le contenu en place.
    ' Chaque UnMerge / Merge oblige Excel a refaire la mise
    ' en page de toute la feuille : repeindre les quarante
    ' boutons coutait 1,1 seconde a chaque clic.
    '
    ' Comparer avant d'ecrire fait le reste du travail :
    ' lors d'une remise a zero, trente-neuf boutons sur
    ' quarante portent deja leur style passif, et les douze
    ' styles partagent les memes bordures, le meme
    ' alignement et le meme format de nombre. Il ne reste
    ' alors presque rien a ecrire.
    '
    ' La comparaison se fait sur la cellule d'ancrage : sur
    ' une zone fusionnee, c'est elle qui porte le format
    ' qu'Excel affiche.
    ' =====================================================

    Dim EtatEvenements As Boolean
    Dim ZoneCible As Range
    Dim Ancre As Range
    Dim Cotes As Variant
    Dim i As Long

    EtatEvenements = Application.EnableEvents

    On Error GoTo Fin

    Application.EnableEvents = False

    If TargetCell.MergeCells Then
        Set ZoneCible = TargetCell.MergeArea
    Else
        Set ZoneCible = TargetCell
    End If

    Set Ancre = ZoneCible.Cells(1, 1)

    ' -----------------------------------------------------
    ' Fond
    ' -----------------------------------------------------

    If Ancre.Interior.Pattern <> _
        ModelCell.Interior.Pattern Then

        ZoneCible.Interior.Pattern = _
            ModelCell.Interior.Pattern

    End If

    If Ancre.Interior.Color <> _
        ModelCell.Interior.Color Then

        ZoneCible.Interior.Color = _
            ModelCell.Interior.Color

    End If

    ' -----------------------------------------------------
    ' Police
    ' -----------------------------------------------------

    If Ancre.Font.Name <> ModelCell.Font.Name Then
        ZoneCible.Font.Name = ModelCell.Font.Name
    End If

    If Ancre.Font.Size <> ModelCell.Font.Size Then
        ZoneCible.Font.Size = ModelCell.Font.Size
    End If

    If Ancre.Font.Bold <> ModelCell.Font.Bold Then
        ZoneCible.Font.Bold = ModelCell.Font.Bold
    End If

    If Ancre.Font.Italic <> ModelCell.Font.Italic Then
        ZoneCible.Font.Italic = ModelCell.Font.Italic
    End If

    If Ancre.Font.Color <> ModelCell.Font.Color Then
        ZoneCible.Font.Color = ModelCell.Font.Color
    End If

    ' -----------------------------------------------------
    ' Alignement et format
    ' -----------------------------------------------------

    If Ancre.HorizontalAlignment <> _
        ModelCell.HorizontalAlignment Then

        ZoneCible.HorizontalAlignment = _
            ModelCell.HorizontalAlignment

    End If

    If Ancre.VerticalAlignment <> _
        ModelCell.VerticalAlignment Then

        ZoneCible.VerticalAlignment = _
            ModelCell.VerticalAlignment

    End If

    If Ancre.WrapText <> ModelCell.WrapText Then
        ZoneCible.WrapText = ModelCell.WrapText
    End If

    If Ancre.NumberFormat <> ModelCell.NumberFormat Then
        ZoneCible.NumberFormat = ModelCell.NumberFormat
    End If

    ' -----------------------------------------------------
    ' Bordures
    ' -----------------------------------------------------

    Cotes = Array( _
        xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)

    For i = LBound(Cotes) To UBound(Cotes)

        If Ancre.Borders(Cotes(i)).LineStyle <> _
            ModelCell.Borders(Cotes(i)).LineStyle Then

            ZoneCible.Borders(Cotes(i)).LineStyle = _
                ModelCell.Borders(Cotes(i)).LineStyle

        End If

        If ModelCell.Borders(Cotes(i)).LineStyle <> _
            xlLineStyleNone Then

            If Ancre.Borders(Cotes(i)).Color <> _
                ModelCell.Borders(Cotes(i)).Color Then

                ZoneCible.Borders(Cotes(i)).Color = _
                    ModelCell.Borders(Cotes(i)).Color

            End If

            If Ancre.Borders(Cotes(i)).Weight <> _
                ModelCell.Borders(Cotes(i)).Weight Then

                ZoneCible.Borders(Cotes(i)).Weight = _
                    ModelCell.Borders(Cotes(i)).Weight

            End If

        End If

    Next i

Fin:

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

    ApplyFormat ws.Range("BTN_CF_CONTRE_ADV"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_CF_CONTRE_NOUS"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_PEN_CONTRE_NOUS"), _
        shParametres.Range("STYLE_BTN_PEN")

    ApplyFormat ws.Range("BTN_FRANCHISSEMENT"), _
        shParametres.Range("STYLE_BTN_PTS")
        
    ApplyFormat ws.Range("BTN_AVANCEE"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_ATT_NULLE"), _
        shParametres.Range("STYLE_BTN_PTS")

    ApplyFormat ws.Range("BTN_CARTON_ROUGE"), _
        shParametres.Range("STYLE_BTN_CHRONO_RESET")

    ApplyFormat ws.Range("BTN_CARTON_JAUNE"), _
        shParametres.Range("STYLE_CARTON_JAUNE")

    ApplyFormat ws.Range("BTN_CARTON_BLANC"), _
        shParametres.Range("STYLE_CARTON_BLANC")

    ApplyFormat ws.Range("BTN_CARTON_BLEU"), _
        shParametres.Range("STYLE_CARTON_BLEU")

    ApplyFormat ws.Range("BTN_RETOUR_CARTON"), _
        shParametres.Range("STYLE_BTN_REPRISE")

SortiePropre:

    Application.CutCopyMode = False
    Application.EnableEvents = EtatEvenements
    Application.ScreenUpdating = EtatAffichage

End Sub

Public Sub ReleaseButton()

    Dim EtatEvenements As Boolean
    Dim EtatAffichage As Boolean

    EtatEvenements = Application.EnableEvents
    EtatAffichage = Application.ScreenUpdating

    On Error GoTo SortiePropre

    Application.EnableEvents = False

    ' L'affichage est fige le temps du deplacement : sans
    ' cela Excel montre la selection passer par la case de
    ' parking avant que le bouton ne reprenne sa couleur.
    Application.ScreenUpdating = False

    If ActiveWorkbook Is ThisWorkbook Then
        If ActiveSheet Is shSaisieVideo Then
            shSaisieVideo.Range("BTN_PARKING").Select
        End If
    End If

SortiePropre:

    Application.ScreenUpdating = EtatAffichage
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

    PendingMotifPenalite = ""
    PendingGroupeFautifRequired = False

SortiePropre:

    Application.CutCopyMode = False
    Application.EnableEvents = EtatEvenements
    Application.ScreenUpdating = EtatAffichage

End Sub

