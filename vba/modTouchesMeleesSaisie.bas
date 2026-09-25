Attribute VB_Name = "modTouchesMeleesSaisie"
Option Explicit

' =========================================================
' TOUCHES ET MELEES - SAISIE PAR POPUP
'
' Enchainement, au clic sur une issue de la palette :
'
'   1. la video est mise en pause ;
'   2. la popup s'ouvre, contexte deja rempli ;
'   3. on renseigne ce que la palette ignore ;
'   4. VALIDER ecrit l'action dans le journal et la ligne
'      detaillee dans "Touches Melees", puis rend la main
'      a la saisie video.
'
' ANNULER n'annule pas l'action : il renonce seulement au
' detail. La ligne est ecrite avec son contexte et ses
' autres colonnes vides, pour que le compte des touches
' gagnees et perdues reste juste.
'
' L'action n'est ecrite qu'a la fermeture de la popup :
' c'est elle qui recueille le sauteur, a la place de
' l'attente de joueur d'avant.
'
' Aucun caractere accentue, pas meme dans les chaines :
' le VBE de macOS importe les .bas en Mac Roman.
' =========================================================

Private Const FEUILLE_TABLEAUX As String = "Touches Melees"
Private Const FEUILLE_POPUP_TO As String = "Popup touche"
Private Const FEUILLE_POPUP_ME As String = "Popup melee"

Private Const TAB_TOUCHES As String = "DetailTouches"
Private Const TAB_MELEES As String = "DetailMelees"

' Contexte de la popup ouverte, conserve le temps qu'elle
' reste a l'ecran.
Private PopupAction As String
Private PopupTemps As Double
Private PopupEquipeAction As String
Private PopupPossessionApres As String
Private PopupLancePour As String
Private PopupIssue As String
Private PopupOuverte As Boolean


' ---------------------------------------------------------
' Points d'entree de la palette
'
' La palette ne dit que le bouton clique : l'equipe qui
' porte l'action et celle qui reprend la main se deduisent
' de la possession en cours et de l'issue.
'
' Touche  G  : le ballon revient a Nous, quel que soit le
'              lanceur, puisque G veut dire "pour nous" ;
'         P  : il reste a Adv ;
'         ND : la touche est rejouee par l'autre equipe.
'
' Melee   G  : Nous ;  P : Adv. La palette faisait deja
'              ainsi, sans regarder l'introduction.
' ---------------------------------------------------------

Public Sub OuvrirTouche( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal Issue As String)

    Dim Apres As String

    If CurrentTeam = "" Then

        MsgBox _
            "Selectionne d'abord l'equipe qui effectue " & _
            "la touche.", _
            vbExclamation, _
            "Possession manquante"

        Exit Sub

    End If

    If Not RepriseFaite("ACT_REPRISE_TOUCHE") Then

        MessageRepriseAttendue "touche"
        Exit Sub

    End If

    Select Case Issue

        Case "G"
            Apres = "Nous"

        Case "P"
            Apres = "Adv"

        Case Else
            Apres = IIf(CurrentTeam = "Nous", "Adv", "Nous")

    End Select

    FermerActionEnAttente

    OuvrirPopupTouche _
        ActionTexte, _
        TempsVideo, _
        CurrentTeam, _
        Apres, _
        IIf(CurrentTeam = "Nous", "N", "E"), _
        Issue

End Sub


Public Sub OuvrirMelee( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal Issue As String)

    If CurrentTeam = "" Then

        MsgBox _
            "Selectionne d'abord l'equipe qui introduit " & _
            "la melee.", _
            vbExclamation, _
            "Possession manquante"

        Exit Sub

    End If

    If Not RepriseFaite("ACT_REPRISE_MELEE") Then

        MessageRepriseAttendue "m" & ChrW(234) & "l" & ChrW(233) & "e"
        Exit Sub

    End If

    FermerActionEnAttente

    OuvrirPopupMelee _
        ActionTexte, _
        TempsVideo, _
        CurrentTeam, _
        IIf(Issue = "G", "Nous", "Adv"), _
        IIf(CurrentTeam = "Nous", "N", "E"), _
        Issue

End Sub


' ---------------------------------------------------------
' Ouverture
'
' EquipeAction est l'equipe portee par la ligne du journal.
' PossessionApres est celle qui reprend la main une fois
' l'action ecrite : les deux different souvent.
' ---------------------------------------------------------

Public Sub OuvrirPopupTouche( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal EquipeAction As String, _
    ByVal PossessionApres As String, _
    ByVal LancePour As String, _
    ByVal Issue As String)

    MemoriserContexte ActionTexte, TempsVideo, _
        EquipeAction, PossessionApres, LancePour, Issue

    MettreVideoEnPause

    AfficherPopup FEUILLE_POPUP_TO, _
        "POPUP_TO_CONTEXTE", "POPUP_TO_CHAMPS"

    PoserListeSauteur

End Sub


Public Sub OuvrirPopupMelee( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal EquipeAction As String, _
    ByVal PossessionApres As String, _
    ByVal LancePour As String, _
    ByVal Issue As String)

    MemoriserContexte ActionTexte, TempsVideo, _
        EquipeAction, PossessionApres, LancePour, Issue

    MettreVideoEnPause

    AfficherPopup FEUILLE_POPUP_ME, _
        "POPUP_ME_CONTEXTE", "POPUP_ME_CHAMPS"

End Sub


Private Sub MemoriserContexte( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal EquipeAction As String, _
    ByVal PossessionApres As String, _
    ByVal LancePour As String, _
    ByVal Issue As String)

    PopupAction = ActionTexte
    PopupTemps = TempsVideo
    PopupEquipeAction = EquipeAction
    PopupPossessionApres = PossessionApres
    PopupLancePour = LancePour
    PopupIssue = Issue
    PopupOuverte = True

End Sub


Private Sub AfficherPopup( _
    ByVal NomFeuille As String, _
    ByVal NomContexte As String, _
    ByVal NomChamps As String)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(NomFeuille)

    ws.Visible = xlSheetVisible
    ws.Activate

    ' Les champs repartent vides a chaque ouverture : une
    ' valeur restee de la touche precedente passerait
    ' inapercue.
    ws.Range(NomChamps).ClearContents

    With ws.Range(NomContexte)
        .Cells(1, 1).Value = FormaterTemps(PopupTemps)
        .Cells(2, 1).Value = CurrentHalf
        .Cells(3, 1).Value = EquipeEnToutesLettres(PopupLancePour)
        .Cells(4, 1).Value = IssueEnToutesLettres(PopupIssue)
    End With

    ws.Range(NomChamps).Cells(1, 1).Select

End Sub


' ---------------------------------------------------------
' Validation et abandon
' ---------------------------------------------------------

Public Sub ValiderPopupTouche()

    If Not PopupOuverte Then Exit Sub

    EcrireTouche _
        LireChamp("POPUP_TO_SAUTEUR"), _
        LireChamp("POPUP_TO_ZONE_SAUT"), _
        LireChamp("POPUP_TO_BALLON"), _
        LireChamp("POPUP_TO_ZONE_TERRAIN"), _
        LireChamp("POPUP_TO_ALIGNEMENT"), _
        LireChamp("POPUP_TO_OBS")

    FermerPopup FEUILLE_POPUP_TO

End Sub


Public Sub AnnulerPopupTouche()

    If Not PopupOuverte Then Exit Sub

    ' Le detail est abandonne, pas l'action : la ligne
    ' garde son contexte pour que le compte reste juste.
    EcrireTouche "", "", "", "", "", ""

    FermerPopup FEUILLE_POPUP_TO

End Sub


Public Sub ValiderPopupMelee()

    If Not PopupOuverte Then Exit Sub

    EcrireMelee _
        LireChamp("POPUP_ME_ZONE_LONGUEUR"), _
        LireChamp("POPUP_ME_ZONE_LARGEUR"), _
        LireChamp("POPUP_ME_UTILISATION"), _
        LireChamp("POPUP_ME_OBS")

    FermerPopup FEUILLE_POPUP_ME

End Sub


Public Sub AnnulerPopupMelee()

    If Not PopupOuverte Then Exit Sub

    EcrireMelee "", "", "", ""

    FermerPopup FEUILLE_POPUP_ME

End Sub


Private Sub FermerPopup(ByVal NomFeuille As String)

    Dim ws As Worksheet

    PopupOuverte = False

    Set ws = ThisWorkbook.Sheets(NomFeuille)

    ' shSaisieVideo est le nom de code VBA de la feuille :
    ' son nom Excel porte un accent, que ce module ne peut
    ' pas ecrire.
    shSaisieVideo.Activate

    ws.Visible = xlSheetHidden

End Sub


' ---------------------------------------------------------
' Ecriture
'
' L'action part dans le journal sous l'equipe qui la porte,
' puis la possession est rendue a celle qui enchaine. Meme
' motif que la palette, ou SetCurrentTeam encadre deja
' AjouterAction.
' ---------------------------------------------------------

Private Sub EcrireTouche( _
    ByVal Sauteur As String, _
    ByVal ZoneSaut As String, _
    ByVal Ballon As String, _
    ByVal ZoneTerrain As String, _
    ByVal Alignement As String, _
    ByVal Obs As String)

    Dim Ligne As ListRow

    SetCurrentTeam PopupEquipeAction

    AjouterAction Sauteur, PopupAction, PopupTemps

    SetCurrentTeam PopupPossessionApres

    Set Ligne = LigneLibre(TAB_TOUCHES)

    EcrireTemps Ligne, "Temps video", PopupTemps
    EcrireCellule Ligne, "Mi-temps", CurrentHalf
    EcrireCellule Ligne, "Lance pour", PopupLancePour
    EcrireCellule Ligne, "Issue", PopupIssue
    EcrireCellule Ligne, "Sauteur", Sauteur
    EcrireCellule Ligne, "Zone saut", ZoneSaut
    EcrireCellule Ligne, "Ballon", BallonAbrege(Ballon)
    EcrireCellule Ligne, "Zone terrain", ZoneTerrain
    EcrireCellule Ligne, "Alignement", Alignement
    EcrireCellule Ligne, "Observations", Obs

End Sub


Private Sub EcrireMelee( _
    ByVal ZoneLongueur As String, _
    ByVal ZoneLargeur As String, _
    ByVal Utilisation As String, _
    ByVal Obs As String)

    Dim Ligne As ListRow

    SetCurrentTeam PopupEquipeAction

    AjouterAction "", PopupAction, PopupTemps

    SetCurrentTeam PopupPossessionApres

    Set Ligne = LigneLibre(TAB_MELEES)

    EcrireTemps Ligne, "Temps video", PopupTemps
    EcrireCellule Ligne, "Mi-temps", CurrentHalf
    EcrireCellule Ligne, "Introduction pour", PopupLancePour
    EcrireCellule Ligne, "Issue", PopupIssue
    EcrireCellule Ligne, "Zone longueur", ZoneLongueur
    EcrireCellule Ligne, "Zone largeur", ZoneLargeur
    EcrireCellule Ligne, "Observations", Obs

    ' Une seule liste a la saisie, trois colonnes au
    ' tableau : la puce est posee dans celle qui
    ' correspond.
    Select Case Utilisation

        Case "Avants"
            EcrireCellule Ligne, "Jeu avant", CocheTM

        Case "3/4"
            EcrireCellule Ligne, "Jeu 3/4", CocheTM

        Case "Pied"
            EcrireCellule Ligne, "Jeu pied", CocheTM

    End Select

End Sub


' ---------------------------------------------------------
' Acces aux tableaux
' ---------------------------------------------------------

' Premiere ligne vide du tableau, ou une ligne neuve.
Private Function LigneLibre( _
    ByVal NomTableau As String) As ListRow

    Dim lo As ListObject
    Dim Ligne As ListRow

    Set lo = ThisWorkbook.Sheets(FEUILLE_TABLEAUX) _
        .ListObjects(NomTableau)

    For Each Ligne In lo.ListRows

        If Application.WorksheetFunction.CountA( _
            Ligne.Range) = 0 Then

            Set LigneLibre = Ligne
            Exit Function

        End If

    Next Ligne

    Set LigneLibre = lo.ListRows.Add

End Function


' Le temps s'ecrit comme le journal le fait : une fraction
' de jour, affichee en mm:ss. Une chaine "08:02" serait
' lue par Excel comme huit heures deux, et les deux
' tableaux ne seraient plus comparables.
Private Sub EcrireTemps( _
    ByVal Ligne As ListRow, _
    ByVal NomColonne As String, _
    ByVal Secondes As Double)

    Dim lo As ListObject
    Dim Cellule As Range

    Set lo = Ligne.Parent

    Set Cellule = Ligne.Range.Cells( _
        1, lo.ListColumns(NomColonne).Index)

    Cellule.Value = Secondes / 86400
    Cellule.NumberFormat = "mm:ss"

End Sub


Private Sub EcrireCellule( _
    ByVal Ligne As ListRow, _
    ByVal NomColonne As String, _
    ByVal Valeur As String)

    Dim lo As ListObject

    If Valeur = "" Then Exit Sub

    Set lo = Ligne.Parent

    Ligne.Range.Cells( _
        1, lo.ListColumns(NomColonne).Index).Value = Valeur

End Sub


' ---------------------------------------------------------
' Liste du sauteur
'
' Reprise de la colonne que modPlayers tient a jour pour
' le journal : une seule source de noms dans le classeur.
' ---------------------------------------------------------

Private Sub PoserListeSauteur()

    Dim wsCompo As Worksheet
    Dim Derniere As Long

    ' La colonne AA est reconstruite depuis la compo :
    ' sans cet appel, elle reste vide tant que le journal
    ' n'a pas ete rafraichi une premiere fois.
    '
    ' L'erreur est ignoree a dessein : la procedure finit
    ' en posant une validation sur le journal, ce qui
    ' echoue tant qu'il est vide. AA est deja rempli a ce
    ' stade, et c'est tout ce qui nous interesse ici.
    On Error Resume Next
    ActualiserJoueursJournal
    On Error GoTo 0

    Set wsCompo = ThisWorkbook.Sheets("Compo")

    Derniere = wsCompo.Cells( _
        wsCompo.Rows.Count, "AA").End(xlUp).Row

    If Derniere < 1 Then Exit Sub

    With ThisWorkbook.Sheets(FEUILLE_POPUP_TO) _
        .Range("POPUP_TO_SAUTEUR").Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="='" & wsCompo.Name & "'!$AA$1:$AA$" _
                & Derniere

    End With

End Sub


' ---------------------------------------------------------
' Video
'
' La commande du moteur est une bascule : elle n'est
' envoyee que si la lecture est en cours, sans quoi la
' popup relancerait la video au lieu de l'arreter.
' ---------------------------------------------------------

Private Sub MettreVideoEnPause()

    On Error Resume Next

    If VideoEnLecture Then PlayPauseChronoVideo

    On Error GoTo 0

End Sub


' Le moteur ne dit pas s'il lit : on regarde si le temps
' avance.
'
' La premiere mesure est celle que la palette vient de
' prendre pour horodater l'action : la reutiliser epargne
' une requete au moteur, qui n'en recevra que deux au lieu
' de trois a chaque popup.
Private Function VideoEnLecture() As Boolean

    Dim Apres As Double
    Dim Fin As Single

    If PopupTemps <= 0 Then Exit Function

    Fin = Timer + 0.3

    Do While Timer < Fin
        DoEvents
    Loop

    Apres = GetTimeVideo()

    If Apres < 0 Then Exit Function

    VideoEnLecture = (Abs(Apres - PopupTemps) > 0.05)

End Function


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function LireChamp( _
    ByVal NomPlage As String) As String

    LireChamp = Trim(CStr( _
        ThisWorkbook.Names(NomPlage).RefersToRange.Value))

End Function


' ---------------------------------------------------------
' Reprise du jeu
'
' Une issue ne se saisit qu'apres la reprise qui lui
' correspond : une touche gagnee sans touche jouee ne veut
' rien dire, et la ligne du tableau serait orpheline.
'
' L'action retenue est celle du temps le plus avance de la
' mi-temps en cours, comme le fait JeuActuellementArrete :
' la saisie peut etre reprise dans le desordre.
' ---------------------------------------------------------

' Libelle d'une action, lu par son nom de plage.
'
' Les plages ACT_* sont posees sur la feuille Parametres,
' pas sur celle de saisie : les qualifier d'une feuille
' precise leve l'erreur 1004. Le nom etant defini au
' niveau du classeur, ThisWorkbook.Names le resout seul.
Private Function LibelleAction( _
    ByVal NomPlage As String) As Variant

    LibelleAction = ThisWorkbook.Names(NomPlage) _
        .RefersToRange.Value

End Function


Private Function RepriseFaite( _
    ByVal NomPlageAction As String) As Boolean

    Dim Attendue As String

    Attendue = CStr(LibelleAction(NomPlageAction))

    RepriseFaite = (StrComp( _
        DerniereRepriseOuArret, _
        Attendue, _
        vbTextCompare) = 0)

End Function


' Vrai tant qu'une reprise n'a pas ouvert le jeu : c'est
' le cas apres un arret, mais aussi en debut de mi-temps,
' ou rien n'a encore ete saisi.
Public Function JeuNonRepris() As Boolean

    Dim Derniere As String

    Derniere = DerniereRepriseOuArret

    If Derniere = "" Then
        JeuNonRepris = True
        Exit Function
    End If

    JeuNonRepris = (StrComp( _
        Derniere, _
        CStr(LibelleAction("ACT_ARRET_DU_JEU")), _
        vbTextCompare) = 0)

End Function


' Derniere action de reprise ou d'arret de la mi-temps, au
' temps le plus avance. Les actions de jeu sont ignorees :
' seules celles-ci decrivent l'etat du jeu.
Private Function DerniereRepriseOuArret() As String

    Dim lo As ListObject
    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long
    Dim i As Long
    Dim Temps As Double
    Dim Meilleur As Double
    Dim Action As String

    Set lo = ThisWorkbook _
        .Worksheets("Journal actions") _
        .ListObjects("JournalActions")

    If lo.DataBodyRange Is Nothing Then Exit Function

    ColTemps = lo.ListColumns("temps video").Index
    ColMiTemps = lo.ListColumns("Mi-temps").Index
    ColAction = lo.ListColumns("Action").Index

    Meilleur = -1

    For i = 1 To lo.ListRows.Count

        If CStr(lo.ListRows(i).Range.Cells(1, ColMiTemps) _
            .Value) = CurrentHalf Then

            If IsNumeric(lo.ListRows(i).Range _
                .Cells(1, ColTemps).Value) Then

                Action = Trim(CStr(lo.ListRows(i).Range _
                    .Cells(1, ColAction).Value))

                If EstTexteRepriseOuArret(Action) Then

                    Temps = CDbl(lo.ListRows(i).Range _
                        .Cells(1, ColTemps).Value)

                    ' Strictement superieur : le journal est
                    ' trie du plus recent au plus ancien, donc
                    ' a temps egal la premiere ligne lue est la
                    ' plus recente. Meme choix que
                    ' JeuActuellementArrete.
                    If Temps > Meilleur Then

                        Meilleur = Temps
                        DerniereRepriseOuArret = Action

                    End If

                End If

            End If

        End If

    Next i

End Function


Private Sub MessageRepriseAttendue( _
    ByVal Phase As String)

    MsgBox _
        "Le jeu n'a pas encore " & ChrW(233) & "t" & _
        ChrW(233) & " repris sur " & Phase & "." & _
        vbCrLf & vbCrLf & _
        "Clique d'abord le bouton de reprise " & _
        "correspondant.", _
        vbExclamation, _
        "Reprise du jeu attendue"

End Sub


' ---------------------------------------------------------
' Abreviations
'
' Les popups parlent en toutes lettres, les tableaux
' gardent leurs lettres pour rester denses et comptables.
' ---------------------------------------------------------

Private Function EquipeEnToutesLettres( _
    ByVal Abrege As String) As String

    If Abrege = "N" Then
        EquipeEnToutesLettres = "Nous"
    Else
        EquipeEnToutesLettres = "Adversaire"
    End If

End Function


Private Function IssueEnToutesLettres( _
    ByVal Abrege As String) As String

    Select Case Abrege

        Case "G"
            IssueEnToutesLettres = "Gagn" & ChrW(233) & "e"

        Case "P"
            IssueEnToutesLettres = "Perdue"

        Case Else
            IssueEnToutesLettres = "Pas droite"

    End Select

End Function


Private Function BallonAbrege( _
    ByVal Libelle As String) As String

    Select Case Libelle

        Case "Chaud"
            BallonAbrege = "C"

        Case "Froid"
            BallonAbrege = "F"

    End Select

End Function


Private Function FormaterTemps( _
    ByVal Secondes As Double) As String

    Dim Total As Long

    Total = CLng(Int(Secondes))

    FormaterTemps = Format(Total \ 60, "00") & ":" & _
        Format(Total Mod 60, "00")

End Function


' Caractere de la puce, partage avec le tableau.
Public Function CocheTM() As String

    CocheTM = ChrW(10003)

End Function
