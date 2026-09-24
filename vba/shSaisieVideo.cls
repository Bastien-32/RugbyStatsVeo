Attribute VB_Name = "shSaisieVideo"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private TraitementSelectionEnCours As Boolean

Private Sub Worksheet_Activate()

    ActiverRaccourcisVideo
    DemarrerAffichageChronoVideo
    
    MettreAJourBoutonSwitchMode

End Sub

Private Sub Worksheet_Deactivate()

    ' Pas de reinitialisation ici : la possession et la
    ' mi-temps doivent survivre au passage sur la popup.
    
    DesactiverRaccourcisVideo
    ArreterAffichageChronoVideo

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If TraitementSelectionEnCours Then
        Exit Sub
    End If

    TraitementSelectionEnCours = True

    On Error GoTo SortiePropre

    TraiterSelectionChange Target

SortiePropre:

    TraitementSelectionEnCours = False

End Sub

Private Sub TraiterSelectionChange(ByVal Target As Range)

    Dim ActionTexte As String
    Dim joueur As String
    Dim NumeroJoueur As String
    Dim TempsVideo As Double
    Dim MotifPenalite As String
    Dim GroupeFautif As String
    Dim TempsPenaliteFinalisee As Double
    Dim ActionJoueurFinalisee As String
    Dim AjouterArretApresJoueur As Boolean
    Dim ActionFermee As Boolean

    ' =====================================================
    ' BOUTON FRANCHISSEMENT FUSIONNE
    '
    ' Excel renvoie toute la zone fusionnee lors du clic.
    ' On ramene donc Target a la cellule portant le nom
    ' afin que le traitement normal des boutons fonctionne.
    ' =====================================================
    
    If Target.MergeCells Then
    
        If Not Intersect( _
            Target, _
            Range("BTN_FRANCHISSEMENT").MergeArea _
        ) Is Nothing Then
    
            Set Target = _
                Range("BTN_FRANCHISSEMENT").Cells(1, 1)
    
        End If
    
    End If

    If Target.CountLarge > 1 Then

        If Target.Address(False, False) = _
            Range("CELL_CHRONO_VIDEO").Resize(1, 3).Address(False, False) Then

            FinaliserActionEnAttente
            ReleaseButton
            Exit Sub

        End If

        Exit Sub

    End If
    
    ' =====================================================
    ' PENALITE EN ATTENTE D'UN MOTIF
    ' =====================================================
    
    If PendingPenaltyActive Then
    
        If IsPenaltyReasonButton(Target) Then
    
            MotifPenalite = GetPenaltyReason(Target)
    
            ActiverBoutonPenaliteEnAttente
    
            If PendingPenaltyNeedsPlayer Then
    
                TransformerPenaliteEnAttenteJoueur _
                    MotifPenalite
    
                Exit Sub
    
            Else
    
                ActionTexte = PendingPenaltyBaseAction
                TempsPenaliteFinalisee = PendingPenaltyTime
    
                SetCurrentTeam PendingPenaltyForcedTeam
    
                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsPenaliteFinalisee, _
                    MotifPenalite, _
                    ""
    
                FermerPenaliteEnAttente
    
                AjouterArretJeuAutomatique _
                    TempsPenaliteFinalisee
    
                ReleaseButton
                Exit Sub
    
            End If
    
            ElseIf PendingPenaltyNeedsPlayer And _
                EstBoutonJoueur(Target) Then
    
            ' Joueur choisi sans motif :
            ' le motif devient "?".
            TransformerPenaliteEnAttenteJoueur "?"
    
            ' Ne pas quitter la procedure :
            ' le clic joueur sera traite plus bas.
    
        ElseIf EstBoutonRepriseJeu(Target) Then
    
            ' La reprise finalise la penalite avec les
            ' informations manquantes remplacees par "?".
            FinaliserPenaliteOubliee
    
            ' Ne pas quitter :
            ' la reprise cliquee doit ensuite etre enregistree.
    
        ElseIf Not Intersect( _
            Target, _
            Range("BTN_ARRET_DU_JEU") _
        ) Is Nothing Then
    
            FinaliserPenaliteOubliee
    
            ReleaseButton
            Exit Sub
    
        Else
    
            AfficherAlertePenaliteEnAttente
    
            ReleaseButton
            Exit Sub
    
        End If
    
    End If
    
    
    ' =====================================================
    ' PENALITE EN ATTENTE D'UN JOUEUR
    ' =====================================================
    
    If EstPenaliteEnAttenteJoueur Then
    
        If EstBoutonJoueur(Target) Then
    
            ' Le clic joueur sera traite plus bas.
    
        ElseIf EstBoutonRepriseJeu(Target) Then
    
            ' Le joueur manquant devient "?".
            FinaliserActionEnAttente
    
            ' Ne pas quitter :
            ' la reprise doit ensuite etre enregistree.
    
        ElseIf Not Intersect( _
            Target, _
            Range("BTN_ARRET_DU_JEU") _
        ) Is Nothing Then
    
            FinaliserActionEnAttente
    
            ReleaseButton
            Exit Sub
    
        ElseIf Not Intersect( _
            Target, _
            Range("PALETTE_JOUEURS") _
        ) Is Nothing Then
    
            ' Cellule vide de la palette :
            ' la penalite reste en attente.
            Exit Sub
    
        Else
    
            AfficherAlertePenaliteEnAttente
    
            ReleaseButton
            Exit Sub
    
        End If
    
    End If
    
    
    ' =====================================================
    ' AUTRE ACTION EN ATTENTE D'UN JOUEUR
    ' =====================================================
    
    If WaitingForPlayer Then
    
        If Not EstPenaliteEnAttenteJoueur Then
    
            If EstBoutonJoueur(Target) Then
    
                ' Le clic joueur sera traite plus bas.
    
            ElseIf Not Intersect( _
                Target, _
                Range("PALETTE_JOUEURS") _
            ) Is Nothing Then
    
                Exit Sub
    
            Else
    
                FinaliserActionEnAttente
    
            End If
    
        End If
    
    End If

    If Not Intersect(Target, Range("CELL_CHRONO_VIDEO").Resize(1, 3)) Is Nothing Then
        ReleaseButton
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_NOUS")) Is Nothing Then
        SetCurrentTeam "Nous"
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_ADV")) Is Nothing Then
        SetCurrentTeam "Adv"
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_MT1")) Is Nothing Then
        ResetHalfButtons
        SetActiveButton Range("BTN_MT1")
        CurrentHalf = "MT1"
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_MT2")) Is Nothing Then
        ResetHalfButtons
        SetActiveButton Range("BTN_MT2")
        CurrentHalf = "MT2"
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_CHRONO_PLAY_PAUSE")) Is Nothing Then
        PlayPauseChronoVideo
        ReleaseButton
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_CHRONO_MINUS_5_SEC")) Is Nothing Then
        Recul5SecondesVideo
        ReleaseButton
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_CHRONO_PLUS_5_SEC")) Is Nothing Then
        Avance5SecondesVideo
        ReleaseButton
        Exit Sub
    End If

    If Not Intersect(Target, Range("BTN_CHRONO_RESET")) Is Nothing Then
        ResetChronoVideo
        ReleaseButton
        Exit Sub
    End If
    
    ' =====================================================
    ' POSSESSION OBLIGATOIRE POUR
    ' FRANCHISSEMENT ET CARTONS
    ' =====================================================

    If IsCartonAction(Target) Or _
        EstAttitudeAuContact(Target) Or _
        Not Intersect( _
            Target, _
            Range("BTN_TURNOVER") _
        ) Is Nothing Then

        If CurrentTeam = "" Then

            MsgBox _
                "Selectionne d'abord l'equipe qui possede le ballon.", _
                vbExclamation, _
                "Possession manquante"

            ReleaseButton
            Exit Sub

        End If

    End If

    ActionTexte = GetAction(Target)
    
    If ActionTexte <> "" Then
        If CurrentHalf = "" Then

            MsgBox _
                "S" & ChrW(233) & _
                "lectionne d'abord la mi-temps.", _
                vbExclamation, _
                "Mi-temps manquante"

            ReleaseButton
            Exit Sub

        End If
    
        If JeuActuellementArrete Then
    
            If Not EstActionAutoriseePendantArret(Target) Then
    
                If Intersect( _
                    Target, _
                    Range("BTN_ARRET_DU_JEU") _
                ) Is Nothing Then
    
                    MsgBox _
                        "Le jeu est actuellement arr" & _
                        ChrW(234) & _
                        "t" & ChrW(233) & "." & _
                        vbCrLf & vbCrLf & _
                        "La prochaine action doit " & _
                        ChrW(234) & _
                        "tre une reprise du jeu :" & _
                        vbCrLf & _
                        "- touche" & vbCrLf & _
                        "- m" & ChrW(234) & "l" & ChrW(233) & "e" & vbCrLf & _
                        "- renvoi" & vbCrLf & _
                        "- reprise " & ChrW(224) & " la main.", _
                        vbExclamation, _
                        "Reprise du jeu attendue"
    
                    ReleaseButton
                    Exit Sub
    
                End If
    
            End If
    
        End If
    
        TempsVideo = GetTimeVideo()

        If TempsVideo < 0 Then
            MsgBox _
                "Impossible de r" & ChrW(233) & "cup" & ChrW(233) & "rer le temps r" & ChrW(233) & "el de la vid" & ChrW(233) & "o.", _
                vbExclamation, _
                "VeoVideoControl"
            Exit Sub
        End If

        ResetActionButtons
        SetActiveButton Target

        ' =====================================================
        ' ATTITUDE AU CONTACT
        ' =====================================================

        If EstAttitudeAuContact(Target) Then

            If CurrentTeam = "Nous" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    False, _
                    "", _
                    True

            Else

                FermerActionEnAttente

                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo

                ReleaseButton

            End If

            Exit Sub

        End If

        ' =====================================================
        ' CARTONS
        ' =====================================================

        If IsCartonAction(Target) Then

            If CurrentTeam = "Nous" Then

                FermerActionEnAttente

                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo

                ReleaseButton

            Else

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    False, _
                    "Adv", _
                    False

            End If

            Exit Sub

        End If

        ' =====================================================
        ' EN-AVANT
        ' =====================================================

        If Not Intersect(Target, Range("BTN_EN_AVANT")) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "S" _
                    & ChrW(233) & _
                    "lectionne d'abord l'" & ChrW(233) & "quipe qui poss" _
                    & ChrW(232) & _
                    "de le ballon.", _
                    vbExclamation, _
                    "Possession manquante"
                
                ReleaseButton
                Exit Sub
        
            End If
        
            If CurrentTeam = "Nous" Then
        
                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False
        
            Else
        
                FermerActionEnAttente
        
                SetCurrentTeam "Adv"
        
                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo
        
                SetCurrentTeam "Nous"
        
                ReleaseButton
        
            End If
        
            Exit Sub
        
        End If

        ' =====================================================
        ' JEU AU PIED
        ' =====================================================

        If Not Intersect(Target, Range("BTN_JEU_AU_PIED")) Is Nothing Then

            If CurrentTeam = "" Then
                MsgBox _
                    "S" _
                    & ChrW(233) & _
                    "lectionne d'abord l'" _
                    & ChrW(233) & _
                    "quipe qui poss" _
                    & ChrW(232) & _
                    "de le ballon.", _
                    vbExclamation, _
                    "Possession manquante"

                ReleaseButton
                Exit Sub
            End If

            FermerActionEnAttente

            AjouterAction "", ActionTexte, TempsVideo

            ReleaseButton
            Exit Sub

        End If

        ' =====================================================
        ' RECEPTIONS
        ' =====================================================

        If Not Intersect(Target, Range("BTN_RECEPTION_NOUS")) Is Nothing Then

            FermerActionEnAttente

            SetCurrentTeam "Nous"
            AjouterAction "", ActionTexte, TempsVideo

            ReleaseButton
            Exit Sub

        End If

        If Not Intersect(Target, Range("BTN_RECEPTION_ADV")) Is Nothing Then

            FermerActionEnAttente

            SetCurrentTeam "Adv"
            AjouterAction "", ActionTexte, TempsVideo

            ReleaseButton
            Exit Sub

        End If

        ' =====================================================
        ' TOUCHES
        '
        ' Le resultat d_pend de l'equipe qui effectue la touche,
        ' donc de la possession active.
        ' =====================================================
        
        If Not Intersect( _
            Target, _
            Range("BTN_TO_G") _
        ) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "Selectionne d'abord l'equipe qui effectue la touche.", _
                    vbExclamation, _
                    "Possession manquante"
        
                ReleaseButton
                Exit Sub
        
            End If
        
            If CurrentTeam = "Nous" Then
        
                ' Touche a nous conservee.
                '
                ' Attend le joueur.
                ' Si aucun joueur n'est choisi avant une autre action,
                ' "?" sera ajoute automatiquement.
                ' La possession reste a Nous.
        
                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    False, _
                    "Nous", _
                    False
        
            Else
        
                ' Touche a eux volee.
                '
                ' Attend le joueur qui vole le ballon.
                ' Si aucun joueur n'est choisi avant une autre action,
                ' "?" sera ajoute automatiquement.
                ' La possession passe ensuite a Nous.
        
                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Adv", _
                    False
        
            End If
        
            ReleaseButton
            Exit Sub
        
        End If
        
        
        If Not Intersect( _
            Target, _
            Range("BTN_TO_P") _
        ) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "Selectionne d'abord l'equipe qui effectue la touche.", _
                    vbExclamation, _
                    "Possession manquante"
        
                ReleaseButton
                Exit Sub
        
            End If
        
            If CurrentTeam = "Nous" Then
        
                ' Touche a nous perdue.
                '
                ' Attend le joueur.
                ' Si aucun joueur n'est choisi avant une autre action,
                ' "?" sera ajoute automatiquement.
                ' La possession passe ensuite a Adv.
        
                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False
        
            Else
        
                ' Touche a eux perdue.
                '
                ' Aucun joueur demande.
                ' La ligne est creee immediatement.
                ' La possession reste a adv.
        
                FermerActionEnAttente
        
                SetCurrentTeam "Adv"
        
                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo
        
                ReleaseButton
        
            End If
        
            Exit Sub
        
        End If
        
        
        If Not Intersect( _
            Target, _
            Range("BTN_TO_PAS_DROIT") _
        ) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "Selectionne d'abord l'equipe qui effectue la touche.", _
                    vbExclamation, _
                    "Possession manquante"
        
                ReleaseButton
                Exit Sub
        
            End If
        
            If CurrentTeam = "Nous" Then
        
                ' Touche a nous pas droite.
                '
                ' Attend le joueur.
                ' Si aucun joueur n'est choisi avant une autre action,
                ' "?" sera ajoute automatiquement.
                ' La possession passe ensuite a Adv.
        
                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False
        
            Else
        
                ' Touche a eux pas droite.
                '
                ' Aucun joueur demande.
                ' La ligne est creee immediatement.
                ' La possession passe a Nous.
        
                FermerActionEnAttente
        
                SetCurrentTeam "Adv"
        
                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo
        
                SetCurrentTeam "Nous"
        
                ReleaseButton
        
            End If
        
            Exit Sub
        
        End If

        ' =====================================================
        ' MELEES
        ' =====================================================
        
        If Not Intersect( _
            Target, _
            Range("BTN_ME_G") _
        ) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "Selectionne d'abord l'equipe qui introduit la melee.", _
                    vbExclamation, _
                    "Possession manquante"
        
                ReleaseButton
                Exit Sub
        
            End If
        
            FermerActionEnAttente
        
            AjouterAction _
                "", _
                ActionTexte, _
                TempsVideo
                
            SetCurrentTeam "Nous"
        
            ReleaseButton
            Exit Sub
        
        End If
        
        
        If Not Intersect( _
            Target, _
            Range("BTN_ME_P") _
        ) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "Selectionne d'abord l'equipe qui introduit la melee.", _
                    vbExclamation, _
                    "Possession manquante"
        
                ReleaseButton
                Exit Sub
        
            End If
        
            FermerActionEnAttente
        
            AjouterAction _
                "", _
                ActionTexte, _
                TempsVideo
        
            SetCurrentTeam "Adv"
        
            ReleaseButton
            Exit Sub
        
        End If

        ' =====================================================
        ' PENALITE POUR NOUS
        '
        ' Aucun joueur demande.
        ' La possession passe a Nous.
        ' =====================================================

         If Not Intersect( _
            Target, _
            Range("BTN_PEN_CONTRE_ADV") _
        ) Is Nothing Then

            InitialiserPenalite _
                ActionTexte, _
                TempsVideo, _
                "Nous", _
                False

            Exit Sub

        End If

        ' =====================================================
        ' PENALITE POUR ADV
        '
        ' Attend le joueur de notre equipe ayant commis la faute.
        ' Si aucun joueur n'est saisi, "?" sera ajoute.
        ' =====================================================

        If Not Intersect( _
            Target, _
            Range("BTN_PEN_CONTRE_NOUS") _
        ) Is Nothing Then

            InitialiserPenalite _
                ActionTexte, _
                TempsVideo, _
                "Adv", _
                True

            Exit Sub

        End If

        ' =====================================================
        ' GRATTAGE / CONTRE-RUCK SANS PENALITE
        ' =====================================================

        If IsTurnoverAction(Target) Then

            If CurrentTeam = "" Then
                MsgBox _
                    "S" _
                    & ChrW(233) & _
                    "lectionne d'abord l'" _
                    & ChrW(233) & "quipe qui poss" _
                    & ChrW(232) & _
                    "de le ballon.", _
                    vbExclamation, _
                    "Possession manquante"

                ReleaseButton
                Exit Sub
            End If

            If CurrentTeam = "Adv" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "", _
                    False

            Else

                FermerActionEnAttente

                AjouterAction "", ActionTexte, TempsVideo
                ToggleCurrentTeam

                ReleaseButton

            End If

            Exit Sub

        End If

        ' =====================================================
        ' PLAQUAGES
        '
        ' Possession Adv : attend un ou plusieurs joueurs.
        ' Possession Nous : saisie immediate sans joueur.
        ' =====================================================

        If IsPlaquageAction(Target) Then

            If CurrentTeam = "" Then
                MsgBox _
                    "S" _
                    & ChrW(233) & _
                    "lectionne d'abord l'" _
                    & ChrW(233) & _
                    "quipe qui poss" _
                    & ChrW(232) & _
                    "de le ballon.", _
                    vbExclamation, _
                    "Possession manquante"

                ReleaseButton
                Exit Sub
            End If

            If CurrentTeam = "Adv" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    False, _
                    "", _
                    True

            Else

                FermerActionEnAttente

                AjouterAction "", ActionTexte, TempsVideo
                ReleaseButton

            End If

            Exit Sub

        End If

        ' =====================================================
        ' ARRACHAGE
        ' =====================================================
        
        If Not Intersect( _
            Target, _
            Range("BTN_ARRACHAGE") _
        ) Is Nothing Then
        
            If CurrentTeam = "" Then
        
                MsgBox _
                    "Selectionne d'abord l'equipe qui possede le ballon.", _
                    vbExclamation, _
                    "Possession manquante"
        
                ReleaseButton
                Exit Sub
        
            End If
        
            If CurrentTeam = "Nous" Then
        
                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False
        
            Else
        
                FermerActionEnAttente
        
                SetCurrentTeam "Adv"
                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo
        
                SetCurrentTeam "Nous"
        
                ReleaseButton
        
            End If
        
            Exit Sub
        
        End If

        ' =====================================================
        ' TURN-OVER
        '
        ' Possession Nous : attend le joueur qui perd le ballon.
        ' Si une autre action est cliquee avant,
        ' "?" est ajoute automatiquement.
        ' Possession Adv : recuperation enregistree
        ' immediatement, sans joueur.
        ' La possession bascule dans les deux cas.
        ' =====================================================

        If Not Intersect( _
            Target, _
            Range("BTN_TURNOVER") _
        ) Is Nothing Then

            If CurrentTeam = "Nous" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False

            Else

                FermerActionEnAttente

                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo

                ToggleCurrentTeam

                ReleaseButton

            End If

            Exit Sub

        End If

        ' =====================================================
        ' ARRET DU JEU
        ' =====================================================
        
        If Not Intersect( _
            Target, _
            Range("BTN_ARRET_DU_JEU") _
        ) Is Nothing Then
        
            FermerActionEnAttente
        
            AjouterArretJeuAutomatique _
                TempsVideo
        
            ReleaseButton
            Exit Sub
        
        End If
        
        ' =====================================================
        ' ESSAI
        '
        ' Possession obligatoire.
        ' Possession Nous : attend le joueur qui marque.
        ' Si une autre action est cliquee avant,
        ' "?" est ajoute automatiquement.
        ' Possession Adv : saisie immediate sans joueur.
        ' =====================================================

        If Not Intersect( _
            Target, _
            Range("BTN_PTS_ESSAI") _
        ) Is Nothing Then

            If CurrentTeam = "" Then

                MsgBox _
                    "S" & ChrW(233) & _
                    "lectionne d'abord l'" & ChrW(233) & _
                    "quipe qui poss" & ChrW(232) & _
                    "de le ballon.", _
                    vbExclamation, _
                    "Possession manquante"

                ReleaseButton
                Exit Sub

            End If

            If CurrentTeam = "Nous" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False

            Else

                FermerActionEnAttente

                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo
                    
                ToggleCurrentTeam

                If ActionDeclencheArretAutomatique(ActionTexte) Then

                    AjouterArretJeuAutomatique _
                        TempsVideo

                End If

                ReleaseButton

            End If

            Exit Sub

        End If
        
        ' =====================================================
        ' TRANSFORMATION
        '
        ' Possession Nous : attend le buteur.
        ' Si une autre action est cliquee avant,
        ' "?" est ajoute automatiquement.
        ' Possession Adv : saisie immediate sans joueur.
        ' =====================================================

        If Not Intersect( _
            Target, _
            Range("BTN_PTS_TRANSFO") _
        ) Is Nothing Then

            ' L'essai a deja fait basculer la possession :
            ' celui qui transforme est donc l'equipe qui
            ' vient de perdre le ballon.
            If CurrentTeam = "Adv" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False

            Else

                FermerActionEnAttente

                SetCurrentTeam "Adv"

                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo

                ToggleCurrentTeam

                ReleaseButton

            End If

            Exit Sub

        End If

        ' =====================================================
        ' DROP
        '
        ' Possession Nous : attend le buteur.
        ' Si une autre action est cliquee avant,
        ' "?" est ajoute automatiquement.
        ' Possession Adv : saisie immediate sans joueur.
        ' La possession bascule : le renvoi revient a
        ' l'equipe qui encaisse.
        ' =====================================================

        If Not Intersect( _
            Target, _
            Range("BTN_PTS_DROP") _
        ) Is Nothing Then

            If CurrentTeam = "Nous" Then

                InitialiserActionEnAttente _
                    ActionTexte, _
                    TempsVideo, _
                    True, _
                    "Nous", _
                    False

            Else

                FermerActionEnAttente

                AjouterAction _
                    "", _
                    ActionTexte, _
                    TempsVideo

                ToggleCurrentTeam

                If ActionDeclencheArretAutomatique(ActionTexte) Then

                    AjouterArretJeuAutomatique _
                        TempsVideo

                End If

                ReleaseButton

            End If

            Exit Sub

        End If

        ' =====================================================
        ' AUTRES ACTIONS IMMEDIATES
        ' =====================================================

        If IsImmediateAction(Target) Then
        
            FermerActionEnAttente
        
            AjouterAction _
                "", _
                ActionTexte, _
                TempsVideo
            
            If EstActionSuivieDunRenvoi(ActionTexte) Then
                ToggleCurrentTeam
            End If
            
            If ActionDeclencheArretAutomatique( _
                ActionTexte _
            ) Then
        
                AjouterArretJeuAutomatique _
                    TempsVideo
        
            End If
        
            ReleaseButton
            Exit Sub
        
        End If

        ' =====================================================
        ' AUTRES ACTIONS
        ' =====================================================

        If CurrentTeam = "Nous" Then

            InitialiserActionEnAttente _
                ActionTexte, _
                TempsVideo, _
                EstActionSuivieDunRenvoi(ActionTexte), _
                "", _
                True

        Else

            FermerActionEnAttente

            AjouterAction "", ActionTexte, TempsVideo
            
            If EstActionSuivieDunRenvoi(ActionTexte) Then
                ToggleCurrentTeam
            End If
            
            ReleaseButton

        End If

        Exit Sub

    End If

    ' =====================================================
    ' CLIC SUR UN JOUEUR
    ' =====================================================

    If WaitingForPlayer Then

        If EstBoutonJoueur(Target) Then

            If InStr(1, CStr(Target.Value), vbLf) > 0 Then

                NumeroJoueur = Left( _
                    CStr(Target.Value), _
                    InStr(1, CStr(Target.Value), vbLf) - 1 _
                )

            Else

                NumeroJoueur = CStr(Target.Value)

            End If

            NumeroJoueur = Trim(NumeroJoueur)

            If NumeroJoueur = "" Then
                Exit Sub
            End If

            joueur = GetPlayerName(NumeroJoueur)

            If PendingActionTimeAvailable Then

                TempsVideo = PendingActionTime

                PendingActionTimeAvailable = False
                PendingActionTime = 0

            Else

                TempsVideo = GetTimeVideo()

                If TempsVideo < 0 Then
                    MsgBox _
                        "Impossible de r" _
                    & ChrW(233) & _
                    "cup" _
                    & ChrW(233) & _
                    "rer le temps r" _
                    & ChrW(233) & _
                    "el de la vid" _
                    & ChrW(233) & _
                    "o.", _
                        vbExclamation, _
                        "VeoVideoControl"
                    Exit Sub
                End If

            End If

            If PendingForcedTeam <> "" Then
                SetCurrentTeam PendingForcedTeam
            End If

            GroupeFautif = ""

            If PendingGroupeFautifRequired Then
                GroupeFautif = GetPlayerGroup(joueur)
            End If

            ActionJoueurFinalisee = PendingAction
            
            AjouterArretApresJoueur = _
                ActionDeclencheArretAutomatique( _
                    ActionJoueurFinalisee _
                )
            
            AjouterAction _
                joueur, _
                ActionJoueurFinalisee, _
                TempsVideo, _
                PendingMotifPenalite, _
                GroupeFautif
            
            PendingPlayerCount = PendingPlayerCount + 1
            ActionFermee = False
            
            If PendingActionChangesPossession Then
            
                ToggleCurrentTeam
                FermerActionEnAttente
                ActionFermee = True
            
            ElseIf PendingForcedTeam <> "" Then
            
                FermerActionEnAttente
                ActionFermee = True
            
            ElseIf Not PendingAllowMultiplePlayers Then
            
                FermerActionEnAttente
                ActionFermee = True
            
            End If
            
            If AjouterArretApresJoueur Then
            
                ' Doit etre execute apres la fermeture de l'action,
                ' afin que le bouton Arret reste visuellement actif.
                AjouterArretJeuAutomatique _
                    TempsVideo
            
            ElseIf ActionFermee Then
            
                ResetActionButtons
            
            End If

            ReleaseButton

        End If

    End If

End Sub

Public Sub ActualiserListeJoueursJournal()

    Dim wsParam As Worksheet
    Dim wsCompo As Worksheet
    Dim loListe As ListObject
    Dim rngCompo As Range
    Dim i As Long
    Dim NouvelleLigne As ListRow

    Set wsParam = shParametres
    Set wsCompo = Sheets("Compo")

    Set loListe = wsParam.ListObjects("LstJoueursJournal")
    Set rngCompo = wsCompo.Range("COMPO")

    ' Supprime toutes les anciennes lignes du tableau.
    Do While loListe.ListRows.Count > 0
        loListe.ListRows(1).Delete
    Loop

    ' Ajoute les joueurs pr_sents dans la composition.
    For i = 1 To rngCompo.Rows.Count

        If Trim(CStr(rngCompo.Cells(i, 2).Value)) <> "" Then

            Set NouvelleLigne = loListe.ListRows.Add

            NouvelleLigne.Range.Cells(1, 1).Value = _
                Trim(CStr(rngCompo.Cells(i, 2).Value))

        End If

    Next i

    ' Ajoute Collectif.
    Set NouvelleLigne = loListe.ListRows.Add
    NouvelleLigne.Range.Cells(1, 1).Value = "Collectif"

    ' Ajoute ?.
    Set NouvelleLigne = loListe.ListRows.Add
    NouvelleLigne.Range.Cells(1, 1).Value = "?"

End Sub

Private Function EstBoutonJoueur(ByVal Target As Range) As Boolean

    Dim Boutons As Range

    If Target.Cells.Count <> 1 Then
        Exit Function
    End If

    Set Boutons = PlageBoutonsJoueurs(Me)

    If Boutons Is Nothing Then
        Exit Function
    End If

    If Intersect(Target, Boutons) Is Nothing Then
        Exit Function
    End If

    EstBoutonJoueur = (Trim(CStr(Target.Value)) <> "")

End Function



