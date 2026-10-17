Attribute VB_Name = "modTempsJeu"
Option Explicit

Public Sub RecalculerTempsJeuEffectif()

    Dim TempsMT1 As Double
    Dim TempsMT2 As Double

    Dim EtatAffichage As Boolean
    Dim EtatCalcul As Long

    ' Les trois cellules de resultat alimentent les 1600
    ' NB.SI.ENS de la feuille Stats match. Sans cette
    ' suspension, Excel recalcule la feuille entiere une
    ' deuxieme fois et repeint l'ecran : c'est le temps
    ' mort et le clignotement blanc ressentis a chaque
    ' arret du jeu.
    EtatAffichage = Application.ScreenUpdating
    EtatCalcul = Application.Calculation

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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

    Application.Calculation = EtatCalcul
    Application.ScreenUpdating = EtatAffichage

    Exit Sub

GestionErreur:

    Application.Calculation = EtatCalcul
    Application.ScreenUpdating = EtatAffichage

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
    Dim NbLignes As Long

    Dim Donnees As Variant

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
    '
    ' Le journal est lu d'un seul coup dans un tableau en
    ' memoire. Lire cellule par cellule coutait quatre
    ' allers-retours vers Excel par ligne, soit plus d'un
    ' millier a chaque arret du jeu : c'est ce qui se
    ' voyait a l'ecran.
    '
    ' Les trois tableaux sont dimensionnes une fois pour
    ' toutes, le journal entier servant de borne : un
    ' ReDim Preserve par evenement recopiait tout le
    ' tableau a chaque ajout.
    ' =====================================================

    NbLignes = loJournal.ListRows.Count

    Donnees = loJournal.DataBodyRange.Value

    ReDim TempsEvenements(1 To NbLignes)
    ReDim ActionsEvenements(1 To NbLignes)
    ReDim LignesSource(1 To NbLignes)

    For i = 1 To NbLignes

        MiTempsLigne = _
            Trim(CStr(Donnees(i, ColMiTemps)))

        If StrComp( _
            MiTempsLigne, _
            MiTempsRecherchee, _
            vbTextCompare _
        ) = 0 Then

            If IsNumeric(Donnees(i, ColTemps)) Then

                ActionTexte = _
                    Trim(CStr(Donnees(i, ColAction)))

                If ActionTexte <> "" Then

                    NbEvenements = NbEvenements + 1

                    TempsEvenements(NbEvenements) = _
                        CDbl(Donnees(i, ColTemps))

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

