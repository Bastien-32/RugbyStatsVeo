Attribute VB_Name = "modDureesJeu"
Option Explicit

' =========================================================
' DUREES DES SEQUENCES DE JEU
'
' Une sequence de jeu commence a une reprise, quelle qu'elle
' soit (touche, melee, renvoi, reprise a la main), et se
' termine au premier arret du jeu qui suit.
'
' Comme pour le temps de jeu effectif, la mi-temps ne demarre
' qu'au premier renvoi : tout ce qui precede est ignore.
' La somme des durees comptees ici est donc egale au temps
' de jeu effectif affiche plus haut.
'
' Le tableau est ecrit a droite de celui des durees de
' possession. Sa ligne n'est pas figee : dans un classeur
' recapitulatif, l'agrandissement des tableaux de joueurs
' l'a fait descendre. On la retrouve donc en cherchant
' l'en-tete "Nous" / "Adv" des possessions.
' =========================================================

' La colonne des libelles est fusionnee sur deux cellules :
' les colonnes F a J sont etroites, heritees des tableaux de
' joueurs, et les intitules n'y tiendraient pas. E reste
' vide, comme separation avec le tableau des possessions.
Private Const COLONNE_LIBELLES As String = "F"

Private Const COLONNE_LIBELLES_FIN As String = "G"

Private Const COLONNE_MT1 As String = "H"

Private Const COLONNE_MT2 As String = "I"

Private Const COLONNE_TOTAL As String = "J"

' Tranches plus fines que celles des possessions :
' 0-20, 20-40, 40-60, 60-90, 90-120, 120-150, 150-180, >180.
Private Const NB_TRANCHES As Long = 8


Public Sub RecalculerStatsDureesJeu()

    Dim CompteursMT1(1 To NB_TRANCHES) As Long
    Dim CompteursMT2(1 To NB_TRANCHES) As Long

    Dim LigneEntete As Long

    LigneEntete = LigneEnteteDureesPossession()

    If LigneEntete = 0 Then

        MsgBox _
            "Le tableau des dur" & ChrW(233) & "es de possession " & _
            "est introuvable dans Stats match." & vbCrLf & vbCrLf & _
            "Le tableau des dur" & ChrW(233) & "es de jeu " & _
            "ne peut pas " & ChrW(234) & "tre plac" & ChrW(233) & ".", _
            vbExclamation, _
            "Dur" & ChrW(233) & "es de jeu"

        Exit Sub

    End If

    CompterSequencesJeuMiTemps "MT1", CompteursMT1
    CompterSequencesJeuMiTemps "MT2", CompteursMT2

    EcrireTableauDureesJeu _
        LigneEntete, _
        CompteursMT1, _
        CompteursMT2

End Sub


' L'en-tete du tableau des possessions est la seule ligne
' portant "Nous" en colonne C et "Adv" en colonne D.
Private Function LigneEnteteDureesPossession() As Long

    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Stats match")

    For i = 1 To 400

        If StrComp( _
            Trim(CStr(ws.Cells(i, "C").Value)), _
            "Nous", _
            vbTextCompare _
        ) = 0 Then

            If StrComp( _
                Trim(CStr(ws.Cells(i, "D").Value)), _
                "Adv", _
                vbTextCompare _
            ) = 0 Then

                LigneEnteteDureesPossession = i
                Exit Function

            End If

        End If

    Next i

End Function


Private Sub CompterSequencesJeuMiTemps( _
    ByVal MiTempsRecherchee As String, _
    ByRef Compteurs() As Long _
)

    Dim loJournal As ListObject

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColAction As Long

    Dim i As Long
    Dim NbEvenements As Long

    Dim TempsEvenements() As Double
    Dim ActionsEvenements() As String

    Dim ActionArret As String
    Dim ActionRepriseRenvoi As String

    Dim MiTempsLigne As String
    Dim ActionTexte As String

    Dim JeuDemarre As Boolean
    Dim JeuEnCours As Boolean

    Dim DebutPeriodeJeu As Double
    Dim TempsAction As Double
    Dim Duree As Double

    Set loJournal = _
        ThisWorkbook _
            .Worksheets("Journal actions") _
            .ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then
        Exit Sub
    End If

    ColTemps = loJournal.ListColumns("temps video").Index
    ColMiTemps = loJournal.ListColumns("Mi-temps").Index
    ColAction = loJournal.ListColumns("Action").Index

    ActionArret = _
        CStr(Range("ACT_ARRET_DU_JEU").Value)

    ActionRepriseRenvoi = _
        CStr(Range("ACT_REPRISE_RENVOI").Value)

    ' =====================================================
    ' 1. EVENEMENTS DE LA MI-TEMPS
    ' =====================================================

    For i = 1 To loJournal.ListRows.Count

        MiTempsLigne = _
            Trim(CStr( _
                loJournal.DataBodyRange.Cells(i, ColMiTemps).Value _
            ))

        If StrComp( _
            MiTempsLigne, _
            MiTempsRecherchee, _
            vbTextCompare _
        ) = 0 Then

            If IsNumeric( _
                loJournal.DataBodyRange.Cells(i, ColTemps).Value _
            ) Then

                ActionTexte = _
                    Trim(CStr( _
                        loJournal.DataBodyRange.Cells(i, ColAction).Value _
                    ))

                If ActionTexte <> "" Then

                    NbEvenements = NbEvenements + 1

                    ReDim Preserve TempsEvenements(1 To NbEvenements)
                    ReDim Preserve ActionsEvenements(1 To NbEvenements)

                    TempsEvenements(NbEvenements) = _
                        CDbl( _
                            loJournal.DataBodyRange.Cells(i, ColTemps).Value _
                        )

                    ActionsEvenements(NbEvenements) = ActionTexte

                End If

            End If

        End If

    Next i

    If NbEvenements = 0 Then
        Exit Sub
    End If

    ' =====================================================
    ' 2. TRI CHRONOLOGIQUE
    ' =====================================================

    TrierEvenements _
        TempsEvenements, _
        ActionsEvenements, _
        NbEvenements

    ' =====================================================
    ' 3. DECOUPAGE EN SEQUENCES DE JEU
    ' =====================================================

    For i = 1 To NbEvenements

        TempsAction = TempsEvenements(i)
        ActionTexte = ActionsEvenements(i)

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

            If StrComp( _
                ActionTexte, _
                ActionArret, _
                vbTextCompare _
            ) = 0 Then

                ' La colonne "temps video" est une duree
                ' Excel, donc une fraction de jour :
                ' il faut la convertir en secondes.
                Duree = _
                    (TempsAction - DebutPeriodeJeu) * 86400#

                If Duree >= 0 Then

                    Compteurs(ClasseDureeJeu(Duree)) = _
                        Compteurs(ClasseDureeJeu(Duree)) + 1

                End If

                JeuEnCours = False

            End If

        Else

            If EstActionRepriseJeu(ActionTexte) Then

                DebutPeriodeJeu = TempsAction
                JeuEnCours = True

            End If

        End If

    Next i

End Sub


Private Function ClasseDureeJeu( _
    ByVal DureeSecondes As Double _
) As Long

    If DureeSecondes < 20 Then

        ClasseDureeJeu = 1

    ElseIf DureeSecondes < 40 Then

        ClasseDureeJeu = 2

    ElseIf DureeSecondes < 60 Then

        ClasseDureeJeu = 3

    ElseIf DureeSecondes < 90 Then

        ClasseDureeJeu = 4

    ElseIf DureeSecondes < 120 Then

        ClasseDureeJeu = 5

    ElseIf DureeSecondes < 150 Then

        ClasseDureeJeu = 6

    ElseIf DureeSecondes < 180 Then

        ClasseDureeJeu = 7

    Else

        ClasseDureeJeu = 8

    End If

End Function


Private Function EstActionRepriseJeu( _
    ByVal ActionTexte As String _
) As Boolean

    Dim Reprises As Variant
    Dim Element As Variant

    Reprises = Array( _
        "ACT_REPRISE_TOUCHE", _
        "ACT_REPRISE_MELEE", _
        "ACT_REPRISE_RENVOI", _
        "ACT_REPRISE_A_LA_MAIN" _
    )

    For Each Element In Reprises

        If StrComp( _
            ActionTexte, _
            CStr(Range(CStr(Element)).Value), _
            vbTextCompare _
        ) = 0 Then

            EstActionRepriseJeu = True
            Exit Function

        End If

    Next Element

End Function


Private Sub TrierEvenements( _
    ByRef TempsEvenements() As Double, _
    ByRef ActionsEvenements() As String, _
    ByVal NbEvenements As Long _
)

    Dim i As Long
    Dim j As Long

    Dim TempsTemporaire As Double
    Dim ActionTemporaire As String

    For i = 1 To NbEvenements - 1

        For j = i + 1 To NbEvenements

            If TempsEvenements(j) < TempsEvenements(i) Then

                TempsTemporaire = TempsEvenements(i)
                TempsEvenements(i) = TempsEvenements(j)
                TempsEvenements(j) = TempsTemporaire

                ActionTemporaire = ActionsEvenements(i)
                ActionsEvenements(i) = ActionsEvenements(j)
                ActionsEvenements(j) = ActionTemporaire

            End If

        Next j

    Next i

End Sub


' Chaque intitule occupe deux cellules en largeur. La fusion
' est refaite a chaque recalcul pour rester valable si le
' tableau a ete deplace, et defaite avant d'etre reappliquee
' pour ne pas dependre de l'etat precedent.
Private Sub FusionnerColonneLibelles( _
    ByVal ws As Worksheet, _
    ByVal LigneEntete As Long _
)

    Dim i As Long
    Dim EtatAlertes As Boolean

    EtatAlertes = Application.DisplayAlerts

    Application.DisplayAlerts = False

    For i = 0 To NB_TRANCHES

        With ws.Range( _
            COLONNE_LIBELLES & LigneEntete + i & _
            ":" & COLONNE_LIBELLES_FIN & LigneEntete + i _
        )

            .UnMerge
            .Merge

        End With

    Next i

    Application.DisplayAlerts = EtatAlertes

End Sub


Private Sub EcrireTableauDureesJeu( _
    ByVal LigneEntete As Long, _
    ByRef CompteursMT1() As Long, _
    ByRef CompteursMT2() As Long _
)

    Dim ws As Worksheet
    Dim i As Long
    Dim Libelles As Variant

    Libelles = Array( _
        "0-20 s", _
        "20-40 s", _
        "40-60 s", _
        "60-90 s", _
        "90-120 s", _
        "120-150 s", _
        "150-180 s", _
        ">180 s" _
    )

    Set ws = ThisWorkbook.Worksheets("Stats match")

    With ws

        ' Meme presentation que le tableau des possessions.
        ' Les formats sont repris cellule par cellule, le
        ' tableau des durees de jeu comptant plus de lignes.
        ' La colonne des libelles est encore defusionnee a ce
        ' stade : coller sur une fusion partielle echouerait.
        .Range("B" & LigneEntete).Copy

        .Range( _
            COLONNE_LIBELLES & LigneEntete & _
            ":" & COLONNE_LIBELLES_FIN & LigneEntete _
        ).PasteSpecial Paste:=xlPasteFormats

        .Range("B" & LigneEntete + 1).Copy

        .Range( _
            COLONNE_LIBELLES & LigneEntete + 1 & _
            ":" & COLONNE_LIBELLES_FIN & LigneEntete + NB_TRANCHES _
        ).PasteSpecial Paste:=xlPasteFormats

        .Range("C" & LigneEntete).Copy

        .Range( _
            COLONNE_MT1 & LigneEntete & _
            ":" & COLONNE_TOTAL & LigneEntete _
        ).PasteSpecial Paste:=xlPasteFormats

        .Range("C" & LigneEntete + 1).Copy

        .Range( _
            COLONNE_MT1 & LigneEntete + 1 & _
            ":" & COLONNE_TOTAL & LigneEntete + NB_TRANCHES _
        ).PasteSpecial Paste:=xlPasteFormats

        Application.CutCopyMode = False

    End With

    ' Les formats sont en place : on peut fusionner, puis
    ' ecrire les valeurs dans les cellules ancres.
    FusionnerColonneLibelles ws, LigneEntete

    With ws

        .Range(COLONNE_LIBELLES & LigneEntete).Value = _
            "Dur" & ChrW(233) & "es actions"

        .Range(COLONNE_MT1 & LigneEntete).Value = "MT1"
        .Range(COLONNE_MT2 & LigneEntete).Value = "MT2"
        .Range(COLONNE_TOTAL & LigneEntete).Value = "Total"

        For i = 1 To NB_TRANCHES

            .Range(COLONNE_LIBELLES & LigneEntete + i).Value = _
                Libelles(i - 1)

        Next i

        For i = 1 To NB_TRANCHES

            .Range(COLONNE_MT1 & LigneEntete + i).Value = _
                CompteursMT1(i)

            .Range(COLONNE_MT2 & LigneEntete + i).Value = _
                CompteursMT2(i)

            .Range(COLONNE_TOTAL & LigneEntete + i).Value = _
                CompteursMT1(i) + CompteursMT2(i)

        Next i
        
        With .Range( _
            COLONNE_LIBELLES & LigneEntete & _
            ":" & COLONNE_TOTAL & LigneEntete + NB_TRANCHES _
        )

            .BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlMedium

        End With

    End With

End Sub
