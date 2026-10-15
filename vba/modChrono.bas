Attribute VB_Name = "modChrono"
Option Explicit

Private ProchaineMiseAJour As Date
Private MiseAJourPlanifiee As Boolean

' Now et la date d'un rendez-vous sont des nombres a
' virgule : "Now + une seconde" ne retombe pas exactement
' sur ce que Now rendra une seconde plus tard. Comparer au
' plus juste ferait passer un battement ponctuel pour un
' doublon. Une demi-seconde de marge separe les deux cas
' sans ambiguite.
Private Const MARGE As Double = 0.5 / 86400#


Public Sub DemarrerAffichageChronoVideo()

    ArreterAffichageChronoVideo
    MettreAJourAffichageChronoVideo

End Sub

Public Sub ArreterAffichageChronoVideo()

    On Error Resume Next

    If MiseAJourPlanifiee Then

        Application.OnTime _
            EarliestTime:=ProchaineMiseAJour, _
            Procedure:=NomProcedureChrono(), _
            Schedule:=False

    End If

    MiseAJourPlanifiee = False

    On Error GoTo 0

End Sub

Public Sub MettreAJourAffichageChronoVideo()

    Dim TempsSecondes As Double

    If FermetureEnCours Then Exit Sub

    ' Un rendez-vous rattrape par VerifierChronoVideo peut
    ' se presenter alors qu'un autre bat deja : sans ce
    ' garde-fou, le chrono se mettrait a battre deux fois.
    If MiseAJourPlanifiee Then
        If ProchaineMiseAJour - Now > MARGE Then Exit Sub
    End If

    MiseAJourPlanifiee = False

    ' Le battement suivant est pose avant le travail : une
    ' erreur dans ce qui suit romprait sinon la chaine, que
    ' plus rien ne relancerait.
    PlanifierProchainBattement

    On Error Resume Next

    TempsSecondes = GetTimeVideo()

    If TempsSecondes >= 0 Then

        With shSaisieVideo.Range("CELL_CHRONO_VIDEO")

            .Value = TempsSecondes / 86400#
            .NumberFormat = "[m]:ss"

        End With

        ' La palette suit le temps : les quinze boutons
        ' montrent qui est sur le terrain a cet instant.
        ' Rien n'est reecrit tant que rien ne change.
        EcrirePalettePresents TempsSecondes

    End If

    On Error GoTo 0

End Sub


' Excel abandonne un rendez-vous OnTime quand il est
' occupe a l'heure dite : une cellule en cours de saisie
' suffit. La chaine meurt alors sans bruit et le chrono
' se fige. Ce controle la remet en marche au premier
' clic qui suit.
Public Sub VerifierChronoVideo()

    If FermetureEnCours Then Exit Sub

    If MiseAJourPlanifiee Then
        If Now <= ProchaineMiseAJour + TimeSerial(0, 0, 5) Then
            Exit Sub
        End If
    End If

    PlanifierProchainBattement

End Sub


Private Sub PlanifierProchainBattement()

    On Error Resume Next

    ProchaineMiseAJour = Now + TimeSerial(0, 0, 1)
    MiseAJourPlanifiee = True

    Err.Clear

    Application.OnTime _
        EarliestTime:=ProchaineMiseAJour, _
        Procedure:=NomProcedureChrono(), _
        Schedule:=True

    ' Rendez-vous refuse : le marquer comme absent pour que
    ' le prochain controle retente.
    If Err.Number <> 0 Then MiseAJourPlanifiee = False

    On Error GoTo 0

End Sub

Private Function NomProcedureChrono() As String

    Dim NomClasseurSecurise As String

    NomClasseurSecurise = _
        Replace(ThisWorkbook.Name, "'", "''")

    NomProcedureChrono = _
        "'" & _
        NomClasseurSecurise & _
        "'!MettreAJourAffichageChronoVideo"

End Function
