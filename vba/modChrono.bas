Attribute VB_Name = "modChrono"
Option Explicit

Private ProchaineMiseAJour As Date
Private MiseAJourPlanifiee As Boolean


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
    
    MiseAJourPlanifiee = False

    TempsSecondes = GetTimeVideo()

    If TempsSecondes >= 0 Then

        With shSaisieVideo.Range("CELL_CHRONO_VIDEO")

            .Value = TempsSecondes / 86400#
            .NumberFormat = "[m]:ss"

        End With

    End If

    ProchaineMiseAJour = Now + TimeSerial(0, 0, 1)
    MiseAJourPlanifiee = True

    Application.OnTime _
        EarliestTime:=ProchaineMiseAJour, _
        Procedure:=NomProcedureChrono(), _
        Schedule:=True

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
