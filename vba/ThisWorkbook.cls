Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

Private Sub Workbook_Open()

    InitialiserClasseurStatsRugby

End Sub


Private Sub Workbook_WindowActivate( _
    ByVal Wn As Window _
)

    If FermetureEnCours Then
        Exit Sub
    End If

    On Error Resume Next

    If Wn.ActiveSheet Is shSaisieVideo Then

        ActiverRaccourcisVideo
        DemarrerAffichageChronoVideo

    Else

        DesactiverRaccourcisVideo
        ArreterAffichageChronoVideo

    End If

    On Error GoTo 0

End Sub


Private Sub Workbook_SheetActivate( _
    ByVal Sh As Object _
)

    If FermetureEnCours Then
        Exit Sub
    End If

    If Sh Is shSaisieVideo Then

        ActiverRaccourcisVideo
        DemarrerAffichageChronoVideo

    Else

        DesactiverRaccourcisVideo
        ArreterAffichageChronoVideo

    End If

End Sub


Private Sub Workbook_BeforeClose( _
    Cancel As Boolean _
)

    FermetureEnCours = True

    ArreterHeartbeatVideo
    DesactiverRaccourcisVideo
    ArreterAffichageChronoVideo

End Sub

