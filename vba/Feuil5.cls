Attribute VB_Name = "Feuil5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim ColNomJoueur As Range
    Dim EtatEvenements As Boolean

    On Error GoTo SortiePropre

    Set ColNomJoueur = _
        Me.ListObjects("COMPO") _
            .ListColumns("Nom du joueur") _
            .DataBodyRange

    If Intersect(Target, ColNomJoueur) Is Nothing Then
        Exit Sub
    End If

    EtatEvenements = Application.EnableEvents
    Application.EnableEvents = False

    shSaisieVideo.ActualiserListeJoueursJournal
    ActualiserPaletteJoueurs

SortiePropre:

    Application.EnableEvents = EtatEvenements

End Sub

