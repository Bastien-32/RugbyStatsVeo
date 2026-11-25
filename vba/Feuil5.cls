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

    EtatEvenements = Application.EnableEvents

    If ToucheTempsDeJeu(Target) Then

        Application.EnableEvents = False
        RecalculerTempsDeJeu
        Application.EnableEvents = EtatEvenements

    End If

    Set ColNomJoueur = _
        Me.ListObjects("COMPO") _
            .ListColumns("Nom du joueur") _
            .DataBodyRange

    If Intersect(Target, ColNomJoueur) Is Nothing Then
        Exit Sub
    End If

    Application.EnableEvents = False

    shSaisieVideo.ActualiserListeJoueursJournal
    ActualiserPaletteJoueurs

SortiePropre:

    Application.EnableEvents = EtatEvenements

End Sub


' La duree du match, ou une minute ou une marque de
' remplacement corrigee a la main : colonnes I et
' suivantes, de la ligne des minutes (6) a la derniere
' ligne de joueur (28).
Private Function ToucheTempsDeJeu(ByVal Target As Range) As Boolean

    Dim Zone As Range

    On Error Resume Next
    Set Zone = Me.Range("MATCH_DUREE")
    On Error GoTo 0

    If Not Zone Is Nothing Then
        If Not Intersect(Target, Zone) Is Nothing Then
            ToucheTempsDeJeu = True
            Exit Function
        End If
    End If

    Set Zone = Me.Range(Me.Cells(6, 9), Me.Cells(28, Me.Columns.Count))

    ToucheTempsDeJeu = Not Intersect(Target, Zone) Is Nothing

End Function


