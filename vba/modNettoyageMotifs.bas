Attribute VB_Name = "modNettoyageMotifs"
Option Explicit

' =========================================================
' RETRAIT DES BOUTONS DE MOTIF
'
' A lancer une seule fois, puis ce module peut etre
' supprime.
'
' Les motifs de penalite se choisissent desormais dans la
' popup : les cinq boutons de la palette et leur libelle
' n'ont plus d'objet. Cette macro efface leurs cellules
' puis supprime leurs plages nommees.
'
' Elle passe par les noms tant qu'ils existent : une fois
' supprimes, plus rien ne dit ou etaient les boutons.
' =========================================================

Public Sub RetirerBoutonsMotif()

    Dim Noms As Variant
    Dim i As Long
    Dim Cible As Range
    Dim Effaces As Long

    Noms = Array( _
        "BTN_PEN_MAUL", _
        "BTN_PEN_RUCK", _
        "BTN_PEN_HORS_JEU", _
        "BTN_PEN_PL_A_2", _
        "BTN_PEN_PL_HAUT", _
        "LBL_MOTIF_PENALITE")

    For i = LBound(Noms) To UBound(Noms)

        Set Cible = Nothing

        On Error Resume Next
        Set Cible = ThisWorkbook.Names(CStr(Noms(i))) _
            .RefersToRange
        On Error GoTo 0

        If Not Cible Is Nothing Then

            ' La zone fusionnee entiere, sans quoi Excel
            ' refuse d'en effacer le format.
            If Cible.MergeCells Then Set Cible = Cible.MergeArea

            Cible.UnMerge
            Cible.Clear
            Cible.Interior.Pattern = xlNone
            Cible.Borders.LineStyle = xlNone

            Effaces = Effaces + 1

        End If

        On Error Resume Next
        ThisWorkbook.Names(CStr(Noms(i))).Delete
        On Error GoTo 0

    Next i

    MsgBox _
        Effaces & " boutons de motif retires de la " & _
        "palette." & vbCrLf & vbCrLf & _
        "Relance ensuite BasculerDispositionSaisieVideo " & _
        "deux fois, pour que les deux modes se replacent.", _
        vbInformation, _
        "Motifs de penalite"

End Sub
