Attribute VB_Name = "modParametres"
Option Explicit

Public Sub NommerCellulesDepuisParametres()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nomCellule As String
    Dim Cible As Range
    Dim nbCrees As Long
    Dim formuleRef As String

    Set ws = shParametres

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow

        nomCellule = Trim(ws.Cells(i, "A").Value)

        If nomCellule <> "" Then

            Set Cible = ws.Cells(i, "B")

            formuleRef = "='" & ws.Name & "'!" & Cible.Address(True, True)

            On Error Resume Next
            ThisWorkbook.Names(nomCellule).Delete
            On Error GoTo 0

            On Error GoTo ErreurNom

            ThisWorkbook.Names.Add _
                Name:=nomCellule, _
                RefersToR1C1:="='" & ws.Name & "'!R" & Cible.Row & "C" & Cible.Column

            nbCrees = nbCrees + 1

            On Error GoTo 0

        End If

    Next i

    MsgBox nbCrees & " noms cr" _
                    & ChrW(233) & ChrW(233) & _
                    "s ou mis " & ChrW(224) & " jour.", vbInformation
    Exit Sub

ErreurNom:
    MsgBox "Impossible de cr" _
                    & ChrW(233) & _
                    "er le nom : " & nomCellule & vbCrLf & _
                    "R" _
                    & ChrW(233) & _
                    "f" _
                    & ChrW(233) & _
                    "rence : " & formuleRef, vbCritical

End Sub

