Attribute VB_Name = "Feuil4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Not Intersect(Target, Range("POPUP_TO_VALIDER")) _
        Is Nothing Then

        ValiderPopupTouche

    ElseIf Not Intersect(Target, Range("POPUP_TO_ANNULER")) _
        Is Nothing Then

        AnnulerPopupTouche

    End If

End Sub


