Attribute VB_Name = "modPlayers"
Option Explicit

Private EtatNavigationPopup As Long

Public Function GetPlayerName( _
    ByVal Numero As String) As String

    Dim rngCompo As Range
    Dim i As Long

    If Numero = "Collectif" _
        Or Numero = "?" Then

        GetPlayerName = Numero
        Exit Function

    End If

    Set rngCompo = Sheets("Compo").Range("COMPO")

    For i = 1 To rngCompo.Rows.Count

        If CStr(rngCompo.Cells(i, 1).Value) = Numero Then

            GetPlayerName = rngCompo.Cells(i, 2).Value
            Exit Function

        End If

    Next i

    GetPlayerName = Numero

End Function

Public Sub OuvrirPopupAjouterJoueur()

    With Sheets("Popup")
        .Visible = True
        .Activate
        .Range("POPUP_NOM").Value = ""
        .Range("POPUP_PRENOM").Value = ""
        .Range("POPUP_NOM").Select
    End With

End Sub

Public Sub AnnulerPopupAjouterJoueur()

    Sheets("Popup").Range("POPUP_NOM").Value = ""
    Sheets("Popup").Range("POPUP_PRENOM").Value = ""
    Sheets("Listes").Activate

End Sub

Public Sub ValiderPopupAjouterJoueur()

    Dim Nom As String
    Dim Prenom As String

    Nom = Trim(Sheets("Popup").Range("POPUP_NOM").Value)
    Prenom = Trim(Sheets("Popup").Range("POPUP_PRENOM").Value)

    If Nom = "" Then
        MsgBox "Le nom est obligatoire.", vbExclamation
        Sheets("Popup").Range("POPUP_NOM").Select
        Exit Sub
    End If

    If Prenom = "" Then
        MsgBox "Le pr" _
                    & ChrW(233) & _
                    "nom est obligatoire.", vbExclamation
        Sheets("Popup").Range("POPUP_PRENOM").Select
        Exit Sub
    End If

    AjouterJoueurDansTableau Nom, Prenom

    Sheets("Popup").Range("POPUP_NOM").Value = ""
    Sheets("Popup").Range("POPUP_PRENOM").Value = ""

    Sheets("Listes").Activate

End Sub

Public Sub AjouterJoueurDansTableau( _
    ByVal Nom As String, _
    ByVal Prenom As String)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim NouvelleLigne As ListRow
    Dim maxId As Long

    Set ws = Sheets("Listes")
    Set lo = ws.ListObjects("LstEffectif")

    If lo.ListRows.Count = 0 Then
        maxId = 0
    Else
        maxId = Application.WorksheetFunction.Max( _
            lo.ListColumns("id").DataBodyRange)
    End If

    Set NouvelleLigne = lo.ListRows.Add

    With NouvelleLigne.Range
        .Cells(1, lo.ListColumns("id").Index).Value = maxId + 1
        .Cells(1, lo.ListColumns("nom").Index).Value = UCase(Trim(Nom))
        .Cells(1, lo.ListColumns("prenom").Index).Value = Trim(Prenom)
        .Cells(1, lo.ListColumns("prenom NOM").Index).Formula = "=[@prenom] & "" "" & [@nom]"
        .Cells(1, lo.ListColumns("NOM Prénom").Index).Formula = "=[@nom] & "" "" & [@prenom]"
    End With

End Sub

Public Sub ActualiserJoueursJournal()

    Dim wsCompo As Worksheet
    Dim wsJournal As Worksheet
    Dim rngCompo As Range
    Dim rngListe As Range
    Dim loJournal As ListObject
    Dim joueurs() As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim ligneDest As Long
    Dim formuleListe As String
    Dim temp As String

    Set wsCompo = Sheets("Compo")
    Set wsJournal = Sheets("Journal actions")
    Set rngCompo = wsCompo.Range("COMPO")
    Set loJournal = wsJournal.ListObjects("JournalActions")

    Set rngListe = wsCompo.Range("AA1:AA30")
    rngListe.ClearContents

    For i = 1 To rngCompo.Rows.Count

        If Trim(rngCompo.Cells(i, 2).Value) <> "" Then
            n = n + 1
            ReDim Preserve joueurs(1 To n)
            joueurs(n) = Trim(rngCompo.Cells(i, 2).Value)
        End If

    Next i

    For i = 1 To n - 1
        For j = i + 1 To n

            If joueurs(i) > joueurs(j) Then
                temp = joueurs(i)
                joueurs(i) = joueurs(j)
                joueurs(j) = temp
            End If

        Next j
    Next i

    ligneDest = 1

    For i = 1 To n
        rngListe.Cells(ligneDest, 1).Value = joueurs(i)
        ligneDest = ligneDest + 1
    Next i

    rngListe.Cells(ligneDest, 1).Value = "Collectif"
    rngListe.Cells(ligneDest + 1, 1).Value = "?"

    wsCompo.Columns("AA").Hidden = True

    formuleListe = "='" & wsCompo.Name & "'!" & _
        wsCompo.Range("AA1:AA" & ligneDest + 1).Address

    With loJournal.ListColumns("Joueur").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=formuleListe
    End With

End Sub

Public Sub ActualiserPaletteJoueurs()

    Dim wsSaisie As Worksheet
    Dim wsCompo As Worksheet
    Dim rngCompo As Range
    Dim i As Long
    Dim Poste As String
    Dim Joueur As String
    Dim nomBouton As String
    Dim Mots() As String
    Dim Mot As Variant
    Dim nomFamille As String

    Set wsSaisie = shSaisieVideo
    Set wsCompo = Sheets("Compo")
    Set rngCompo = wsCompo.Range("COMPO")

    For i = 1 To rngCompo.Rows.Count

        Poste = Trim(CStr(rngCompo.Cells(i, 1).Value))
        Joueur = Trim(CStr(rngCompo.Cells(i, 2).Value))

        If Poste <> "" Then

            nomBouton = "BTN_JO_" & Poste

            nomFamille = ""

            If Joueur <> "" Then

                Mots = Split(Joueur, " ")

                For Each Mot In Mots

                    If CStr(Mot) = UCase(CStr(Mot)) Then

                        If nomFamille = "" Then
                            nomFamille = CStr(Mot)
                        Else
                            nomFamille = nomFamille & " " & CStr(Mot)
                        End If

                    Else

                        Exit For

                    End If

                Next Mot

                Joueur = nomFamille

                wsSaisie.Range(nomBouton).Value = Poste & vbLf & Joueur

            Else

                Joueur = ""
                wsSaisie.Range(nomBouton).Value = Poste

            End If

            With wsSaisie.Range(nomBouton)

                .WrapText = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter

                .Characters(1, Len(Poste)).Font.Size = 18
                .Characters(1, Len(Poste)).Font.Bold = True

                If Joueur <> "" Then
                    .Characters(Len(Poste) + 2, Len(Joueur)).Font.Size = 9
                    .Characters(Len(Poste) + 2, Len(Joueur)).Font.Bold = False
                End If

            End With

        End If

    Next i

End Sub

Public Function GetPlayerGroup(ByVal NomJoueur As String) As String

    Dim rngCompo As Range
    Dim i As Long

    If NomJoueur = "?" Then
        GetPlayerGroup = "?"
        Exit Function
    End If

    Set rngCompo = Sheets("Compo").Range("COMPO")

    For i = 1 To rngCompo.Rows.Count

        If Trim(CStr(rngCompo.Cells(i, 2).Value)) = _
            Trim(NomJoueur) Then

            GetPlayerGroup = _
                Trim(CStr(rngCompo.Cells(i, 3).Value))

            Exit Function

        End If

    Next i

    GetPlayerGroup = "?"

End Function

' ============================================================
' NAVIGATION CLAVIER POPUP AJOUT JOUEUR
'
' 1 = Nom
' 2 = Prénom
' 3 = Ajouter
' 4 = Annuler
' ============================================================


Public Sub ActiverNavigationPopup()

    EtatNavigationPopup = 1

    Application.OnKey "{TAB}", "TabPopupSuivant"
    Application.OnKey "+{TAB}", "TabPopupPrecedent"
    Application.OnKey "~", "EntreePopup"

End Sub


Public Sub DesactiverNavigationPopup()

    Application.OnKey "{TAB}"
    Application.OnKey "+{TAB}"
    Application.OnKey "~"

    EtatNavigationPopup = 0

End Sub


Public Sub TabPopupSuivant()

    If ActiveSheet.Name <> "Popup" Then
        Exit Sub
    End If

    EtatNavigationPopup = _
        EtatNavigationPopup + 1

    If EtatNavigationPopup > 4 Then
        EtatNavigationPopup = 1
    End If

    AfficherEtatNavigationPopup

End Sub


Public Sub TabPopupPrecedent()

    If ActiveSheet.Name <> "Popup" Then
        Exit Sub
    End If

    EtatNavigationPopup = _
        EtatNavigationPopup - 1

    If EtatNavigationPopup < 1 Then
        EtatNavigationPopup = 4
    End If

    AfficherEtatNavigationPopup

End Sub


Private Sub AfficherEtatNavigationPopup()

    Dim ws As Worksheet

    Set ws = Sheets("Popup")

    ' =====================================================
    ' REMISE A ZERO DES BORDURES DES BOUTONS
    ' =====================================================

    With ws.Shapes("BTN_POPUP_AJOUTER").Line
        .Visible = msoFalse
    End With

    With ws.Shapes("BTN_POPUP_ANNULER").Line
        .Visible = msoFalse
    End With

    ' =====================================================
    ' ETAT COURANT
    ' =====================================================

    Select Case EtatNavigationPopup

        Case 1

            ws.Range("POPUP_NOM").Select

        Case 2

            ws.Range("POPUP_PRENOM").Select

        Case 3

            ' Déplace la sélection hors des champs visibles.
            ws.Range("F9").Select

            ' Focus visuel sur AJOUTER.
            With ws.Shapes("BTN_POPUP_AJOUTER").Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 4
            End With

        Case 4

            ' Déplace la sélection hors des champs visibles.
            ws.Range("F9").Select

            ' Focus visuel sur ANNULER.
            With ws.Shapes("BTN_POPUP_ANNULER").Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 4
            End With

    End Select

End Sub


Public Sub EntreePopup()

    If ActiveSheet.Name <> "Popup" Then
        Exit Sub
    End If

    Select Case EtatNavigationPopup

        Case 1

            EtatNavigationPopup = 2

            AfficherEtatNavigationPopup

        Case 2

            ValiderPopupAjouterJoueur

        Case 3

            ValiderPopupAjouterJoueur

        Case 4

            AnnulerPopupAjouterJoueur

    End Select

End Sub

