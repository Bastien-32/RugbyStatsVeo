Attribute VB_Name = "modJournal"
Option Explicit

Public Sub AjouterAction( _
    ByVal joueur As String, _
    ByVal ActionTexte As String, _
    ByVal TempsVideoSecondes As Double, _
    Optional ByVal MotifPenalite As String = "", _
    Optional ByVal GroupeFautif As String = "")

    Dim ws As Worksheet
    Dim wsCompo As Worksheet
    Dim lo As ListObject
    Dim NouvelleLigne As ListRow
    Dim MatchID As String

    Set ws = Sheets("Journal actions")
    Set wsCompo = Sheets("Compo")
    Set lo = ws.ListObjects("JournalActions")

    MatchID = Trim(CStr(wsCompo.Range("MATCH_ID").Value))

    If MatchID = "" Then

        MsgBox _
            "Renseigne d'abord les informations du match dans l'onglet Compo.", _
            vbExclamation, _
            "Match non identifi" & ChrW(232)

        Exit Sub

    End If

    Set NouvelleLigne = lo.ListRows.Add(Position:=1)

    With NouvelleLigne.Range

        ' =====================================================
        ' IDENTITE DU MATCH
        ' =====================================================

        .Cells(1, lo.ListColumns("ID match").Index).Value = _
            MatchID

        .Cells(1, lo.ListColumns("Saison").Index).Value = _
            wsCompo.Range("MATCH_SAISON").Value

        .Cells(1, lo.ListColumns("Date match").Index).Value = _
            wsCompo.Range("MATCH_DATE").Value

        .Cells(1, lo.ListColumns("Date match").Index).NumberFormat = _
            "dd/mm/yyyy"

        .Cells(1, lo.ListColumns("Adversaire").Index).Value = _
            wsCompo.Range("MATCH_ADV").Value

        .Cells(1, lo.ListColumns("Lieu").Index).Value = _
            wsCompo.Range("MATCH_LIEU").Value

        .Cells(1, lo.ListColumns("Phase").Index).Value = _
            wsCompo.Range("MATCH_PHASE").Value

        .Cells(1, lo.ListColumns("Journee").Index).Value = _
            wsCompo.Range("MATCH_JOURNEE").Value

        .Cells(1, lo.ListColumns("Resultat").Index).Value = _
            wsCompo.Range("MATCH_RESULTAT").Value

        ' =====================================================
        ' INFORMATIONS DE L'ACTION
        ' =====================================================

        .Cells(1, lo.ListColumns("Temps video").Index).Value = _
            TempsVideoSecondes / 86400#

        .Cells(1, lo.ListColumns("Temps video").Index).NumberFormat = _
            "[m]:ss"

        .Cells(1, lo.ListColumns("Mi-temps").Index).Value = _
            CurrentHalf

        .Cells(1, lo.ListColumns("Joueur").Index).Value = _
            joueur

        .Cells(1, lo.ListColumns("Groupe fautif").Index).Value = _
            GroupeFautif

        .Cells(1, lo.ListColumns("Action").Index).Value = _
            ActionTexte

        .Cells(1, lo.ListColumns("Motif penalite").Index).Value = _
            MotifPenalite

        .Cells(1, lo.ListColumns("Possession").Index).Value = _
            CurrentTeam

        .Cells(1, lo.ListColumns("Commentaire").Index).Value = ""

        .Cells(1, lo.ListColumns("Horodatage saisie").Index).Value = _
            Now

        .Cells(1, lo.ListColumns("Horodatage saisie").Index).NumberFormat = _
            "dd/mm/yyyy hh:mm"

    End With

    AppliquerValidationsJournal NouvelleLigne.Range, lo

End Sub

Public Sub AppliquerValidationsJournal( _
    ByVal LigneJournal As Range, _
    ByVal lo As ListObject)

    AppliquerValidationListe _
        LigneJournal.Cells(1, lo.ListColumns("Mi-temps").Index), _
        "LISTE_MT_JOURNAL"

    AppliquerValidationListe _
        LigneJournal.Cells(1, lo.ListColumns("Joueur").Index), _
        "LISTE_JOUEURS_JOURNAL"

    AppliquerValidationListe _
        LigneJournal.Cells(1, lo.ListColumns("Action").Index), _
        "LISTE_ACTIONS_JOURNAL"

    AppliquerValidationListe _
        LigneJournal.Cells(1, lo.ListColumns("Possession").Index), _
        "LISTE_EQUIPES_JOURNAL"
        
    AppliquerValidationListe _
        LigneJournal.Cells(1, lo.ListColumns("Motif penalite").Index), _
        "LISTE_MOTIFS_PENALITE"

End Sub

Private Sub AppliquerValidationListe( _
    ByVal Cellule As Range, _
    ByVal NomListe As String)

    Dim nomExiste As Boolean
    Dim nomTest As Name

    nomExiste = False

    For Each nomTest In ThisWorkbook.Names

        If UCase(Split(nomTest.Name, "!")(UBound(Split(nomTest.Name, "!")))) = _
            UCase(NomListe) Then

            nomExiste = True
            Exit For

        End If

    Next nomTest

    If Not nomExiste Then

        MsgBox _
            "La plage nomm" & ChrW(233) & _
            "e """ & NomListe & """ est introuvable.", _
            vbExclamation, _
            "Validation du journal"

        Exit Sub

    End If

    With Cellule.Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="=" & NomListe

        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True

    End With

End Sub

