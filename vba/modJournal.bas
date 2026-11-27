Attribute VB_Name = "modJournal"
Option Explicit

Public Sub AjouterAction( _
    ByVal Joueur As String, _
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
            MiTempsDeLAction(TempsVideoSecondes)

        .Cells(1, lo.ListColumns("Joueur").Index).Value = _
            Joueur

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

    PoserLienTempsVideo _
        NouvelleLigne.Range.Cells(1, lo.ListColumns("Temps video").Index)

    AppliquerValidationsJournal NouvelleLigne.Range, lo

End Sub

' =========================================================
' MI-TEMPS D'UNE ACTION
'
' Une action saisie apres coup, en revenant en arriere
' dans la video, appartient a la mi-temps de ce moment-la,
' et non a celle selectionnee dans la palette. On la lit
' sur les actions qui l'encadrent : si celle d'avant et
' celle d'apres dans la video sont de la meme mi-temps,
' c'est la sienne.
'
' Sinon - journal vide, action la plus recente, ou pile
' pendant la pause - la mi-temps selectionnee fait foi,
' comme lors d'une saisie en direct.
' =========================================================

Public Function MiTempsDeLAction( _
    ByVal TempsVideoSecondes As Double) As String

    Dim lo As ListObject
    Dim Temps As Variant
    Dim MiTemps As Variant
    Dim i As Long
    Dim T As Double
    Dim Avant As Double
    Dim Apres As Double
    Dim MiTempsAvant As String
    Dim MiTempsApres As String

    MiTempsDeLAction = CurrentHalf

    Set lo = Sheets("Journal actions").ListObjects("JournalActions")

    ' Il faut une action de chaque cote, et .Value ne rend
    ' un tableau qu'a partir de deux lignes.
    If lo.ListRows.Count < 2 Then Exit Function

    Temps = lo.ListColumns("Temps video").DataBodyRange.Value
    MiTemps = lo.ListColumns("Mi-temps").DataBodyRange.Value

    Avant = -1
    Apres = -1

    For i = 1 To UBound(Temps, 1)

        If Not IsEmpty(Temps(i, 1)) And IsNumeric(Temps(i, 1)) Then

            ' Le journal stocke des jours.
            T = CDbl(Temps(i, 1)) * 86400#

            ' Une action au meme instant, l'arret de jeu
            ' automatique par exemple, n'encadre rien.
            If T < TempsVideoSecondes And T > Avant Then

                Avant = T
                MiTempsAvant = Trim(CStr(MiTemps(i, 1)))

            ElseIf T > TempsVideoSecondes And _
                (Apres < 0 Or T < Apres) Then

                Apres = T
                MiTempsApres = Trim(CStr(MiTemps(i, 1)))

            End If

        End If

    Next i

    If Avant < 0 Or Apres < 0 Then Exit Function
    If MiTempsAvant = "" Then Exit Function

    If StrComp(MiTempsAvant, MiTempsApres, vbTextCompare) <> 0 Then
        Exit Function
    End If

    MiTempsDeLAction = MiTempsAvant

End Function

' =========================================================
' LIENS VERS LA VIDEO
'
' Chaque temps du journal est un lien vers la cellule
' Joueur de sa ligne : un clic y amene directement, et
' declenche Worksheet_FollowHyperlink, qui positionne la
' video a cet instant.
'
' Un lien ne reagit qu'a la souris : parcourir le journal
' au clavier ne fait pas sauter la video.
'
' Excel fige l'adresse d'un lien a sa pose. Les actions
' s'ajoutant en tete du journal, les lignes descendent et
' leurs liens visent des lignes plus hautes : on les
' recale a l'ouverture du journal.
' =========================================================

Public Sub PoserLienTempsVideo(ByVal Cellule As Range)

    Dim FormatTemps As String

    If Not EstTempsVideo(Cellule) Then Exit Sub

    If Cellule.Hyperlinks.Count > 0 Then

        If Not LienAJour(Cellule) Then
            Cellule.Hyperlinks(1).SubAddress = CibleLien(Cellule)
        End If

        Exit Sub

    End If

    FormatTemps = Cellule.NumberFormat

    Cellule.Worksheet.Hyperlinks.Add _
        Anchor:=Cellule, _
        Address:="", _
        SubAddress:=CibleLien(Cellule), _
        ScreenTip:="Aller " & ChrW(224) & " ce moment de la vid" & _
            ChrW(233) & "o"

    ' Le style Lien hypertexte ne doit pas changer
    ' l'affichage du temps.
    Cellule.NumberFormat = FormatTemps

End Sub

' Pose les liens manquants, ceux des lignes saisies avant
' leur apparition, et recale ceux que les ajouts ont
' decales. Les ajouts se faisant en tete, la premiere et
' la derniere ligne a jour suffisent a tout savoir a jour.
Public Sub PoserLiensTempsVideo()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim Temps As Range
    Dim Cellule As Range

    Set ws = Sheets("Journal actions")
    Set lo = ws.ListObjects("JournalActions")

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set Temps = lo.ListColumns("Temps video").DataBodyRange

    If ws.Hyperlinks.Count >= lo.ListRows.Count Then

        If LienAJour(Temps.Cells(1)) And _
           LienAJour(Temps.Cells(Temps.Cells.Count)) Then

            Exit Sub

        End If

    End If

    Application.ScreenUpdating = False

    For Each Cellule In Temps.Cells

        PoserLienTempsVideo Cellule

    Next Cellule

    Application.ScreenUpdating = True

End Sub

Public Sub SuivreLienTempsVideo(ByVal Lien As Hyperlink)

    Dim Cellule As Range

    Set Cellule = Lien.Range

    If Not EstTempsVideo(Cellule) Then Exit Sub

    ' Le journal stocke des jours, le moteur attend des
    ' secondes.
    AllerAuTempsVideo Cellule.Value * 86400#

    ' Filet de securite : une ligne supprimee depuis
    ' l'ouverture du journal a pu decaler le lien.
    If Not LienAJour(Cellule) Then
        CelluleJoueur(Cellule).Select
    End If

End Sub

Private Function CelluleJoueur(ByVal Cellule As Range) As Range

    Dim lo As ListObject

    Set lo = Cellule.Worksheet.ListObjects("JournalActions")

    Set CelluleJoueur = Cellule.Worksheet.Cells( _
        Cellule.Row, _
        lo.ListColumns("Joueur").Range.Column)

End Function

Private Function CibleLien(ByVal Cellule As Range) As String

    CibleLien = _
        "'" & Cellule.Worksheet.Name & "'!" & _
        CelluleJoueur(Cellule).Address(False, False)

End Function

' Compare la cible du lien a la cellule Joueur de sa
' ligne, quelle que soit la facon dont Excel a reecrit
' l'adresse (guillemets, dollars).
Private Function LienAJour(ByVal Cellule As Range) As Boolean

    Dim Cible As String

    If Cellule.Hyperlinks.Count = 0 Then Exit Function

    Cible = Cellule.Hyperlinks(1).SubAddress
    Cible = Mid(Cible, InStr(Cible, "!") + 1)
    Cible = Replace(Cible, "$", "")

    LienAJour = _
        (Cible = CelluleJoueur(Cellule).Address(False, False))

End Function

Private Function EstTempsVideo(ByVal Cellule As Range) As Boolean

    Dim lo As ListObject

    Set lo = Cellule.Worksheet.ListObjects("JournalActions")

    If lo.DataBodyRange Is Nothing Then Exit Function

    If Intersect( _
        Cellule, _
        lo.ListColumns("Temps video").DataBodyRange) Is Nothing Then

        Exit Function

    End If

    If IsEmpty(Cellule.Value) Then Exit Function
    If Not IsNumeric(Cellule.Value) Then Exit Function

    EstTempsVideo = True

End Function

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

    Dim NomExiste As Boolean
    Dim nomTest As Name

    NomExiste = False

    For Each nomTest In ThisWorkbook.Names

        If UCase(Split(nomTest.Name, "!")(UBound(Split(nomTest.Name, "!")))) = _
            UCase(NomListe) Then

            NomExiste = True
            Exit For

        End If

    Next nomTest

    If Not NomExiste Then

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

