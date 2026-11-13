Attribute VB_Name = "modLiensVideo"
Option Explicit

' =========================================================
' LIENS VERS LA VIDEO, POUR LES MATCHS DEJA SAISIS
'
' A importer uniquement dans un fichier de match cree
' avant le 13/11/2026. Les fichiers plus recents portent
' deja ce code dans modJournal et modVideo : l'importer
' en plus y creerait des procedures en double.
'
' Le module de la feuille Journal actions doit recevoir :
'
'   Private Sub Worksheet_Activate()
'       PoserLiensTempsVideo
'   End Sub
'
'   Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
'       SuivreLienTempsVideo Target
'   End Sub
'
' Chaque temps du journal devient un lien vers la cellule
' Joueur de sa ligne. Un clic y amene et positionne la
' video a cet instant. Le moteur doit etre en 0.6.0 ou
' plus : les plus anciens ne connaissent pas seek_to.
' =========================================================

Private Const URL_SAUT_VIDEO As String = _
    "http://127.0.0.1:48652/command/active/"

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

    Set Temps = ColonneTemps(lo).DataBodyRange

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
        ColonneTemps(lo).DataBodyRange) Is Nothing Then

        Exit Function

    End If

    If IsEmpty(Cellule.Value) Then Exit Function
    If Not IsNumeric(Cellule.Value) Then Exit Function

    EstTempsVideo = True

End Function

Private Sub AllerAuTempsVideo(ByVal Secondes As Double)

    ' Le moteur attend un point decimal, quels que soient
    ' les reglages regionaux.
    EnvoyerSautVideo _
        "seek_to/" & Replace(Format(Secondes, "0.0"), ",", ".")

End Sub

' Les premiers fichiers de la saison nomment la colonne
' "temps video", les suivants "Temps video".
Private Function ColonneTemps(ByVal lo As ListObject) As ListColumn

    Dim Colonne As ListColumn

    For Each Colonne In lo.ListColumns

        If LCase(Colonne.Name) = "temps video" Then
            Set ColonneTemps = Colonne
            Exit Function
        End If

    Next Colonne

End Function

Private Function EnvoyerSautVideo( _
    ByVal Commande As String) As Boolean

    Dim url As String

    url = URL_SAUT_VIDEO & Commande

    On Error GoTo GestionErreur

#If Mac Then

    Dim ScriptApple As String
    Dim Resultat As String

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/curl " & _
        "--silent " & _
        "--show-error " & _
        "--fail " & _
        "--max-time 2 " & _
        "--request POST " & _
        url & _
        Chr(34)

    Resultat = MacScript(ScriptApple)

#Else

    Dim Requete As Object

    Set Requete = _
        CreateObject("WinHttp.WinHttpRequest.5.1")

    Requete.Open "POST", url, False

    Requete.SetTimeouts _
        1000, _
        1000, _
        2000, _
        2000

    Requete.Send

    If Requete.Status < 200 Or _
       Requete.Status >= 300 Then

        Err.Raise _
            vbObjectError + 1000, _
            "EnvoyerSautVideo", _
            "R" & ChrW(233) & "ponse HTTP " & Requete.Status

    End If

#End If

    EnvoyerSautVideo = True
    Exit Function

GestionErreur:

    EnvoyerSautVideo = False

    MsgBox _
        "Impossible de piloter la vid" & ChrW(233) & "o." & _
        vbCrLf & vbCrLf & _
        "Clique d'abord sur Connecter lecteur vid" & ChrW(233) & "o." & _
        vbCrLf & _
        "Commande : " & Commande & _
        vbCrLf & vbCrLf & _
        "D" & ChrW(233) & "tail : " & Err.Description, _
        vbExclamation, _
        "VeoVideoControl"

End Function
