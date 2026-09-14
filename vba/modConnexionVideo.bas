Attribute VB_Name = "modConnexionVideo"
Option Explicit

Private Const NOM_TEXTE As String = "ETAT_CONNEXION_VIDEO"
Private Const NOM_PASTILLE As String = "PASTILLE_CONNEXION_VIDEO"
Private Const NOM_BOUTON As String = "BTN_CONNECTER_VIDEO"

Private Const URL_CONNEXION_VIDEO As String = _
    "http://127.0.0.1:48652/connect-video"

Private ConnexionEnCours As Boolean

Public Sub ConnecterVideo()

    Dim ReponseJSON As String
    Dim Lecteur As String

    If ConnexionEnCours Then
        Exit Sub
    End If

    ConnexionEnCours = True

    On Error GoTo GestionErreur

    DesactiverBoutonConnecter
    AfficherEtatRecherche

    DoEvents

    ReponseJSON = AppelerConnexionVideo()

    If JSONContientVrai(ReponseJSON, "connected") Then

        Lecteur = ExtraireTexteJSON( _
            ReponseJSON, _
            "player" _
        )

        Select Case UCase$(Trim$(Lecteur))

            Case "GOOGLE CHROME", "CHROME"
                AfficherEtatChrome

            Case "SAFARI"
                AfficherEtatSafari

            Case "VLC"
                AfficherEtatVLC

            Case Else
                AfficherEtatLecteur Lecteur

        End Select

    Else

        AfficherEtatAucunLecteur

    End If

Fin:

    ActiverBoutonConnecter
    ConnexionEnCours = False
    Exit Sub

GestionErreur:

    AfficherEtatMoteurIndisponible

    Resume Fin

End Sub


Private Function AppelerConnexionVideo() As String

#If Mac Then

    Dim ScriptApple As String

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/curl " & _
        "--silent " & _
        "--show-error " & _
        "--max-time 5 " & _
        "--request POST " & _
        URL_CONNEXION_VIDEO & _
        " 2>&1 || true" & _
        Chr(34)

    AppelerConnexionVideo = MacScript(ScriptApple)

#Else

    Dim Requete As Object

    Set Requete = _
        CreateObject("WinHttp.WinHttpRequest.5.1")

    Requete.Open _
        "POST", _
        URL_CONNEXION_VIDEO, _
        False

    Requete.SetTimeouts _
        1000, _
        1000, _
        5000, _
5000

    Requete.Send

    If Requete.Status < 200 Or _
       Requete.Status >= 300 Then

        Err.Raise _
            vbObjectError + 1500, _
            "AppelerConnexionVideo", _
            "R" & ChrW(233) & "ponse HTTP " & Requete.Status

    End If

    AppelerConnexionVideo = Requete.ResponseText

#End If

End Function


Private Function JSONContientVrai( _
    ByVal JSON As String, _
    ByVal NomPropriete As String) As Boolean

    Dim TexteRecherche As String

    TexteRecherche = _
        Chr(34) & NomPropriete & Chr(34) & _
        ":true"

    JSONContientVrai = _
        (InStr( _
            1, _
            Replace(JSON, " ", ""), _
            TexteRecherche, _
            vbTextCompare _
        ) > 0)

End Function


Private Function ExtraireTexteJSON( _
    ByVal JSON As String, _
    ByVal NomPropriete As String) As String

    Dim PositionPropriete As Long
    Dim PositionDeuxPoints As Long
    Dim PositionGuillemetDebut As Long
    Dim PositionGuillemetFin As Long

    PositionPropriete = InStr( _
        1, _
        JSON, _
        Chr(34) & NomPropriete & Chr(34), _
        vbTextCompare _
    )

    If PositionPropriete = 0 Then
        ExtraireTexteJSON = ""
        Exit Function
    End If

    PositionDeuxPoints = InStr( _
        PositionPropriete, _
        JSON, _
        ":" _
    )

    PositionGuillemetDebut = InStr( _
        PositionDeuxPoints + 1, _
        JSON, _
        Chr(34) _
    )

    PositionGuillemetFin = InStr( _
        PositionGuillemetDebut + 1, _
        JSON, _
        Chr(34) _
    )

    If PositionGuillemetDebut = 0 Or _
       PositionGuillemetFin = 0 Then

        ExtraireTexteJSON = ""
        Exit Function

    End If

    ExtraireTexteJSON = Mid$( _
        JSON, _
        PositionGuillemetDebut + 1, _
        PositionGuillemetFin - _
            PositionGuillemetDebut - 1 _
    )

End Function


Public Sub AfficherEtatRecherche()

    MettreEtatConnexion _
        "Recherche d'un lecteur...", _
        RGB(255, 193, 7)

End Sub

Public Sub InitialiserEtatConnexionVideo()

    ConnexionEnCours = False

    MettreEtatConnexion _
        "Aucun lecteur connect" & ChrW(233) & vbCrLf & _
        "Cliquez sur Connecter lorsque votre vid" & ChrW(233) & "o est pr" & ChrW(234) & "te.", _
        RGB(220, 53, 69)

    ActiverBoutonConnecter

End Sub

Public Sub AfficherEtatAucunLecteur()

    MettreEtatConnexion _
        "Aucun lecteur connect" & ChrW(233) & vbCrLf & _
        "Veuillez lancer la vid" & ChrW(233) & "o dans VLC ou sur le site Veo.", _
        RGB(220, 53, 69)

End Sub


Public Sub AfficherEtatMoteurIndisponible()

    MettreEtatConnexion _
        "VeoVideoControl indisponible" & vbCrLf & _
        "V" & ChrW(233) & "rifiez que le moteur est lanc" & ChrW(233) & ".", _
        RGB(220, 53, 69)

End Sub


Public Sub AfficherEtatChrome()

    MettreEtatConnexion _
        "Google Chrome connect" & ChrW(233), _
        RGB(40, 167, 69)

End Sub


Public Sub AfficherEtatSafari()

    MettreEtatConnexion _
        "Safari connect" & ChrW(233), _
        RGB(40, 167, 69)

End Sub


Public Sub AfficherEtatVLC()

    MettreEtatConnexion _
        "VLC connect" & ChrW(233), _
        RGB(40, 167, 69)

End Sub


Private Sub AfficherEtatLecteur( _
    ByVal NomLecteur As String)

    If Trim$(NomLecteur) = "" Then

        AfficherEtatAucunLecteur

    Else

        MettreEtatConnexion _
            NomLecteur & " connect" & ChrW(233), _
            RGB(40, 167, 69)

    End If

End Sub


Private Sub MettreEtatConnexion( _
    ByVal Texte As String, _
    ByVal Couleur As Long)

    Dim ws As Worksheet

    Set ws = shSaisieVideo

    With ws.Shapes(NOM_TEXTE) _
        .TextFrame2.TextRange

        .Text = Texte

    End With

    With ws.Shapes(NOM_PASTILLE)

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = Couleur

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = Couleur

    End With

End Sub


Private Sub DesactiverBoutonConnecter()

    Dim Bouton As Shape

    Set Bouton = shSaisieVideo.Shapes(NOM_BOUTON)

    With Bouton

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = _
            RGB(210, 210, 210)

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = _
            RGB(180, 180, 180)

        .TextFrame2.TextRange.Text = _
            "Recherche..."

        .TextFrame2.TextRange.Font.Fill _
            .ForeColor.RGB = _
            RGB(120, 120, 120)

    End With

End Sub


Private Sub ActiverBoutonConnecter()

    Dim Bouton As Shape

    Set Bouton = shSaisieVideo.Shapes(NOM_BOUTON)

    With Bouton

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = _
            RGB(47, 126, 32)

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = _
            RGB(31, 90, 21)

        .TextFrame2.TextRange.Text = _
            "Connecter lecteur vid" & ChrW(233) & "o"

        .TextFrame2.TextRange.Font.Fill _
            .ForeColor.RGB = _
            RGB(255, 255, 255)

    End With

End Sub

