Attribute VB_Name = "modVideo"
Option Explicit

Private Const URL_COMMANDES_VIDEO As String = _
    "http://127.0.0.1:48652/command/active/"
    
Private Const URL_ETAT_VIDEO As String = _
    "http://127.0.0.1:48652/video/state/active"


Public Sub PlayPauseChronoVideo()

    EnvoyerCommandeVideo "playpause"

End Sub


Public Sub Recul5SecondesVideo()

    EnvoyerCommandeVideo "seek_minus_5"

End Sub


Public Sub Avance5SecondesVideo()

    EnvoyerCommandeVideo "seek_plus_5"

End Sub


Public Sub ResetChronoVideo()

    EnvoyerCommandeVideo "reset"

End Sub

Public Sub ActionPrecedenteVideo()

    EnvoyerCommandeVideo "previous_action"

End Sub


Public Sub ActionSuivanteVideo()

    EnvoyerCommandeVideo "next_action"

End Sub

Private Function EnvoyerCommandeVideo( _
    ByVal Commande As String) As Boolean

    Dim url As String

    url = URL_COMMANDES_VIDEO & Commande

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
            "EnvoyerCommandeVideo", _
            "R" & ChrW(233) & "ponse HTTP " & Requete.Status

    End If

#End If

    EnvoyerCommandeVideo = True
    Exit Function

GestionErreur:

    EnvoyerCommandeVideo = False

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

Public Sub ActiverRaccourcisVideo()

    Application.OnKey " ", "PlayPauseChronoVideo"
    Application.OnKey "{LEFT}", "Recul5SecondesVideo"
    Application.OnKey "{RIGHT}", "Avance5SecondesVideo"
    Application.OnKey "{UP}", "ActionPrecedenteVideo"
    Application.OnKey "{DOWN}", "ActionSuivanteVideo"
    Application.OnKey "r", "BasculerRewindVideo"
    Application.OnKey "R", "BasculerRewindVideo"

    
End Sub


Public Sub DesactiverRaccourcisVideo()

    Application.OnKey " "
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "r"
    Application.OnKey "R"

End Sub

Public Function GetTimeVideo() As Double

    Dim url As String
    Dim JSON As String

    url = URL_ETAT_VIDEO

    On Error GoTo GestionErreur

#If Mac Then

    Dim ScriptApple As String

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/curl " & _
        "--silent " & _
        "--show-error " & _
        "--fail " & _
        "--max-time 2 " & _
        url & _
        Chr(34)

    JSON = MacScript(ScriptApple)

#Else

    Dim Requete As Object

    Set Requete = CreateObject("WinHttp.WinHttpRequest.5.1")

    Requete.Open "GET", url, False

    Requete.SetTimeouts _
        1000, _
        1000, _
        2000, _
        2000

    Requete.Send

    If Requete.Status < 200 Or _
       Requete.Status >= 300 Then

        Err.Raise _
            vbObjectError + 1100, _
            "GetTimeVideo", _
            "R" & ChrW(233) & "ponse HTTP " & Requete.Status

    End If

    JSON = Requete.ResponseText

#End If

    GetTimeVideo = ExtraireNombreJSON( _
        JSON, _
        "currentTime" _
    )

    Exit Function

GestionErreur:

    GetTimeVideo = -1

End Function


Private Function ExtraireNombreJSON( _
    ByVal JSON As String, _
    ByVal NomPropriete As String) As Double

    Dim PositionPropriete As Long
    Dim PositionDeuxPoints As Long
    Dim positionFin As Long
    Dim texteNombre As String
    Dim Caractere As String

    PositionPropriete = InStr( _
        1, _
        JSON, _
        Chr(34) & NomPropriete & Chr(34), _
        vbTextCompare _
    )

    If PositionPropriete = 0 Then
        Err.Raise _
            vbObjectError + 1101, _
            "ExtraireNombreJSON", _
            "Propri" & ChrW(233) & "t" & ChrW(233) & " JSON introuvable : " & NomPropriete
    End If

    PositionDeuxPoints = InStr( _
        PositionPropriete, _
        JSON, _
        ":" _
    )

    If PositionDeuxPoints = 0 Then
        Err.Raise _
            vbObjectError + 1102, _
            "ExtraireNombreJSON", _
            "Valeur JSON invalide."
    End If

    positionFin = PositionDeuxPoints + 1

    Do While positionFin <= Len(JSON)

        Caractere = Mid$(JSON, positionFin, 1)

        If InStr("0123456789.-", Caractere) = 0 Then

            If Len(texteNombre) > 0 Then
                Exit Do
            End If

        Else

            texteNombre = texteNombre & Caractere

        End If

        positionFin = positionFin + 1

    Loop

    If texteNombre = "" Then
        Err.Raise _
            vbObjectError + 1103, _
            "ExtraireNombreJSON", _
            "Nombre JSON introuvable."
    End If

    ' Le moteur renvoie toujours le separateur decimal avec un point.
    ' Val permet de le convertir independamment des reglages francais.
    ExtraireNombreJSON = Val(texteNombre)

End Function

Public Sub BasculerRewindVideo()

    EnvoyerCommandeVideo "rewind_toggle"

End Sub

