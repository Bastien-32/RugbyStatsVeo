Attribute VB_Name = "modHeartbeatVideo"
Option Explicit

Private Const URL_HEARTBEAT_VIDEO As String = _
    "http://127.0.0.1:48652/heartbeat"

Private Const INTERVALLE_HEARTBEAT_SECONDES As Long = 1

Private ProchainHeartbeat As Date
Private HeartbeatActif As Boolean


Public Sub DemarrerHeartbeatVideo()

    If HeartbeatActif Then
        Exit Sub
    End If

    HeartbeatActif = True

    EnvoyerHeartbeatVideo

End Sub


Public Sub EnvoyerHeartbeatVideo()

    If FermetureEnCours Then Exit Sub
    
    If Not HeartbeatActif Then
        Exit Sub
    End If

    AppelerHeartbeatVideo

    ' Programme le signal suivant dans une seconde.
    ProchainHeartbeat = _
        Now + TimeSerial(0, 0, INTERVALLE_HEARTBEAT_SECONDES)

    On Error Resume Next

    Application.OnTime _
        EarliestTime:=ProchainHeartbeat, _
        Procedure:=NomProcedureHeartbeat(), _
        Schedule:=True

    On Error GoTo 0

End Sub


Public Sub ArreterHeartbeatVideo()

    If Not HeartbeatActif Then
        Exit Sub
    End If

    HeartbeatActif = False

    On Error Resume Next

    Application.OnTime _
        EarliestTime:=ProchainHeartbeat, _
        Procedure:=NomProcedureHeartbeat(), _
        Schedule:=False

    On Error GoTo 0

End Sub


Private Sub AppelerHeartbeatVideo()

#If Mac Then

    Dim ScriptApple As String
    Dim Resultat As String

    On Error Resume Next

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/curl " & _
        "--silent " & _
        "--fail " & _
        "--max-time 0.8 " & _
        "--request POST " & _
        URL_HEARTBEAT_VIDEO & _
        Chr(34)

    Resultat = MacScript(ScriptApple)

    On Error GoTo 0

#Else

    Dim Requete As Object

    On Error Resume Next

    Set Requete = _
        CreateObject("WinHttp.WinHttpRequest.5.1")

    Requete.Open _
        "POST", _
        URL_HEARTBEAT_VIDEO, _
        False

    Requete.SetTimeouts _
        500, _
        500, _
        800, _
        800

    Requete.Send

    On Error GoTo 0

#End If

End Sub


Private Function NomProcedureHeartbeat() As String

    Dim NomClasseurSecurise As String

    NomClasseurSecurise = _
        Replace(ThisWorkbook.Name, "'", "''")

    NomProcedureHeartbeat = _
        "'" & _
        NomClasseurSecurise & _
        "'!EnvoyerHeartbeatVideo"

End Function

