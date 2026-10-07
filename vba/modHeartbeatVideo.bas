Attribute VB_Name = "modHeartbeatVideo"
Option Explicit

Private Const URL_HEARTBEAT_VIDEO As String = _
    "http://127.0.0.1:48652/heartbeat"

Private Const INTERVALLE_HEARTBEAT_SECONDES As Long = 1

Private ProchainHeartbeat As Date
Private HeartbeatActif As Boolean
Private HeartbeatPlanifie As Boolean


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

    ' Un rendez-vous rattrape par VerifierHeartbeatVideo
    ' peut se presenter alors qu'un autre bat deja : sans
    ' ce garde-fou, deux chaines tourneraient en parallele.
    If HeartbeatPlanifie Then
        If ProchainHeartbeat > Now Then Exit Sub
    End If

    HeartbeatPlanifie = False

    ' Le signal suivant est programme avant l'appel : si
    ' celui-ci echoue, la chaine survit quand meme.
    PlanifierProchainHeartbeat

    AppelerHeartbeatVideo

End Sub


' Excel abandonne un rendez-vous OnTime quand il est
' occupe a l'heure dite : saisir une observation dans une
' cellule suffit. Le signal s'arrete alors sans bruit, et
' le moteur video se ferme au bout de trente secondes
' sans nouvelles. Ce controle relance la chaine au
' premier clic qui suit la saisie.
Public Sub VerifierHeartbeatVideo()

    If FermetureEnCours Then Exit Sub

    If Not HeartbeatActif Then Exit Sub

    If HeartbeatPlanifie Then
        If Now <= ProchainHeartbeat + TimeSerial(0, 0, 5) Then
            Exit Sub
        End If
    End If

    PlanifierProchainHeartbeat
    AppelerHeartbeatVideo

End Sub


Private Sub PlanifierProchainHeartbeat()

    On Error Resume Next

    ProchainHeartbeat = _
        Now + TimeSerial(0, 0, INTERVALLE_HEARTBEAT_SECONDES)

    HeartbeatPlanifie = True

    Err.Clear

    Application.OnTime _
        EarliestTime:=ProchainHeartbeat, _
        Procedure:=NomProcedureHeartbeat(), _
        Schedule:=True

    ' Rendez-vous refuse : le marquer comme absent pour que
    ' le prochain controle retente.
    If Err.Number <> 0 Then HeartbeatPlanifie = False

    On Error GoTo 0

End Sub


Public Sub ArreterHeartbeatVideo()

    If Not HeartbeatActif Then
        Exit Sub
    End If

    HeartbeatActif = False

    On Error Resume Next

    If HeartbeatPlanifie Then

        Application.OnTime _
            EarliestTime:=ProchainHeartbeat, _
            Procedure:=NomProcedureHeartbeat(), _
            Schedule:=False

    End If

    HeartbeatPlanifie = False

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

