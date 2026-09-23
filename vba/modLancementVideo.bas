Attribute VB_Name = "modLancementVideo"
Option Explicit

Private Const URL_TEST_MOTEUR As String = _
    "http://127.0.0.1:48652/openapi.json"

Private Const NOM_MOTEUR_WINDOWS As String = _
    "VeoVideoControl.exe"

Public Sub DemarrerVeoVideoControl()

#If Mac Then

    Dim HeureDebut As Double

    If MoteurVideoDisponible() Then
        Exit Sub
    End If

    If Not LancerApplicationMac() Then

        MsgBox _
            "VeoVideoControl n'a pas pu etre lance.", _
            vbExclamation, _
            "VeoVideoControl"

        Exit Sub

    End If

    HeureDebut = Timer

    Do

        DoEvents

        If MoteurVideoDisponible() Then
            Exit Sub
        End If

        AttendreCourteDuree 0.25

    Loop While DureeEcoulee(HeureDebut) < 15

    MsgBox _
        "VeoVideoControl a ete lance, mais le moteur ne repond pas encore.", _
        vbExclamation, _
        "VeoVideoControl"

#Else

    Dim ShellWindows As Object
    Dim CheminRacine As String
    Dim CheminMoteur As String
    Dim DossierMoteur As String
    Dim HeureDebut As Double

    If MoteurVideoDisponible() Then
        Exit Sub
    End If

    ' =====================================================
    ' PREMIER CAS :
    ' le classeur est dans le dossier principal STATS
    ' =====================================================

    CheminRacine = _
        ThisWorkbook.Path

    DossierMoteur = _
        CheminRacine & _
        "\VeoVideoControl\VeoVideoControlEngine"

    CheminMoteur = _
        DossierMoteur & _
        "\" & NOM_MOTEUR_WINDOWS

    ' =====================================================
    ' SECOND CAS :
    ' le classeur est un fichier de match dans Matchs
    ' =====================================================

    If Dir(CheminMoteur) = "" Then

        CheminRacine = _
            ThisWorkbook.Path & _
            "\.."

        DossierMoteur = _
            CheminRacine & _
            "\VeoVideoControl\VeoVideoControlEngine"

        CheminMoteur = _
            DossierMoteur & _
            "\" & NOM_MOTEUR_WINDOWS

    End If

    ' =====================================================
    ' VERIFICATION
    ' =====================================================

    If Dir(CheminMoteur) = "" Then

        MsgBox _
            "Impossible de trouver le moteur Windows." & _
            vbCrLf & vbCrLf & _
            "Moteur recherche :" & vbCrLf & _
            CheminMoteur & _
            vbCrLf & vbCrLf & _
            "Classeur :" & vbCrLf & _
            ThisWorkbook.FullName, _
            vbExclamation, _
            "VeoVideoControl"

        Exit Sub

    End If

    ' =====================================================
    ' LANCEMENT DE L'EXECUTABLE
    ' =====================================================

    Set ShellWindows = _
        CreateObject("WScript.Shell")

    ShellWindows.CurrentDirectory = _
        DossierMoteur

    ShellWindows.Run _
        Chr(34) & _
        CheminMoteur & _
        Chr(34), _
        0, _
        False

    ' =====================================================
    ' ATTENTE DU DEMARRAGE DU MOTEUR
    ' =====================================================

    HeureDebut = Timer

    Do

        DoEvents

        If MoteurVideoDisponible() Then
            Exit Sub
        End If

        AttendreCourteDuree 0.25

    Loop While DureeEcoulee(HeureDebut) < 15

    MsgBox _
        "VeoVideoControl a ete lance, mais le moteur ne repond pas.", _
        vbExclamation, _
        "VeoVideoControl"

#End If

End Sub

Public Function MoteurVideoDisponible() As Boolean

#If Mac Then

    Dim ScriptApple As String
    Dim Resultat As String
    
    On Error GoTo Indisponible

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/curl " & _
        "--silent " & _
        "--show-error " & _
        "--fail " & _
        "--max-time 1 " & _
        URL_TEST_MOTEUR & _
        Chr(34)

       Resultat = MacScript(ScriptApple)

    MoteurVideoDisponible = _
        (Len(Trim$(Resultat)) > 0)

    Exit Function

#Else

    Dim Requete As Object

    On Error GoTo Indisponible

    Set Requete = _
        CreateObject("WinHttp.WinHttpRequest.5.1")

    Requete.Open _
        "GET", _
        URL_TEST_MOTEUR, _
        False

    Requete.SetTimeouts _
        500, _
        500, _
        1000, _
        1000

    Requete.Send

    MoteurVideoDisponible = _
        (Requete.Status >= 200 And _
         Requete.Status < 300)

    Exit Function

#End If

Indisponible:

    MoteurVideoDisponible = False

End Function

Private Function LancerApplicationMac() As Boolean

    Dim ScriptApple As String
    Dim Resultat As String
    Dim CheminApplication As String

    On Error GoTo GestionErreur

    CheminApplication = _
        ThisWorkbook.Path & _
        "/VeoVideoControl/VeoVideoControl.app"

    If Dir(CheminApplication, vbDirectory) = "" Then

        CheminApplication = _
            ThisWorkbook.Path & _
            "/../VeoVideoControl/VeoVideoControl.app"

    End If

    If Dir(CheminApplication, vbDirectory) = "" Then

        LancerApplicationMac = False
        Exit Function

    End If

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/open " & _
        Chr(39) & _
        CheminApplication & _
        Chr(39) & _
        Chr(34)

    Resultat = MacScript(ScriptApple)

    LancerApplicationMac = True
    Exit Function

GestionErreur:

    LancerApplicationMac = False

End Function


Private Sub AttendreCourteDuree( _
    ByVal NombreSecondes As Double)

    Dim HeureDebut As Double

    HeureDebut = Timer

    Do While DureeEcoulee(HeureDebut) < NombreSecondes
        DoEvents
    Loop

End Sub


Private Function DureeEcoulee( _
    ByVal HeureDebut As Double) As Double

    If Timer >= HeureDebut Then

        DureeEcoulee = Timer - HeureDebut

    Else

        DureeEcoulee = _
            (86400# - HeureDebut) + Timer

    End If

End Function

Public Sub ArreterVeoVideoControl()

#If Mac Then

    Dim ScriptApple As String
    Dim Resultat As String

    On Error Resume Next

    ScriptApple = _
        "do shell script " & _
        Chr(34) & _
        "/usr/bin/open -a 'Arreter VeoVideoControl'" & _
        Chr(34)

    Resultat = MacScript(ScriptApple)

    On Error GoTo 0

#End If

End Sub

Public Sub AttendreUneSeconde()

    AttendreCourteDuree 1

End Sub

Public Function AutreClasseurRugbyOuvert( _
    ByVal ClasseurQuiFerme As Workbook) As Boolean

    Dim wb As Workbook

    For Each wb In Application.Workbooks

        If Not wb Is ClasseurQuiFerme Then

            If EstClasseurStatsRugby(wb) Then

                AutreClasseurRugbyOuvert = True
                Exit Function

            End If

        End If

    Next wb

    AutreClasseurRugbyOuvert = False

End Function

Private Function EstClasseurStatsRugby( _
    ByVal wb As Workbook) As Boolean

    Dim NomMarqueur As Name

    On Error Resume Next

    Set NomMarqueur = wb.Names("STATS_RUGBY")

    On Error GoTo 0

    EstClasseurStatsRugby = Not NomMarqueur Is Nothing

End Function
