Attribute VB_Name = "modGlobals"
Option Explicit

' Met en sourdine les messages de fin des procedures de
' construction, que la reinitialisation enchaine.
Public ModeSilencieux As Boolean

Public CurrentTeam As String
Public CurrentHalf As String

Public PendingAction As String
Public WaitingForPlayer As Boolean
Public FermetureEnCours As Boolean

Public PendingActionTime As Double
Public PendingActionTimeAvailable As Boolean
Public PendingActionChangesPossession As Boolean
Public PendingPlayerCount As Long
Public PendingAllowMultiplePlayers As Boolean
Public PendingForcedTeam As String

Public PendingMotifPenalite As String
Public PendingGroupeFautifRequired As Boolean

Public ChronoRunning As Boolean
Public ChronoStart As Date
Public ChronoOffset As Double
