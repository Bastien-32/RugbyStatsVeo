Attribute VB_Name = "modGlobals"
Option Explicit

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

Public PendingPenaltyActive As Boolean
Public PendingPenaltyBaseAction As String
Public PendingPenaltyTime As Double
Public PendingPenaltyForcedTeam As String
Public PendingPenaltyNeedsPlayer As Boolean

Public PendingMotifPenalite As String
Public PendingGroupeFautifRequired As Boolean

Public ChronoRunning As Boolean
Public ChronoStart As Date
Public ChronoOffset As Double
