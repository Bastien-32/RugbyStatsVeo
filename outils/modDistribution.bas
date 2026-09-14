Attribute VB_Name = "modDistribution"
Option Explicit

' =========================================================
' REINITIALISATION AVANT DISTRIBUTION
'
' Retire du classeur tout ce qui appartient a un club en
' particulier, pour qu'il puisse etre remis a quelqu'un
' d'autre : effectif, equipes de la poule, composition,
' en-tete du match, journal et statistiques.
'
' Ce qui reste est le parametrage commun a tous les clubs :
' libelles d'actions, lieux, phases, resultats, niveaux
' d'equipe, listes du journal, styles et dispositions.
'
' A executer sur une COPIE du classeur, jamais sur celui
' de travail : l'effacement est definitif.
' =========================================================

Private Const NOM_FEUILLE_LISTES As String = "Listes"

Private Const NOM_FEUILLE_COMPO As String = "Compo"

Private Const NOM_FEUILLE_JOURNAL As String = "Journal actions"


Public Sub ReinitialiserPourDistribution()

    Dim Reponse As VbMsgBoxResult

    Dim EtatEvenements As Boolean
    Dim EtatAffichage As Boolean

    Reponse = MsgBox( _
        "Ce classeur va " & ChrW(234) & "tre vid" & ChrW(233) & _
        " de toutes les donn" & ChrW(233) & "es du club :" & _
        vbCrLf & vbCrLf & _
        "- l'effectif ;" & vbCrLf & _
        "- les " & ChrW(233) & "quipes de la poule ;" & vbCrLf & _
        "- la composition et l'en-t" & ChrW(234) & "te du match ;" & vbCrLf & _
        "- le journal d'actions et les statistiques." & _
        vbCrLf & vbCrLf & _
        "L'effacement est d" & ChrW(233) & "finitif." & vbCrLf & _
        "N'utilise cette macro que sur une copie." & _
        vbCrLf & vbCrLf & _
        "Continuer ?", _
        vbExclamation + vbYesNo + vbDefaultButton2, _
        "R" & ChrW(233) & "initialiser pour distribution" _
    )

    If Reponse <> vbYes Then
        Exit Sub
    End If

    EtatEvenements = Application.EnableEvents
    EtatAffichage = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ViderTableau NOM_FEUILLE_LISTES, "LstEffectif"
    ViderTableau NOM_FEUILLE_LISTES, "LstEquipesPoule"
    ViderTableau NOM_FEUILLE_JOURNAL, "JournalActions"

    ViderComposition
    ViderEnteteMatch

    ' Le journal est vide : les tableaux de statistiques
    ' ecrits par le VBA repassent a zero.
    RecalculerStatsMatch

    shSaisieVideo.ActualiserListeJoueursJournal

    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    MsgBox _
        "Classeur r" & ChrW(233) & "initialis" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & _
        "Enregistre-le, puis fabrique le paquet avec " & _
        "outils/preparer_distribution.py", _
        vbInformation, _
        "R" & ChrW(233) & "initialisation termin" & ChrW(233) & "e"

    Exit Sub

GestionErreur:

    Dim DescriptionErreur As String

    DescriptionErreur = Err.Description

    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    MsgBox _
        "La r" & ChrW(233) & "initialisation a " & _
        ChrW(233) & "chou" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & _
        DescriptionErreur & _
        vbCrLf & vbCrLf & _
        "Ferme ce classeur sans enregistrer.", _
        vbCritical, _
        "Erreur"

End Sub


' =========================================================
' Vide un tableau structure en conservant ses en-tetes et
' une ligne vide, pour que les formules qui s'y referent
' continuent de fonctionner.
' =========================================================

Private Sub ViderTableau( _
    ByVal NomFeuille As String, _
    ByVal NomTableau As String _
)

    Dim loTableau As ListObject

    Set loTableau = _
        ThisWorkbook.Worksheets(NomFeuille).ListObjects(NomTableau)

    If loTableau.DataBodyRange Is Nothing Then
        Exit Sub
    End If

    loTableau.DataBodyRange.ClearContents

    If loTableau.ListRows.Count > 1 Then

        loTableau.Resize _
            loTableau.Range.Resize(2, loTableau.ListColumns.Count)

    End If

End Sub


' =========================================================
' La composition perd ses joueurs et leur temps de jeu.
' Les postes et la ligne avant / trois-quarts restent :
' ils decrivent la feuille de match, pas le club.
' =========================================================

Private Sub ViderComposition()

    Dim loCompo As ListObject

    Set loCompo = _
        ThisWorkbook _
            .Worksheets(NOM_FEUILLE_COMPO) _
            .ListObjects("COMPO")

    If loCompo.DataBodyRange Is Nothing Then
        Exit Sub
    End If

    loCompo.ListColumns("Nom du joueur") _
        .DataBodyRange.ClearContents

    loCompo.ListColumns("temps de jeu") _
        .DataBodyRange.ClearContents

End Sub


Private Sub ViderEnteteMatch()

    Dim Champs As Variant
    Dim Champ As Variant

    Champs = Array( _
        "MATCH_SAISON", _
        "MATCH_DATE", _
        "MATCH_ADV", _
        "MATCH_LIEU", _
        "MATCH_PHASE", _
        "MATCH_JOURNEE", _
        "MATCH_RESULTAT", _
        "MATCH_CAT", _
        "MATCH_ID" _
    )

    For Each Champ In Champs

        On Error Resume Next

        ThisWorkbook.Names(CStr(Champ)) _
            .RefersToRange.ClearContents

        On Error GoTo 0

    Next Champ

End Sub
