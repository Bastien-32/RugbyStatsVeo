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
' La macro commence par enregistrer le classeur sous un
' autre nom : c'est cette copie qu'elle vide, et le
' fichier de travail reste sur le disque avec toutes ses
' donnees. Il n'y a donc plus a penser a dupliquer avant.
' =========================================================

Private Const NOM_FEUILLE_LISTES As String = "Listes"

Private Const NOM_FEUILLE_COMPO As String = "Compo"

Private Const NOM_FEUILLE_JOURNAL As String = "Journal actions"

Private Const NOM_FEUILLE_TM As String = "Touches Melees"

' La composition occupe les colonnes E a H : les colonnes
' de remplacement commencent juste apres.
Private Const PREMIERE_COL_REMPLACEMENT As Long = 9

' Nomme l'etape en cours pour que le message d'erreur dise
' ou la reinitialisation s'est arretee.
Private EtapeDistribution As String


Public Sub ReinitialiserPourDistribution()

    Dim Reponse As VbMsgBoxResult
    Dim Destination As String

    Dim EtatEvenements As Boolean
    Dim EtatAffichage As Boolean

    ' Le fichier a produire est choisi avant tout : la
    ' fenetre d'enregistrement vaut aussi autorisation
    ' d'ecrire la ou macOS la demande.
    Destination = ChoisirFichierADistribuer()

    If Destination = "" Then
        Exit Sub
    End If

    Reponse = MsgBox( _
        "La copie suivante va " & ChrW(234) & "tre cr" & _
        ChrW(233) & ChrW(233) & "e, puis vid" & ChrW(233) & _
        "e de toutes les donn" & ChrW(233) & "es du club :" & _
        vbCrLf & vbCrLf & Destination & vbCrLf & vbCrLf & _
        "Ce qui sera efface de la copie :" & _
        vbCrLf & vbCrLf & _
        "- l'effectif ;" & vbCrLf & _
        "- les " & ChrW(233) & "quipes de la poule ;" & vbCrLf & _
        "- la composition et l'en-t" & ChrW(234) & "te du match ;" & vbCrLf & _
        "- le journal d'actions et les statistiques ;" & vbCrLf & _
        "- le detail des touches et des melees ;" & vbCrLf & _
        "- les remplacements et le temps de jeu." & _
        vbCrLf & vbCrLf & _
        "Ton classeur de travail reste intact sur le " & _
        "disque : c'est la copie qui est vid" & ChrW(233) & "e." & _
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

    ' A partir d'ici, ThisWorkbook designe la copie : le
    ' fichier de travail est ferme, intact, sur le disque.
    EtapeDistribution = "enregistrement de la copie"

    ThisWorkbook.SaveAs _
        Filename:=Destination, _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled

    EtapeDistribution = "effectif"
    ViderTableau NOM_FEUILLE_LISTES, "LstEffectif"

    EtapeDistribution = "equipes de la poule"
    ViderTableau NOM_FEUILLE_LISTES, "LstEquipesPoule"

    EtapeDistribution = "journal d'actions"
    ViderTableau NOM_FEUILLE_JOURNAL, "JournalActions"

    EtapeDistribution = "detail des touches"
    ViderTableau NOM_FEUILLE_TM, "DetailTouches"

    EtapeDistribution = "detail des melees"
    ViderTableau NOM_FEUILLE_TM, "DetailMelees"

    EtapeDistribution = "composition"
    ViderComposition

    EtapeDistribution = "colonnes de remplacement"
    ViderColonnesRemplacement

    EtapeDistribution = "en-tete du match"
    ViderEnteteMatch

    ' Le journal est vide : les tableaux de statistiques
    ' ecrits par le VBA repassent a zero.
    EtapeDistribution = "statistiques du match"
    RecalculerStatsMatch

    ' Les deux tableaux sont vides : leurs recapitulatifs
    ' aussi.
    EtapeDistribution = "recapitulatifs touches et melees"
    ConstruireRecapitulatifs

    ' Les trois listes de joueurs se reconstruisent depuis
    ' une composition desormais vide : le tableau du
    ' journal, la colonne qui alimente sa validation, et
    ' les boutons de la palette.
    EtapeDistribution = "liste des joueurs du journal"
    shSaisieVideo.ActualiserListeJoueursJournal

    EtapeDistribution = "validation des joueurs du journal"
    ActualiserJoueursJournal

    EtapeDistribution = "palette des joueurs"
    ActualiserPaletteJoueurs

    ' Les popups gardent la derniere saisie affichee, noms
    ' de joueurs compris : on les redessine.
    EtapeDistribution = "popups"
    ReconstruirePopups

    EtapeDistribution = ""

    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    ThisWorkbook.Save

    MsgBox _
        "Copie " & ChrW(224) & " distribuer pr" & ChrW(234) & _
        "te :" & vbCrLf & vbCrLf & Destination & _
        vbCrLf & vbCrLf & _
        "ATTENTION : le classeur ouvert est maintenant " & _
        "cette copie vierge, pas ton classeur de travail." & _
        vbCrLf & _
        "Ferme-le et rouvre " & ChrW(171) & " Createur de " & _
        "match.xlsm " & ChrW(187) & " pour reprendre ta saisie.", _
        vbInformation, _
        "R" & ChrW(233) & "initialisation termin" & ChrW(233) & "e"

    Exit Sub

GestionErreur:

    Dim DescriptionErreur As String
    Dim NumeroErreur As Long

    DescriptionErreur = Err.Description
    NumeroErreur = Err.Number

    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    MsgBox _
        "La r" & ChrW(233) & "initialisation a " & _
        ChrW(233) & "chou" & ChrW(233) & " pendant : " & _
        EtapeDistribution & "." & _
        vbCrLf & vbCrLf & _
        "Erreur " & NumeroErreur & " : " & _
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
'
' Le tableau retreci laisse derriere lui des lignes qui ont
' garde son habillage : bordures et bandes de couleur sur
' des cellules desormais hors du tableau. On efface donc ce
' format, puis on repose un fond blanc.
'
' Le fond blanc n'est pas une coquetterie : sans remplissage,
' le quadrillage d'Excel reapparait, alors que le reste de la
' feuille le masque avec un blanc uni. Une zone simplement
' effacee se verrait donc comme un rectangle quadrille.
' =========================================================

' =========================================================
' Le fichier a produire.
'
' Le dossier dist est impose, avec un nom propose que l'on
' peut changer. Les fenetres d'enregistrement d'Excel pour
' Mac rendent des chemins inexploitables - extension
' doublee, caractere parasite - la ou une simple saisie
' donne toujours ce qu'on attend.
' =========================================================
Private Function ChoisirFichierADistribuer() As String

    Dim Dossier As String
    Dim Nom As String
    Dim Chemin As String

    Dossier = ThisWorkbook.Path & _
        Application.PathSeparator & "dist"

    If Dir(Dossier, vbDirectory) = "" Then

        On Error Resume Next
        MkDir Dossier
        On Error GoTo 0

    End If

    Nom = InputBox( _
        "Nom du fichier " & ChrW(224) & " produire, dans le " & _
        "dossier dist :" & vbCrLf & vbCrLf & Dossier, _
        "Copie " & ChrW(224) & " distribuer", _
        "Createur de match - vierge.xlsm")

    Nom = Trim$(Nom)

    ' Annulation, ou champ vide.
    If Len(Nom) = 0 Then
        Exit Function
    End If

    If LCase$(Right$(Nom, 5)) <> ".xlsm" Then
        Nom = Nom & ".xlsm"
    End If

    Chemin = Dossier & Application.PathSeparator & Nom

    ' Un fichier deja la serait ecrase sans bruit par
    ' SaveAs : on demande.
    If Dir(Chemin) <> "" Then

        If MsgBox( _
            Nom & " existe d" & ChrW(233) & "j" & ChrW(224) & _
            " dans dist." & vbCrLf & vbCrLf & "Le remplacer ?", _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            "Fichier existant") <> vbYes Then

            Exit Function

        End If

        ' DisplayAlerts est deja a False plus loin, mais
        ' SaveAs refuserait d'ecraser un fichier ouvert
        ' ailleurs : on le supprime d'abord.
        On Error Resume Next
        Kill Chemin
        On Error GoTo 0

    End If

    ChoisirFichierADistribuer = Chemin

End Function


Private Sub ViderTableau( _
    ByVal NomFeuille As String, _
    ByVal NomTableau As String _
)

    Dim ws As Worksheet
    Dim loTableau As ListObject

    Dim ColonneDebut As Long
    Dim NbColonnes As Long

    Dim DerniereLigneAvant As Long
    Dim PremiereLigneLiberee As Long

    Set ws = ThisWorkbook.Worksheets(NomFeuille)

    Set loTableau = ws.ListObjects(NomTableau)

    If loTableau.DataBodyRange Is Nothing Then
        Exit Sub
    End If

    ColonneDebut = loTableau.Range.Column

    NbColonnes = loTableau.ListColumns.Count

    DerniereLigneAvant = _
        loTableau.Range.Row + loTableau.Range.Rows.Count - 1

    loTableau.DataBodyRange.ClearContents

    If loTableau.ListRows.Count > 1 Then

        loTableau.Resize _
            loTableau.Range.Resize(2, NbColonnes)

    End If

    PremiereLigneLiberee = _
        loTableau.Range.Row + loTableau.Range.Rows.Count

    If PremiereLigneLiberee <= DerniereLigneAvant Then

        With ws.Range( _
            ws.Cells(PremiereLigneLiberee, ColonneDebut), _
            ws.Cells( _
                DerniereLigneAvant, _
                ColonneDebut + NbColonnes - 1 _
            ) _
        )

            .ClearFormats

            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(255, 255, 255)

        End With

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
        "MATCH_NOUS", _
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


' =========================================================
' Les remplacements ecrivent une colonne par validation a
' droite de la composition. Elles sont effacees en entier,
' habillage compris, et retrouvent leur largeur d'origine.
'
' Contrairement aux tableaux vides, il n'y a pas de fond
' blanc a reposer : cette partie de la feuille n'en a
' jamais eu.
' =========================================================

Private Sub ViderColonnesRemplacement()

    Dim ws As Worksheet
    Dim DerniereColonne As Long

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_COMPO)

    DerniereColonne = ws.UsedRange.Column + _
        ws.UsedRange.Columns.Count - 1

    If DerniereColonne < PREMIERE_COL_REMPLACEMENT Then
        Exit Sub
    End If

    With ws.Range( _
        ws.Columns(PREMIERE_COL_REMPLACEMENT), _
        ws.Columns(DerniereColonne) _
    )

        .Clear
        .ColumnWidth = ws.StandardWidth

    End With

End Sub


' =========================================================
' Les cinq popups sont redessinees a partir de leur code,
' puis remasquees : leur construction les laisse visibles
' pour que l'on en verifie le rendu.
' =========================================================

Private Sub ReconstruirePopups()

    ModeSilencieux = True

    On Error GoTo Sortie

    ConstruirePopupsTouchesMelees
    ConstruirePopupPenalite
    ConstruirePopupRemplacement
    ConstruirePopupRetourCarton

Sortie:

    ModeSilencieux = False

    MasquerPopup "Popup touche"
    MasquerPopup "Popup melee"
    MasquerPopup "Popup penalite"
    MasquerPopup "Popup remplacement"
    MasquerPopup FEUILLE_POPUP_RETOUR

End Sub


Private Sub MasquerPopup(ByVal NomFeuille As String)

    On Error Resume Next

    ThisWorkbook.Worksheets(NomFeuille).Visible = _
        xlSheetHidden

    On Error GoTo 0

End Sub
