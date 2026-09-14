Attribute VB_Name = "modRecapitulatif"
Option Explicit

' =========================================================
' RECAPITULATIF D'UN MATCH SAISI EN DEUX FICHIERS
'
' A importer dans une copie du fichier de premiere mi-temps,
' renommee 26-27_AMI_RC-Trie_Premiere_DOM.xlsm, puis a
' executer depuis ce classeur : la macro y ajoute le journal
' et la composition de la seconde mi-temps, etend le tableau
' des stats par joueur, et ne laisse visibles que
' Journal actions et Stats match.
'
' Usage unique : ce module n'a pas sa place dans le modele.
'
' Ordre des operations impose par le code existant :
' RecalculerStatsMatch ecrit le tableau des sequences de
' possession a des adresses fixes (B92:D97). Le recalcul est
' donc lance AVANT d'inserer les lignes du tableau des
' joueurs, sinon les valeurs partiraient au mauvais endroit.
' Pour la meme raison, le bouton de recalcul est masque.
' =========================================================

Private Const FICHIER_MT2 As String = _
    "26-27_MT2_RC-Trie_Premiere_DOM.xlsm"

' Stats match comporte deux tableaux indexes par joueur,
' tous deux dimensionnes pour 21 noms et alimentes par une
' formule SORT(FILTER(COMPO...)) :
'
'   - stats generales, lignes 8 a 28, colonnes C a V ;
'   - touches, lignes 62 a 82, colonnes C a K.
'
' C a V couvre les colonnes ajoutees pour le franchissement
' et les penalites decalees. Sur un classeur ou ces colonnes
' n'existent pas encore, les dernieres sont vides : la copie
' reste sans effet.

Private Const STATS_PREMIERE_LIGNE As Long = 8

Private Const STATS_DERNIERE_LIGNE As Long = 28

Private Const STATS_DERNIERE_COLONNE As String = "V"

Private Const TOUCHES_PREMIERE_LIGNE As Long = 62

Private Const TOUCHES_DERNIERE_LIGNE As Long = 82

Private Const TOUCHES_DERNIERE_COLONNE As String = "K"


Public Sub CreerRecapitulatifMatch()

    Dim CheminMT2 As String

    Dim wbMT2 As Workbook

    Dim NbJoueurs As Long
    Dim NbLignesAjoutees As Long

    Dim EtatEvenements As Boolean
    Dim EtatAffichage As Boolean
    Dim EtatAlertes As Boolean

#If Mac Then

    Dim FichiersDemandes As Variant

#End If

    CheminMT2 = _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        FICHIER_MT2

    ' Une seconde execution doublerait le journal.
    If JournalContientMiTemps("MT2") Then

        MsgBox _
            "Le journal contient d" & ChrW(233) & "j" & ChrW(224) & _
            " des lignes de seconde mi-temps." & _
            vbCrLf & vbCrLf & _
            "La fusion a donc d" & ChrW(233) & "j" & ChrW(224) & _
            " " & ChrW(233) & "t" & ChrW(233) & " faite.", _
            vbExclamation, _
            "Fusion d" & ChrW(233) & "j" & ChrW(224) & " effectu" & ChrW(233) & "e"

        Exit Sub

    End If

#If Mac Then

    FichiersDemandes = Array(CheminMT2)

    If Not GrantAccessToMultipleFiles(FichiersDemandes) Then

        MsgBox _
            "L'autorisation d'acc" & ChrW(232) & _
            "s au fichier de seconde mi-temps n'a pas " & _
            ChrW(233) & "t" & ChrW(233) & " accord" & ChrW(233) & "e.", _
            vbExclamation, _
            "Autorisation refus" & ChrW(233) & "e"

        Exit Sub

    End If

#End If

    EtatEvenements = Application.EnableEvents
    EtatAffichage = Application.ScreenUpdating
    EtatAlertes = Application.DisplayAlerts

    On Error GoTo GestionErreur

    ' EnableEvents = False empeche le Workbook_Open du
    ' fichier de seconde mi-temps de relancer la video.
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbMT2 = Workbooks.Open(CheminMT2)

    NbLignesAjoutees = FusionnerJournal(wbMT2)

    NbJoueurs = FusionnerCompositions(wbMT2)

    wbMT2.Close SaveChanges:=False

    Set wbMT2 = Nothing

    ' Temps de jeu effectif et sequences de possession,
    ' pendant que les adresses fixes sont encore valides.
    RecalculerStatsMatch

    ' Du bas vers le haut : inserer des lignes dans le
    ' tableau du haut decalerait celui des touches.
    EtendreTableauJoueurs _
        TOUCHES_PREMIERE_LIGNE, _
        TOUCHES_DERNIERE_LIGNE, _
        TOUCHES_DERNIERE_COLONNE, _
        NbJoueurs

    EtendreTableauJoueurs _
        STATS_PREMIERE_LIGNE, _
        STATS_DERNIERE_LIGNE, _
        STATS_DERNIERE_COLONNE, _
        NbJoueurs

    MasquerBoutonRecalcul

    MasquerFeuillesInutiles

    Application.CalculateFull

    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    MsgBox _
        "Fusion termin" & ChrW(233) & "e." & _
        vbCrLf & vbCrLf & _
        NbLignesAjoutees & " lignes de journal ajout" & ChrW(233) & "es." & _
        vbCrLf & _
        NbJoueurs & " joueurs dans le tableau des stats." & _
        vbCrLf & vbCrLf & _
        "V" & ChrW(233) & "rifie le r" & ChrW(233) & "sultat, puis enregistre.", _
        vbInformation, _
        "R" & ChrW(233) & "capitulatif"

    Exit Sub

GestionErreur:

    Dim DescriptionErreur As String

    DescriptionErreur = Err.Description

    On Error Resume Next

    If Not wbMT2 Is Nothing Then
        wbMT2.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    On Error GoTo 0

    MsgBox _
        "La fusion a " & ChrW(233) & "chou" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & _
        DescriptionErreur & _
        vbCrLf & vbCrLf & _
        "Ferme ce classeur sans enregistrer, " & _
        "puis repars d'une copie neuve.", _
        vbCritical, _
        "Erreur"

End Sub


Private Function JournalContientMiTemps( _
    ByVal MiTemps As String _
) As Boolean

    Dim loJournal As ListObject
    Dim ColMiTemps As Long
    Dim i As Long

    Set loJournal = _
        ThisWorkbook _
            .Worksheets("Journal actions") _
            .ListObjects("JournalActions")

    If loJournal.DataBodyRange Is Nothing Then
        Exit Function
    End If

    ColMiTemps = loJournal.ListColumns("Mi-temps").Index

    For i = 1 To loJournal.ListRows.Count

        If StrComp( _
            Trim(CStr( _
                loJournal.DataBodyRange.Cells(i, ColMiTemps).Value _
            )), _
            MiTemps, _
            vbTextCompare _
        ) = 0 Then

            JournalContientMiTemps = True
            Exit Function

        End If

    Next i

End Function


' =========================================================
' JOURNAL : les lignes de la seconde mi-temps sont
' ajoutees a la suite de celles de la premiere.
' =========================================================

Private Function FusionnerJournal( _
    ByVal wbSource As Workbook _
) As Long

    Dim loSource As ListObject
    Dim loCible As ListObject

    Dim Donnees As Variant
    Dim Destination As Range

    Dim NbLignesSource As Long
    Dim NbLignesCible As Long
    Dim NbColonnes As Long

    Set loSource = _
        wbSource _
            .Worksheets("Journal actions") _
            .ListObjects("JournalActions")

    Set loCible = _
        ThisWorkbook _
            .Worksheets("Journal actions") _
            .ListObjects("JournalActions")

    If loSource.DataBodyRange Is Nothing Then
        Exit Function
    End If

    Donnees = loSource.DataBodyRange.Value

    NbLignesSource = UBound(Donnees, 1)
    NbColonnes = UBound(Donnees, 2)

    NbLignesCible = loCible.ListRows.Count

    ' Cells(1, 1) du ListObject est l'en-tete :
    ' la premiere ligne libre est donc en NbLignesCible + 2.
    Set Destination = _
        loCible.Range.Cells(NbLignesCible + 2, 1) _
            .Resize(NbLignesSource, NbColonnes)

    Destination.Value = Donnees

    loCible.Resize _
        loCible.Range.Resize( _
            NbLignesCible + NbLignesSource + 1, _
            loCible.ListColumns.Count _
        )

    FusionnerJournal = NbLignesSource

End Function


' =========================================================
' COMPOSITION : les joueurs de la seconde mi-temps absents
' de la premiere sont ajoutes au tableau COMPO.
'
' Renvoie le nombre total de joueurs nommes, qui determine
' la hauteur du tableau des stats.
' =========================================================

Private Function FusionnerCompositions( _
    ByVal wbSource As Workbook _
) As Long

    Dim loSource As ListObject
    Dim loCible As ListObject

    Dim NomsPresents As Collection
    Dim NouvelleLigne As ListRow

    Dim ColPoste As Long
    Dim ColNom As Long
    Dim ColGroupe As Long

    Dim i As Long
    Dim Nom As String
    Dim Total As Long

    Set loSource = _
        wbSource.Worksheets("Compo").ListObjects("COMPO")

    Set loCible = _
        ThisWorkbook.Worksheets("Compo").ListObjects("COMPO")

    ColPoste = loCible.ListColumns("Poste").Index
    ColNom = loCible.ListColumns("Nom du joueur").Index
    ColGroupe = loCible.ListColumns("Avant ou 3/4").Index

    Set NomsPresents = New Collection

    If Not loCible.DataBodyRange Is Nothing Then

        For i = 1 To loCible.ListRows.Count

            Nom = Trim(CStr( _
                loCible.DataBodyRange.Cells(i, ColNom).Value _
            ))

            If Nom <> "" Then

                On Error Resume Next
                NomsPresents.Add Nom, Nom
                On Error GoTo 0

            End If

        Next i

    End If

    If Not loSource.DataBodyRange Is Nothing Then

        For i = 1 To loSource.ListRows.Count

            Nom = Trim(CStr( _
                loSource.DataBodyRange.Cells(i, ColNom).Value _
            ))

            If Nom <> "" Then

                If Not NomDejaPresent(NomsPresents, Nom) Then

                    Set NouvelleLigne = loCible.ListRows.Add

                    NouvelleLigne.Range.Cells(1, ColPoste).Value = _
                        loSource.DataBodyRange.Cells(i, ColPoste).Value

                    NouvelleLigne.Range.Cells(1, ColNom).Value = Nom

                    NouvelleLigne.Range.Cells(1, ColGroupe).Value = _
                        loSource.DataBodyRange.Cells(i, ColGroupe).Value

                    NomsPresents.Add Nom, Nom

                End If

            End If

        Next i

    End If

    Total = 0

    For i = 1 To loCible.ListRows.Count

        Nom = Trim(CStr( _
            loCible.DataBodyRange.Cells(i, ColNom).Value _
        ))

        If Nom <> "" Then
            Total = Total + 1
        End If

    Next i

    FusionnerCompositions = Total

End Function


Private Function NomDejaPresent( _
    ByVal NomsPresents As Collection, _
    ByVal Nom As String _
) As Boolean

    Dim Valeur As String

    On Error Resume Next

    Valeur = NomsPresents.Item(Nom)

    NomDejaPresent = (Err.Number = 0)

    Err.Clear

    On Error GoTo 0

End Function


' =========================================================
' STATS MATCH : un tableau de joueurs est dimensionne pour
' 21 noms. On insere ce qui manque a l'interieur du tableau,
' ce qui decale automatiquement les sommes, les blocs situes
' en dessous et le graphique.
'
' La colonne B n'est pas recopiee : elle est remplie par la
' formule SORT(FILTER(COMPO...)) en tete de tableau, qui
' deborde toute seule des que la place existe.
' =========================================================

Private Sub EtendreTableauJoueurs( _
    ByVal PremiereLigne As Long, _
    ByVal DerniereLigne As Long, _
    ByVal DerniereColonne As String, _
    ByVal NbJoueurs As Long _
)

    Dim ws As Worksheet

    Dim LignesAInserer As Long
    Dim PremiereLigneAjoutee As Long
    Dim DerniereLigneAjoutee As Long

    Set ws = ThisWorkbook.Worksheets("Stats match")

    LignesAInserer = _
        NbJoueurs - _
        (DerniereLigne - PremiereLigne + 1)

    If LignesAInserer <= 0 Then
        Exit Sub
    End If

    PremiereLigneAjoutee = DerniereLigne + 1

    DerniereLigneAjoutee = DerniereLigne + LignesAInserer

    ws.Rows( _
        PremiereLigneAjoutee & ":" & DerniereLigneAjoutee _
    ).Insert Shift:=xlDown

    ' Mise en forme de la colonne des noms.
    ws.Range("B" & DerniereLigne).Copy

    ws.Range( _
        "B" & PremiereLigneAjoutee & _
        ":B" & DerniereLigneAjoutee _
    ).PasteSpecial Paste:=xlPasteFormats

    ' Formules de comptage.
    ws.Range( _
        "C" & DerniereLigne & _
        ":" & DerniereColonne & DerniereLigne _
    ).Copy

    ws.Range( _
        "C" & PremiereLigneAjoutee & _
        ":" & DerniereColonne & DerniereLigneAjoutee _
    ).PasteSpecial Paste:=xlPasteAll

    Application.CutCopyMode = False

End Sub


' =========================================================
' Un clic sur le bouton de recalcul reecrirait le tableau
' des sequences aux anciennes adresses : on le retire.
' =========================================================

Private Sub MasquerBoutonRecalcul()

    Dim ws As Worksheet
    Dim Forme As Shape
    Dim Macro As String

    Set ws = ThisWorkbook.Worksheets("Stats match")

    For Each Forme In ws.Shapes

        Macro = ""

        On Error Resume Next
        Macro = Forme.OnAction
        On Error GoTo 0

        If InStr( _
            1, _
            Macro, _
            "RecalculerStatsMatch", _
            vbTextCompare _
        ) > 0 Then

            Forme.Visible = msoFalse

        End If

    Next Forme

End Sub


' =========================================================
' Seules Journal actions et Stats match restent visibles.
'
' Les autres feuilles sont masquees et non supprimees :
' Parametres porte les libelles d'actions (ACT_*) utilises
' par toutes les formules COUNTIFS, et Compo alimente la
' liste des joueurs ainsi que le titre du tableau.
' =========================================================

Private Sub MasquerFeuillesInutiles()

    Dim ws As Worksheet

    ThisWorkbook.Worksheets("Stats match").Activate

    For Each ws In ThisWorkbook.Worksheets

        Select Case ws.Name

            Case "Journal actions", "Stats match"

                ws.Visible = xlSheetVisible

            Case Else

                ws.Visible = xlSheetVeryHidden

        End Select

    Next ws

End Sub
