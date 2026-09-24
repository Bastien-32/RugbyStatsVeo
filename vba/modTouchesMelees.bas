Attribute VB_Name = "modTouchesMelees"
Option Explicit

' =========================================================
' TOUCHES ET MELEES - FABRICATION DE LA FEUILLE
'
' A executer une fois depuis "Createur de match.xlsm".
'
' La macro fabrique la feuille "Touches Melees", ses deux
' tableaux, ses listes deroulantes et ses mises en forme
' conditionnelles. Elle est idempotente : relancee, elle
' remet en place ce qui manque sans toucher aux lignes
' deja saisies.
'
' Aucun caractere accentue dans ce module, pas meme dans
' les chaines : le VBE de macOS importe les .bas en Mac
' Roman, et tout accent ressort casse dans les en-tetes.
'
' Les lignes sont ecrites par la popup au moment de la
' saisie video. Les listes deroulantes ne sont la que pour
' corriger apres coup une possession mal renseignee.
' =========================================================

Public Const FEUILLE_TM As String = "Touches Melees"

Private Const TAB_TOUCHES As String = "DetailTouches"
Private Const TAB_MELEES As String = "DetailMelees"

' Colonne ou commence la reserve de listes, masquee.
Private Const COL_LISTES As Long = 23
Private Const NB_LISTES As Long = 8

' Couleurs des mises en forme conditionnelles. Const
' n'accepte pas RGB(), d'ou les valeurs calculees :
' RGB(r, g, b) = r + g * 256 + b * 65536.
Private Const ROUGE As Long = 192          ' RGB(192, 0, 0)
Private Const VERT As Long = 32768         ' RGB(0, 128, 0)
Private Const BLEU As Long = 12611584      ' RGB(0, 112, 192)

' Etape en cours, citee par le message d'erreur.
Private EtapeEnCours As String

' Mises en forme conditionnelles refusees par Excel.
Private MFCEchouees As Long


Public Sub ConstruireFeuilleTouchesMelees()

    Dim ws As Worksheet
    Dim EtatAffichage As Boolean

    EtatAffichage = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    MFCEchouees = 0

    EtapeEnCours = "ouverture de la feuille"
    Set ws = FeuilleTouchesMelees

    ' Les listes en premier : les validations des deux
    ' tableaux s'y referent par adresse.
    EtapeEnCours = "ecriture des listes"
    EcrireListes ws

    EtapeEnCours = "tableau des touches"
    ConstruireTableauTouches ws

    EtapeEnCours = "tableau des melees"
    ConstruireTableauMelees ws

    EtapeEnCours = "mise en forme"
    MettreEnFormeFeuille ws

    Application.ScreenUpdating = EtatAffichage

    If MFCEchouees = 0 Then

        MsgBox _
            "La feuille """ & FEUILLE_TM & """ est prete.", _
            vbInformation, _
            "Touches et melees"

    Else

        MsgBox _
            "La feuille """ & FEUILLE_TM & """ est prete." & _
            vbCrLf & vbCrLf & _
            MFCEchouees & " mise(s) en forme " & _
            "conditionnelle refusee(s) par Excel : " & _
            "les tableaux et les listes sont en place, " & _
            "les couleurs non.", _
            vbExclamation, _
            "Touches et melees"

    End If

    Exit Sub

GestionErreur:

    Application.ScreenUpdating = EtatAffichage

    MsgBox _
        "La construction a echoue pendant : " & _
        EtapeEnCours & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


' ---------------------------------------------------------
' Remise a zero
'
' Utile pendant la mise au point : supprime la feuille
' pour repartir d'une construction propre.
' ---------------------------------------------------------

Public Sub SupprimerFeuilleTouchesMelees()

    Dim ws As Worksheet
    Dim EtatAlertes As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_TM)
    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "Aucune feuille """ & FEUILLE_TM & """.", _
            vbInformation, _
            "Touches et melees"

        Exit Sub

    End If

    If MsgBox( _
        "Supprimer la feuille """ & FEUILLE_TM & _
        """ et tout son contenu ?", _
        vbYesNo + vbExclamation, _
        "Touches et melees") <> vbYes Then Exit Sub

    EtatAlertes = Application.DisplayAlerts
    Application.DisplayAlerts = False

    SupprimerNomsListes
    ws.Delete

    Application.DisplayAlerts = EtatAlertes

End Sub


' ---------------------------------------------------------
' La feuille elle-meme
' ---------------------------------------------------------

Private Function FeuilleTouchesMelees() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_TM)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = FEUILLE_TM

    End If

    Set FeuilleTouchesMelees = ws

End Function


' ---------------------------------------------------------
' Reserve de listes
'
' Les valeurs vivent dans des colonnes masquees de la
' feuille plutot que dans la formule de validation :
' "Z1,5" contient une virgule, qui serait prise pour un
' separateur d'elements.
' ---------------------------------------------------------

Private Sub EcrireListes(ByVal ws As Worksheet)

    EcrireUneListe ws, 0, "LST_TM_LANCE_POUR", _
        Array("N", "E")

    EcrireUneListe ws, 1, "LST_TM_ISSUE_TOUCHE", _
        Array("G", "P", "ND")

    EcrireUneListe ws, 2, "LST_TM_ISSUE_MELEE", _
        Array("G", "P")

    EcrireUneListe ws, 3, "LST_TM_ZONE_SAUT", _
        Array("Z0", "Z1", "Z1,5", "Z2", "Z3")

    EcrireUneListe ws, 4, "LST_TM_BALLON", _
        Array("C", "F")

    EcrireUneListe ws, 5, "LST_TM_ZONE_LONGUEUR", _
        Array("N en-but", "N 15m", "N 22m", "N 30m", _
              "N 40m", "50m", "E 40m", "E 30m", _
              "E 22m", "E 15m", "E en-but")

    EcrireUneListe ws, 6, "LST_TM_ALIGNEMENT", _
        Array("2", "3", "4", "5", "6", "Complet")

    EcrireUneListe ws, 7, "LST_TM_ZONE_LARGEUR", _
        Array("Gauche", "Milieu", "Droit")

    ws.Range( _
        ws.Columns(COL_LISTES), _
        ws.Columns(COL_LISTES + NB_LISTES - 1)).Hidden = True

End Sub


Private Sub EcrireUneListe( _
    ByVal ws As Worksheet, _
    ByVal Decalage As Long, _
    ByVal NomPlage As String, _
    ByVal Valeurs As Variant)

    Dim Colonne As Long
    Dim i As Long
    Dim Nombre As Long

    Colonne = COL_LISTES + Decalage
    Nombre = UBound(Valeurs) - LBound(Valeurs) + 1

    ws.Cells(1, Colonne).Value = NomPlage

    For i = 0 To Nombre - 1
        ws.Cells(2 + i, Colonne).Value = _
            Valeurs(LBound(Valeurs) + i)
    Next i

    On Error Resume Next
    ThisWorkbook.Names(NomPlage).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=NomPlage, _
        RefersTo:=PlageListe(ws, Decalage)

End Sub


' Plage des valeurs d'une liste, en-tete exclu.
Private Function PlageListe( _
    ByVal ws As Worksheet, _
    ByVal Decalage As Long) As Range

    Dim Colonne As Long
    Dim Derniere As Long

    Colonne = COL_LISTES + Decalage

    Derniere = ws.Cells(ws.Rows.Count, Colonne) _
        .End(xlUp).Row

    If Derniere < 2 Then Derniere = 2

    Set PlageListe = ws.Range( _
        ws.Cells(2, Colonne), _
        ws.Cells(Derniere, Colonne))

End Function


Private Sub SupprimerNomsListes()

    Dim Noms As Variant
    Dim i As Long

    Noms = Array( _
        "LST_TM_LANCE_POUR", "LST_TM_ISSUE_TOUCHE", _
        "LST_TM_ISSUE_MELEE", "LST_TM_ZONE_SAUT", _
        "LST_TM_BALLON", "LST_TM_ZONE_LONGUEUR", _
        "LST_TM_ALIGNEMENT", "LST_TM_ZONE_LARGEUR")

    For i = LBound(Noms) To UBound(Noms)

        On Error Resume Next
        ThisWorkbook.Names(CStr(Noms(i))).Delete
        On Error GoTo 0

    Next i

End Sub


' ---------------------------------------------------------
' Tableau des touches
' ---------------------------------------------------------

Private Sub ConstruireTableauTouches(ByVal ws As Worksheet)

    Dim lo As ListObject

    If TableauExiste(ws, TAB_TOUCHES) Then Exit Sub

    ws.Range("A1").Value = "TOUCHES"

    ws.Range("A2:J2").Value = Array( _
        "Temps video", "Mi-temps", "Lance pour", "Issue", _
        "Sauteur", "Zone saut", "Ballon", _
        "Zone terrain", "Alignement", "Observations")

    Set lo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("A2:J3"), _
        XlListObjectHasHeaders:=xlYes)

    lo.Name = TAB_TOUCHES
    lo.TableStyle = "TableStyleMedium2"

    PoserValidation ws, lo, "Lance pour", 0
    PoserValidation ws, lo, "Issue", 1
    PoserValidation ws, lo, "Zone saut", 3
    PoserValidation ws, lo, "Ballon", 4
    PoserValidation ws, lo, "Zone terrain", 5
    PoserValidation ws, lo, "Alignement", 6

    PoserMFCIssue lo, "Lance pour", "Issue"
    PoserMFCBallon lo, "Ballon"

End Sub


' ---------------------------------------------------------
' Tableau des melees
'
' Pas de "ND" ici : une melee non jouable n'existe pas
' dans la palette, qui n'a que BTN_ME_G et BTN_ME_P.
' ---------------------------------------------------------

Private Sub ConstruireTableauMelees(ByVal ws As Worksheet)

    Dim lo As ListObject

    If TableauExiste(ws, TAB_MELEES) Then Exit Sub

    ws.Range("L1").Value = "MELEES"

    ws.Range("L2:U2").Value = Array( _
        "Temps video", "Mi-temps", "Lance pour", "Issue", _
        "Zone longueur", "Zone largeur", _
        "Jeu avant", "Jeu 3/4", "Jeu pied", "Observations")

    Set lo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("L2:U3"), _
        XlListObjectHasHeaders:=xlYes)

    lo.Name = TAB_MELEES
    lo.TableStyle = "TableStyleMedium3"

    PoserValidation ws, lo, "Lance pour", 0
    PoserValidation ws, lo, "Issue", 2
    PoserValidation ws, lo, "Zone longueur", 5
    PoserValidation ws, lo, "Zone largeur", 7

    PoserMFCIssue lo, "Lance pour", "Issue"

    ' Les trois colonnes de jeu recevront leurs puces a
    ' l'etape suivante : un clic y ecrira la coche, et
    ' l'exclusivite sera tenue par Worksheet_Change.
    lo.ListColumns("Jeu avant").DataBodyRange _
        .HorizontalAlignment = xlCenter
    lo.ListColumns("Jeu 3/4").DataBodyRange _
        .HorizontalAlignment = xlCenter
    lo.ListColumns("Jeu pied").DataBodyRange _
        .HorizontalAlignment = xlCenter

End Sub


' ---------------------------------------------------------
' Listes deroulantes
'
' La source est donnee par adresse, pas par plage nommee :
' Excel pour Mac refuse une bonne partie des validations
' qui passent par un nom. Les listes etant sur la meme
' feuille, l'adresse suffit.
'
' Meme appel que modPlayers.ActualiserJoueursJournal, le
' seul de ce classeur qui tourne deja sous Excel pour Mac.
' Une case vide reste permise : IgnoreBlank vaut True par
' defaut, et la popup n'est donc pas bloquante.
' ---------------------------------------------------------

Private Sub PoserValidation( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject, _
    ByVal NomColonne As String, _
    ByVal Decalage As Long)

    Dim Source As String
    Dim rng As Range

    EtapeEnCours = "validation, colonne " & NomColonne

    Set rng = lo.ListColumns(NomColonne).DataBodyRange

    If rng Is Nothing Then
        Err.Raise 5, , _
            "Le tableau " & lo.Name & " n'a aucune " & _
            "ligne de donnees."
    End If

    Source = "=" & PlageListe(ws, Decalage) _
        .Address(True, True)

    EtapeEnCours = EtapeEnCours & " (source " & Source & ")"

    With rng.Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=Source

    End With

End Sub


' ---------------------------------------------------------
' Mise en forme conditionnelle de l'issue
'
' Seuls les deux cas remarquables sont colores : perdre
' son propre lancer, et prendre celui de l'adversaire.
' Les deux issues attendues restent neutres.
' ---------------------------------------------------------

Private Sub PoserMFCIssue( _
    ByVal lo As ListObject, _
    ByVal ColonneLance As String, _
    ByVal ColonneIssue As String)

    Dim rng As Range
    Dim LettreLance As String
    Dim LettreIssue As String
    Dim Ligne As Long

    EtapeEnCours = "MFC issue, tableau " & lo.Name

    Set rng = lo.ListColumns(ColonneIssue).DataBodyRange

    LettreLance = LettreColonne( _
        lo.ListColumns(ColonneLance).Range.Column)

    LettreIssue = LettreColonne( _
        lo.ListColumns(ColonneIssue).Range.Column)

    Ligne = rng.Row

    rng.FormatConditions.Delete

    ' Concatenation plutot que AND : aucun separateur
    ' d'arguments, donc rien qui depende de la langue.
    AjouterMFC rng, _
        "=($" & LettreLance & Ligne & "&$" & _
        LettreIssue & Ligne & ")=""NP""", ROUGE

    AjouterMFC rng, _
        "=($" & LettreLance & Ligne & "&$" & _
        LettreIssue & Ligne & ")=""EG""", VERT

End Sub


' ---------------------------------------------------------
' Mise en forme conditionnelle du ballon
'
' La lettre reste lisible, en blanc sur fond colore : une
' case qui parait vide invite a la remplir, et le comptage
' des C et des F doit rester verifiable a l'oeil.
' ---------------------------------------------------------

Private Sub PoserMFCBallon( _
    ByVal lo As ListObject, _
    ByVal NomColonne As String)

    Dim rng As Range
    Dim Lettre As String
    Dim Ligne As Long

    EtapeEnCours = "MFC ballon"

    Set rng = lo.ListColumns(NomColonne).DataBodyRange

    Lettre = LettreColonne( _
        lo.ListColumns(NomColonne).Range.Column)

    Ligne = rng.Row

    rng.FormatConditions.Delete

    AjouterMFC rng, _
        "=$" & Lettre & Ligne & "=""F""", BLEU

    AjouterMFC rng, _
        "=$" & Lettre & Ligne & "=""C""", ROUGE

End Sub


Private Sub AjouterMFC( _
    ByVal rng As Range, _
    ByVal Formule As String, _
    ByVal Couleur As Long)

    Dim fc As FormatCondition

    ' Les mises en forme conditionnelles sont le seul point
    ' de ce classeur sans precedent sous Excel pour Mac :
    ' un refus est compte et signale, il n'interrompt pas
    ' la construction des tableaux.
    On Error GoTo MFCRefusee

    Set fc = rng.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:=Formule)

    fc.Interior.Color = Couleur
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True

    Exit Sub

MFCRefusee:

    MFCEchouees = MFCEchouees + 1

End Sub


' ---------------------------------------------------------
' Presentation
' ---------------------------------------------------------

Private Sub MettreEnFormeFeuille(ByVal ws As Worksheet)

    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("L1").Font.Bold = True
    ws.Range("L1").Font.Size = 14

    ws.Columns("A:J").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 30
    ws.Columns("K").ColumnWidth = 3
    ws.Columns("L:U").ColumnWidth = 12
    ws.Columns("U").ColumnWidth = 30

    ' Les deux tableaux commencent ligne 3 : les volets
    ' gardent titres et en-tetes visibles au defilement.
    ws.Activate
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function TableauExiste( _
    ByVal ws As Worksheet, _
    ByVal Nom As String) As Boolean

    Dim lo As ListObject

    For Each lo In ws.ListObjects

        If StrComp(lo.Name, Nom, vbTextCompare) = 0 Then
            TableauExiste = True
            Exit Function
        End If

    Next lo

End Function


Private Function LettreColonne( _
    ByVal Colonne As Long) As String

    Dim Adresse As String

    Adresse = ThisWorkbook.Sheets(1) _
        .Cells(1, Colonne).Address(True, False)

    LettreColonne = Left( _
        Adresse, InStr(Adresse, "$") - 1)

End Function
