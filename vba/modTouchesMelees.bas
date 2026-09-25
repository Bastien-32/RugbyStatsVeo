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
Private Const NB_LISTES As Long = 10

' Couleurs des mises en forme conditionnelles. Const
' n'accepte pas RGB(), d'ou les valeurs calculees :
' RGB(r, g, b) = r + g * 256 + b * 65536.
Private Const ROUGE As Long = 192          ' RGB(192, 0, 0)
Private Const VERT As Long = 32768         ' RGB(0, 128, 0)
Private Const BLEU As Long = 12611584      ' RGB(0, 112, 192)

' Ancres des recapitulatifs : des plages nommees, qui
' suivent les decalages quand les tableaux grandissent.
Private Const ANCRE_TOUCHES As String = "TM_RECAP_TOUCHES"
Private Const ANCRE_MELEES As String = "TM_RECAP_MELEES"

' Chaque recapitulatif reprend la teinte du tableau qu'il
' resume : bleu pour les touches, orange pour les melees.
Private Const BLEU_ENTETE As Long = 12874308  ' RGB(68,114,196)
Private Const BLEU_CLAIR As Long = 15983321   ' RGB(217,226,243)
Private Const ORANGE_ENTETE As Long = 3243501 ' RGB(237,125,49)
Private Const ORANGE_CLAIR As Long = 14083324 ' RGB(252,228,214)

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

    ' Les popups parlent en toutes lettres, les tableaux
    ' en abrege : deux listes, une conversion a l'ecriture.
    EcrireUneListe ws, 8, "LST_TM_UTILISATION", _
        Array("Avants", "3/4", "Pied")

    EcrireUneListe ws, 9, "LST_TM_BALLON_LONG", _
        Array("Chaud", "Froid")

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
        "LST_TM_ALIGNEMENT", "LST_TM_ZONE_LARGEUR", _
        "LST_TM_UTILISATION", "LST_TM_BALLON_LONG")

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
        "Temps video", "Mi-temps", "Introduction pour", _
        "Issue", _
        "Zone longueur", "Zone largeur", _
        "Jeu avant", "Jeu 3/4", "Jeu pied", "Observations")

    Set lo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("L2:U3"), _
        XlListObjectHasHeaders:=xlYes)

    lo.Name = TAB_MELEES
    lo.TableStyle = "TableStyleMedium3"

    PoserValidation ws, lo, "Introduction pour", 0
    PoserValidation ws, lo, "Issue", 2
    PoserValidation ws, lo, "Zone longueur", 5
    PoserValidation ws, lo, "Zone largeur", 7

    PoserMFCIssue lo, "Introduction pour", "Issue"

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
' Recapitulatifs
'
' Deux lignes sous chaque tableau, du point de vue du
' lanceur : "pour l'adversaire, perdue" est donc une
' touche que nous avons prise.
'
' Les formules sont en references structurees, donc elles
' suivent le tableau quand il grandit. Et Excel decale ce
' qui se trouve sous un tableau, dans ses colonnes
' seulement : les deux blocs bougent independamment.
'
' L'ancre est une plage nommee, qui suit elle aussi les
' decalages : c'est par elle qu'on retrouve le bloc pour
' le reconstruire.
' ---------------------------------------------------------

Public Sub ConstruireRecapitulatifs()

    Dim ws As Worksheet

    On Error GoTo GestionErreur

    Set ws = ThisWorkbook.Sheets(FEUILLE_TM)

    EcrireRecap ws, ANCRE_TOUCHES, TAB_TOUCHES, _
        "Lance pour", "Touches", True, _
        BLEU_ENTETE, BLEU_CLAIR

    EcrireRecap ws, ANCRE_MELEES, TAB_MELEES, _
        "Introduction pour", "Melees", False, _
        ORANGE_ENTETE, ORANGE_CLAIR

    Exit Sub

GestionErreur:

    MsgBox _
        "Le recapitulatif n'a pas pu etre construit." & _
        vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


Private Sub EcrireRecap( _
    ByVal ws As Worksheet, _
    ByVal NomAncre As String, _
    ByVal NomTableau As String, _
    ByVal ColonneEquipe As String, _
    ByVal Titre As String, _
    ByVal AvecPasDroites As Boolean, _
    ByVal CouleurEntete As Long, _
    ByVal CouleurClaire As Long)

    Dim lo As ListObject
    Dim Ancre As Range
    Dim NbColonnes As Long

    Set lo = ws.ListObjects(NomTableau)

    NbColonnes = IIf(AvecPasDroites, 4, 3)

    Set Ancre = AncreRecap(ws, NomAncre, lo)

    Ancre.Resize(3, NbColonnes).Clear

    ' En-tete
    Ancre.Value = Titre
    Ancre.Offset(0, 1).Value = "Gagnees"
    Ancre.Offset(0, 2).Value = "Perdues"

    If AvecPasDroites Then
        Ancre.Offset(0, 3).Value = "Pas droites"
    End If

    ' Notre lancer : G nous revient, P nous echappe.
    Ancre.Offset(1, 0).Value = "Pour nous"
    Ancre.Offset(1, 1).Formula = _
        Comptage(NomTableau, ColonneEquipe, "N", "G")
    Ancre.Offset(1, 2).Formula = _
        Comptage(NomTableau, ColonneEquipe, "N", "P")

    If AvecPasDroites Then
        Ancre.Offset(1, 3).Formula = _
            Comptage(NomTableau, ColonneEquipe, "N", "ND")
    End If

    ' Leur lancer : ce qu'ils gagnent est ce que nous
    ' n'avons pas pris, d'ou l'inversion.
    Ancre.Offset(2, 0).Value = "Pour l'adversaire"
    Ancre.Offset(2, 1).Formula = _
        Comptage(NomTableau, ColonneEquipe, "E", "P")
    Ancre.Offset(2, 2).Formula = _
        Comptage(NomTableau, ColonneEquipe, "E", "G")

    If AvecPasDroites Then
        Ancre.Offset(2, 3).Formula = _
            Comptage(NomTableau, ColonneEquipe, "E", "ND")
    End If

    MettreEnFormeRecap Ancre, NbColonnes, _
        CouleurEntete, CouleurClaire

End Sub


Private Function Comptage( _
    ByVal NomTableau As String, _
    ByVal ColonneEquipe As String, _
    ByVal Equipe As String, _
    ByVal Issue As String) As String

    Comptage = "=COUNTIFS(" & _
        NomTableau & "[" & ColonneEquipe & "]," & _
        """" & Equipe & """," & _
        NomTableau & "[Issue]," & _
        """" & Issue & """)"

End Function


Private Function AncreRecap( _
    ByVal ws As Worksheet, _
    ByVal NomAncre As String, _
    ByVal lo As ListObject) As Range

    Dim Ligne As Long

    On Error Resume Next
    Set AncreRecap = ThisWorkbook.Names(NomAncre) _
        .RefersToRange
    On Error GoTo 0

    If Not AncreRecap Is Nothing Then Exit Function

    ' Deux lignes de respiration sous le tableau.
    Ligne = lo.Range.Row + lo.Range.Rows.Count + 1

    Set AncreRecap = ws.Cells(Ligne, lo.Range.Column)

    ThisWorkbook.Names.Add _
        Name:=NomAncre, _
        RefersTo:=AncreRecap

End Function


Private Sub MettreEnFormeRecap( _
    ByVal Ancre As Range, _
    ByVal NbColonnes As Long, _
    ByVal CouleurEntete As Long, _
    ByVal CouleurClaire As Long)

    With Ancre.Resize(1, NbColonnes)
        .Interior.Color = CouleurEntete
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    Ancre.HorizontalAlignment = xlLeft

    With Ancre.Offset(1, 0).Resize(1, NbColonnes)
        .Interior.Color = CouleurClaire
    End With

    With Ancre.Resize(3, NbColonnes)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
    End With

    With Ancre.Offset(1, 1).Resize(2, NbColonnes - 1)
        .HorizontalAlignment = xlCenter
    End With

End Sub


' ---------------------------------------------------------
' Puces du tableau des melees
'
' Appele par Worksheet_BeforeDoubleClick de la feuille :
' un double-clic dans Jeu avant, Jeu 3/4 ou Jeu pied pose
' la coche, un second la retire.
'
' Le double-clic plutot que le clic simple : parcourir le
' tableau ne doit jamais effacer une saisie par megarde.
'
' Les trois colonnes s'excluent : poser une coche efface
' les deux autres, qu'elles viennent de la popup ou d'un
' double-clic precedent.
' ---------------------------------------------------------

Public Sub BasculerPuceMelee(ByVal Target As Range)

    Dim lo As ListObject
    Dim Colonnes As Variant
    Dim Ligne As Range
    Dim Cellule As Range
    Dim i As Long
    Dim Index As Long
    Dim EtaitCochee As Boolean

    If Target.Cells.Count > 1 Then Exit Sub

    On Error GoTo Sortie

    Set lo = ThisWorkbook.Sheets(FEUILLE_TM) _
        .ListObjects(TAB_MELEES)

    If lo.DataBodyRange Is Nothing Then Exit Sub

    If Intersect(Target, lo.DataBodyRange) Is Nothing Then
        Exit Sub
    End If

    Colonnes = Array("Jeu avant", "Jeu 3/4", "Jeu pied")

    ' Le clic doit tomber dans l'une des trois colonnes.
    Index = -1

    For i = 0 To 2

        If Target.Column = lo.ListColumns( _
            CStr(Colonnes(i))).Range.Column Then

            Index = i
            Exit For

        End If

    Next i

    If Index < 0 Then Exit Sub

    Set Ligne = Intersect( _
        Target.EntireRow, lo.DataBodyRange)

    EtaitCochee = (Trim(CStr(Target.Value)) <> "")

    Application.EnableEvents = False

    For i = 0 To 2

        Set Cellule = Ligne.Cells( _
            1, lo.ListColumns(CStr(Colonnes(i))).Index)

        Cellule.ClearContents

    Next i

    If Not EtaitCochee Then Target.Value = CocheTM

Sortie:

    Application.EnableEvents = True

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
