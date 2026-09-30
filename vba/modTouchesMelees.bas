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

' Le premier refus, cite tel quel dans le message : un
' compteur seul ne dit pas ce qu'Excel a rejete.
Private DetailRefus As String


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

    EtapeEnCours = "bouton d'export"
    ConstruireRecapitulatifs

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
    ' "3/4" serait converti en date par Excel dans une
    ' liste deroulante, et s'afficherait "03-avr".
    EcrireUneListe ws, 8, "LST_TM_UTILISATION", _
        Array("Avants", "Trois-quart", "Pied")

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

    PoserListeSauteur lo

    PoserValidation ws, lo, "Lance pour", 0
    PoserValidation ws, lo, "Issue", 1
    PoserValidation ws, lo, "Zone saut", 3
    PoserValidation ws, lo, "Ballon", 4
    PoserValidation ws, lo, "Zone terrain", 5
    PoserValidation ws, lo, "Alignement", 6

    PoserMFCIssue lo, "Lance pour", "Issue"
    PoserMFCBallon lo, "Ballon"
    PoserMFCLigne lo, "Lance pour", BLEU_CLAIR, BLEU_ENTETE

End Sub


' ---------------------------------------------------------
' Tableau des melees
'
' Pas de "ND" ici : une melee non jouable n'existe pas
' dans la palette, qui n'a que BTN_ME_G et BTN_ME_P.
' ---------------------------------------------------------

Private Sub ConstruireTableauMelees(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim Ligne As Long

    If TableauExiste(ws, TAB_MELEES) Then Exit Sub

    Ligne = LigneTitreMelees(ws)

    ws.Cells(Ligne, 1).Value = "MELEES"

    ws.Range( _
        ws.Cells(Ligne + 1, 1), _
        ws.Cells(Ligne + 1, 10)).Value = Array( _
        "Temps video", "Mi-temps", "Introduction pour", _
        "Issue", _
        "Zone longueur", "Zone largeur", _
        "Jeu avant", "Jeu 3/4", "Jeu pied", "Observations")

    Set lo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range( _
            ws.Cells(Ligne + 1, 1), _
            ws.Cells(Ligne + 2, 10)), _
        XlListObjectHasHeaders:=xlYes)

    lo.Name = TAB_MELEES
    lo.TableStyle = "TableStyleMedium3"

    PoserValidation ws, lo, "Introduction pour", 0
    PoserValidation ws, lo, "Issue", 2
    PoserValidation ws, lo, "Zone longueur", 5
    PoserValidation ws, lo, "Zone largeur", 7

    PoserMFCIssue lo, "Introduction pour", "Issue"
    PoserMFCLigne lo, "Introduction pour", _
        ORANGE_CLAIR, ORANGE_ENTETE

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
' Ou commence le bloc des melees
'
' Sous le tableau des touches : une ligne de respiration,
' les trois lignes du recapitulatif, puis les deux lignes
' vides qui separent les deux blocs.
' ---------------------------------------------------------

Private Function LigneTitreMelees( _
    ByVal ws As Worksheet) As Long

    Dim lo As ListObject

    Set lo = ws.ListObjects(TAB_TOUCHES)

    LigneTitreMelees = lo.Range.Row + _
        lo.Range.Rows.Count + 6

End Function


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

' =========================================================
' LISTE DES SAUTEURS
'
' La popup remplit deja cette colonne, mais rien ne
' permettait de corriger une erreur dans le tableau : la
' meme liste que le journal y est donc posee.
'
' La source vit sur Parametres, colonne AA, tenue a jour
' par ActualiserJoueursJournal a chaque changement de
' composition. La validation pointe la plage, pas son
' contenu : elle suit donc les remplacements sans qu'on
' ait a la refaire.
' =========================================================

Public Sub ActualiserListeSauteur()

    On Error GoTo GestionErreur

    PoserListeSauteur _
        FeuilleTouchesMelees.ListObjects(TAB_TOUCHES)

    MsgBox _
        "La colonne Sauteur a retrouv" & ChrW(233) & _
        " sa liste.", _
        vbInformation, _
        "Touches et melees"

    Exit Sub

GestionErreur:

    MsgBox _
        "La pose de la liste a " & ChrW(233) & "chou" & _
        ChrW(233) & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


Private Sub PoserListeSauteur(ByVal lo As ListObject)

    Dim rng As Range
    Dim Derniere As Long
    Dim Source As String

    EtapeEnCours = "liste des sauteurs"

    Set rng = lo.ListColumns("Sauteur").DataBodyRange

    If rng Is Nothing Then Exit Sub

    ' La colonne AA porte les joueurs, puis Collectif et
    ' l'inconnu : on s'arrete a sa derniere valeur, sinon
    ' la liste deroulante se termine par des vides.
    Derniere = shParametres.Cells( _
        shParametres.Rows.Count, "AA").End(xlUp).Row

    If Derniere < 1 Then Exit Sub

    ' Le nom de la feuille porte un accent : il est lu,
    ' jamais ecrit, pour que ce module reste importable.
    Source = "='" & shParametres.Name & "'!" & _
        shParametres.Range("AA1:AA" & Derniere) _
            .Address(True, True)

    With rng.Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=Source

    End With

End Sub


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


' =========================================================
' BOUTON D'EXPORT
'
' En colonne AG, juste apres les listes : les tableaux
' grandissent vers le bas, le bouton ne serait jamais au
' meme endroit s'il les suivait.
'
' Le clic est capte par le module de la feuille, comme
' partout ailleurs dans ce classeur.
' =========================================================

Public Sub AjouterBoutonExport()

    On Error GoTo GestionErreur

    ' Passe par les recapitulatifs : le bouton se place
    ' sous celui des melees, qui doit donc exister.
    ConstruireRecapitulatifs

    MsgBox _
        "Le bouton d'export est en place sous le " & _
        "recapitulatif des melees.", _
        vbInformation, _
        "Touches et melees"

    Exit Sub

GestionErreur:

    MsgBox _
        "La pose du bouton a " & ChrW(233) & "chou" & _
        ChrW(233) & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


Private Sub PoserBoutonExport(ByVal ws As Worksheet)

    Dim Ancre As Range
    Dim Cible As Range

    ' Deux cases sous le recapitulatif des melees, qui en
    ' occupe trois lignes.
    On Error Resume Next
    Set Ancre = ThisWorkbook.Names(ANCRE_MELEES) _
        .RefersToRange
    On Error GoTo 0

    If Ancre Is Nothing Then Exit Sub

    ' L'emplacement precedent est libere avant le nouveau.
    On Error Resume Next
    ThisWorkbook.Names("TM_BTN_EXPORT") _
        .RefersToRange.MergeArea.Clear
    On Error GoTo 0

    Set Cible = Ancre.Offset(4, 0).Resize(2, 2)

    With Cible

        .Merge
        .Value = "EXPORTER" & vbLf & "LA FEUILLE"
        .Interior.Color = BLEU_ENTETE
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True

        .Borders.LineStyle = xlContinuous
        .Borders.Color = BLEU_ENTETE

    End With

    On Error Resume Next
    ThisWorkbook.Names("TM_BTN_EXPORT").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:="TM_BTN_EXPORT", _
        RefersTo:=Cible.Cells(1, 1)

End Sub


' ---------------------------------------------------------
' Appele par le module de la feuille a chaque selection.
' ---------------------------------------------------------

Public Sub ClicFeuilleTouchesMelees(ByVal Target As Range)

    Dim Bouton As Range

    On Error Resume Next
    Set Bouton = ThisWorkbook.Names("TM_BTN_EXPORT") _
        .RefersToRange
    On Error GoTo 0

    If Bouton Is Nothing Then Exit Sub

    If Intersect(Target, Bouton.MergeArea) Is Nothing Then
        Exit Sub
    End If

    ' La selection quitte le bouton, sinon un second clic
    ' au meme endroit ne declencherait rien. K2 separe les
    ' deux tableaux : on revient en haut de la feuille
    ' plutot qu'a leur pied.
    Target.Worksheet.Range("K2").Select

    ExporterTouchesMelees

End Sub


' =========================================================
' EXPORT DE LA FEUILLE SEULE
'
' Le contenu est recopie dans un classeur neuf, puis
' enregistre au format .xlsx.
'
' Copier la feuille elle-meme emporterait son module de
' code, qui appelle des procedures absentes du nouveau
' classeur : la moindre selection y declencherait une
' erreur de compilation. Seuls les valeurs et les formats
' voyagent donc.
'
' Tout est fige : le destinataire recoit des nombres, non
' des formules qui chercheraient des tableaux restes dans
' le fichier de match. Les couleurs conditionnelles, elles,
' continuent de fonctionner, car elles ne lisent que des
' cellules de la feuille.
' =========================================================

Public Sub ExporterTouchesMelees()

    Dim wsSource As Worksheet
    Dim wsCible As Worksheet
    Dim wbCible As Workbook
    Dim Chemin As Variant
    Dim Nom As String
    Dim EtatAlertes As Boolean

    On Error GoTo GestionErreur

    Set wsSource = FeuilleTouchesMelees

    Nom = ThisWorkbook.Name

    If InStrRev(Nom, ".") > 0 Then
        Nom = Left(Nom, InStrRev(Nom, ".") - 1)
    End If

    ' Le nom du match d'abord : les exports d'une meme
    ' saison se rangent ainsi les uns a la suite des
    ' autres.
    Nom = Nom & " - Touches et melees.xlsx"

    ' L'emplacement passe par la boite de dialogue, et non
    ' par un chemin ecrit dans le code : Excel pour Mac
    ' vit dans un bac a sable et refuse d'ecrire la ou
    ' l'utilisateur ne lui a pas ouvert l'acces.
    '
    ' Sans FileFilter : Excel pour Mac n'en veut pas.
    On Error Resume Next
    Chemin = Application.GetSaveAsFilename( _
        InitialFileName:=Nom, _
        Title:="Enregistrer la feuille des touches")
    On Error GoTo GestionErreur

    ' Annule par l'utilisateur.
    If VarType(Chemin) = vbBoolean Then
        Exit Sub
    End If

    ' La boite de dialogue n'a pas repondu : le dossier
    ' par defaut d'Excel est, lui, toujours accessible.
    If Trim(CStr(Chemin)) = "" Then

        Chemin = Application.DefaultFilePath & _
            Application.PathSeparator & Nom

    End If

    Set wbCible = Workbooks.Add(xlWBATWorksheet)
    Set wsCible = wbCible.Worksheets(1)

    wsCible.Name = FEUILLE_TM

    wsSource.UsedRange.Copy

    wsCible.Range("A1").PasteSpecial xlPasteColumnWidths
    wsCible.Range("A1").PasteSpecial xlPasteAll

    Application.CutCopyMode = False

    RendreAutonome wsCible

    ' Les deux tableaux sont refaits a l'identique : sans
    ' eux, le destinataire perdrait le filtrage.
    RecreerTableau wsSource, wsCible, TAB_TOUCHES
    RecreerTableau wsSource, wsCible, TAB_MELEES

    ReposerValidations wsCible
    RanimerRecapitulatifs wsCible

    EtatAlertes = Application.DisplayAlerts
    Application.DisplayAlerts = False

    wbCible.SaveAs _
        Filename:=CStr(Chemin), _
        FileFormat:=xlOpenXMLWorkbook

    Application.DisplayAlerts = EtatAlertes

    MsgBox _
        "La feuille est export" & ChrW(233) & "e dans :" & _
        vbCrLf & vbCrLf & wbCible.FullName & vbCrLf & _
        vbCrLf & "Le classeur reste ouvert pour v" & _
        ChrW(233) & "rification.", _
        vbInformation, _
        "Touches et melees"

    Exit Sub

GestionErreur:

    Application.CutCopyMode = False
    Application.DisplayAlerts = EtatAlertes

    MsgBox _
        "L'export a " & ChrW(233) & "chou" & ChrW(233) & _
        "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


' ---------------------------------------------------------
' Ce qui n'a pas de sens hors du fichier de match :
' les formules, qui pointeraient vers des tableaux restes
' la-bas ; le bouton d'export, sans macro pour l'entendre ;
' les listes deroulantes et leurs colonnes de reference.
' ---------------------------------------------------------

Private Sub RendreAutonome(ByVal ws As Worksheet)

    Dim Zone As Range

    Set Zone = ws.UsedRange

    ' Seules les formules sont figees : elles pointeraient
    ' des tableaux restes dans le fichier de match. Les
    ' listes deroulantes, elles, sont conservees pour que
    ' le destinataire puisse corriger une saisie.
    Zone.Value = Zone.Value

    EffacerBoutonExport ws

    ' Plus de quadrillage : hors des tableaux, la feuille
    ' est unie. Peindre les cellules en blanc donnerait le
    ' meme rendu, mais alourdirait le fichier et risquerait
    ' d'effacer les couleurs au passage.
    On Error Resume Next
    ws.Parent.Windows(1).DisplayGridlines = False
    On Error GoTo 0

End Sub


' ---------------------------------------------------------
' Les comptages de la copie
'
' Le figeage a transforme les recapitulatifs en nombres :
' ils ne bougeraient plus si le destinataire corrigeait
' une issue. Les formules sont donc reecrites sur les
' tableaux refaits, aux memes emplacements que dans la
' source.
' ---------------------------------------------------------

Private Sub RanimerRecapitulatifs(ByVal ws As Worksheet)

    On Error Resume Next

    EcrireComptages ws, ANCRE_TOUCHES, TAB_TOUCHES, _
        "Lance pour", True

    EcrireComptages ws, ANCRE_MELEES, TAB_MELEES, _
        "Introduction pour", False

    On Error GoTo 0

End Sub


Private Sub EcrireComptages( _
    ByVal ws As Worksheet, _
    ByVal NomAncre As String, _
    ByVal NomTableau As String, _
    ByVal ColonneEquipe As String, _
    ByVal AvecPasDroites As Boolean)

    Dim Source As Range
    Dim Ancre As Range

    ' L'ancre est celle du fichier de match : la copie a
    ' les memes coordonnees.
    Set Source = ThisWorkbook.Names(NomAncre).RefersToRange

    Set Ancre = ws.Cells(Source.Row, Source.Column)

    Ancre.Offset(1, 1).Formula = _
        Comptage(NomTableau, ColonneEquipe, "N", "G")
    Ancre.Offset(1, 2).Formula = _
        Comptage(NomTableau, ColonneEquipe, "N", "P")

    Ancre.Offset(2, 1).Formula = _
        Comptage(NomTableau, ColonneEquipe, "E", "P")
    Ancre.Offset(2, 2).Formula = _
        Comptage(NomTableau, ColonneEquipe, "E", "G")

    If AvecPasDroites Then

        Ancre.Offset(1, 3).Formula = _
            Comptage(NomTableau, ColonneEquipe, "N", "ND")
        Ancre.Offset(2, 3).Formula = _
            Comptage(NomTableau, ColonneEquipe, "E", "ND")

    End If

End Sub


' ---------------------------------------------------------
' Les listes deroulantes de la copie
'
' Elles sont reposees plutot que reprises du collage : une
' validation copiee d'un classeur a l'autre garde souvent
' une reference vers le classeur d'origine, et le fichier
' envoye ne doit dependre de rien.
'
' Les colonnes de choix ont voyage avec la feuille, aux
' memes adresses : les sources restent donc locales. Seuls
' les sauteurs vivaient sur Parametres, feuille qui n'est
' pas de l'export : leur liste est recopiee a cote des
' autres.
' ---------------------------------------------------------

Private Sub ReposerValidations(ByVal ws As Worksheet)

    Dim lo As ListObject

    On Error Resume Next

    Set lo = ws.ListObjects(TAB_TOUCHES)

    PoserValidation ws, lo, "Lance pour", 0
    PoserValidation ws, lo, "Issue", 1
    PoserValidation ws, lo, "Zone saut", 3
    PoserValidation ws, lo, "Ballon", 4
    PoserValidation ws, lo, "Zone terrain", 5
    PoserValidation ws, lo, "Alignement", 6

    CopierListeSauteur ws, lo

    Set lo = ws.ListObjects(TAB_MELEES)

    PoserValidation ws, lo, "Introduction pour", 0
    PoserValidation ws, lo, "Issue", 2
    PoserValidation ws, lo, "Zone longueur", 5
    PoserValidation ws, lo, "Zone largeur", 7

    ' Les colonnes de choix restent hors de vue.
    ws.Range( _
        ws.Columns(COL_LISTES), _
        ws.Columns(COL_LISTES + NB_LISTES)).Hidden = True

    On Error GoTo 0

End Sub


Private Sub CopierListeSauteur( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject)

    Dim Derniere As Long
    Dim Colonne As Long
    Dim rng As Range

    ' Juste apres les autres listes.
    Colonne = COL_LISTES + NB_LISTES

    Derniere = shParametres.Cells( _
        shParametres.Rows.Count, "AA").End(xlUp).Row

    If Derniere < 1 Then Exit Sub

    ws.Cells(1, Colonne).Value = "LST_SAUTEURS"

    ws.Range( _
        ws.Cells(2, Colonne), _
        ws.Cells(Derniere + 1, Colonne) _
    ).Value = shParametres.Range( _
        "AA1:AA" & Derniere).Value

    Set rng = lo.ListColumns("Sauteur").DataBodyRange

    If rng Is Nothing Then Exit Sub

    With rng.Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="=" & ws.Range( _
                ws.Cells(2, Colonne), _
                ws.Cells(Derniere + 1, Colonne) _
            ).Address(True, True)

    End With

End Sub


' ---------------------------------------------------------
' Le tableau structure de la source, refait sur la meme
' plage dans la copie : meme nom, meme style, et toujours
' sans bandes, que les lignes colorees remplacent.
' ---------------------------------------------------------

Private Sub RecreerTableau( _
    ByVal wsSource As Worksheet, _
    ByVal wsCible As Worksheet, _
    ByVal NomTableau As String)

    Dim loSource As ListObject
    Dim loCible As ListObject
    Dim Adresse As String
    Dim Style As String

    On Error GoTo Sortie

    Set loSource = wsSource.ListObjects(NomTableau)

    Adresse = loSource.Range.Address
    Style = loSource.TableStyle

    Set loCible = wsCible.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=wsCible.Range(Adresse), _
        XlListObjectHasHeaders:=xlYes)

    loCible.Name = NomTableau
    loCible.TableStyle = Style
    loCible.ShowTableStyleRowStripes = False

Sortie:

End Sub


Private Sub EffacerBoutonExport(ByVal ws As Worksheet)

    Dim Cellule As Range
    Dim Zone As Range

    For Each Cellule In ws.UsedRange

        If InStr(1, CStr(Cellule.Value), "EXPORTER", _
            vbTextCompare) > 0 Then

            ' La zone est retenue avant la defusion :
            ' apres, MergeArea ne rendrait plus que la
            ' cellule d'ancrage, et les autres garderaient
            ' leur fond bleu.
            Set Zone = Cellule.MergeArea
            Zone.UnMerge
            Zone.Clear

            Exit For

        End If

    Next Cellule

End Sub


' =========================================================
' REORGANISATION EN COLONNE
'
' Les deux tableaux etaient cote a cote. Le bloc des
' melees passe sous le recapitulatif des touches, dans les
' memes colonnes.
'
' A lancer une seule fois par classeur : elle ne fait rien
' si le deplacement a deja eu lieu.
'
' Le bloc est deplace d'un seul couper-coller, si bien
' qu'Excel emporte avec lui le tableau structure, son
' style, ses listes deroulantes et ses mises en forme
' conditionnelles.
' =========================================================

Public Sub ReorganiserFeuilleTouchesMelees()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim Source As Range
    Dim Ligne As Long
    Dim DerniereLigne As Long

    On Error GoTo GestionErreur

    Set ws = FeuilleTouchesMelees
    Set lo = ws.ListObjects(TAB_MELEES)

    If lo.Range.Column = 1 Then

        MsgBox _
            "Les deux tableaux sont d" & ChrW(233) & _
            "j" & ChrW(224) & " l'un sous l'autre.", _
            vbInformation, _
            "Touches et melees"

        Exit Sub

    End If

    Ligne = LigneTitreMelees(ws)

    DerniereLigne = ws.UsedRange.Row + _
        ws.UsedRange.Rows.Count - 1

    ' Tout ce qui vit a droite : le titre, le tableau, le
    ' recapitulatif et le bouton d'export.
    Set Source = ws.Range( _
        ws.Cells(1, lo.Range.Column), _
        ws.Cells(DerniereLigne, lo.Range.Column + 9))

    Source.Cut Destination:=ws.Cells(Ligne, 1)

    Application.CutCopyMode = False

    ' La colonne qui separait les deux tableaux n'a plus
    ' d'objet.
    ws.Columns("K").ColumnWidth = 12

    ws.Cells(Ligne, 1).Font.Bold = True
    ws.Cells(Ligne, 1).Font.Size = 14

    ' Recapitulatifs et bouton reprennent leur place sous
    ' leurs tableaux respectifs.
    ConstruireRecapitulatifs

    MsgBox _
        "Le bloc des m" & ChrW(234) & "l" & ChrW(233) & _
        "es est pass" & ChrW(233) & " sous celui des " & _
        "touches.", _
        vbInformation, _
        "Touches et melees"

    Exit Sub

GestionErreur:

    Application.CutCopyMode = False

    MsgBox _
        "La r" & ChrW(233) & "organisation a " & _
        ChrW(233) & "chou" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


' =========================================================
' RECOLORATION DES TABLEAUX EXISTANTS
'
' Les tableaux deja construits gardent leurs bandes une
' ligne sur deux : cette macro repose toutes leurs mises
' en forme conditionnelles.
'
' A lancer aussi sur un fichier de match deja saisi.
' =========================================================

Public Sub RecolorerTouchesMelees()

    Dim ws As Worksheet
    Dim lo As ListObject

    MFCEchouees = 0
    DetailRefus = ""

    On Error GoTo GestionErreur

    Set ws = FeuilleTouchesMelees

    EtapeEnCours = "tableau des touches"
    Set lo = ws.ListObjects(TAB_TOUCHES)
    AssurerUneLigne lo
    lo.DataBodyRange.FormatConditions.Delete
    PoserMFCIssue lo, "Lance pour", "Issue"
    PoserMFCBallon lo, "Ballon"
    PoserMFCLigne lo, "Lance pour", BLEU_CLAIR, BLEU_ENTETE

    EtapeEnCours = "tableau des melees"
    Set lo = ws.ListObjects(TAB_MELEES)
    AssurerUneLigne lo
    lo.DataBodyRange.FormatConditions.Delete
    PoserMFCIssue lo, "Introduction pour", "Issue"
    PoserMFCLigne lo, "Introduction pour", _
        ORANGE_CLAIR, ORANGE_ENTETE

    If MFCEchouees = 0 Then

        MsgBox _
            "Les deux tableaux sont recolores.", _
            vbInformation, _
            "Touches et melees"

    Else

        MsgBox _
            MFCEchouees & " mise(s) en forme refusee(s) " & _
            "par Excel." & vbCrLf & vbCrLf & DetailRefus, _
            vbExclamation, _
            "Touches et melees"

    End If

    Exit Sub

GestionErreur:

    MsgBox _
        "La recoloration a echoue pendant : " & _
        EtapeEnCours & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Touches et melees"

End Sub


' ---------------------------------------------------------
' Coloration des lignes
'
' Les bandes une ligne sur deux ne disaient rien : la
' couleur sert desormais a distinguer nos touches et nos
' melees de celles de l'adversaire.
'
' Deux regles, posees en derniere priorite pour que les
' mises en forme de l'issue et du ballon gardent la main
' sur les cellules qu'elles colorent :
'
'   - un trait sous chaque ligne remplie, sans quoi deux
'     lignes de meme couleur se confondraient ;
'   - le fond colore quand le lancer ou l'introduction
'     nous revient.
' ---------------------------------------------------------

' ---------------------------------------------------------
' Un tableau sans ligne de donnees n'a pas de
' DataBodyRange : il n'y a alors aucune cellule a mettre
' en forme, et Excel repond par l'erreur 91.
'
' La ligne vide ajoutee ici est celle que les tableaux
' portent de toute facon a leur construction.
' ---------------------------------------------------------

Private Sub AssurerUneLigne(ByVal lo As ListObject)

    If lo.ListRows.Count = 0 Then
        lo.ListRows.Add
    End If

End Sub


Private Sub PoserMFCLigne( _
    ByVal lo As ListObject, _
    ByVal NomColonne As String, _
    ByVal CouleurFond As Long, _
    ByVal CouleurBordure As Long)

    Dim rng As Range
    Dim Lettre As String
    Dim Ligne As Long

    EtapeEnCours = "MFC ligne"

    Set rng = lo.DataBodyRange

    Lettre = LettreColonne( _
        lo.ListColumns(NomColonne).Range.Column)

    Ligne = rng.Row

    On Error Resume Next
    lo.ShowTableStyleRowStripes = False
    On Error GoTo 0

    ' Une seule regle par ligne, et dans cet ordre : Excel
    ' pour Mac marque toute regle posee par VBA en "arreter
    ' si vrai", si bien qu'une premiere regle vraie partout
    ' empecherait la seconde d'etre evaluee. Chacune porte
    ' donc son trait.
    AjouterMFCLigne rng, _
        "=$" & Lettre & Ligne & "=""N""", _
        CouleurFond, CouleurBordure

    AjouterMFCLigne rng, _
        "=$" & Lettre & Ligne & "<>""""", _
        -1, CouleurBordure

End Sub


Private Sub AjouterMFCLigne( _
    ByVal rng As Range, _
    ByVal Formule As String, _
    ByVal CouleurFond As Long, _
    ByVal CouleurBordure As Long)

    Dim fc As FormatCondition

    ' Seul l'ajout est surveille : ajoutee en dernier, la
    ' regle porte deja la priorite la plus faible, et
    ' SetLastPriority n'a pas lieu d'etre.
    On Error GoTo MFCRefusee

    Set fc = rng.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:=Formule)

    ' Excel pour Mac refuse certaines proprietes sans que
    ' la regle elle-meme soit perdue : chacune est posee
    ' separement.
    On Error Resume Next

    If CouleurFond >= 0 Then
        fc.Interior.Color = CouleurFond
    End If

    If CouleurBordure >= 0 Then
        fc.Borders(xlBottom).LineStyle = xlContinuous
        fc.Borders(xlBottom).Color = CouleurBordure
    End If

    Exit Sub

MFCRefusee:

    MFCEchouees = MFCEchouees + 1

    If DetailRefus = "" Then
        DetailRefus = "Erreur " & Err.Number & " : " & _
            Err.Description & vbCrLf & "Formule : " & Formule
    End If

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

    Dim Titre As Range

    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Les deux tableaux partagent les memes colonnes : le
    ' titre des melees se trouve ou le bloc commence.
    Set Titre = ws.Cells(LigneTitreMelees(ws), 1)
    Titre.Font.Bold = True
    Titre.Font.Size = 14

    ws.Columns("A:J").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 30

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

    PoserBoutonExport ws

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


' Le recapitulatif se replace sous le tableau a chaque
' ecriture : fige, il finirait recouvert par les lignes
' qui s'ajoutent.
'
' L'emplacement precedent est efface avant le calcul du
' nouveau, sinon le recapitulatif se dupliquerait a
' chaque descente.
Private Function AncreRecap( _
    ByVal ws As Worksheet, _
    ByVal NomAncre As String, _
    ByVal lo As ListObject) As Range

    Dim Ligne As Long
    Dim Ancienne As Range

    On Error Resume Next
    Set Ancienne = ThisWorkbook.Names(NomAncre) _
        .RefersToRange
    On Error GoTo 0

    If Not Ancienne Is Nothing Then
        Ancienne.Resize(3, 4).Clear
    End If

    ' Deux lignes de respiration sous le tableau.
    Ligne = lo.Range.Row + lo.Range.Rows.Count + 1

    Set AncreRecap = ws.Cells(Ligne, lo.Range.Column)

    On Error Resume Next
    ThisWorkbook.Names(NomAncre).Delete
    On Error GoTo 0

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
