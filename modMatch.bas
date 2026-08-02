Option Explicit

Private FermetureConfirmee As Boolean


Public Sub InitialiserClasseurStatsRugby()

    FermetureEnCours = False

    DemarrerVeoVideoControl

    ' Le premier heartbeat est envoye immediatement.
    DemarrerHeartbeatVideo

    InitialiserEtatConnexionVideo
    AttendreUneSeconde
    ConnecterVideo
    ReinitialiserSelectionSaisieVideo

    If ActiveSheet Is shSaisieVideo Then

        ActiverRaccourcisVideo
        DemarrerAffichageChronoVideo

    Else

        DesactiverRaccourcisVideo
        ArreterAffichageChronoVideo

    End If

End Sub


Public Sub CreerFichierMatch()

    Dim wsCompo As Worksheet
    Dim wbModele As Workbook
    Dim wbMatch As Workbook

    Dim MatchID As String
    Dim NomFichier As String
    Dim DossierMatchs As String
    Dim CheminChoisi As String
    Dim Reponse As VbMsgBoxResult

    Dim EtatAffichage As Boolean
    Dim EtatAlertes As Boolean
    Dim EtatEvenements As Boolean

    Dim NomClasseurSecurise As String
    Dim NomProcedureInitialisation As String
    Dim BoutonIntrouvable As Boolean

    Dim NumeroErreur As Long
    Dim DescriptionErreur As String

    #If Mac Then
        Dim AccesAccorde As Boolean
        Dim FichiersDemandes As Variant
    #End If

    Set wbModele = ThisWorkbook
    Set wsCompo = wbModele.Worksheets("Compo")

    ' =====================================================
    ' 1. VERIFICATION DES INFORMATIONS OBLIGATOIRES
    ' =====================================================

    If Trim(CStr(wsCompo.Range("MATCH_SAISON").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_SAISON"), "Renseigne la saison."
        Exit Sub
    End If

    If Trim(CStr(wsCompo.Range("MATCH_DATE").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_DATE"), "Renseigne la date du match."
        Exit Sub
    End If

    If Trim(CStr(wsCompo.Range("MATCH_ADV").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_ADV"), "Renseigne l'equipe adverse."
        Exit Sub
    End If

    If Trim(CStr(wsCompo.Range("MATCH_CAT").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_CAT"), "Renseigne la categorie : Premiere, Reserve, Feminines ou Juniors."
        Exit Sub
    End If

    If Trim(CStr(wsCompo.Range("MATCH_LIEU").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_LIEU"), "Indique si le match est joue a domicile ou a l'exterieur."
        Exit Sub
    End If

    If Trim(CStr(wsCompo.Range("MATCH_PHASE").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_PHASE"), "Indique s'il s'agit de la phase aller ou retour."
        Exit Sub
    End If

    If Trim(CStr(wsCompo.Range("MATCH_JOURNEE").Value)) = "" Then
        AfficherChampMatchManquant wsCompo.Range("MATCH_JOURNEE"), "Renseigne le numero de journee."
        Exit Sub
    End If

    If wbModele.Path = "" Then
        MsgBox _
            "Enregistre d'abord le modele dans un dossier local.", _
            vbExclamation, _
            "Modele non enregistre"
        Exit Sub
    End If

    ' =====================================================
    ' 2. CONSTRUCTION DE MATCH_ID
    ' =====================================================

    MatchID = ConstruireMatchID(wsCompo)

    If MatchID = "" Then
        MsgBox _
            "L'identifiant du match n'a pas pu etre cree.", _
            vbExclamation, _
            "Match non identifie"
        Exit Sub
    End If

    NomFichier = NettoyerNomFichier(MatchID) & ".xlsm"

    DossierMatchs = _
        wbModele.Path & _
        Application.PathSeparator & _
        "Matchs"

    If Dir(DossierMatchs, vbDirectory) = "" Then
        MkDir DossierMatchs
    End If

    CheminChoisi = _
        DossierMatchs & _
        Application.PathSeparator & _
        NomFichier

    If StrComp(CheminChoisi, wbModele.FullName, vbTextCompare) = 0 Then
        MsgBox _
            "Le fichier du match ne peut pas remplacer le modele.", _
            vbExclamation, _
            "Nom de fichier incorrect"
        Exit Sub
    End If

    ' =====================================================
    ' 3. CONFIRMATION AVANT CREATION
    ' =====================================================

    Reponse = MsgBox( _
        "Le fichier de match suivant va " & ChrW(234) & "tre cr" & ChrW(233) & ChrW(233) & " :" & _
        vbCrLf & vbCrLf & _
        NomFichier & _
        vbCrLf & vbCrLf & _
        "Dans le dossier :" & _
        vbCrLf & _
        DossierMatchs & _
        vbCrLf & vbCrLf & _
        "Confirmer la cr" & ChrW(233) & "ation ?", _
        vbQuestion + vbOKCancel + vbDefaultButton2, _
        "Cr" & ChrW(233) & "er le fichier de match" _
    )

    If Reponse <> vbOK Then
        Exit Sub
    End If

    wsCompo.Range("MATCH_ID").Value = MatchID

    ' =====================================================
    ' 4. CONTROLE SI LE FICHIER EXISTE DEJA
    ' =====================================================

    If Dir(CheminChoisi) <> "" Then

        Reponse = MsgBox( _
            "Un fichier portant deja ce nom existe." & vbCrLf & vbCrLf & _
            "Veux-tu le remplacer ?", _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            "Fichier existant")

        If Reponse <> vbYes Then Exit Sub

        On Error Resume Next
        Kill CheminChoisi

        If Err.Number <> 0 Then
            MsgBox _
                "Le fichier existant n'a pas pu etre remplace." & vbCrLf & vbCrLf & _
                "Il est peut-etre deja ouvert.", _
                vbExclamation, _
                "Remplacement impossible"
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If

        On Error GoTo 0

    End If

    ' =====================================================
    ' 5. CREATION ET PREPARATION DE LA COPIE
    ' =====================================================

    EtatAffichage = Application.ScreenUpdating
    EtatAlertes = Application.DisplayAlerts
    EtatEvenements = Application.EnableEvents

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    wbModele.SaveCopyAs CheminChoisi

    #If Mac Then

        FichiersDemandes = Array(CheminChoisi)

        AccesAccorde = _
            GrantAccessToMultipleFiles(FichiersDemandes)

        If Not AccesAccorde Then

            On Error Resume Next
            Kill CheminChoisi
            On Error GoTo GestionErreur

            Application.DisplayAlerts = EtatAlertes
            Application.ScreenUpdating = EtatAffichage
            Application.EnableEvents = EtatEvenements

            MsgBox _
                "Le fichier du match a " & ChrW(233) & "t" & ChrW(233) & " cr" & ChrW(233) & ChrW(233) & " dans le dossier Matchs." & _
                vbCrLf & vbCrLf & _
                "Il n'a pas pu " & ChrW(234) & "tre ouvert car l'autorisation d'acc" & ChrW(232) & "s n'a pas " & _
                ChrW(233) & "t" & ChrW(233) & " accord" & ChrW(233) & "e.", _
                vbExclamation, _
                "Fichier cr" & ChrW(233) & ChrW(233)

            Exit Sub

        End If

    #End If

    ' Empeche le Workbook_Open de la copie de se lancer pendant
    ' que le modele execute encore cette macro.
    Application.EnableEvents = False

    Set wbMatch = Workbooks.Open( _
        Filename:=CheminChoisi, _
        ReadOnly:=False)

    ' Masque le bouton uniquement dans la copie.
    On Error Resume Next
    wbMatch.Worksheets("Compo") _
        .Shapes("BTN_CREER_FICHIER_MATCH") _
        .Visible = msoFalse

    BoutonIntrouvable = (Err.Number <> 0)
    Err.Clear
    On Error GoTo GestionErreur

    wbMatch.Save

    wbMatch.Activate
    wbMatch.Worksheets("Compo").Activate

    ' Prepare l'initialisation de la copie apres la fermeture du modele.
    NomClasseurSecurise = _
        Replace(wbMatch.Name, "'", "''")

    NomProcedureInitialisation = _
        "'" & _
        NomClasseurSecurise & _
        "'!InitialiserClasseurStatsRugby"

    ' Arrete les taches du modele avant de transmettre la main a la copie.
    FermetureEnCours = True
    ArreterHeartbeatVideo
    DesactiverRaccourcisVideo
    ArreterAffichageChronoVideo

    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    Application.OnTime _
        EarliestTime:=Now + TimeSerial(0, 0, 1), _
        Procedure:=NomProcedureInitialisation, _
        Schedule:=True

    If BoutonIntrouvable Then
        MsgBox _
            "Le fichier du match a ete cree, mais la forme " & _
            "BTN_CREER_FICHIER_MATCH est introuvable.", _
            vbExclamation, _
            "Bouton introuvable"
    Else
        MsgBox _
            "Le fichier du match a ete cree :" & vbCrLf & vbCrLf & _
            wbMatch.Name & vbCrLf & vbCrLf & _
            "Tu peux maintenant terminer la composition.", _
            vbInformation, _
            "Fichier du match cree"
    End If

    ' Derniere instruction : le Workbook_Open de la copie a ete
    ' remplace par une initialisation planifiee dans son propre projet VBA.
    wbModele.Close SaveChanges:=False
    Exit Sub

GestionErreur:

    NumeroErreur = Err.Number
    DescriptionErreur = Err.Description

    On Error Resume Next

    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    If Not wbMatch Is Nothing Then
        wbMatch.Close SaveChanges:=False
    End If

    On Error GoTo 0

    MsgBox _
        "Le fichier du match n'a pas pu etre cree." & vbCrLf & vbCrLf & _
        "Erreur " & NumeroErreur & " : " & DescriptionErreur, _
        vbCritical, _
        "Erreur de creation"

End Sub


Private Function ConstruireMatchID( _
    ByVal wsCompo As Worksheet) As String

    Dim Saison As String
    Dim PhaseCode As String
    Dim Journee As String
    Dim Adversaire As String
    Dim Categorie As String
    Dim LieuCode As String

    Saison = Trim(CStr(wsCompo.Range("MATCH_SAISON").Value))
    Journee = Trim(CStr(wsCompo.Range("MATCH_JOURNEE").Value))
    Adversaire = Trim(CStr(wsCompo.Range("MATCH_ADV").Value))
    Categorie = Trim(CStr(wsCompo.Range("MATCH_CAT").Value))

    If StrComp( _
        Trim(CStr(wsCompo.Range("MATCH_PHASE").Value)), _
        "Aller", _
        vbTextCompare) = 0 Then

        PhaseCode = "A"
    Else
        PhaseCode = "R"
    End If

    If StrComp( _
        Trim(CStr(wsCompo.Range("MATCH_LIEU").Value)), _
        "Domicile", _
        vbTextCompare) = 0 Then

        LieuCode = "DOM"
    Else
        LieuCode = "EXT"
    End If

    ConstruireMatchID = _
        Saison & "_" & _
        PhaseCode & Journee & "_" & _
        Adversaire & "_" & _
        Categorie & "_" & _
        LieuCode

End Function

Private Sub AfficherChampMatchManquant( _
    ByVal Cellule As Range, _
    ByVal Message As String)

    MsgBox Message, vbExclamation, "Informations du match incompletes"

    Cellule.Worksheet.Activate
    Cellule.Select

End Sub

Private Function NettoyerNomFichier( _
    ByVal Texte As String) As String

    Dim CaracteresInterdits As Variant
    Dim Caractere As Variant
    Dim Resultat As String

    Resultat = Trim(Texte)

    ' Retire les principaux accents afin de fiabiliser Excel Mac.
    Resultat = Replace(Resultat, ChrW(224), "a")
    Resultat = Replace(Resultat, ChrW(225), "a")
    Resultat = Replace(Resultat, ChrW(226), "a")
    Resultat = Replace(Resultat, ChrW(228), "a")
    Resultat = Replace(Resultat, ChrW(192), "A")
    Resultat = Replace(Resultat, ChrW(193), "A")
    Resultat = Replace(Resultat, ChrW(194), "A")
    Resultat = Replace(Resultat, ChrW(196), "A")

    Resultat = Replace(Resultat, ChrW(233), "e")
    Resultat = Replace(Resultat, ChrW(232), "e")
    Resultat = Replace(Resultat, ChrW(234), "e")
    Resultat = Replace(Resultat, ChrW(235), "e")
    Resultat = Replace(Resultat, ChrW(201), "E")
    Resultat = Replace(Resultat, ChrW(200), "E")
    Resultat = Replace(Resultat, ChrW(202), "E")
    Resultat = Replace(Resultat, ChrW(203), "E")

    Resultat = Replace(Resultat, ChrW(238), "i")
    Resultat = Replace(Resultat, ChrW(239), "i")
    Resultat = Replace(Resultat, ChrW(206), "I")
    Resultat = Replace(Resultat, ChrW(207), "I")

    Resultat = Replace(Resultat, ChrW(244), "o")
    Resultat = Replace(Resultat, ChrW(246), "o")
    Resultat = Replace(Resultat, ChrW(212), "O")
    Resultat = Replace(Resultat, ChrW(214), "O")

    Resultat = Replace(Resultat, ChrW(249), "u")
    Resultat = Replace(Resultat, ChrW(251), "u")
    Resultat = Replace(Resultat, ChrW(252), "u")
    Resultat = Replace(Resultat, ChrW(217), "U")
    Resultat = Replace(Resultat, ChrW(219), "U")
    Resultat = Replace(Resultat, ChrW(220), "U")

    Resultat = Replace(Resultat, ChrW(231), "c")
    Resultat = Replace(Resultat, ChrW(199), "C")

    CaracteresInterdits = Array( _
        "/", "\", ":", "*", "?", """", "<", ">", "|")

    For Each Caractere In CaracteresInterdits
        Resultat = Replace(Resultat, CStr(Caractere), "-")
    Next Caractere

    Resultat = Replace(Resultat, " ", "-")

    Do While InStr(Resultat, "--") > 0
        Resultat = Replace(Resultat, "--", "-")
    Loop

    NettoyerNomFichier = Resultat

End Function
