Attribute VB_Name = "modMatch"
Option Explicit

Private FermetureConfirmee As Boolean


Public Sub InitialiserClasseurStatsRugby()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False

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

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_SAISON"), _
            "Renseigne la saison."

        Exit Sub

    End If

    If Trim(CStr(wsCompo.Range("MATCH_DATE").Value)) = "" Then

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_DATE"), _
            "Renseigne la date du match."

        Exit Sub

    End If

    If Trim(CStr(wsCompo.Range("MATCH_ADV").Value)) = "" Then

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_ADV"), _
            "Renseigne l'equipe adverse."

        Exit Sub

    End If

    If Trim(CStr(wsCompo.Range("MATCH_CAT").Value)) = "" Then

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_CAT"), _
            "Renseigne la categorie : Premiere, Reserve, Feminines ou Juniors."

        Exit Sub

    End If

    If Trim(CStr(wsCompo.Range("MATCH_LIEU").Value)) = "" Then

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_LIEU"), _
            "Indique si le match est joue a domicile ou a l'exterieur."

        Exit Sub

    End If

    If Trim(CStr(wsCompo.Range("MATCH_PHASE").Value)) = "" Then

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_PHASE"), _
            "Indique s'il s'agit de la phase aller ou retour."

        Exit Sub

    End If

    If Trim(CStr(wsCompo.Range("MATCH_JOURNEE").Value)) = "" Then

        AfficherChampMatchManquant _
            wsCompo.Range("MATCH_JOURNEE"), _
            "Renseigne le numero de journee."

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
    ' 2. CONSTRUCTION DU MATCH_ID
    ' =====================================================

    MatchID = ConstruireMatchID(wsCompo)

    If MatchID = "" Then

        MsgBox _
            "L'identifiant du match n'a pas pu etre cree.", _
            vbExclamation, _
            "Match non identifie"

        Exit Sub

    End If

    NomFichier = _
        NettoyerNomFichier(MatchID) & ".xlsm"

    DossierMatchs = _
        wbModele.Path & _
        Application.PathSeparator & _
        "Matchs"

    #If Mac Then

        ' Sur Mac, ne pas tester le dossier avec Dir :
        ' Excel peut le d_clarer introuvable malgr_ son existence
        ' ö cause des autorisations de la sandbox.
        ' L'acc_s au fichier sera trait_ plus bas avec
        ' GrantAccessToMultipleFiles.
    
    #Else
    
        ' Sous Windows, cr_e automatiquement le dossier
        ' Matchs s'il n'existe pas.
        If Dir(DossierMatchs, vbDirectory) = "" Then
            MkDir DossierMatchs
        End If
    
    #End If

    CheminChoisi = _
        DossierMatchs & _
        Application.PathSeparator & _
        NomFichier

    If StrComp( _
        CheminChoisi, _
        wbModele.FullName, _
        vbTextCompare _
    ) = 0 Then

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
        "Le fichier de match suivant va " & _
        ChrW(234) & _
        "tre cr" & ChrW(233) & ChrW(233) & " :" & _
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
            "Un fichier portant deja ce nom existe." & _
            vbCrLf & vbCrLf & _
            "Veux-tu le remplacer ?", _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            "Fichier existant" _
        )

        If Reponse <> vbYes Then
            Exit Sub
        End If

        On Error Resume Next

        Kill CheminChoisi

        If Err.Number <> 0 Then

            MsgBox _
                "Le fichier existant n'a pas pu etre remplace." & _
                vbCrLf & vbCrLf & _
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

#If Mac Then

    FichiersDemandes = Array(CheminChoisi)

    AccesAccorde = _
        GrantAccessToMultipleFiles(FichiersDemandes)

    If Not AccesAccorde Then

        ' Redonne la main a l'ecran pour afficher
        ' le selecteur de dossier.
        Application.DisplayAlerts = EtatAlertes
        Application.ScreenUpdating = EtatAffichage

        AccesAccorde = DemanderAutorisationDossierMatchs()

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        If AccesAccorde Then

            FichiersDemandes = Array(CheminChoisi)

            AccesAccorde = _
                GrantAccessToMultipleFiles(FichiersDemandes)

        End If

    End If

    If Not AccesAccorde Then

        Application.DisplayAlerts = EtatAlertes
        Application.ScreenUpdating = EtatAffichage
        Application.EnableEvents = EtatEvenements

        MsgBox _
            "Le fichier du match n'a pas pu " & _
            ChrW(234) & "tre cr" & ChrW(233) & ChrW(233) & _
            " car l'autorisation d'acc" & ChrW(232) & _
            "s au dossier Matchs n'a pas " & _
            ChrW(233) & "t" & ChrW(233) & " accord" & ChrW(233) & "e." & _
            vbCrLf & vbCrLf & _
            "Relance la cr" & ChrW(233) & "ation puis s" & ChrW(233) & _
            "lectionne le dossier :" & _
            vbCrLf & DossierMatchs, _
            vbExclamation, _
            "Autorisation refus" & ChrW(233) & "e"

        Exit Sub

    End If

#End If

    wbModele.SaveCopyAs CheminChoisi

    ' Empeche le Workbook_Open de la copie de se lancer
    ' pendant que le modele execute encore cette macro.
    Application.EnableEvents = False

    Set wbMatch = Workbooks.Open( _
        Filename:=CheminChoisi, _
        ReadOnly:=False _
    )
    
    Application.EnableEvents = EtatEvenements

    ' Masque le bouton uniquement dans la copie.
    On Error Resume Next

    wbMatch.Worksheets("Compo") _
        .Shapes("BTN_CREER_FICHIER_MATCH") _
        .Visible = msoFalse

    BoutonIntrouvable = _
        (Err.Number <> 0)

    Err.Clear

    On Error GoTo GestionErreur

    wbMatch.Save

    wbMatch.Activate
    wbMatch.Worksheets("Compo").Activate

    ' Prepare l'initialisation de la copie apres
    ' la fermeture du modele.
    NomClasseurSecurise = _
        Replace(wbMatch.Name, "'", "''")

    NomProcedureInitialisation = _
        "'" & _
        NomClasseurSecurise & _
        "'!InitialiserClasseurStatsRugby"

    ' Arrete les taches du modele avant de transmettre
    ' la main a la copie.
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

    If NumeroErreur = 1004 Then
    
        MsgBox _
            "Le fichier du match n'a pas pu etre cree." & _
            vbCrLf & vbCrLf & _
            "Erreur " & NumeroErreur & " : " & DescriptionErreur & _
            vbCrLf & vbCrLf & _
            "Chemin vise :" & vbCrLf & CheminChoisi, _
            vbExclamation, _
            "Creation du fichier de match"
    
    Else


        MsgBox _
            "Le fichier du match a ete cree :" & _
            vbCrLf & vbCrLf & _
            wbMatch.Name & _
            vbCrLf & vbCrLf & _
            "Tu peux maintenant terminer la composition.", _
            vbInformation, _
            "Fichier du match cree"

    End If

    ' Cette instruction doit rester la derniere :
    ' l'initialisation de la copie est deja planifiee.
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

    If NumeroErreur = 1004 Then
    
        MsgBox _
            "Le fichier du match n'a pas pu etre cree." & _
            vbCrLf & vbCrLf & _
            "Verifie qu'un dossier nomme 'Matchs' existe dans le meme dossier que ce classeur, puis recommence." & _
            vbCrLf & vbCrLf & _
            "Si le dossier a ete supprime, recree simplement un dossier nomme 'Matchs'.", _
            vbExclamation, _
            "Creation du fichier de match"
    
    Else
    
        MsgBox _
            "Le fichier du match n'a pas pu etre cree." & _
            vbCrLf & vbCrLf & _
            "Erreur " & NumeroErreur & _
            " : " & DescriptionErreur, _
            vbCritical, _
            "Erreur de creation"
    
    End If

End Sub


Private Function ConstruireMatchID( _
    ByVal wsCompo As Worksheet _
) As String

    Dim Saison As String
    Dim PhaseCode As String
    Dim Journee As String
    Dim Adversaire As String
    Dim Categorie As String
    Dim LieuCode As String

    Saison = _
        Trim(CStr(wsCompo.Range("MATCH_SAISON").Value))

    Journee = _
        Trim(CStr(wsCompo.Range("MATCH_JOURNEE").Value))

    Adversaire = _
        Trim(CStr(wsCompo.Range("MATCH_ADV").Value))

    Categorie = _
        Trim(CStr(wsCompo.Range("MATCH_CAT").Value))

    Select Case UCase(Trim(CStr(wsCompo.Range("MATCH_PHASE").Value)))
    
        Case "ALLER"
            PhaseCode = "A"
    
        Case "RETOUR"
            PhaseCode = "R"
    
        Case "PHASE FINALE"
            PhaseCode = "PF"
    
        Case Else
            PhaseCode = ""
    
    End Select
    
    Select Case UCase(Trim(CStr(wsCompo.Range("MATCH_LIEU").Value)))
    
        Case "DOMICILE"
            LieuCode = "DOM"
    
        Case "EXTƒRIEUR", "EXTERIEUR"
            LieuCode = "EXT"
    
        Case "TERRAIN NEUTRE"
            LieuCode = "NEU"
    
        Case Else
            LieuCode = ""
    
    End Select

    If PhaseCode <> "" Then

        ConstruireMatchID = _
            Saison & "_" & _
            PhaseCode & Journee & "_" & _
            Adversaire & "_" & _
            Categorie & "_" & _
            LieuCode

    Else

        ConstruireMatchID = _
            Saison & "_" & _
            Journee & "_" & _
            Adversaire & "_" & _
            Categorie & "_" & _
            LieuCode

    End If

End Function


Private Sub AfficherChampMatchManquant( _
    ByVal Cellule As Range, _
    ByVal Message As String _
)

    MsgBox _
        Message, _
        vbExclamation, _
        "Informations du match incompletes"

    Cellule.Worksheet.Activate
    Cellule.Select

End Sub


Private Function NettoyerNomFichier( _
    ByVal Texte As String _
) As String

    Dim CaracteresInterdits As Variant
    Dim Caractere As Variant
    Dim Resultat As String

    Resultat = Trim(Texte)

    ' Retire les principaux accents.
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
        "/", _
        "\", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|" _
    )

    For Each Caractere In CaracteresInterdits

        Resultat = _
            Replace( _
                Resultat, _
                CStr(Caractere), _
                "-" _
            )

    Next Caractere

    Resultat = _
        Replace(Resultat, " ", "-")

    Do While InStr(Resultat, "--") > 0

        Resultat = _
            Replace(Resultat, "--", "-")

    Loop

    NettoyerNomFichier = Resultat

End Function

Sub AutoriserDossierMatchs()

    If DemanderAutorisationDossierMatchs() Then

        MsgBox _
            "Autorisation enregistr" & ChrW(233) & "e pour le dossier Matchs.", _
            vbInformation, _
            "Dossier Matchs"

    End If

End Sub

Private Function DemanderAutorisationDossierMatchs() As Boolean

    Dim Chemin As String
    Dim Reponse As VbMsgBoxResult

#If Mac Then

    Dim DossiersDemandes As Variant

#End If

    Reponse = MsgBox( _
        "macOS doit autoriser Excel " & ChrW(224) & _
        " " & ChrW(233) & "crire dans le dossier Matchs." & _
        vbCrLf & vbCrLf & _
        "S" & ChrW(233) & "lectionne le dossier Matchs " & _
        "dans la fen" & ChrW(234) & "tre suivante.", _
        vbInformation + vbOKCancel, _
        "Autorisation requise" _
    )

    If Reponse <> vbOK Then
        Exit Function
    End If

    Chemin = _
        ChoisirDossier( _
            "Selectionne le dossier Matchs", _
            ThisWorkbook.Path _
        )

    If Len(Chemin) = 0 Then
        Exit Function
    End If

#If Mac Then

    DossiersDemandes = Array(Chemin)

    DemanderAutorisationDossierMatchs = _
        GrantAccessToMultipleFiles(DossiersDemandes)

#Else

    DemanderAutorisationDossierMatchs = True

#End If

End Function


Private Function ChoisirDossier( _
    ByVal Titre As String, _
    ByVal DossierDepart As String _
) As String

#If Mac Then

    Dim ScriptApple As String
    Dim Resultat As String

    ScriptApple = _
        "set Titre to " & _
        Chr(34) & Titre & Chr(34) & vbLf

    If Len(DossierDepart) > 0 Then

        ScriptApple = ScriptApple & _
            "set Cible to POSIX file " & _
            Chr(34) & DossierDepart & Chr(34) & vbLf & _
            "try" & vbLf & _
            "set Dossier to choose folder with prompt Titre " & _
            "default location Cible" & vbLf & _
            "on error Message number Numero" & vbLf & _
            "if Numero is -128 then return " & Chr(34) & Chr(34) & vbLf & _
            "try" & vbLf & _
            "set Dossier to choose folder with prompt Titre" & vbLf & _
            "on error" & vbLf & _
            "return " & Chr(34) & Chr(34) & vbLf & _
            "end try" & vbLf & _
            "end try" & vbLf

    Else

        ScriptApple = ScriptApple & _
            "try" & vbLf & _
            "set Dossier to choose folder with prompt Titre" & vbLf & _
            "on error" & vbLf & _
            "return " & Chr(34) & Chr(34) & vbLf & _
            "end try" & vbLf

    End If

    ScriptApple = ScriptApple & _
        "return POSIX path of Dossier"

    On Error GoTo Annule

    Resultat = MacScript(ScriptApple)

    Resultat = Trim$(Resultat)

    ' POSIX path renvoie un dossier termine par "/".
    Do While Len(Resultat) > 1 And Right$(Resultat, 1) = "/"

        Resultat = Left$(Resultat, Len(Resultat) - 1)

    Loop

    ChoisirDossier = Resultat

    Exit Function

Annule:

    ChoisirDossier = ""

#Else

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = Titre
        .InitialFileName = DossierDepart

        If .Show <> -1 Then
            Exit Function
        End If

        ChoisirDossier = .SelectedItems(1)

    End With

#End If

End Function


