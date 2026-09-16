Attribute VB_Name = "modRecapImport"
Option Explicit

' =========================================================
' RECAPITULATIF DE SAISON - AJOUT DES MATCHS NOUVEAUX
'
' Le classeur garde chez lui le journal d'actions et les
' compositions de tous les matchs deja ajoutes. Le bouton
' d'actualisation n'ouvre donc que les fichiers qu'il ne
' connait pas encore : une fois un match integre, son
' fichier n'est plus jamais relu, et la duree de
' l'actualisation ne depend plus du nombre de matchs de la
' saison mais du nombre de matchs joues depuis la derniere
' fois.
'
' Seuls les fichiers poses a la RACINE du dossier sont
' regardes, et l'emplacement est le seul critere : ce qui
' est range dans un sous-dossier n'entre pas dans la base,
' ce qui est a la racine y entre. Un amical depose a la
' racine est donc traite comme les autres, et la feuille
' Filtres permet ensuite de le comparer au championnat.
'
' Un match saisi en deux fichiers de mi-temps n'a pas sa
' place ici : c'est le fichier fusionne, qui porte le match
' entier, qu'on depose a la racine. Le classement ne
' distingue pas les mi-temps, il totalise le match.
'
' UN FICHIER MODIFIE EST RELU. La date de modification du
' fichier est notee lors de son integration ; a chaque
' actualisation, elle est comparee a celle du fichier, ce
' qui ne demande pas de l'ouvrir. Si elle a bouge, toutes
' les lignes du match sont effacees puis reecrites.
'
' C'est un remplacement, jamais une fusion : on ne cherche
' pas quelles lignes ont ete ajoutees. Le journal n'a pas
' d'identifiant de ligne stable, et une saisie reprise
' comporte aussi des lignes corrigees ou supprimees. Tout
' remplacer est la seule facon de garantir qu'aucune ligne
' ne soit ni doublee ni perdue.
'
' Les fichiers de match ne sont jamais modifies : ils sont
' ouverts en lecture seule et refermes sans enregistrer.
' =========================================================

Private Type InfosMatch

    Fichier As String

    IdMatch As String
    Saison As String
    DateMatch As Double
    Categorie As String
    Adversaire As String
    Lieu As String
    Phase As String
    Journee As String
    Resultat As String

    NbActions As Long
    NbJoueurs As Long
    NbMinutes As Long

    Lu As Boolean

    Actions As Variant
    Compositions As Variant

End Type


' =========================================================
' Bouton "Ajouter les nouveaux matchs".
' =========================================================
Public Sub AjouterNouveauxMatchs()

    Dim Fichiers As Collection
    Dim Fiches As Collection

    Dim Match As InfosMatch
    Dim ALire As Collection
    Dim DejaEnBase As Collection

    Dim Fichier As String

    Dim NbAjoutes As Long
    Dim NbRelus As Long
    Dim NbIllisibles As Long
    Dim NbDoublons As Long

    Dim EtatAffichage As Boolean
    Dim EtatEvenements As Boolean
    Dim EtatAlertes As Boolean
    Dim EtatCalcul As XlCalculation

    Dim i As Long

    If Len(ThisWorkbook.Path) = 0 Then

        MsgBox _
            "Enregistre d'abord ce classeur dans le " & _
            "dossier Matchs, " & ChrW(224) & " c" & ChrW(244) & _
            "t" & ChrW(233) & " des fichiers de championnat.", _
            vbExclamation, _
            "Classeur non enregistr" & ChrW(233)

        Exit Sub

    End If

    Tracer "--- Ajout des nouveaux matchs ---"

    Set Fichiers = FichiersDeLaRacine()

    Tracer "Fichiers trouves a la racine : " & Fichiers.Count

    If Fichiers.Count = 0 Then

        MsgBox _
            "Aucun fichier de match " & ChrW(224) & " la racine de" & _
            vbCrLf & ThisWorkbook.Path & vbCrLf & vbCrLf & _
            "Les fichiers de championnat doivent " & ChrW(234) & _
            "tre pos" & ChrW(233) & "s directement dans ce " & _
            "dossier, sans sous-dossier.", _
            vbExclamation, _
            "Aucun match"

        Exit Sub

    End If

    Set Fiches = FichesParFichier()

    Tracer "Matchs deja en base : " & Fiches.Count

    ' Un fichier est a lire s'il est inconnu, ou si sa date
    ' de modification a bouge depuis son integration.
    Set ALire = New Collection
    Set DejaEnBase = New Collection

    For i = 1 To Fichiers.Count

        Fichier = CStr(Fichiers.Item(i))

        If FichierAIntegrer(Fiches, Fichier) Then

            ALire.Add Fichier

            If CleExiste(Fiches, Normaliser(Fichier)) Then
                DejaEnBase.Add True, Normaliser(Fichier)
            End If

        End If

    Next i

    If ALire.Count = 0 Then

        AppliquerFiltres

        MsgBox _
            "La base est " & ChrW(224) & " jour : les " & _
            Fichiers.Count & " fichier(s) de la racine sont " & _
            "int" & ChrW(233) & "gr" & ChrW(233) & "s, et aucun " & _
            "n'a " & ChrW(233) & "t" & ChrW(233) & " modifi" & _
            ChrW(233) & " depuis.", _
            vbInformation, _
            "Rien " & ChrW(224) & " ajouter"

        Exit Sub

    End If

    Tracer "A lire : " & ALire.Count

    Tracer "Demande d'autorisation macOS"

    If Not AutoriserLecture(ALire) Then
        Tracer "Autorisation refusee"
        Exit Sub
    End If

    Tracer "Autorisation accordee"

    EtatAffichage = Application.ScreenUpdating
    EtatEvenements = Application.EnableEvents
    EtatAlertes = Application.DisplayAlerts
    EtatCalcul = Application.Calculation

    On Error GoTo GestionErreur

    ' EnableEvents = False empeche le Workbook_Open des
    ' fichiers de match de lancer VeoVideoControl. Le
    ' calcul manuel evite qu'Excel recalcule les stats de
    ' chaque fichier ouvert, ce qui serait tres long.
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    For i = 1 To ALire.Count

        Fichier = CStr(ALire.Item(i))

        Tracer "Ouverture : " & Fichier

        LireFichierMatch Fichier, Match

        Tracer "  lu=" & Match.Lu & " actions=" & Match.NbActions & _
            " joueurs=" & Match.NbJoueurs

        If Not Match.Lu Then

            NbIllisibles = NbIllisibles + 1

        ElseIf CleExiste(DejaEnBase, Normaliser(Fichier)) Then

            ' Fichier deja integre mais modifie depuis : ses
            ' lignes sont effacees avant d'etre reecrites,
            ' donc aucune n'est doublee.
            Tracer "  remplacement des lignes"

            EffacerMatchDeLaBase Fichier

            AjouterMatchEnBase Match, Fichier

            NbRelus = NbRelus + 1

        ElseIf IdDejaEnBase(Match.IdMatch) Then

            ' Deux fichiers differents pour le meme match :
            ' le second doublerait chaque statistique.
            NbDoublons = NbDoublons + 1

        Else

            Tracer "  ecriture en base"

            AjouterMatchEnBase Match, Fichier

            NbAjoutes = NbAjoutes + 1

        End If

    Next i

    Tracer "Lecture terminee, construction des listes"

    ConstruireListesFiltres

    Application.Calculation = EtatCalcul
    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    Tracer "Calcul du classement"

    AppliquerFiltres

    Tracer "Termine"

    MsgBox _
        BilanAjout( _
            NbAjoutes, NbRelus, NbDoublons, NbIllisibles _
        ), _
        vbInformation, _
        "Actualisation termin" & ChrW(233) & "e"

    Exit Sub

GestionErreur:

    Dim DescriptionErreur As String

    DescriptionErreur = Err.Description

    On Error Resume Next

    Application.Calculation = EtatCalcul
    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    On Error GoTo 0

    MsgBox _
        "L'ajout a " & ChrW(233) & "chou" & ChrW(233) & "." & _
        vbCrLf & vbCrLf & DescriptionErreur & vbCrLf & vbCrLf & _
        "Les matchs d" & ChrW(233) & "j" & ChrW(224) & " ajout" & _
        ChrW(233) & "s sont intacts.", _
        vbCritical, _
        "Erreur"

End Sub


Private Function BilanAjout( _
    ByVal NbAjoutes As Long, _
    ByVal NbRelus As Long, _
    ByVal NbDoublons As Long, _
    ByVal NbIllisibles As Long _
) As String

    Dim Texte As String

    Texte = NbAjoutes & " match(s) ajout" & ChrW(233) & "(s) " & _
        ChrW(224) & " la base."

    If NbRelus > 0 Then

        Texte = Texte & vbCrLf & vbCrLf & _
            NbRelus & " match(s) relu(s) : leur fichier " & _
            "avait " & ChrW(233) & "t" & ChrW(233) & " modifi" & _
            ChrW(233) & " depuis." & vbCrLf & _
            "Leurs anciennes lignes ont " & ChrW(233) & "t" & _
            ChrW(233) & " remplac" & ChrW(233) & "es, pas " & _
            "compl" & ChrW(233) & "t" & ChrW(233) & "es."

    End If

    If NbDoublons > 0 Then

        Texte = Texte & vbCrLf & vbCrLf & _
            NbDoublons & " fichier(s) portant un match " & _
            "d" & ChrW(233) & "j" & ChrW(224) & " en base : ignor" & _
            ChrW(233) & "(s) pour ne pas doubler les " & _
            "statistiques."

    End If

    If NbIllisibles > 0 Then

        Texte = Texte & vbCrLf & vbCrLf & _
            NbIllisibles & " fichier(s) illisible(s) : " & _
            "v" & ChrW(233) & "rifie qu'il s'agit bien de " & _
            "fichiers de match."

    End If

    BilanAjout = Texte

End Function


' =========================================================
' Fichiers de match poses a la racine du dossier. Le
' balayage ne descend dans aucun sous-dossier : c'est ce
' qui met les amicaux hors perimetre.
' =========================================================
Private Function FichiersDeLaRacine() As Collection

    Dim Fichiers As Collection
    Dim Fichier As String

    Set Fichiers = New Collection

    Fichier = _
        Dir(ThisWorkbook.Path & Application.PathSeparator & "*.xls*")

    Do While Len(Fichier) > 0

        If EstFichierMatch(Fichier) Then
            Fichiers.Add Fichier
        End If

        Fichier = Dir

    Loop

    Set FichiersDeLaRacine = Fichiers

End Function


' =========================================================
' Ecarte le classeur recapitulatif lui-meme, les fichiers
' de travail laisses par Excel et tout ce qui n'est pas un
' classeur.
' =========================================================
Private Function EstFichierMatch( _
    ByVal Fichier As String _
) As Boolean

    Dim Extension As String

    If Left$(Fichier, 2) = "~$" Then
        Exit Function
    End If

    If Left$(Fichier, 1) = "." Then
        Exit Function
    End If

    If StrComp(Fichier, ThisWorkbook.Name, vbTextCompare) = 0 Then
        Exit Function
    End If

    Extension = LCase$(Mid$(Fichier, InStrRev(Fichier, ".") + 1))

    EstFichierMatch = _
        (Extension = "xlsm" Or Extension = "xlsx")

End Function


' =========================================================
' Fichiers deja integres, avec la date de modification
' relevee lors de leur lecture.
'
' L'inventaire de la feuille Matchs fait office de
' memoire : un match dont le journal serait vide y figure
' quand meme, et ne sera donc pas rouvert a chaque
' actualisation.
' =========================================================
Private Function FichesParFichier() As Collection

    Dim Fiches As Collection
    Dim Donnees As Variant
    Dim Cle As String
    Dim Ligne As Long

    Set Fiches = New Collection

    Donnees = DonneesTableau(FEUILLE_MATCHS, TBL_MATCHS)

    If IsEmpty(Donnees) Then
        Set FichesParFichier = Fiches
        Exit Function
    End If

    For Ligne = 1 To UBound(Donnees, 1)

        Cle = Normaliser(TexteCellule(Donnees(Ligne, MTC_FICHIER)))

        If Len(Cle) > 0 Then

            If Not CleExiste(Fiches, Cle) Then

                Fiches.Add _
                    NombreCellule(Donnees(Ligne, MTC_MODIFIE)), _
                    Cle

            End If

        End If

    Next Ligne

    Set FichesParFichier = Fiches

End Function


' =========================================================
' Vrai si le fichier doit etre lu : soit il est inconnu,
' soit il a ete modifie depuis son integration.
'
' La date de modification se lit sans ouvrir le fichier :
' c'est ce qui permet de passer en revue toute une saison
' en un instant.
'
' Une fiche sans date notee est celle d'un match integre
' avant que ce controle existe. Elle est relue une fois,
' pour prendre sa date au passage.
' =========================================================
Private Function FichierAIntegrer( _
    ByVal Fiches As Collection, _
    ByVal Fichier As String _
) As Boolean

    Dim Cle As String
    Dim DateNotee As Double
    Dim DateReelle As Double

    Cle = Normaliser(Fichier)

    If Not CleExiste(Fiches, Cle) Then
        FichierAIntegrer = True
        Exit Function
    End If

    DateNotee = CDbl(Fiches.Item(Cle))

    If DateNotee <= 0 Then
        FichierAIntegrer = True
        Exit Function
    End If

    DateReelle = DateModificationFichier(Fichier)

    If DateReelle <= 0 Then
        Exit Function
    End If

    FichierAIntegrer = _
        (DateReelle > DateNotee + TOLERANCE_DATE)

End Function


' =========================================================
' Date de derniere modification d'un fichier du dossier, ou
' zero si elle n'a pas pu etre lue.
' =========================================================
Private Function DateModificationFichier( _
    ByVal Fichier As String _
) As Double

    On Error Resume Next

    DateModificationFichier = _
        CDbl(FileDateTime(CheminComplet(Fichier)))

    If Err.Number <> 0 Then
        DateModificationFichier = 0
        Err.Clear
    End If

    On Error GoTo 0

End Function


' =========================================================
' Vrai si un match portant cet identifiant est deja en
' base sous un autre nom de fichier.
' =========================================================
Private Function IdDejaEnBase( _
    ByVal IdMatch As String _
) As Boolean

    Dim Donnees As Variant
    Dim Cle As String
    Dim Ligne As Long

    Cle = Normaliser(IdMatch)

    If Len(Cle) = 0 Then
        Exit Function
    End If

    Donnees = DonneesTableau(FEUILLE_MATCHS, TBL_MATCHS)

    If IsEmpty(Donnees) Then
        Exit Function
    End If

    For Ligne = 1 To UBound(Donnees, 1)

        If Normaliser(TexteCellule(Donnees(Ligne, MTC_ID))) = Cle Then
            IdDejaEnBase = True
            Exit Function
        End If

    Next Ligne

End Function


' =========================================================
' macOS demande une autorisation explicite avant qu'Excel
' n'ouvre des fichiers que l'utilisateur n'a pas designes
' lui-meme. Elle est demandee une fois pour tous les
' fichiers a lire.
' =========================================================
Private Function AutoriserLecture( _
    ByVal Fichiers As Collection _
) As Boolean

#If Mac Then

    ' Le tableau doit etre un Variant contenant des
    ' Variants, comme ce que renvoie Array() : c'est la
    ' forme qu'attend cette fonction native de la sandbox
    ' macOS, et celle qu'emploie deja le classeur de
    ' creation des matchs. Lui passer un tableau de String
    ' typees ne provoque pas une erreur VBA mais un arret
    ' brutal d'Excel.
    Dim FichiersDemandes As Variant
    Dim Liste() As Variant
    Dim i As Long

    If Fichiers.Count = 0 Then
        AutoriserLecture = True
        Exit Function
    End If

    ReDim Liste(0 To Fichiers.Count - 1)

    For i = 1 To Fichiers.Count
        Liste(i - 1) = CheminComplet(CStr(Fichiers.Item(i)))
    Next i

    FichiersDemandes = Liste

    AutoriserLecture = _
        GrantAccessToMultipleFiles(FichiersDemandes)

    If Not AutoriserLecture Then

        MsgBox _
            "macOS n'a pas autoris" & ChrW(233) & " la lecture " & _
            "des fichiers de match." & vbCrLf & vbCrLf & _
            "Relance l'actualisation et accepte la " & _
            "demande d'acc" & ChrW(232) & "s.", _
            vbExclamation, _
            "Autorisation refus" & ChrW(233) & "e"

    End If

#Else

    AutoriserLecture = True

#End If

End Function


Private Function CheminComplet( _
    ByVal Fichier As String _
) As String

    CheminComplet = _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        Fichier

End Function


' =========================================================
' Ouvre un fichier de match et en extrait le journal et la
' composition. Un fichier illisible est compte dans le
' bilan, mais n'interrompt pas l'ajout des autres.
' =========================================================
Private Sub LireFichierMatch( _
    ByVal Fichier As String, _
    ByRef Match As InfosMatch _
)

    Dim wb As Workbook
    Dim Vierge As InfosMatch

    Match = Vierge
    Match.Fichier = Fichier

    On Error GoTo FichierIllisible

    Set wb = _
        Workbooks.Open( _
            Filename:=CheminComplet(Fichier), _
            UpdateLinks:=0, _
            ReadOnly:=True _
        )

    ' L'en-tete est d'abord lue dans la feuille Compo :
    ' elle y figure meme quand la saisie video n'a pas
    ' commence. Le journal, plus loin, la confirmera.
    Match.Categorie = LibelleDepuisCompo(wb, "categorie")
    Match.Phase = LibelleDepuisCompo(wb, "phase")
    Match.IdMatch = LibelleDepuisCompo(wb, "idmatch")

    LireJournal wb, Match

    LireComposition wb, Match

    wb.Close SaveChanges:=False

    Set wb = Nothing

    Match.Lu = True

    Exit Sub

FichierIllisible:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    On Error GoTo 0

    Match.Lu = False

End Sub


' =========================================================
' Valeur d'un champ de l'en-tete du match, lue dans la
' feuille Compo : le libelle est cherche en colonne B, la
' valeur est prise en colonne C.
'
' Le reperage se fait sur le libelle normalise, et non sur
' une adresse : la feuille Compo a deja change de
' disposition d'une saison a l'autre.
'
' La categorie de l'equipe est le seul champ que le journal
' ne recopie pas. La phase et l'identifiant y sont, mais ne
' servent ici qu'a reconnaitre un match dont la saisie
' video n'a pas encore commence : sans cela, un amical au
' journal vide passerait le filtre.
' =========================================================
Private Function LibelleDepuisCompo( _
    ByVal wb As Workbook, _
    ByVal ChampNormalise As String _
) As String

    Dim ws As Worksheet
    Dim Ligne As Long

    On Error Resume Next
    Set ws = wb.Worksheets(SRC_FEUILLE_COMPO)
    On Error GoTo 0

    If ws Is Nothing Then
        Exit Function
    End If

    ' L'identifiant du match est ecrit bien plus bas que
    ' l'en-tete, sous la composition.
    For Ligne = 4 To 70

        If Normaliser(TexteCellule(ws.Cells(Ligne, 2).Value)) _
            = ChampNormalise Then

            LibelleDepuisCompo = _
                TexteCellule(ws.Cells(Ligne, 3).Value)

            Exit Function

        End If

    Next Ligne

    If ChampNormalise = "categorie" Then

        LibelleDepuisCompo = _
            TexteCellule(ws.Range(SRC_CELLULE_CATEGORIE).Value)

    End If

End Function


' =========================================================
' Recopie le journal d'actions dans la disposition des
' colonnes de ce classeur.
'
' La colonne Mi-temps est conservee telle quelle : elle ne
' sert pas au classement, qui totalise le match, mais elle
' reste disponible pour qui voudrait monter un tableau
' croise sur la feuille Journal actions.
' =========================================================
Private Sub LireJournal( _
    ByVal wb As Workbook, _
    ByRef Match As InfosMatch _
)

    Dim Tableau As ListObject
    Dim Source As Variant
    Dim Cible As Variant

    Dim ColTemps As Long
    Dim ColMiTemps As Long
    Dim ColJoueur As Long
    Dim ColGroupe As Long
    Dim ColAction As Long
    Dim ColMotif As Long
    Dim ColPossession As Long
    Dim ColId As Long
    Dim ColSaison As Long
    Dim ColDate As Long
    Dim ColAdversaire As Long
    Dim ColLieu As Long
    Dim ColPhase As Long
    Dim ColJournee As Long
    Dim ColResultat As Long

    Dim NbLignes As Long
    Dim Ligne As Long
    Dim Ecrites As Long

    On Error Resume Next

    Set Tableau = _
        wb.Worksheets(SRC_FEUILLE_JOURNAL) _
            .ListObjects(SRC_TBL_JOURNAL)

    On Error GoTo 0

    If Tableau Is Nothing Then
        Exit Sub
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Sub
    End If

    ColTemps = IndexColonneSouple(Tableau, "temps video")
    ColMiTemps = IndexColonneSouple(Tableau, "Mi-temps")
    ColJoueur = IndexColonneSouple(Tableau, "Joueur")
    ColGroupe = IndexColonneSouple(Tableau, "Groupe fautif")
    ColAction = IndexColonneSouple(Tableau, "Action")
    ColMotif = IndexColonneSouple(Tableau, "Motif penalite")
    ColPossession = IndexColonneSouple(Tableau, "Possession")
    ColId = IndexColonneSouple(Tableau, "ID match")
    ColSaison = IndexColonneSouple(Tableau, "Saison")
    ColDate = IndexColonneSouple(Tableau, "Date match")
    ColAdversaire = IndexColonneSouple(Tableau, "Adversaire")
    ColLieu = IndexColonneSouple(Tableau, "Lieu")
    ColPhase = IndexColonneSouple(Tableau, "Phase")
    ColJournee = IndexColonneSouple(Tableau, "Journee")
    ColResultat = IndexColonneSouple(Tableau, "Resultat")

    If ColAction = 0 Then
        Exit Sub
    End If

    Source = Tableau.DataBodyRange.Value

    NbLignes = UBound(Source, 1)

    ReDim Cible(1 To NbLignes, 1 To NB_COLONNES_JOURNAL)

    For Ligne = 1 To NbLignes

        ' Le tableau du journal est dimensionne large : ses
        ' dernieres lignes restent vides tant que la saisie
        ' n'a pas rempli le match.
        If Len(ValeurTexte(Source, Ligne, ColAction)) > 0 Then

            Ecrites = Ecrites + 1

            Cible(Ecrites, 1) = ValeurTexte(Source, Ligne, ColId)
            Cible(Ecrites, 2) = Match.Fichier
            Cible(Ecrites, 3) = ValeurTexte(Source, Ligne, ColSaison)
            Cible(Ecrites, 4) = ValeurBrute(Source, Ligne, ColDate)
            Cible(Ecrites, 5) = Match.Categorie
            Cible(Ecrites, 6) = ValeurTexte(Source, Ligne, ColAdversaire)
            Cible(Ecrites, 7) = ValeurTexte(Source, Ligne, ColLieu)
            Cible(Ecrites, 8) = ValeurTexte(Source, Ligne, ColPhase)
            Cible(Ecrites, 9) = ValeurTexte(Source, Ligne, ColJournee)
            Cible(Ecrites, 10) = ValeurTexte(Source, Ligne, ColResultat)
            Cible(Ecrites, 11) = ValeurTexte(Source, Ligne, ColMiTemps)
            Cible(Ecrites, 12) = ValeurTexte(Source, Ligne, ColJoueur)
            Cible(Ecrites, 13) = ValeurTexte(Source, Ligne, ColGroupe)
            Cible(Ecrites, 14) = ValeurTexte(Source, Ligne, ColAction)
            Cible(Ecrites, 15) = ValeurTexte(Source, Ligne, ColMotif)
            Cible(Ecrites, 16) = ValeurTexte(Source, Ligne, ColPossession)
            Cible(Ecrites, 17) = ValeurBrute(Source, Ligne, ColTemps)

            If Ecrites = 1 Then
                RetenirEnteteMatch Match, Cible
            End If

        End If

    Next Ligne

    Match.NbActions = Ecrites

    If Ecrites > 0 Then
        Match.Actions = Cible
    End If

End Sub


' =========================================================
' L'en-tete du match est lue sur la premiere ligne utile du
' journal, ou toutes ces informations sont deja recopiees.
'
' Un champ vide dans le journal ne remplace pas ce que la
' feuille Compo avait deja donne.
' =========================================================
Private Sub RetenirEnteteMatch( _
    ByRef Match As InfosMatch, _
    ByRef Cible As Variant _
)

    RemplacerSiRenseigne Match.IdMatch, Cible(1, 1)
    RemplacerSiRenseigne Match.Saison, Cible(1, 3)
    RemplacerSiRenseigne Match.Adversaire, Cible(1, 6)
    RemplacerSiRenseigne Match.Lieu, Cible(1, 7)
    RemplacerSiRenseigne Match.Phase, Cible(1, 8)
    RemplacerSiRenseigne Match.Journee, Cible(1, 9)
    RemplacerSiRenseigne Match.Resultat, Cible(1, 10)

    Match.DateMatch = NombreCellule(Cible(1, 4))

End Sub


Private Sub RemplacerSiRenseigne( _
    ByRef Champ As String, _
    ByVal Valeur As Variant _
)

    Dim Texte As String

    Texte = TexteCellule(Valeur)

    If Len(Texte) > 0 Then
        Champ = Texte
    End If

End Sub


' =========================================================
' Recopie la composition : un joueur par ligne, avec le
' temps de jeu saisi par l'entraineur dans le fichier de
' match.
' =========================================================
Private Sub LireComposition( _
    ByVal wb As Workbook, _
    ByRef Match As InfosMatch _
)

    Dim Tableau As ListObject
    Dim Source As Variant
    Dim Cible As Variant

    Dim ColPoste As Long
    Dim ColNom As Long
    Dim ColLigne As Long
    Dim ColTemps As Long

    Dim NbLignes As Long
    Dim Ligne As Long
    Dim Ecrites As Long
    Dim Nom As String
    Dim Minutes As Double

    On Error Resume Next

    Set Tableau = _
        wb.Worksheets(SRC_FEUILLE_COMPO) _
            .ListObjects(SRC_TBL_COMPO)

    On Error GoTo 0

    If Tableau Is Nothing Then
        Exit Sub
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Sub
    End If

    ColPoste = IndexColonneSouple(Tableau, "Poste")
    ColNom = IndexColonneSouple(Tableau, "Nom du joueur")
    ColLigne = IndexColonneSouple(Tableau, "Avant ou 3/4")
    ColTemps = IndexColonneSouple(Tableau, "temps de jeu")

    If ColNom = 0 Then
        Exit Sub
    End If

    Source = Tableau.DataBodyRange.Value

    NbLignes = UBound(Source, 1)

    ReDim Cible(1 To NbLignes, 1 To NB_COLONNES_COMPOSITIONS)

    For Ligne = 1 To NbLignes

        Nom = ValeurTexte(Source, Ligne, ColNom)

        ' Les lignes de remplacants restees vides et les
        ' entrees collectives du journal ne sont pas des
        ' joueurs.
        If EstJoueur(Nom) Then

            Ecrites = Ecrites + 1

            Minutes = _
                MinutesDepuisCompo( _
                    ValeurBrute(Source, Ligne, ColTemps) _
                )

            If Minutes >= 0 Then
                Match.NbMinutes = Match.NbMinutes + 1
            End If

            Cible(Ecrites, 1) = Match.IdMatch
            Cible(Ecrites, 2) = Match.Fichier
            Cible(Ecrites, 3) = Match.Saison
            Cible(Ecrites, 4) = Match.DateMatch
            Cible(Ecrites, 5) = Match.Categorie
            Cible(Ecrites, 6) = Match.Adversaire
            Cible(Ecrites, 7) = Match.Lieu
            Cible(Ecrites, 8) = Match.Phase
            Cible(Ecrites, 9) = Match.Journee
            Cible(Ecrites, 10) = Match.Resultat
            Cible(Ecrites, 11) = Nom
            Cible(Ecrites, 12) = ValeurBrute(Source, Ligne, ColPoste)
            Cible(Ecrites, 13) = ValeurTexte(Source, Ligne, ColLigne)

            If Minutes >= 0 Then
                Cible(Ecrites, 14) = Minutes
            End If

        End If

    Next Ligne

    Match.NbJoueurs = Ecrites

    If Ecrites > 0 Then
        Match.Compositions = Cible
    End If

End Sub


Private Function ValeurTexte( _
    ByRef Source As Variant, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long _
) As String

    If Colonne = 0 Then
        Exit Function
    End If

    ValeurTexte = TexteCellule(Source(Ligne, Colonne))

End Function


Private Function ValeurBrute( _
    ByRef Source As Variant, _
    ByVal Ligne As Long, _
    ByVal Colonne As Long _
) As Variant

    If Colonne = 0 Then
        Exit Function
    End If

    If IsError(Source(Ligne, Colonne)) Then
        Exit Function
    End If

    ValeurBrute = Source(Ligne, Colonne)

End Function


' =========================================================
' Ajoute un match lu aux trois tableaux du classeur.
' =========================================================
Private Sub AjouterMatchEnBase( _
    ByRef Match As InfosMatch, _
    ByVal Fichier As String _
)

    Dim Fiche As Variant
    Dim DateModifiee As Double

    AjouterLignesTableau _
        TableauRecap(FEUILLE_JOURNAL, TBL_JOURNAL), _
        Match.Actions, _
        Match.NbActions, _
        NB_COLONNES_JOURNAL

    AjouterLignesTableau _
        TableauRecap(FEUILLE_COMPOSITIONS, TBL_COMPOSITIONS), _
        Match.Compositions, _
        Match.NbJoueurs, _
        NB_COLONNES_COMPOSITIONS

    ReDim Fiche(1 To 1, 1 To NB_COLONNES_MATCHS)

    Fiche(1, MTC_ID) = Match.IdMatch
    Fiche(1, MTC_FICHIER) = Match.Fichier
    Fiche(1, MTC_SAISON) = Match.Saison

    If Match.DateMatch > 0 Then
        Fiche(1, MTC_DATE) = Match.DateMatch
    End If

    Fiche(1, MTC_CATEGORIE) = Match.Categorie
    Fiche(1, MTC_ADVERSAIRE) = Match.Adversaire
    Fiche(1, MTC_LIEU) = Match.Lieu
    Fiche(1, MTC_PHASE) = Match.Phase
    Fiche(1, MTC_JOURNEE) = Match.Journee
    Fiche(1, MTC_RESULTAT) = Match.Resultat
    Fiche(1, MTC_ACTIONS) = Match.NbActions
    Fiche(1, MTC_JOUEURS) = Match.NbJoueurs
    Fiche(1, MTC_MINUTES) = Match.NbMinutes
    Fiche(1, MTC_AJOUTE) = Now

    ' La date est relevee apres la lecture : si le fichier
    ' venait d'etre enregistre, c'est bien la version qu'on
    ' a lue qui est notee.
    DateModifiee = DateModificationFichier(Fichier)

    If DateModifiee > 0 Then
        Fiche(1, MTC_MODIFIE) = DateModifiee
    End If

    AjouterLignesTableau _
        TableauRecap(FEUILLE_MATCHS, TBL_MATCHS), _
        Fiche, _
        1, _
        NB_COLONNES_MATCHS

End Sub


' =========================================================
' BOUTONS DE LA FEUILLE MATCHS
' =========================================================

' =========================================================
' Retire de la base le match de la ligne selectionnee.
' =========================================================
Public Sub RetirerMatch()

    Dim Fichier As String

    Fichier = FichierDeLaLigneSelectionnee()

    If Len(Fichier) = 0 Then
        Exit Sub
    End If

    If MsgBox( _
        "Retirer de la base le match du fichier" & vbCrLf & _
        Fichier & " ?" & vbCrLf & vbCrLf & _
        "Ses actions et sa composition seront " & _
        "effac" & ChrW(233) & "es. Le fichier de match, lui, " & _
        "n'est pas touch" & ChrW(233) & ".", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Retirer un match" _
    ) <> vbYes Then

        Exit Sub

    End If

    EffacerMatchDeLaBase Fichier

    ConstruireListesFiltres

    AppliquerFiltres

    MsgBox _
        "Match retir" & ChrW(233) & " de la base." & vbCrLf & vbCrLf & _
        "Il sera r" & ChrW(233) & "int" & ChrW(233) & "gr" & ChrW(233) & _
        " " & ChrW(224) & " la prochaine actualisation si son " & _
        "fichier est toujours " & ChrW(224) & " la racine.", _
        vbInformation, _
        "Retrait termin" & ChrW(233)

End Sub


' =========================================================
' Relit le fichier du match selectionne : c'est ce qu'il
' faut faire quand la saisie d'un match deja ajoute a ete
' completee ou corrigee.
' =========================================================
Public Sub ReimporterMatch()

    Dim Fichier As String
    Dim Match As InfosMatch
    Dim Fichiers As Collection

    Dim EtatAffichage As Boolean
    Dim EtatEvenements As Boolean
    Dim EtatAlertes As Boolean
    Dim EtatCalcul As XlCalculation

    Fichier = FichierDeLaLigneSelectionnee()

    If Len(Fichier) = 0 Then
        Exit Sub
    End If

    If Dir(CheminComplet(Fichier)) = "" Then

        MsgBox _
            "Le fichier" & vbCrLf & Fichier & vbCrLf & _
            "n'est plus " & ChrW(224) & " la racine du dossier." & _
            vbCrLf & vbCrLf & _
            "Utilise " & ChrW(171) & " Retirer ce match " & _
            ChrW(187) & " si le match ne doit plus compter.", _
            vbExclamation, _
            "Fichier introuvable"

        Exit Sub

    End If

    Set Fichiers = New Collection
    Fichiers.Add Fichier

    If Not AutoriserLecture(Fichiers) Then
        Exit Sub
    End If

    EtatAffichage = Application.ScreenUpdating
    EtatEvenements = Application.EnableEvents
    EtatAlertes = Application.DisplayAlerts
    EtatCalcul = Application.Calculation

    On Error GoTo GestionErreur

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    LireFichierMatch Fichier, Match

    If Match.Lu Then

        ' Le retrait n'a lieu qu'une fois la relecture
        ' reussie : un fichier devenu illisible ne doit pas
        ' faire perdre ce qui etait deja en base.
        EffacerMatchDeLaBase Fichier

        AjouterMatchEnBase Match, Fichier

    End If

    ConstruireListesFiltres

    Application.Calculation = EtatCalcul
    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    AppliquerFiltres

    If Match.Lu Then

        MsgBox _
            "Match relu : " & Match.NbActions & " action(s), " & _
            Match.NbJoueurs & " joueur(s), " & _
            Match.NbMinutes & " temps de jeu saisi(s).", _
            vbInformation, _
            "R" & ChrW(233) & "import termin" & ChrW(233)

    Else

        MsgBox _
            "Le fichier n'a pas pu " & ChrW(234) & "tre lu." & _
            vbCrLf & vbCrLf & _
            "Ce qui " & ChrW(233) & "tait en base a " & ChrW(233) & _
            "t" & ChrW(233) & " conserv" & ChrW(233) & ".", _
            vbExclamation, _
            "Fichier illisible"

    End If

    Exit Sub

GestionErreur:

    Dim DescriptionErreur As String

    DescriptionErreur = Err.Description

    On Error Resume Next

    Application.Calculation = EtatCalcul
    Application.DisplayAlerts = EtatAlertes
    Application.ScreenUpdating = EtatAffichage
    Application.EnableEvents = EtatEvenements

    On Error GoTo 0

    MsgBox _
        "Le r" & ChrW(233) & "import a " & ChrW(233) & "chou" & _
        ChrW(233) & "." & vbCrLf & vbCrLf & DescriptionErreur, _
        vbCritical, _
        "Erreur"

End Sub


' =========================================================
' Nom du fichier porte par la ligne ou se trouve le
' curseur, dans l'inventaire des matchs.
' =========================================================
Private Function FichierDeLaLigneSelectionnee() As String

    Dim Tableau As ListObject
    Dim ws As Worksheet
    Dim Ligne As Long

    Set Tableau = TableauRecap(FEUILLE_MATCHS, TBL_MATCHS)

    If Tableau Is Nothing Then
        Exit Function
    End If

    If Tableau.ListRows.Count = 0 Then

        MsgBox _
            "La base ne contient encore aucun match.", _
            vbInformation, _
            "Base vide"

        Exit Function

    End If

    Set ws = Tableau.Parent

    If Not ActiveSheet Is ws Then
        ws.Activate
    End If

    Ligne = ActiveCell.Row

    If Ligne < Tableau.DataBodyRange.Row _
        Or Ligne > Tableau.DataBodyRange.Row _
            + Tableau.ListRows.Count - 1 Then

        MsgBox _
            "Clique d'abord sur la ligne du match, dans " & _
            "le tableau ci-dessous.", _
            vbExclamation, _
            "Aucun match s" & ChrW(233) & "lectionn" & ChrW(233)

        Exit Function

    End If

    FichierDeLaLigneSelectionnee = _
        TexteCellule( _
            ws.Cells( _
                Ligne, _
                Tableau.Range.Column + MTC_FICHIER - 1 _
            ).Value _
        )

End Function


' =========================================================
' Efface d'un tableau toutes les lignes d'un match.
'
' Les lignes ne sont pas supprimees une par une : on
' recopie celles qu'on garde, on vide, puis on reecrit d'un
' bloc. Sur plusieurs milliers de lignes, la difference se
' compte en minutes.
' =========================================================
Private Sub EffacerMatchDeLaBase( _
    ByVal Fichier As String _
)

    EffacerLignesDuFichier _
        TableauRecap(FEUILLE_JOURNAL, TBL_JOURNAL), _
        Fichier, JRN_FICHIER, NB_COLONNES_JOURNAL

    EffacerLignesDuFichier _
        TableauRecap(FEUILLE_COMPOSITIONS, TBL_COMPOSITIONS), _
        Fichier, CMP_FICHIER, NB_COLONNES_COMPOSITIONS

    EffacerLignesDuFichier _
        TableauRecap(FEUILLE_MATCHS, TBL_MATCHS), _
        Fichier, MTC_FICHIER, NB_COLONNES_MATCHS

End Sub


Private Sub EffacerLignesDuFichier( _
    ByVal Tableau As ListObject, _
    ByVal Fichier As String, _
    ByVal ColonneFichier As Long, _
    ByVal NbColonnes As Long _
)

    Dim Donnees As Variant
    Dim Gardees As Variant
    Dim Cle As String
    Dim NbGardees As Long
    Dim Ligne As Long
    Dim Colonne As Long

    If Tableau Is Nothing Then
        Exit Sub
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Sub
    End If

    Donnees = Tableau.DataBodyRange.Value

    Cle = Normaliser(Fichier)

    ReDim Gardees(1 To UBound(Donnees, 1), 1 To NbColonnes)

    For Ligne = 1 To UBound(Donnees, 1)

        If Normaliser( _
            TexteCellule(Donnees(Ligne, ColonneFichier)) _
        ) <> Cle Then

            NbGardees = NbGardees + 1

            For Colonne = 1 To NbColonnes
                Gardees(NbGardees, Colonne) = Donnees(Ligne, Colonne)
            Next Colonne

        End If

    Next Ligne

    If NbGardees = UBound(Donnees, 1) Then
        Exit Sub
    End If

    ViderTableauRecap Tableau

    AjouterLignesTableau Tableau, Gardees, NbGardees, NbColonnes

End Sub


' =========================================================
' Reconstruit les listes qui alimentent les menus de la
' feuille Filtres, pour ne proposer que des valeurs
' presentes dans les matchs de la base.
' =========================================================
Public Sub ConstruireListesFiltres()

    Dim ws As Worksheet
    Dim Donnees As Variant

    Set ws = FeuilleRecap(FEUILLE_PARAMETRES)

    If ws Is Nothing Then
        Exit Sub
    End If

    Donnees = DonneesTableau(FEUILLE_MATCHS, TBL_MATCHS)

    EcrireListe ws, "I", "LST_SAISONS", _
        ValeursDistinctes(Donnees, MTC_SAISON)

    EcrireListe ws, "J", "LST_CATEGORIES", _
        ValeursDistinctes(Donnees, MTC_CATEGORIE)

    EcrireListe ws, "K", "LST_PHASES", _
        ValeursDistinctes(Donnees, MTC_PHASE)

    EcrireListe ws, "L", "LST_LIEUX", _
        ValeursDistinctes(Donnees, MTC_LIEU)

    EcrireListe ws, "M", "LST_RESULTATS", _
        ValeursDistinctes(Donnees, MTC_RESULTAT)

    EcrireListe ws, "N", "LST_ADVERSAIRES", _
        ValeursDistinctes(Donnees, MTC_ADVERSAIRE)

End Sub


Private Function ValeursDistinctes( _
    ByRef Donnees As Variant, _
    ByVal Colonne As Long _
) As Collection

    Dim Valeurs As Collection
    Dim Valeur As String
    Dim Ligne As Long

    Set Valeurs = New Collection

    Valeurs.Add VALEUR_TOUS, Normaliser(VALEUR_TOUS)

    If IsEmpty(Donnees) Then
        Set ValeursDistinctes = Valeurs
        Exit Function
    End If

    For Ligne = 1 To UBound(Donnees, 1)

        Valeur = TexteCellule(Donnees(Ligne, Colonne))

        If Len(Valeur) > 0 Then

            If Not CleExiste(Valeurs, Normaliser(Valeur)) Then
                Valeurs.Add Valeur, Normaliser(Valeur)
            End If

        End If

    Next Ligne

    Set ValeursDistinctes = Valeurs

End Function


Private Sub EcrireListe( _
    ByVal ws As Worksheet, _
    ByVal Colonne As String, _
    ByVal NomPlage As String, _
    ByVal Valeurs As Collection _
)

    Dim i As Long

    ws.Range(Colonne & "7:" & Colonne & "500").ClearContents

    For i = 1 To Valeurs.Count
        ws.Range(Colonne & (6 + i)).Value = CStr(Valeurs.Item(i))
    Next i

    On Error Resume Next
    ThisWorkbook.Names(NomPlage).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=NomPlage, _
        RefersTo:= _
            ws.Range( _
                Colonne & "7:" & Colonne & (6 + Valeurs.Count) _
            )

End Sub


' =========================================================
' DIAGNOSTIC
'
' A lancer a la main depuis Outils > Macro > Macros quand
' l'actualisation s'arrete brutalement. La macro fait le
' strict minimum : elle ouvre le premier fichier de la
' racine, compte ses lignes, et le referme, sans rien
' ecrire dans le classeur.
'
' Si elle passe, le probleme n'est pas l'ouverture des
' fichiers de match mais l'ecriture en base. Si elle
' echoue, recap-trace.txt dit a quelle etape.
' =========================================================
Public Sub TesterLecturePremierFichier()

    Dim Fichiers As Collection
    Dim Match As InfosMatch
    Dim Fichier As String

    Dim EtatEvenements As Boolean
    Dim EtatCalcul As XlCalculation

    If Len(ThisWorkbook.Path) = 0 Then

        MsgBox _
            "Enregistre d'abord ce classeur.", _
            vbExclamation, _
            "Classeur non enregistr" & ChrW(233)

        Exit Sub

    End If

    Tracer "--- Test de lecture ---"

    Set Fichiers = FichiersDeLaRacine()

    If Fichiers.Count = 0 Then

        MsgBox _
            "Aucun fichier " & ChrW(224) & " la racine.", _
            vbExclamation, _
            "Test"

        Exit Sub

    End If

    Fichier = CStr(Fichiers.Item(1))

    Tracer "Fichier retenu : " & Fichier

    If Not AutoriserLecture(Fichiers) Then
        Tracer "Autorisation refusee"
        Exit Sub
    End If

    Tracer "Autorisation accordee"

    EtatEvenements = Application.EnableEvents
    EtatCalcul = Application.Calculation

    ' L'affichage reste actif : voir le fichier s'ouvrir
    ' indique deja jusqu'ou on est alle.
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Tracer "Ouverture"

    LireFichierMatch Fichier, Match

    Tracer "Ferme"

    Application.Calculation = EtatCalcul
    Application.EnableEvents = EtatEvenements

    MsgBox _
        "Fichier : " & Fichier & vbCrLf & vbCrLf & _
        "Lu : " & Match.Lu & vbCrLf & _
        "Actions : " & Match.NbActions & vbCrLf & _
        "Joueurs : " & Match.NbJoueurs & vbCrLf & _
        "Temps de jeu saisis : " & Match.NbMinutes & vbCrLf & _
        "Cat" & ChrW(233) & "gorie : " & Match.Categorie & vbCrLf & _
        "Phase : " & Match.Phase & vbCrLf & _
        "ID : " & Match.IdMatch, _
        vbInformation, _
        "Test de lecture"

    Tracer "Test termine"

End Sub
