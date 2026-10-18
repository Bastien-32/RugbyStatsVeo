Attribute VB_Name = "modRecapDistribution"
Option Explicit

' =========================================================
' REINITIALISATION AVANT DISTRIBUTION
'
' Retire du recapitulatif tout ce qui appartient a un club
' en particulier : les matchs integres, le journal de la
' saison, les compositions, les listes de filtres qui en
' decoulent, le classement et les temps de jeu.
'
' Ce qui reste est le parametrage commun a tous les clubs :
' le tableau RECAP_STATS, qui dit quelles statistiques
' afficher et sous quel libelle, ainsi que la mise en page
' et les boutons.
'
' Comme pour le classeur de saisie, la macro commence par
' enregistrer sous un autre nom : c'est la copie qui est
' videe, et le fichier de travail reste intact sur le
' disque avec toute la saison.
' =========================================================

' Nomme l'etape en cours pour que le message d'erreur dise
' ou la reinitialisation s'est arretee.
Private EtapeRecapDistribution As String


Public Sub ReinitialiserRecapPourDistribution()

    Dim Reponse As VbMsgBoxResult
    Dim Destination As String

    Dim EtatEvenements As Boolean
    Dim EtatAffichage As Boolean

    ' Le fichier a produire est choisi avant tout : la
    ' fenetre d'enregistrement vaut aussi autorisation
    ' d'ecrire la ou macOS la demande.
    Destination = ChoisirRecapADistribuer()

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
        "- les matchs integres ;" & vbCrLf & _
        "- le journal d'actions de la saison ;" & vbCrLf & _
        "- les compositions ;" & vbCrLf & _
        "- les listes de filtres (saisons, adversaires...) ;" & vbCrLf & _
        "- le classement et les temps de jeu." & _
        vbCrLf & vbCrLf & _
        "Le parametrage des statistiques est conserve." & _
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
    EtapeRecapDistribution = "enregistrement de la copie"

    ThisWorkbook.SaveAs _
        Filename:=Destination, _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled

    EtapeRecapDistribution = "journal de la saison"
    ViderTableauRecap _
        TableauRecap(FEUILLE_JOURNAL, TBL_JOURNAL)

    EtapeRecapDistribution = "compositions"
    ViderTableauRecap _
        TableauRecap(FEUILLE_COMPOSITIONS, TBL_COMPOSITIONS)

    ' L'inventaire des matchs est vide en dernier : les
    ' listes de filtres s'en deduisent, et le classement
    ' les relit.
    EtapeRecapDistribution = "inventaire des matchs"
    ViderTableauRecap _
        TableauRecap(FEUILLE_MATCHS, TBL_MATCHS)

    ' Les menus de la feuille Filtres sont alimentes par
    ' des colonnes masquees de Parametres, remplies depuis
    ' l'inventaire : elles portent donc les adversaires et
    ' les saisons du club.
    EtapeRecapDistribution = "listes de filtres"
    ConstruireListesFiltres

    ' Remet les criteres a leurs valeurs par defaut, puis
    ' rejoue le calcul : le classement et les temps de jeu
    ' repassent a vide.
    EtapeRecapDistribution = "filtres, classement et temps de jeu"
    ReinitialiserFiltres

    EtapeRecapDistribution = ""

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
        "Ferme-le et rouvre " & ChrW(171) & " Recapitulatif " & _
        "saison.xlsm " & ChrW(187) & " pour retrouver ta saison.", _
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
        EtapeRecapDistribution & "." & _
        vbCrLf & vbCrLf & _
        "Erreur " & NumeroErreur & " : " & _
        DescriptionErreur & _
        vbCrLf & vbCrLf & _
        "Ferme ce classeur sans enregistrer.", _
        vbCritical, _
        "Erreur"

End Sub


' =========================================================
' Le fichier a produire.
'
' Le recapitulatif vit dans Matchs, et le dossier dist a la
' racine du projet : on accepte les deux emplacements,
' comme le classeur de saisie le fait pour trouver le
' moteur video.
'
' Les fenetres d'enregistrement d'Excel pour Mac rendent
' des chemins inexploitables - extension doublee, caractere
' parasite - la ou une simple saisie donne toujours ce
' qu'on attend.
' =========================================================
Private Function ChoisirRecapADistribuer() As String

    Dim Dossier As String
    Dim Nom As String
    Dim Chemin As String

    Dossier = DossierDist()

    Nom = InputBox( _
        "Nom du fichier " & ChrW(224) & " produire, dans le " & _
        "dossier dist :" & vbCrLf & vbCrLf & Dossier, _
        "Copie " & ChrW(224) & " distribuer", _
        "Recapitulatif saison - vierge.xlsm")

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

        ' SaveAs refuserait d'ecraser un fichier ouvert
        ' ailleurs : on le supprime d'abord.
        On Error Resume Next
        Kill Chemin
        On Error GoTo 0

    End If

    ChoisirRecapADistribuer = Chemin

End Function


Private Function DossierDist() As String

    Dim Candidat As String

    ' Le projet range dist a sa racine, et le
    ' recapitulatif dans Matchs : le dossier est donc un
    ' niveau au-dessus du classeur.
    Candidat = ThisWorkbook.Path & _
        Application.PathSeparator & ".." & _
        Application.PathSeparator & "dist"

    If Dir(Candidat, vbDirectory) <> "" Then
        DossierDist = Candidat
        Exit Function
    End If

    ' Classeur pose a la racine du projet.
    Candidat = ThisWorkbook.Path & _
        Application.PathSeparator & "dist"

    If Dir(Candidat, vbDirectory) = "" Then

        On Error Resume Next
        MkDir Candidat
        On Error GoTo 0

    End If

    DossierDist = Candidat

End Function
