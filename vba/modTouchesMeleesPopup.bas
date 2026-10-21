Attribute VB_Name = "modTouchesMeleesPopup"
Option Explicit

' =========================================================
' TOUCHES ET MELEES - FABRICATION DES DEUX POPUPS
'
' A executer une fois depuis "Createur de match.xlsm".
'
' Deux feuilles separees plutot qu'une seule : une touche
' et une melee ne demandent pas les memes informations,
' et la feuille "Popup" qui sert a ajouter un joueur reste
' intacte.
'
' Aucun caractere accentue, pas meme dans les chaines :
' le VBE de macOS importe les .bas en Mac Roman.
'
' Les quatre premieres lignes affichent ce que la saisie
' video a deja determine. Elles sont grisees pour signaler
' qu'on n'a pas a y toucher : une correction se fait dans
' le tableau.
'
' Les choix se font au clic, sur des boutons poses en
' ligne : derouler une liste pour chaque champ coutait
' trop de temps pendant le depouillement. Seul le sauteur
' garde sa liste, la composition etant trop longue pour
' tenir en boutons.
'
' La valeur choisie ne s'ecrit pas dans le bouton mais
' dans une cellule a part, en colonne masquee : la saisie
' continue de la lire par son nom de plage, comme du temps
' des listes deroulantes.
' =========================================================

Public Const FEUILLE_POPUP_TO As String = "Popup touche"
Public Const FEUILLE_POPUP_ME As String = "Popup melee"

Private Const FEUILLE_TABLEAUX As String = "Touches Melees"

Private Const GRIS As Long = 15132390     ' RGB(230, 230, 230)
Private Const BLEU As Long = 12611584     ' RGB(0, 112, 192)
Private Const VERT As Long = 32768        ' RGB(0, 128, 0)
Private Const ROUGE As Long = 192         ' RGB(192, 0, 0)
Private Const SABLE As Long = 14083324    ' RGB(252, 228, 214)
Private Const PELOUSE As Long = 14282978  ' RGB(226, 240, 217)
Private Const ENBUT As Long = 11854022    ' RGB(198, 224, 180)

' Disposition commune aux deux popups.
Private Const COL_LIBELLE As Long = 2
Private Const COL_PREMIER_BOUTON As Long = 3
Private Const LARGEUR_BOUTON As Long = 2

' Colonne masquee ou atterrissent les valeurs choisies.
Private Const COL_VALEURS As Long = 30

Private Const CELLULE_SENS As String = "J5:L6"

Private EtapeEnCours As String


' ---------------------------------------------------------
' Disposition des boutons
'
' Une entree par bouton, lue aussi bien pour dessiner la
' popup que pour reconnaitre le bouton clique :
'
'     nom | ligne | rang | libelle | valeur
'     nom | ligne | rang | libelle | valeur | largeur
'     nom | ligne | rang | libelle | valeur | largeur | nom2 | valeur2
'
' Le rang est la position du bouton sur sa ligne, a partir
' de zero. Le libelle est ce qui s'affiche, la valeur ce
' qui part dans le tableau : les zones de terrain disent
' "en-but" des deux cotes, mais ecrivent "N en-but" ou
' "E en-but" selon le camp.
'
' La largeur est le nombre de colonnes du bouton, deux
' par defaut. Les cases du terrain n'en prennent qu'une,
' sans quoi ses onze zones deborderaient de l'ecran.
'
' Les deux derniers champs servent au terrain de la
' melee : une case y renseigne a la fois la longueur et
' la largeur, d'un seul clic.
' ---------------------------------------------------------

Public Function OptionsTouche() As Variant

    Dim T As String

    T = _
        "POPUP_TO_ZONE_SAUT|13|0|Z0|Z0" & vbTab & _
        "POPUP_TO_ZONE_SAUT|13|1|Z1|Z1" & vbTab & _
        "POPUP_TO_ZONE_SAUT|13|2|Z1,5|Z1,5" & vbTab & _
        "POPUP_TO_ZONE_SAUT|13|3|Z2|Z2" & vbTab & _
        "POPUP_TO_ZONE_SAUT|13|4|Z3|Z3"

    T = T & vbTab & _
        "POPUP_TO_BALLON|15|0|Chaud|Chaud" & vbTab & _
        "POPUP_TO_BALLON|15|1|Froid|Froid"

    T = T & vbTab & BandeTerrain("POPUP_TO_ZONE_TERRAIN", 19)

    T = T & vbTab & _
        "POPUP_TO_ALIGNEMENT|11|0|2|2" & vbTab & _
        "POPUP_TO_ALIGNEMENT|11|1|3|3" & vbTab & _
        "POPUP_TO_ALIGNEMENT|11|2|4|4" & vbTab & _
        "POPUP_TO_ALIGNEMENT|11|3|5|5" & vbTab & _
        "POPUP_TO_ALIGNEMENT|11|4|6|6" & vbTab & _
        "POPUP_TO_ALIGNEMENT|11|5|Complet|Complet"

    OptionsTouche = Split(T, vbTab)

End Function


Public Function OptionsMelee() As Variant

    Dim T As String

    ' Une seule grille pour les deux champs : le couloir
    ' se lit du rang clique, la distance de sa colonne.
    T = GrilleTerrain( _
        "POPUP_ME_ZONE_LONGUEUR", "POPUP_ME_ZONE_LARGEUR", 11)

    T = T & vbTab & _
        "POPUP_ME_UTILISATION|15|0|Avants|Avants" & vbTab & _
        "POPUP_ME_UTILISATION|15|1|Trois-quart|Trois-quart" & vbTab & _
        "POPUP_ME_UTILISATION|15|2|Pied|Pied"

    OptionsMelee = Split(T, vbTab)

End Function


' ---------------------------------------------------------
' Le terrain
'
' Onze zones dans la longueur, de notre en-but a celui
' d'en face. La touche n'en retient que la distance, et
' se contente d'une bande ; la melee se joue n'importe ou
' et demande la grille entiere, ou chaque case porte sa
' longueur et sa largeur.
' ---------------------------------------------------------

' La distance affichee sur chaque case, de gauche a droite.
Private Function DistancesTerrain() As Variant

    DistancesTerrain = Array( _
        "en-but", "15m", "22m", "30m", "40m", "50m", _
        "40m", "30m", "22m", "15m", "en-but")

End Function


' Ce que la case ecrit dans le tableau : notre camp a
' gauche, le milieu neutre, celui d'en face a droite.
'
' Quand nous jouons dans l'autre sens, la bande se lit a
' l'envers : la case de gauche appartient alors au camp
' adverse. Les distances affichees, elles, ne changent
' pas - elles sont symetriques.
Private Function ValeurTerrain(ByVal Rang As Long) As String

    Dim Distances As Variant
    Dim r As Long

    Distances = DistancesTerrain

    If SensJeuInverse() Then
        r = UBound(Distances) - Rang
    Else
        r = Rang
    End If

    If r = 5 Then
        ValeurTerrain = "50m"
    ElseIf r < 5 Then
        ValeurTerrain = "N " & Distances(r)
    Else
        ValeurTerrain = "E " & Distances(r)
    End If

End Function


' Le couloir, vu depuis notre equipe. Inverser le sens de
' jeu fait pivoter le terrain d'un demi-tour : ce qui est
' en haut de l'ecran passe a droite pour nous. Les
' libelles affiches ne bougent pas, seule la valeur
' enregistree suit.
Private Function ValeurLargeur(ByVal Rang As Long) As String

    Dim Couloirs As Variant

    Couloirs = Array("Gauche", "Milieu", "Droit")

    If SensJeuInverse() Then
        ValeurLargeur = Couloirs(UBound(Couloirs) - Rang)
    Else
        ValeurLargeur = Couloirs(Rang)
    End If

End Function


' La bande de la touche : une seule rangee, la longueur
' seule, le tableau des touches n'ayant pas de colonne
' pour le cote.
Private Function BandeTerrain( _
    ByVal NomPlage As String, _
    ByVal Ligne As Long) As String

    Dim T As String
    Dim Distances As Variant
    Dim i As Long

    Distances = DistancesTerrain

    For i = 0 To UBound(Distances)

        T = T & NomPlage & "|" & Ligne & "|" & i & "|" & _
            Distances(i) & "|" & ValeurTerrain(i) & "|1" & vbTab

    Next i

    BandeTerrain = Left$(T, Len(T) - 1)

End Function


' La grille de la melee : trois rangees, un clic pour les
' deux champs.
Private Function GrilleTerrain( _
    ByVal NomLongueur As String, _
    ByVal NomLargeur As String, _
    ByVal PremiereLigne As Long) As String

    Dim T As String
    Dim Distances As Variant
    Dim Couloirs As Variant
    Dim i As Long
    Dim j As Long

    Distances = DistancesTerrain
    Couloirs = Array("Gauche", "Milieu", "Droit")

    For j = 0 To UBound(Couloirs)

        For i = 0 To UBound(Distances)

            T = T & NomLongueur & "|" & (PremiereLigne + j) & _
                "|" & i & "|" & Distances(i) & "|" & _
                ValeurTerrain(i) & "|1|" & NomLargeur & "|" & _
                ValeurLargeur(j) & vbTab

        Next i

    Next j

    GrilleTerrain = Left$(T, Len(T) - 1)

End Function


' ---------------------------------------------------------
' Bouton du sens de jeu
'
' Le meme que celui de la feuille de saisie, repris ici :
' on s'apercoit souvent qu'on est dans le mauvais sens en
' voyant le terrain, et il serait penible de fermer la
' popup pour le corriger.
' ---------------------------------------------------------


Private Sub EcrireBoutonSensJeu( _
    ByVal ws As Worksheet, _
    ByVal NomPlage As String)

    With ws.Range(CELLULE_SENS)

        .Merge
        .Interior.Color = GRIS
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)

    End With

    NommerPlage ws, CELLULE_SENS, NomPlage

    AfficherSensJeuPopup ws

End Sub


' La fleche est agrandie, le libelle garde sa taille.
Private Sub AfficherSensJeuPopup(ByVal ws As Worksheet)

    Dim Texte As String

    On Error GoTo Sortie

    Texte = LibelleSensJeu()

    ws.Range(CELLULE_SENS).Cells(1, 1).Value = Texte

    ws.Range(CELLULE_SENS).Cells(1, 1) _
        .Characters(Len(Texte), 1).Font.Size = 18

Sortie:

End Sub


' Apres bascule : le terrain garde ses cases, mais elles
' ne valent plus la meme chose, et les deux camps ont
' change de cote.
Private Sub RafraichirTerrainPopup( _
    ByVal ws As Worksheet, _
    ByVal Options As Variant, _
    ByVal LigneEnTetes As Long)

    PoserValeurs ws, Options
    EcrireEnTetesTerrain ws, LigneEnTetes
    AfficherSensJeuPopup ws
    LibererSelectionPopup ws

End Sub

' Le terrain est dessine une fois a la construction : il
' faut le reorienter a chaque ouverture, le sens ayant pu
' changer depuis la feuille de saisie ou l'autre popup.
Public Sub OrienterPopupTouche()

    RafraichirTerrainPopup _
        ThisWorkbook.Sheets(FEUILLE_POPUP_TO), _
        OptionsTouche, 17

End Sub

Public Sub OrienterPopupMelee()

    RafraichirTerrainPopup _
        ThisWorkbook.Sheets(FEUILLE_POPUP_ME), _
        OptionsMelee, 9

End Sub

' Excel ne redeclenche rien si la selection ne bouge pas :
' sans ce deplacement, un second clic sur le meme bouton
' resterait sans effet. Les evenements sont coupes pendant
' le saut, sans quoi A1 serait traitee comme un clic.
Private Sub LibererSelectionPopup(ByVal ws As Worksheet)

    Dim EtatEvenements As Boolean

    EtatEvenements = Application.EnableEvents

    On Error GoTo Sortie

    Application.EnableEvents = False

    If ActiveSheet Is ws Then
        ws.Range("A1").Select
    End If

Sortie:

    Application.EnableEvents = EtatEvenements

End Sub


' ---------------------------------------------------------
' Construction
' ---------------------------------------------------------

Public Sub ConstruirePopupsTouchesMelees()

    Dim EtatAffichage As Boolean

    EtatAffichage = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    EtapeEnCours = "popup touche"
    ConstruirePopupTouche

    EtapeEnCours = "popup melee"
    ConstruirePopupMelee

    Application.ScreenUpdating = EtatAffichage

    If ModeSilencieux Then Exit Sub

    MsgBox _
        "Les deux popups sont pretes." & vbCrLf & vbCrLf & _
        "Elles sont visibles pour que tu en verifies le " & _
        "rendu. Le VBA de liaison les masquera.", _
        vbInformation, _
        "Touches et melees"

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


Private Sub ConstruirePopupTouche()

    Dim ws As Worksheet

    Set ws = FeuillePopup(FEUILLE_POPUP_TO)

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws, "TOUCHE"
    EcrireContexte ws, "Lance pour"

    ' Le sauteur garde sa liste : la composition ne tient
    ' pas en boutons.
    EcrireChampListe ws, 9, "Sauteur", "POPUP_TO_SAUTEUR"

    ' L'alignement suit le sauteur : les deux se lisent
    ' sur la meme image, avant d'en venir au ballon.
    EcrireLibelle ws, 11, "Alignement"

    EcrireLibelle ws, 13, "Zone de saut"
    EcrireLibelle ws, 15, "Ballon"

    ' L'intitule tient sur la rangee des deux camps :
    ' une ligne pour lui seul ne servait a rien.
    EcrireSousTitre ws, 17, "ZONE TERRAIN"
    EcrireEnTetesTerrain ws, 17
    EcrireLibelle ws, 19, "Zone de touche"

    DessinerOptions ws, OptionsTouche

    PoserValeurs ws, OptionsTouche

    EcrireChampTexte ws, 21, "Observations", "POPUP_TO_OBS"

    ' La saisie la repose a chaque ouverture, en suivant
    ' la composition : la poser ici donne une popup
    ' complete des sa construction.
    PoserListeSauteur

    EcrireBoutons ws, 23, "POPUP_TO_VALIDER", _
        "POPUP_TO_ANNULER"

    EcrireBoutonSensJeu ws, "POPUP_TO_SENS_JEU"

    NommerPlage ws, "E4", "POPUP_TO_PLAY_PAUSE"
    NommerPlage ws, "C4:C7", "POPUP_TO_CONTEXTE"

    ' Ce que l'ouverture remet a blanc : les deux champs
    ' libres et les valeurs des boutons.
    NommerPlage ws, _
        AdresseNommee("POPUP_TO_SAUTEUR") & "," & _
        AdresseValeurs(ws, OptionsTouche) & "," & _
        AdresseNommee("POPUP_TO_OBS"), _
        "POPUP_TO_CHAMPS"

    MettreEnFormePopup ws, 23

End Sub


Private Sub ConstruirePopupMelee()

    Dim ws As Worksheet

    Set ws = FeuillePopup(FEUILLE_POPUP_ME)

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws, "MELEE"
    EcrireContexte ws, "Introduction pour"

    EcrireSousTitre ws, 9, "ZONE DE JEU"
    EcrireEnTetesTerrain ws, 9
    EcrireLibelle ws, 11, "Gauche"
    EcrireLibelle ws, 12, "Milieu"
    EcrireLibelle ws, 13, "Droit"

    EcrireLibelle ws, 15, "Utilisation"

    DessinerOptions ws, OptionsMelee

    PoserValeurs ws, OptionsMelee

    EcrireChampTexte ws, 17, "Observations", "POPUP_ME_OBS"

    EcrireBoutons ws, 19, "POPUP_ME_VALIDER", _
        "POPUP_ME_ANNULER"

    EcrireBoutonSensJeu ws, "POPUP_ME_SENS_JEU"

    NommerPlage ws, "E4", "POPUP_ME_PLAY_PAUSE"
    NommerPlage ws, "C4:C7", "POPUP_ME_CONTEXTE"

    NommerPlage ws, _
        AdresseValeurs(ws, OptionsMelee) & "," & _
        AdresseNommee("POPUP_ME_OBS"), _
        "POPUP_ME_CHAMPS"

    MettreEnFormePopup ws, 19

End Sub


' ---------------------------------------------------------
' Boutons d'options
' ---------------------------------------------------------

Private Sub DessinerOptions( _
    ByVal ws As Worksheet, _
    ByVal Entrees As Variant)

    Dim i As Long

    For i = 0 To UBound(Entrees)

        ' Le fond suit l'ancre d'une cellule fusionnee,
        ' mais pas les bordures : tout se pose sur la zone
        ' entiere, sans quoi le bouton n'a de contour que
        ' sur sa premiere colonne.
        With ZoneBouton(ws, CStr(Entrees(i)))

            .Merge

            If Not EstCaseTerrain(CStr(Entrees(i))) Then
                .Value = Champ(CStr(Entrees(i)), 3)
            End If

            .Interior.Color = CouleurRepos(CStr(Entrees(i)))
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous

            ' Les lignes du terrain sont blanches, comme
            ' sur une pelouse ; les boutons gardent leur
            ' contour gris.
            If EstCaseTerrain(CStr(Entrees(i))) Then
                .Borders.Color = RGB(255, 255, 255)
            Else
                .Borders.Color = RGB(150, 150, 150)
            End If

        End With

    Next i

End Sub


' Les champs renseignes par ces boutons, dans l'ordre ou
' ils apparaissent et sans doublon. Une case du terrain
' en porte deux : la longueur et le couloir.
Private Function NomsDesChamps( _
    ByVal Entrees As Variant) As Variant

    Dim i As Long
    Dim j As Long
    Dim Noms As String
    Dim Vus As String
    Dim Candidat As String

    For i = 0 To UBound(Entrees)

        For j = 0 To 6 Step 6

            Candidat = Champ(CStr(Entrees(i)), j)

            If Candidat <> "" Then

                If InStr(1, Vus, "|" & Candidat & "|") = 0 Then

                    Vus = Vus & "|" & Candidat & "|"

                    If Noms <> "" Then Noms = Noms & vbTab
                    Noms = Noms & Candidat

                End If

            End If

        Next j

    Next i

    NomsDesChamps = Split(Noms, vbTab)

End Function


' Une cellule par champ, dans la colonne masquee : c'est
' elle que la saisie lit, et c'est elle que le clic sur un
' bouton remplit. Elles sont empilees a partir du haut,
' sans rapport avec la ligne des boutons : une case du
' terrain en nourrit deux, qui se marcheraient dessus.
Private Sub PoserValeurs( _
    ByVal ws As Worksheet, _
    ByVal Entrees As Variant)

    Dim Noms As Variant
    Dim i As Long

    Noms = NomsDesChamps(Entrees)

    For i = 0 To UBound(Noms)

        NommerPlage ws, _
            ws.Cells(2 + i, COL_VALEURS).Address, _
            CStr(Noms(i))

    Next i

    ws.Columns(COL_VALEURS).Hidden = True

End Sub


' Les memes cellules, citees d'un trait : l'ouverture de
' la popup les vide toutes en une fois.
Private Function AdresseValeurs( _
    ByVal ws As Worksheet, _
    ByVal Entrees As Variant) As String

    Dim Noms As Variant

    Noms = NomsDesChamps(Entrees)

    AdresseValeurs = ws.Range( _
        ws.Cells(2, COL_VALEURS), _
        ws.Cells(2 + UBound(Noms), COL_VALEURS)).Address

End Function


Private Function ZoneBouton( _
    ByVal ws As Worksheet, _
    ByVal Entree As String) As Range

    Set ZoneBouton = CelluleOption(ws, Entree) _
        .Resize(1, LargeurOption(Entree))

End Function


Private Function CelluleOption( _
    ByVal ws As Worksheet, _
    ByVal Entree As String) As Range

    Set CelluleOption = ws.Cells( _
        CLng(Champ(Entree, 1)), _
        COL_PREMIER_BOUTON + _
            CLng(Champ(Entree, 2)) * LargeurOption(Entree))

End Function


' Deux colonnes par defaut, une seule pour les cases du
' terrain : onze zones de deux colonnes sortiraient de
' l'ecran.
Private Function LargeurOption( _
    ByVal Entree As String) As Long

    Dim Largeur As String

    Largeur = Champ(Entree, 5)

    If Largeur = "" Then
        LargeurOption = LARGEUR_BOUTON
    Else
        LargeurOption = CLng(Largeur)
    End If

End Function


' Les cases du terrain se reconnaissent a leur champ :
' elles seules portent la pelouse, et restent muettes,
' la distance etant ecrite une fois pour toutes en tete
' de colonne.
Private Function EstCaseTerrain( _
    ByVal Entree As String) As Boolean

    EstCaseTerrain = _
        InStr(1, Champ(Entree, 0), "ZONE_TERRAIN") > 0 Or _
        InStr(1, Champ(Entree, 0), "ZONE_LONGUEUR") > 0

End Function


' La couleur de repos : le terrain porte la sienne, ses
' en-buts plus soutenus, et les autres choix gardent le
' sable commun aux popups.
Private Function CouleurRepos( _
    ByVal Entree As String) As Long

    If Not EstCaseTerrain(Entree) Then

        CouleurRepos = SABLE

    ElseIf InStr(1, Champ(Entree, 3), "en-but") > 0 Then

        CouleurRepos = ENBUT

    Else

        CouleurRepos = PELOUSE

    End If

End Function


' Un champ d'une entree, decoupee sur ses barres.
Private Function Champ( _
    ByVal Entree As String, _
    ByVal Rang As Long) As String

    Dim Morceaux As Variant

    Morceaux = Split(Entree, "|")

    If Rang > UBound(Morceaux) Then Exit Function

    Champ = Morceaux(Rang)

End Function


' ---------------------------------------------------------
' Clics
'
' Des cellules nommees, pas des formes : la palette de
' saisie fonctionne deja ainsi, et Worksheet_SelectionChange
' suffit a les detecter.
' ---------------------------------------------------------

Public Sub ClicPopupTouche(ByVal Target As Range)

    Dim ws As Worksheet
    Dim Cellule As Range

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_TO)

    Set Cellule = AncreCliquee(Target)
    If Cellule Is Nothing Then Exit Sub

    If Touche(ws, Cellule, "POPUP_TO_PLAY_PAUSE") Then
        BasculerLectureDepuisPopup
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_TO_SENS_JEU") Then
        BasculerSensJeu
        RafraichirTerrainPopup ws, OptionsTouche, 17
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_TO_VALIDER") Then
        ValiderPopupTouche
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_TO_ANNULER") Then
        AnnulerPopupTouche
        Exit Sub
    End If

    AppliquerClicOption ws, Cellule, OptionsTouche

End Sub


Public Sub ClicPopupMelee(ByVal Target As Range)

    Dim ws As Worksheet
    Dim Cellule As Range

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_ME)

    Set Cellule = AncreCliquee(Target)
    If Cellule Is Nothing Then Exit Sub

    If Touche(ws, Cellule, "POPUP_ME_PLAY_PAUSE") Then
        BasculerLectureDepuisPopup
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_ME_SENS_JEU") Then
        BasculerSensJeu
        RafraichirTerrainPopup ws, OptionsMelee, 9
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_ME_VALIDER") Then
        ValiderPopupMelee
        Exit Sub
    End If

    If Touche(ws, Cellule, "POPUP_ME_ANNULER") Then
        AnnulerPopupMelee
        Exit Sub
    End If

    AppliquerClicOption ws, Cellule, OptionsMelee

End Sub


' Un clic sur une cellule fusionnee livre toute la plage :
' on ne garde que son ancre, la seule que la disposition
' connaisse.
Private Function AncreCliquee( _
    ByVal Target As Range) As Range

    If Target.Areas.Count > 1 Then Exit Function

    If Target.Cells.Count > 1 Then

        If Not Target.MergeCells Then Exit Function
        Set AncreCliquee = Target.Cells(1, 1)
        Exit Function

    End If

    Set AncreCliquee = Target

End Function


Private Function Touche( _
    ByVal ws As Worksheet, _
    ByVal Cellule As Range, _
    ByVal NomPlage As String) As Boolean

    On Error Resume Next

    Touche = Not Intersect( _
        Cellule, ws.Range(NomPlage)) Is Nothing

    On Error GoTo 0

End Function


' Le bouton clique prend la couleur, ses voisins de champ
' la perdent, et la valeur part dans la cellule masquee.
' Reclique, il se rend : le champ repart vide.
Private Sub AppliquerClicOption( _
    ByVal ws As Worksheet, _
    ByVal Cellule As Range, _
    ByVal Entrees As Variant)

    Dim i As Long
    Dim Choisie As Long
    Dim NomPlage As String
    Dim Autre As String
    Dim DejaChoisi As Boolean

    Choisie = -1

    For i = 0 To UBound(Entrees)

        If CelluleOption(ws, CStr(Entrees(i))).Address = _
            Cellule.Address Then

            Choisie = i
            Exit For

        End If

    Next i

    If Choisie < 0 Then Exit Sub

    NomPlage = Champ(CStr(Entrees(Choisie)), 0)
    DejaChoisi = (Cellule.Interior.Color = BLEU)

    ' Les evenements sont rendus meme en cas d'erreur :
    ' coupes, plus aucun clic ne passerait ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    For i = 0 To UBound(Entrees)

        If Champ(CStr(Entrees(i)), 0) = NomPlage Then

            With ZoneBouton(ws, CStr(Entrees(i)))
                .Interior.Color = CouleurRepos(CStr(Entrees(i)))
                .Font.Color = RGB(0, 0, 0)
            End With

        End If

    Next i

    Autre = Champ(CStr(Entrees(Choisie)), 6)

    If DejaChoisi Then

        ws.Range(NomPlage).ClearContents

        ' La case du terrain porte deux champs : les
        ' abandonner ensemble.
        If Autre <> "" Then ws.Range(Autre).ClearContents

    Else

        With ZoneBouton(ws, CStr(Entrees(Choisie)))
            .Interior.Color = BLEU
            .Font.Color = RGB(255, 255, 255)
        End With

        ws.Range(NomPlage).Value = _
            Champ(CStr(Entrees(Choisie)), 4)

        If Autre <> "" Then
            ws.Range(Autre).Value = _
                Champ(CStr(Entrees(Choisie)), 7)
        End If

    End If

Sortie:

    Application.EnableEvents = True

End Sub


' Appelee a chaque ouverture : sans elle, les boutons de
' la touche precedente resteraient allumes.
Public Sub EffacerSelectionsOptions(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    If ws.Name = FEUILLE_POPUP_TO Then
        Entrees = OptionsTouche
    ElseIf ws.Name = FEUILLE_POPUP_ME Then
        Entrees = OptionsMelee
    Else
        Exit Sub
    End If

    On Error Resume Next

    For i = 0 To UBound(Entrees)

        With ZoneBouton(ws, CStr(Entrees(i)))
            .Interior.Color = CouleurRepos(CStr(Entrees(i)))
            .Font.Color = RGB(0, 0, 0)
        End With

    Next i

    On Error GoTo 0

End Sub


' ---------------------------------------------------------
' Briques communes
' ---------------------------------------------------------

Private Sub EcrireTitre( _
    ByVal ws As Worksheet, _
    ByVal Titre As String)

    With ws.Range("B2")
        .Value = Titre
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = BLEU
    End With

End Sub


' Les quatre valeurs que la saisie video connait deja.
Private Sub EcrireContexte( _
    ByVal ws As Worksheet, _
    ByVal LibelleEquipe As String)

    Dim Libelles As Variant
    Dim i As Long

    Libelles = Array( _
        "Temps video", "Mi-temps", LibelleEquipe, "Issue")

    For i = 0 To 3

        ws.Cells(4 + i, 2).Value = Libelles(i)
        ws.Cells(4 + i, 2).Font.Italic = True

        With ws.Cells(4 + i, 3).Resize(1, 2)
            .Merge
            .Interior.Color = GRIS
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

    Next i

    ' Lecture et pause sans quitter la popup : la video
    ' reste pilotable pendant la saisie.
    With ws.Cells(4, 5).Resize(1, 3)

        .Merge
        .Value = "PLAY / PAUSE"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub


Private Sub EcrireLibelle( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String)

    With ws.Cells(Ligne, COL_LIBELLE)
        .Value = Libelle
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

End Sub


' Les trois rangs du terrain vont ensemble : un intertitre
' les rassemble sous un meme nom.
Private Sub EcrireSousTitre( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Titre As String)

    With ws.Cells(Ligne, COL_LIBELLE)
        .Value = Titre
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = BLEU
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

End Sub


' Au-dessus du terrain : les deux camps, puis la distance
' de chaque colonne. Du texte, pas des boutons : seules
' les cases du terrain se cliquent.
Private Sub EcrireEnTetesTerrain( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long)

    Dim Distances As Variant
    Dim i As Long

    ' Les cinq colonnes de chaque cote, le milieu laisse
    ' a decouvert entre les deux.
    With ws.Cells(Ligne, COL_PREMIER_BOUTON).Resize(1, 5)
        .Merge
        .Value = IIf(SensJeuInverse(), _
            "CAMP ADVERSE", "NOTRE CAMP")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = BLEU
    End With

    With ws.Cells(Ligne, COL_PREMIER_BOUTON + 6).Resize(1, 5)
        .Merge
        .Value = IIf(SensJeuInverse(), _
            "NOTRE CAMP", "CAMP ADVERSE")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = BLEU
    End With

    Distances = DistancesTerrain

    For i = 0 To UBound(Distances)

        With ws.Cells(Ligne + 1, COL_PREMIER_BOUTON + i)
            .Value = Distances(i)
            .HorizontalAlignment = xlCenter
            .Font.Size = 9
            .Font.Italic = True
        End With

    Next i

End Sub


' Le champ du sauteur : une cellule, sa liste posee plus
' tard par PoserListeSauteur, qui suit la composition.
Private Sub EcrireChampListe( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String, _
    ByVal NomPlage As String)

    EcrireLibelle ws, Ligne, Libelle

    With ws.Cells(Ligne, COL_PREMIER_BOUTON).Resize(1, 4)

        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter

    End With

    ' La zone entiere, pas son ancre : Excel refuse de
    ' vider une partie d'une cellule fusionnee, et
    ' l'ouverture de la popup s'arreterait la.
    NommerPlage ws, _
        ws.Cells(Ligne, COL_PREMIER_BOUTON) _
            .Resize(1, 4).Address, NomPlage

End Sub


' Le champ libre des observations, large comme la popup.
Private Sub EcrireChampTexte( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal Libelle As String, _
    ByVal NomPlage As String)

    EcrireLibelle ws, Ligne, Libelle

    With ws.Cells(Ligne, COL_PREMIER_BOUTON).Resize(1, 12)

        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlLeft

    End With

    ' Meme raison : la zone entiere, pas son ancre.
    NommerPlage ws, _
        ws.Cells(Ligne, COL_PREMIER_BOUTON) _
            .Resize(1, 12).Address, NomPlage

End Sub


Private Sub EcrireBoutons( _
    ByVal ws As Worksheet, _
    ByVal Ligne As Long, _
    ByVal NomValider As String, _
    ByVal NomAnnuler As String)

    ' Les deux boutons commencent sous les choix, pas
    ' sous leurs intitules.
    With ws.Cells(Ligne, COL_LIBELLE + 1).Resize(1, 2)
        .Merge
        .Value = "ANNULER"
        .Interior.Color = ROUGE
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(Ligne, COL_LIBELLE + 4).Resize(1, 4)
        .Merge
        .Value = "VALIDER  (Entree)"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    NommerPlage ws, _
        ws.Cells(Ligne, COL_LIBELLE + 1).Address, NomAnnuler

    NommerPlage ws, _
        ws.Cells(Ligne, COL_LIBELLE + 4).Address, NomValider

End Sub


Private Sub MettreEnFormePopup( _
    ByVal ws As Worksheet, _
    ByVal LigneBoutons As Long)

    Dim Ligne As Long

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 16

    ' Douze colonnes de boutons : les onze zones du terrain,
    ' une par colonne, et de quoi loger six options de deux
    ' colonnes sur une meme ligne.
    ws.Range(ws.Columns(COL_PREMIER_BOUTON), _
        ws.Columns(COL_PREMIER_BOUTON + 11)).ColumnWidth = 6.5

    ' Les rangs qui portent quelque chose respirent, les
    ' rangs intercalaires restent fins.
    For Ligne = 4 To LigneBoutons

        If Application.WorksheetFunction.CountA( _
            ws.Range( _
                ws.Cells(Ligne, 2), _
                ws.Cells(Ligne, COL_PREMIER_BOUTON + 11))) > 0 Then

            ws.Rows(Ligne).RowHeight = 20

        Else

            ws.Rows(Ligne).RowHeight = 8

        End If

    Next Ligne

    ws.Rows(LigneBoutons).RowHeight = 26

    ws.Cells.Font.Name = "Calibri"

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function FeuillePopup( _
    ByVal Nom As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Nom)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = Nom

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePopup = ws

End Function


' L'adresse d'une plage deja nommee : les champs libres
' sont fusionnes, et c'est leur etendue qui compte.
Private Function AdresseNommee( _
    ByVal NomPlage As String) As String

    AdresseNommee = _
        ThisWorkbook.Names(NomPlage).RefersToRange.Address

End Function


Private Sub NommerPlage( _
    ByVal ws As Worksheet, _
    ByVal Adresse As String, _
    ByVal NomPlage As String)

    On Error Resume Next
    ThisWorkbook.Names(NomPlage).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=NomPlage, _
        RefersTo:=ws.Range(Adresse)

End Sub


Private Sub SupprimerFeuille(ByVal Nom As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Nom)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ws.Visible = xlSheetVisible
    ws.Delete

End Sub
