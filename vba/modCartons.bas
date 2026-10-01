Attribute VB_Name = "modCartons"
Option Explicit

' =========================================================
' CARTONS
'
' Un carton n'est pas un evenement isole : il ouvre une
' periode pendant laquelle un joueur n'est plus sur le
' terrain. Le journal porte les deux bouts de cette
' periode, le carton et son retour, et c'est lui qui fait
' foi : rien n'est garde a part. La palette retrouve donc
' l'etat du moment meme quand on revient en arriere dans la
' video, comme elle le fait deja pour les remplacements.
'
' Ce module repond a une seule question : quels cartons sont
' en cours a tel instant de la video ? La palette s'en sert
' pour colorer les boutons et refuser les clics, la popup de
' retour pour proposer ce qui peut revenir.
'
' Un carton est reconnu a son libelle, pas a la presence
' d'un joueur : "Carton jaune contre nous" nous designe,
' "contre adversaire" les designe. Le libelle survit a une
' correction faite a la main dans le journal, la colonne
' Joueur beaucoup moins.
'
' Aucun caractere accentue en clair : le VBE de macOS
' importe les .bas en Mac Roman.
' =========================================================

' Camps, tels qu'ils terminent les noms des libelles :
' ACT_CARTON_JAUNE_NS, ACT_RETOUR_CARTON_JAUNE_ADV.
Public Const CAMP_NOUS As String = "NS"
Public Const CAMP_ADV As String = "ADV"

Private Const FEUILLE_JOURNAL As String = "Journal actions"
Private Const TABLEAU_JOURNAL As String = "JournalActions"

' Un evenement retenu du journal, garde sous forme de
' chaine : type | camp | couleur | joueur | temps video.
' Le type vaut "C" pour un carton, "R" pour son retour.
Private Evenements() As String
Private NbEvenements As Long

' Le journal n'est relu que lorsque son nombre de lignes
' change : la palette interroge ce module chaque seconde.
Private LignesLues As Long
Private CacheValide As Boolean


' ---------------------------------------------------------
' Libelles
' ---------------------------------------------------------

Private Function CouleursCarton() As Variant

    CouleursCarton = Array("ROUGE", "JAUNE", "BLANC", "BLEU")

End Function


Private Function LibelleCarton( _
    ByVal Couleur As String, _
    ByVal Camp As String) As String

    LibelleCarton = ValeurNommee( _
        "ACT_CARTON_" & Couleur & "_" & Camp)

End Function


Public Function LibelleRetour( _
    ByVal Couleur As String, _
    ByVal Camp As String) As String

    LibelleRetour = ValeurNommee( _
        "ACT_RETOUR_CARTON_" & Couleur & "_" & Camp)

End Function


' Le carton en minuscules, pour les phrases : "carton
' jaune" plutot que "carton JAUNE".
Public Function CouleurLisible( _
    ByVal Couleur As String) As String

    CouleurLisible = LCase(Couleur)

End Function


' Vrai si cette action est la pose d'un carton, quel que
' soit le camp et la couleur.
'
' Sert a lever le blocage du bouton d'un joueur deja sous
' carton : le second jaune devient rouge, et il faut bien
' pouvoir designer celui qui le prend.
Public Function EstActionCarton( _
    ByVal ActionTexte As String) As Boolean

    Dim Couleurs As Variant
    Dim Camps As Variant
    Dim i As Long
    Dim j As Long

    If Trim(ActionTexte) = "" Then Exit Function

    Couleurs = CouleursCarton
    Camps = Array(CAMP_NOUS, CAMP_ADV)

    For i = 0 To UBound(Couleurs)
        For j = 0 To UBound(Camps)

            If EstLeMemeLibelle( _
                ActionTexte, _
                LibelleCarton( _
                    CStr(Couleurs(i)), CStr(Camps(j)))) Then

                EstActionCarton = True
                Exit Function

            End If

        Next j
    Next i

End Function


' Vide plutot qu'une erreur si le nom manque encore : les
' libelles de retour n'existent qu'une fois
' PreparerLibellesRetourCarton passe.
Private Function ValeurNommee( _
    ByVal NomPlage As String) As String

    Dim Cible As Range

    On Error Resume Next
    Set Cible = ThisWorkbook.Names(NomPlage).RefersToRange
    On Error GoTo 0

    If Cible Is Nothing Then Exit Function

    ValeurNommee = Trim(CStr(Cible.Cells(1, 1).Value))

End Function


' ---------------------------------------------------------
' Lecture du journal
' ---------------------------------------------------------

' Les cartons et les retours du journal, ranges du plus
' ancien au plus recent.
Private Sub ChargerEvenements()

    Dim lo As ListObject
    Dim Donnees As Variant

    Dim LibellesCarton() As String
    Dim LibellesRetour() As String
    Dim Camps() As String
    Dim Couleurs() As String
    Dim NbLibelles As Long

    Dim ColTemps As Long
    Dim ColJoueur As Long
    Dim ColAction As Long

    Dim i As Long
    Dim k As Long

    Dim Action As String
    Dim Joueur As String
    Dim Brut As Variant
    Dim Temps As Double

    NbEvenements = 0
    ReDim Evenements(1 To 1)

    LignesLues = -1
    CacheValide = True

    On Error GoTo Sortie

    Set lo = ThisWorkbook _
        .Worksheets(FEUILLE_JOURNAL) _
        .ListObjects(TABLEAU_JOURNAL)

    LignesLues = lo.ListRows.Count

    If lo.DataBodyRange Is Nothing Then Exit Sub

    ColTemps = lo.ListColumns("Temps video").Index
    ColJoueur = lo.ListColumns("Joueur").Index
    ColAction = lo.ListColumns("Action").Index

    NbLibelles = ConstruireTableLibelles( _
        LibellesCarton, LibellesRetour, Camps, Couleurs)

    ' Tout le journal en une seule lecture : cellule par
    ' cellule, la relecture peserait sur le rafraichissement
    ' de la palette, qui a lieu chaque seconde.
    Donnees = lo.DataBodyRange.Value

    ReDim Evenements(1 To UBound(Donnees, 1) + 1)

    For i = 1 To UBound(Donnees, 1)

        Action = Trim(CStr(Donnees(i, ColAction)))
        Brut = Donnees(i, ColTemps)

        If Action <> "" Then

            If Len(Trim(CStr(Brut))) > 0 Then

                ' Le format [m]:ss peut rendre une date
                ' plutot qu'un nombre : CDbl les ramene tous
                ' deux a la fraction de jour attendue.
                If IsNumeric(Brut) Or IsDate(Brut) Then

                    Temps = CDbl(Brut) * 86400#
                    Joueur = Trim(CStr(Donnees(i, ColJoueur)))

                    For k = 1 To NbLibelles

                        If EstLeMemeLibelle( _
                            Action, LibellesCarton(k)) Then

                            AjouterEvenement "C", Camps(k), _
                                Couleurs(k), Joueur, Temps

                            Exit For

                        End If

                        If EstLeMemeLibelle( _
                            Action, LibellesRetour(k)) Then

                            AjouterEvenement "R", Camps(k), _
                                Couleurs(k), Joueur, Temps

                            Exit For

                        End If

                    Next k

                End If

            End If

        End If

    Next i

    TrierEvenements

Sortie:

End Sub


' Les huit cartons et les huit retours : quatre couleurs
' pour chacun des deux camps.
Private Function ConstruireTableLibelles( _
    ByRef LibellesCarton() As String, _
    ByRef LibellesRetour() As String, _
    ByRef Camps() As String, _
    ByRef Couleurs() As String) As Long

    Dim ListeCouleurs As Variant
    Dim ListeCamps As Variant
    Dim i As Long
    Dim j As Long
    Dim n As Long

    ListeCouleurs = CouleursCarton
    ListeCamps = Array(CAMP_NOUS, CAMP_ADV)

    n = (UBound(ListeCouleurs) + 1) * (UBound(ListeCamps) + 1)

    ReDim LibellesCarton(1 To n)
    ReDim LibellesRetour(1 To n)
    ReDim Camps(1 To n)
    ReDim Couleurs(1 To n)

    n = 0

    For i = 0 To UBound(ListeCouleurs)
        For j = 0 To UBound(ListeCamps)

            n = n + 1

            Couleurs(n) = CStr(ListeCouleurs(i))
            Camps(n) = CStr(ListeCamps(j))

            LibellesCarton(n) = _
                LibelleCarton(Couleurs(n), Camps(n))

            LibellesRetour(n) = _
                LibelleRetour(Couleurs(n), Camps(n))

        Next j
    Next i

    ConstruireTableLibelles = n

End Function


' Un libelle absent de Parametres vaut la chaine vide : il
' ne doit alors rien reconnaitre du tout.
Private Function EstLeMemeLibelle( _
    ByVal Action As String, _
    ByVal Libelle As String) As Boolean

    If Libelle = "" Then Exit Function

    EstLeMemeLibelle = _
        (StrComp(Action, Libelle, vbTextCompare) = 0)

End Function


Private Sub AjouterEvenement( _
    ByVal TypeEvenement As String, _
    ByVal Camp As String, _
    ByVal Couleur As String, _
    ByVal Joueur As String, _
    ByVal Temps As Double)

    NbEvenements = NbEvenements + 1

    Evenements(NbEvenements) = _
        TypeEvenement & "|" & _
        Camp & "|" & _
        Couleur & "|" & _
        Joueur & "|" & _
        CStr(Temps)

End Sub


' Le journal range ses lignes de la plus recente a la plus
' ancienne. L'appariement d'un carton avec son retour, lui,
' se fait dans l'ordre du match.
Private Sub TrierEvenements()

    Dim i As Long
    Dim j As Long
    Dim Pivot As String

    For i = 2 To NbEvenements

        Pivot = Evenements(i)
        j = i - 1

        Do While j >= 1
            If Not EstAvant(Pivot, Evenements(j)) Then Exit Do

            Evenements(j + 1) = Evenements(j)
            j = j - 1
        Loop

        Evenements(j + 1) = Pivot

    Next i

End Sub


' Le temps d'abord, et a temps egal le carton avant son
' retour.
'
' Les deux portent la meme seconde des que la video est en
' pause : un carton pose puis rendu sans avoir avance d'une
' image doit quand meme s'annuler, et non laisser le joueur
' bloque. Sans cette regle, l'ordre du journal decidait, et
' il va du plus recent au plus ancien.
Private Function EstAvant( _
    ByVal Gauche As String, _
    ByVal Droite As String) As Boolean

    If TempsCarton(Gauche) < TempsCarton(Droite) Then
        EstAvant = True
        Exit Function
    End If

    If TempsCarton(Gauche) > TempsCarton(Droite) Then
        Exit Function
    End If

    EstAvant = _
        (TypeEvenement(Gauche) = "C") _
        And (TypeEvenement(Droite) = "R")

End Function


Private Sub AssurerCache()

    If CacheValide Then
        If NbLignesJournal = LignesLues Then Exit Sub
    End If

    ChargerEvenements

End Sub


Private Function NbLignesJournal() As Long

    NbLignesJournal = -1

    On Error GoTo Sortie

    NbLignesJournal = ThisWorkbook _
        .Worksheets(FEUILLE_JOURNAL) _
        .ListObjects(TABLEAU_JOURNAL) _
        .ListRows.Count

Sortie:

End Function


' Le journal a change autrement qu'en gagnant une ligne :
' la prochaine question le relira.
Public Sub InvaliderCacheCartons()

    CacheValide = False

End Sub


' ---------------------------------------------------------
' Cartons en cours
' ---------------------------------------------------------

' Les cartons poses avant cet instant de la video et dont le
' retour n'a pas encore ete journalise. Un carton rouge ou
' bleu n'en sort jamais : personne ne declarera son retour.
Public Function CartonsEnCours( _
    ByVal TempsVideo As Double) As Collection

    Dim Resultat As Collection
    Dim i As Long
    Dim j As Long
    Dim Rendu As Long

    Set Resultat = New Collection

    AssurerCache

    For i = 1 To NbEvenements

        If TempsCarton(Evenements(i)) > TempsVideo Then
            Exit For
        End If

        If TypeEvenement(Evenements(i)) = "C" Then

            Resultat.Add Evenements(i)

        Else

            ' Le retour efface le plus ancien carton de meme
            ' camp, de meme couleur et de meme joueur.
            Rendu = 0

            For j = 1 To Resultat.Count

                If CleCarton(Resultat(j)) = _
                    CleCarton(Evenements(i)) Then

                    Rendu = j
                    Exit For

                End If

            Next j

            If Rendu > 0 Then Resultat.Remove Rendu

        End If

    Next i

    Set CartonsEnCours = Resultat

End Function


' La couleur du carton que porte un joueur, la plus grave
' s'il en cumule. Vide s'il n'en porte aucun.
Public Function CartonDeJoueurDansListe( _
    ByVal Liste As Collection, _
    ByVal Joueur As String) As String

    Dim i As Long
    Dim Pire As Long
    Dim Rang As Long

    If Trim(Joueur) = "" Then Exit Function

    ' Un carton pose sans designer personne va au journal
    ' avec "?". Il ne doit pas pour autant bloquer le bouton
    ' "?" de la palette, ni celui du collectif.
    If Joueur = "?" Then Exit Function
    If Joueur = "Collectif" Then Exit Function

    For i = 1 To Liste.Count

        If CampCarton(Liste(i)) = CAMP_NOUS Then

            If StrComp( _
                JoueurCarton(Liste(i)), _
                Joueur, _
                vbTextCompare) = 0 Then

                Rang = Gravite(CouleurCarton(Liste(i)))

                If Rang > Pire Then

                    Pire = Rang

                    CartonDeJoueurDansListe = _
                        CouleurCarton(Liste(i))

                End If

            End If

        End If

    Next i

End Function


' Le carton que porte l'occupant d'un emplacement de la
' palette, au temps que le chrono affiche.
'
' Le chrono plutot que le moteur video : la verification a
' lieu a chaque clic, et c'est de toute facon l'etat montre
' a l'ecran que l'on veut faire respecter.
Public Function CartonActuelDuPoste( _
    ByVal Poste As String) As String

    If Trim(Poste) = "" Then Exit Function

    CartonActuelDuPoste = CartonDeJoueurDansListe( _
        CartonsEnCours(TempsChronoAffiche), _
        GetPlayerName(Poste))

End Function


Private Function TempsChronoAffiche() As Double

    Dim Valeur As Variant

    On Error GoTo Sortie

    Valeur = shSaisieVideo.Range("CELL_CHRONO_VIDEO").Value

    If IsNumeric(Valeur) Then
        TempsChronoAffiche = CDbl(Valeur) * 86400#
    End If

Sortie:

End Function


' Habille une cellule aux couleurs d'un carton. La palette
' de saisie et la composition de la popup penalite s'en
' servent toutes deux : une seule table de couleurs, pas
' deux qui finiraient par diverger.
'
' Le blanc pose sa bordure : sur fond blanc, celle d'origine
' se verrait a peine et la case se fondrait dans la feuille.
Public Sub AppliquerCouleursCarton( _
    ByVal Zone As Range, _
    ByVal Couleur As String)

    Select Case Couleur

        Case "ROUGE"

            Zone.Interior.Color = RGB(192, 0, 0)
            Zone.Font.Color = RGB(255, 255, 255)

        Case "JAUNE"

            Zone.Interior.Color = RGB(255, 255, 0)
            Zone.Font.Color = RGB(0, 0, 0)

        Case "BLANC"

            Zone.Interior.Color = RGB(255, 255, 255)
            Zone.Font.Color = RGB(0, 0, 0)

            Zone.Borders.LineStyle = xlContinuous
            Zone.Borders.Color = RGB(0, 0, 0)

        Case "BLEU"

            Zone.Interior.Color = RGB(0, 176, 240)
            Zone.Font.Color = RGB(255, 255, 255)

    End Select

End Sub


Private Function Gravite(ByVal Couleur As String) As Long

    Select Case Couleur

        Case "ROUGE"
            Gravite = 4

        Case "BLEU"
            Gravite = 3

        Case "JAUNE"
            Gravite = 2

        Case "BLANC"
            Gravite = 1

    End Select

End Function


' ---------------------------------------------------------
' Lecture d'un evenement
' ---------------------------------------------------------

Private Function TypeEvenement( _
    ByVal Evenement As String) As String

    TypeEvenement = Split(Evenement, "|")(0)

End Function


Public Function CampCarton( _
    ByVal Evenement As String) As String

    CampCarton = Split(Evenement, "|")(1)

End Function


Public Function CouleurCarton( _
    ByVal Evenement As String) As String

    CouleurCarton = Split(Evenement, "|")(2)

End Function


Public Function JoueurCarton( _
    ByVal Evenement As String) As String

    JoueurCarton = Split(Evenement, "|")(3)

End Function


Public Function TempsCarton( _
    ByVal Evenement As String) As Double

    TempsCarton = CDbl(Split(Evenement, "|")(4))

End Function


' Camp, couleur et joueur : ce qui fait qu'un retour parle
' bien du carton pose plus tot.
Private Function CleCarton( _
    ByVal Evenement As String) As String

    Dim Morceaux() As String

    Morceaux = Split(Evenement, "|")

    CleCarton = Morceaux(1) & "|" & _
        Morceaux(2) & "|" & _
        UCase(Morceaux(3))

End Function


' =========================================================
' PREPARATION DU CLASSEUR
'
' A lancer une fois. Deux choses :
'
'   les huit libelles de retour rejoignent
'   TblActionsJournal, ce qui les met du meme coup dans la
'   liste deroulante du journal ;
'
'   tous les noms ACT_ sont recrees depuis ce tableau.
'   ACT_CARTON_ROUGE_NS pointait sur le libelle de
'   l'adversaire : un carton rouge contre nous partait au
'   journal comme un carton contre eux.
' =========================================================

Public Sub PreparerLibellesRetourCarton()

    Dim lo As ListObject
    Dim Ajoutes As Long
    Dim Nommes As Long

    On Error GoTo GestionErreur

    Set lo = shParametres.ListObjects("TblActionsJournal")

    Ajoutes = AjouterLibellesManquants(lo)
    Nommes = NommerActionsDepuisTableau(lo)

    InvaliderCacheCartons

    If ModeSilencieux Then Exit Sub

    MsgBox _
        Ajoutes & " libell" & ChrW(233) & "s ajout" & _
        ChrW(233) & "s, " & Nommes & " noms recr" & _
        ChrW(233) & ChrW(233) & "s.", _
        vbInformation, _
        "Cartons"

    Exit Sub

GestionErreur:

    MsgBox _
        "La pr" & ChrW(233) & "paration a " & ChrW(233) & _
        "chou" & ChrW(233) & "." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Cartons"

End Sub


Private Function AjouterLibellesManquants( _
    ByVal lo As ListObject) As Long

    Dim Couleurs As Variant
    Dim Camps As Variant
    Dim i As Long
    Dim j As Long

    Dim Nom As String
    Dim Manquants As Collection
    Dim Element As Variant

    Dim PremiereLigne As Long
    Dim Ligne As Long

    Set Manquants = New Collection

    Couleurs = CouleursCarton
    Camps = Array(CAMP_NOUS, CAMP_ADV)

    For i = 0 To UBound(Couleurs)
        For j = 0 To UBound(Camps)

            Nom = "ACT_RETOUR_CARTON_" & _
                CStr(Couleurs(i)) & "_" & CStr(Camps(j))

            If LigneDuNom(lo, Nom) = 0 Then
                Manquants.Add Nom & "|" & _
                    LibelleRetourAttendu( _
                        CStr(Couleurs(i)), CStr(Camps(j)))
            End If

        Next j
    Next i

    AjouterLibellesManquants = Manquants.Count

    If Manquants.Count = 0 Then Exit Function

    ' Le tableau s'etend sans inserer de lignes : une
    ' insertion decalerait ce qui vit plus bas sur la
    ' feuille.
    PremiereLigne = lo.Range.Row + lo.Range.Rows.Count

    lo.Resize lo.Range.Resize( _
        lo.Range.Rows.Count + Manquants.Count)

    Ligne = PremiereLigne

    For Each Element In Manquants

        shParametres.Cells(Ligne, lo.Range.Column).Value = _
            Split(CStr(Element), "|")(0)

        shParametres.Cells(Ligne, lo.Range.Column + 1).Value = _
            Split(CStr(Element), "|")(1)

        Ligne = Ligne + 1

    Next Element

End Function


' Les libelles tels qu'ils doivent s'ecrire dans la
' feuille. Les accents y sont permis, c'est une valeur de
' cellule et non du code.
Private Function LibelleRetourAttendu( _
    ByVal Couleur As String, _
    ByVal Camp As String) As String

    Dim Qui As String

    If Camp = CAMP_NOUS Then
        Qui = "nous"
    Else
        Qui = "adversaire"
    End If

    LibelleRetourAttendu = _
        "Retour de carton " & LCase(Couleur) & _
        " - " & Qui

End Function


' La ligne de la feuille ou le tableau porte ce nom, zero
' s'il n'y figure pas.
Private Function LigneDuNom( _
    ByVal lo As ListObject, _
    ByVal Nom As String) As Long

    Dim i As Long
    Dim Cellule As Range

    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count

        Set Cellule = lo.DataBodyRange.Cells(i, 1)

        If StrComp(Trim(CStr(Cellule.Value)), Nom, _
            vbTextCompare) = 0 Then

            LigneDuNom = Cellule.Row
            Exit Function

        End If

    Next i

End Function


' Chaque ligne du tableau donne un nom en premiere colonne
' et le libelle qu'il doit designer en seconde.
Private Function NommerActionsDepuisTableau( _
    ByVal lo As ListObject) As Long

    Dim i As Long
    Dim Nom As String
    Dim Cible As Range
    Dim n As Long

    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count

        Nom = Trim(CStr(lo.DataBodyRange.Cells(i, 1).Value))

        If Nom <> "" Then

            Set Cible = lo.DataBodyRange.Cells(i, 2)

            On Error Resume Next
            ThisWorkbook.Names(Nom).Delete
            On Error GoTo 0

            ThisWorkbook.Names.Add _
                Name:=Nom, _
                RefersTo:=Cible

            n = n + 1

        End If

    Next i

    NommerActionsDepuisTableau = n

End Function
