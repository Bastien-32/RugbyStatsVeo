Attribute VB_Name = "modRecapBase"
Option Explicit

' =========================================================
' RECAPITULATIF DE SAISON - SOCLE COMMUN
'
' Constantes, acces aux feuilles et petits utilitaires
' partages par les trois autres modules :
'
'   modRecapConstruction : fabrique les feuilles ;
'   modRecapImport       : ajoute les matchs nouveaux ;
'   modRecapCalcul       : applique les filtres et ecrit
'                          le classement.
'
' Ce classeur vit dans le dossier Matchs, a cote des
' fichiers de championnat. Il garde chez lui une copie du
' journal d'actions et des compositions de chaque match
' deja ajoute : une fois un match integre, son fichier
' n'est plus jamais rouvert.
' =========================================================

' ---------- Feuilles ----------

Public Const FEUILLE_CLASSEMENT As String = "Classement"

Public Const FEUILLE_TEMPS As String = "Temps de jeu"

Public Const FEUILLE_FILTRES As String = "Filtres"

Public Const FEUILLE_MATCHS As String = "Matchs"

Public Const FEUILLE_JOURNAL As String = "Journal actions"

Public Const FEUILLE_COMPOSITIONS As String = "Compositions"

Public Const FEUILLE_PARAMETRES As String = "Parametres"

' ---------- Tableaux structures ----------

Public Const TBL_MATCHS As String = "RECAP_MATCHS"

Public Const TBL_JOURNAL As String = "RECAP_JOURNAL"

Public Const TBL_COMPOSITIONS As String = "RECAP_COMPOSITIONS"

Public Const TBL_STATS As String = "RECAP_STATS"

' ---------- Colonnes de RECAP_MATCHS ----------

Public Const MTC_CIBLER As Long = 1

Public Const MTC_RETENU As Long = 2

Public Const MTC_ID As Long = 3

Public Const MTC_FICHIER As Long = 4

Public Const MTC_SAISON As Long = 5

Public Const MTC_DATE As Long = 6

Public Const MTC_CATEGORIE As Long = 7

Public Const MTC_ADVERSAIRE As Long = 8

Public Const MTC_LIEU As Long = 9

Public Const MTC_PHASE As Long = 10

Public Const MTC_JOURNEE As Long = 11

Public Const MTC_RESULTAT As Long = 12

Public Const MTC_ACTIONS As Long = 13

Public Const MTC_JOUEURS As Long = 14

Public Const MTC_MINUTES As Long = 15

Public Const MTC_AJOUTE As Long = 16

' Date de modification du fichier de match au moment ou il
' a ete lu. C'est en la comparant a celle du fichier qu'on
' sait, sans rien ouvrir, qu'une saisie a ete completee.
Public Const MTC_MODIFIE As Long = 17

Public Const NB_COLONNES_MATCHS As Long = 17

' ---------- Colonnes de RECAP_JOURNAL ----------

Public Const JRN_ID As Long = 1

Public Const JRN_FICHIER As Long = 2

Public Const JRN_JOUEUR As Long = 12

Public Const JRN_ACTION As Long = 14

Public Const JRN_MOTIF As Long = 15

Public Const NB_COLONNES_JOURNAL As Long = 17

' ---------- Colonnes de RECAP_COMPOSITIONS ----------

Public Const CMP_ID As Long = 1

Public Const CMP_FICHIER As Long = 2

Public Const CMP_CATEGORIE As Long = 5

Public Const CMP_JOUEUR As Long = 11

Public Const CMP_LIGNE As Long = 13

Public Const CMP_MINUTES As Long = 14

Public Const NB_COLONNES_COMPOSITIONS As Long = 14

' ---------- Reperes dans les fichiers de match ----------

Public Const SRC_FEUILLE_COMPO As String = "Compo"

Public Const SRC_FEUILLE_JOURNAL As String = "Journal actions"

Public Const SRC_TBL_JOURNAL As String = "JournalActions"

Public Const SRC_TBL_COMPO As String = "COMPO"

' Categorie de l'equipe : seule information du match qui
' ne soit pas recopiee sur chaque ligne du journal.
Public Const SRC_CELLULE_CATEGORIE As String = "C20"

' ---------- Mise en page ----------

' Ligne d'en-tete des tableaux fabriques par le module de
' construction. Au-dessus : le titre en 2, le sous-titre en
' 3, les boutons en 4, et en 5 le rappel des filtres
' appliques.
Public Const LIGNE_ENTETE As Long = 6

Public Const PREMIERE_COLONNE As Long = 2

Public Const MODE_TOTAUX As String = "Totaux"

Public Const MODE_PAR_MATCH As String = "Par match"

Public Const MODE_PAR_80 As String = "Par 80 minutes"

Public Const VALEUR_TOUS As String = "(Tous)"

Public Const MARQUE_COCHE As String = "X"

' Un joueur ne peut pas depasser le temps reglementaire.
' Au-dela, la valeur lue dans Compo n'est pas un nombre de
' minutes mais autre chose, et elle est ignoree.
Public Const MINUTES_MAXI As Double = 200

' Deux secondes, exprimees en fraction de jour. En deca, un
' ecart entre la date de modification d'un fichier et celle
' qu'on avait notee ne prouve rien : les horloges de
' fichiers s'arrondissent, et une synchronisation iCloud
' peut retoucher la date sans que le contenu bouge.
Public Const TOLERANCE_DATE As Double = 2# / 86400#


' =========================================================
' JOURNAL DE DIAGNOSTIC
'
' Un arret brutal d'Excel ne laisse ni message ni pile
' d'appels : il faut donc savoir ou on en etait. Chaque
' etape est ecrite dans recap-trace.txt, a cote du
' classeur, et le fichier est referme aussitot pour que la
' ligne soit sur le disque avant l'etape suivante.
'
' A passer a False une fois la mise au point terminee.
' =========================================================
Public Const TRACE_ACTIVE As Boolean = True

Public Const FICHIER_TRACE As String = "recap-trace.txt"


Public Sub Tracer( _
    ByVal Etape As String _
)

    Dim Sortie As Integer

    If Not TRACE_ACTIVE Then
        Exit Sub
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        Exit Sub
    End If

    ' Le diagnostic ne doit jamais faire echouer ce qu'il
    ' observe.
    On Error Resume Next

    Sortie = FreeFile

    Open ThisWorkbook.Path & Application.PathSeparator & _
        FICHIER_TRACE For Append As #Sortie

    Print #Sortie, Format$(Now, "hh:mm:ss") & "  " & Etape

    Close #Sortie

    Err.Clear

    On Error GoTo 0

End Sub


' =========================================================
' Renvoie une feuille du classeur recapitulatif, ou Nothing
' si elle n'existe pas encore.
' =========================================================
Public Function FeuilleRecap( _
    ByVal NomFeuille As String _
) As Worksheet

    On Error Resume Next

    Set FeuilleRecap = ThisWorkbook.Worksheets(NomFeuille)

    On Error GoTo 0

End Function


' =========================================================
' Renvoie une feuille en la creant au besoin, placee en fin
' de classeur. Les feuilles sont ainsi toujours dans
' l'ordre ou le module de construction les demande.
' =========================================================
Public Function FeuilleRecapOuCreee( _
    ByVal NomFeuille As String _
) As Worksheet

    Dim ws As Worksheet

    Set ws = FeuilleRecap(NomFeuille)

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count _
                ) _
            )

        ws.Name = NomFeuille

    End If

    Set FeuilleRecapOuCreee = ws

End Function


' =========================================================
' Renvoie un tableau structure du classeur recapitulatif,
' ou Nothing s'il n'a pas encore ete fabrique.
' =========================================================
Public Function TableauRecap( _
    ByVal NomFeuille As String, _
    ByVal NomTableau As String _
) As ListObject

    Dim ws As Worksheet

    Set ws = FeuilleRecap(NomFeuille)

    If ws Is Nothing Then
        Exit Function
    End If

    On Error Resume Next

    Set TableauRecap = ws.ListObjects(NomTableau)

    On Error GoTo 0

End Function


' =========================================================
' Contenu d'un tableau structure, ou Empty s'il est vide.
' =========================================================
Public Function DonneesTableau( _
    ByVal NomFeuille As String, _
    ByVal NomTableau As String _
) As Variant

    Dim Tableau As ListObject

    Set Tableau = TableauRecap(NomFeuille, NomTableau)

    If Tableau Is Nothing Then
        Exit Function
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Function
    End If

    DonneesTableau = Tableau.DataBodyRange.Value

End Function


' =========================================================
' Vide un tableau structure de ses donnees en gardant une
' ligne vide, pour qu'il conserve son nom et son habillage.
' =========================================================
Public Sub ViderTableauRecap( _
    ByVal Tableau As ListObject _
)

    If Tableau Is Nothing Then
        Exit Sub
    End If

    If Tableau.ListRows.Count = 0 Then
        Exit Sub
    End If

    Tableau.DataBodyRange.Delete

End Sub


' =========================================================
' Ajoute des lignes vides a la fin d'un tableau structure
' et renvoie la premiere ligne ajoutee, en numero de ligne
' de feuille. Ecrire d'un bloc dans la plage ainsi obtenue
' est bien plus rapide que d'appeler ListRows.Add pour
' chaque ligne.
' =========================================================
Public Function PreparerLignesTableau( _
    ByVal Tableau As ListObject, _
    ByVal NbLignes As Long _
) As Long

    Dim Existantes As Long

    If NbLignes <= 0 Then
        Exit Function
    End If

    Existantes = Tableau.ListRows.Count

    ' Un tableau neuf porte une ligne vide, laissee par sa
    ' creation : on ecrit dedans plutot qu'en dessous, sans
    ' quoi la base commencerait par une ligne blanche que
    ' les filtres compteraient comme un match.
    If Existantes = 1 Then

        If Application.CountA(Tableau.DataBodyRange) = 0 Then
            Existantes = 0
        End If

    End If

    If Existantes = 0 Then

        If Tableau.ListRows.Count = 0 Then
            Tableau.ListRows.Add
        End If

        PreparerLignesTableau = Tableau.DataBodyRange.Row

        If NbLignes > 1 Then

            Tableau.Resize _
                Tableau.Range.Resize(NbLignes + 1)

        End If

    Else

        PreparerLignesTableau = _
            Tableau.DataBodyRange.Row + Existantes

        ' La plage du tableau compte aussi sa ligne
        ' d'en-tete.
        Tableau.Resize _
            Tableau.Range.Resize(Existantes + NbLignes + 1)

    End If

End Function


' =========================================================
' Ecrit un bloc de valeurs a la fin d'un tableau structure.
' =========================================================
Public Sub AjouterLignesTableau( _
    ByVal Tableau As ListObject, _
    ByRef Valeurs As Variant, _
    ByVal NbLignes As Long, _
    ByVal NbColonnes As Long _
)

    Dim PremiereLigne As Long

    If NbLignes <= 0 Then
        Exit Sub
    End If

    PremiereLigne = PreparerLignesTableau(Tableau, NbLignes)

    Tableau.Parent.Cells( _
        PremiereLigne, _
        Tableau.Range.Column _
    ).Resize(NbLignes, NbColonnes).Value = Valeurs

End Sub


' =========================================================
' Rapproche deux libelles en ignorant la casse, les
' accents, les espaces et la ponctuation.
'
' Les fichiers de match n'ecrivent pas toujours leurs
' en-tetes de la meme facon d'une saison a l'autre :
' "temps video" ici, "Temps vid" & ChrW(233) & "o" ailleurs.
' Chercher une colonne sur son nom exact reviendrait a
' casser l'import au premier accent deplace.
' =========================================================
Public Function Normaliser( _
    ByVal Texte As String _
) As String

    ' Les deux tables ne peuvent pas etre des constantes :
    ' VBA exige qu'une constante soit connue a la
    ' compilation, et ChrW s'evalue a l'execution. Static
    ' les construit au premier appel seulement, ce qui
    ' compte : cette fonction est appelee des milliers de
    ' fois par calcul.
    Static Accentues As String
    Static Simples As String

    Dim Resultat As String
    Dim Caractere As String
    Dim Position As Long
    Dim i As Long

    If Len(Accentues) = 0 Then

        Accentues = _
            ChrW(224) & ChrW(226) & ChrW(228) & ChrW(233) & _
            ChrW(232) & ChrW(234) & ChrW(235) & ChrW(238) & _
            ChrW(239) & ChrW(244) & ChrW(246) & ChrW(249) & _
            ChrW(251) & ChrW(252) & ChrW(231)

        Simples = "aaaeeeeiioouuuc"

    End If

    Resultat = LCase$(Trim$(Texte))

    For i = 1 To Len(Resultat)

        Caractere = Mid$(Resultat, i, 1)

        Position = InStr(1, Accentues, Caractere, vbTextCompare)

        If Position > 0 Then

            Normaliser = Normaliser & Mid$(Simples, Position, 1)

        ElseIf (Caractere >= "a" And Caractere <= "z") _
            Or (Caractere >= "0" And Caractere <= "9") Then

            Normaliser = Normaliser & Caractere

        End If

    Next i

End Function


' =========================================================
' Index d'une colonne dans un tableau structure, cherchee
' sur son nom normalise. Renvoie 0 si elle est absente.
' =========================================================
Public Function IndexColonneSouple( _
    ByVal Tableau As ListObject, _
    ByVal NomColonne As String _
) As Long

    Dim Recherche As String
    Dim i As Long

    Recherche = Normaliser(NomColonne)

    For i = 1 To Tableau.ListColumns.Count

        If Normaliser(Tableau.ListColumns(i).Name) = Recherche Then

            IndexColonneSouple = i
            Exit Function

        End If

    Next i

End Function


' =========================================================
' Plus grand de deux entiers. Ecrit ici plutot qu'appele
' sur WorksheetFunction : ces bornes servent a dimensionner
' des tableaux, et un ReDim ne doit pas dependre du moteur
' de calcul de la feuille.
' =========================================================
Public Function Maxi( _
    ByVal Premier As Long, _
    ByVal Second As Long _
) As Long

    If Premier > Second Then
        Maxi = Premier
    Else
        Maxi = Second
    End If

End Function


' =========================================================
' Vrai si la cle figure deja dans la collection. Excel pour
' Mac ne fournit pas Scripting.Dictionary : les index par
' cle du classeur passent donc par une Collection.
' =========================================================
Public Function CleExiste( _
    ByVal Liste As Collection, _
    ByVal Cle As String _
) As Boolean

    Dim Valeur As Variant

    On Error Resume Next

    Valeur = Liste.Item(Cle)

    CleExiste = (Err.Number = 0)

    Err.Clear

    On Error GoTo 0

End Function


' =========================================================
' Vrai si la cellule porte la marque de coche, quelle que
' soit la casse et les espaces autour.
' =========================================================
Public Function EstCoche( _
    ByVal Valeur As Variant _
) As Boolean

    EstCoche = _
        (StrComp( _
            Trim(CStr(Valeur)), _
            MARQUE_COCHE, _
            vbTextCompare _
        ) = 0)

End Function


' =========================================================
' Convertit en minutes ce qui a ete saisi dans la colonne
' "temps de jeu" de Compo.
'
' Deux ecritures sont acceptees : un nombre de minutes
' (80) ou une duree Excel (1:20, stockee en fraction de
' jour). Une duree de match etant toujours inferieure a un
' jour, une valeur strictement inferieure a 1 ne peut etre
' qu'une duree ; au-dela, ce sont des minutes.
'
' Renvoie -1 quand la cellule est vide ou illisible, pour
' distinguer "pas saisi" de "zero minute".
' =========================================================
Public Function MinutesDepuisCompo( _
    ByVal Valeur As Variant _
) As Double

    Dim Nombre As Double
    Dim Texte As String

    MinutesDepuisCompo = -1

    If IsEmpty(Valeur) Then
        Exit Function
    End If

    If IsError(Valeur) Then
        Exit Function
    End If

    Texte = Trim(CStr(Valeur))

    If Len(Texte) = 0 Then
        Exit Function
    End If

    If Not IsNumeric(Valeur) Then

        ' Une saisie du type "1:20" arrive en texte quand
        ' la cellule n'est pas au format horaire.
        If Not IsDate(Valeur) Then
            Exit Function
        End If

        Nombre = CDbl(CDate(Valeur))

    Else

        Nombre = CDbl(Valeur)

    End If

    If Nombre < 0 Then
        Exit Function
    End If

    If Nombre < 1 Then
        Nombre = Nombre * 1440
    End If

    If Nombre > MINUTES_MAXI Then
        Exit Function
    End If

    MinutesDepuisCompo = Nombre

End Function


' =========================================================
' Texte d'une cellule, debarrasse des espaces et des
' erreurs de formule.
' =========================================================
Public Function TexteCellule( _
    ByVal Valeur As Variant _
) As String

    If IsError(Valeur) Then
        Exit Function
    End If

    If IsEmpty(Valeur) Then
        Exit Function
    End If

    TexteCellule = Trim(CStr(Valeur))

End Function


' =========================================================
' Valeur numerique d'une cellule, ou zero si elle n'en
' porte pas.
' =========================================================
Public Function NombreCellule( _
    ByVal Valeur As Variant _
) As Double

    If IsError(Valeur) Then
        Exit Function
    End If

    If Not IsNumeric(Valeur) Then
        Exit Function
    End If

    NombreCellule = CDbl(Valeur)

End Function


' =========================================================
' Le journal impute certaines actions a l'equipe entiere,
' ou a un joueur non identifie. Elles ne peuvent pas entrer
' dans un classement individuel.
' =========================================================
Public Function EstJoueur( _
    ByVal Nom As String _
) As Boolean

    If Len(Nom) = 0 Then
        Exit Function
    End If

    If Nom = "?" Then
        Exit Function
    End If

    If Normaliser(Nom) = "collectif" Then
        Exit Function
    End If

    EstJoueur = True

End Function
