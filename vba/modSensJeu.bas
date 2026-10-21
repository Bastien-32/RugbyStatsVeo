Attribute VB_Name = "modSensJeu"
Option Explicit

' =========================================================
' SENS DE JEU DE NOTRE EQUIPE
'
' Les popups touche et melee dessinent le terrain de
' gauche a droite. Encore faut-il savoir de quel cote se
' trouve notre camp sur la video : il change a la mi-temps,
' et d'un match a l'autre.
'
' L'etat tient dans une cellule plutot que dans une
' variable : il doit survivre a la fermeture du classeur,
' et les popups le relisent a chaque ouverture.
'
' Faux  : notre camp est a gauche de l'ecran, nous
'         attaquons vers la droite. C'est le sens d'origine.
' Vrai  : l'inverse.
'
' Rien de tout cela ne change ce qui est enregistre : une
' touche a "N 22m" reste a "N 22m". Seule la place de la
' case a l'ecran change, et le couloir des melees, qui se
' lit toujours depuis notre equipe.
' =========================================================

Private Const NOM_ETAT As String = "SENS_JEU"

' Une colonne libre de Parametres, loin des listes.
Private Const CELLULE_ETAT As String = "$V$1"

Private Const NOM_BOUTON As String = "BTN_SENS_JEU"


' ---------------------------------------------------------
' Lecture
' ---------------------------------------------------------

Public Function SensJeuInverse() As Boolean

    On Error GoTo Defaut

    SensJeuInverse = (UCase$(Trim$(CStr( _
        ThisWorkbook.Names(NOM_ETAT) _
            .RefersToRange.Value))) = "G")

    Exit Function

Defaut:

    SensJeuInverse = False

End Function


' ---------------------------------------------------------
' Ecriture
' ---------------------------------------------------------

Public Sub DefinirSensJeu(ByVal Inverse As Boolean)

    Dim EtatEvenements As Boolean

    EtatEvenements = Application.EnableEvents

    On Error GoTo Sortie

    Application.EnableEvents = False

    AssurerCelluleEtat

    ThisWorkbook.Names(NOM_ETAT).RefersToRange.Value = _
        IIf(Inverse, "G", "D")

    AfficherSensJeu

Sortie:

    Application.EnableEvents = EtatEvenements

End Sub


Public Sub BasculerSensJeu()

    DefinirSensJeu Not SensJeuInverse()

End Sub


' La mi-temps renverse le terrain : les equipes changent
' de cote. Le bouton reste corrigeable, dans la feuille de
' saisie comme dans les deux popups.
Public Sub AppliquerSensJeuMiTemps(ByVal MiTemps As String)

    DefinirSensJeu (UCase$(Trim$(MiTemps)) = "MT2")

End Sub


' ---------------------------------------------------------
' Affichage du bouton
' ---------------------------------------------------------

Public Sub AfficherSensJeu()

    Const TAILLE_FLECHE As Double = 22

    Dim Cible As Range
    Dim Texte As String

    On Error GoTo Sortie

    Set Cible = ThisWorkbook.Names(NOM_BOUTON).RefersToRange

    Texte = LibelleSensJeu()

    Cible.Cells(1, 1).Value = Texte

    ' La fleche est le dernier caractere : elle seule est
    ' agrandie, le libelle gardant la taille du style.
    ' Characters ne travaille que sur une cellule dont le
    ' contenu est deja ecrit, d'ou l'ordre.
    Cible.Cells(1, 1).Characters( _
        Len(Texte), 1).Font.Size = TAILLE_FLECHE

Sortie:

End Sub


Public Function LibelleSensJeu() As String

    ' Fleche pleine vers la droite, puis vers la gauche.
    Dim Fleche As String

    If SensJeuInverse() Then
        Fleche = ChrW(11013)
    Else
        Fleche = ChrW(10145)
    End If

    LibelleSensJeu = _
        "Sens de jeu de votre " & ChrW(233) & "quipe" & _
        vbLf & Fleche

End Function


' ---------------------------------------------------------
' Installation
'
' Cree la cellule d'etat et nomme le bouton. A lancer une
' fois : les deux noms survivent ensuite dans le classeur.
' ---------------------------------------------------------

Public Sub InstallerSensJeu()

    Dim ws As Worksheet

    AssurerCelluleEtat

    Set ws = shSaisieVideo

    ThisWorkbook.Names.Add _
        Name:=NOM_BOUTON, _
        RefersTo:="='" & ws.Name & "'!$BP$4"

    AfficherSensJeu

    MsgBox _
        "Bouton du sens de jeu installe en BP4." & _
        vbCrLf & vbCrLf & _
        "Etat actuel : " & _
        IIf(SensJeuInverse(), _
            "notre camp a droite", _
            "notre camp a gauche") & ".", _
        vbInformation, _
        "Sens de jeu"

End Sub


Private Sub AssurerCelluleEtat()

    Dim Existe As Boolean
    Dim n As Name

    For Each n In ThisWorkbook.Names

        If UCase$(n.Name) = UCase$(NOM_ETAT) Then
            Existe = True
            Exit For
        End If

    Next n

    If Not Existe Then

        ThisWorkbook.Names.Add _
            Name:=NOM_ETAT, _
            RefersTo:="='" & shParametres.Name & "'!" & _
                CELLULE_ETAT

        ThisWorkbook.Names(NOM_ETAT) _
            .RefersToRange.Value = "D"

    End If

End Sub
