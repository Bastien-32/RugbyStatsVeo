Attribute VB_Name = "modPenalitePopup"
Option Explicit

' =========================================================
' PENALITES - POPUP DE SAISIE
'
' Au clic sur "penalite contre nous" ou "contre
' adversaire", une popup s'ouvre avec les motifs a gauche
' et la composition a droite. On choisit l'un, l'autre,
' les deux ou rien, puis on valide.
'
' Ce qui n'est pas renseigne vaut "?", comme le faisait
' l'attente de la palette. Le groupe fautif n'est pas
' demande : GetPlayerGroup le deduit du joueur.
'
' La disposition reprend la maquette dessinee dans le
' classeur : motifs groupes par famille, joueurs places
' comme sur le terrain, l'attaque vers le bas.
'
' Ce que la case affiche est le nom du joueur, mais c'est
' sa POSITION qui l'identifie : un nom ne permettrait pas
' de remonter au poste, et la compo change d'un match a
' l'autre.
'
' Aucun caractere accentue en clair : le VBE de macOS
' importe les .bas en Mac Roman. Les accents passent donc
' par ChrW, comme ailleurs dans ce classeur.
' =========================================================

Public Const FEUILLE_POPUP_PEN As String = "Popup penalite"

' Colonnes des trois familles de motifs, fusionnees par
' deux : B:C, E:F, H:I.
Private Const COL_MOTIF_1 As Long = 2
Private Const COL_MOTIF_2 As Long = 5
Private Const COL_MOTIF_3 As Long = 8
Private Const LARGEUR_MOTIF As Long = 2

' Les valeurs du contexte se collent a leurs libelles.
Private Const COL_CONTEXTE As Long = 3

' Les boutons joueurs tiennent sur trois colonnes.
Private Const LARGEUR_JOUEUR As Long = 3

Private Const PREMIERE_LIGNE As Long = 7

' Bandeau de la composition : K4:U5, fusionne et mis en
' forme dans la feuille. K4 en est l'ancre.
Private Const LIGNE_TITRE_JOUEURS As Long = 4
Private Const COL_TITRE_JOUEURS As Long = 11
Private Const LIGNE_BOUTONS As Long = 32

Private Const GRIS As Long = 15132390     ' RGB(230,230,230)
Private Const BLEU As Long = 12611584     ' RGB(0,112,192)
Private Const VERT As Long = 32768        ' RGB(0,128,0)
Private Const ROUGE As Long = 192         ' RGB(192,0,0)
Private Const SABLE As Long = 14083324    ' RGB(252,228,214)
Private Const CIEL As Long = 15983321     ' RGB(217,226,243)
Private Const MARINE As Long = 6567967    ' RGB(31,56,100)

' Contexte de la penalite en cours de saisie.
Private PenaliteAction As String
Private PenaliteTemps As Double
Private PenaliteEquipe As String
Private PenaliteOuverte As Boolean

' Vrai quand le fautif est des notres : seule une penalite
' contre nous designe un joueur.
Private PenaliteAvecJoueur As Boolean

' "PEN" ou "CF" : la popup sert aux deux, seuls les
' motifs et le titre changent.
Private PenaliteType As String

' Poste occupant chaque emplacement du terrain au moment
' de la penalite : l'emplacement 12 peut porter le 20 si
' un remplacement a eu lieu.
Private PostesAffiches(1 To 15) As String

' Carton porte par l'occupant de chaque emplacement au
' moment de la faute. Une case ne revient donc pas toujours
' au ciel quand la selection change.
Private CartonsAffiches(1 To 15) As String


' ---------------------------------------------------------
' Disposition
'
' Chaque entree vaut "ligne:colonne:contenu". Lignes et
' colonnes sont celles de la feuille, pour que la maquette
' reste lisible telle qu'elle a ete dessinee.
' ---------------------------------------------------------

' Motifs de la faute en cours. La popup est commune aux
' penalites et aux coups francs : deux tables, une seule
' mecanique.
Private Function MotifsDisposes() As Variant

    If PenaliteType = "CF" Then
        MotifsDisposes = MotifsCoupFranc
    Else
        MotifsDisposes = MotifsPenalite
    End If

End Function


Private Function MotifsCoupFranc() As Variant

    Dim T As String

    ' La liste est assemblee famille par famille : VBA
    ' n'accepte que vingt-quatre continuations de ligne
    ' dans une meme instruction.

    T = "9:2:Melee - Mise en place trop longue" & vbTab & _
        "11:2:Melee - Engagement anticip" & ChrW(233) & vbTab & _
        "13:2:Melee - Pouss" & ChrW(233) & "e anticip" & ChrW(233) & "e" & vbTab & _
        "15:2:Melee - Introduction irreguliere" & vbTab & _
        "17:2:Melee - Talonnage irregulier" & vbTab & _
        "23:2:Jeu-rapide 10m de plus"

    T = T & vbTab & _
        "9:5:Touche - Mise en place trop longue" & vbTab & _
        "11:5:Touche - Alignement irregulier" & vbTab & _
        "13:5:Touche - Lancer non conforme" & vbTab & _
        "15:5:Touche - Saut ou soutien premature" & vbTab & _
        "17:5:Touche - Sauteur mal redescendu" & vbTab & _
        "19:5:Touche - Soutient devant sauteur retomb" & ChrW(233) & "e" & vbTab & _
        "23:5:Autre"

    T = T & vbTab & _
        "9:8:Ruck - Fausse sortie" & vbTab & _
        "11:8:Perte de temps" & vbTab & _
        "15:8:Passage " & ChrW(224) & " vide" & vbTab & _
        "17:8:Ecran"

    MotifsCoupFranc = Split(T, vbTab)

End Function


Private Function MotifsPenalite() As Variant

    Dim T As String

    ' La liste est assemblee famille par famille : VBA
    ' n'accepte que vingt-quatre continuations de ligne
    ' dans une meme instruction.

    T = "9:2:En-avant volontaire" & vbTab & _
        "13:2:Maul - Entrer sur le c" & ChrW(244) & "t" & ChrW(233) & vbTab & _
        "15:2:Maul - Ecroulement" & vbTab & _
        "17:2:Melee - Poussee avant introduction" & vbTab & _
        "21:2:Touche - Saisie bras sauteur" & vbTab & _
        "23:2:Touche - Plaquage sauteur avant retomb" & ChrW(233) & "e" & vbTab & _
        "25:2:Touche - Pouss" & ChrW(233) & "e avant retomb" & ChrW(233) & "e sauteur"

    T = T & vbTab & _
        "9:5:Parle arbitre" & vbTab & _
        "11:5:Ruck - Va au-del" & ChrW(224) & vbTab & _
        "13:5:Ruck - Soutient couch" & ChrW(233) & " sur porteur" & vbTab & _
        "15:5:Ruck - Retard soutient" & vbTab & _
        "17:5:Ruck - Garde le ballon au sol" & vbTab & _
        "19:5:Ruck - D" & ChrW(233) & "blayage sur le cot" & ChrW(233) & vbTab & _
        "21:5:Ruck - Saisie relayeur" & vbTab & _
        "23:5:Ruck - Talonnage " & ChrW(224) & " la main" & vbTab & _
        "25:5:Ruck - Plaqueur qui ne sort pas" & vbTab & _
        "27:5:Ruck - 4 appuis" & vbTab & _
        "29:5:Ruck - Gratteur ne rel" & ChrW(226) & "che pas"

    T = T & vbTab & _
        "9:8:Brutalit" & ChrW(233) & vbTab & _
        "11:8:Plaquage - " & ChrW(224) & " 2" & vbTab & _
        "13:8:Plaquage - haut" & vbTab & _
        "15:8:Plaquage - sans ballon" & vbTab & _
        "17:8:Plaquage - " & ChrW(224) & " retardement" & vbTab & _
        "19:8:Plaquage - en l'air" & vbTab & _
        "23:8:Passage " & ChrW(224) & " vide" & vbTab & _
        "25:8:Hors-jeu" & vbTab & _
        "27:8:10m de plus" & vbTab & _
        "29:8:Autre"

    MotifsPenalite = Split(T, vbTab)

End Function


' L'equipe vue de dessus, l'attaque vers le bas.
Private Function JoueursDisposes() As Variant

    JoueursDisposes = Array( _
        "7:11:1", "7:15:2", "7:19:3", _
        "9:13:4", "9:17:5", _
        "11:11:6", "11:15:8", "11:19:7", _
        "13:13:9", _
        "15:16:10", _
        "17:13:12", "17:17:13", _
        "19:11:11", "19:15:15", "19:19:14", _
        "21:11:16", "21:15:17", "21:19:18", _
        "23:11:19", "23:15:20", "23:19:21", _
        "25:15:22", _
        "27:11:Collectif", "27:19:?")

End Function


' Element d'une entree : 0 la ligne, 1 la colonne, 2 le
' contenu. Le contenu peut lui-meme contenir un deux-
' points, d'ou le recollage.
Private Function Champ( _
    ByVal Entree As String, _
    ByVal Rang As Long) As String

    Dim Morceaux() As String
    Dim i As Long

    Morceaux = Split(Entree, ":")

    If Rang > UBound(Morceaux) Then Exit Function

    If Rang < 2 Then
        Champ = Morceaux(Rang)
        Exit Function
    End If

    Champ = Morceaux(2)

    For i = 3 To UBound(Morceaux)
        Champ = Champ & ":" & Morceaux(i)
    Next i

End Function


' ---------------------------------------------------------
' Construction
' ---------------------------------------------------------

Public Sub ConstruirePopupPenalite()

    Dim ws As Worksheet

    On Error GoTo GestionErreur

    Set ws = FeuillePopup

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    EcrireTitre ws
    EcrireContexte ws
    EcrireBandeauJoueurs ws
    EcrireMotifs ws
    EcrireJoueurs ws
    EcrireBoutons ws
    MettreEnForme ws

    If ModeSilencieux Then Exit Sub

    MsgBox _
        "La popup des penalites est prete.", _
        vbInformation, _
        "Penalites"

    Exit Sub

GestionErreur:

    MsgBox _
        "La construction a echoue." & vbCrLf & vbCrLf & _
        "Erreur " & Err.Number & " : " & Err.Description, _
        vbExclamation, _
        "Penalites"

End Sub


Private Sub EcrireTitre(ByVal ws As Worksheet)

    With ws.Cells(2, COL_MOTIF_1)
        .Value = "PENALITE"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = BLEU
    End With

End Sub


' Bandeau au-dessus de la composition. Son texte est pose
' a l'ouverture, selon qu'une penalite designe ou non un
' de nos joueurs ; sa mise en forme, elle, appartient a la
' construction et doit survivre a une reconstruction.
Private Sub EcrireBandeauJoueurs(ByVal ws As Worksheet)

    With ws.Range( _
        ws.Cells(LIGNE_TITRE_JOUEURS, COL_TITRE_JOUEURS), _
        ws.Cells(LIGNE_TITRE_JOUEURS + 1, 21))

        .Merge
        .Interior.Color = MARINE
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ws.Cells(LIGNE_TITRE_JOUEURS, COL_TITRE_JOUEURS).Value = _
        "JOUEUR FAUTIF"

End Sub


Private Sub EcrireContexte(ByVal ws As Worksheet)

    Dim Libelles As Variant
    Dim i As Long

    Libelles = Array("Temps video", "Mi-temps", "Contre")

    For i = 0 To 2

        With ws.Cells(4 + i, COL_MOTIF_1)
            .Value = Libelles(i)
            .Font.Italic = True
        End With

        With ws.Cells(4 + i, COL_CONTEXTE).Resize(1, 2)
            .Merge
            .Interior.Color = GRIS
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(4, COL_CONTEXTE), _
            ws.Cells(6, COL_CONTEXTE)).Address, _
        "PEN_CONTEXTE"

    ' Lecture et pause sans quitter la popup : la video
    ' reste pilotable pendant qu'on cherche le motif.
    With ws.Cells(4, COL_MOTIF_2).Resize(1, LARGEUR_MOTIF)

        .Merge
        .Value = "PLAY / PAUSE"
        .Interior.Color = VERT
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    NommerPlage ws, _
        ws.Cells(4, COL_MOTIF_2).Address, _
        "PEN_PLAY_PAUSE"

End Sub


' Efface la zone des motifs puis y ecrit ceux du type en
' cours. Les deux tables n'occupent pas les memes cases :
' un reste de l'autre faute resterait cliquable.
Private Sub RedessinerMotifs(ByVal ws As Worksheet)

    Dim Zone As Range

    Set Zone = ws.Range( _
        ws.Cells(PREMIERE_LIGNE, COL_MOTIF_1), _
        ws.Cells(LIGNE_BOUTONS - 1, _
            COL_MOTIF_3 + LARGEUR_MOTIF - 1))

    Zone.UnMerge
    Zone.ClearContents
    Zone.Borders.LineStyle = xlNone

    ' Blanc, et non "aucun remplissage" : sans cela le
    ' quadrillage de la feuille reapparait sous les
    ' emplacements laisses vides par l'autre liste.
    Zone.Interior.Color = RGB(255, 255, 255)

    EcrireMotifs ws

End Sub


Private Sub EcrireMotifs(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    Entrees = MotifsDisposes

    For i = 0 To UBound(Entrees)

        With ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))) _
            .Resize(1, LARGEUR_MOTIF)

            .Merge
            .Value = Champ(CStr(Entrees(i)), 2)
            .Interior.Color = SABLE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)

        End With

    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(PREMIERE_LIGNE, COL_MOTIF_1), _
            ws.Cells(LIGNE_BOUTONS - 1, _
                COL_MOTIF_3 + LARGEUR_MOTIF - 1)).Address, _
        "PEN_MOTIFS"

End Sub


Private Sub EcrireJoueurs(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        With ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))) _
            .Resize(1, LARGEUR_JOUEUR)

            .Merge
            .Interior.Color = CIEL
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)

        End With

        EcrireCaseJoueur _
            ws.Cells( _
                CLng(Champ(CStr(Entrees(i)), 0)), _
                CLng(Champ(CStr(Entrees(i)), 1))), _
            Champ(CStr(Entrees(i)), 2)

    Next i

    ' Les sept rangs du terrain portent deux lignes de
    ' texte : le numero puis le nom.
    For i = PREMIERE_LIGNE To 19 Step 2
        ws.Rows(i).RowHeight = 32
    Next i

    NommerPlage ws, _
        ws.Range( _
            ws.Cells(PREMIERE_LIGNE, 11), _
            ws.Cells(LIGNE_BOUTONS - 1, 21)).Address, _
        "PEN_JOUEURS"

End Sub


Private Sub EcrireBoutons(ByVal ws As Worksheet)

    With ws.Cells(LIGNE_BOUTONS, COL_MOTIF_1).Resize(1, 2)
        .Merge
        .Value = "ANNULER"
        .Interior.Color = ROUGE
    End With

    With ws.Cells(LIGNE_BOUTONS, COL_MOTIF_2).Resize(1, 4)
        .Merge
        .Value = "VALIDER  (Entree)"
        .Interior.Color = VERT
    End With

    With ws.Range( _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_1), _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_3 + 1))

        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ws.Rows(LIGNE_BOUTONS).RowHeight = 26

    NommerPlage ws, _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_1).Address, _
        "PEN_ANNULER"

    NommerPlage ws, _
        ws.Cells(LIGNE_BOUTONS, COL_MOTIF_2).Address, _
        "PEN_VALIDER"

End Sub


Private Sub MettreEnForme(ByVal ws As Worksheet)

    ' Largeurs de la maquette : une colonne etroite separe
    ' les familles de motifs, et la composition tient sur
    ' des colonnes fines.
    ws.Columns("B:C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 3.5
    ws.Columns("E:F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 4
    ws.Columns("H:I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 4
    ws.Columns("K:U").ColumnWidth = 4.3

    ws.Cells.Font.Name = "Calibri"

    AjusterHauteurs ws

End Sub


' Les rangs qui portent quelque chose font 20, les rangs
' intercalaires 10 : la grille respire sans s'etaler.
Private Sub AjusterHauteurs(ByVal ws As Worksheet)

    Dim Ligne As Long

    For Ligne = PREMIERE_LIGNE To LIGNE_BOUTONS - 2

        If Application.WorksheetFunction.CountA( _
            ws.Range( _
                ws.Cells(Ligne, 2), _
                ws.Cells(Ligne, 21))) > 0 Then

            ws.Rows(Ligne).RowHeight = 20

        Else

            ws.Rows(Ligne).RowHeight = 10

        End If

    Next Ligne

End Sub


' ---------------------------------------------------------
' Libelles
' ---------------------------------------------------------

' Nom de famille du joueur, ou le numero si la compo ne le
' donne pas. Meme regle que la palette de saisie : les
' mots en majuscules forment le nom de famille.
Private Function LibelleJoueur( _
    ByVal Poste As String) As String

    Dim Complet As String
    Dim Mots() As String
    Dim Mot As Variant
    Dim Famille As String

    LibelleJoueur = Poste

    If Not IsNumeric(Poste) Then Exit Function

    Complet = Trim(GetPlayerName(Poste))

    If Complet = "" Then Exit Function

    Mots = Split(Complet, " ")

    For Each Mot In Mots

        If CStr(Mot) = UCase(CStr(Mot)) Then

            If Famille = "" Then
                Famille = CStr(Mot)
            Else
                Famille = Famille & " " & CStr(Mot)
            End If

        Else

            Exit For

        End If

    Next Mot

    If Famille <> "" Then LibelleJoueur = Famille

End Function


' Le numero en gras, le nom dessous : meme presentation
' que les boutons de la palette de saisie.
Private Sub EcrireCaseJoueur( _
    ByVal Cellule As Range, _
    ByVal Poste As String, _
    Optional ByVal Carton As String = "")

    Dim Nom As String

    Nom = LibelleJoueur(Poste)

    With Cellule

        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        ' Une case eteinte a perdu son fond et ses
        ' bordures : il faut les lui rendre. Les bordures
        ' d'abord : le carton blanc pose les siennes en noir
        ' et doit avoir le dernier mot.
        If .MergeCells Then
            .MergeArea.Borders.LineStyle = xlContinuous
            .MergeArea.Borders.Color = RGB(150, 150, 150)
        Else
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End If

        If Carton = "" Then

            .Interior.Color = CIEL
            .Font.Color = RGB(0, 0, 0)

        Else

            AppliquerCouleursCarton ZoneCase(Cellule), Carton

        End If

        If Nom = Poste Then

            ' Sans compo, ou pour Collectif et "?", il n'y
            ' a que l'etiquette.
            .Value = Poste
            .Font.Size = 14
            .Font.Bold = True

            Exit Sub

        End If

        .Value = Poste & vbLf & Nom

        .Characters(1, Len(Poste)).Font.Size = 14
        .Characters(1, Len(Poste)).Font.Bold = True

        .Characters(Len(Poste) + 2, Len(Nom)).Font.Size = 9
        .Characters(Len(Poste) + 2, Len(Nom)).Font.Bold = False

    End With

End Sub


' Seuls les joueurs presents a cet instant de la video
' sont montres : les quinze emplacements du terrain
' portent leur occupant du moment, les cases des
' remplacants sont eteintes.
'
' Les lignes ne peuvent pas etre masquees : les motifs
' occupent les memes. Les cases inutiles sont donc videes
' et rendues invisibles.
Private Sub RafraichirLibelles( _
    ByVal ws As Worksheet, _
    ByVal TempsVideo As Double, _
    ByVal AvecJoueur As Boolean)

    Dim Entrees As Variant
    Dim i As Long
    Dim Etiquette As String
    Dim Emplacement As Long
    Dim Cellule As Range
    Dim Cartons As Collection

    ' Les evenements sont rendus meme en cas d'erreur :
    ' coupes, plus aucun clic ne passerait ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    ' Les cartons en cours au moment de la faute, releves
    ' une seule fois pour toute la composition.
    Set Cartons = CartonsEnCours(TempsVideo)

    For i = 1 To 15
        CartonsAffiches(i) = ""
    Next i

    ' Deux lignes de texte par case : sans cette hauteur,
    ' le renvoi a la ligne masque le nom et seul le numero
    ' reste visible.
    For i = PREMIERE_LIGNE To 19 Step 2
        ws.Rows(i).RowHeight = 32
    Next i

    ' Une penalite contre l'adversaire ne designe pas un
    ' de nos joueurs : la composition disparait alors.
    With ws.Cells(LIGNE_TITRE_JOUEURS, COL_TITRE_JOUEURS)

        .Value = IIf(AvecJoueur, "JOUEUR FAUTIF", "")

        ' Le fond marine s'en va aussi : vide, le bandeau
        ' resterait visible comme un titre sans texte.
        If AvecJoueur Then
            .MergeArea.Interior.Color = MARINE
        Else
            .MergeArea.Interior.Color = RGB(255, 255, 255)
        End If

    End With

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        Etiquette = Champ(CStr(Entrees(i)), 2)

        Set Cellule = ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1)))

        If Not AvecJoueur Then

            EteindreCase Cellule

        ElseIf Not IsNumeric(Etiquette) Then

            ' Collectif et l'inconnu restent toujours la.
            EcrireCaseJoueur Cellule, Etiquette

        Else

            Emplacement = CLng(Etiquette)

            If Emplacement <= 15 Then

                PostesAffiches(Emplacement) = _
                    OccupantEmplacement(Emplacement, TempsVideo)

                CartonsAffiches(Emplacement) = _
                    CartonDeJoueurDansListe( _
                        Cartons, _
                        GetPlayerName( _
                            PostesAffiches(Emplacement)))

                EcrireCaseJoueur Cellule, _
                    PostesAffiches(Emplacement), _
                    CartonsAffiches(Emplacement)

            Else

                EteindreCase Cellule

            End If

        End If

    Next i

Sortie:

    Application.EnableEvents = True

    If Err.Number <> 0 Then

        MsgBox _
            "La composition n'a pas pu etre mise a jour." & _
            vbCrLf & vbCrLf & _
            "Erreur " & Err.Number & " : " & _
            Err.Description & vbCrLf & vbCrLf & _
            "Emplacement en cours : " & Etiquette, _
            vbExclamation, _
            "Penalites"

    End If

End Sub


Private Sub EteindreCase(ByVal Cellule As Range)

    Dim Zone As Range

    ' Les bordures se posent sur la zone fusionnee entiere,
    ' jamais sur son ancre seule : Excel refuse la seconde.
    If Cellule.MergeCells Then
        Set Zone = Cellule.MergeArea
    Else
        Set Zone = Cellule
    End If

    With Zone
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone
    End With

End Sub


' ---------------------------------------------------------
' Ouverture et fermeture
' ---------------------------------------------------------

Public Sub OuvrirPopupPenalite( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal Equipe As String)

    OuvrirPopupFaute ActionTexte, TempsVideo, Equipe, "PEN"

End Sub


Public Sub OuvrirPopupCoupFranc( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal Equipe As String)

    OuvrirPopupFaute ActionTexte, TempsVideo, Equipe, "CF"

End Sub


Private Sub OuvrirPopupFaute( _
    ByVal ActionTexte As String, _
    ByVal TempsVideo As Double, _
    ByVal Equipe As String, _
    ByVal TypeFaute As String)

    Dim ws As Worksheet

    PenaliteType = TypeFaute
    PenaliteAction = ActionTexte
    PenaliteTemps = TempsVideo
    PenaliteEquipe = Equipe
    PenaliteOuverte = True
    PenaliteAvecJoueur = (Equipe = "Adv")

    ' Une action restee en attente doit etre close avant
    ' que la penalite prenne la main.
    FermerActionEnAttente
    SuspendreVideoPourPopup TempsVideo

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    ws.Visible = xlSheetVisible
    ws.Activate

    ws.Cells(2, COL_MOTIF_1).Value = _
        IIf(TypeFaute = "CF", "COUP FRANC", "PENALITE")

    RedessinerMotifs ws

    EffacerSelections ws
    RafraichirLibelles ws, TempsVideo, PenaliteAvecJoueur

    With ws.Range("PEN_CONTEXTE")
        ' Format texte impose : sans lui, Excel lit
        ' "64:48" comme 64 heures 48 et affiche 16:48.
        .Cells(1, 1).NumberFormat = "@"
        .Cells(1, 1).Value = FormaterTemps(TempsVideo)
        .Cells(2, 1).Value = CurrentHalf
        ' Equipe est la possession qui suit la penalite :
        ' si le ballon nous revient, c'est l'adversaire
        ' qui a commis la faute.
        .Cells(3, 1).Value = _
            IIf(Equipe = "Nous", "Adversaire", "Nous")
    End With

    ' La touche Entree vaut validation, comme le fait deja
    ' la popup d'ajout de joueur pour son bouton.
    ActiverEspacePopup

    Application.OnKey "~", "ValiderPopupPenalite"
    Application.OnKey "{ENTER}", "ValiderPopupPenalite"

End Sub


Public Sub ValiderPopupPenalite()

    Dim Motif As String
    Dim Joueur As String

    If Not PenaliteOuverte Then Exit Sub

    Motif = SelectionCourante("PEN_MOTIFS")
    Joueur = PosteSelectionne

    If Motif = "" Then Motif = "?"

    If Not PenaliteAvecJoueur Then

        ' Penalite contre l'adversaire : le fautif est
        ' chez eux, les deux colonnes restent vides.
        Joueur = ""

    ElseIf Joueur = "" Then

        Joueur = "?"

    Else

        ' GetPlayerName accepte un numero, "Collectif" ou
        ' "?" et rend le nom porte par la compo.
        Joueur = GetPlayerName(Joueur)

        If Trim(Joueur) = "" Then Joueur = "?"

    End If

    SetCurrentTeam PenaliteEquipe

    AjouterAction _
        Joueur, _
        PenaliteAction, _
        PenaliteTemps, _
        Motif, _
        IIf(Joueur = "", "", GetPlayerGroup(Joueur))

    ArreterLeJeu

    FermerPopupPenalite

End Sub


Public Sub AnnulerPopupPenalite()

    If Not PenaliteOuverte Then Exit Sub

    ' Comme la palette avant elle : la penalite est
    ' enregistree, motif et joueur inconnus.
    SetCurrentTeam PenaliteEquipe

    AjouterAction _
        IIf(PenaliteAvecJoueur, "?", ""), _
        PenaliteAction, _
        PenaliteTemps, _
        "?", _
        IIf(PenaliteAvecJoueur, "?", "")

    ArreterLeJeu

    FermerPopupPenalite

End Sub


' Une penalite arrete le jeu : la palette ajoutait cette
' ligne d'elle-meme, la popup doit continuer de le faire.
Private Sub ArreterLeJeu()

    If ActionDeclencheArretAutomatique(PenaliteAction) Then
        AjouterArretJeuAutomatique PenaliteTemps
    End If

End Sub


Private Sub FermerPopupPenalite()

    Dim ws As Worksheet

    PenaliteOuverte = False

    ReprendreVideoApresPopup

    Application.OnKey "~"
    Application.OnKey "{ENTER}"

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    shSaisieVideo.Activate
    ws.Visible = xlSheetHidden

End Sub


' ---------------------------------------------------------
' Selection
'
' Un clic met la case en evidence et eteint les autres :
' un seul motif et un seul joueur.
' ---------------------------------------------------------

Public Sub ClicPopupPenalite(ByVal Target As Range)

    Dim ws As Worksheet
    Dim Carton As String

    If Target.Areas.Count > 1 Then Exit Sub

    ' Un clic sur une cellule fusionnee livre toute la
    ' plage : on ne garde que son ancre.
    If Target.Cells.Count > 1 Then

        If Not Target.MergeCells Then Exit Sub
        Set Target = Target.Cells(1, 1)

    End If

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    If Not Intersect(Target, ws.Range("PEN_PLAY_PAUSE")) _
        Is Nothing Then

        BasculerLectureDepuisPopup
        Exit Sub

    End If

    If Not Intersect(Target, ws.Range("PEN_VALIDER")) _
        Is Nothing Then

        ValiderPopupPenalite
        Exit Sub

    End If

    If Not Intersect(Target, ws.Range("PEN_ANNULER")) _
        Is Nothing Then

        AnnulerPopupPenalite
        Exit Sub

    End If

    ' Un joueur sous carton n'etait pas sur le terrain a cet
    ' instant : il ne peut pas etre le fautif.
    Carton = CartonDeLaCase(ws, Target)

    If Carton <> "" Then

        MsgBox _
            "Ce joueur portait un carton " & _
            CouleurLisible(Carton) & " " & ChrW(224) & _
            " cet instant : il n'" & ChrW(233) & "tait pas " & _
            "sur le terrain et ne peut pas " & ChrW(234) & _
            "tre le fautif.", _
            vbExclamation, _
            "Joueur sous carton"

        Exit Sub

    End If

    BasculerSelection ws, Target, "PEN_MOTIFS", SABLE
    BasculerSelection ws, Target, "PEN_JOUEURS", CIEL

End Sub


' Rend aux cases de la composition leur fond d'origine : le
' ciel habituel, ou la couleur du carton pour qui en porte
' un. La selection, elle, eteint tout sans distinction.
Private Sub RendreCasesJoueurs(ByVal ws As Worksheet)

    Dim Entrees As Variant
    Dim i As Long
    Dim Etiquette As String
    Dim Emplacement As Long
    Dim Cellule As Range

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        Etiquette = Champ(CStr(Entrees(i)), 2)

        Set Cellule = ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1)))

        If Trim(CStr(Cellule.Value)) <> "" Then

            If Not IsNumeric(Etiquette) Then

                EcrireCaseJoueur Cellule, Etiquette

            Else

                Emplacement = CLng(Etiquette)

                If Emplacement <= 15 Then

                    EcrireCaseJoueur Cellule, _
                        PostesAffiches(Emplacement), _
                        CartonsAffiches(Emplacement)

                End If

            End If

        End If

    Next i

End Sub


' Le carton porte par la case cliquee, vide si elle n'en
' montre pas ou si ce n'est pas une case de joueur.
Private Function CartonDeLaCase( _
    ByVal ws As Worksheet, _
    ByVal Target As Range) As String

    Dim Entrees As Variant
    Dim i As Long
    Dim Etiquette As String
    Dim Emplacement As Long

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        If ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))).Address = _
            Target.Address Then

            Etiquette = Champ(CStr(Entrees(i)), 2)

            If Not IsNumeric(Etiquette) Then Exit Function

            Emplacement = CLng(Etiquette)

            If Emplacement > 15 Then Exit Function

            CartonDeLaCase = CartonsAffiches(Emplacement)
            Exit Function

        End If

    Next i

End Function


' Le fond et les bordures se posent sur la zone fusionnee
' entiere, jamais sur son ancre seule.
Private Function ZoneCase(ByVal Cellule As Range) As Range

    If Cellule.MergeCells Then
        Set ZoneCase = Cellule.MergeArea
    Else
        Set ZoneCase = Cellule
    End If

End Function


Private Sub BasculerSelection( _
    ByVal ws As Worksheet, _
    ByVal Target As Range, _
    ByVal NomGrille As String, _
    ByVal CouleurNormale As Long)

    Dim Grille As Range
    Dim Cellule As Range
    Dim DejaChoisie As Boolean

    Set Grille = ws.Range(NomGrille)

    If Intersect(Target, Grille) Is Nothing Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub

    DejaChoisie = (Target.Interior.Color = BLEU)

    ' Les evenements sont rendus meme en cas d'erreur :
    ' coupes, plus aucun clic ne passerait ensuite.
    On Error GoTo Sortie

    Application.EnableEvents = False

    If NomGrille = "PEN_JOUEURS" Then

        ' Les cases ne reviennent pas toutes au ciel : celle
        ' d'un joueur sous carton garde sa couleur.
        RendreCasesJoueurs ws

    Else

        For Each Cellule In Grille

            If Trim(CStr(Cellule.Value)) <> "" Then
                Cellule.Interior.Color = CouleurNormale
                Cellule.Font.Color = RGB(0, 0, 0)
            End If

        Next Cellule

    End If

    If Not DejaChoisie Then
        Target.Interior.Color = BLEU
        Target.Font.Color = RGB(255, 255, 255)
    End If

Sortie:

    Application.EnableEvents = True

    If Err.Number <> 0 Then

        MsgBox _
            "La selection n'a pas pu etre appliquee." & _
            vbCrLf & vbCrLf & _
            "Erreur " & Err.Number & " : " & _
            Err.Description, _
            vbExclamation, _
            "Penalites"

    End If

End Sub


' Poste choisi dans la composition, rendu par sa position
' et non par ce que la case affiche.
Private Function PosteSelectionne() As String

    Dim ws As Worksheet
    Dim Entrees As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)

    Entrees = JoueursDisposes

    For i = 0 To UBound(Entrees)

        If ws.Cells( _
            CLng(Champ(CStr(Entrees(i)), 0)), _
            CLng(Champ(CStr(Entrees(i)), 1))) _
            .Interior.Color = BLEU Then

            PosteSelectionne = Champ(CStr(Entrees(i)), 2)

            ' Un emplacement rend son occupant du moment,
            ' qui n'est pas forcement le titulaire.
            If IsNumeric(PosteSelectionne) Then

                If CLng(PosteSelectionne) <= 15 Then
                    PosteSelectionne = _
                        PostesAffiches(CLng(PosteSelectionne))
                End If

            End If

            Exit Function

        End If

    Next i

End Function


Private Function SelectionCourante( _
    ByVal NomGrille As String) As String

    Dim Cellule As Range

    For Each Cellule In ThisWorkbook _
        .Sheets(FEUILLE_POPUP_PEN).Range(NomGrille)

        If Cellule.Interior.Color = BLEU Then
            SelectionCourante = Trim(CStr(Cellule.Value))
            Exit Function
        End If

    Next Cellule

End Function


Private Sub EffacerSelections(ByVal ws As Worksheet)

    On Error GoTo Sortie

    Application.EnableEvents = False

    RendreGrille ws, "PEN_MOTIFS", SABLE
    RendreGrille ws, "PEN_JOUEURS", CIEL

Sortie:

    Application.EnableEvents = True

End Sub


Private Sub RendreGrille( _
    ByVal ws As Worksheet, _
    ByVal NomGrille As String, _
    ByVal Couleur As Long)

    Dim Cellule As Range

    For Each Cellule In ws.Range(NomGrille)

        If Trim(CStr(Cellule.Value)) <> "" Then
            Cellule.Interior.Color = Couleur
            Cellule.Font.Color = RGB(0, 0, 0)
        End If

    Next Cellule

End Sub


' ---------------------------------------------------------
' Utilitaires
' ---------------------------------------------------------

Private Function FormaterTemps( _
    ByVal Secondes As Double) As String

    Dim Total As Long

    Total = CLng(Int(Secondes))

    FormaterTemps = Format(Total \ 60, "00") & ":" & _
        Format(Total Mod 60, "00")

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


Private Function FeuillePopup() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_POPUP_PEN)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count))

        ws.Name = FEUILLE_POPUP_PEN

    Else

        ws.Visible = xlSheetVisible

    End If

    Set FeuillePopup = ws

End Function
