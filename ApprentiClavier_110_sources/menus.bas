Attribute VB_Name = "Module_menus"
'*******************  MENU_RESET : RECRÉE le fichier TEXTE des MENUS  ********************
Public Sub menu_reset(ficmnu)
nfree = FreeFile
Module_routines.clean

' menu_principal
If ficmnu = "menu_principal.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "Présentation Générale."
    Print #nfree, "Pour qui ? Pourquoi ?"
    Print #nfree, "Pour la frappe, des conseils !"
    Print #nfree, "1.  Les touches essentielles."
    Print #nfree, "2.  La frappe des lettres."
    Print #nfree, "3.  La fin de l'alphabet."
    Print #nfree, "4.  Approfondir la frappe de l'alphabet."
    Print #nfree, "5.  Mots, proverbes, phrases."
    Print #nfree, "6.  Régularité, l'alphabet au hasard."
    Print #nfree, "7.  Essai de vitesse."
    Print #nfree, "8.  Majuscules, accents, ponctuations."
    Print #nfree, "9.  Ponctuations de la rangée du haut."
    Print #nfree, "10. Les chiffres au clavier principal."
    Print #nfree, "11. La vitesse, le trot et le galop."
    Print #nfree, "12. Insertion, suppression, déplacement."
    Print #nfree, "13. Alt, AltGr, Windows, agir, rédiger."
    Print #nfree, "14. Des dictées."
    Print #nfree, "15. Encore des dictées, plus rapides."
    Print #nfree, "16. Les fonctions du pavé numérique."
    Print #nfree, "17. Toutes les touches, avec dextérité."
    Print #nfree, "18. Jouer avec les mots."
    Print #nfree, "19. Des dictées au galop."
    Print #nfree, "Consulter le fichier des résultats."
    Print #nfree, "Quitter."
    Close #nfree
End If

' menu_leçon1
If ficmnu = "menu_leçon1.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les touches ESPACE, Entrée, Échap."
    Print #nfree, "B. Les Flèches et F1, F2, F3."
    Print #nfree, "C. Les touches ALT et Control."
    Print #nfree, "D. Exercice des touches essentielles."
    Close #nfree
End If

' menu_leçon2
If ficmnu = "menu_leçon2.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. La rangée de départ Q S D F."
    Print #nfree, "B. Les lettres G et H."
    Print #nfree, "C. La rangée A Z E R."
    Print #nfree, "D. Des groupes de mots courts."
    Print #nfree, "E. Les lettres T et Y."
    Print #nfree, "F. Des petits groupes de mots."
    Print #nfree, "G. Des groupes de trois mots."
    Print #nfree, "H. Des phrases courtes."
    Close #nfree
End If

' menu_leçon3
If ficmnu = "menu_leçon3.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. La rangée W X C V."
    Print #nfree, "B. Inclure le B."
    Print #nfree, "C. Des mots ambigus."
    Close #nfree
End If

' menu_leçon4
If ficmnu = "menu_leçon4.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Q D avec S F."
    Print #nfree, "B. La rangée J K L."
    Print #nfree, "C. La rangée A Z E R."
    Print #nfree, "D. Le G avec le T."
    Print #nfree, "E. La rangée U I O P."
    Print #nfree, "F. Le H avec le Y."
    Print #nfree, "G. Utiliser W X C V."
    Print #nfree, "H. Utiliser le B et le N."
    Close #nfree
End If

' menu_leçon5
If ficmnu = "menu_leçon5.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Des mots avec tout l'alphabet."
    Print #nfree, "B. Des proverbes."
    Print #nfree, "C. Des phrases avec tout l'alphabet."
    Close #nfree
End If

' menu_leçon6
If ficmnu = "menu_leçon6.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Des petits mots au hasard."
    Print #nfree, "B. Des petits mots au hasard, vite."
    Print #nfree, "C. Des mots longs au hasard."
    Print #nfree, "D. Des mots se terminant par ""ation""."
    Close #nfree
End If

' menu_leçon7
If ficmnu = "menu_leçon7.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premier essai de vitesse."
    Print #nfree, "B. Deuxième essai de vitesse."
    Print #nfree, "C. Troisième essai."
    Close #nfree
End If

' menu_leçon8
If ficmnu = "menu_leçon8.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les touches pour majuscules et minuscules."
    Print #nfree, "B. Ponctuations accessibles en minuscules."
    Print #nfree, "C. Ponctuations en majuscules et la Barre-Oblique."
    Print #nfree, "D. Quelques phrases ponctuées."
    Print #nfree, "E. Le U grave, le Circonflexe et le Tréma."
    Print #nfree, "F. Les signes Astérisque, Inférieur à, Supérieur à."
    Print #nfree, "G. Les signes PourCent, Mu, Dollar, Livre."
    Print #nfree, "H. Les lettres accentuées E aigu, E grave, A grave."
    Close #nfree
End If

' menu_leçon9
If ficmnu = "menu_leçon9.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les signes AuCarré et Écommercial."
    Print #nfree, "B. Guillemet, apostrophe, parenthèse, tiret."
    Print #nfree, "C. Souligné, parenthèse droite."
    Print #nfree, "D. Le ç, les signes Degré, Égal, Plus."
    Close #nfree
End If

' menu_leçon10
If ficmnu = "menu_leçon10.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les quatre premiers chiffres."
    Print #nfree, "B. De 5 à 7"
    Print #nfree, "C. Les chiffres 8, 9 et 0."
    Close #nfree
End If

' menu_leçon11
If ficmnu = "menu_leçon11.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premier exercice de vitesse."
    Print #nfree, "B. Encore de la vitesse."
    Print #nfree, "C. Toujours de la vitesse."
    Close #nfree
End If

' menu_leçon12
If ficmnu = "menu_leçon12.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Insertion, remplacement, suppression."
    Print #nfree, "B. Les touches Début et Fin."
    Print #nfree, "C. Touches PagePrécédente, PageSuivante."
    Print #nfree, "D. Touches Impression, ArrêtDéfil, Pause."
    Close #nfree
End If

' menu_leçon13
If ficmnu = "menu_leçon13.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Control, Windows, Alt, Menu-Contextuel."
    Print #nfree, "B. Les touches Tab et Retour-Arrière."
    Print #nfree, "C. Les touches de fonction F1 à F12."
    Print #nfree, "D. Les raccourcis clavier."
    Print #nfree, "E. Dièse, Contre-Oblique, Acommercial."
    Print #nfree, "F. Crochets, accolades."
    Print #nfree, "G. Saisir un nom de chemin, ou une commande."
    Close #nfree
End If

' menu_leçon14
If ficmnu = "menu_leçon14.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Première dictée."
    Print #nfree, "B. Deuxième dictée."
    Print #nfree, "C. Troisième dictée."
    Close #nfree
End If

' menu_leçon15
If ficmnu = "menu_leçon15.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Première dictée."
    Print #nfree, "B. Deuxième dictée."
    Print #nfree, "C. Troisième dictée."
    Close #nfree
End If

' menu_leçon16
If ficmnu = "menu_leçon16.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les chiffres."
    Print #nfree, "B. Plus, Moins, Multiplier, Diviser."
    Print #nfree, "C. Les caractères Ascii et Ansi."
    Print #nfree, "D. Les touches de direction."
    Close #nfree
End If

' menu_leçon17
If ficmnu = "menu_leçon17.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Délier les doigts."
    Print #nfree, "B. Caractères et autres touches."
    Print #nfree, "C. La vitesse sur le clavier."
    Print #nfree, "D. La vitesse sur pavé numérique."
    Close #nfree
End If

' menu_leçon18
If ficmnu = "menu_leçon18.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les mots accentués."
    Print #nfree, "B. Les consonnes doubles."
    Print #nfree, "C. Des terminaisons usuelles."
    Print #nfree, "D. La vitesse sur mots semblables."
    Print #nfree, "E. La frappe du programmeur."
    Close #nfree
End If

' menu_leçon19
If ficmnu = "menu_leçon19.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premier texte."
    Print #nfree, "B. Deuxième texte."
    Print #nfree, "C. Troisième texte."
    Print #nfree, "D. Quatrième texte."
    Close #nfree
End If

End Sub
