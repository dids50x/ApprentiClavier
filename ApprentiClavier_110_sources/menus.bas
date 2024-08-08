Attribute VB_Name = "Module_menus"
'*******************  MENU_RESET : RECR�E le fichier TEXTE des MENUS  ********************
Public Sub menu_reset(ficmnu)
nfree = FreeFile
Module_routines.clean

' menu_principal
If ficmnu = "menu_principal.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "Pr�sentation G�n�rale."
    Print #nfree, "Pour qui ? Pourquoi ?"
    Print #nfree, "Pour la frappe, des conseils !"
    Print #nfree, "1.  Les touches essentielles."
    Print #nfree, "2.  La frappe des lettres."
    Print #nfree, "3.  La fin de l'alphabet."
    Print #nfree, "4.  Approfondir la frappe de l'alphabet."
    Print #nfree, "5.  Mots, proverbes, phrases."
    Print #nfree, "6.  R�gularit�, l'alphabet au hasard."
    Print #nfree, "7.  Essai de vitesse."
    Print #nfree, "8.  Majuscules, accents, ponctuations."
    Print #nfree, "9.  Ponctuations de la rang�e du haut."
    Print #nfree, "10. Les chiffres au clavier principal."
    Print #nfree, "11. La vitesse, le trot et le galop."
    Print #nfree, "12. Insertion, suppression, d�placement."
    Print #nfree, "13. Alt, AltGr, Windows, agir, r�diger."
    Print #nfree, "14. Des dict�es."
    Print #nfree, "15. Encore des dict�es, plus rapides."
    Print #nfree, "16. Les fonctions du pav� num�rique."
    Print #nfree, "17. Toutes les touches, avec dext�rit�."
    Print #nfree, "18. Jouer avec les mots."
    Print #nfree, "19. Des dict�es au galop."
    Print #nfree, "Consulter le fichier des r�sultats."
    Print #nfree, "Quitter."
    Close #nfree
End If

' menu_le�on1
If ficmnu = "menu_le�on1.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les touches ESPACE, Entr�e, �chap."
    Print #nfree, "B. Les Fl�ches et F1, F2, F3."
    Print #nfree, "C. Les touches ALT et Control."
    Print #nfree, "D. Exercice des touches essentielles."
    Close #nfree
End If

' menu_le�on2
If ficmnu = "menu_le�on2.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. La rang�e de d�part Q S D F."
    Print #nfree, "B. Les lettres G et H."
    Print #nfree, "C. La rang�e A Z E R."
    Print #nfree, "D. Des groupes de mots courts."
    Print #nfree, "E. Les lettres T et Y."
    Print #nfree, "F. Des petits groupes de mots."
    Print #nfree, "G. Des groupes de trois mots."
    Print #nfree, "H. Des phrases courtes."
    Close #nfree
End If

' menu_le�on3
If ficmnu = "menu_le�on3.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. La rang�e W X C V."
    Print #nfree, "B. Inclure le B."
    Print #nfree, "C. Des mots ambigus."
    Close #nfree
End If

' menu_le�on4
If ficmnu = "menu_le�on4.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Q D avec S F."
    Print #nfree, "B. La rang�e J K L."
    Print #nfree, "C. La rang�e A Z E R."
    Print #nfree, "D. Le G avec le T."
    Print #nfree, "E. La rang�e U I O P."
    Print #nfree, "F. Le H avec le Y."
    Print #nfree, "G. Utiliser W X C V."
    Print #nfree, "H. Utiliser le B et le N."
    Close #nfree
End If

' menu_le�on5
If ficmnu = "menu_le�on5.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Des mots avec tout l'alphabet."
    Print #nfree, "B. Des proverbes."
    Print #nfree, "C. Des phrases avec tout l'alphabet."
    Close #nfree
End If

' menu_le�on6
If ficmnu = "menu_le�on6.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Des petits mots au hasard."
    Print #nfree, "B. Des petits mots au hasard, vite."
    Print #nfree, "C. Des mots longs au hasard."
    Print #nfree, "D. Des mots se terminant par ""ation""."
    Close #nfree
End If

' menu_le�on7
If ficmnu = "menu_le�on7.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premier essai de vitesse."
    Print #nfree, "B. Deuxi�me essai de vitesse."
    Print #nfree, "C. Troisi�me essai."
    Close #nfree
End If

' menu_le�on8
If ficmnu = "menu_le�on8.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les touches pour majuscules et minuscules."
    Print #nfree, "B. Ponctuations accessibles en minuscules."
    Print #nfree, "C. Ponctuations en majuscules et la Barre-Oblique."
    Print #nfree, "D. Quelques phrases ponctu�es."
    Print #nfree, "E. Le U grave, le Circonflexe et le Tr�ma."
    Print #nfree, "F. Les signes Ast�risque, Inf�rieur �, Sup�rieur �."
    Print #nfree, "G. Les signes PourCent, Mu, Dollar, Livre."
    Print #nfree, "H. Les lettres accentu�es E aigu, E grave, A grave."
    Close #nfree
End If

' menu_le�on9
If ficmnu = "menu_le�on9.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les signes AuCarr� et �commercial."
    Print #nfree, "B. Guillemet, apostrophe, parenth�se, tiret."
    Print #nfree, "C. Soulign�, parenth�se droite."
    Print #nfree, "D. Le �, les signes Degr�, �gal, Plus."
    Close #nfree
End If

' menu_le�on10
If ficmnu = "menu_le�on10.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les quatre premiers chiffres."
    Print #nfree, "B. De 5 � 7"
    Print #nfree, "C. Les chiffres 8, 9 et 0."
    Close #nfree
End If

' menu_le�on11
If ficmnu = "menu_le�on11.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premier exercice de vitesse."
    Print #nfree, "B. Encore de la vitesse."
    Print #nfree, "C. Toujours de la vitesse."
    Close #nfree
End If

' menu_le�on12
If ficmnu = "menu_le�on12.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Insertion, remplacement, suppression."
    Print #nfree, "B. Les touches D�but et Fin."
    Print #nfree, "C. Touches PagePr�c�dente, PageSuivante."
    Print #nfree, "D. Touches Impression, Arr�tD�fil, Pause."
    Close #nfree
End If

' menu_le�on13
If ficmnu = "menu_le�on13.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Control, Windows, Alt, Menu-Contextuel."
    Print #nfree, "B. Les touches Tab et Retour-Arri�re."
    Print #nfree, "C. Les touches de fonction F1 � F12."
    Print #nfree, "D. Les raccourcis clavier."
    Print #nfree, "E. Di�se, Contre-Oblique, Acommercial."
    Print #nfree, "F. Crochets, accolades."
    Print #nfree, "G. Saisir un nom de chemin, ou une commande."
    Close #nfree
End If

' menu_le�on14
If ficmnu = "menu_le�on14.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premi�re dict�e."
    Print #nfree, "B. Deuxi�me dict�e."
    Print #nfree, "C. Troisi�me dict�e."
    Close #nfree
End If

' menu_le�on15
If ficmnu = "menu_le�on15.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premi�re dict�e."
    Print #nfree, "B. Deuxi�me dict�e."
    Print #nfree, "C. Troisi�me dict�e."
    Close #nfree
End If

' menu_le�on16
If ficmnu = "menu_le�on16.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les chiffres."
    Print #nfree, "B. Plus, Moins, Multiplier, Diviser."
    Print #nfree, "C. Les caract�res Ascii et Ansi."
    Print #nfree, "D. Les touches de direction."
    Close #nfree
End If

' menu_le�on17
If ficmnu = "menu_le�on17.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. D�lier les doigts."
    Print #nfree, "B. Caract�res et autres touches."
    Print #nfree, "C. La vitesse sur le clavier."
    Print #nfree, "D. La vitesse sur pav� num�rique."
    Close #nfree
End If

' menu_le�on18
If ficmnu = "menu_le�on18.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Les mots accentu�s."
    Print #nfree, "B. Les consonnes doubles."
    Print #nfree, "C. Des terminaisons usuelles."
    Print #nfree, "D. La vitesse sur mots semblables."
    Print #nfree, "E. La frappe du programmeur."
    Close #nfree
End If

' menu_le�on19
If ficmnu = "menu_le�on19.txt" Then
    Open vpath & "menu_courant.txt" For Output As #nfree
    Print #nfree, "A. Premier texte."
    Print #nfree, "B. Deuxi�me texte."
    Print #nfree, "C. Troisi�me texte."
    Print #nfree, "D. Quatri�me texte."
    Close #nfree
End If

End Sub
