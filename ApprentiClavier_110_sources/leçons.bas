Attribute VB_Name = "Module_leçons"
'******************  RESET_STANDARD : RESET des leçons STANDARD  **************************
Public Sub reset_standard(force As Byte)
' Création des reps Standard et Personnalisé s'il n'existent pas encore
On Error Resume Next
MkDir vpath & "Leçons"
On Error Resume Next
MkDir vpath & "Leçons\Standard"
On Error Resume Next
MkDir vpath & "Leçons\Personnalisé"
nfree = FreeFile

' Copier info.txt sur la façon de personnaliser les leçons
If Dir(vpath & "info.txt") <> "" Then FileCopy vpath & "info.txt", vpath & "Leçons\Personnalisé\info.txt"

' leçon1A
If Dir(vpath & "Leçons\Standard\leçon1A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon1A.txt" For Output As #nfree
    Print #nfree, vvEspace
    Print #nfree, vvEntrée
    Print #nfree, vvÉchap
    Print #nfree, vvEspace
    Print #nfree, vvEntrée
    Print #nfree, vvÉchap
    Print #nfree, vvEntrée
    Print #nfree, vvÉchap
    Print #nfree, vvEspace
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon1A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon1A.txt", vpath & "Leçons\Personnalisé\leçon1A.txt"
End If

' leçon1B
If Dir(vpath & "Leçons\Standard\leçon1B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon1B.txt" For Output As #nfree
    Print #nfree, vvFlecheGauche
    Print #nfree, vvFlecheDroite
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheGauche
    Print #nfree, vvFlecheDroite
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFlecheBas
    Print #nfree, "F1"
    Print #nfree, vvÉchap
    Print #nfree, "F1"
    Print #nfree, "F2"
    Print #nfree, vvÉchap
    Print #nfree, "F2"
    Print #nfree, "F3"
    Print #nfree, vvÉchap
    Print #nfree, "F3"
    Print #nfree, vvFlecheBas
    Print #nfree, "F1"
    Print #nfree, vvFlecheGauche
    Print #nfree, "F2"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon1B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon1B.txt", vpath & "Leçons\Personnalisé\leçon1B.txt"
End If

' leçon1C
If Dir(vpath & "Leçons\Standard\leçon1C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon1C.txt" For Output As #nfree
    Print #nfree, vvAlt
    Print #nfree, vvEspace
    Print #nfree, vvAlt
    Print #nfree, vvEspace
    Print #nfree, vvControlGauche
    Print #nfree, vvEspace  ' Nécessaire après CONTROL, sinon AltGr = Crtl+Alt serait accepté
    Print #nfree, vvAlt
    Print #nfree, vvControlDroit
    Print #nfree, vvEspace  ' Nécessaire après CONTROL, sinon AltGr = Crtl+Alt serait accepté
    Print #nfree, vvControlGauche
    Print #nfree, vvAlt
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon1C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon1C.txt", vpath & "Leçons\Personnalisé\leçon1C.txt"
End If

' leçon1D
If Dir(vpath & "Leçons\Standard\leçon1D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon1D.txt" For Output As #nfree
    Print #nfree, vvFlecheHaut
    Print #nfree, vvÉchap
    Print #nfree, vvControl
    Print #nfree, vvEspace
    Print #nfree, vvAlt
    Print #nfree, vvFlecheBas
    Print #nfree, vvEntrée
    Print #nfree, vvFlecheGauche
    Print #nfree, vvEspace
    Print #nfree, vvAlt
    Print #nfree, "F1"
    Print #nfree, vvFlecheHaut
    Print #nfree, vvÉchap
    Print #nfree, "F2"
    Print #nfree, vvFlecheDroite
    Print #nfree, vvControlDroit
    Print #nfree, "F3"
    Print #nfree, vvFlecheBas
    Print #nfree, vvEntrée
    Print #nfree, "F1"
    Print #nfree, vvFlecheGauche
    Print #nfree, "F2"
    Print #nfree, vvAlt
    Print #nfree, vvControlGauche
    Print #nfree, vvEspace
    Print #nfree, "F3"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon1D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon1D.txt", vpath & "Leçons\Personnalisé\leçon1D.txt"
End If

' leçon2A
If Dir(vpath & "Leçons\Standard\leçon2A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2A.txt" For Output As #nfree
    Print #nfree, "q"
    Print #nfree, "s"  'attention au type de police sinon certaines lettres ne sont pas prononcées
    Print #nfree, "d"  'mais ce pb est résolu en faisant suivre le caractère par un blanc dur Alt255
    Print #nfree, "f"
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "j"
    Print #nfree, "k"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "l"
    Print #nfree, "k"
    Print #nfree, "m"
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "j"
    Print #nfree, "l"
    Print #nfree, "f"
    Print #nfree, "s"
    Print #nfree, "m"
    Print #nfree, "k"
    Print #nfree, "f"
    Print #nfree, "s"
    Print #nfree, "d"
    Print #nfree, "q"
    Print #nfree, "k"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "l"
    Print #nfree, "q"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2A.txt", vpath & "Leçons\Personnalisé\leçon2A.txt"
End If

' leçon2B
If Dir(vpath & "Leçons\Standard\leçon2B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2B.txt" For Output As #nfree
    Print #nfree, "g"
    Print #nfree, "h"
    Print #nfree, "h"
    Print #nfree, "g"
    Print #nfree, "g"
    Print #nfree, "h"
    Print #nfree, "h"
    Print #nfree, "g"
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "f"
    Print #nfree, "g"
    Print #nfree, "f"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "g"
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "f"
    Print #nfree, "g"
    Print #nfree, "f"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "g"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "m"
    Print #nfree, "h"
    Print #nfree, "l"
    Print #nfree, "h"
    Print #nfree, "k"
    Print #nfree, "h"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "m"
    Print #nfree, "h"
    Print #nfree, "l"
    Print #nfree, "h"
    Print #nfree, "k"
    Print #nfree, "h"
    Print #nfree, "s"
    Print #nfree, "d"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "f"
    Print #nfree, "h"
    Print #nfree, "s"
    Print #nfree, "l"
    Print #nfree, "g"
    Print #nfree, "q"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "k"
    Print #nfree, "s"
    Print #nfree, "m"
    Print #nfree, "q"
    Print #nfree, "g"
    Print #nfree, "s"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "s"
    Print #nfree, "k"
    Print #nfree, "d"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "f"
    Print #nfree, "h"
    Print #nfree, "s"
    Print #nfree, "l"
    Print #nfree, "g"
    Print #nfree, "q"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "k"
    Print #nfree, "s"
    Print #nfree, "m"
    Print #nfree, "q"
    Print #nfree, "g"
    Print #nfree, "s"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "s"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2B.txt", vpath & "Leçons\Personnalisé\leçon2B.txt"
End If

' leçon2C
If Dir(vpath & "Leçons\Standard\leçon2C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2C.txt" For Output As #nfree
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "z"  '20
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "l"
    Print #nfree, "o"
    Print #nfree, "l"
    Print #nfree, "k"
    Print #nfree, "i"
    Print #nfree, "k"
    Print #nfree, "m"
    Print #nfree, "p"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "l"  '40
    Print #nfree, "o"
    Print #nfree, "l"
    Print #nfree, "k"
    Print #nfree, "i"
    Print #nfree, "k"
    Print #nfree, "m"
    Print #nfree, "p"
    Print #nfree, "m"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "k"
    Print #nfree, "i"
    Print #nfree, "k"
    Print #nfree, "l"
    Print #nfree, "o"
    Print #nfree, "d"  '60
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "k"
    Print #nfree, "i"
    Print #nfree, "l"
    Print #nfree, "o"
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "q"
    Print #nfree, "m"
    Print #nfree, "p"
    Print #nfree, "m"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "q"  '80
    Print #nfree, "a"
    Print #nfree, "m"
    Print #nfree, "p"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "a"
    Print #nfree, "e"
    Print #nfree, "i"
    Print #nfree, "o"
    Print #nfree, "u"
    Print #nfree, "p"
    Print #nfree, "z"
    Print #nfree, "q"
    Print #nfree, "e"
    Print #nfree, "s"
    Print #nfree, "a"
    Print #nfree, "l"
    Print #nfree, "i"  '100
    Print #nfree, "k"
    Print #nfree, "o"
    Print #nfree, "a"
    Print #nfree, "e"
    Print #nfree, "i"
    Print #nfree, "o"
    Print #nfree, "u"
    Print #nfree, "p"
    Print #nfree, "z"
    Print #nfree, "q"
    Print #nfree, "e"
    Print #nfree, "s"
    Print #nfree, "a"
    Print #nfree, "l"
    Print #nfree, "i"
    Print #nfree, "k"
    Print #nfree, "o"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2C.txt", vpath & "Leçons\Personnalisé\leçon2C.txt"
End If

' leçon2D
If Dir(vpath & "Leçons\Standard\leçon2D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2D.txt" For Output As #nfree
    Print #nfree, "la rue"
    Print #nfree, "la rue"
    Print #nfree, "oui papa"
    Print #nfree, "oui papa"
    Print #nfree, "le roi"
    Print #nfree, "le roi"
    Print #nfree, "les rires"
    Print #nfree, "les rires"
    Print #nfree, "la foi"
    Print #nfree, "la foi"
    Print #nfree, "au ras du sol"
    Print #nfree, "au ras du sol"
    Print #nfree, "le zigzag"
    Print #nfree, "le zigzag"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2D.txt", vpath & "Leçons\Personnalisé\leçon2D.txt"
End If

' leçon2E
If Dir(vpath & "Leçons\Standard\leçon2E.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2E.txt" For Output As #nfree
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "g"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "g"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "t"
    Print #nfree, "p"
    Print #nfree, "t"
    Print #nfree, "p"
    Print #nfree, "t"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2E.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2E.txt", vpath & "Leçons\Personnalisé\leçon2E.txt"
End If

' leçon2F
If Dir(vpath & "Leçons\Standard\leçon2F.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2F.txt" For Output As #nfree
    Print #nfree, "papa fume"
    Print #nfree, "papa fume"
    Print #nfree, "du riz"
    Print #nfree, "du riz"
    Print #nfree, "la mer"
    Print #nfree, "la mer"
    Print #nfree, "sa pipe"
    Print #nfree, "sa pipe"
    Print #nfree, "qui parle"
    Print #nfree, "qui parle"
    Print #nfree, "je dis"
    Print #nfree, "je dis"
    Print #nfree, "des toits"
    Print #nfree, "des toits"
    Print #nfree, "du gaz"
    Print #nfree, "du gaz"
    Print #nfree, "le sel"
    Print #nfree, "le sel"
    Print #nfree, "je frappe"
    Print #nfree, "je frappe"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2F.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2F.txt", vpath & "Leçons\Personnalisé\leçon2F.txt"
End If

' leçon2G
If Dir(vpath & "Leçons\Standard\leçon2G.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2G.txt" For Output As #nfree
    Print #nfree, "le gros tas"
    Print #nfree, "le gros tas"
    Print #nfree, "elle aime lire"
    Print #nfree, "elle aime lire"
    Print #nfree, "sa jolie tasse"
    Print #nfree, "sa jolie tasse"
    Print #nfree, "du papier gris"
    Print #nfree, "du papier gris"
    Print #nfree, "la loi juste"
    Print #nfree, "la loi juste"
    Print #nfree, "le soleil luit"
    Print #nfree, "le soleil luit"
    Print #nfree, "quatre jupes rouges"
    Print #nfree, "quatre jupes rouges"
    Print #nfree, "faire le mur"
    Print #nfree, "faire le mur"
    Print #nfree, "la pierre roule"
    Print #nfree, "la pierre roule"
    Print #nfree, "le jus frais"
    Print #nfree, "le jus frais"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2G.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2G.txt", vpath & "Leçons\Personnalisé\leçon2G.txt"
End If

' leçon2H
If Dir(vpath & "Leçons\Standard\leçon2H.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon2H.txt" For Output As #nfree
    Print #nfree, "que dit le maire"
    Print #nfree, "que dit le maire"
    Print #nfree, "la pluie est froide"
    Print #nfree, "la pluie est froide"
    Print #nfree, "papa a pris le marteau"
    Print #nfree, "papa a pris le marteau"
    Print #nfree, "tu es si fragile"
    Print #nfree, "tu es si fragile"
    Print #nfree, "il est le premier et il progresse"
    Print #nfree, "il est le premier et il progresse"
    Print #nfree, "elle passe des films mais il les regarde peu"
    Print #nfree, "elle passe des films mais il les regarde peu"
    Print #nfree, "la roulotte de la jolie dame"
    Print #nfree, "la roulotte de la jolie dame"
    Print #nfree, "les hommes du hameau"
    Print #nfree, "les hommes du hameau"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon2H.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon2H.txt", vpath & "Leçons\Personnalisé\leçon2H.txt"
End If

' leçon3A
If Dir(vpath & "Leçons\Standard\leçon3A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon3A.txt" For Output As #nfree
    Print #nfree, "q"
    Print #nfree, "w"  'attention au type de police sinon w non prononcé !
    Print #nfree, "q"
    Print #nfree, "q"
    Print #nfree, "w"
    Print #nfree, "q"
    Print #nfree, "s" 'sx 7
    Print #nfree, "x"
    Print #nfree, "s"
    Print #nfree, "q"
    Print #nfree, "x"
    Print #nfree, "s"
    Print #nfree, "d" 'dc 13
    Print #nfree, "c"
    Print #nfree, "d"
    Print #nfree, "q"
    Print #nfree, "c"
    Print #nfree, "d"
    Print #nfree, "f" 'fv 19
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "q"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "w" 'reprise 25
    Print #nfree, "q"
    Print #nfree, "x"
    Print #nfree, "s"
    Print #nfree, "w"
    Print #nfree, "q"
    Print #nfree, "x"
    Print #nfree, "s"
    Print #nfree, "c"
    Print #nfree, "d"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "c"
    Print #nfree, "d"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "q"
    Print #nfree, "a" '3rangées, main gauche 42
    Print #nfree, "q"
    Print #nfree, "w"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "x"
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "q"
    Print #nfree, "w"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "x"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "c"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "v"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "c"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "c"
    Print #nfree, "v"
    Print #nfree, "j" 'j et n, main droite 74
    Print #nfree, "n"
    Print #nfree, "j"
    Print #nfree, "n"
    Print #nfree, "j"
    Print #nfree, "n"
    Print #nfree, "j"
    Print #nfree, "u" '3rangées, main droite 81
    Print #nfree, "j"
    Print #nfree, "n"
    Print #nfree, "u"
    Print #nfree, "n"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "n"
    Print #nfree, "u"
    Print #nfree, "n"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon3A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon3A.txt", vpath & "Leçons\Personnalisé\leçon3A.txt"
End If

' leçon3B
If Dir(vpath & "Leçons\Standard\leçon3B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon3B.txt" For Output As #nfree
    Print #nfree, "f"
    Print #nfree, "b"
    Print #nfree, "f"
    Print #nfree, "b"
    Print #nfree, "f"
    Print #nfree, "b"
    Print #nfree, "f" 'msg différence entre B et V
    Print #nfree, "b"
    Print #nfree, "f"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "b"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "p"
    Print #nfree, "v"
    Print #nfree, "b" 'msg attention, les autres consonnes
    Print #nfree, "c"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "p"
    Print #nfree, "b"
    Print #nfree, "v"
    Print #nfree, "p"
    Print #nfree, "u"
    Print #nfree, "i"
    Print #nfree, "u"
    Print #nfree, "i"
    Print #nfree, "d"
    Print #nfree, "b"
    Print #nfree, "t"
    Print #nfree, "v"
    Print #nfree, "p"
    Print #nfree, "i"
    Print #nfree, "u"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon3B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon3B.txt", vpath & "Leçons\Personnalisé\leçon3B.txt"
End If

' leçon3C
If Dir(vpath & "Leçons\Standard\leçon3C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon3C.txt" For Output As #nfree
    Print #nfree, "bercer"
    Print #nfree, "percer"
    Print #nfree, "cerner"
    Print #nfree, "verser"
    Print #nfree, "rester"
    Print #nfree, "taper"
    Print #nfree, "ferrer"
    Print #nfree, "presser"
    Print #nfree, "dresser"
    Print #nfree, "butiner"
    Print #nfree, "bitumer"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon3C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon3C.txt", vpath & "Leçons\Personnalisé\leçon3C.txt"
End If

' leçon4A
If Dir(vpath & "Leçons\Standard\leçon4A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4A.txt" For Output As #nfree
    Print #nfree, "q"
    Print #nfree, "d"
    Print #nfree, "d"
    Print #nfree, "q"
    Print #nfree, "f" 'msg : distinguer le f et le s
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "q"
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "d"
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "s"
    Print #nfree, "q"
    Print #nfree, "s"
    Print #nfree, "f"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4A.txt", vpath & "Leçons\Personnalisé\leçon4A.txt"
End If

' leçon4B
If Dir(vpath & "Leçons\Standard\leçon4B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4B.txt" For Output As #nfree
    Print #nfree, "j"
    Print #nfree, "k"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "k"
    Print #nfree, "m"
    Print #nfree, "j"
    Print #nfree, "l"
    Print #nfree, "m"
    Print #nfree, "l"
    Print #nfree, "j"
    Print #nfree, "k"
    Print #nfree, "s" 'msg avec les 2 mains
    Print #nfree, "l"
    Print #nfree, "q"
    Print #nfree, "m"
    Print #nfree, "d"
    Print #nfree, "k"
    Print #nfree, "l"
    Print #nfree, "q"
    Print #nfree, "m"
    Print #nfree, "s"
    Print #nfree, "f"
    Print #nfree, "j"
    Print #nfree, "f"
    Print #nfree, "s"
    Print #nfree, "k"
    Print #nfree, "m"
    Print #nfree, "l"
    Print #nfree, "s"
    Print #nfree, "m"
    Print #nfree, "q"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4B.txt", vpath & "Leçons\Personnalisé\leçon4B.txt"
End If

' leçon4C
If Dir(vpath & "Leçons\Standard\leçon4C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4C.txt" For Output As #nfree
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "q"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "s"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "f"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "d"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "s" ' avec des mots
    Print #nfree, "raz"
    Print #nfree, "sera"
    Print #nfree, "assez"
    Print #nfree, "rade"
    Print #nfree, "raser"
    Print #nfree, "fera"
    Print #nfree, "raz"
    Print #nfree, "sera"
    Print #nfree, "assez"
    Print #nfree, "assez"
    Print #nfree, "rade"
    Print #nfree, "raser"
    Print #nfree, "fera"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4C.txt", vpath & "Leçons\Personnalisé\leçon4C.txt"
End If

' leçon4D
If Dir(vpath & "Leçons\Standard\leçon4D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4D.txt" For Output As #nfree
    Print #nfree, "f"
    Print #nfree, "g"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "f"
    Print #nfree, "s"
    Print #nfree, "d"
    Print #nfree, "g"
    Print #nfree, "d"
    Print #nfree, "t"
    Print #nfree, "z"
    Print #nfree, "t"
    Print #nfree, "d"
    Print #nfree, "f"
    Print #nfree, "raser" ' avec des mots
    Print #nfree, "gaz"
    Print #nfree, "tasser"
    Print #nfree, "rate"
    Print #nfree, "fera"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4D.txt", vpath & "Leçons\Personnalisé\leçon4D.txt"
End If

' leçon4E
If Dir(vpath & "Leçons\Standard\leçon4E.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4E.txt" For Output As #nfree
    Print #nfree, "m"
    Print #nfree, "p"
    Print #nfree, "m"
    Print #nfree, "p"
    Print #nfree, "m"
    Print #nfree, "l"
    Print #nfree, "o"
    Print #nfree, "l"
    Print #nfree, "o"
    Print #nfree, "l"
    Print #nfree, "k"
    Print #nfree, "i"
    Print #nfree, "k"
    Print #nfree, "i"
    Print #nfree, "k"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "k"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "pour" 'avec des mots
    Print #nfree, "oui"
    Print #nfree, "osier"
    Print #nfree, "outillage"
    Print #nfree, "proie"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4E.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4E.txt", vpath & "Leçons\Personnalisé\leçon4E.txt"
End If

' leçon4F
If Dir(vpath & "Leçons\Standard\leçon4F.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4F.txt" For Output As #nfree
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "j"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "j"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "h"
    Print #nfree, "y"
    Print #nfree, "j"
    Print #nfree, "yaourts" 'avec des mots
    Print #nfree, "haie"
    Print #nfree, "yole"
    Print #nfree, "hall"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4F.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4F.txt", vpath & "Leçons\Personnalisé\leçon4F.txt"
End If

' leçon4G
If Dir(vpath & "Leçons\Standard\leçon4G.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4G.txt" For Output As #nfree
    Print #nfree, "q"
    Print #nfree, "w"
    Print #nfree, "q"
    Print #nfree, "a"
    Print #nfree, "w"
    Print #nfree, "q"
    Print #nfree, "w"
    Print #nfree, "a"
    Print #nfree, "s"
    Print #nfree, "x"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "x"
    Print #nfree, "s"
    Print #nfree, "x"
    Print #nfree, "z"
    Print #nfree, "d"
    Print #nfree, "c"
    Print #nfree, "d"
    Print #nfree, "e"
    Print #nfree, "c"
    Print #nfree, "d"
    Print #nfree, "c"
    Print #nfree, "e"
    Print #nfree, "f"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "v"
    Print #nfree, "r"
    Print #nfree, "c"
    Print #nfree, "w"
    Print #nfree, "v"
    Print #nfree, "x"
    Print #nfree, "v"
    Print #nfree, "c"
    Print #nfree, "x"
    Print #nfree, "w"
    Print #nfree, "saxe" 'avec des mots
    Print #nfree, "whist"
    Print #nfree, "trace"
    Print #nfree, "voir"
    Print #nfree, "watt"
    Print #nfree, "choix"
    Print #nfree, "vexer"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4G.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4G.txt", vpath & "Leçons\Personnalisé\leçon4G.txt"
End If

' leçon4H
If Dir(vpath & "Leçons\Standard\leçon4H.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon4H.txt" For Output As #nfree
    Print #nfree, "f"
    Print #nfree, "b"
    Print #nfree, "f"
    Print #nfree, "v"
    Print #nfree, "f"
    Print #nfree, "t"
    Print #nfree, "b"
    Print #nfree, "f"
    Print #nfree, "r"
    Print #nfree, "b"
    Print #nfree, "v"
    Print #nfree, "j"
    Print #nfree, "n"
    Print #nfree, "j"
    Print #nfree, "u"
    Print #nfree, "n"
    Print #nfree, "j"
    Print #nfree, "h"
    Print #nfree, "n"
    Print #nfree, "y"
    Print #nfree, "n"
    Print #nfree, "b"
    Print #nfree, "n"
    Print #nfree, "u"
    Print #nfree, "braver le danger" 'avec des mots
    Print #nfree, "braver le danger"
    Print #nfree, "un biberon de bouillie"
    Print #nfree, "un biberon de bouillie"
    Print #nfree, "bien valider le bon choix"
    Print #nfree, "bien valider le bon choix"
    Print #nfree, "au bazar du coin"
    Print #nfree, "au bazar du coin"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon4H.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon4H.txt", vpath & "Leçons\Personnalisé\leçon4H.txt"
End If

' leçon5A
If Dir(vpath & "Leçons\Standard\leçon5A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon5A.txt" For Output As #nfree
    Print #nfree, "le convoi"
    Print #nfree, "le convoi"
    Print #nfree, "ton xylophone"
    Print #nfree, "ton xylophone"
    Print #nfree, "que comprendre"
    Print #nfree, "que comprendre"
    Print #nfree, "les manuels"
    Print #nfree, "les manuels"
    Print #nfree, "boire du whisky"
    Print #nfree, "boire du whisky"
    Print #nfree, "il a du nez"
    Print #nfree, "il a du nez"
    Print #nfree, "avoir du jeu"
    Print #nfree, "avoir du jeu"
    Print #nfree, "un beau western"
    Print #nfree, "un beau western"
    Print #nfree, "une grande girafe"
    Print #nfree, "une grande girafe"
    Print #nfree, "des vieux ponts"
    Print #nfree, "des vieux ponts"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon5A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon5A.txt", vpath & "Leçons\Personnalisé\leçon5A.txt"
End If

' leçon5B
If Dir(vpath & "Leçons\Standard\leçon5B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon5B.txt" For Output As #nfree
    Print #nfree, "le chien aboie la caravane passe"
    Print #nfree, "le chien aboie la caravane passe"
    Print #nfree, "tel qui rit vendredi dimanche pleurera"
    Print #nfree, "tel qui rit vendredi dimanche pleurera"
    Print #nfree, "A bon chat bon rat"
    Print #nfree, "A bon chat bon rat"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon5B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon5B.txt", vpath & "Leçons\Personnalisé\leçon5B.txt"
End If

' leçon5C
If Dir(vpath & "Leçons\Standard\leçon5C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon5C.txt" For Output As #nfree
    Print #nfree, "Il faut faire des exercices qui se compliquent"
    Print #nfree, "Il faut faire des exercices qui se compliquent"
    Print #nfree, "Papa nous passe beaucoup de films en couleurs"
    Print #nfree, "Papa nous passe beaucoup de films en couleurs"
    Print #nfree, "Le cheval cavale au fond du vallon"
    Print #nfree, "Le cheval cavale au fond du vallon"
    Print #nfree, "Elle file de la laine bleue pour un uniforme de grand zouave"
    Print #nfree, "Elle file de la laine bleue pour un uniforme de grand zouave"
    Print #nfree, "Un tramway tire six wagons"
    Print #nfree, "Un tramway tire six wagons"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon5C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon5C.txt", vpath & "Leçons\Personnalisé\leçon5C.txt"
End If

' leçon6A
If Dir(vpath & "Leçons\Standard\leçon6A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon6A.txt" For Output As #nfree
    Print #nfree, "amiral"
    Print #nfree, "artiste"
    Print #nfree, "bermuda"
    Print #nfree, "bobine"
    Print #nfree, "briller"
    Print #nfree, "bravo"
    Print #nfree, "cheval"
    Print #nfree, "choc"
    Print #nfree, "clou"
    Print #nfree, "coton"
    Print #nfree, "cri"
    Print #nfree, "culot"
    Print #nfree, "delta"
    Print #nfree, "distrait"
    Print #nfree, "ennemi"
    Print #nfree, "ensuite"
    Print #nfree, "esprit"
    Print #nfree, "facile"
    Print #nfree, "farine"
    Print #nfree, "gel"
    Print #nfree, "gouffre"
    Print #nfree, "goulot"
    Print #nfree, "grossier"
    Print #nfree, "haricot"
    Print #nfree, "histoire"
    Print #nfree, "incendie"
    Print #nfree, "insecte"
    Print #nfree, "ivoire"
    Print #nfree, "jaloux"
    Print #nfree, "janvier"
    Print #nfree, "kilo"
    Print #nfree, "klaxon"
    Print #nfree, "limite"
    Print #nfree, "lune"
    Print #nfree, "magique"
    Print #nfree, "malin"
    Print #nfree, "mariage"
    Print #nfree, "marteau"
    Print #nfree, "meilleur"
    Print #nfree, "mourir"
    Print #nfree, "nacre"
    Print #nfree, "navet"
    Print #nfree, "parasol"
    Print #nfree, "pauvre"
    Print #nfree, "poison"
    Print #nfree, "pratique"
    Print #nfree, "prime"
    Print #nfree, "quota"
    Print #nfree, "rapace"
    Print #nfree, "rigueur"
    Print #nfree, "sabot"
    Print #nfree, "servir"
    Print #nfree, "soleil"
    Print #nfree, "soldat"
    Print #nfree, "table"
    Print #nfree, "taxe"
    Print #nfree, "tumeur"
    Print #nfree, "usine"
    Print #nfree, "valse"
    Print #nfree, "viaduc"
    Print #nfree, "voyou"
    Print #nfree, "wagon"
    Print #nfree, "yoga"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon6A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon6A.txt", vpath & "Leçons\Personnalisé\leçon6A.txt"
End If

' leçon6B
If Dir(vpath & "Leçons\Standard\leçon6B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon6B.txt" For Output As #nfree
    Print #nfree, "amiral"
    Print #nfree, "artiste"
    Print #nfree, "bermuda"
    Print #nfree, "bobine"
    Print #nfree, "briller"
    Print #nfree, "bravo"
    Print #nfree, "cheval"
    Print #nfree, "choc"
    Print #nfree, "clou"
    Print #nfree, "coton"
    Print #nfree, "cri"
    Print #nfree, "culot"
    Print #nfree, "delta"
    Print #nfree, "distrait"
    Print #nfree, "ennemi"
    Print #nfree, "ensuite"
    Print #nfree, "esprit"
    Print #nfree, "facile"
    Print #nfree, "farine"
    Print #nfree, "gel"
    Print #nfree, "gouffre"
    Print #nfree, "goulot"
    Print #nfree, "grossier"
    Print #nfree, "haricot"
    Print #nfree, "histoire"
    Print #nfree, "incendie"
    Print #nfree, "insecte"
    Print #nfree, "ivoire"
    Print #nfree, "jaloux"
    Print #nfree, "janvier"
    Print #nfree, "kilo"
    Print #nfree, "klaxon"
    Print #nfree, "limite"
    Print #nfree, "lune"
    Print #nfree, "magique"
    Print #nfree, "malin"
    Print #nfree, "mariage"
    Print #nfree, "marteau"
    Print #nfree, "meilleur"
    Print #nfree, "mourir"
    Print #nfree, "nacre"
    Print #nfree, "navet"
    Print #nfree, "parasol"
    Print #nfree, "pauvre"
    Print #nfree, "poison"
    Print #nfree, "pratique"
    Print #nfree, "prime"
    Print #nfree, "quota"
    Print #nfree, "rapace"
    Print #nfree, "rigueur"
    Print #nfree, "sabot"
    Print #nfree, "servir"
    Print #nfree, "soleil"
    Print #nfree, "soldat"
    Print #nfree, "table"
    Print #nfree, "taxe"
    Print #nfree, "tumeur"
    Print #nfree, "usine"
    Print #nfree, "valse"
    Print #nfree, "viaduc"
    Print #nfree, "voyou"
    Print #nfree, "wagon"
    Print #nfree, "yoga"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon6B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon6A.txt", vpath & "Leçons\Personnalisé\leçon6B.txt"
End If


' leçon6C
If Dir(vpath & "Leçons\Standard\leçon6C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon6C.txt" For Output As #nfree
    Print #nfree, "analogie"
    Print #nfree, "artillerie"
    Print #nfree, "auparavant"
    Print #nfree, "bavardage"
    Print #nfree, "balourdise"
    Print #nfree, "binoculaire"
    Print #nfree, "blanchisseur"
    Print #nfree, "califourchon"
    Print #nfree, "camouflage"
    Print #nfree, "cassoulet"
    Print #nfree, "compagnie"
    Print #nfree, "courbature"
    Print #nfree, "critiquable"
    Print #nfree, "dictionnaire"
    Print #nfree, "enneigement"
    Print #nfree, "enthousiasme"
    Print #nfree, "excellent"
    Print #nfree, "fortune"
    Print #nfree, "funiculaire"
    Print #nfree, "gastronome"
    Print #nfree, "harmonie"
    Print #nfree, "implacable"
    Print #nfree, "involontaire"
    Print #nfree, "jonquille"
    Print #nfree, "kermesse"
    Print #nfree, "lexique"
    Print #nfree, "lamentable"
    Print #nfree, "maladresse"
    Print #nfree, "menuisier"
    Print #nfree, "moustache"
    Print #nfree, "naufrage"
    Print #nfree, "nostalgie"
    Print #nfree, "obstacle"
    Print #nfree, "paratonnerre"
    Print #nfree, "parcimonie"
    Print #nfree, "partenariat"
    Print #nfree, "photographie"
    Print #nfree, "piraterie"
    Print #nfree, "profession"
    Print #nfree, "quotidien"
    Print #nfree, "rafistoler"
    Print #nfree, "ralentissement"
    Print #nfree, "ritournelle"
    Print #nfree, "sacrifice"
    Print #nfree, "sauvegarde"
    Print #nfree, "spectacle"
    Print #nfree, "symphonie"
    Print #nfree, "tapisserie"
    Print #nfree, "territoire"
    Print #nfree, "transparent"
    Print #nfree, "uniforme"
    Print #nfree, "utilisable"
    Print #nfree, "ventriloque"
    Print #nfree, "vitamine"
    Print #nfree, "xylophone"
    Print #nfree, "zodiaque"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon6C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon6B.txt", vpath & "Leçons\Personnalisé\leçon6C.txt"
End If

' leçon6D
If Dir(vpath & "Leçons\Standard\leçon6D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon6D.txt" For Output As #nfree
    Print #nfree, "abrogation"
    Print #nfree, "accentuation"
    Print #nfree, "association"
    Print #nfree, "attribution"
    Print #nfree, "bifurcation"
    Print #nfree, "calcification"
    Print #nfree, "circulation"
    Print #nfree, "civilisation"
    Print #nfree, "codification"
    Print #nfree, "cohabitation"
    Print #nfree, "confirmation"
    Print #nfree, "conjuration"
    Print #nfree, "consolation"
    Print #nfree, "contestation"
    Print #nfree, "distribution"
    Print #nfree, "documentation"
    Print #nfree, "embarcation"
    Print #nfree, "exploration"
    Print #nfree, "civilisation"
    Print #nfree, "fabrication"
    Print #nfree, "facturation"
    Print #nfree, "fluctuation"
    Print #nfree, "graduation"
    Print #nfree, "habitation"
    Print #nfree, "liquidation"
    Print #nfree, "machination"
    Print #nfree, "malversation"
    Print #nfree, "normalisation"
    Print #nfree, "observation"
    Print #nfree, "obligation"
    Print #nfree, "participation"
    Print #nfree, "personnalisation"
    Print #nfree, "plantation"
    Print #nfree, "procuration"
    Print #nfree, "recommandation"
    Print #nfree, "scolarisation"
    Print #nfree, "situation"
    Print #nfree, "tergiversation"
    Print #nfree, "tarification"
    Print #nfree, "titularisation"
    Print #nfree, "tractation"
    Print #nfree, "tribulation"
    Print #nfree, "variation"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon6D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon6C.txt", vpath & "Leçons\Personnalisé\leçon6D.txt"
End If

' leçon7A
If Dir(vpath & "Leçons\Standard\leçon7A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon7A.txt" For Output As #nfree
    Print #nfree, "Il est heureux de tisser un vrai tapis"
    Print #nfree, "Elle accepte de visiter un grand magasin"
    Print #nfree, "Mon mari est un homme au sourire si doux"  '12/2011 ajout
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon7A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon7A.txt", vpath & "Leçons\Personnalisé\leçon7A.txt"
End If

' leçon7B
If Dir(vpath & "Leçons\Standard\leçon7B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon7B.txt" For Output As #nfree
    Print #nfree, "Les clochettes du muguet tapissent le jardin"
    Print #nfree, "Le myosotis est une fleur du joli mois de mai"
    Print #nfree, "Elle quitta le porche pour venir vers lui"  '12/2011 ajout
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon7B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon7B.txt", vpath & "Leçons\Personnalisé\leçon7B.txt"
End If

' leçon7C
If Dir(vpath & "Leçons\Standard\leçon7C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon7C.txt" For Output As #nfree
    Print #nfree, "Xavier est le premier de sa classe et il progresse sans cesse"
    Print #nfree, "Bernard se rapproche beaucoup du grand radiateur"
    Print #nfree, "Ses beaux cheveux blonds ondulent dans le vent"  '12/2011 ajout
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon7C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon7C.txt", vpath & "Leçons\Personnalisé\leçon7C.txt"
End If

End Sub
