Attribute VB_Name = "Module_leçons2"
'******************  RESET_STANDARD : RESET des leçons STANDARD  **************************
Public Sub reset2_standard(force As Byte)
' Création des reps Standard et Personnalisé s'il n'existent pas encore
On Error Resume Next
MkDir vpath & "Leçons"
On Error Resume Next
MkDir vpath & "Leçons\Standard"
On Error Resume Next
MkDir vpath & "Leçons\Personnalisé"
nfree = FreeFile

' leçon8A
If Dir(vpath & "Leçons\Standard\leçon8A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8A.txt" For Output As #nfree
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajGauche
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajGauche
    Print #nfree, vvMajDroit
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajGauche
    Print #nfree, vvMajDroit
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajDroit
    Print #nfree, vvMajGauche
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajDroit
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajGauche
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8A.txt", vpath & "Leçons\Personnalisé\leçon8A.txt"
End If

' leçon8B
If Dir(vpath & "Leçons\Standard\leçon8B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8B.txt" For Output As #nfree
    Print #nfree, ","
    Print #nfree, ";"
    Print #nfree, ":"
    Print #nfree, "!"
    Print #nfree, ";"
    Print #nfree, ","
    Print #nfree, "!"
    Print #nfree, ","
    Print #nfree, ":"
    Print #nfree, "!"
    Print #nfree, ";"
    Print #nfree, ","
    Print #nfree, "!"
    Print #nfree, ","
    Print #nfree, ":"
    Print #nfree, ";"
    Print #nfree, "Le mur, il faut le repeindre ; quel boulot ! Tant pis !"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8B.txt", vpath & "Leçons\Personnalisé\leçon8B.txt"
End If

' leçon8C
If Dir(vpath & "Leçons\Standard\leçon8C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8C.txt" For Output As #nfree
    Print #nfree, "?"
    Print #nfree, "."
    Print #nfree, "/"
    Print #nfree, "§"
    Print #nfree, "?"
    Print #nfree, "."
    Print #nfree, "?"
    Print #nfree, "/"
    Print #nfree, "."
    Print #nfree, "§"
    Print #nfree, "?"
    Print #nfree, "/" ' Attention, toutes les ponctuations
    Print #nfree, ","
    Print #nfree, "?"
    Print #nfree, ";"
    Print #nfree, ":"
    Print #nfree, "/"
    Print #nfree, "!"
    Print #nfree, ","
    Print #nfree, "?"
    Print #nfree, ";"
    Print #nfree, ":"
    Print #nfree, "/"
    Print #nfree, "!"
    Print #nfree, "?"
    Print #nfree, ":"
    Print #nfree, ","
    Print #nfree, ";"
    Print #nfree, "/"
    Print #nfree, ":"
    Print #nfree, "!"
    Print #nfree, "?"
    Print #nfree, ":"
    Print #nfree, ","
    Print #nfree, ";"
    Print #nfree, "/"
    Print #nfree, "!"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8C.txt", vpath & "Leçons\Personnalisé\leçon8C.txt"
End If

' leçon8D
If Dir(vpath & "Leçons\Standard\leçon8D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8D.txt" For Output As #nfree
    Print #nfree, "Quel est ton nom ?"
    Print #nfree, "Quel est ton nom ?"
    Print #nfree, "Quelle chance tu as !"
    Print #nfree, "Quelle chance tu as !"
    Print #nfree, "Le kiosque est proche ; Pierre ira tout de suite."
    Print #nfree, "Le kiosque est proche ; Pierre ira tout de suite."
    Print #nfree, "Le T.G.V. traverse la C.E.E."
    Print #nfree, "Le T.G.V. traverse la C.E.E."
    Print #nfree, "Jacques, comme Robert, fait son chemin."
    Print #nfree, "Jacques, comme Robert, fait son chemin."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8D.txt", vpath & "Leçons\Personnalisé\leçon8D.txt"
End If

' leçon8E
If Dir(vpath & "Leçons\Standard\leçon8E.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8E.txt" For Output As #nfree
    Print #nfree, "â"
    Print #nfree, "ë"
    Print #nfree, "â"
    Print #nfree, "î"
    Print #nfree, "ù"
    Print #nfree, "ë"
    Print #nfree, "ù"
    Print #nfree, "Un âne pour Noël"
    Print #nfree, "Un âne pour Noël"
    Print #nfree, "Une ouïe fine"
    Print #nfree, "Une ouïe fine"
    Print #nfree, "La boîte de Mikaël"
    Print #nfree, "La boîte de Mikaël"
    Print #nfree, "Où dîner ?"
    Print #nfree, "Où dîner ?"
    Print #nfree, "Une mûre fraîche"
    Print #nfree, "Une mûre fraîche"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8E.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8E.txt", vpath & "Leçons\Personnalisé\leçon8E.txt"
End If

' leçon8F
If Dir(vpath & "Leçons\Standard\leçon8F.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8F.txt" For Output As #nfree
    Print #nfree, "*"
    Print #nfree, "<"
    Print #nfree, ">"
    Print #nfree, "*"
    Print #nfree, ">"
    Print #nfree, "<"
    Print #nfree, "*"
    Print #nfree, "<"
    Print #nfree, ">"
    Print #nfree, "*"
    Print #nfree, ">"
    Print #nfree, "<"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8F.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8F.txt", vpath & "Leçons\Personnalisé\leçon8F.txt"
End If

' leçon8G
If Dir(vpath & "Leçons\Standard\leçon8G.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8G.txt" For Output As #nfree
    Print #nfree, "%"
    Print #nfree, "µ"
    Print #nfree, "$"
    Print #nfree, "£"
    Print #nfree, "µ"
    Print #nfree, "$"
    Print #nfree, "£"
    Print #nfree, "%"
    Print #nfree, "£"
    Print #nfree, "µ"
    Print #nfree, "$"
    Print #nfree, "%"
    Print #nfree, "$"
    Print #nfree, "µ"
    Print #nfree, "%"
    Print #nfree, "$"
    Print #nfree, "£"
    Print #nfree, "µ"
    Print #nfree, "$"
    Print #nfree, "%"
    Print #nfree, "µ"
    Print #nfree, "£"
    Print #nfree, "%"
    Print #nfree, "£"
    Print #nfree, "µ"
    Print #nfree, "$"
    Print #nfree, "%"
    Print #nfree, "ù"  ' Attention
    Print #nfree, "£"
    Print #nfree, "*"
    Print #nfree, "%"
    Print #nfree, "µ"
    Print #nfree, "ù"
    Print #nfree, "$"
    Print #nfree, "£"
    Print #nfree, "*"
    Print #nfree, "£"
    Print #nfree, "$"
    Print #nfree, "ù"
    Print #nfree, "$"
    Print #nfree, "µ"
    Print #nfree, "ù"
    Print #nfree, "£"
    Print #nfree, "*"
    Print #nfree, "%"
    Print #nfree, "µ"
    Print #nfree, "ù"
    Print #nfree, "$"
    Print #nfree, "£"
    Print #nfree, "*"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8G.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8G.txt", vpath & "Leçons\Personnalisé\leçon8G.txt"
End If

' leçon8H
If Dir(vpath & "Leçons\Standard\leçon8H.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon8H.txt" For Output As #nfree
    Print #nfree, "é"
    Print #nfree, "à"
    Print #nfree, "è"
    Print #nfree, "é"
    Print #nfree, "à"
    Print #nfree, "é"
    Print #nfree, "à"
    Print #nfree, "è"
    Print #nfree, "é"
    Print #nfree, "è"
    Print #nfree, "à"
    Print #nfree, "è"
    Print #nfree, "à"
    Print #nfree, "é"
    Print #nfree, "è"
    Print #nfree, "Un été tempéré vient après un piètre hiver."
    Print #nfree, "Un été tempéré vient après un piètre hiver."
    Print #nfree, "Il gère très bien le travail à Paris."
    Print #nfree, "Il gère très bien le travail à Paris."
    Print #nfree, "Il emmène sa fille préférée à Moscou."
    Print #nfree, "Il emmène sa fille préférée à Moscou."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon8H.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon8H.txt", vpath & "Leçons\Personnalisé\leçon8H.txt"
End If

' leçon9A
If Dir(vpath & "Leçons\Standard\leçon9A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon9A.txt" For Output As #nfree
    Print #nfree, "&"
    Print #nfree, "²"
    Print #nfree, "&"
    Print #nfree, "²"
    Print #nfree, "%"
    Print #nfree, "&"
    Print #nfree, "²"
    Print #nfree, "*"
    Print #nfree, "&"
    Print #nfree, "$"
    Print #nfree, "²"
    Print #nfree, "é"
    Print #nfree, "&"
    Print #nfree, "²"
    Print #nfree, "ù"
    Print #nfree, "%"
    Print #nfree, "²"
    Print #nfree, "*"
    Print #nfree, "&"
    Print #nfree, "$"
    Print #nfree, "²"
    Print #nfree, "é"
    Print #nfree, "ù"
    Print #nfree, "&"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon9A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon9A.txt", vpath & "Leçons\Personnalisé\leçon9A.txt"
End If

' leçon9B
If Dir(vpath & "Leçons\Standard\leçon9B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon9B.txt" For Output As #nfree
    Print #nfree, "-"
    Print #nfree, "("
    Print #nfree, "'"
    Print #nfree, """"
    Print #nfree, "-"
    Print #nfree, "'"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, """"
    Print #nfree, "-"
    Print #nfree, "("
    Print #nfree, "'"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, "-"
    Print #nfree, "'"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, "C'est-à-dire"
    Print #nfree, "C'est-à-dire"
    Print #nfree, "J'irai, iras-tu aussi ?"
    Print #nfree, "J'irai, iras-tu aussi ?"
    Print #nfree, "Y a-t-il de l'eau, potable bien sûr !"
    Print #nfree, "Y a-t-il de l'eau, potable bien sûr !"
    Print #nfree, "Il dit simplement : ""Viens vite !"""
    Print #nfree, "Il dit simplement : ""Viens vite !"""
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon9B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon9B.txt", vpath & "Leçons\Personnalisé\leçon9B.txt"
End If

' leçon9C
If Dir(vpath & "Leçons\Standard\leçon9C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon9C.txt" For Output As #nfree
    Print #nfree, "_"
    Print #nfree, ")"
    Print #nfree, "_"
    Print #nfree, ")"
    Print #nfree, "-"
    Print #nfree, "'"
    Print #nfree, "_"
    Print #nfree, ")"
    Print #nfree, "_"
    Print #nfree, "-"
    Print #nfree, "'"
    Print #nfree, "&"
    Print #nfree, "("
    Print #nfree, """"
    Print #nfree, "_"
    Print #nfree, "²"
    Print #nfree, ")"
    Print #nfree, "'"
    Print #nfree, "&"
    Print #nfree, "("
    Print #nfree, "_"
    Print #nfree, "-"
    Print #nfree, """"
    Print #nfree, "²"
    Print #nfree, ")"
    Print #nfree, "-"
    Print #nfree, "'"
    Print #nfree, "&"
    Print #nfree, "("
    Print #nfree, """"
    Print #nfree, "²"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon9C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon9C.txt", vpath & "Leçons\Personnalisé\leçon9C.txt"
End If

' leçon9D
If Dir(vpath & "Leçons\Standard\leçon9D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon9D.txt" For Output As #nfree
    Print #nfree, "°"
    Print #nfree, "="
    Print #nfree, "a+b"
    Print #nfree, "°"
    Print #nfree, "ç"
    Print #nfree, "="
    Print #nfree, "c+d"
    Print #nfree, "ç"
    Print #nfree, "°"
    Print #nfree, "x+y"
    Print #nfree, "="
    Print #nfree, "°"
    Print #nfree, "ç"
    Print #nfree, "="
    Print #nfree, "g+h"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, ")"
    Print #nfree, "°"
    Print #nfree, "-"
    Print #nfree, "="
    Print #nfree, "_"
    Print #nfree, "'"
    Print #nfree, "i+j"
    Print #nfree, "ç"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, ")"
    Print #nfree, "°"
    Print #nfree, "-"
    Print #nfree, "="
    Print #nfree, "_"
    Print #nfree, "'"
    Print #nfree, "k+l"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon9D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon9D.txt", vpath & "Leçons\Personnalisé\leçon9D.txt"
End If

' leçon10A
If Dir(vpath & "Leçons\Standard\leçon10A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon10A.txt" For Output As #nfree
    Print #nfree, "1"
    Print #nfree, "2"
    Print #nfree, "3"
    Print #nfree, "4"
    Print #nfree, "2"
    Print #nfree, "1"
    Print #nfree, "4"
    Print #nfree, "3"
    Print #nfree, "2"
    Print #nfree, "4"
    Print #nfree, "1"
    Print #nfree, "2"
    Print #nfree, "3"
    Print #nfree, "4"
    Print #nfree, "2"
    Print #nfree, "1"
    Print #nfree, "4"
    Print #nfree, "3"
    Print #nfree, "2"
    Print #nfree, "4"
    Print #nfree, "13"   'Attention nombres à 2 chiffres
    Print #nfree, "24"
    Print #nfree, "41"
    Print #nfree, "43"
    Print #nfree, "32"
    Print #nfree, "13"
    Print #nfree, "24"
    Print #nfree, "41"
    Print #nfree, "42"
    Print #nfree, "32"
    Print #nfree, "123"  'Attention nombres à 3 chiffres"
    Print #nfree, "234"
    Print #nfree, "341"
    Print #nfree, "412"
    Print #nfree, "132"
    Print #nfree, "234"
    Print #nfree, "341"
    Print #nfree, "412"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon10A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon10A.txt", vpath & "Leçons\Personnalisé\leçon10A.txt"
    End If

' leçon10B
If Dir(vpath & "Leçons\Standard\leçon10B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon10B.txt" For Output As #nfree
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "5"
    Print #nfree, "6"
    Print #nfree, "7"
    Print #nfree, "5"
    Print #nfree, "6"
    Print #nfree, "5"
    Print #nfree, "6"
    Print #nfree, "7"
    Print #nfree, "5"
    Print #nfree, "1"  'Attention les 7 premiers chiffres
    Print #nfree, "3"
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "2"
    Print #nfree, "4"
    Print #nfree, "6"
    Print #nfree, "7"
    Print #nfree, "1"
    Print #nfree, "4"
    Print #nfree, "3"
    Print #nfree, "6"
    Print #nfree, "2"
    Print #nfree, "17"   'Attention voici des nombres
    Print #nfree, "36"
    Print #nfree, "75"
    Print #nfree, "25"
    Print #nfree, "72"
    Print #nfree, "47"
    Print #nfree, "567"
    Print #nfree, "675"
    Print #nfree, "756"
    Print #nfree, "765"
    Print #nfree, "567"
    Print #nfree, "675"
    Print #nfree, "756"
    Print #nfree, "135"
    Print #nfree, "246"
    Print #nfree, "72"
    Print #nfree, "357"
    Print #nfree, "61"
    Print #nfree, "36"
    Print #nfree, "627"
    Print #nfree, "25"
    Print #nfree, "135"
    Print #nfree, "17"
    Print #nfree, "246"
    Print #nfree, "72"
    Print #nfree, "357"
    Print #nfree, "61"
    Print #nfree, "36"
    Print #nfree, "627"
    Print #nfree, "25"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon10B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon10B.txt", vpath & "Leçons\Personnalisé\leçon10B.txt"
    End If

' leçon10C
If Dir(vpath & "Leçons\Standard\leçon10C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon10C.txt" For Output As #nfree
    Print #nfree, "8"
    Print #nfree, "6"
    Print #nfree, "8"
    Print #nfree, "6"
    Print #nfree, "8"
    Print #nfree, "6"
    Print #nfree, "9"
    Print #nfree, "0"
    Print #nfree, "9"
    Print #nfree, "8"
    Print #nfree, "0"
    Print #nfree, "9"
    Print #nfree, "0"
    Print #nfree, "8"
    Print #nfree, "9"
    Print #nfree, "8"
    Print #nfree, "0"
    Print #nfree, "9"
    Print #nfree, "8"
    Print #nfree, "0"
    Print #nfree, "8"
    Print #nfree, "18"
    Print #nfree, "81"
    Print #nfree, "68"
    Print #nfree, "86"
    Print #nfree, "93"
    Print #nfree, "39"
    Print #nfree, "10"
    Print #nfree, "50"
    Print #nfree, "98"
    Print #nfree, "80"
    Print #nfree, "98"
    Print #nfree, "08"
    Print #nfree, "80"
    Print #nfree, "901"
    Print #nfree, "279"
    Print #nfree, "980"
    Print #nfree, "630"
    Print #nfree, "591"
    Print #nfree, "333"
    Print #nfree, "929"
    Print #nfree, "135"
    Print #nfree, "402"
    Print #nfree, "1608"
    Print #nfree, "37,2"
    Print #nfree, "51.8"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon10C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon10C.txt", vpath & "Leçons\Personnalisé\leçon10C.txt"
    End If

' leçon11A
If Dir(vpath & "Leçons\Standard\leçon11A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon11A.txt" For Output As #nfree
    Print #nfree, "Qui va à la chasse perd sa place"
    Print #nfree, "Qu'il est doux de ne rien faire quand tout s'agite autour de vous"
    Print #nfree, "C'était un espagnol de l'armée en déroute" '12/2011 ajout
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon11A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon11A.txt", vpath & "Leçons\Personnalisé\leçon11A.txt"
End If

' leçon11B
If Dir(vpath & "Leçons\Standard\leçon11B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon11B.txt" For Output As #nfree
    Print #nfree, "Noël au balcon, Pâques aux tisons"
    Print #nfree, "Nous passerons Noël au château"
    Print #nfree, "Donne-lui tout de même à boire, dit mon père" '12/2011 ajout
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon11B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon11B.txt", vpath & "Leçons\Personnalisé\leçon11B.txt"
End If

' leçon11C
If Dir(vpath & "Leçons\Standard\leçon11C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon11C.txt" For Output As #nfree
    Print #nfree, "Le général De Gaulle arriva à la tête de la France en 1958"
    Print #nfree, "Les révolutionnaires ont pris la Bastille le 14 juillet 1789"
    Print #nfree, "Il vécut au Canada et aux Etats-Unis de 1908 à 1921" '12/2011 ajout
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon11C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon11C.txt", vpath & "Leçons\Personnalisé\leçon11C.txt"
End If

' leçon12A
If Dir(vpath & "Leçons\Standard\leçon12A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon12A.txt" For Output As #nfree
    Print #nfree, vvSuppression
    Print #nfree, vvInsertion
    Print #nfree, vvSuppression
    Print #nfree, vvInsertion
    Print #nfree, "m"  ' Avec les lettres du clavier
    Print #nfree, vvSuppression
    Print #nfree, "s"
    Print #nfree, "j"
    Print #nfree, vvSuppression
    Print #nfree, "l"
    Print #nfree, vvInsertion
    Print #nfree, "q"
    Print #nfree, vvSuppression
    Print #nfree, "d"
    Print #nfree, "k"
    Print #nfree, vvInsertion
    Print #nfree, "f"
    Print #nfree, vvInsertion
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon12A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon12A.txt", vpath & "Leçons\Personnalisé\leçon12A.txt"
End If

' leçon12B
If Dir(vpath & "Leçons\Standard\leçon12B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon12B.txt" For Output As #nfree
    Print #nfree, vvDébut
    Print #nfree, vvFin
    Print #nfree, vvFin
    Print #nfree, vvDébut
    Print #nfree, vvSuppression  ' Avec les autres touches
    Print #nfree, vvDébut
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, vvSuppression
    Print #nfree, vvDébut
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, vvSuppression
    Print #nfree, vvDébut
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, "q"            ' Avec les lettres
    Print #nfree, vvDébut
    Print #nfree, "m"
    Print #nfree, vvSuppression
    Print #nfree, "f"
    Print #nfree, vvFin
    Print #nfree, "n"
    Print #nfree, vvInsertion
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon12B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon12B.txt", vpath & "Leçons\Personnalisé\leçon12B.txt"
End If

' leçon12C
If Dir(vpath & "Leçons\Standard\leçon12C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon12C.txt" For Output As #nfree
    Print #nfree, vvPagePrécédente
    Print #nfree, vvPageSuivante
    Print #nfree, vvPagePrécédente
    Print #nfree, vvPageSuivante
    Print #nfree, vvFin
    Print #nfree, vvPagePrécédente
    Print #nfree, vvPageSuivante
    Print #nfree, vvSuppression
    Print #nfree, vvDébut
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, vvSuppression
    Print #nfree, vvPagePrécédente
    Print #nfree, vvDébut
    Print #nfree, vvInsertion
    Print #nfree, vvPageSuivante
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon12C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon12C.txt", vpath & "Leçons\Personnalisé\leçon12C.txt"
End If

' leçon12D
If Dir(vpath & "Leçons\Standard\leçon12D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon12D.txt" For Output As #nfree
    Print #nfree, vvImpression
    Print #nfree, vvPause
    Print #nfree, vvArrêtDéfil
    Print #nfree, vvPause
    Print #nfree, vvImpression
    Print #nfree, vvArrêtDéfil
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon12D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon12D.txt", vpath & "Leçons\Personnalisé\leçon12D.txt"
End If

' leçon13A
' Attention, Jaws401 perturbe le AltGr, il faut toujours le faire suivre par Control puis Espace
If Dir(vpath & "Leçons\Standard\leçon13A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13A.txt" For Output As #nfree
    Print #nfree, vvControlGauche
    Print #nfree, vvEspace
    Print #nfree, vvAlt
    Print #nfree, vvWindowsGauche
    Print #nfree, vvControlGauche
    Print #nfree, vvEspace
    Print #nfree, vvWindowsGauche
    Print #nfree, vvAlt
    Print #nfree, vvAltGr         'Ajoute AltGr, Windows-droit, Control-Droit
    Print #nfree, vvControlDroit
    Print #nfree, vvEspace
    Print #nfree, vvWindowsDroit
    Print #nfree, vvControlDroit
    Print #nfree, vvWindowsDroit
    Print #nfree, vvAltGr
    Print #nfree, vvControlDroit
    Print #nfree, vvEspace
    Print #nfree, vvMenuContextuel 'Ajoute menu-contextuel
    Print #nfree, vvControlDroit
    Print #nfree, vvEspace
    Print #nfree, vvWindowsDroit
    Print #nfree, vvWindowsGauche
    Print #nfree, vvMenuContextuel
    Print #nfree, vvMajDroit  'Avec d'autres touches
    Print #nfree, vvControlDroit
    Print #nfree, vvEspace
    Print #nfree, vvFlecheGauche
    Print #nfree, vvMenuContextuel
    Print #nfree, vvWindowsDroit
    Print #nfree, vvAltGr
    Print #nfree, vvControlGauche
    Print #nfree, vvEspace
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13A.txt", vpath & "Leçons\Personnalisé\leçon13A.txt"
End If

' leçon13B
If Dir(vpath & "Leçons\Standard\leçon13B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13B.txt" For Output As #nfree
    Print #nfree, vvRetourArrière
    Print #nfree, vvTab
    Print #nfree, vvRetourArrière
    Print #nfree, vvEspace
    Print #nfree, vvTab
    Print #nfree, vvControlDroit
    Print #nfree, vvTab
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvRetourArrière
    Print #nfree, vvMajGauche
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvTab
    Print #nfree, vvMajDroit
    Print #nfree, vvRetourArrière
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13B.txt", vpath & "Leçons\Personnalisé\leçon13B.txt"
End If

' leçon13C
If Dir(vpath & "Leçons\Standard\leçon13C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13C.txt" For Output As #nfree
    Print #nfree, "F4"
    Print #nfree, "F5"
    Print #nfree, "F2"
    Print #nfree, "F1"
    Print #nfree, "F3"
    Print #nfree, "F1"
    Print #nfree, "F4"
    Print #nfree, "F6"
    Print #nfree, "F8"
    Print #nfree, "F7"
    Print #nfree, "F9"
    Print #nfree, "F10"
    Print #nfree, "F8"
    Print #nfree, "F1"
    Print #nfree, "F4"
    Print #nfree, "F7"
    Print #nfree, "F10"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13C.txt", vpath & "Leçons\Personnalisé\leçon13C.txt"
End If

' leçon13D
If Dir(vpath & "Leçons\Standard\leçon13D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13D.txt" For Output As #nfree
    Print #nfree, "n"
    Print #nfree, "MAJ+N"
    Print #nfree, "CONTROL+N"
    Print #nfree, "ALT+N"
    Print #nfree, "v"
    Print #nfree, "MAJ+V"
    Print #nfree, "CONTROL+V"
    Print #nfree, "ALT+V"
    Print #nfree, "CONTROL+MAJ+N"
    Print #nfree, "CONTROL+ALT+N"
    Print #nfree, "CONTROL+MAJ+S"
    Print #nfree, "CONTROL+ALT+V"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13D.txt", vpath & "Leçons\Personnalisé\leçon13D.txt"
End If

' leçon13E
If Dir(vpath & "Leçons\Standard\leçon13E.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13E.txt" For Output As #nfree
    Print #nfree, "#"
    Print #nfree, "\"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "#"
    Print #nfree, "\"
    Print #nfree, "@"
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "#"
    Print #nfree, "@"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13E.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13E.txt", vpath & "Leçons\Personnalisé\leçon13E.txt"
End If

' leçon13F
If Dir(vpath & "Leçons\Standard\leçon13F.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13F.txt" For Output As #nfree
    Print #nfree, "["
    Print #nfree, "{"
    Print #nfree, "["
    Print #nfree, "{"
    Print #nfree, "]"
    Print #nfree, "}"
    Print #nfree, "]"
    Print #nfree, "}"
    Print #nfree, "["
    Print #nfree, "]"
    Print #nfree, "{"
    Print #nfree, "}"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13F.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13F.txt", vpath & "Leçons\Personnalisé\leçon13F.txt"
End If

' leçon13G
If Dir(vpath & "Leçons\Standard\leçon13G.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon13G.txt" For Output As #nfree
    Print #nfree, "c:\program"
    Print #nfree, "www.meteo.com/demain/tendance.htm"
    Print #nfree, "http://microsoft.com/introduction.htm"
    Print #nfree, "maison-de-la-famille.paris_14@tiscali.fr"
    Print #nfree, "c:\program files\microsoft office"
    Print #nfree, "dir a:\élèves\résultats /p"
    Print #nfree, "path c:\;c:\windows\system"
    Print #nfree, "type toto.txt > tata.txt"
    Print #nfree, "d:\setup.exe"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon13G.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon13G.txt", vpath & "Leçons\Personnalisé\leçon13G.txt"
End If

End Sub
