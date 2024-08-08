Attribute VB_Name = "Module_le�ons2"
'******************  RESET_STANDARD : RESET des le�ons STANDARD  **************************
Public Sub reset2_standard(force As Byte)
' Cr�ation des reps Standard et Personnalis� s'il n'existent pas encore
On Error Resume Next
MkDir vpath & "Le�ons"
On Error Resume Next
MkDir vpath & "Le�ons\Standard"
On Error Resume Next
MkDir vpath & "Le�ons\Personnalis�"
nfree = FreeFile

' le�on8A
If Dir(vpath & "Le�ons\Standard\le�on8A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8A.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8A.txt", vpath & "Le�ons\Personnalis�\le�on8A.txt"
End If

' le�on8B
If Dir(vpath & "Le�ons\Standard\le�on8B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8B.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8B.txt", vpath & "Le�ons\Personnalis�\le�on8B.txt"
End If

' le�on8C
If Dir(vpath & "Le�ons\Standard\le�on8C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8C.txt" For Output As #nfree
    Print #nfree, "?"
    Print #nfree, "."
    Print #nfree, "/"
    Print #nfree, "�"
    Print #nfree, "?"
    Print #nfree, "."
    Print #nfree, "?"
    Print #nfree, "/"
    Print #nfree, "."
    Print #nfree, "�"
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8C.txt", vpath & "Le�ons\Personnalis�\le�on8C.txt"
End If

' le�on8D
If Dir(vpath & "Le�ons\Standard\le�on8D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8D.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8D.txt", vpath & "Le�ons\Personnalis�\le�on8D.txt"
End If

' le�on8E
If Dir(vpath & "Le�ons\Standard\le�on8E.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8E.txt" For Output As #nfree
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "Un �ne pour No�l"
    Print #nfree, "Un �ne pour No�l"
    Print #nfree, "Une ou�e fine"
    Print #nfree, "Une ou�e fine"
    Print #nfree, "La bo�te de Mika�l"
    Print #nfree, "La bo�te de Mika�l"
    Print #nfree, "O� d�ner ?"
    Print #nfree, "O� d�ner ?"
    Print #nfree, "Une m�re fra�che"
    Print #nfree, "Une m�re fra�che"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8E.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8E.txt", vpath & "Le�ons\Personnalis�\le�on8E.txt"
End If

' le�on8F
If Dir(vpath & "Le�ons\Standard\le�on8F.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8F.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8F.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8F.txt", vpath & "Le�ons\Personnalis�\le�on8F.txt"
End If

' le�on8G
If Dir(vpath & "Le�ons\Standard\le�on8G.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8G.txt" For Output As #nfree
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "%"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "%"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "%"
    Print #nfree, "�"  ' Attention
    Print #nfree, "�"
    Print #nfree, "*"
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "*"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "*"
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "*"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8G.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8G.txt", vpath & "Le�ons\Personnalis�\le�on8G.txt"
End If

' le�on8H
If Dir(vpath & "Le�ons\Standard\le�on8H.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on8H.txt" For Output As #nfree
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "Un �t� temp�r� vient apr�s un pi�tre hiver."
    Print #nfree, "Un �t� temp�r� vient apr�s un pi�tre hiver."
    Print #nfree, "Il g�re tr�s bien le travail � Paris."
    Print #nfree, "Il g�re tr�s bien le travail � Paris."
    Print #nfree, "Il emm�ne sa fille pr�f�r�e � Moscou."
    Print #nfree, "Il emm�ne sa fille pr�f�r�e � Moscou."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on8H.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on8H.txt", vpath & "Le�ons\Personnalis�\le�on8H.txt"
End If

' le�on9A
If Dir(vpath & "Le�ons\Standard\le�on9A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on9A.txt" For Output As #nfree
    Print #nfree, "&"
    Print #nfree, "�"
    Print #nfree, "&"
    Print #nfree, "�"
    Print #nfree, "%"
    Print #nfree, "&"
    Print #nfree, "�"
    Print #nfree, "*"
    Print #nfree, "&"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "&"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "%"
    Print #nfree, "�"
    Print #nfree, "*"
    Print #nfree, "&"
    Print #nfree, "$"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "&"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on9A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on9A.txt", vpath & "Le�ons\Personnalis�\le�on9A.txt"
End If

' le�on9B
If Dir(vpath & "Le�ons\Standard\le�on9B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on9B.txt" For Output As #nfree
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
    Print #nfree, "C'est-�-dire"
    Print #nfree, "C'est-�-dire"
    Print #nfree, "J'irai, iras-tu aussi ?"
    Print #nfree, "J'irai, iras-tu aussi ?"
    Print #nfree, "Y a-t-il de l'eau, potable bien s�r !"
    Print #nfree, "Y a-t-il de l'eau, potable bien s�r !"
    Print #nfree, "Il dit simplement : ""Viens vite !"""
    Print #nfree, "Il dit simplement : ""Viens vite !"""
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on9B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on9B.txt", vpath & "Le�ons\Personnalis�\le�on9B.txt"
End If

' le�on9C
If Dir(vpath & "Le�ons\Standard\le�on9C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on9C.txt" For Output As #nfree
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
    Print #nfree, "�"
    Print #nfree, ")"
    Print #nfree, "'"
    Print #nfree, "&"
    Print #nfree, "("
    Print #nfree, "_"
    Print #nfree, "-"
    Print #nfree, """"
    Print #nfree, "�"
    Print #nfree, ")"
    Print #nfree, "-"
    Print #nfree, "'"
    Print #nfree, "&"
    Print #nfree, "("
    Print #nfree, """"
    Print #nfree, "�"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on9C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on9C.txt", vpath & "Le�ons\Personnalis�\le�on9C.txt"
End If

' le�on9D
If Dir(vpath & "Le�ons\Standard\le�on9D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on9D.txt" For Output As #nfree
    Print #nfree, "�"
    Print #nfree, "="
    Print #nfree, "a+b"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "="
    Print #nfree, "c+d"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "x+y"
    Print #nfree, "="
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "="
    Print #nfree, "g+h"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, ")"
    Print #nfree, "�"
    Print #nfree, "-"
    Print #nfree, "="
    Print #nfree, "_"
    Print #nfree, "'"
    Print #nfree, "i+j"
    Print #nfree, "�"
    Print #nfree, """"
    Print #nfree, "("
    Print #nfree, ")"
    Print #nfree, "�"
    Print #nfree, "-"
    Print #nfree, "="
    Print #nfree, "_"
    Print #nfree, "'"
    Print #nfree, "k+l"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on9D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on9D.txt", vpath & "Le�ons\Personnalis�\le�on9D.txt"
End If

' le�on10A
If Dir(vpath & "Le�ons\Standard\le�on10A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on10A.txt" For Output As #nfree
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
    Print #nfree, "13"   'Attention nombres � 2 chiffres
    Print #nfree, "24"
    Print #nfree, "41"
    Print #nfree, "43"
    Print #nfree, "32"
    Print #nfree, "13"
    Print #nfree, "24"
    Print #nfree, "41"
    Print #nfree, "42"
    Print #nfree, "32"
    Print #nfree, "123"  'Attention nombres � 3 chiffres"
    Print #nfree, "234"
    Print #nfree, "341"
    Print #nfree, "412"
    Print #nfree, "132"
    Print #nfree, "234"
    Print #nfree, "341"
    Print #nfree, "412"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on10A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on10A.txt", vpath & "Le�ons\Personnalis�\le�on10A.txt"
    End If

' le�on10B
If Dir(vpath & "Le�ons\Standard\le�on10B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on10B.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on10B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on10B.txt", vpath & "Le�ons\Personnalis�\le�on10B.txt"
    End If

' le�on10C
If Dir(vpath & "Le�ons\Standard\le�on10C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on10C.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on10C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on10C.txt", vpath & "Le�ons\Personnalis�\le�on10C.txt"
    End If

' le�on11A
If Dir(vpath & "Le�ons\Standard\le�on11A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on11A.txt" For Output As #nfree
    Print #nfree, "Qui va � la chasse perd sa place"
    Print #nfree, "Qu'il est doux de ne rien faire quand tout s'agite autour de vous"
    Print #nfree, "C'�tait un espagnol de l'arm�e en d�route" '12/2011 ajout
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on11A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on11A.txt", vpath & "Le�ons\Personnalis�\le�on11A.txt"
End If

' le�on11B
If Dir(vpath & "Le�ons\Standard\le�on11B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on11B.txt" For Output As #nfree
    Print #nfree, "No�l au balcon, P�ques aux tisons"
    Print #nfree, "Nous passerons No�l au ch�teau"
    Print #nfree, "Donne-lui tout de m�me � boire, dit mon p�re" '12/2011 ajout
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on11B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on11B.txt", vpath & "Le�ons\Personnalis�\le�on11B.txt"
End If

' le�on11C
If Dir(vpath & "Le�ons\Standard\le�on11C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on11C.txt" For Output As #nfree
    Print #nfree, "Le g�n�ral De Gaulle arriva � la t�te de la France en 1958"
    Print #nfree, "Les r�volutionnaires ont pris la Bastille le 14 juillet 1789"
    Print #nfree, "Il v�cut au Canada et aux Etats-Unis de 1908 � 1921" '12/2011 ajout
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on11C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on11C.txt", vpath & "Le�ons\Personnalis�\le�on11C.txt"
End If

' le�on12A
If Dir(vpath & "Le�ons\Standard\le�on12A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on12A.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on12A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on12A.txt", vpath & "Le�ons\Personnalis�\le�on12A.txt"
End If

' le�on12B
If Dir(vpath & "Le�ons\Standard\le�on12B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on12B.txt" For Output As #nfree
    Print #nfree, vvD�but
    Print #nfree, vvFin
    Print #nfree, vvFin
    Print #nfree, vvD�but
    Print #nfree, vvSuppression  ' Avec les autres touches
    Print #nfree, vvD�but
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, vvSuppression
    Print #nfree, vvD�but
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, vvSuppression
    Print #nfree, vvD�but
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, "q"            ' Avec les lettres
    Print #nfree, vvD�but
    Print #nfree, "m"
    Print #nfree, vvSuppression
    Print #nfree, "f"
    Print #nfree, vvFin
    Print #nfree, "n"
    Print #nfree, vvInsertion
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on12B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on12B.txt", vpath & "Le�ons\Personnalis�\le�on12B.txt"
End If

' le�on12C
If Dir(vpath & "Le�ons\Standard\le�on12C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on12C.txt" For Output As #nfree
    Print #nfree, vvPagePr�c�dente
    Print #nfree, vvPageSuivante
    Print #nfree, vvPagePr�c�dente
    Print #nfree, vvPageSuivante
    Print #nfree, vvFin
    Print #nfree, vvPagePr�c�dente
    Print #nfree, vvPageSuivante
    Print #nfree, vvSuppression
    Print #nfree, vvD�but
    Print #nfree, vvInsertion
    Print #nfree, vvFin
    Print #nfree, vvSuppression
    Print #nfree, vvPagePr�c�dente
    Print #nfree, vvD�but
    Print #nfree, vvInsertion
    Print #nfree, vvPageSuivante
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on12C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on12C.txt", vpath & "Le�ons\Personnalis�\le�on12C.txt"
End If

' le�on12D
If Dir(vpath & "Le�ons\Standard\le�on12D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on12D.txt" For Output As #nfree
    Print #nfree, vvImpression
    Print #nfree, vvPause
    Print #nfree, vvArr�tD�fil
    Print #nfree, vvPause
    Print #nfree, vvImpression
    Print #nfree, vvArr�tD�fil
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on12D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on12D.txt", vpath & "Le�ons\Personnalis�\le�on12D.txt"
End If

' le�on13A
' Attention, Jaws401 perturbe le AltGr, il faut toujours le faire suivre par Control puis Espace
If Dir(vpath & "Le�ons\Standard\le�on13A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13A.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13A.txt", vpath & "Le�ons\Personnalis�\le�on13A.txt"
End If

' le�on13B
If Dir(vpath & "Le�ons\Standard\le�on13B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13B.txt" For Output As #nfree
    Print #nfree, vvRetourArri�re
    Print #nfree, vvTab
    Print #nfree, vvRetourArri�re
    Print #nfree, vvEspace
    Print #nfree, vvTab
    Print #nfree, vvControlDroit
    Print #nfree, vvTab
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvRetourArri�re
    Print #nfree, vvMajGauche
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvTab
    Print #nfree, vvMajDroit
    Print #nfree, vvRetourArri�re
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13B.txt", vpath & "Le�ons\Personnalis�\le�on13B.txt"
End If

' le�on13C
If Dir(vpath & "Le�ons\Standard\le�on13C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13C.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13C.txt", vpath & "Le�ons\Personnalis�\le�on13C.txt"
End If

' le�on13D
If Dir(vpath & "Le�ons\Standard\le�on13D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13D.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13D.txt", vpath & "Le�ons\Personnalis�\le�on13D.txt"
End If

' le�on13E
If Dir(vpath & "Le�ons\Standard\le�on13E.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13E.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13E.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13E.txt", vpath & "Le�ons\Personnalis�\le�on13E.txt"
End If

' le�on13F
If Dir(vpath & "Le�ons\Standard\le�on13F.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13F.txt" For Output As #nfree
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
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13F.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13F.txt", vpath & "Le�ons\Personnalis�\le�on13F.txt"
End If

' le�on13G
If Dir(vpath & "Le�ons\Standard\le�on13G.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on13G.txt" For Output As #nfree
    Print #nfree, "c:\program"
    Print #nfree, "www.meteo.com/demain/tendance.htm"
    Print #nfree, "http://microsoft.com/introduction.htm"
    Print #nfree, "maison-de-la-famille.paris_14@tiscali.fr"
    Print #nfree, "c:\program files\microsoft office"
    Print #nfree, "dir a:\�l�ves\r�sultats /p"
    Print #nfree, "path c:\;c:\windows\system"
    Print #nfree, "type toto.txt > tata.txt"
    Print #nfree, "d:\setup.exe"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on13G.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on13G.txt", vpath & "Le�ons\Personnalis�\le�on13G.txt"
End If

End Sub
