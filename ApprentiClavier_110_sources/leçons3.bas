Attribute VB_Name = "Module_le�ons3"
'******************  RESET_STANDARD : RESET des le�ons STANDARD  **************************
Public Sub reset3_standard(force As Byte)
' Cr�ation des reps Standard et Personnalis� s'il n'existent pas encore
On Error Resume Next
MkDir vpath & "Le�ons"
On Error Resume Next
MkDir vpath & "Le�ons\Standard"
On Error Resume Next
MkDir vpath & "Le�ons\Personnalis�"
nfree = FreeFile

' le�on14A
If Dir(vpath & "Le�ons\Standard\le�on14A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on14A.txt" For Output As #nfree
    'Print #nfree, "Les lettres de Capri, de Mario Soldati."
    'Print #nfree, "Je me suis remise au lit. Sur le dos, les yeux ferm�s, j'entendais encore les cris des hirondelles. Alors, pour la premi�re fois, j'ai r�alis� une chose que je savais depuis longtemps. Les hirondelles rient aussi parce qu'en filant dans l'air tr�s rapidement, le bec ouvert, elles d�vorent des insectes. Elles d�vorent, elles tuent. J'ai imagin� leur bec ouvert, leurs petits yeux vifs, voraces et rapaces. Et brusquement, ces cris que j'aimais tant m'ont sembl� horribles."
    Print #nfree, "Les hommes de bonne volont�, de Jules Romains." 'septembre 2007
    Print #nfree, "Nos rendez-vous, en g�n�ral, nous nous les donnions de vive voix, d'une fois sur l'autre. Mais comme il pouvait se produire des emp�chements, des changements d'heure ou de lieu, nous avions conserv�, en le d�veloppant, notre syst�me de signaux. Il �tait devenu d'une grande subtilit�, et d'une grande souplesse. Nous �tions arriv�s � tout exprimer par de petits dessins. Sans user de mots, ni de chiffres. Nos signaux restaient ainsi davantage notre propri�t�. Ils risquaient moins d'�tre effac�s par une main �trang�re. Surtout, personne ne pouvait en soup�onner le sens."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on14A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on14A.txt", vpath & "Le�ons\Personnalis�\le�on14A.txt"
End If

' le�on14B
If Dir(vpath & "Le�ons\Standard\le�on14B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on14B.txt" For Output As #nfree
    Print #nfree, "Les vagues, de Virginia Woolf."
    Print #nfree, "Voici qu'elle traverse le champ avec un balancement nonchalant de tout le corps, pour nous tromper. Elle arrive dans un creux ; elle se croit invisible. Elle commence � courir, tenant devant elle ses poings ferm�s. Ses ongles se rejoignent sur son mouchoir roul� en boule. Elle fonce vers le bois de h�tres, hors du grand jour. En entrant dans le bois, elle ouvre les bras, et plonge dans l'ombre comme une nageuse. Mais le grand jour l'avait aveugl�e ; elle tr�buche ; elle se jette � terre parmi les racines des arbres. Les branches s'inclinent, puis se redressent. Tout ici est plein de trouble et d'agitation. Tout est lugubre. Les racines dessinent � terre une esp�ce de squelette, et il y a des tas de feuilles mortes dans les coins. Suzanne �tale ici sa d�tresse."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on14B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on14B.txt", vpath & "Le�ons\Personnalis�\le�on14B.txt"
End If

' le�on14C
If Dir(vpath & "Le�ons\Standard\le�on14C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on14C.txt" For Output As #nfree
    Print #nfree, "Les effets de la crise de 1929."
    Print #nfree, "La r�cession atteignit l'industrie d�s 1930, m�me dans l'automobile. Le textile fut tout de suite le plus touch�. Les prix agricoles s'effondraient : -30% pour le bl�, -20% pour le vin. Mais la loi Loucheur d'aide au logement encourageait la construction, et les �changes avec les colonies se d�veloppaient. C'est en 1935 que les effets de la crise furent les plus sensibles : on produisait moiti� moins d'acier, 2/3 de moins de fer. Le coton �tait touch� � 35%, comme l'automobile. M�me le b�timent reculait. La France avait 400 000 ch�meurs, ce qui paraissait �norme."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on14C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on14C.txt", vpath & "Le�ons\Personnalis�\le�on14C.txt"
End If

' le�on15A
If Dir(vpath & "Le�ons\Standard\le�on15A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on15A.txt" For Output As #nfree
    Print #nfree, "Le paysan parvenu, de Marivaux."
    Print #nfree, "Tu t'expliques plaisamment, me dit-elle ; mais si na�vement que tu plais. Dis-moi, Jacob, que font tes parents � la campagne ? H�las ! mademoiselle, lui dis-je, ils ne sont pas riches ; mais pour honorables, oh ! c'est la cr�me de notre paroisse. Pour ce qui est de la profession, mon p�re est le vigneron et le fermier du seigneur de notre village. Mais je dis mal, je ne sais plus ce qu'il est, il n'y a plus ni vignes, ni ferme. Car notre seigneur est mort. Pour ce qui est de mes autres parents, j'ai deux oncles dont l'un est cur�, qui a toujours du bon vin chez lui, et l'autre a pens� l'�tre plus de trois fois ; mais il va toujours son train de vicaire en attendant mieux."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on15A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on15A.txt", vpath & "Le�ons\Personnalis�\le�on15A.txt"
End If

' le�on15B
If Dir(vpath & "Le�ons\Standard\le�on15B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on15B.txt" For Output As #nfree
    Print #nfree, "Les pouvoirs de l'empereur Auguste."
    Print #nfree, "Au 1er si�cle avant J�sus-Christ (en l'an 43), Octave avait re�u l�galement des pouvoirs exorbitants : imperium consulaire, droit de vie et de mort, pouvoirs constituants. En 36, il se voit attribuer des droits tribunitiens qui font de lui un �tre inviolable et sacro-saint. En 32, spontan�ment, l'Italie et l'Occident lui pr�tent serment de fid�lit�. En 31, il rev�t derechef le consulat. En 29, le titre de prince du S�nat consacre sa place �minente dans l'Etat. D�sormais il prend comme pr�nom Imperator ; les historiens l'appelleront l'empereur Auguste."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on15B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on15B.txt", vpath & "Le�ons\Personnalis�\le�on15B.txt"
End If

' le�on15C
If Dir(vpath & "Le�ons\Standard\le�on15C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on15C.txt" For Output As #nfree
    Print #nfree, "Correspondance, de Jean-Jacques Rousseau."
    Print #nfree, "J'aurais moins tard�, monsieur, � vous remercier de la derni�re lettre dont vous m'avez honor�, si j'avais mesur� ma diligence � r�pondre sur le plaisir qu'elle m'a fait. Mais, outre qu'il m'en co�te beaucoup d'�crire, j'ai pens� qu'il fallait donner quelques jours aux importunit�s de ces temps-ci, pour ne pas vous accabler des miennes."
    Print #nfree, "Quoique je ne me console point de ce qui vient de se passer, je suis tr�s content que vous en soyez instruit, puisque cela ne m'a point �t� votre estime ; elle en sera plus � moi quand vous ne me croirez pas meilleur que je ne suis."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on15C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on15C.txt", vpath & "Le�ons\Personnalis�\le�on15C.txt"
End If

' le�on16A
If Dir(vpath & "Le�ons\Standard\le�on16A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on16A.txt" For Output As #nfree
    Print #nfree, "5"
    Print #nfree, "4"
    Print #nfree, "6"
    Print #nfree, "2"
    Print #nfree, "1"
    Print #nfree, "3"
    Print #nfree, "0"
    Print #nfree, "4"
    Print #nfree, "6"
    Print #nfree, "5"
    Print #nfree, "8"
    Print #nfree, "7"
    Print #nfree, "9"
    Print #nfree, "1"
    Print #nfree, "0"
    Print #nfree, "4"
    Print #nfree, "2"
    Print #nfree, "7"
    Print #nfree, "9"
    Print #nfree, "6"
    Print #nfree, "3"
    Print #nfree, "13"   'Attention nombres � 2 chiffres
    Print #nfree, "71"
    Print #nfree, "63"
    Print #nfree, "25"
    Print #nfree, "87"
    Print #nfree, "350"  'Attention nombres � 3 chiffres"
    Print #nfree, "159"
    Print #nfree, "852"
    Print #nfree, "654"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on16A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on16A.txt", vpath & "Le�ons\Personnalis�\le�on16A.txt"
    End If

' le�on16B
If Dir(vpath & "Le�ons\Standard\le�on16B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on16B.txt" For Output As #nfree
    Print #nfree, vvPlus
    Print #nfree, vvTiret
    Print #nfree, vvBarreOblique
    Print #nfree, vvPoint
    Print #nfree, vvAst�risque
    Print #nfree, vvPlus
    Print #nfree, vvBarreOblique
    Print #nfree, vvTiret
    Print #nfree, vvAst�risque
    Print #nfree, vvPoint
    Print #nfree, vvBarreOblique
    Print #nfree, vvTiret
    Print #nfree, vvAst�risque
    Print #nfree, vvPlus
    Print #nfree, vvPoint
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on16B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on16B.txt", vpath & "Le�ons\Personnalis�\le�on16B.txt"
End If

' le�on16C
If Dir(vpath & "Le�ons\Standard\le�on16C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on16C.txt" For Output As #nfree
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "~"
    Print #nfree, "�"
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "~"
    Print #nfree, "�"
    Print #nfree, "�" 'Suite
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "�"
    Print #nfree, "~"
    Print #nfree, "�"
    Print #nfree, "#"
    Print #nfree, "�"
    Print #nfree, "�"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on16C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on16C.txt", vpath & "Le�ons\Personnalis�\le�on16C.txt"
End If

' le�on16D
If Dir(vpath & "Le�ons\Standard\le�on16D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on16D.txt" For Output As #nfree
    Print #nfree, vvVerrouillageNum�rique
    Print #nfree, vvFlecheGauche
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFlecheDroite
    Print #nfree, vvD�but
    Print #nfree, vvFin
    Print #nfree, vvPageSuivante
    Print #nfree, vvPagePr�c�dente
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheGauche
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFlecheDroite
    Print #nfree, vvFin
    Print #nfree, vvD�but
    Print #nfree, vvPagePr�c�dente
    Print #nfree, vvPageSuivante
    Print #nfree, vvInsertion ' Suite
    Print #nfree, vvSuppression
    Print #nfree, vvEntr�e
    Print #nfree, vvFlecheGauche
    Print #nfree, vvD�but
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFin
    Print #nfree, vvEntr�e
    Print #nfree, vvPageSuivante
    Print #nfree, vvInsertion
    Print #nfree, vvFlecheDroite
    Print #nfree, vvSuppression
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on16D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on16D.txt", vpath & "Le�ons\Personnalis�\le�on16D.txt"
End If

' le�on17A
If Dir(vpath & "Le�ons\Standard\le�on17A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on17A.txt" For Output As #nfree
    Print #nfree, "�"
    Print #nfree, "x"
    Print #nfree, "'"
    Print #nfree, """"
    Print #nfree, "c"
    Print #nfree, "-"
    Print #nfree, "w"
    Print #nfree, "�"
    Print #nfree, "b"
    Print #nfree, "v"
    Print #nfree, "("
    Print #nfree, """"
    Print #nfree, "�"
    Print #nfree, "'"
    Print #nfree, "b"
    Print #nfree, "�"
    Print #nfree, "v"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "c"
    Print #nfree, "&"
    Print #nfree, "x"
    Print #nfree, "q"
    Print #nfree, "s"
    Print #nfree, "z"
    Print #nfree, "t"
    Print #nfree, "w"
    Print #nfree, "-"
    Print #nfree, "M"
    Print #nfree, "�"
    Print #nfree, "+"
    Print #nfree, "="
    Print #nfree, ";"
    Print #nfree, ">"
    Print #nfree, "9"
    Print #nfree, "*"
    Print #nfree, "$"
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, "0"
    Print #nfree, "�"
    Print #nfree, "1"
    Print #nfree, "x"
    Print #nfree, "/"
    Print #nfree, "\"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "8"
    Print #nfree, "&"
    Print #nfree, "%"
    Print #nfree, "C"
    Print #nfree, "$"
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, ")"
    Print #nfree, "/"
    Print #nfree, "\"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on17A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on17A.txt", vpath & "Le�ons\Personnalis�\le�on17A.txt"
End If

' le�on17B
If Dir(vpath & "Le�ons\Standard\le�on17B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on17B.txt" For Output As #nfree
    Print #nfree, ";"
    Print #nfree, "C"
    Print #nfree, vvRetourArri�re
    Print #nfree, "A"
    Print #nfree, vvD�but
    Print #nfree, "F5"
    Print #nfree, "<"
    Print #nfree, vvEntr�e
    Print #nfree, vvAlt
    Print #nfree, ">"
    Print #nfree, "L"
    Print #nfree, vvEspace
    Print #nfree, vvMajGauche
    Print #nfree, vvPageSuivante
    Print #nfree, "x"
    Print #nfree, vvInsertion
    Print #nfree, vvRetourArri�re
    Print #nfree, "F4"
    Print #nfree, "_"
    Print #nfree, "�"
    Print #nfree, """"
    Print #nfree, vvFin
    Print #nfree, "�"
    Print #nfree, vvFlecheDroite
    Print #nfree, ","
    Print #nfree, "W"
    Print #nfree, "�"
    Print #nfree, vvSuppression
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajDroit
    Print #nfree, "�"
    Print #nfree, "!"
    Print #nfree, vvTab
    Print #nfree, vvFlecheHaut
    Print #nfree, "-"
    Print #nfree, vvPageSuivante
    Print #nfree, "f"
    Print #nfree, vvInsertion
    Print #nfree, "F4"
    Print #nfree, "_"
    Print #nfree, vvVerrouillageNum�rique
    Print #nfree, "�"
    Print #nfree, """"
    Print #nfree, vvFin
    Print #nfree, "�"
    Print #nfree, vvFlecheDroite
    Print #nfree, vvSuppression
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on17B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on17B.txt", vpath & "Le�ons\Personnalis�\le�on17B.txt"
End If

' le�on17C
If Dir(vpath & "Le�ons\Standard\le�on17C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on17C.txt" For Output As #nfree
    Print #nfree, ":"
    Print #nfree, "F"    ' Attention, touches en 2 temps telles que � inacceptables par rac2ligne1
    Print #nfree, "^"
    Print #nfree, ":"
    Print #nfree, "1"
    Print #nfree, "F9"
    Print #nfree, "*"
    Print #nfree, "�"
    Print #nfree, "�"
    Print #nfree, "\"
    Print #nfree, vvMajGauche
    Print #nfree, "CONTROL+F8"
    Print #nfree, "B"
    Print #nfree, "F6"
    Print #nfree, "?"
    Print #nfree, "!"
    Print #nfree, "MAJ+F2"
    Print #nfree, "2"
    Print #nfree, vvRetourArri�re
    Print #nfree, vvTab
    Print #nfree, vvAlt
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on17C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on17C.txt", vpath & "Le�ons\Personnalis�\le�on17C.txt"
End If

' le�on17D
If Dir(vpath & "Le�ons\Standard\le�on17D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on17D.txt" For Output As #nfree
    Print #nfree, "1"
    Print #nfree, vvPoint
    Print #nfree, "8"
    Print #nfree, "3"
    Print #nfree, vvAst�risque
    Print #nfree, vvBarreOblique
    Print #nfree, "4"
    Print #nfree, vvPlus
    Print #nfree, "6"
    Print #nfree, vvEntr�e
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "0"
    Print #nfree, vvTiret
    Print #nfree, "9"
    Print #nfree, vvPoint
    Print #nfree, "2"
    Print #nfree, "3"
    Print #nfree, vvAst�risque
    Print #nfree, vvBarreOblique
    Print #nfree, "4"
    Print #nfree, vvPlus
    Print #nfree, "6"
    Print #nfree, vvTiret
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "0"
    Print #nfree, vvTiret
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on17D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on17D.txt", vpath & "Le�ons\Personnalis�\le�on17D.txt"
End If

' le�on18A   Mots accentu�s
If Dir(vpath & "Le�ons\Standard\le�on18A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on18A.txt" For Output As #nfree
    Print #nfree, "ab�m�"
    Print #nfree, "aigu�"
    Print #nfree, "app�t�"
    Print #nfree, "apr�s-midi"
    Print #nfree, "� l'instant o�"
    Print #nfree, "b�t�"
    Print #nfree, "b�ch�"
    Print #nfree, "bl�m�"
    Print #nfree, "d�l�gu�"
    Print #nfree, "d�p�che"
    Print #nfree, "dext�rit�"
    Print #nfree, "�br�ch�"
    Print #nfree, "entrec�te"
    Print #nfree, "g�om�tre"
    Print #nfree, "h�b�t�"
    Print #nfree, "h�t�rog�ne"
    Print #nfree, "inf�mit�"
    Print #nfree, "malhonn�tet�"
    Print #nfree, "m�rier"
    Print #nfree, "pr�l�vement"
    Print #nfree, "rez-de-chauss�e"
    Print #nfree, "sto�que"
    Print #nfree, "th��tre"
    Print #nfree, "th�i�re"
    Print #nfree, "z�zayer"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on18A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on18A.txt", vpath & "Le�ons\Personnalis�\le�on18A.txt"
End If

' le�on18B     Doubles consonnes
If Dir(vpath & "Le�ons\Standard\le�on18B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on18B.txt" For Output As #nfree
    Print #nfree, "accessible"
    Print #nfree, "affichette"
    Print #nfree, "atterrissage"
    Print #nfree, "avant-veille"
    Print #nfree, "appropriation"
    Print #nfree, "ba�onnette"
    Print #nfree, "d�marreur"
    Print #nfree, "eccl�siastique"
    Print #nfree, "embrouillamini"
    Print #nfree, "emmailloter"
    Print #nfree, "excellement"
    Print #nfree, "flottement"
    Print #nfree, "footballeur"
    Print #nfree, "hippopotame"
    Print #nfree, "ill�galit�"
    Print #nfree, "impressionnant"
    Print #nfree, "insuffisamment"
    Print #nfree, "kidnapper"
    Print #nfree, "laisser-aller"
    Print #nfree, "loggia"
    Print #nfree, "myrrhe"
    Print #nfree, "parall�logramme"
    Print #nfree, "piailler"
    Print #nfree, "quintessence"
    Print #nfree, "r��ducation"
    Print #nfree, "toboggan"
    Print #nfree, "vieillissement"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on18B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on18B.txt", vpath & "Le�ons\Personnalis�\le�on18B.txt"
End If

' le�on18C   Terminaisons usuelles
If Dir(vpath & "Le�ons\Standard\le�on18C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on18C.txt" For Output As #nfree
    Print #nfree, "illusion"
    Print #nfree, "allusion"
    Print #nfree, "occlusion"
    Print #nfree, "�pilogue"
    Print #nfree, "psychologue"
    Print #nfree, "dialogue"
    Print #nfree, "catalogue"
    Print #nfree, "liquidation"
    Print #nfree, "acclimatation"
    Print #nfree, "canalisation"
    Print #nfree, "prestidigitation"
    Print #nfree, "pr�c�demment"
    Print #nfree, "dysfonctionnement"
    Print #nfree, "apparemment"
    Print #nfree, "douillettement"
    Print #nfree, "�go�stement"
    Print #nfree, "ballottement"
    Print #nfree, "envo�tement"
    Print #nfree, "identit�"
    Print #nfree, "mixit�"
    Print #nfree, "excit�"
    Print #nfree, "unicit�"
    Print #nfree, "cha�nette"
    Print #nfree, "d�nette"
    Print #nfree, "�chantillonner"
    Print #nfree, "papillonner"
    Print #nfree, "carilonner"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on18C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on18C.txt", vpath & "Le�ons\Personnalis�\le�on18C.txt"
End If

' le�on18D  Vitesse sur mots semblables
If Dir(vpath & "Le�ons\Standard\le�on18D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on18D.txt" For Output As #nfree
    Print #nfree, "psychoth�rapie"
    Print #nfree, "thalassoth�rapie"
    Print #nfree, "usure"
    Print #nfree, "luxure"
    Print #nfree, "br�lure"
    Print #nfree, "p�ture"
    Print #nfree, "illusoire"
    Print #nfree, "accessoire"
    Print #nfree, "assesseur"
    Print #nfree, "pr�voir"
    Print #nfree, "pressoir"
    Print #nfree, "option"
    Print #nfree, "adoption"
    Print #nfree, "exaction"
    Print #nfree, "luxure"
    Print #nfree, "habitude"
    Print #nfree, "attitude"
    Print #nfree, "inaptitude"
    Print #nfree, "clo�tre"
    Print #nfree, "go�tre"
    Print #nfree, "cro�tre"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on18D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on18D.txt", vpath & "Le�ons\Personnalis�\le�on18D.txt"
End If

' le�on18E    La frappe du programmeur
If Dir(vpath & "Le�ons\Standard\le�on18E.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on18E.txt" For Output As #nfree
    Print #nfree, "@echo off"
    Print #nfree, "prompt $P$G"
    Print #nfree, "If le�on.text1.text = ""@"" Then keyexpect = 48"
    Print #nfree, "PATH c:\windows\system32;c:\windows;c:\util"
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on18E.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on18E.txt", vpath & "Le�ons\Personnalis�\le�on18E.txt"
End If

' le�on19A
If Dir(vpath & "Le�ons\Standard\le�on19A.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on19A.txt" For Output As #nfree
    Print #nfree, "Electre, de Sophocle."
    Print #nfree, "Songe � la gloire que nous aurons, aux �loges qu'on fera de nous ! Citoyens, �trangers, tous diront de nous : "" Regardez ces deux s�urs qui ont sauv� la maison de leur p�re, qui ont, au m�pris de leur vie, affront� des ennemis plus puissants qu'elles et qui se sont veng�es de leur dette de sang ! Nous devons les respecter, les admirer, les honorer partout comme on honore les h�ros ! "". Telle est la gloire qui nous attend, dans notre vie et apr�s notre mort. Ecoute-moi, s�ur ch�rie : sauve ton p�re, sauve ton fr�re et avec eux, nos deux bonheurs ! Des c�urs comme les n�tres ne peuvent tol�rer de vivre dans la honte."
    Close #nfree
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on19A.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on19A.txt", vpath & "Le�ons\Personnalis�\le�on19A.txt"
End If

' le�on19B
If Dir(vpath & "Le�ons\Standard\le�on19B.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on19B.txt" For Output As #nfree
    Print #nfree, "C�tes de sanglier � la romaine."
    Print #nfree, ""
    Print #nfree, "Ingr�dients :"
    Print #nfree, "6 c�telettes."
    Print #nfree, "6 tranches de pain de mie."
    Print #nfree, "125 grammes de beurre."
    Print #nfree, "1 verre de vin blanc."
    Print #nfree, "1/4 litre de sauce chasseur."
    Print #nfree, "Sel."
    Print #nfree, "Poivre."
    Print #nfree, ""
    Print #nfree, "Pr�paration 15 mn. Cuisson 20 mn."
    Print #nfree, ""
    Print #nfree, "Parer les c�telettes. Ciseler les bords. Faire mac�rer 2 jours dans une marinade cuite."
    Print #nfree, ""
    Print #nfree, "Retirer, �goutter, �ponger, mettre � la po�le avec 50 grammes de beurre et laisser mijoter 15 minutes. Dresser sur des canap�s de la grandeur des c�telettes, frits dans le beurre de chaque c�t�. Tenir au chaud."
    Print #nfree, ""
    Print #nfree, "D�glacer le jus des c�telettes avec le vin blanc. Ajouter la sauce chasseur sans gel�e de groseilles. Assaisonner, faire bouillir 5 minutes � feu vif. Napper les c�telettes et servir."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on19B.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on19B.txt", vpath & "Le�ons\Personnalis�\le�on19B.txt"
End If

' le�on19C
If Dir(vpath & "Le�ons\Standard\le�on19C.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on19C.txt" For Output As #nfree
    Print #nfree, "A propos du subjonctif imparfait."
    Print #nfree, "Le subjonctif imparfait, apr�s un verbe principal au pr�sent, correspond � un indicatif imparfait dans une proposition ind�pendante. Fontenelle : "" Vous ne doutez pas que les autres le traitassent de fou "". Gide : "" Il n'est pas un mouvement de ma phrase qui ne r�pondit pas � un besoin de mon esprit ""."
    Print #nfree, ""
    Print #nfree, "Dans la langue ancienne, l'imparfait du subjonctif exprimait dans la subordonn�e une id�e d'�ventualit� voisine du conditionnel. Le fran�ais s'est appauvri en perdant cet emploi. Marot : "" Tu n'as couteau, serpe, ni serpillon, qui s�t couper corde ni cordillon "". Moli�re : "" Il n'y a rien au monde que je ne fisse pour votre service "". Rousseau : "" Pensez-vous que je me fisse faute de pleurer, si je pouvais d�jeuner de mes larmes ? ""."
    Print #nfree, ""
    Print #nfree, "Mais l'imparfait du subjonctif s'emploie surtout dans les propositions subordonn�es r�gies par un verbe � un temps pass�. Moli�re : "" Ne pouvait-il pas bien attendre qu'il f�t jour ""."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on19C.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on19C.txt", vpath & "Le�ons\Personnalis�\le�on19C.txt"
End If

' le�on19D
If Dir(vpath & "Le�ons\Standard\le�on19D.txt") = "" Or force = 1 Then
    Open vpath & "Le�ons\Standard\le�on19D.txt" For Output As #nfree
    Print #nfree, "Les troubles du m�tabolisme calcique."
    Print #nfree, "L'insuffisance parathyro�dienne provoque une baisse de l'activit� des ost�oclastes et une baisse du calcium sanguin et interstitiel. Elle se traduit par un syndrome d'hyperexcitabilit� neuro-musculaire o� les neurones peuvent d�charger spontan�ment et causer des secousses musculaires, des spasmes et m�me de la t�tanie."
    Print #nfree, ""
    Print #nfree, "Les spasmes du larynx peuvent facilement bloquer la respiration et causer la mort de l'organisme. Le niveau normal de la calc�mie est d'environ 10 mg/ml. La t�tanie se produit lorsque la concentration tombe � environ 6 mg/ml, soit une baisse de 40%."
    Print #nfree, ""
    Print #nfree, "Les sympt�mes de l'hypoparathyro�die disparaissent rapidement apr�s injection de calcium. On traite g�n�ralement cet �tat avec de grandes quantit�s de vitamine D."
    Close #nfree
    ' Copier la le�on Standard dans Personnalis� seulement si la le�on personnalis�e n'existe pas
    If Dir(vpath & "Le�ons\Personnalis�\le�on19D.txt") = "" And force = 1 Then FileCopy vpath & "Le�ons\Standard\le�on19D.txt", vpath & "Le�ons\Personnalis�\le�on19D.txt"
End If

End Sub
