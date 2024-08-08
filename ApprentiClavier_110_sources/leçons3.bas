Attribute VB_Name = "Module_leçons3"
'******************  RESET_STANDARD : RESET des leçons STANDARD  **************************
Public Sub reset3_standard(force As Byte)
' Création des reps Standard et Personnalisé s'il n'existent pas encore
On Error Resume Next
MkDir vpath & "Leçons"
On Error Resume Next
MkDir vpath & "Leçons\Standard"
On Error Resume Next
MkDir vpath & "Leçons\Personnalisé"
nfree = FreeFile

' leçon14A
If Dir(vpath & "Leçons\Standard\leçon14A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon14A.txt" For Output As #nfree
    'Print #nfree, "Les lettres de Capri, de Mario Soldati."
    'Print #nfree, "Je me suis remise au lit. Sur le dos, les yeux fermés, j'entendais encore les cris des hirondelles. Alors, pour la première fois, j'ai réalisé une chose que je savais depuis longtemps. Les hirondelles rient aussi parce qu'en filant dans l'air très rapidement, le bec ouvert, elles dévorent des insectes. Elles dévorent, elles tuent. J'ai imaginé leur bec ouvert, leurs petits yeux vifs, voraces et rapaces. Et brusquement, ces cris que j'aimais tant m'ont semblé horribles."
    Print #nfree, "Les hommes de bonne volonté, de Jules Romains." 'septembre 2007
    Print #nfree, "Nos rendez-vous, en général, nous nous les donnions de vive voix, d'une fois sur l'autre. Mais comme il pouvait se produire des empêchements, des changements d'heure ou de lieu, nous avions conservé, en le développant, notre système de signaux. Il était devenu d'une grande subtilité, et d'une grande souplesse. Nous étions arrivés à tout exprimer par de petits dessins. Sans user de mots, ni de chiffres. Nos signaux restaient ainsi davantage notre propriété. Ils risquaient moins d'être effacés par une main étrangère. Surtout, personne ne pouvait en soupçonner le sens."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon14A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon14A.txt", vpath & "Leçons\Personnalisé\leçon14A.txt"
End If

' leçon14B
If Dir(vpath & "Leçons\Standard\leçon14B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon14B.txt" For Output As #nfree
    Print #nfree, "Les vagues, de Virginia Woolf."
    Print #nfree, "Voici qu'elle traverse le champ avec un balancement nonchalant de tout le corps, pour nous tromper. Elle arrive dans un creux ; elle se croit invisible. Elle commence à courir, tenant devant elle ses poings fermés. Ses ongles se rejoignent sur son mouchoir roulé en boule. Elle fonce vers le bois de hêtres, hors du grand jour. En entrant dans le bois, elle ouvre les bras, et plonge dans l'ombre comme une nageuse. Mais le grand jour l'avait aveuglée ; elle trébuche ; elle se jette à terre parmi les racines des arbres. Les branches s'inclinent, puis se redressent. Tout ici est plein de trouble et d'agitation. Tout est lugubre. Les racines dessinent à terre une espèce de squelette, et il y a des tas de feuilles mortes dans les coins. Suzanne étale ici sa détresse."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon14B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon14B.txt", vpath & "Leçons\Personnalisé\leçon14B.txt"
End If

' leçon14C
If Dir(vpath & "Leçons\Standard\leçon14C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon14C.txt" For Output As #nfree
    Print #nfree, "Les effets de la crise de 1929."
    Print #nfree, "La récession atteignit l'industrie dès 1930, même dans l'automobile. Le textile fut tout de suite le plus touché. Les prix agricoles s'effondraient : -30% pour le blé, -20% pour le vin. Mais la loi Loucheur d'aide au logement encourageait la construction, et les échanges avec les colonies se développaient. C'est en 1935 que les effets de la crise furent les plus sensibles : on produisait moitié moins d'acier, 2/3 de moins de fer. Le coton était touché à 35%, comme l'automobile. Même le bâtiment reculait. La France avait 400 000 chômeurs, ce qui paraissait énorme."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon14C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon14C.txt", vpath & "Leçons\Personnalisé\leçon14C.txt"
End If

' leçon15A
If Dir(vpath & "Leçons\Standard\leçon15A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon15A.txt" For Output As #nfree
    Print #nfree, "Le paysan parvenu, de Marivaux."
    Print #nfree, "Tu t'expliques plaisamment, me dit-elle ; mais si naïvement que tu plais. Dis-moi, Jacob, que font tes parents à la campagne ? Hélas ! mademoiselle, lui dis-je, ils ne sont pas riches ; mais pour honorables, oh ! c'est la crème de notre paroisse. Pour ce qui est de la profession, mon père est le vigneron et le fermier du seigneur de notre village. Mais je dis mal, je ne sais plus ce qu'il est, il n'y a plus ni vignes, ni ferme. Car notre seigneur est mort. Pour ce qui est de mes autres parents, j'ai deux oncles dont l'un est curé, qui a toujours du bon vin chez lui, et l'autre a pensé l'être plus de trois fois ; mais il va toujours son train de vicaire en attendant mieux."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon15A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon15A.txt", vpath & "Leçons\Personnalisé\leçon15A.txt"
End If

' leçon15B
If Dir(vpath & "Leçons\Standard\leçon15B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon15B.txt" For Output As #nfree
    Print #nfree, "Les pouvoirs de l'empereur Auguste."
    Print #nfree, "Au 1er siècle avant Jésus-Christ (en l'an 43), Octave avait reçu légalement des pouvoirs exorbitants : imperium consulaire, droit de vie et de mort, pouvoirs constituants. En 36, il se voit attribuer des droits tribunitiens qui font de lui un être inviolable et sacro-saint. En 32, spontanément, l'Italie et l'Occident lui prêtent serment de fidélité. En 31, il revêt derechef le consulat. En 29, le titre de prince du Sénat consacre sa place éminente dans l'Etat. Désormais il prend comme prénom Imperator ; les historiens l'appelleront l'empereur Auguste."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon15B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon15B.txt", vpath & "Leçons\Personnalisé\leçon15B.txt"
End If

' leçon15C
If Dir(vpath & "Leçons\Standard\leçon15C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon15C.txt" For Output As #nfree
    Print #nfree, "Correspondance, de Jean-Jacques Rousseau."
    Print #nfree, "J'aurais moins tardé, monsieur, à vous remercier de la dernière lettre dont vous m'avez honoré, si j'avais mesuré ma diligence à répondre sur le plaisir qu'elle m'a fait. Mais, outre qu'il m'en coûte beaucoup d'écrire, j'ai pensé qu'il fallait donner quelques jours aux importunités de ces temps-ci, pour ne pas vous accabler des miennes."
    Print #nfree, "Quoique je ne me console point de ce qui vient de se passer, je suis très content que vous en soyez instruit, puisque cela ne m'a point ôté votre estime ; elle en sera plus à moi quand vous ne me croirez pas meilleur que je ne suis."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon15C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon15C.txt", vpath & "Leçons\Personnalisé\leçon15C.txt"
End If

' leçon16A
If Dir(vpath & "Leçons\Standard\leçon16A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon16A.txt" For Output As #nfree
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
    Print #nfree, "13"   'Attention nombres à 2 chiffres
    Print #nfree, "71"
    Print #nfree, "63"
    Print #nfree, "25"
    Print #nfree, "87"
    Print #nfree, "350"  'Attention nombres à 3 chiffres"
    Print #nfree, "159"
    Print #nfree, "852"
    Print #nfree, "654"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon16A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon16A.txt", vpath & "Leçons\Personnalisé\leçon16A.txt"
    End If

' leçon16B
If Dir(vpath & "Leçons\Standard\leçon16B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon16B.txt" For Output As #nfree
    Print #nfree, vvPlus
    Print #nfree, vvTiret
    Print #nfree, vvBarreOblique
    Print #nfree, vvPoint
    Print #nfree, vvAstérisque
    Print #nfree, vvPlus
    Print #nfree, vvBarreOblique
    Print #nfree, vvTiret
    Print #nfree, vvAstérisque
    Print #nfree, vvPoint
    Print #nfree, vvBarreOblique
    Print #nfree, vvTiret
    Print #nfree, vvAstérisque
    Print #nfree, vvPlus
    Print #nfree, vvPoint
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon16B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon16B.txt", vpath & "Leçons\Personnalisé\leçon16B.txt"
End If

' leçon16C
If Dir(vpath & "Leçons\Standard\leçon16C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon16C.txt" For Output As #nfree
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "~"
    Print #nfree, "€"
    Print #nfree, "#"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "~"
    Print #nfree, "€"
    Print #nfree, "É" 'Suite
    Print #nfree, "œ"
    Print #nfree, "±"
    Print #nfree, "½"
    Print #nfree, "€"
    Print #nfree, "½"
    Print #nfree, "É"
    Print #nfree, "@"
    Print #nfree, "\"
    Print #nfree, "É"
    Print #nfree, "~"
    Print #nfree, "½"
    Print #nfree, "#"
    Print #nfree, "œ"
    Print #nfree, "±"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon16C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon16C.txt", vpath & "Leçons\Personnalisé\leçon16C.txt"
End If

' leçon16D
If Dir(vpath & "Leçons\Standard\leçon16D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon16D.txt" For Output As #nfree
    Print #nfree, vvVerrouillageNumérique
    Print #nfree, vvFlecheGauche
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFlecheDroite
    Print #nfree, vvDébut
    Print #nfree, vvFin
    Print #nfree, vvPageSuivante
    Print #nfree, vvPagePrécédente
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheGauche
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFlecheDroite
    Print #nfree, vvFin
    Print #nfree, vvDébut
    Print #nfree, vvPagePrécédente
    Print #nfree, vvPageSuivante
    Print #nfree, vvInsertion ' Suite
    Print #nfree, vvSuppression
    Print #nfree, vvEntrée
    Print #nfree, vvFlecheGauche
    Print #nfree, vvDébut
    Print #nfree, vvFlecheBas
    Print #nfree, vvFlecheHaut
    Print #nfree, vvFin
    Print #nfree, vvEntrée
    Print #nfree, vvPageSuivante
    Print #nfree, vvInsertion
    Print #nfree, vvFlecheDroite
    Print #nfree, vvSuppression
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon16D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon16D.txt", vpath & "Leçons\Personnalisé\leçon16D.txt"
End If

' leçon17A
If Dir(vpath & "Leçons\Standard\leçon17A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon17A.txt" For Output As #nfree
    Print #nfree, "é"
    Print #nfree, "x"
    Print #nfree, "'"
    Print #nfree, """"
    Print #nfree, "c"
    Print #nfree, "-"
    Print #nfree, "w"
    Print #nfree, "²"
    Print #nfree, "b"
    Print #nfree, "v"
    Print #nfree, "("
    Print #nfree, """"
    Print #nfree, "é"
    Print #nfree, "'"
    Print #nfree, "b"
    Print #nfree, "²"
    Print #nfree, "v"
    Print #nfree, "é"
    Print #nfree, "è"
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
    Print #nfree, "à"
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
    Print #nfree, "°"
    Print #nfree, "1"
    Print #nfree, "x"
    Print #nfree, "/"
    Print #nfree, "\"
    Print #nfree, "ç"
    Print #nfree, "é"
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
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon17A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon17A.txt", vpath & "Leçons\Personnalisé\leçon17A.txt"
End If

' leçon17B
If Dir(vpath & "Leçons\Standard\leçon17B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon17B.txt" For Output As #nfree
    Print #nfree, ";"
    Print #nfree, "C"
    Print #nfree, vvRetourArrière
    Print #nfree, "A"
    Print #nfree, vvDébut
    Print #nfree, "F5"
    Print #nfree, "<"
    Print #nfree, vvEntrée
    Print #nfree, vvAlt
    Print #nfree, ">"
    Print #nfree, "L"
    Print #nfree, vvEspace
    Print #nfree, vvMajGauche
    Print #nfree, vvPageSuivante
    Print #nfree, "x"
    Print #nfree, vvInsertion
    Print #nfree, vvRetourArrière
    Print #nfree, "F4"
    Print #nfree, "_"
    Print #nfree, "ç"
    Print #nfree, """"
    Print #nfree, vvFin
    Print #nfree, "à"
    Print #nfree, vvFlecheDroite
    Print #nfree, ","
    Print #nfree, "W"
    Print #nfree, "µ"
    Print #nfree, vvSuppression
    Print #nfree, vvVerrouillageMajuscules
    Print #nfree, vvMajDroit
    Print #nfree, "é"
    Print #nfree, "!"
    Print #nfree, vvTab
    Print #nfree, vvFlecheHaut
    Print #nfree, "-"
    Print #nfree, vvPageSuivante
    Print #nfree, "f"
    Print #nfree, vvInsertion
    Print #nfree, "F4"
    Print #nfree, "_"
    Print #nfree, vvVerrouillageNumérique
    Print #nfree, "ç"
    Print #nfree, """"
    Print #nfree, vvFin
    Print #nfree, "à"
    Print #nfree, vvFlecheDroite
    Print #nfree, vvSuppression
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon17B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon17B.txt", vpath & "Leçons\Personnalisé\leçon17B.txt"
End If

' leçon17C
If Dir(vpath & "Leçons\Standard\leçon17C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon17C.txt" For Output As #nfree
    Print #nfree, ":"
    Print #nfree, "F"    ' Attention, touches en 2 temps telles que ê inacceptables par rac2ligne1
    Print #nfree, "^"
    Print #nfree, ":"
    Print #nfree, "1"
    Print #nfree, "F9"
    Print #nfree, "*"
    Print #nfree, "²"
    Print #nfree, "¨"
    Print #nfree, "\"
    Print #nfree, vvMajGauche
    Print #nfree, "CONTROL+F8"
    Print #nfree, "B"
    Print #nfree, "F6"
    Print #nfree, "?"
    Print #nfree, "!"
    Print #nfree, "MAJ+F2"
    Print #nfree, "2"
    Print #nfree, vvRetourArrière
    Print #nfree, vvTab
    Print #nfree, vvAlt
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon17C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon17C.txt", vpath & "Leçons\Personnalisé\leçon17C.txt"
End If

' leçon17D
If Dir(vpath & "Leçons\Standard\leçon17D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon17D.txt" For Output As #nfree
    Print #nfree, "1"
    Print #nfree, vvPoint
    Print #nfree, "8"
    Print #nfree, "3"
    Print #nfree, vvAstérisque
    Print #nfree, vvBarreOblique
    Print #nfree, "4"
    Print #nfree, vvPlus
    Print #nfree, "6"
    Print #nfree, vvEntrée
    Print #nfree, "5"
    Print #nfree, "7"
    Print #nfree, "0"
    Print #nfree, vvTiret
    Print #nfree, "9"
    Print #nfree, vvPoint
    Print #nfree, "2"
    Print #nfree, "3"
    Print #nfree, vvAstérisque
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
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon17D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon17D.txt", vpath & "Leçons\Personnalisé\leçon17D.txt"
End If

' leçon18A   Mots accentués
If Dir(vpath & "Leçons\Standard\leçon18A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon18A.txt" For Output As #nfree
    Print #nfree, "abîmé"
    Print #nfree, "aiguë"
    Print #nfree, "appâté"
    Print #nfree, "après-midi"
    Print #nfree, "à l'instant où"
    Print #nfree, "bâté"
    Print #nfree, "bêché"
    Print #nfree, "blâmé"
    Print #nfree, "délégué"
    Print #nfree, "dépêche"
    Print #nfree, "dextérité"
    Print #nfree, "ébréché"
    Print #nfree, "entrecôte"
    Print #nfree, "géomètre"
    Print #nfree, "hébété"
    Print #nfree, "hétérogène"
    Print #nfree, "infâmité"
    Print #nfree, "malhonnêteté"
    Print #nfree, "mûrier"
    Print #nfree, "prélèvement"
    Print #nfree, "rez-de-chaussée"
    Print #nfree, "stoïque"
    Print #nfree, "théâtre"
    Print #nfree, "théière"
    Print #nfree, "zézayer"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon18A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon18A.txt", vpath & "Leçons\Personnalisé\leçon18A.txt"
End If

' leçon18B     Doubles consonnes
If Dir(vpath & "Leçons\Standard\leçon18B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon18B.txt" For Output As #nfree
    Print #nfree, "accessible"
    Print #nfree, "affichette"
    Print #nfree, "atterrissage"
    Print #nfree, "avant-veille"
    Print #nfree, "appropriation"
    Print #nfree, "baïonnette"
    Print #nfree, "démarreur"
    Print #nfree, "ecclésiastique"
    Print #nfree, "embrouillamini"
    Print #nfree, "emmailloter"
    Print #nfree, "excellement"
    Print #nfree, "flottement"
    Print #nfree, "footballeur"
    Print #nfree, "hippopotame"
    Print #nfree, "illégalité"
    Print #nfree, "impressionnant"
    Print #nfree, "insuffisamment"
    Print #nfree, "kidnapper"
    Print #nfree, "laisser-aller"
    Print #nfree, "loggia"
    Print #nfree, "myrrhe"
    Print #nfree, "parallélogramme"
    Print #nfree, "piailler"
    Print #nfree, "quintessence"
    Print #nfree, "rééducation"
    Print #nfree, "toboggan"
    Print #nfree, "vieillissement"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon18B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon18B.txt", vpath & "Leçons\Personnalisé\leçon18B.txt"
End If

' leçon18C   Terminaisons usuelles
If Dir(vpath & "Leçons\Standard\leçon18C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon18C.txt" For Output As #nfree
    Print #nfree, "illusion"
    Print #nfree, "allusion"
    Print #nfree, "occlusion"
    Print #nfree, "épilogue"
    Print #nfree, "psychologue"
    Print #nfree, "dialogue"
    Print #nfree, "catalogue"
    Print #nfree, "liquidation"
    Print #nfree, "acclimatation"
    Print #nfree, "canalisation"
    Print #nfree, "prestidigitation"
    Print #nfree, "précédemment"
    Print #nfree, "dysfonctionnement"
    Print #nfree, "apparemment"
    Print #nfree, "douillettement"
    Print #nfree, "égoïstement"
    Print #nfree, "ballottement"
    Print #nfree, "envoûtement"
    Print #nfree, "identité"
    Print #nfree, "mixité"
    Print #nfree, "excité"
    Print #nfree, "unicité"
    Print #nfree, "chaînette"
    Print #nfree, "dînette"
    Print #nfree, "échantillonner"
    Print #nfree, "papillonner"
    Print #nfree, "carilonner"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon18C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon18C.txt", vpath & "Leçons\Personnalisé\leçon18C.txt"
End If

' leçon18D  Vitesse sur mots semblables
If Dir(vpath & "Leçons\Standard\leçon18D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon18D.txt" For Output As #nfree
    Print #nfree, "psychothérapie"
    Print #nfree, "thalassothérapie"
    Print #nfree, "usure"
    Print #nfree, "luxure"
    Print #nfree, "brûlure"
    Print #nfree, "pâture"
    Print #nfree, "illusoire"
    Print #nfree, "accessoire"
    Print #nfree, "assesseur"
    Print #nfree, "prévoir"
    Print #nfree, "pressoir"
    Print #nfree, "option"
    Print #nfree, "adoption"
    Print #nfree, "exaction"
    Print #nfree, "luxure"
    Print #nfree, "habitude"
    Print #nfree, "attitude"
    Print #nfree, "inaptitude"
    Print #nfree, "cloître"
    Print #nfree, "goître"
    Print #nfree, "croître"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon18D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon18D.txt", vpath & "Leçons\Personnalisé\leçon18D.txt"
End If

' leçon18E    La frappe du programmeur
If Dir(vpath & "Leçons\Standard\leçon18E.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon18E.txt" For Output As #nfree
    Print #nfree, "@echo off"
    Print #nfree, "prompt $P$G"
    Print #nfree, "If leçon.text1.text = ""@"" Then keyexpect = 48"
    Print #nfree, "PATH c:\windows\system32;c:\windows;c:\util"
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon18E.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon18E.txt", vpath & "Leçons\Personnalisé\leçon18E.txt"
End If

' leçon19A
If Dir(vpath & "Leçons\Standard\leçon19A.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon19A.txt" For Output As #nfree
    Print #nfree, "Electre, de Sophocle."
    Print #nfree, "Songe à la gloire que nous aurons, aux éloges qu'on fera de nous ! Citoyens, étrangers, tous diront de nous : "" Regardez ces deux sœurs qui ont sauvé la maison de leur père, qui ont, au mépris de leur vie, affronté des ennemis plus puissants qu'elles et qui se sont vengées de leur dette de sang ! Nous devons les respecter, les admirer, les honorer partout comme on honore les héros ! "". Telle est la gloire qui nous attend, dans notre vie et après notre mort. Ecoute-moi, sœur chérie : sauve ton père, sauve ton frère et avec eux, nos deux bonheurs ! Des cœurs comme les nôtres ne peuvent tolérer de vivre dans la honte."
    Close #nfree
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon19A.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon19A.txt", vpath & "Leçons\Personnalisé\leçon19A.txt"
End If

' leçon19B
If Dir(vpath & "Leçons\Standard\leçon19B.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon19B.txt" For Output As #nfree
    Print #nfree, "Côtes de sanglier à la romaine."
    Print #nfree, ""
    Print #nfree, "Ingrédients :"
    Print #nfree, "6 côtelettes."
    Print #nfree, "6 tranches de pain de mie."
    Print #nfree, "125 grammes de beurre."
    Print #nfree, "1 verre de vin blanc."
    Print #nfree, "1/4 litre de sauce chasseur."
    Print #nfree, "Sel."
    Print #nfree, "Poivre."
    Print #nfree, ""
    Print #nfree, "Préparation 15 mn. Cuisson 20 mn."
    Print #nfree, ""
    Print #nfree, "Parer les côtelettes. Ciseler les bords. Faire macérer 2 jours dans une marinade cuite."
    Print #nfree, ""
    Print #nfree, "Retirer, égoutter, éponger, mettre à la poêle avec 50 grammes de beurre et laisser mijoter 15 minutes. Dresser sur des canapés de la grandeur des côtelettes, frits dans le beurre de chaque côté. Tenir au chaud."
    Print #nfree, ""
    Print #nfree, "Déglacer le jus des côtelettes avec le vin blanc. Ajouter la sauce chasseur sans gelée de groseilles. Assaisonner, faire bouillir 5 minutes à feu vif. Napper les côtelettes et servir."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon19B.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon19B.txt", vpath & "Leçons\Personnalisé\leçon19B.txt"
End If

' leçon19C
If Dir(vpath & "Leçons\Standard\leçon19C.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon19C.txt" For Output As #nfree
    Print #nfree, "A propos du subjonctif imparfait."
    Print #nfree, "Le subjonctif imparfait, après un verbe principal au présent, correspond à un indicatif imparfait dans une proposition indépendante. Fontenelle : "" Vous ne doutez pas que les autres le traitassent de fou "". Gide : "" Il n'est pas un mouvement de ma phrase qui ne répondit pas à un besoin de mon esprit ""."
    Print #nfree, ""
    Print #nfree, "Dans la langue ancienne, l'imparfait du subjonctif exprimait dans la subordonnée une idée d'éventualité voisine du conditionnel. Le français s'est appauvri en perdant cet emploi. Marot : "" Tu n'as couteau, serpe, ni serpillon, qui sût couper corde ni cordillon "". Molière : "" Il n'y a rien au monde que je ne fisse pour votre service "". Rousseau : "" Pensez-vous que je me fisse faute de pleurer, si je pouvais déjeuner de mes larmes ? ""."
    Print #nfree, ""
    Print #nfree, "Mais l'imparfait du subjonctif s'emploie surtout dans les propositions subordonnées régies par un verbe à un temps passé. Molière : "" Ne pouvait-il pas bien attendre qu'il fût jour ""."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon19C.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon19C.txt", vpath & "Leçons\Personnalisé\leçon19C.txt"
End If

' leçon19D
If Dir(vpath & "Leçons\Standard\leçon19D.txt") = "" Or force = 1 Then
    Open vpath & "Leçons\Standard\leçon19D.txt" For Output As #nfree
    Print #nfree, "Les troubles du métabolisme calcique."
    Print #nfree, "L'insuffisance parathyroïdienne provoque une baisse de l'activité des ostéoclastes et une baisse du calcium sanguin et interstitiel. Elle se traduit par un syndrome d'hyperexcitabilité neuro-musculaire où les neurones peuvent décharger spontanément et causer des secousses musculaires, des spasmes et même de la tétanie."
    Print #nfree, ""
    Print #nfree, "Les spasmes du larynx peuvent facilement bloquer la respiration et causer la mort de l'organisme. Le niveau normal de la calcémie est d'environ 10 mg/ml. La tétanie se produit lorsque la concentration tombe à environ 6 mg/ml, soit une baisse de 40%."
    Print #nfree, ""
    Print #nfree, "Les symptômes de l'hypoparathyroïdie disparaissent rapidement après injection de calcium. On traite généralement cet état avec de grandes quantités de vitamine D."
    Close #nfree
    ' Copier la leçon Standard dans Personnalisé seulement si la leçon personnalisée n'existe pas
    If Dir(vpath & "Leçons\Personnalisé\leçon19D.txt") = "" And force = 1 Then FileCopy vpath & "Leçons\Standard\leçon19D.txt", vpath & "Leçons\Personnalisé\leçon19D.txt"
End If

End Sub
