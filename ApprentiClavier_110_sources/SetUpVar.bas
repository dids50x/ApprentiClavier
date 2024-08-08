Attribute VB_Name = "Module_SetUpVar"

' **************** MAIN *******************************************************************
Public Sub SetUpVariables()

' VARIABLES pour le clavier et la LANGUE
clavierType = "AZERTY fran�ais (FRANCE)"
country = "Langue : fran�aise"
repjawscountry = "\settings\fra\"

' VARIABLES � TRADUIRE
annul = " L'INSTALLATION d'ApprentiClavier est ANNUL�E. "
bannerversion = "ApprentiClavier Version 1.10"
bannercopyright = "Copyleft 2008-2019 GNU/GPL"
info_lance = CRLF + "  ApprentiClavier a �t� INSTALL�." + CRLF + "  Pour le LANCER, vous utiliserez l'ic�ne " + CRLF + "  ApprentiClavier, " + CRLF + "  qui a �t� plac�e sur votre BUREAU Windows. "
msgAide = "Aide" + CRLF + "(F1)"
msgAnnuler = "Annuler" + CRLF + "(�chap)"
msgASupprimer = " Veuillez SUPPRIMER le DOSSIER C:\ApprentiClavier."
msgAttention = "Attention"
msgAurevoir = CRLF4 + "       AU REVOIR       " + CRLF4
msgKeyboard = " Clavier : "
msgContinuer = "&Continuer (Entr�e)"
msgD�tect� = " d�tect�. "
msgD�sinst = " D�SINSTALLATION. "
msgD�sinstall = "D�SINSTALLATION d'ApprentiClavier." + CRLF + "Voulez-vous CONTINUER (Oui, Non) ?"
msgD�sinstaller = "&D�sinstaller"
msgEchap = "   �chap.  "
msgErreur1 = " ERREUR : Ne placez pas l'ex�cutable d'installation dans c:\ApprentiClavier. "
msgInfo = "Information"
msgInstall = "INSTALLATION de "
msgInstaller = "&Installer"
msgNoFic = "Aucun fichier "
msgNoSono = "On peut utiliser ApprentiClavier avec ou sans vocalisation." + CRLF2 + "Les utilisateurs non-voyants peuvent utiliser les lecteurs d'�cran Jaws ou NVDA. Consultez l'aide sur la vocalisation, dans la barre de menus d'ApprentiClavier."
msgPage = "Page "
msgPatientez = " PATIENTEZ. "
msgQuitter = "&Quitter (�chap)"
msgRecommencez = "Veuillez RECOMMENCER."
msgSupprim� = " L'INSTALLATION d'ApprentiClavier a �t� SUPPRIM�E. " + CRLF4 + "          AU REVOIR. "
msgVousEtiez = "   Vous �tiez dans un message d'explications." + CRLF2 + "Pressez la touche ESPACE pour R�P�TER le MESSAGE," + CRLF + " puis les fl�ches pour R�P�TER chaque LIGNE," + CRLF2 + " ou �chap pour SORTIR," + CRLF + " ou Entr�e pour CONTINUER."
pressez = CRLF2 + "      Pressez ESPACE pour R�P�TER,      " + CRLF + "       ou Entr�e pour CONTINUER.        "
texte_bienvenue = "ApprentiClavier sera install�, " + CRLF + " dans le dossier C:\ApprentiClavier." & CRLF & " Pressez la touche " & CRLF & "      I       pour        INSTALLER," & CRLF + "      D       pour     D�SINSTALLER," & CRLF & "      F1      pour des EXPLICATIONS," & CRLF & " ou �chap     pour          ANNULER."

End Sub
