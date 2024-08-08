Attribute VB_Name = "Module_SetUpVar"

' **************** MAIN *******************************************************************
Public Sub SetUpVariables()

' VARIABLES pour le clavier et la LANGUE
clavierType = "AZERTY français (FRANCE)"
country = "Langue : française"
repjawscountry = "\settings\fra\"

' VARIABLES à TRADUIRE
annul = " L'INSTALLATION d'ApprentiClavier est ANNULÉE. "
bannerversion = "ApprentiClavier Version 1.10"
bannercopyright = "Copyleft 2008-2019 GNU/GPL"
info_lance = CRLF + "  ApprentiClavier a été INSTALLÉ." + CRLF + "  Pour le LANCER, vous utiliserez l'icône " + CRLF + "  ApprentiClavier, " + CRLF + "  qui a été placée sur votre BUREAU Windows. "
msgAide = "Aide" + CRLF + "(F1)"
msgAnnuler = "Annuler" + CRLF + "(Échap)"
msgASupprimer = " Veuillez SUPPRIMER le DOSSIER C:\ApprentiClavier."
msgAttention = "Attention"
msgAurevoir = CRLF4 + "       AU REVOIR       " + CRLF4
msgKeyboard = " Clavier : "
msgContinuer = "&Continuer (Entrée)"
msgDétecté = " détecté. "
msgDésinst = " DÉSINSTALLATION. "
msgDésinstall = "DÉSINSTALLATION d'ApprentiClavier." + CRLF + "Voulez-vous CONTINUER (Oui, Non) ?"
msgDésinstaller = "&Désinstaller"
msgEchap = "   Échap.  "
msgErreur1 = " ERREUR : Ne placez pas l'exécutable d'installation dans c:\ApprentiClavier. "
msgInfo = "Information"
msgInstall = "INSTALLATION de "
msgInstaller = "&Installer"
msgNoFic = "Aucun fichier "
msgNoSono = "On peut utiliser ApprentiClavier avec ou sans vocalisation." + CRLF2 + "Les utilisateurs non-voyants peuvent utiliser les lecteurs d'écran Jaws ou NVDA. Consultez l'aide sur la vocalisation, dans la barre de menus d'ApprentiClavier."
msgPage = "Page "
msgPatientez = " PATIENTEZ. "
msgQuitter = "&Quitter (Échap)"
msgRecommencez = "Veuillez RECOMMENCER."
msgSupprimé = " L'INSTALLATION d'ApprentiClavier a été SUPPRIMÉE. " + CRLF4 + "          AU REVOIR. "
msgVousEtiez = "   Vous étiez dans un message d'explications." + CRLF2 + "Pressez la touche ESPACE pour RÉPÉTER le MESSAGE," + CRLF + " puis les flèches pour RÉPÉTER chaque LIGNE," + CRLF2 + " ou Échap pour SORTIR," + CRLF + " ou Entrée pour CONTINUER."
pressez = CRLF2 + "      Pressez ESPACE pour RÉPÉTER,      " + CRLF + "       ou Entrée pour CONTINUER.        "
texte_bienvenue = "ApprentiClavier sera installé, " + CRLF + " dans le dossier C:\ApprentiClavier." & CRLF & " Pressez la touche " & CRLF & "      I       pour        INSTALLER," & CRLF + "      D       pour     DÉSINSTALLER," & CRLF & "      F1      pour des EXPLICATIONS," & CRLF & " ou Échap     pour          ANNULER."

End Sub
