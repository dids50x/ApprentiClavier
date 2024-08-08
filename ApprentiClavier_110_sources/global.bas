Attribute VB_Name = "Module_global"
'Ce logiciel libre est disponible sous licence GNU/GPL,
'dont une copie se trouvera dans le fichier gpl.txt,
'avec une traduction française non officielle gpl-fr.txt.

Option Explicit
' ***************  CONTIENT DECLARATIONS, MAIN, puis ROUTINES à TRADUIRE *********************
' API declarations:
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
' Ajout avril 2008 pour les beep sonores
Public Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' TYPES, notamment OS Version Win95/98 ou XP
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128   '  Maintenance string for PSS usage
End Type

' VARIABLES
Global BLANCS6, BLANCS12 As String '6 espaces, 12 espaces
Global CRLF, CRLF2, CRLF3 As String 'carrierreturn-linefeed (13-10) une, deux, trois fois
Global cara1, cara2, letter, old, tempo, clavierType As String 'variables
Global nivo, nivoRep, nom, nom_temp, vpath, vfile, vfileresults As String 'niveau_leçons, niveau_nom_répertoire_immuable, nom_utilisateur, nom_temporaire, chemin_programme, chemin_avec_nom_fichier, chemin_avec_résultats
Global bannerThanks, bannerNosell, bannerFunction, bannerVersion, bannerCopyright, bannerAuthorAddress As String  'remerciements, non_vendu, fonction_du_logiciel, bannière_app_version, bannière_copyright, bannière_auteur
Global bannerPrincipal, bannerMenu, bannerLeçon As String  'mot_principal, mot_menu, mot_leçon

Global biplevel, debexplilevel, debexplivalue, msgtext0 As String 'niveau_sonore_des_bips, débit_des_explications, string_repère_débit, texte_msg_msgform
Global msgtext1(150), msgtext2(150), datatext1(150) As String 'texte_msg_text1(iter), texte_msg_boite_de_dialogue(iter), données_text1(ligne)
Global repj(9), repjexe(9) As String 'rep_Config_jaws_si_plusieurs_versions_jaws, rep_Exe_Jaws
Global fautesur(150), fautecourante, fauteprec As String 'cara_demandé_fauté(iwrong)
Global currentline, currentmenuline, scorecourant As String 'ligne_complète_en_cours_incluant_retour-chariot, ligne_courante_menu, score_courant
Global country, repjawscountry, repjaws, repUsers, repNVDA, repjawsjsb, ujaws(9) As String 'langue, rep_jaws_settings, test version Windows, rép_jaws_pour_jcf, idem_valid, idem_rép_name, idem_fichier_jss, unité_rép_jaws
Global perso_methode, pressez, pressez_entrée, pressez_quit, pressez_suivant, pressez_LeçonSuivante, pressez_précédent, pressez_F2, pressez_touche As String 'msg_methode_pour_créer_perso_files, msg_pressez_espace_ou_Entrée, msg_pressez_entrée, msg_pressez_quit, msg_pressez_suivant, msg_pressez_précédent, msg_si_F2_inutile, msg_minimum
Global pressez_basic, pressez_ligne As String 'msg_pressez_de_base, msg_pressez_ligne_par_ligne
Global repjawsnames, repjawsfra As String 'nom_reps_jaws_type_c:\jaws401, nom_rep_settings_fra_courant
Global exo_courant As String 'exo_courant.txt (attention pas d'autre variable sur cette ligne déclaration)
Global debgenlevel, debgenvalue As String 'débit_reperé_par_string, débit_session_info

Global msgAide, msgAideF3, msgAtt, msgWith, msgBienvenue, msgEnter, msgLevel, msgNofic, msgSpeedExp, msgSpeedGen, msgUser As String
Global msgBienvFaudra, msgPage, msgPressEnter, msgSonori, msgTapez, msgTapez2, msgTapezTouche As String
Global msgBienvUsername, msgBienvRedo, msgBienvRep, msgBienvRepeat, msgBienvRetape, msgRelaunch As String
Global msgMots, msgCommandes, msgCommandesEn, msgExoFautes, msgExoIdem, msgExoSuivant, msgLevelIs, msgStandard, msgPersonnalisé As String
Global msgPrincPour, msgPrincDansniveau, msgPrincContenu, msgPrincTerminé As String
Global msgFormPressez, msgFormRecommencer, msgFormVousétiez As String
Global msgAvec, msgMotsEn, msgSecondes, msgFautesSur, msgPressezF1, msgRéussià, msgRéussià100, msgPourcent, msgPourcentSeulement, msgTranslator As String
Global msgTypeClavier, msgScore, msgCommandesDispo, msgF1Aide, msgF2Loc, msgF3AM, msgEspace, msgCtrlEspace, msgAltEspace, msgMajEspace, msgSortir, msgSortir2, msgSortir3, msgAltF4 As String
Global msgAurevoir, msgNoficSono, msgEntréeContinuer, msgDétecté, msgUserIs, msgSpeedExpIs, msgSpeedGenIs, msgBipsAre, msgChoisir, msgChoisissez, msgOptions As String
Global msgBip, msgBipComment 'avril 2008
Global msgÉchap, msgÉchap2, msgÉchap3, msgContinuer, msgPrécédent, msgSuivant, msgQuitter, msgQuitterAM, msgQuitterMP, msgQuitterVers As String
Global msgF1F2F3, msgLaLeçon, msgEstTerminée, msgRésultats, msgMotsMinute, msgSes, msgRéussite As String
Global msgClassique, msgDifférent, msgDictée, msgDébit, msgExpli, msgLent, msgMoyen, msgVite, msgNormal, msgRapide As String
Global msgNiveauStandard, msgNiveauPersonnalisé, msgConseils, msgHit, msgNoSono, msgKeyboard As String
Global msgRestartTitle, msgRestart, msgRestartCmd, meRestart, msgResetTitle, msgReset, meReset As String

Global vvAccentGrave, vvAccoladeDroite, vvAccoladeGauche, vvAlaligne, vvAlt, vvAltGr, vvAltOuAltGr, vvApostrophe, vvArrêtDéfil, vvAstérisque As String
Global vvBarreOblique, vvBarreObliqueInversée, vvBarreVerticale, vvControl, vvControlDroit, vvControlGauche, vvCrochetDroit, vvCrochetGauche, vvCtrl, vvDébut, vvDeuxPoints, vvDiviser As String
Global vvEchap, vvÉchap, vvEntrée, vvEspace, vvÉtoile, vvFin, vvFlecheBas, vvFlecheDroite, vvFlecheGauche, vvFlecheHaut, vvGuillemet, vvImpression, vvInférieur, vvInsertion As String
Global vvMaj, vvMajDroit, vvMajGauche, vvMajuscule, vvMenuContextuel, vvMinuscule, vvMoins, vvMultiplier As String
Global vvPagePrécédente, vvPageSuivante, vvParenthèseDroite, vvParenthèseGauche, vvPause, vvPlus, vvPoint, vvPointExclamation, vvPointInterrogation, vvPointVirgule As String
Global vvRetourArriere, vvRetourArrière, vvSansNom, vvSouligné, vvSupérieur, vvSuppression, vvTab, vvTabulationAvant, vvTabulationArrière, vvTilde, vvTiret As String
Global vvVerrouillageMajuscules, vvVerrouillageNumérique, vvVirgule, vvWindowsDroit, vvWindowsGauche As String

Global nboccur(150) As Integer 'nb_occurences_sur_cara_demandé_fauté(iwrong)
Global rrs(9), rrt(9) As Integer 'variables_selon_le_type_de_version_Jaws
Global ii, iiold, ll, llold, nbli, zz As Integer 'indice_courant_text2, old_ii, length_ligne_text1, old_length, nb_lignes_text1, length_currentline
Global iistart, iistartp, iistop, iistop0, iistop1, iistop2, iistop3, iistop9, iistopf As Integer 'indices_déb_phrase, indices_fin_phrase (détecte le "." ou "!" ou "?")
Global jj, kk, mm, nn, pp, qq As Integer 'variables_de_boucle
Global iwrong, iwrongbis, iwrongbismax, iwrongl, irecur As Integer 'nb_fautes, nb_fautes_bissées, nb_max, nb_fautes_dans_ligne, nb_fois_qu'on_recommence
Global iwrongCR, iwrongCRmax As Integer 'nb_fautes_pour_retour_charriot_à_la_ligne
Global nbcaras, nbonscaras, lt1 As Integer 'nb_caras_exo, nbonscaras, length_réelle_text1
Global iter, iiante, iiprec As Integer 'indice_avancement, indice_antérieur, indice_précédent
Global cadencecara, cadencemot, cadenceligne As Integer 'cadences_demande_suivante
Global elapsed, elapsedtot As Integer 'temps_passé_pour_ligne_courante, temps_total
Global nbmots, keyforce As Integer 'nombre_mots, force_keycode
Global echapbis, echapbismax, echapoff As Integer 'echap_bissée, max_de_echapbis, echap_offset
Global fsize, fsizedefault, leçonfontsize, leçonfontsize5 As Integer 'font_size, font_size_default_msgform, fontsize_normale_pour_leçon, idem_leçon5
Global numpad As Integer 'pavé_numérique:-1=évite_lock__0=non__1=uniquement__2=oui_ainsi_que_chiffres_clavier_principal
Global firstmove As Integer 'first_cursor_move_in_msgform

Global indif, KeyAscii, Keycode, KeyExpect, ShiftExpect As Byte 'indifférent_majuscule_minuscule, réponse_KeyPress, réponse_KeyUp_or_down, réponse_attendue_code_touche, réponse_attendue_shift_touche
Global KeyFirst, KeySecond, KeyThird As Byte 'Codes successifs générs par Jaws pour les touches Echap, Alt, Control
Global derligne, nextleçon, typeleçon, stopscroll As Byte 'dernière_ligne, passer_à_leçon_suivante, typeleçon_1_2_3_7_14, stop_scroll_results
Global f2link As Byte 'Enchaine_après_f2
Global keyinhibit, t2inhibit, mcinhibit, fullscreeninhibit As Byte 'inhibit_keyup_after_msgbox, inhibit_event_change_text2, inhibit_after_menu-contextuel, inhibit_fullscreen_display
Global concatf, timevalid, winstop As Byte 'mode_où_text1_s'ajoute, autorise_affich&calcul_vitesse, renforce_stop_touche_windows
Global notab, sonocara, timein, timeover, timeout As Byte 'touche_tab_sans_effet, sel_cara_pour_sonoriser, compteur_temps, temps_de_réponse_dépassé, timeout=1_pour_quitter
Global menucount, numleçon, numexo, nbexo, nblines As Byte 'nb_choix_menu_courant, numéro_leçon, numéro_exercice, nb_d'exercices_pour_la_leçon, nb_lignes_leçon
Global f1msgform, inexo, iwait, ff, incomplet As Byte 'F1_valid_msgform, entrée_leçon, var_wait, flag_faute, leçon_où_manque_au_moins_un_résultat_pctok_d'un_exercice
Global numindex, tempnum As Byte 'numéro_index_courant_menu, temporaire
Global pctok(50, 10) As Integer '25li_Standard_puis_25_Personnalisé ; 1ère_col_pct_réussite_moyenne_numleçon ; cols_réussite_exos ; der_col_visualise_numleçon
Global vitok(50, 10) As Integer '25li_Standard_puis_25_Personnalisé ; 1ère_col_vitesse_moyenne_numleçon ; cols_vitesse_exos ; der_col_visualise_numleçon
Global pctt, pct1, erepeat, lrepeat, wrepeat As Byte 'pourcentage_réussite_exercice, pourcent_limite_pour passer_au_suivant, épelle, line_repeat, word_repeat
Global nfree As Byte 'numéros_libres_pour_ouvrir_fichiers
Global msgf As Byte 'sortie_de_msgform_pour_la_fonction_msghb_avec_0=Échap,1=Entrée,2=Répéter
Global passb As Byte 'mode_de_sortie_de_la_ligne_avec_0=Suivante_en_fin_de_ligne,1=Suivante_dès_1ère_erreur,2=Répéter_la_ligne_si_2erreurs
Global avecf2, avecf3, nobip, noF1, noechapF1, pasdepoint As Byte 'en_mode_aidef2, en_mode_aidef3, pas_de_bip, pas_aide_F1_fautes, pas_echap_msgform_de_F1_vers_menu, chargement_non_limité_par_un_"."
Global espacevalid, nodoublesono, noalt As Byte 'accepte_espace_pour_répéter_dans_leçon1, évite_double_sono_cas_particuliers, évite_répéter_fin_phrase_par_Alt
Global pagenum, pagemax As Byte 'numéro_page (0_interdit_numéros_page), dernière_page
Global forcepause, consult, quitactive, altf4 As Byte 'évite_de_skipper_score, mode_consulter_résultats, quitquit_en_cours, quit_par_Alt+F4
Global bascule, bipinhibit, quitF2, FullScreenSwitch As Byte 'bascule, bip_inhibit, quit_msg_F2_inutile, passer_en_plein_écran

Global leçon_courante, menu_courant, menu_suivant As Object 'exemples: leçon3, menu_leçon3, menu_leçon4
Global emax, startline, starttop As Date 'elapsed_max_allowed, temps_départ_nouvelle_ligne, temps_top_départ
Global bNumLockState, bCapsLockState, bScrollLockState As Boolean
Global scrw, frmw, scrh, frmh, vcomp, zfactor As Variant 'détecteurs de résolution d'écran
Global fbc, fbc_default, fbc_quit, fbc_f1, f_orange, f_rouge, f_rougevif, f_gris, f_violet, f_bleuclair, f_bleufoncé, f_vert As Variant 'font_back_colors
Global ffc, ffc_default, ffc_quit, ffc_f1, f_noirfoncé, f_noirgris, f_noirnoir As Variant 'font_fore_color

Global txtTps, lblFreq As String  'ajout avril 2008 pour les beep sonores

' Variables textes des pages d'explications
Global pgia1, pgia2, pgia3, pgia4, pgib1, pgib2, pgib3, pgib4, pgic1, pgic2, pgic3, pgic4, pgic5 As String
Global pg1a1, pg1a2, pg1a3, pg1am1, pg1am2, pg1b1, pg1b2, pg1bm1, pg1bm2, pg1bm3, pg1c0, pg1cm0, pg1d1, pg1d2 As String
Global pg2a1, pg2a2, pg2am1, pg2am2, pg2b1, pg2b2, pg2bm1, pg2bm2, pg2bm3, pg2c1, pg2c2, pg2cm1, pg2cm2, pg2cm3, pg2d1, pg2d2, pg2d3, pg2e1, pg2e2, pg2e3, pg2em1, pg2f0, pg2g1, pg2g2, pg2g3, pg2h0 As String
Global pg3a1, pg3a2, pg3a3, pg3a4, pg3am1, pg3am2, pg3b1, pg3b2, pg3c0 As String
Global pg4a1, pg4a2, pg4a3, pg4a4, pg4am1, pg4b0, pg4c0, pg4d0, pg4e0, pg4f0, pg4g0, pg4h0 As String
Global pg5a1, pg5a2, pg5b1, pg5b2, pg5c0 As String
Global pg6a1, pg6a2a, pg6a2b, pg6a3, pg6b1, pg6b2a, pg6b2b, pg6b3, pg6c0a, pg6c0b, pg6d0a, pg6d0b As String
Global pg7a1, pg7a2, pg7b1, pg7b2, pg7c0 As String
Global pg8a1, pg8a2, pg8a3, pg8a4, pg8am1, pg8b1, pg8b2, pg8c1, pg8c2, pg8cm1, pg8d1, pg8d2, pg8e1, pg8e2, pg8e3, pg8e4, pg8f1, pg8f2, pg8f3, pg8f4, pg8g1, pg8g2, pg8g3, pg8g4, pg8h1, pg8h2, pg8h3 As String
Global pg9a1, pg9a2, pg9b1, pg9b2, pg9b3, pg9c1, pg9c2, pg9d1, pg9d2 As String
Global pg10a1, pg10a2, pg10a3, pg10am1, pg10am2, pg10b1, pg10b2, pg10b3, pg10bm1, pg10bm2, pg10c1, pg10c2 As String
Global pg11a0, pg11b1, pg11b2, pg11c0 As String
Global pg12a1, pg12a2, pg12a3, pg12a4, pg12am1, pg12am2, pg12b1, pg12b2, pg12bm1, pg12c1, pg12c2, pg12d1, pg12d2 As String
Global pg13a1, pg13a2, pg13a3, pg13a4, pg13am1, pg13am2, pg13b1, pg13b2, pg13c1, pg13c2, pg13c3, pg13d1, pg13d2, pg13d3, pg13e1, pg13e2, pg13e3, pg13e4, pg13f1, pg13f2, pg13f3, pg13f4, pg13g0 As String
Global pg14a0, pg14b0, pg14c0 As String
Global pg15a0, pg15b0, pg15c0 As String
Global pg16a1, pg16a2, pg16a3, pg16a4, pg16b1, pg16b2, pg16b3, pg16c1, pg16c2, pg16cm1, pg16d1, pg16d2, pg16dm1 As String
Global pg17a0, pg17b0, pg17c0, pg17d0 As String
Global pg18a0, pg18b0, pg18c0, pg18d0, pg18e0 As String
Global pg19a0, pg19b0, pg19c0, pg19d0 As String

' Variables pour les traductions des barres de menus
Global meFichier, meQuitter_bm, meOptions, meStandard, mePersonnalisé, meDebExpliNormal, meDebExpliRapide, meDebGenLent, meDebGenMoyen, meDebGenVite As String
Global meBipClassique, meBipDifférent, meAide, meAideGénérale, meAideMémoire, meEnseignant, meSonorisation, meAproposde As String

'12/2011 zoom et couleurs
Global zoomfactor, zoomvalue As Variant
Global zoomlevel, msgNoZoom, msgWithZoom, msgDisplay As String
Global colorslevel, msgBasicColors, msgOtherColors As String
Global meNoZoom, meWithZoom, meBasicColors, meOtherColors As String
Global f_grispâle, f_grisfoncé, f_vertsombre, f_marronsombre, f_jaunevif, f_jaunetrèsvif, f_violetsombre, f_bleuvif, f_violetvif, f_vertpâle, f_vertvif, f_noirpresque, f_blanc, f_orangeclair As Variant

' Constant declarations, pour keybd_event, pour touches de Verrouillage CAPSLOCK, NUMLOCK, SCROLL,
' et pour remplacer SENDKEYS (simulation d'envoi de touches) pour Windows Vista
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_CAPITAL = &H14
Public Const VK_CAPITAL_BIS = &H10
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_DELETE = &H47
Public Const VK_ESCAPE = &H1B
' Ajouts pour le son en avril 2008
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000





'Touches à pb, voir les leçons 1 et 13 lorsque Jaws rajoute des codes parasites sur ces touches :
'Control 17
'Alt-Droit 17+18
'Alt-Gauche 18
'Win-Gauche 91, traitement spécifique pour casser la priorité donnée par Windows à la touche Win
'Win-Droit 92, traitement spécifique pour casser la priorité donnée par Windows à la touche Win
'Menu-Contextuel 93, nécessite de faire suivre par Echap
'Retour-De-Sendkeys 145

' **************** MAIN *******************************************************************
Public Sub main()
Module_routines.resetmsg
BLANCS6 = Space(6)
BLANCS12 = BLANCS6 + BLANCS6
CRLF = Chr(13) + Chr(10)
CRLF2 = Chr(13) + Chr(10) + CRLF
CRLF3 = Chr(13) + Chr(10) + CRLF2
old = "": tempo = "": nom_temp = "": scorecourant = "100 %"
repjawsnames = ""

'Suite
Module_global.main1
End Sub


' **************** MAIN1 **** TRADUIRE à DROITE DU SIGNE ÉGAL ******************************
Public Sub main1()

' Variables messages
bannerAuthorAddress = "herve.beranger@neuf.fr"
bannerCopyright = "  Copyleft 2008-2019 GNU/GPL"
bannerFunction = "logiciel d'apprentissage du clavier"
bannerLeçon = "Leçon"
bannerMenu = "Menu"
bannerNosell = "Ce logiciel libre est disponible sous licence GNU/GPL," & CRLF & "dont une copie se trouvera dans le fichier c:\ApprentiClavier\gpl.txt," & CRLF & "avec une traduction française non officielle gpl-fr.txt." & CRLF2 & "ApprentiClavier est diffusé gratuitement sur les sites" & CRLF & "www.apprenticlavier.com, http://apprenticlavier.wifeo.com, http://www.winaide.net"
bannerPrincipal = "Principal"
bannerThanks = "L'auteur adresse ses vifs remerciements" & CRLF & "aux précurseurs bénévoles de l'association Club Micro Son," & CRLF & "Robert Agro, Thierry Bertrand, Alain Rousseau."
bannerVersion = " ApprentiClavier Version 1.10"
clavierType = "AZERTY français (FRANCE)"
country = "Langue : française."
msgAide = "AIDE."
msgAideF3 = "Mode Aide-Mémoire, Échap pour sortir"
msgAltEspace = "Alt+ESPACE : pour ÉPELER."
msgAltF4 = "Alt+F4 : pour QUITTER IMMÉDIATEMENT."
msgAtt = "Attention "
msgAurevoir = "AU REVOIR."
msgAvec = "avec  "
msgBienvenue = "Bienvenue dans "
msgBienvFaudra = "Il faudra souvent appuyer sur la touche Entrée :"
msgBienvRedo = "  ERREUR.  Recommençons."
msgBienvRep = " Une deuxième fois, TAPEZ votre NOM."
msgBienvRepeat = "Par sécurité, une deuxième fois, TAPEZ votre NOM :"
msgBienvRetape = " TAPEZ votre NOM, ou sur Entrée."
msgBienvUsername = "Tapez votre NOM, ou tapez simplement sur Entrée :"
msgBip = "Bip sur fautes : " 'avril 2008
msgBipComment = "Voir ""Options"" dans la barre des menus." 'avril 2008
msgBipsAre = "Les BIPS sur les fautes sont réglés sur "
msgChoisir = "Flèches et Entrée : pour CHOISIR."
msgChoisissez = "  Choisissez."
msgClassique = "Classique"
msgCommandes = " commandes.       "
msgCommandesEn = "  commandes en  "
msgCommandesDispo = "Après ce MESSAGE, les COMMANDES DISPONIBLES seront :"
msgConseils = "CONSEILS PRATIQUES."
msgContinuer = "&Continuer (Entrée)"
msgCtrlEspace = "CONTROL+ESPACE : pour RÉPÉTER le MOT."
msgDébit = " Le débit général sera réglé sur "
msgDétecté = "DÉTECTÉ."
msgDictée = "Dictée "
msgDifférent = "Différent"
msgÉchap = "(Échap)"
msgÉchap2 = "(2 fois Échap)"
msgÉchap3 = "(3 fois Échap)"
msgEnter = "  Remontez le long du clavier principal par la droite." + CRLF + "  Vous rencontrez une grande touche verticale." + CRLF + "  C'est la touche Entrée."
msgEntréeContinuer = "Touche Entrée : pour CONTINUER."
msgEspace = "ESPACE : pour RÉPÉTER la TOUCHE demandée."
msgEstTerminée = "  est terminée."
msgExoFautes = "Exercice autour des fautes."
msgExoIdem = "Voulez-vous REFAIRE L'EXERCICE ? "
msgExoSuivant = "Voulez-vous FAIRE L'EXERCICE SUIVANT ? "
msgExpli = " Les explications seront données avec un débit "
msgF1F2F3 = " F1=Aide générale       F2=Description de la touche       F3=Aide-Mémoire "
msgF1Aide = "Touche F1 : pour l'AIDE."
msgF2Loc = "Touche F2 : pour LOCALISER la TOUCHE demandée."
msgF3AM = "Touche F3 : pour l'AIDE-MÉMOIRE."
msgFautesSur = " FAUTES SUR "
msgFormPressez = " Pressez " + CRLF + " ESPACE pour RÉPÉTER la PAGE d'explications," + CRLF + " puis les flèches pour RÉPÉTER chaque LIGNE," + CRLF2 + " ou Échap pour SORTIR," + CRLF + " ou Entrée pour CONTINUER."
msgFormRecommencer = "Veuillez RECOMMENCER."
msgFormVousétiez = " Vous étiez dans une page d'explications."
msgHit = "CONSIGNES DE FRAPPE."
msgKeyboard = "    Clavier : "
msgLaLeçon = "La leçon  "
msgLent = "Lent"
msgLevel = "         Niveau : "
msgLevelIs = "Le niveau des leçons est "
msgMajEspace = "MAJ+ESPACE : pour RÉPÉTER la FIN de la PHRASE."
msgMots = " mots.   "
msgMotsEn = "  mots en  "
msgMotsMinute = "mots-minute"
msgMoyen = "Moyen"
msgNofic = "Pas de fichier  "
msgNoficSono = "Pas de fichier de vocalisation JAWS."
msgNoSono = "Consultez ""Aide"" dans la barre supérieure."
msgNormal = "Normal"
msgOptions = "Touche Alt, puis O, OPTIONS : NIVEAU, DÉBIT, BIP, ZOOM, COULEURS."  ' 12/2011
msgPage = "Page "
msgPersonnalisé = "Personnalisé"
msgPourcent = " pourcent"
msgPourcentSeulement = " pourcent seulement"
msgPrécédent = "&Précédent"
msgPressEnter = " Appuyez maintenant sur la touche Entrée ! "
msgPressezF1 = "Pressez F1 pour vous exercer sur les fautes."
msgPrincContenu = "Voici le contenu du fichier "
msgPrincDansniveau = "  dans le niveau "
msgPrincPour = "  pour  "
msgPrincTerminé = "  TERMINÉ.  "
msgQuitter = "&Quitter "
msgQuitterAM = "&Quitter l'aide-mémoire "
msgQuitterMP = "&Quitter vers Menu Principal  "
msgQuitterVers = "&Quitter vers "
msgRapide = "Rapide"
msgRelaunch = "Il va falloir RELANCER ApprentiClavier."
msgReset = "Pour effacer tous les résultats d'un utilisateur et redémarrer à la première leçon, vous devrez entrer dans le menu principal." + CRLF2 + "Puis vous taperez Alt+O Options, puis I pour ""Redémarrer.""" + CRLF2 + "Appuyez maintenant sur la touche Entrée."
msgResetTitle = "Information."
msgRestart = "Voulez-vous Redémarrer à la première leçon, et effacer tous les résultats de l'utilisateur?"
msgRestartCmd = CRLF2 + "Pressez sur Entrée pour redémarrer, ou sur Échap pour annuler."
msgRestartTitle = "Confirmer le redémarrage."
msgRésultats = "    RÉSULTATS DE "
msgRéussià = "Cet exercice est réussi à "
msgRéussià100 = "nous enregistrons ici un taux de réussite de 100 pourcent. "
msgRéussite = " exercices réussis en moyenne à "
msgScore = "Score : "
msgSecondes = " secondes."
msgSes = "Ses "
msgSonori = "   Vocalisation : "
msgSortir = "Alt+Q ou Touche Échap : pour SORTIR."
msgSortir2 = "Alt+Q ou Touche Échap 2 fois : pour SORTIR."
msgSortir3 = "Alt+Q ou Touche Échap 3 fois : pour SORTIR."
msgSpeedExp = "    Avec débits : explications="
msgSpeedExpIs = "Le DÉBIT des EXPLICATIONS est "
msgSpeedGen = ", général="
msgSpeedGenIs = "Le DÉBIT GÉNÉRAL est réglé sur "
msgStandard = "Standard"
msgSuivant = "&Suivant"
msgTapez = "Tapez"
msgTapez2 = "Tapez :"
msgTapezTouche = "Tapez la touche :"
msgTranslator = "La traduction en anglais est due à Hervé Béranger."
msgTypeClavier = "Cette version est pour clavier " & clavierType & ", " & country
msgUser = "   Utilisateur : "
msgUserIs = "L'utilisateur est "
msgVite = "Vite"
msgWith = "Vous allez taper des mots "
pressez_basic = "      Pressez ESPACE pour RÉPÉTER,      " + CRLF + "       ou Entrée pour CONTINUER.        " + CRLF
pressez = CRLF2 + pressez_basic
pressez_ligne = CRLF2 + "   Pressez Flèche-Bas pour RÉPÉTER LIGNE PAR LIGNE, " + CRLF + pressez_basic
pressez_entrée = CRLF3 + "      Pressez Entrée pour CONTINUER.        " + CRLF
pressez_quit = CRLF2 + BLANCS12 + "  ATTENTION.             " + CRLF + BLANCS6 + "    Vous allez QUITTER         " + CRLF + BLANCS6 + "     ApprentiClavier.          " + CRLF2 + BLANCS12 + "   Pressez               " + CRLF + "      Entrée pour NE PAS QUITTER,    " + CRLF + "        ou ÉCHAP  pour QUITTER.      "
pressez_LeçonSuivante = CRLF2 + "     Pressez ESPACE pour RÉPÉTER ce message," + CRLF + "     ou Entrée pour PASSER à la leçon SUIVANTE."
pressez_suivant = CRLF2 + "     Pressez ESPACE pour RÉPÉTER ce message," + CRLF + "     ou Entrée pour PASSER à l'EXERCICE SUIVANT."
pressez_précédent = CRLF2 + "     Pressez ESPACE pour RÉPÉTER ce message," + CRLF + "     ou Entrée pour REFAIRE l'EXERCICE."
pressez_F2 = CRLF + " F2 ne peut pas vous renseigner," + CRLF + " car aucune TOUCHE n'est DEMANDÉE." + CRLF2
pressez_touche = "     Appuyez sur une touche."
perso_methode = CRLF2 + "L'enseignant peut créer chaque fichier texte," + CRLF + "en copiant celui du sous-dossier leçons\standard" + CRLF + "vers le sous-dossier leçons\personnalisé," + CRLF + "puis en le modifiant à son gré." + CRLF2 + "L'utilisateur choisira le niveau Standard, ou Personnalisé," + CRLF + "en se plaçant dans n'importe quel menu," + CRLF + "puis en tapant sur Alt, puis sur O, Options ;" + CRLF + "puis S, Standard, ou P, Personnalisé."

msgNoZoom = "Sans zoom" ' 12/2011
msgWithZoom = "Avec zoom"
msgBasicColors = "Couleurs basiques"
msgOtherColors = "Autres couleurs"
msgDisplay = "      Affichage : "

' Variables noms de touches
vvAccentGrave = "ACCENT GRAVE"
vvAccoladeDroite = "ACCOLADE DROITE"
vvAccoladeGauche = "ACCOLADE GAUCHE"
vvAlaligne = "à la ligne"
vvAlt = "ALT"
vvAltGr = "AltGr"
vvAltOuAltGr = "ALT ou AltGr"
vvApostrophe = "APOSTROPHE"
vvArrêtDéfil = "ArrêtDéfil"
vvAstérisque = "ASTÉRISQUE"
vvBarreOblique = "BARRE-OBLIQUE"
vvBarreObliqueInversée = "BARRE-OBLIQUE-INVERSÉE"
vvBarreVerticale = "BARRE-VERTICALE"
vvControl = "CONTROL"
vvControlDroit = "CONTROL-DROIT"
vvControlGauche = "CONTROL-GAUCHE"
vvCrochetDroit = "CROCHET DROIT"
vvCrochetGauche = "CROCHET GAUCHE"
vvCtrl = "CTRL"
vvDébut = "DÉBUT"
vvDeuxPoints = "DEUX-POINTS"
vvDiviser = "DIVISER"
vvEchap = "ECHAP"  'Préférer vvÉchap avec l'accent
vvÉchap = "ÉCHAP"
vvEntrée = "Entrée"
vvEspace = "ESPACE"
vvÉtoile = "ÉTOILE"
vvFin = "FIN"
vvFlecheBas = "FLECHE-BAS"
vvFlecheDroite = "FLECHE-DROITE"
vvFlecheGauche = "FLECHE-GAUCHE"
vvFlecheHaut = "FLECHE-HAUT"
vvGuillemet = "GUILLEMET"
vvImpression = "IMPRESSION"
vvInférieur = "INFÉRIEUR"
vvInsertion = "INSERTION"
vvMaj = "MAJ"
vvMajDroit = "MAJ-DROIT"
vvMajGauche = "MAJ-GAUCHE"
vvMajuscule = "  MAJUSCULE"
vvMenuContextuel = "MENU-CONTEXTUEL"
vvMinuscule = "  MINUSCULE"
vvMoins = "MOINS"
vvMultiplier = "MULTIPLIER"
vvPagePrécédente = "PAGE-PRÉCÉDENTE"
vvPageSuivante = "PAGE-SUIVANTE"
vvParenthèseDroite = "Parenthèse DROITE"
vvParenthèseGauche = "Parenthèse GAUCHE"
vvPause = "PAUSE"
vvPlus = "PLUS"
vvPoint = "POINT"
vvPointExclamation = "POINT d'EXCLAMATION"
vvPointInterrogation = "POINT d'INTERROGATION"
vvPointVirgule = "POINT-VIRGULE"
vvRetourArriere = "RETOUR-ARRIERE"
vvRetourArrière = "RETOUR-Arrière"
vvSansNom = "SansNom"
vvSouligné = "SOULIGNÉ"
vvSupérieur = "SUPÉRIEUR"
vvSuppression = "SUPPRESSION"
vvTab = "TAB"
vvTabulationAvant = "TABULATION-AVANT"
vvTabulationArrière = "TABULATION-Arrière"
vvTilde = "TILDE"
vvTiret = "TIRET"
vvVerrouillageMajuscules = "VERROUILLAGE-MAJUSCULES"
vvVerrouillageNumérique = "VERROUILLAGE-NUMÉRIQUE"
vvVirgule = "VIRGULE"
vvWindowsDroit = "WINDOWS-DROIT"
vvWindowsGauche = "WINDOWS-GAUCHE"

' Variables pages d'explications
' Présentation
pgia1 = bannerVersion & ", " + CRLF + bannerCopyright + ", Hervé Béranger," + CRLF + "   " + bannerAuthorAddress + "." + CRLF2 + "ATTENTION." + CRLF + "À la fin de chaque page d'explications, veuillez utiliser :" + CRLF2 + " - la touche Entrée pour CONTINUER," + CRLF + " - la touche ESPACE pour RÉPÉTER," + CRLF + " - la touche Échap  pour SORTIR." + CRLF2 + "La leçon 1 expliquera ces 3 touches." + CRLF2 + "Pour sortir, la touche Échap se trouve" + CRLF + "dans le coin en haut et à gauche du clavier."
'pgia2 = "ApprentiClavier permet l'apprentissage " + CRLF + "du clavier d'un ordinateur." + CRLF2 + "Il est dérivé du logiciel CLAVSON, mis au point par" + CRLF + "l'association Club Micro Son, notamment par" + CRLF + "Robert Agro, Thierry Bertrand, Alain Rousseau." + CRLF2 + msgTypeClavier + CRLF2 + msgTranslator
pgia2 = "ApprentiClavier permet l'apprentissage d'un clavier classique" + CRLF + "d'ordinateur,type 105 touches avec pavé numérique." + CRLF2 + msgTypeClavier
pgia3 = "Les enseignants peuvent modifier les exercices." + CRLF2 + "Ils peuvent éditer directement les fichiers textes" + CRLF + "du sous-dossier leçons\Personnalisé." + CRLF2 + "Puis, dans n'importe quel menu," + CRLF + "on pourra alors basculer le niveau," + CRLF + "en tapant sur Alt, puis sur O, Options." + CRLF + "Puis S, Standard, ou P, Personnalisé."
pgia4 = "Cette présentation est TERMINÉE." + CRLF2 + "Puisque vous l'avez suivie entièrement," + CRLF + msgRéussià100

' Pour qui, Pourquoi ?
pgib1 = "ApprentiClavier est développé à l'intention des non-voyants." + CRLF + "Ils utiliseront un lecteur d'écran tel que Jaws ou NVDA." + CRLF2 + "Dès que vous connaîtrez la touche Alt et quelques lettres," + CRLF + "vous pourrez modifier le débit des explications." + CRLF + "Dans n'importe quel menu, vous frapperez Alt, puis O pour Options." + CRLF + "Alors, dans le sous-menu, vous choisirez N pour Normal, ou R pour Rapide." + CRLF2 + "ApprentiClavier est utilisable par les voyants, sans vocalisation." + CRLF2 + "Les exercices progressifs sont conçus pour un apprentissage autonome." + CRLF2 + "Pour cette version," + CRLF + "certaines indications sur les emplacements des touches" + CRLF + "seraient incorrectes pour les ordinateurs portables."
pgib2 = "Il vous sera proposé successivement 5 sortes de leçons." + CRLF2 + "- l'entraînement à la frappe des lettres, des mots, des phrases;" + CRLF + "- des exercices de régularité;" + CRLF + "- des exercices de vitesse;" + CRLF + "- des dictées." + CRLF + "- des exercices sur les raccourcis clavier, et le pavé numérique."
pgib3 = "Il y a 3 objectifs :" + CRLF + "Premièrement : permettre le repérage et la frappe des touches." + CRLF + "Deuxièmement : faciliter l'utilisation d'une synthèse vocale." + CRLF + "Troisièmement : utiliser les combinaisons de touches."
pgib4 = "Cette explication est TERMINÉE. " + CRLF + "Puisque vous l'avez suivie entièrement," + CRLF + msgRéussià100

' Pour la frappe, des conseils
pgic1 = "INSTALLEZ-VOUS CONFORTABLEMENT." + CRLF + "Le dos est appuyé au dossier du siège." + CRLF + "Les bras sont en souplesse contre le corps." + CRLF + "Les poignets restent arrondis, détendus." + CRLF + "Chaque doigt se détache de la main pour frapper," + CRLF + "et revient à sa position de départ," + CRLF + "avant de relancer une nouvelle frappe." + CRLF2 + "Il est recommandé de pratiquer chaque jour," + CRLF + "pendant 20 à 25 minutes." + CRLF + "L'utilisateur est fatigué non seulement par la frappe," + CRLF + "mais aussi par l'écoute de la synthèse vocale."
pgic2 = "À PROPOS DES MENUS." + CRLF + "Le menu principal permet de choisir l'une des leçons." + CRLF + "Chaque leçon comporte aussi un menu," + CRLF + "avec un choix des exercices." + CRLF2 + "Dans tous les menus," + CRLF + "on peut utiliser les touches Flèche Haut ou Flèche Bas," + CRLF + "puis la touche Entrée pour valider votre choix." + CRLF2 + "La touche Échappement est très importante." + CRLF + "Échap vous fait quitter l'exercice," + CRLF + "revenir au menu des exercices," + CRLF + "puis revenir au menu principal," + CRLF + "et même quitter complètement."
pgic3 = "À PROPOS DES ERREURS DE FRAPPE." + CRLF + "En général, ApprentiClavier passe à la lettre suivante," + CRLF + "au bout de 5 fautes de frappe sur la lettre demandée." + CRLF + "Il est préférable de ne pas s'acharner sur une lettre," + CRLF + "il vaut mieux refaire l'exercice." + CRLF2 + "Dans certains exercices tels que les phrases et les dictées," + CRLF + "on passe à la lettre suivante dès la deuxième erreur." + CRLF + "Pour les raccourcis clavier, le but doit être zéro faute."
pgic4 = "À PROPOS DES RÉSULTATS." + CRLF + "Le taux de réussite sera annoncé," + CRLF + "à la fin de chaque exercice," + CRLF + "sous la forme d'un pourcentage." + CRLF + "Il est aussi enregistré et associé à votre nom." + CRLF + "Il sera affiché à la fin de chaque ligne des menus." + CRLF2 + "Il faut obtenir au moins 85 pourcent." + CRLF + "Même si le résultat est bon," + CRLF + "vous pourrez refaire l'exercice," + CRLF + "en vous déplaçant par les flèches dans les menus." + CRLF2 + "Vous pouvez consulter vos résultats," + CRLF + "en choisissant dans le menu principal" + CRLF + "l'avant-dernière option, Consulter."
pgic5 = "Ces conseils de frappe sont TERMINÉS." + CRLF + "Puisque vous les avez suivis entièrement," + CRLF + msgRéussià100

' Leçon1
pg1a1 = "Voici une DESCRIPTION du CLAVIER classique." + CRLF + "Le clavier principal," + CRLF + "c'est la plus grande partie du clavier" + CRLF + "qui se situe à gauche." + CRLF2 + "Ce clavier principal se compose de 5 rangées." + CRLF + "Attention, au-dessus," + CRLF + "on trouve encore une sixième rangée séparée." + CRLF + "C'est la rangée des touches de fonction." + CRLF2 + "À droite du clavier principal," + CRLF + "on trouve des groupes de touches," + CRLF + "puis encore à droite un damier de 5 rangées et 4 colonnes," + CRLF + "appelé pavé numérique."
pg1a2 = "Voici les PREMIERES TOUCHES indispensables." + CRLF2 + "La TOUCHE ESPACE." + CRLF2 + "Partez du devant du clavier principal." + CRLF2 + "Au centre de la ligne du bas, il y a une grande barre," + CRLF + "c'est la TOUCHE ESPACE, ou barre d'espacement." + CRLF2 + "Utilisez le pouce gauche ou le pouce droit."
pg1a3 = "Si vous utilisez une synthèse vocale, ne soyez pas surpris." + CRLF + "Dans la plupart des exercices," + CRLF + "vous n'entendrez pas la touche que vous venez de taper," + CRLF + "car ApprentiClavier désactive l'écho clavier." + CRLF2 + "Vous allez entendre deux notes de musique, " + CRLF + "car c'est le début de votre exercice."
pg1am1 = "La TOUCHE Entrée." + CRLF2 + msgEnter + CRLF2 + "Entrée vous permet de valider votre choix," + CRLF + "ou de progresser dans les pages d'explications." + CRLF2 + "Utilisez l'auriculaire droit."
pg1am2 = "La TOUCHE Échap." + CRLF + "Attendez d'avoir validé ce message," + CRLF + "avant d'exercer la touche Échap." + CRLF2 + "Partez du devant, et avec l'auriculaire gauche," + CRLF + "contournez le clavier principal." + CRLF2 + "La touche détachée qui fait le coin," + CRLF + "en haut à gauche du clavier, c'est la touche Échap." + CRLF2 + "Échap vous permet de quitter l'exercice" + CRLF + "ou l'action en cours," + CRLF + "et même de quitter progressivement ApprentiClavier." + CRLF2 + "Attention, exceptionnellement, il faudra frapper 2 fois Échap," + CRLF + "pour interrompre la leçon 1."
pg1b1 = "Les TOUCHES FLÈCHES." + CRLF2 + "Partez du devant du clavier principal," + CRLF + "allez à droite de la barre d'espacement," + CRLF + "sautez quatre touches." + CRLF2 + "Vous rencontrez un groupe isolé de quatre touches," + CRLF + "dont trois sont horizontales et une au-dessus," + CRLF + "ce sont les FLÈCHES." + CRLF2 + "De gauche à droite, on trouve :" + CRLF + "Flèche Gauche, Flèche Bas, Flèche Droite," + CRLF + "et au-dessus on trouve Flèche Haut." + CRLF2 + "Vous placerez l'index sur Flèche Gauche."
pg1b2 = "Avec les flèches, vous pourrez vous déplacer dans les menus," + CRLF + "ou dans les pages d'explications." + CRLF2 + "Par exemple, dans un menu," + CRLF + "Flèche Haut vous positionne sur la ligne au-dessus," + CRLF + "c'est-à-dire sur l'exercice précédent." + CRLF2 + "Dans toutes les pages d'explications," + CRLF + "si vous trouvez le débit des explications trop rapide," + CRLF + "vous pouvez relire immédiatement ligne par ligne," + CRLF + "en tapant sur Flèche Bas."
pg1bm1 = "La TOUCHE F1." + CRLF2 + "Partez du devant. Contournez par la gauche jusqu'en haut." + CRLF + "À droite de la touche Échap," + CRLF + "vous avez un groupe de 4 touches horizontales," + CRLF + "qui commençe à gauche par F1." + CRLF2 + "F1 est une touche qui vous offre de l'aide," + CRLF + "dans ApprentiClavier, comme dans la plupart des logiciels." + CRLF2 + "Utilisez l'annulaire gauche."
pg1bm2 = "La TOUCHE F2." + CRLF2 + "C'est la touche située à droite de la touche F1." + CRLF2 + "Dans ApprentiClavier, F2 est une touche" + CRLF + "qui vous indique l'emplacement et le doigt prévu" + CRLF + "pour chaque touche demandée." + CRLF2 + "Utilisez l'annulaire gauche comme pour F1."
pg1bm3 = "La TOUCHE F3." + CRLF2 + "C'est la touche située à droite de la touche F2." + CRLF2 + "Dans ApprentiClavier, F3 est la touche d'aide-mémoire," + CRLF + "qui vous informe sur chaque touche." + CRLF + "Vous pouvez utiliser F3 partout," + CRLF + "dans les leçons, dans les menus," + CRLF + "et même dès le lancement de la page de bienvenue." + CRLF2 + "Utilisez le majeur gauche."
pg1c0 = "Il y a trois touches à gauche de la barre ESPACE." + CRLF + "Ce sont les touches ALT, Windows, et CONTROL." + CRLF + "Attendez la leçon 13 pour le doigté de la touche Windows." + CRLF2 + "Voici la TOUCHE gauche ALT." + CRLF + "En général, cette touche lance la barre de menu." + CRLF + "Dans ApprentiClavier, après vous être placé dans le menu principal," + CRLF + "vous pourrez par exemple modifier le débit de la voix," + CRLF + "en tapant Alt, puis Flèche Droite pour passer aux Options," + CRLF + "puis des Flèches Bas, puis Entrée pour valider." + CRLF2 + "Avec le pouce gauche." + CRLF + "Partez de la barre ESPACE." + CRLF + "La touche Alt est la première touche à sa gauche."
pg1cm0 = "Les TOUCHES CONTROL. On les appelle aussi C t r l." + CRLF2 + "Avec l'auriculaire gauche." + CRLF + "Allez au coin en bas et à gauche du clavier principal." + CRLF + "C'est la touche Control de gauche." + CRLF2 + "Semblablement avec l'auriculaire droit," + CRLF + "allez au coin droit en bas du clavier principal." + CRLF + "C'est aussi une touche CONTROL." + CRLF2 + "On peut utiliser indifféremment" + CRLF + "la touche droite ou la touche gauche," + CRLF + "mais il vaut mieux s'entraîner sur les 2 touches."
pg1d1 = "Maintenant vous allez frapper les touches," + CRLF + "que vous venez de découvrir." + CRLF2 + msgConseils + CRLF2 + "Mettez vos mains sur la rangée de départ," + CRLF + "qui est la troisième rangée de touches" + CRLF + "en partant du devant." + CRLF2 + "L'index gauche est sur le point en relief de gauche, lettre F." + CRLF2 + "L'index droit est sur le point en relief de droite, lettre J."
pg1d2 = "Placez chaque pouce sur la barre d'espacement." + CRLF2 + "Gardez les poignets arrondis." + CRLF2 + "Tapez les touches demandées d'un coup sec," + CRLF + "en ramenant le doigt à sa position de départ."

' Leçon2
pg2a1 = msgConseils + CRLF2 + "Gardez vos mains sur la troisième ligne de touches" + CRLF + "en partant du devant du clavier." + CRLF2 + "L'index gauche doit être sur le point en relief à gauche," + CRLF + "lettre F." + CRLF2 + "L'index droit doit être sur le point en relief" + CRLF + "situé trois touches plus loin à droite, lettre J."
pg2a2 = msgHit + CRLF2 + "Avec la main gauche, vous devrez taper :" + CRLF2 + "Le Q avec l'auriculaire." + CRLF + "Le S avec l'annulaire." + CRLF + "Le D avec le majeur." + CRLF + "Le F avec l'index."
pg2am1 = "Avec la main droite, vous devrez taper :" + CRLF2 + "Le J : avec l'index." + CRLF + "Le K : avec le majeur." + CRLF + "Le L : avec l'annulaire." + CRLF + "Le M : avec l'auriculaire."
pg2am2 = " Attention." + CRLF2 + " Reprenez avec les 2 mains."
pg2b1 = msgConseils + CRLF2 + "Gardez les mains sur la troisième rangée de touches" + CRLF + "en partant du devant." + CRLF2 + "N'appuyez pas sur le clavier."
pg2b2 = msgHit + CRLF2 + "Tapez la touche prévue, mais seulement avec le doigt prévu." + CRLF + "À la cinquième erreur," + CRLF + "ApprentiClavier proposera la lettre qui suit." + CRLF2 + "Utilisez l'index pour les lettres G et H." + CRLF2 + "Pour le G, l'index gauche, en allant d'une touche vers la droite" + CRLF + "à partir du point en relief F." + CRLF2 + "Pour le H, l'index droit, en allant d'une touche vers la gauche" + CRLF + "à partir du deuxième point en relief J."
pg2bm1 = "Attention." + CRLF2 + "Avec la main gauche."
pg2bm2 = "Attention." + CRLF2 + "Avec la main droite."
pg2bm3 = "Attention." + CRLF2 + "Avec les deux mains."
pg2c1 = "Maintenant vous allez frapper les touches de la rangée" + CRLF + "juste au-dessus de celle que vous venez d'étudier." + CRLF2 + msgConseils + CRLF2 + "Frappez d'un coup sec et uniquement avec le doigt prévu." + CRLF2 + "Gardez les autres doigts à leur place" + CRLF + "et sans appuyer sur les touches."
pg2c2 = msgHit + CRLF + "Vous allez taper avec la main gauche." + CRLF2 + "Le A. Avec l'auriculaire." + CRLF + "Partez du Q, frappez nettement," + CRLF + "juste au-dessus, légèrement à gauche." + CRLF2 + "Le Z. Avec l'annulaire." + CRLF + "Partez du S. Frappez juste au-dessus, légèrement à gauche." + CRLF2 + "Le E. Avec le majeur. Partez du D." + CRLF + "Frappez juste au-dessus, légèrement à gauche." + CRLF2 + "Le R. Avec l'index. Partez du F." + CRLF + "Frappez juste au-dessus, légèrement à gauche."
pg2cm1 = "Maintenant, avec la main droite, vous devrez taper :" + CRLF2 + "Le U : avec l'index. Partez du J. Juste au-dessus, légèrement à gauche." + CRLF2 + "Le I : avec le majeur. Partez du K. Juste au-dessus, légèrement à gauche" + CRLF2 + "Le O : avec l'annulaire. Partez du L. Juste au-dessus, légèrement à gauche" + CRLF2 + "Le P : avec l'auriculaire. Partez du M. Juste au-dessus, légèrement à gauche."
pg2cm2 = "Attention." + CRLF2 + "Reprenez avec les 2 mains."
pg2cm3 = "C'est bientôt fini "
pg2d1 = "Maintenant vous allez taper des groupes de plusieurs mots," + CRLF + "avec les lettres déjà vues." + CRLF2 + "Les voyants doivent éviter de regarder le clavier." + CRLF2 + "Le caractère ESPACE n'est pas sonorisé habituellement," + CRLF + "sauf ici pour la frappe."
pg2d2 = msgConseils + CRLF2 + "Gardez les mains sur la rangée de départ sans appuyer." + CRLF2 + "Pour faire répéter la lettre, appuyez sur ESPACE." + CRLF2 + "Pour faire répéter le mot," + CRLF + "appuyez sur CONTROL en le gardant enfoncé, et frappez ESPACE." + CRLF2 + "On vous proposera 2 fois chaque groupe de mots."
pg2d3 = msgHit + CRLF2 + "Ecoutez les mots. Tapez la lettre demandée." + CRLF + "Tapez toujours un ESPACE entre les mots." + CRLF2 + "Pour faire un ESPACE," + CRLF + "tapez d'un coup sec sur la barre d'espacement." + CRLF + "Utilisez le pouce de la main," + CRLF + "qui n'a pas frappé la dernière lettre."
pg2e1 = "Maintenant vous allez utiliser l'index de chaque main."
pg2e2 = msgConseils + CRLF2 + "Gardez les poignets arrondis." + CRLF2 + "Ne bougez que le doigt prévu." + CRLF + "Ne faites aucune pression avec les autres doigts."
pg2e3 = msgHit + CRLF2 + "Tapez la lettre demandée." + CRLF + "À la cinquième erreur, on passera à la lettre suivante." + CRLF2 + "Main gauche, le T." + CRLF + "L'index part du F. Allez en diagonale juste au-dessus à droite, près de la lettre R." + CRLF2 + "Main droite, le Y." + CRLF + "L'index part du J. Allez en diagonale juste au-dessus à gauche, près de la lettre U."
pg2em1 = "Écoutez pour faire la différence entre le T de Thérèse, et le P de Patrick."
pg2f0 = "Maintenant vous allez taper à nouveau des groupes de 2 mots, séparés par un espace."
pg2g1 = "Maintenant vous allez taper des groupes de 3 mots." + CRLF2 + "Les mots sont courts, ils sont séparés par un espace."
pg2g2 = msgConseils + CRLF2 + "Tapez à votre rythme." + CRLF + "Les mains sont en place, elles sont souples." + CRLF2 + "Si vous écoutez une synthèse vocale," + CRLF + "ne la mettez pas trop fort, cela fatigue."
pg2g3 = msgHit + CRLF2 + "Ecoutez les mots. Tapez la lettre demandée."
pg2h0 = "Maintenant vous allez taper des phrases courtes." + CRLF2 + "Écoutez, puis frappez à votre rythme." + CRLF2 + "Tapez un espace entre les mots." + CRLF + "Mais ne tapez pas d'espace entre les phrases."

' Leçon3
pg3a1 = "Maintenant vous allez taper les lettres" + CRLF + "juste en-dessous de la rangée de départ." + CRLF2 + "RAPPEL." + CRLF2 + "F2 donne le nom et le doigt prévu" + CRLF + "pour la touche demandée." + CRLF2 + "F3 est la touche d'aide-mémoire," + CRLF + "qui donne le nom de la touche frappée."
pg3a2 = msgConseils + CRLF2 + "Pour chaque touche, partez de la rangée de départ," + CRLF + "Allez à la rangée inférieure vers la droite." + CRLF2 + "Frappez tranquillement. Déplacez seulement le doigt prévu." + CRLF + "Il faut que le doigt soit indépendant de la main."
pg3a3 = msgHit + CRLF2 + "Frappez la lettre demandée." + CRLF2 + "Ramenez le doigt à la position de départ." + CRLF2 + "N'appuyez pas avec les autres doigts."
pg3a4 = "Avec la main gauche vous devrez taper :" + CRLF2 + "Le W." + CRLF + "Avec l'auriculaire. Partez de Q. Juste en-dessous à droite." + CRLF2 + "Le X." + CRLF + "Avec l'annulaire. Partez de S. Allez en bas à droite." + CRLF2 + "Le C." + CRLF + "Avec le majeur. Partez de D. Allez en bas à droite." + CRLF2 + "Le V." + CRLF + "Avec l'index. Partez de F. Allez en bas à droite."
pg3am1 = "Les 3 rangées. Main gauche."
pg3am2 = "Maintenant, avec la main droite." + CRLF2 + "Le N. Avec l'index. Partez de J. Allez en bas, vers la gauche."
pg3b1 = msgConseils + CRLF2 + "Ecoutez et distinguez le B de Bernard et le V de Véronique." + CRLF2 + "Ecoutez et distinguez aussi le B de Bernard et le P de Patrick."
pg3b2 = msgHit + CRLF2 + "Maintenant vous allez taper le B." + CRLF + "Avec l'index. Partez de F. Allez en bas nettement à droite."
pg3c0 = msgConseils + CRLF2 + "Ecoutez et distinguez les mots qui se ressemblent."

' Leçon4
pg4a1 = "Cette leçon approfondit l'étude de l'alphabet." + CRLF2 + "Elle permet de reprendre des habitudes oubliées." + CRLF2 + "Attention. Le rythme sera plus rapide." + CRLF + "On vous proposera la lettre suivante" + CRLF + "dès que vous aurez frappé."
pg4a2 = "RAPPEL." + CRLF2 + "F1 donne l'aide sur les commandes." + CRLF + "F2 donne le nom et le doigt prévu pour la touche demandée." + CRLF + "F3 suivi d'une touche, donne le nom de la touche."
pg4a3 = msgConseils + CRLF2 + "Gardez les mains sur la troisième ligne de touches " + CRLF + "en partant du devant du clavier." + CRLF2 + "L'index gauche doit être sur le point en relief à gauche," + CRLF + "lettre F." + CRLF + "C'est la cinquième touche en partant du bord gauche du clavier." + CRLF2 + "L'index droit doit être sur le point en relief" + CRLF + "situé trois touches plus loin à droite, lettre J."
pg4a4 = "Toutes les indications partent de cette troisième rangée," + CRLF + "appelée rangée de départ." + CRLF2 + "Après 5 erreurs, on vous proposera la lettre suivante."
pg4am1 = "Attention." + CRLF2 + "Il faudra distinguer le F de François et le S de Simone."
pg4b0 = "Maintenant vous allez encore taper des touches" + CRLF + "de la rangée de départ," + CRLF + "celles à utiliser avec la main droite."
pg4c0 = "Maintenant vous allez taper des lettres" + CRLF + "de la rangée juste au-dessus de la rangée de départ." + CRLF2 + msgHit + CRLF2 + "Partez de la lettre de départ." + CRLF + "Déplacez le doigt." + CRLF + "Frappez d'un coup sec." + CRLF + "Ramenez le doigt à sa place." + CRLF + "Ne déplacez pas les autres doigts."
pg4d0 = msgConseils + CRLF2 + "Le G et le T se frappent avec l'index gauche." + CRLF2 + "Le G est juste à droite du F." + CRLF2 + "Le T est à droite du R, c'est-à-dire au-dessus et à droite du F." + CRLF2 + "Vous distinguerez le T de Thomas d'avec le D de Denis."
pg4e0 = msgConseils + CRLF2 + "Maintenant avec la main droite," + CRLF + "allez de la touche de départ vers le haut légèrement à gauche."
pg4f0 = msgConseils + CRLF2 + "Le H et le Y se frappent avec l'index droit." + CRLF2 + "Le H est juste à gauche du J." + CRLF2 + "Le Y est à gauche de U, c'est-à-dire au-dessus" + CRLF + "et très à gauche du J."
pg4g0 = msgConseils + CRLF2 + "Avec la main gauche, vous allez taper les lettres de la rangée" + CRLF + "juste en-dessous de la rangée de départ." + CRLF2 + "Les 4 lettres sont nettement décalées vers la droite."
pg4h0 = msgConseils + CRLF2 + "Maintenant, vous continuez dans la rangée" + CRLF + "en-dessous de la rangée de départ, avec l'index." + CRLF2 + "Avec la main gauche." + CRLF + "Le B. En-dessous du F, en extension très à droite." + CRLF2 + "Avec la main droite." + CRLF + "Le N. En-dessous du J. Légèrement à gauche."

' Leçon5
pg5a1 = msgConseils + CRLF2 + "Gardez vos mains sur la troisième ligne de touches" + CRLF + "en partant du devant du clavier." + CRLF2 + "L'index gauche doit être sur le point en relief à gauche," + CRLF + "lettre F." + CRLF2 + "L'index droit doit être sur le point en relief" + CRLF + "situé trois touches plus loin à droite, lettre J."
pg5a2 = msgHit + CRLF2 + "Avec la main gauche, vous devrez taper :" + CRLF2 + "Le Q avec l'auriculaire." + CRLF2 + "Le D avec le majeur."
pg5b1 = "Maintenant vous allez taper des proverbes connus." + CRLF2 + "La synthèse vocale ne dira la phrase qu'une fois," + CRLF + "puis chaque mot est prononcé une fois." + CRLF2 + "RAPPEL." + CRLF2 + "ESPACE répète la lettre demandée." + CRLF + "CONTROL+ESPACE répète le mot demandé."
pg5b2 = msgHit + CRLF2 + "Ne tapez qu'avec le doigt prévu." + CRLF2 + "Utilisez la touche F2 pour rappeler le doigt prévu."
pg5c0 = "Maintenant vous allez taper des phrases" + CRLF + "reprenant toutes les lettres de l'alphabet."


' Suite (sinon : procedure too large)
Module_global.main2
End Sub


' **************** MAIN2 **** TRADUIRE à DROITE DU SIGNE ÉGAL ******************************
Public Sub main2()
' Leçon 6
pg6a1 = "Maintenant vous allez taper des mots courts," + CRLF + "qui vous seront envoyés au hasard." + CRLF2 + "À chaque nouvel essai de cet exercice," + CRLF + "la séquence sera différente."
pg6a2a = "Attention." + CRLF2 + "Vous disposez de beaucoup de temps," + CRLF + "avec "
pg6a2b = " secondes pour chaque mot." + CRLF2 + "Vous pouvez faire épeler," + CRLF + "avec le pouce gauche enfoncé sur la touche ALT," + CRLF + "et un coup bref sur ESPACE," + CRLF + "mais le compteur de temps continuera."
pg6a3 = "On changera de mot à chaque erreur de frappe." + CRLF2 + "Si votre frappe est correcte, la synthèse vocale est muette." + CRLF2 + "Dès que le mot sera réussi," + CRLF + "ou bien si une frappe est incorrecte," + CRLF + "ou bien si le temps est écoulé," + CRLF + "on vous demandera de taper sur ESPACE."

pg6b1 = "Maintenant vous allez encore taper des mots courts," + CRLF + "qui vous seront envoyés au hasard." + CRLF2 + "À chaque nouvel essai de cet exercice," + CRLF + "la séquence sera différente."
pg6b2a = "Attention." + CRLF2 + "Vous disposez seulement de "
pg6b2b = " secondes pour chaque mot." + CRLF2 + "Vous pouvez faire épeler," + CRLF + "avec le pouce gauche enfoncé sur la touche ALT," + CRLF + "et un coup bref sur ESPACE," + CRLF + "mais le compteur de temps continuera."
pg6b3 = "On changera de mot à chaque erreur de frappe." + CRLF2 + "Si votre frappe est correcte, la synthèse vocale est muette." + CRLF2 + "Dès que le mot sera réussi," + CRLF + "ou bien si une frappe est incorrecte," + CRLF + "ou bien si le temps est écoulé," + CRLF + "on vous demandera de taper sur ESPACE."

pg6c0a = "Maintenant vous allez taper des mots plus longs," + CRLF + "très régulièrement." + CRLF2 + "Vous disposez de "
pg6c0b = " secondes pour chaque mot."
pg6d0a = "Maintenant vous allez taper des mots se terminant par : ation." + CRLF2 + "Vous disposez de "
pg6d0b = " secondes pour chaque mot."

' Leçon 7
pg7a1 = "Vous allez taper une phrase." + CRLF2 + msgConseils + CRLF2 + "Vous aurez 2 résultats." + CRLF2 + "     Le pourcentage de réussite." + CRLF2 + "     La vitesse de frappe."
pg7a2 = "Tapez la combinaison Control+ESPACE pour faire RÉPÉTER le MOT."
pg7b1 = msgConseils + CRLF2 + "N'augmentez votre vitesse que si vous obtenez " + CRLF + "au moins 80 pourcent de réussite." + CRLF2 + "RAPPEL." + CRLF2 + "Tapez en même temps sur CONTROL et ESPACE" + CRLF + "pour faire répéter le mot demandé."
pg7b2 = msgHit + CRLF2 + "Frappe régulière. Mains souples et sur la rangée de départ." + CRLF2 + "Dès la deuxième erreur, on vous proposera la lettre suivante."
pg7c0 = "Voici encore des phrases pour la vitesse."

' Leçon8
pg8a1 = "Maintenant vous allez frapper les trois touches" + CRLF + "pour les majuscules et les minuscules." + CRLF2 + "Attention. Vous remarquerez" + CRLF + "que la synthèse vocale n'indique MAJUSCULE ou MINUSCULE" + CRLF + "que lorsqu'on vient effectivement de changer d'état." + CRLF2 + "Par exemple," + CRLF + "un deuxième appel temporaire par la touche majuscule" + CRLF + "ne sera pas sonorisé en tant que MAJUSCULE."
pg8a2 = "Dans ApprentiClavier, on appelle VERROUILLAGE-MAJUSCULES" + CRLF + "ou FIXE-MAJUSCULES, en anglais CAPSLOCK," + CRLF + "la touche qui bloque le clavier en mode MAJUSCULES." + CRLF2 + "On appelle la touche MAJUSCULE en abrégé MAJ," + CRLF + "en anglais SHIFT," + CRLF + "la touche qui fait passer temporairement en majuscule." + CRLF2 + "Il y a en réalité 2 touches MAJUSCULES," + CRLF + "avec la même fonction." + CRLF2 + "On appelle MAJ-GAUCHE la touche située à gauche" + CRLF + "pour le passage temporaire en majuscule." + CRLF + "On appelle MAJ-DROIT la touche semblable située à droite."
pg8a3 = "La touche VERROUILLAGE-MAJUSCULES est à gauche du Q." + CRLF + "Pressez-la une seule fois." + CRLF + "Elle permettra alors d'écrire tout un texte en majuscules." + CRLF2 + "Attention. Pour déverrouiller," + CRLF + "c'est-à-dire pour revenir à des minuscules," + CRLF + "la touche à utiliser dépend du réglage dans Windows." + CRLF2 + "Parfois, ce sont les touches MAJ-GAUCHE et MAJ-DROIT" + CRLF + "qui suppriment le verrouillage, dès qu'on les relâche." + CRLF + "Il suffira alors de presser MAJ-GAUCHE ou MAJ-DROIT brièvement." + CRLF2 + "Mais en général, il faut appuyer à nouveau " + CRLF + "sur la touche VERROUILLAGE-MAJUSCULES elle-même."
pg8a4 = "Avec la main gauche." + CRLF2 + "Le VERROUILLAGE-MAJUSCULES. Avec l'auriculaire. Partez de Q." + CRLF + "Juste à gauche. C'est une touche plus grande." + CRLF2 + "Le MAJ-GAUCHE. Avec l'auriculaire. Partez de Q." + CRLF + "Allez en bas nettement à gauche en descendant d'une rangée." + CRLF + "C'est une touche assez grande."
pg8am1 = "Attention." + CRLF + "Sur certains claviers, la touche MAJ-DROIT est raccourcie," + CRLF + "pour laisser la place à une touche à sa gauche." + CRLF + "Ici nous donnons l'emplacement le plus fréquent." + CRLF2 + "Avec la main droite." + CRLF2 + "Le MAJ-DROIT. Avec l'auriculaire. Partez de M." + CRLF + "Descendez à la rangée inférieure, très à droite." + CRLF + "C'est une touche souvent très grande."

pg8b1 = "Maintenant vous allez taper les signes de ponctuation" + CRLF + "accessibles en minuscules," + CRLF + "c'est-à-dire la virgule, le point-virgule, le deux-points," + CRLF + "et le point d'exclamation." + CRLF2 + msgConseils + CRLF + "On rajoute un espace avant les ponctuations" + CRLF + "constituées de 2 signes," + CRLF + "donc on rajoute toujours un espace" + CRLF + "avant point-virgule, deux-points," + CRLF + "point d'interrogation et point d'exclamation." + CRLF + "Mais on n'en rajoute pas avant une virgule ou un point." + CRLF2 + "Et on place toujours un espace derrière," + CRLF + "si on continue sur la même ligne."
pg8b2 = msgHit + CRLF2 + "Maintenant vous allez taper avec la main droite." + CRLF2 + "La virgule." + CRLF + "Avec l'index. Partez de J. Allez en bas nettement à droite." + CRLF + "Le point-virgule." + CRLF + "Avec le majeur. Partez de K. Allez en bas nettement à droite." + CRLF + "Les deux-points." + CRLF + "Avec l'annulaire. Partez de L. Allez en bas bien à droite." + CRLF + "Le point d'exclamation." + CRLF + "Avec l'auriculaire. Partez de M. Allez en bas bien à droite."

pg8c1 = "Maintenant vous allez taper les signes de ponctuation" + CRLF + "accessibles en majuscules," + CRLF + "c'est-à-dire le point d'interrogation et le point." + CRLF2 + "Nous étudierons aussi " + CRLF + "la Barre-Oblique," + CRLF + "et le signe Section, utilisé parfois " + CRLF + "pour indiquer une nouvelle section dans un document."
pg8c2 = msgHit + CRLF2 + "Mettez-vous en majuscules bloquées," + CRLF + "en appuyant une fois sur la touche VERROUILLAGE-MAJUSCULES." + CRLF2 + "Puis avec la main droite." + CRLF2 + "Le point d'interrogation." + CRLF + "Avec l'index. Partez de J. Allez en bas nettement à droite." + CRLF + "Le point." + CRLF + "Avec le majeur. Partez de K. Allez en bas nettement à droite." + CRLF + "La Barre-Oblique." + CRLF + "Avec l'annulaire. Partez de L. Allez en bas bien à droite." + CRLF + "La Section, appelée aussi Chapitre ou Paragraphe." + CRLF + "Avec l'auriculaire. Partez de M. Allez en bas bien à droite."
pg8cm1 = "Attention." + CRLF + "Déverrouillez les majuscules, utilisez les touches MAJ-GAUCHE ou MAJ-DROIT." + CRLF + "Voici toutes les ponctuations de la rangée."

pg8d1 = "Maintenant vous allez taper des groupes de mots" + CRLF + "contenant des majuscules et des minuscules, avec des ponctuations."
pg8d2 = msgHit + CRLF2 + "Tapez MAJ-GAUCHE pour taper une lettre avec la main droite." + CRLF2 + "Tapez MAJ-DROIT pour taper une lettre avec la main gauche."

pg8e1 = "Maintenant vous allez taper le U grave," + CRLF + "ainsi que 2 accents particuliers." + CRLF2 + "Attention." + CRLF + "L'accent circonflexe et le tréma ne sont prononcés" + CRLF + "qu'après la frappe de la voyelle à accentuer." + CRLF2 + "Vous pouvez appeler la touche F2 pour vous aider" + CRLF + "au moment de taper une lettre accentuée."
pg8e2 = "Pour le circonflexe et le tréma, l'accent se frappe d'abord," + CRLF + "puis juste après, on frappe la voyelle." + CRLF2 + "Le tréma se fait en appuyant d'abord sur MAJ-GAUCHE," + CRLF + "en maintenant l'appui enfoncé, et en tapant la touche tréma." + CRLF + "Puis on tape la voyelle demandée."
pg8e3 = msgConseils + CRLF2 + "Avec la main droite." + CRLF2 + "Le  U grave." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Allez sur la même rangée à droite."
pg8e4 = "L'accent circonflexe." + CRLF + "En minuscule. Avec l'auriculaire. Partez de M." + CRLF + "Allez à la rangée au-dessus et à droite." + CRLF2 + "Le tréma." + CRLF + "En majuscule. D'abord pressez et maintenez Maj-Gauche." + CRLF + "Puis avec l'auriculaire" + CRLF + "allez à la même touche que l'accent circonflexe."

pg8f1 = "Maintenant vous allez taper des signes" + CRLF + "souvent employés en informatique." + CRLF2 + "Selon les synthèses vocales," + CRLF + "le signe Étoile est prononcé Étoile ou Astérisque." + CRLF + "Le signe  Inférieur à  est prononcé Inférieur." + CRLF + "Le signe  Supérieur à  est prononcé Supérieur."
pg8f2 = "Les signes Astérisque et Inférieur se frappent en minuscules." + CRLF2 + "Le Supérieur se fait en majuscule avec MAJ-DROIT."
pg8f3 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Attention. Sur certains claviers," + CRLF + "la touche Astérisque se trouve sous la touche Entrée." + CRLF + "Ici nous donnons l'emplacement le plus fréquent." + CRLF2 + "L'Astérisque." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Allez 2 touches plus loin sur la même rangée à droite."
pg8f4 = "Avec la main gauche." + CRLF2 + "L'Inférieur." + CRLF + "En minuscule. Avec l'auriculaire. Partez de Q." + CRLF + "Allez à la rangée inférieure et à gauche, juste avant MAJ-GAUCHE." + CRLF2 + "Le Supérieur." + CRLF + "En majuscule. Allez à la même touche que pour l'Inférieur."

pg8g1 = "Maintenant vous allez taper 4 signes" + CRLF + "souvent employés dans les textes et en informatique." + CRLF2 + "Le pourcent, la lettre grecque Mu ou Micro, le dollar" + CRLF + "et la livre anglaise."
pg8g2 = msgHit + CRLF2 + "Le PourCent et le Mu se frappent en majuscules avec MAJ-GAUCHE." + CRLF2 + "Avec la synthèse vocale, le Mu est souvent prononcé MICRO."
pg8g3 = "Avec la main droite." + CRLF2 + "Le PourCent." + CRLF + "Avec l'auriculaire en majuscule. Partez de M." + CRLF + "Allez une touche à droite." + CRLF2 + "Attention. Sur certains claviers," + CRLF + "la touche Mu se trouve en dessous de la touche Entrée." + CRLF + "Ici nous donnons l'emplacement le plus fréquent." + CRLF + "Le Mu." + CRLF + "Avec l'auriculaire en majuscule. Partez de M." + CRLF + "Allez 2 touches plus loin à droite." + CRLF + "C'est donc en général la même touche que l'astérisque," + CRLF + "mais en majuscule."
pg8g4 = "Avec la main droite." + CRLF2 + "Le Dollar." + CRLF + "Avec l'auriculaire en minuscule. Partez de M." + CRLF + "Allez à la rangée supérieure très en extension à droite," + CRLF + "avant la grande touche." + CRLF2 + "La Livre." + CRLF + "Avec l'auriculaire en majuscule. Partez de M." + CRLF + "Allez à la même touche que le Dollar."

pg8h1 = "Maintenant vous allez taper les lettres accentuées," + CRLF + "e aigu, e grave, a grave," + CRLF + "situées dans la rangée la plus éloignée du clavier principal." + CRLF + "C'est la deuxième rangée au-dessus de la rangée de départ." + CRLF2 + msgConseils + CRLF2 + "Repérez la touche à frapper." + CRLF + "Le doigt se libère de la main sans l'entraîner." + CRLF + "Il revient à sa position de départ" + CRLF + "avant de frapper une autre touche."
pg8h2 = msgHit + CRLF2 + "Avec la main gauche." + CRLF + "Le é." + CRLF + "Avec l'auriculaire. Partez de Q." + CRLF + "Montez de 2 rangées directement au-dessus." + CRLF + "C'est la troisième touche en partant de la gauche" + CRLF + "dans cette rangée."
pg8h3 = "Avec la main droite." + CRLF2 + "Le è." + CRLF + "Avec l'index. Partez de J." + CRLF + "Montez de 2 rangées au-dessus légèrement à gauche." + CRLF2 + "Le à." + CRLF + "Avec l'annulaire. Partez de L." + CRLF + "Montez de 2 rangées au-dessus légèrement à droite."

' Leçon9
pg9a1 = "Maintenant vous allez taper le signe ²" + CRLF + "qui correspond à l'exposant au carré." + CRLF + "Vous taperez aussi le signe & prononcé Écommercial," + CRLF + "qui remplace la préposition ET." + CRLF2 + msgHit + CRLF2 + "Avec la main gauche, frappez sur la rangée du haut," + CRLF + "2 rangées au-dessus de la rangée de départ."
pg9a2 = "Le ²." + CRLF + "C'est la touche la plus à gauche. Avec l'auriculaire." + CRLF + "En minuscule. Partez de Q. Montez 2 rangées au-dessus." + CRLF + "En extension maximum à gauche." + CRLF2 + "Le &." + CRLF + "Avec l'auriculaire. En minuscule. Partez de Q." + CRLF + "Montez 2 rangées au-dessus, légèrement à gauche."

pg9b1 = "Maintenant vous allez taper des signes de ponctuation" + CRLF + "de la rangée supérieure." + CRLF + "Cet exercice étudie les guillemets, l'apostrophe," + CRLF + "la parenthèse gauche et le tiret." + CRLF2 + msgConseils + CRLF + "Toutes les ponctuations de la rangée supérieure " + CRLF + "se font en minuscules." + CRLF2 + "On ne rajoute pas d'espace, sauf dans les 4 cas suivants :" + CRLF + "avant d'ouvrir le guillemet," + CRLF + "avant d'ouvrir une parenthèse," + CRLF + "après avoir fermé le deuxième guillemet," + CRLF + "après avoir fermé la parenthèse droite."
pg9b2 = "PRONONCIATION." + CRLF2 + "Selon les synthèses, la première parenthèse se prononce : " + CRLF + "ouvre parenthèse, parenthèse gauche ou parenthèse ouverte." + CRLF2 + "La parenthèse qui termine se prononce : " + CRLF + "ferme parenthèse, parenthèse droite ou parenthèse fermée." + CRLF + "Il en sera de même pour les crochets et les accolades."
pg9b3 = msgHit + CRLF + "Avec la main gauche." + CRLF + "Le guillemet." + CRLF + "Avec l'annulaire. Partez de S." + CRLF + "Montez de 2 rangées au-dessus directement." + CRLF + "L'apostrophe." + CRLF + "Avec le majeur. Partez de D." + CRLF + "Montez de 2 rangées au-dessus directement." + CRLF + "La parenthèse gauche." + CRLF + "Avec l'index. Partez de F." + CRLF + "Montez de 2 rangées au-dessus directement." + CRLF + "Le tiret." + CRLF + "Avec l'index en extension. Partez de F." + CRLF + "Montez de 2 rangées au-dessus et à droite."

pg9c1 = "Maintenant vous allez taper des signes de ponctuation" + CRLF + "de la rangée du haut à main droite." + CRLF2 + "Vous allez étudier le souligné et la parenthèse droite."
pg9c2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Le signe Souligné." + CRLF + "C'est un caractère qui est une sorte de tiret." + CRLF + "Avec l'index. Partez de J." + CRLF + "Montez de 2 rangées au-dessus, et légèrement à droite." + CRLF2 + "La parenthèse droite." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Montez de 2 rangées au-dessus directement."

pg9d1 = "Maintenant vous allez taper encore 4 signes" + CRLF + "de la rangée supérieure." + CRLF2 + "Le ç, le Degré, le signe Égal, et le signe d'addition Plus." + CRLF2 + msgConseils + CRLF2 + "Le ç se frappe uniquement en minuscule." + CRLF + "En majuscule, on tape simplement le C majuscule." + CRLF2 + "Le signe Égal se fait uniquement en minuscule," + CRLF + "même au milieu des chiffres." + CRLF2 + "Par contre, on tape en majuscules les signes Degré et Plus."
pg9d2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Le ç. Avec le majeur. Partez de K." + CRLF + "Montez de 2 rangées au-dessus directement." + CRLF + "Le °. En majuscule. Avec l'auriculaire." + CRLF + "Partez de M. Montez de 2 rangées au-dessus directement." + CRLF + "Le Égal. En minuscule. Avec l'auriculaire." + CRLF + "Partez de M. Montez de 2 rangées au-dessus, et à droite." + CRLF + "Le Plus. En majuscule. Avec l'auriculaire." + CRLF + "Partez de M. Montez de 2 rangées au-dessus, et à droite." + CRLF + "C'est la même touche que le signe Égal."

' Leçon10
pg10a1 = "Maintenant vous allez taper les chiffres." + CRLF2 + "PRONONCIATION." + CRLF2 + "Les nombres sont prononcés globalement" + CRLF + "en prononciation française," + CRLF + "par exemple 10000 est prononcé dix mille." + CRLF2 + "Mais certaines synthèses vocales ne prononcent que" + CRLF + "chiffre après chiffre, soit 1 0 0 0 0." + CRLF2 + "Pour une prononciation globale," + CRLF + "écrivez le nombre complet sans espaces."
pg10a2 = msgConseils + CRLF + "Les chiffres se trouvent à 2 endroits." + CRLF + "D'une part sur la rangée supérieure du clavier principal," + CRLF + "d'autre part dans le pavé numérique de droite." + CRLF + "Ici nous étudions seulement le clavier principal." + CRLF2 + "Les chiffres du clavier principal se tapent en majuscules." + CRLF2 + "En France, les décimales se marquent par une virgule," + CRLF + "et les milliers par un point ou un espace." + CRLF2 + "Dans beaucoup de pays anglo-saxons, c'est l'inverse." + CRLF + "Les décimales anglo-saxonnes se marquent par un point," + CRLF + "et les milliers par une virgule."
pg10a3 = msgHit + CRLF + "En majuscules. Avec la main gauche." + CRLF + "Le 1. Avec l'auriculaire." + CRLF + "Partez de Q. Montez de 2 rangées au-dessus, nettement à gauche." + CRLF + "C'est la deuxième touche de cette rangée." + CRLF + "Le 2. Avec l'auriculaire." + CRLF + "Partez de Q. Montez de 2 rangées au-dessus, légèrement à droite." + CRLF + "Le 3. Avec l'annulaire." + CRLF + "Partez de S. Montez de 2 rangées au-dessus, légèrement à droite." + CRLF + "Le 4. Avec le majeur." + CRLF + "Partez de D. Montez de 2 rangées au-dessus, légèrement à droite."
pg10am1 = "Attention, nombres à 2 chiffres, tapez le nombre entièrement."
pg10am2 = "Attention, nombres à 3 chiffres, tapez le nombre entièrement."

pg10b1 = "Maintenant vous allez taper les trois chiffres 5, 6, 7." + CRLF2 + "PRONONCIATION." + CRLF2 + "Avec la synthèse vocale, écoutez la différence entre 5 et 7." + CRLF2 + "Utilisez les touches F2 et F3 " + CRLF + "pour rappeler l'emplacement et le doigt prévu."
pg10b2 = msgHit + CRLF + "Avec la main gauche." + CRLF + "Le 5. Avec l'index." + CRLF + "Partez de F. Montez de 2 rangées au-dessus légèrement à droite." + CRLF2 + "Le 6. Avec l'index." + CRLF + "Partez de F. Montez de 2 rangées au-dessus et très à droite."
pg10b3 = "Avec la main droite." + CRLF2 + "Le 7. Avec l'index." + CRLF + "Partez de J. Montez de 2 rangées au-dessus légèrement à gauche."
pg10bm1 = "Attention tous les chiffres de 1 à 7 "
pg10bm2 = "Attention voici des nombres "

pg10c1 = "Maintenant vous allez taper le 8, le 9, et le 0." + CRLF2 + "PRONONCIATION." + CRLF2 + "Avec la synthèse vocale, le 8 peut se confondre avec le 6," + CRLF + "quand il est pris isolément."
pg10c2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Le 8." + CRLF + "Avec l'index. Partez de J." + CRLF + "Montez de 2 rangées au-dessus, et légèrement à droite." + CRLF + "Le 9." + CRLF + "Avec le majeur. Partez de K." + CRLF + "Montez de 2 rangées au-dessus, légèrement à droite." + CRLF + "Le 0." + CRLF + "Avec l'annulaire. Partez de L. Montez de 2 rangées au-dessus, légèrement à droite."

' Leçon11
pg11a0 = "Vous allez taper une phrase." + CRLF2 + msgConseils + CRLF2 + "Vous aurez 2 résultats." + CRLF2 + "     Le pourcentage de réussite." + CRLF2 + "     La vitesse de frappe."

pg11b1 = msgConseils + CRLF2 + "N'augmentez votre vitesse que si vous obtenez au moins" + CRLF + "80 pourcent de réussite." + CRLF2 + "RAPPEL." + CRLF2 + "Tapez en même temps sur CONTROL et ESPACE pour répéter" + CRLF + "le mot demandé."
pg11b2 = msgHit + CRLF2 + "Frappe régulière. Mains souples et sur la rangée de départ." + CRLF2 + "ATTENTION." + CRLF + "Dès la deuxième erreur, on vous proposera la lettre suivante."

pg11c0 = "Voici encore des phrases pour la vitesse."

' Leçon12
pg12a1 = "Maintenant vous allez frapper 2 touches" + CRLF + "qui agissent sur le texte." + CRLF2 + "Ce sont les touches INSERTION, en anglais INSERT, ou INS," + CRLF + "et SUPPRESSION, en anglais DELETE, ou DEL." + CRLF2 + "Attention." + CRLF2 + "Le pilote de synthèse peut modifier le comportement" + CRLF + "de certaines touches." + CRLF2 + "Par exemple Jaws peut être configuré" + CRLF + "pour supprimer l'effet bascule de la touche INSERTION."
pg12a2 = "Normalement, la touche INSERTION est une bascule." + CRLF2 + "Quand vous tapez du texte par-dessus un texte existant," + CRLF + "en général, par défaut il s'insère," + CRLF + "c'est-à-dire il s'ajoute au texte existant." + CRLF2 + "En appuyant sur la touche INSERTION," + CRLF + "vous basculez en mode REMPLACEMENT," + CRLF + "c'est-à-dire que votre future frappe écrasera l'ancien texte," + CRLF + "si votre curseur se trouve par-dessus un texte." + CRLF2 + "Si vous appuyez à nouveau plus tard sur la touche INSERTION," + CRLF + "vous reviendrez au mode INSERTION."
pg12a3 = "La touche SUPPRESSION supprime immédiatement" + CRLF + "le caractère courant de votre texte," + CRLF + "celui qui se trouve juste après le curseur." + CRLF2 + "PRONONCIATION." + CRLF + "Pour la touche SUPPRESSION," + CRLF + "la synthèse prononce généralement le caractère qui se présente" + CRLF + "à droite du caractère qui vient d'être supprimé."
pg12a4 = msgConseils + CRLF2 + "A droite du clavier principal," + CRLF + "dans le prolongement des 2 rangées supérieures," + CRLF + "on trouve un ensemble de 2 ou 3 lignes de 3 touches." + CRLF2 + "Voici la touche SUPPRESSION." + CRLF + "Avec la main droite." + CRLF + "Partez de M. Allez à droite hors du clavier principal." + CRLF + "Placez l'index sur la première touche" + CRLF + "de la petite rangée rencontrée."
pg12am1 = "Touche INSERTION." + CRLF + "En général, la synthèse vocale prononce le mot INSERTION" + CRLF + "ou REMPLACEMENT." + CRLF + "Parfois, au lieu de REMPLACEMENT, le logiciel indique" + CRLF + "SURFRAPPE ou REFRAPPE." + CRLF2 + "C'est la première touche à gauche" + CRLF + "dans la petite rangée au-dessus." + CRLF2 + "Elle se trouve donc au-dessus de la touche SUPPRESSION."
pg12am2 = "Attention, avec des lettres du clavier."

pg12b1 = "Maintenant vous allez frapper les touches" + CRLF + "qui vous mettent en début ou en fin de ligne." + CRLF2 + "Ceci est vrai en général pour les traitements de texte" + CRLF + "tels que WORD." + CRLF2 + msgConseils + CRLF2 + "La touche DÉBUT, en anglais HOME," + CRLF + "vous place au début de la ligne courante." + CRLF2 + "La touche FIN, en anglais END," + CRLF + "vous place à la fin de la ligne courante."
pg12b2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "La touche FIN." + CRLF + "Partez de M. Placez la main sur la rangée du bas" + CRLF + "du groupe des 2 ou 3 rangées de 3 touches." + CRLF + "Le majeur frappe la deuxième touche de cette rangée." + CRLF2 + "La touche DÉBUT." + CRLF + "Partez de M. C'est le même groupe de touches." + CRLF + "Avec le majeur, frappez la deuxième touche de la rangée au-dessus."
pg12bm1 = "Avec des lettres "

pg12c1 = "Maintenant vous allez frapper les touches" + CRLF + "permettant de changer de page." + CRLF2 + msgConseils + CRLF2 + "PAGE-PRÉCÉDENTE, en anglais Page-Up," + CRLF + "permet de passer à la page précédente ou à l'écran précédent." + CRLF2 + "PAGE-SUIVANTE, en anglais Page-Down," + CRLF + "permet de passer à la page suivante ou à l'écran suivant." + CRLF2 + "Par exemple avec PAGE-PRÉCÉDENTE, dans ApprentiClavier," + CRLF + "vous pouvez revenir en arrière dans les pages d'explications."
pg12c2 = msgHit + CRLF2 + "Partez du M. Allez à droite hors du clavier principal." + CRLF2 + "La touche PAGE-SUIVANTE." + CRLF + "Avec l'annulaire." + CRLF + "Frappez la troisième et dernière touche du petit groupe," + CRLF + "dans sa rangée inférieure." + CRLF2 + "La touche PAGE-PRÉCÉDENTE." + CRLF + "C'est la touche au-dessus de PAGE-SUIVANTE."

pg12d1 = "Maintenant vous allez frapper 3 touches" + CRLF + "situées au-dessus du groupe des 6 touches." + CRLF2 + "Il s'agit des touches IMPRESSION, ArrêtDéfil, et PAUSE." + CRLF2 + "IMPRESSION, en anglais PrintScreen, permet d'imprimer l'écran" + CRLF + "ou le document." + CRLF2 + "ArrêtDéfil, en anglais ScrollLock, supprime la possibilité" + CRLF + "de faire défiler les pages." + CRLF2 + "PAUSE, en anglais ATTENTION, stoppe l'exécution" + CRLF + "du logiciel en cours." + CRLF2 + "Ces commandes n'agissent que si le logiciel actif le permet."
pg12d2 = msgHit + CRLF2 + "Touche IMPRESSION." + CRLF + "Avec la main droite, au-dessus de la touche INSERTION." + CRLF + "Partez de M. Allez à droite hors du clavier principal." + CRLF + "Placez l'index sur la première touche de la rangée supérieure." + CRLF2 + "A droite d'IMPRESSION, le majeur trouve ArrêtDéfil." + CRLF2 + "A sa droite, l'annulaire trouve la PAUSE."

' Suite (sinon procedure too large)
Module_global.main3
End Sub


' **************** MAIN3 **** TRADUIRE à DROITE DU SIGNE ÉGAL ******************************
Public Sub main3()

' Leçon13
pg13a1 = "Maintenant vous allez apprendre à utiliser les touches" + CRLF + "pour les menus ou les boîtes de dialogue de Windows." + CRLF2 + "Ces touches sont situées à gauche et à droite" + CRLF + "de la grande barre ESPACE." + CRLF2 + "Les touches CONTROL, WINDOWS, ALT sont doublées." + CRLF + "On les trouve à gauche et à droite." + CRLF2 + "La touche Menu-Contextuel se trouve seulement à droite." + CRLF2 + "Attention." + CRLF + "L'objectif est de ne faire aucune erreur."
pg13a2 = "PRONONCIATION." + CRLF2 + "Ces touches ne sont pas prononcées par la synthèse vocale," + CRLF + "sauf MENU-CONTEXTUEL qui est parfois prononcé APPLICATION." + CRLF2 + "C'est seulement le résultat de leur action qui est prononcé." + CRLF2 + "Attention. En général," + CRLF + "la touche CONTROL stoppe la prononciation du message en cours."
pg13a3 = msgConseils + CRLF2 + "Avec la main gauche." + CRLF2 + "La touche CONTROL." + CRLF + "Avec l'auriculaire. Partez de Q. Descendez très en bas à gauche." + CRLF + "C'est la touche d'angle du clavier." + CRLF2 + "La touche ALT." + CRLF + "Avec le pouce. Partez de la barre ESPACE." + CRLF + "C'est la première touche à sa gauche."
pg13a4 = "Les touches Windows n'existaient pas sur les claviers très anciens." + CRLF + "On les appelle aussi touches LOGO." + CRLF2 + "Attention, normalement elles lancent le menu Démarrer de Windows." + CRLF + "Il faudrait alors presser Échap pour annuler cette action." + CRLF2 + "La touche WINDOWS de gauche." + CRLF + "Avec l'auriculaire. Partez de Q. Descendez  de 2 rangées." + CRLF + "Elle se trouve entre CONTROL et ALT."
pg13am1 = "Avec la main droite." + CRLF2 + "La touche AltGr." + CRLF + "Elle peut avoir une action différente de celle de la touche ALT." + CRLF + "Avec le pouce. Partez de la barre ESPACE." + CRLF + "C'est la première touche à sa droite." + CRLF2 + "La touche WINDOWS de droite." + CRLF + "Avec l'auriculaire. Partez de M. Descendez de 2 rangées." + CRLF + "C'est juste à droite de AltGr." + CRLF2 + "La touche CONTROL de droite." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Descendez de 2 rangées en extension à droite." + CRLF + "C'est la touche d'angle du clavier."
pg13am2 = "Le MENU-CONTEXTUEL." + CRLF + "MENU-CONTEXTUEL lance normalement un menu lié au contexte." + CRLF + "On l'annule par Échap." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Descendez de 2 rangées et à droite." + CRLF + "C'est juste à gauche du CONTROL-DROIT."

pg13b1 = "Maintenant voici 2 touches qui agissent dans le texte." + CRLF2 + "La touche TABULATION-AVANT, appelée TAB," + CRLF + "déplace le curseur texte vers la droite sur la même ligne." + CRLF + "C'est un déplacement défini de quelques caractères." + CRLF + "La touche TAB n'est pas prononcée." + CRLF2 + "En mode Majuscule, cette TAB devient une TABULATION-Arrière" + CRLF + "qui recule le curseur de quelques caractères vers la gauche." + CRLF2 + "Le RETOUR-Arrière, en anglais BACKSPACE," + CRLF + "recule le curseur d'un seul caractère," + CRLF + "effaçant ainsi le caractère que vous venez de taper." + CRLF + "En général, la synthèse vocale prononce le caractère effacé."
pg13b2 = msgHit + CRLF2 + "Avec la main gauche." + CRLF + "La touche TAB." + CRLF + "Avec l'auriculaire. Partez de Q." + CRLF + "Allez à la rangée au-dessus et très à gauche." + CRLF + "C'est une touche plus grande." + CRLF2 + "Avec la main droite." + CRLF + "Le RETOUR-Arrière." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Montez de 2 rangées au-dessus, en extension à droite." + CRLF + "C'est la touche plus grande dans le coin du clavier principal."

pg13c1 = "Maintenant vous allez frapper les touches de fonction F1 à F12." + CRLF2 + msgConseils + CRLF2 + "Repérez les touches sur la rangée la plus éloignée de vous," + CRLF + "dans le clavier principal." + CRLF + "Partez de la rangée de départ, en extension maximum." + CRLF + "Frappez uniquement avec le doigt prévu."
pg13c2 = "CONSIGNES DE FRAPPE. Avec la main gauche." + CRLF2 + "F1." + CRLF + "Avec l'annulaire. En extension maximum, nettement à gauche." + CRLF + "F2." + CRLF + "Avec l'annulaire. En extension maximum, légèrement à droite." + CRLF + "F3." + CRLF + "Avec le majeur. En extension maximum légèrement à droite." + CRLF + "F4." + CRLF + "Avec l'index. En extension maximum, légèrement à droite." + CRLF + "F5." + CRLF + "Avec l'index. En extension maximum et nettement à droite."
pg13c3 = "Attention. Avec la main droite." + CRLF2 + "F6." + CRLF + "Avec l'index. En extension maximum, directement." + CRLF + "C'est la deuxième touche du deuxième groupe de 4 touches." + CRLF2 + "F7." + CRLF + "Avec le majeur. En extension maximum, directement." + CRLF2 + "F8." + CRLF + "Avec l'annulaire. En extension maximum, directement." + CRLF2 + "F9 à F12." + CRLF + "Avec l'annulaire. En extension maximum, vers la droite."

pg13d1 = "Un raccourci-clavier est une combinaison de 2 ou 3 touches." + CRLF2 + "Par exemple, si vous tenez une touche CONTROL enfoncée" + CRLF + "pendant que vous frappez brièvement une autre touche," + CRLF + "vous exécutez une combinaison qui peut agir" + CRLF + "dans l'application que vous utilisez." + CRLF2 + "Par exemple, dans Word, Control tenu F, noté Ctrl+F," + CRLF + "lance la recherche d'une chaîne de caractères."
pg13d2 = "Les touches MAJ, Ctrl, Windows, Alt," + CRLF + "suivies par exemple d'une lettre," + CRLF + "sont utilisées dans les raccourcis clavier." + CRLF2 + "Chaque combinaison a une action différente." + CRLF2 + "Certains raccourcis clavier demandent 2 touches enfoncées," + CRLF + "avant de frapper brièvement la touche finale."
pg13d3 = "La touche AltGr est différente." + CRLF2 + "Quand on appuie sur AltGr, à droite de ESPACE," + CRLF + "c'est comme si on appuyait à la fois sur CONTROL et ALT." + CRLF2 + "Pour la combinaison Control+Alt+V," + CRLF + "Il suffit donc de maintenir enfoncé AltGr, puis de taper V."

pg13e1 = "Maintenant vous allez taper 3 caractères," + CRLF + "à l'aide d'une combinaison de touches," + CRLF + "démarrant par l'enfoncement de la touche AltGr." + CRLF2 + "Cette frappe ressemble à celle d'un raccourci." + CRLF + "Pourtant il s'agit seulement de caractères moins accessibles." + CRLF2 + "Dans une autre leçon, on verra une autre méthode," + CRLF + "avec le pavé numérique."
pg13e2 = "PRONONCIATION." + CRLF2 + "Le Dièse est prononcé Dièse." + CRLF2 + "La Barre-Oblique-Inversée, ou Contre-Oblique," + CRLF + "ou encore Antislash, en anglais Backslash," + CRLF + "est parfois prononcée Contre-Oblique." + CRLF2 + "La Barre-Oblique-Inversée est souvent utilisée pour préciser" + CRLF + "l'emplacement hiérarchique d'un dossier ou d'un fichier." + CRLF2 + "Le Acommercial, en anglais arobace," + CRLF + "est parfois prononcé at, mais plus souvent, arobase." + CRLF2 + "Le Acommercial est souvent utilisé dans les adresses internet."
pg13e3 = msgConseils + CRLF2 + "Appuyez et maintenez d'abord AltGr," + CRLF + "à droite de la barre ESPACE, avec le pouce droit." + CRLF2 + "Maintenez l'appui" + CRLF + "et frappez alors brièvement la touche souhaitée." + CRLF2 + "Enfin vous pouvez relâcher l'appui de AltGr." + CRLF2 + "Peu importe la position majuscule ou minuscule."
pg13e4 = "Le Dièse." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "et frappez brièvement avec l'annulaire gauche" + CRLF + "sur la touche du chiffre 3 ou guillemet." + CRLF2 + "La Barre-Oblique-Inversée." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "et frappez brièvement avec l'index sur le chiffre 8 ou souligné." + CRLF2 + "Le Arobase." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "et frappez brièvement avec l'annulaire sur le chiffre 0 ou A grave."

pg13f1 = "Maintenant vous allez taper des crochets ou des accolades," + CRLF + "à l'aide d'une combinaison de touches," + CRLF + "démarrant par l'enfoncement de la touche AltGr." + CRLF2 + "Cette frappe ressemble à celle d'un raccourci." + CRLF + "Pourtant il s'agit seulement de caractères moins accessibles."
pg13f2 = "PRONONCIATION." + CRLF2 + "Les crochets sont prononcés crochet gauche ou crochet droit." + CRLF2 + "Les accolades sont prononcées accolade gauche ou accolade droite." + CRLF2 + "Les crochets et les accolades ressemblent aux parenthèses," + CRLF + "mais en mathématiques," + CRLF + "ils expriment la hiérarchie des regroupements."
pg13f3 = msgConseils + CRLF2 + "Appuyez et maintenez d'abord AltGr," + CRLF + "à droite de la barre ESPACE, avec le pouce droit." + CRLF2 + "Maintenez l'appui" + CRLF + "et frappez alors brièvement la touche souhaitée." + CRLF2 + "Enfin vous pouvez relâcher l'appui de AltGr." + CRLF2 + "Peu importe la position majuscule ou minuscule."
'12/2011 texte plus court pour pg13f4
pg13f4 = "Le crochet gauche." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez brièvement avec l'auriculaire droit sur la touche du è." + CRLF2 + "Le crochet droit." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez brièvement avec l'auriculaire droit sur le tréma." + CRLF2 + "L'accolade gauche." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez brièvement avec l'auriculaire droit sur la touche du à." + CRLF2 + "L'accolade droite." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez brièvement avec l'auriculaire droit sur la touche du Dollar."

pg13g0 = "Dans les boîtes de dialogue," + CRLF + "certains champs sont difficiles à remplir." + CRLF2 + "Par exemple, les chemins des fichiers ou des dossiers " + CRLF + "exigent les deux-points et la Barre-Oblique-Inversée." + CRLF2 + "Certains noms comportent un ESPACE à respecter," + CRLF + "un caractère tiret ou un caractère souligné." + CRLF2 + "Voici un exercice d'entraînement."

' Leçon14
pg14a0 = "La dictée sera lue une seule fois." + CRLF + "Puis, la dictée sera faite phrase par phrase." + CRLF + "On passe à la lettre suivante dès la deuxième erreur." + CRLF2 + "Le mot suivant à taper sera prononcé." + CRLF2 + "Quand on vous demande ""à la ligne""," + CRLF + "appuyez simplement sur la touche Entrée." + CRLF2 + "Tapez la combinaison CONTROL+ESPACE pour RÉPÉTER le MOT." + CRLF2 + "Tapez Alt+ESPACE pour ÉPELER." + CRLF2 + "Tapez MAJ+ESPACE pour RÉPÉTER la FIN de la PHRASE."
pg14b0 = "Voici une deuxième dictée."
pg14c0 = "Voici une troisième dictée."

' Leçon15
pg15a0 = "La synthèse vocale va parler plus vite." + CRLF2 + "La dictée sera lue une seule fois." + CRLF + "Puis, la dictée sera faite phrase par phrase." + CRLF2 + "Le mot suivant à taper sera prononcé." + CRLF2 + "Tapez la combinaison CONTROL+ESPACE pour RÉPÉTER le MOT." + CRLF2 + "Tapez MAJ+ESPACE pour RÉPÉTER la PHRASE."
pg15b0 = "Voici une deuxième dictée." + CRLF2 + "La synthèse vocale va parler plus vite."
pg15c0 = "Voici une troisième dictée." + CRLF2 + "La synthèse vocale va parler encore plus vite."

' Leçon16
pg16a1 = "Maintenant vous allez apprendre à utiliser les touches" + CRLF + "du pavé numérique." + CRLF2 + "Le pavé numérique, c'est le groupe de 17 touches," + CRLF + "situé le plus à droite du clavier principal."
pg16a2 = msgConseils + CRLF2 + "Partez du bas complètement à droite du clavier." + CRLF + "Montez à la troisième rangée." + CRLF2 + "Une touche porte un point en relief, c'est le chiffre 5." + CRLF + "Placez le majeur sur cette touche 5." + CRLF2 + "Placez l'index à gauche sur le 4," + CRLF + "et placez l'annulaire à droite sur le 6." + CRLF2 + "Pour les 7, 8, 9, c'est la rangée au-dessus." + CRLF + "Pour 1, 2, 3, c'est la rangée au-dessous." + CRLF + "Pour le 0, c'est une grande touche encore plus bas sous le 1."
pg16a3 = "ATTENTION." + CRLF2 + "La touche dans le coin en haut et à gauche du pavé" + CRLF + "est une bascule." + CRLF2 + "C'est-à-dire, à chaque appui de cette touche spéciale," + CRLF + "on bascule toutes les touches du pavé" + CRLF + "du mode numérique au mode flèche," + CRLF + "ou du mode flèche au mode numérique." + CRLF2 + "Si la touche 5 ne répond pas, par exemple," + CRLF + "appuyez sur la bascule VERROUILLAGE-NUMÉRIQUE," + CRLF + "située en haut et à gauche du pavé."
pg16a4 = "RAPPEL." + CRLF + "Utilisez les aides F2 et F3." + CRLF2 + "Utilisez la touche bascule du coin en haut et à gauche du pavé," + CRLF + "si nécessaire."

pg16b1 = "Maintenant vous allez taper les 5 signes d'opérations," + CRLF + "avec le pavé numérique." + CRLF2 + "C'est-à-dire :" + CRLF + "la touche PLUS," + CRLF + "la touche MOINS prononcée TIRET," + CRLF + "la touche MULTIPLIER prononcée ÉTOILE ou ASTÉRISQUE," + CRLF + "la touche DIVISER prononcée SLASH ou BARRE-OBLIQUE," + CRLF + "et la touche POINT."
pg16b2 = msgHit + CRLF2 + "Le majeur reste sur le point en relief du 5." + CRLF2 + "La BARRE-OBLIQUE." + CRLF + "Partez du 5. Avec le majeur en extension," + CRLF + "montez à la deuxième rangée au-dessus." + CRLF2 + "L'ASTÉRISQUE." + CRLF + "Partez du 6. Avec l'annulaire," + CRLF + "montez à la deuxième rangée au-dessus."
pg16b3 = "Le TIRET." + CRLF + "Partez du 6. Avec l'auriculaire en extension," + CRLF + "montez à la deuxième rangée au-dessus et à droite." + CRLF2 + "Le PLUS." + CRLF + "Avec l'annulaire." + CRLF + "C'est la touche à droite du 6." + CRLF2 + "Le POINT." + CRLF + "Partez du 6. Avec l'annulaire," + CRLF + "descendez à la deuxième rangée au-dessous."

pg16c1 = "Chaque caractère possède un code chiffré pour le représenter." + CRLF + "Ce code est appelé nombre Ascii (prononcez ASKI)" + CRLF + "jusqu'à 3 chiffres, ou nombre ANSI à 4 chiffres." + CRLF + "Ce code peut dépendre des options linguistiques et régionales" + CRLF + "sélectionnées sur votre ordinateur," + CRLF + "notamment s'il dépasse la valeur 127." + CRLF + "Ceci permet de taper des caractères d'accès malcommode," + CRLF + "ou qui n'existent pas sur le clavier." + CRLF2 + "Il faudra tenir d'abord la touche ALT enfoncée," + CRLF + "avec le pouce gauche," + CRLF + "et taper le nombre Ascii ou ANSI," + CRLF + "avec le pavé numérique." + CRLF2 + "Le caractère apparaîtra quand vous relâcherez la touche ALT."
pg16c2 = "Maintenant vous allez taper les caractères suivants." + CRLF2 + "Le Dièse (#) : ALT tenu avec nombre 35." + CRLF + "La Barre-Oblique-Inversée (\): ALT tenu avec nombre 92." + CRLF + "Le Acommercial (@) : ALT tenu avec nombre 64." + CRLF + "Le Tilde (~): ALT tenu avec nombre 126." + CRLF + "Le € (euro) : ALT tenu avec nombre à 4 chiffres 0128."
pg16cm1 = "Maintenant vous allez taper les autres caractères suivants." + CRLF2 + "Le E aigu majuscule (É) : ALT tenu avec nombre 144." + CRLF + "Le œ (ligature du e collé dans l'o): ALT tenu avec nombre 0156." + CRLF + "Le ± (plus ou moins): ALT tenu avec nombre 0177." + CRLF + "Le ½ : ALT tenu avec nombre 0189."

pg16d1 = "Lorsque la touche VERROUILLAGE-NUMÉRIQUE," + CRLF + "dans le coin en haut et à gauche du pavé numérique," + CRLF + "est basculée sur le mode Flèche," + CRLF + "les touches du pavé deviennent des touches" + CRLF + "de direction du curseur." + CRLF2 + "Par exemple," + CRLF + "le 2 devient FLECHE-BAS," + CRLF + "le 4 devient FLECHE-GAUCHE," + CRLF + "le 6 devient FLECHE-DROITE," + CRLF + "le 8 devient FLECHE-HAUT."
pg16d2 = "Les autres touches de direction sont les suivantes :" + CRLF2 + "DÉBUT. Avec l'index. Au-dessus du 4." + CRLF2 + "FIN. Avec l'index. En-dessous du 4." + CRLF2 + "PAGE-PRÉCÉDENTE. Avec l'annulaire. Au-dessus du 6." + CRLF2 + "PAGE-SUIVANTE. Avec l'annulaire. Au-dessous du 6."
pg16dm1 = "Il y a aussi une touche bascule INSERTION-REMPLACEMENT," + CRLF + "une touche SUPPRESSION, et une touche Entrée." + CRLF2 + "INSERTION. Avec l'index." + CRLF + "Partez du 4. Descendez de 2 rangées en-dessous." + CRLF2 + "SUPPRESSION. Avec l'annulaire." + CRLF + "Partez du 6. Descendez de 2 rangées en-dessous." + CRLF2 + "Entrée. Avec l'auriculaire. Allez au coin en bas à droite." + CRLF2 + "Pour cet exercice, n'utilisez pas les touches équivalentes," + CRLF + "situées à gauche du pavé numérique."

' Leçon17
pg17a0 = "Maintenant vous allez utiliser les lettres, les chiffres," + CRLF + "et les ponctuations du clavier principal."
pg17b0 = "Maintenant vous allez utiliser toutes les touches" + CRLF + "du clavier général."
pg17c0 = "Maintenant vous allez taper les caractères et les combinaisons," + CRLF + "en allant plus vite."
pg17d0 = "Maintenant vous allez utiliser tout le pavé numérique," + CRLF + "en allant plus vite."

' Leçon18
pg18a0 = "Maintenant vous allez taper des mots accentués." + CRLF2 + "Attention." + CRLF + "Les substantifs sont au singulier." + CRLF + "Les verbes sont à l'infinitif."
pg18b0 = "Maintenant vous allez taper des mots avec doubles consonnes." + CRLF2 + "Il y aura des majuscules," + CRLF + "et une prononciation difficile."
pg18c0 = "Maintenant vous allez taper des mots dont les terminaisons" + CRLF + "sont courantes." + CRLF2 + "Attention." + CRLF + "La synthèse vocale ne va pas épeler les mots," + CRLF + "sauf si vous faites une erreur." + CRLF2 + "Si vous faites 2 erreurs dans un mot," + CRLF + "ce mot sera à nouveau proposé."
pg18d0 = "Maintenant vous allez taper rapidement des mots," + CRLF + "dont la prononciation est voisine." + CRLF2 + "La synthèse vocale parlera vite, sans épeler."
pg18e0 = "Maintenant vous allez taper des instructions," + CRLF + "dont se servent les programmeurs."

' Leçon19
pg19a0 = "La synthèse vocale va parler plus vite." + CRLF2 + "Le texte sera lu une seule fois." + CRLF + "Puis, la dictée sera faite phrase par phrase." + CRLF2 + "Le mot suivant à taper sera prononcé." + CRLF2 + "Tapez la combinaison CONTROL+ESPACE pour RÉPÉTER le MOT." + CRLF2 + "Tapez sur Alt+ESPACE pour ÉPELER." + CRLF2 + "Tapez MAJ+ESPACE pour RÉPÉTER la FIN de la PHRASE."
pg19b0 = "Voici un deuxième texte." + CRLF2 + "La synthèse vocale va parler plus vite."
pg19c0 = "Voici un troisième texte." + CRLF2 + "La synthèse vocale va parler très vite."
pg19d0 = "Voici un quatrième texte." + CRLF2 + "La synthèse vocale va parler extrèmement vite."

' Menu Editor
meFichier = "&Fichier"
meQuitter_bm = "&Quitter"
meOptions = "&Options"
meStandard = "Niveau &Standard"
mePersonnalisé = "Niveau &Personnalisé"
meDebExpliNormal = "Débit des explications &Normal"
meDebExpliRapide = "Débit des explications &Rapide"
meDebGenLent = "Débit général &Lent"
meDebGenMoyen = "Débit général &Moyen"
meDebGenVite = "Débit général &Vite"
meBipClassique = "Bip &Classique"
meBipDifférent = "Bip &Différent"
meAide = "&Aide"
meAideGénérale = "Aide générale"
meAideMémoire = "Aide-Mémoire"
meEnseignant = "Aide pour l'&Enseignant"
meSonorisation = "Aide pour la &Vocalisation"
meAproposde = "À &Propos de"
meReset = "Redémarrer à la prem&ière leçon"
meRestart = "Redémarrer à la prem&ière leçon"

meNoZoom = "Sans z&oom"  ' 12/2011
meWithZoom = "Avec &zoom"
meBasicColors = "Couleurs &basiques"
meOtherColors = "A&utres couleurs"

' Suite
repjawscountry = "\settings\fra\"
Module_routines.inits
End Sub


' *************  HELP_F2  ******  TRADUIRE SEULEMENT à DROITE DE .text4.Text =  ************
Public Sub help_f2(leçon)
With leçon

' Détecter le Alt255 final éventuel
If Right(.text1.Text, 1) = " " Then
    lt1 = Len(.text1.Text) - 1
Else
    lt1 = Len(.text1.Text)
End If

' Reset
.text4.Visible = False
Call Sleep(10) 'attention, pas trop long sinon pb quand AltGr reste enfoncé
pp = ii + 1 - ff
If pp <= 0 Then pp = 1

' Lettres d'1 seul caractère minuscule
If Mid(.text1.Text, pp, 1) = "a" Then .text4.Text = " a.  Minuscule.   Auriculaire gauche.   Rangée au-dessus de Q, et légèrement à droite.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "b" Then .text4.Text = " b.  Minuscule.   Index gauche.        Rangée au-dessous de G, et à droite.  Voir leçon 3B."
If Mid(.text1.Text, pp, 1) = "c" Then .text4.Text = " c.  Minuscule.   Majeur gauche.       Rangée au-dessous de D, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "d" Then .text4.Text = " d.  Minuscule.   Majeur gauche.       Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "e" Then .text4.Text = " e.  Minuscule.   Majeur gauche.       Rangée au-dessus de D, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "f" Then .text4.Text = " f.  Minuscule.   Index gauche.        Rangée de départ, point en relief.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "g" Then .text4.Text = " g.  Minuscule.   Index gauche.         Rangée de départ, à droite du F.  Voir leçon 2B."
If Mid(.text1.Text, pp, 1) = "h" Then .text4.Text = " h.  Minuscule.   Index droit.         Rangée de départ, à gauche du J.  Voir leçon 2B."
If Mid(.text1.Text, pp, 1) = "i" Then .text4.Text = " i.  Minuscule.   Majeur droit.        Rangée au-dessus de K, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "j" Then .text4.Text = " j.  Minuscule.   Index droit.         Rangée de départ, point en relief.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "k" Then .text4.Text = " k.  Minuscule.   Majeur droit.        Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "l" Then .text4.Text = " l.  Minuscule.   Annulaire droit.     Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "m" Then .text4.Text = " m.  Minuscule.   Auriculaire droit.   Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "n" Then .text4.Text = " n.  Minuscule.   Index droit.         Rangée au-dessous de J, et à gauche.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "o" Then .text4.Text = " o.  Minuscule.   Annulaire droit.     Rangée au-dessus de L, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "p" Then .text4.Text = " p.  Minuscule.   Auriculaire droit.   Rangée au-dessus de M, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "q" Then .text4.Text = " q.  Minuscule.   Auriculaire gauche.  Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "r" Then .text4.Text = " r.  Minuscule.   Index gauche.        Rangée au-dessus de F, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "s" Then .text4.Text = " s.  Minuscule.   Annulaire gauche.    Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "t" Then .text4.Text = " t.  Minuscule.   Index gauche.        Rangée au-dessus de F, et à droite.  Voir leçon 2E."
If Mid(.text1.Text, pp, 1) = "u" Then .text4.Text = " u.  Minuscule.   Index droit.         Rangée au-dessus de J, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "v" Then .text4.Text = " v.  Minuscule.   Index gauche.        Rangée au-dessous de F, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "w" Then .text4.Text = " w.  Minuscule.   Auriculaire gauche.  Rangée au-dessous de Q, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "x" Then .text4.Text = " x.  Minuscule.   Annulaire gauche.    Rangée au-dessous de S, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "y" Then .text4.Text = " y.  Minuscule.   Index droit.         Rangée au-dessus de H, et à gauche.  Voir leçon 2E."
If Mid(.text1.Text, pp, 1) = "z" Then .text4.Text = " z.  Minuscule.   Annulaire gauche.    Rangée au-dessus de S, et à gauche.  Voir leçon 2C."

' Lettres d'1 seul caractère majuscule
If Mid(.text1.Text, pp, 1) = "A" Then .text4.Text = " A.  Majuscule.   Auriculaire gauche.  Rangée au-dessus de Q, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "B" Then .text4.Text = " B.  Majuscule.   Index gauche.        Rangée au-dessous de G, et à droite.  Voir leçon 3B."
If Mid(.text1.Text, pp, 1) = "C" Then .text4.Text = " C.  Majuscule.   Majeur gauche.       Rangée au-dessous de D, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "D" Then .text4.Text = " D.  Majuscule.   Majeur gauche.       Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "E" Then .text4.Text = " E.  Majuscule.   Majeur gauche.       Rangée au-dessus de D, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "F" Then .text4.Text = " F.  Majuscule.   Index gauche.        Rangée de départ, point en relief.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "G" Then .text4.Text = " G.  Majuscule.   Index gauche.         Rangée de départ, à droite du F.  Voir leçon 2B."
If Mid(.text1.Text, pp, 1) = "H" Then .text4.Text = " H.  Majuscule.   Index droit.         Rangée de départ, à gauche du J.  Voir leçon 2B."
If Mid(.text1.Text, pp, 1) = "I" Then .text4.Text = " I.  Majuscule.   Majeur droit.        Rangée au-dessus de K, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "J" Then .text4.Text = " J.  Majuscule.   Index droit.         Rangée de départ, point en relief.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "K" Then .text4.Text = " K.  Majuscule.   Majeur droit.        Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "L" Then .text4.Text = " L.  Majuscule.   Annulaire droit.     Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "M" Then .text4.Text = " M.  Majuscule.   Auriculaire droit.   Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "N" Then .text4.Text = " N.  Majuscule.   Index droit.         Rangée au-dessous de J, et à gauche.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "O" Then .text4.Text = " O.  Majuscule.   Annulaire droit.     Rangée au-dessus de L, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "P" Then .text4.Text = " P.  Majuscule.   Auriculaire droit.   Rangée au-dessus de M, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "Q" Then .text4.Text = " Q.  Majuscule.   Auriculaire gauche.  Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "R" Then .text4.Text = " R.  Majuscule.   Index gauche.        Rangée au-dessus de F, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "S" Then .text4.Text = " S.  Majuscule.   Annulaire gauche.    Rangée de départ.  Voir leçon 2A."
If Mid(.text1.Text, pp, 1) = "T" Then .text4.Text = " T.  Majuscule.   Index gauche.        Rangée au-dessus de F, et à droite.  Voir leçon 2E."
If Mid(.text1.Text, pp, 1) = "U" Then .text4.Text = " U.  Majuscule.   Index droit.         Rangée au-dessus de J, et à gauche.  Voir leçon 2C."
If Mid(.text1.Text, pp, 1) = "V" Then .text4.Text = " V.  Majuscule.   Index gauche.        Rangée au-dessous de F, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "W" Then .text4.Text = " W.  Majuscule.   Auriculaire gauche.  Rangée au-dessous de Q, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "X" Then .text4.Text = " X.  Majuscule.   Annulaire gauche.    Rangée au-dessous de S, et à droite.  Voir leçon 3A."
If Mid(.text1.Text, pp, 1) = "Y" Then .text4.Text = " Y.  Majuscule.   Index droit.         Rangée au-dessus de H, et à gauche.  Voir leçon 2E."
If Mid(.text1.Text, pp, 1) = "Z" Then .text4.Text = " Z.  Majuscule.   Annulaire gauche.    Rangée au-dessus de S, et à gauche.  Voir leçon 2C."

' Chiffres
If Module_routines.IsNumLockOn() = "False" Then
    If Mid(.text1.Text, pp, 1) = "1" Then .text4.Text = " Chiffre 1.  Majuscule.  Auriculaire gauche.  2 rangées au-dessus de Q, et légèrement à gauche.  Voir leçon 10A."
    If Mid(.text1.Text, pp, 1) = "2" Then .text4.Text = " Chiffre 2.  Majuscule.  Auriculaire gauche.  2 rangées au-dessus de Q, et légèrement à droite.  Voir leçon 10A."
    If Mid(.text1.Text, pp, 1) = "3" Then .text4.Text = " Chiffre 3.  Majuscule.  Annulaire gauche.   2 rangées au-dessus de S, et légèrement à droite.  Voir leçon 10A."
    If Mid(.text1.Text, pp, 1) = "4" Then .text4.Text = " Chiffre 4.  Majuscule.  Majeur gauche.   2 rangées au-dessus de D, et légèrement à droite.  Voir leçon 10A."
    If Mid(.text1.Text, pp, 1) = "5" Then .text4.Text = " Chiffre 5.  Majuscule.  Index gauche.   2 rangées au-dessus de F, et légèrement à droite.  Voir leçon 10B."
    If Mid(.text1.Text, pp, 1) = "6" Then .text4.Text = " Chiffre 6.  Majuscule.  Index gauche.   2 rangées au-dessus de F, et en extension à droite.  Voir leçon 10B."
    If Mid(.text1.Text, pp, 1) = "7" Then .text4.Text = " Chiffre 7.  Majuscule.  Index droit.   2 rangées au-dessus de J, et nettement à gauche.  Voir leçon 10B."
    If Mid(.text1.Text, pp, 1) = "8" Then .text4.Text = " Chiffre 8.  Majuscule.  Index droit.   2 rangées au-dessus de J, et légèrement à droite.  Voir leçon 10C."
    If Mid(.text1.Text, pp, 1) = "9" Then .text4.Text = " Chiffre 9.  Majuscule.  Majeur droit.   2 rangées au-dessus de K, et légèrement à droite.  Voir leçon 10C."
    If Mid(.text1.Text, pp, 1) = "0" Then .text4.Text = " Chiffre 0.  Majuscule.  Annulaire droit.   2 rangées au-dessus de L, et légèrement à droite.  Voir leçon 10C."
Else
    If Mid(.text1.Text, pp, 1) = "1" Then .text4.Text = " Chiffre 1.  Pavé Mode Numérique.  Index droit.   En-dessous du 5 et à gauche.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "2" Then .text4.Text = " Chiffre 2.  Pavé Mode Numérique.  Majeur droit.   En-dessous du 5.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "3" Then .text4.Text = " Chiffre 3.  Pavé Mode Numérique.  Annulaire droit.   En-dessous du 5 et à droite.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "4" Then .text4.Text = " Chiffre 4.  Pavé Mode Numérique.  Index droit.   A gauche du 5.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "5" Then .text4.Text = " Chiffre 5.  Pavé Mode Numérique.  Majeur droit.   Au centre du pavé, touche avec relief.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "6" Then .text4.Text = " Chiffre 6.  Pavé Mode Numérique.  Annulaire droit.   A droite du 5.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "7" Then .text4.Text = " Chiffre 7.  Pavé Mode Numérique.  Index droit.    Au-dessus du 5 et à gauche.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "8" Then .text4.Text = " Chiffre 8.  Pavé Mode Numérique.  Index droit.   Au-dessus du 5.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "9" Then .text4.Text = " Chiffre 9.  Pavé Mode Numérique.  Annulaire droit.   Au-dessus du 5 et à droite.  Voir leçon 16A."
    If Mid(.text1.Text, pp, 1) = "0" Then .text4.Text = " Chiffre 0.  Pavé Mode Numérique.  Index droit.   2 rangées en-dessous du 5, et à gauche.  Voir leçon 16A."
End If

' Ponctuations et signes
If Mid(.text1.Text, pp, 1) = " " Then .text4.Text = " ESPACE.   Pouce gauche, ou pouce droit.  Grande barre au-devant du clavier principal.  Voir leçon 1A."
If Mid(.text1.Text, pp, 1) = "²" Then .text4.Text = " Au Carré.  Minuscule.  Auriculaire gauche.  2 rangées au-dessus de Q, en extension à gauche.  Voir leçon 9A."
If Mid(.text1.Text, pp, 1) = "&" Then .text4.Text = " Et Commercial.  Minuscule.  Auriculaire gauche.  2 rangées au-dessus de Q, et à gauche.  Voir leçon 9A."
If Mid(.text1.Text, pp, 1) = """" Then .text4.Text = "Guillemet.  Minuscule.   Annulaire gauche.    2 rangées au-dessus de S, et légèrement à droite.  Voir leçon 9B."
If Mid(.text1.Text, pp, 1) = "'" Then .text4.Text = " Apostrophe.  Minuscule.  Majeur gauche.   2 rangées au-dessus de D, et légèrement à droite.  Voir leçon 9B."
If Mid(.text1.Text, pp, 1) = "(" Then .text4.Text = " Parenthèse gauche. Minuscule. Index gauche.  2 rangées au-dessus de F, et légèrement à droite.  Voir leçon 9B."
If Mid(.text1.Text, pp, 1) = ")" Then .text4.Text = " Parenthèse droite.    Minuscule.    Auriculaire droit.     2 rangées au-dessus de M.  Voir leçon 9C."
If Mid(.text1.Text, pp, 1) = "_" Then .text4.Text = " Souligné.    Minuscule.     Index droit.      2 rangées au-dessus de J.  Voir leçon 9C."
If Mid(.text1.Text, pp, 1) = "," Then .text4.Text = " Virgule.  Minuscule.    Index droit.       Rangée au-dessous de J, et à droite.  Voir leçon 8B."
If Mid(.text1.Text, pp, 1) = "?" Then .text4.Text = " Point d'interrogation.    Majuscule.    Index droit.    Rangée au-dessous de J, et à droite.  Voir leçon 8C."
If Mid(.text1.Text, pp, 1) = ";" Then .text4.Text = " Point-Virgule.    Minuscule.    Majeur droit.    Rangée au-dessous de K, et à droite.  Voir leçon 8B."
If Mid(.text1.Text, pp, 1) = ":" Then .text4.Text = " Deux-Points.    Minuscule.    Annulaire droit.     Rangée au-dessous de L, et à droite.  Voir leçon 8B."
If Mid(.text1.Text, pp, 1) = "!" Then .text4.Text = " Point d'exclamation.   Minuscule.   Auriculaire droit.   Rangée au-dessous de M, et à droite.  Voir leçon 8B."
If Mid(.text1.Text, pp, 1) = "§" Then .text4.Text = " Section.   Majuscule.   Auriculaire droit.   Rangée au-dessous de M, et à droite.  Voir leçon 8C."
If Mid(.text1.Text, pp, 1) = "<" Then .text4.Text = " Inférieur.   Minuscule.   Auriculaire gauche.   Rangée au-dessous de Q, et légèrement à gauche.  Voir leçon 8F."
If Mid(.text1.Text, pp, 1) = ">" Then .text4.Text = " Supérieur.   Majuscule.   Auriculaire gauche.   Rangée au-dessous de Q, et légèrement à gauche.  Voir leçon 8F."
If Mid(.text1.Text, pp, 1) = "%" Then .text4.Text = " PourCent.    Majuscule.    Auriculaire droit.    Rangée de départ, à droite de M.  Voir leçon 8G."
If Mid(.text1.Text, pp, 1) = "$" Then .text4.Text = " Dollar.   Minuscule.   Auriculaire droit.   Rangée au-dessus de M, et en extension à droite.  Voir leçon 8G."
If Mid(.text1.Text, pp, 1) = "£" Then .text4.Text = " Livre.    Majuscule.    Auriculaire droit.    Rangée au-dessus de M, et en extension à droite.  Voir leçon 8G."
If Mid(.text1.Text, pp, 1) = "µ" Then .text4.Text = " Mu, ou Micro.    Majuscule.   Auriculaire droit.    Rangée de départ, 2 touches à droite de M.  Voir leçon 8G."
If Mid(.text1.Text, pp, 1) = "ç" Then .text4.Text = " c cédille.   Minuscule.    Majeur droit.    2 rangées au-dessus de K.  Voir leçon 9D."
If Mid(.text1.Text, pp, 1) = "°" Then .text4.Text = " Degré.    Majuscule.    Auriculaire droit.    2 rangées au-dessus de M.  Voir leçon 9D."
If Mid(.text1.Text, pp, 1) = "=" Then .text4.Text = " Égal.    Minuscule.    Auriculaire droit.    2 rangées au-dessus de M, nettement à droite.  Voir leçon 9D."
If Mid(.text1.Text, pp, 1) = "^" Then .text4.Text = " Circonflexe.  Minuscule.  Auriculaire droit.  Au-dessus de M, et à droite.  Tapez ensuite la voyelle.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "¨" Then .text4.Text = " Tréma.  Majuscule.  Auriculaire droit.  Au-dessus de M, et à droite.  Tapez ensuite la voyelle.  Voir leçon 8E."

' Caractères 4 opérations et POINT selon numpad ?
If Mid(.text1.Text, pp, 1) = "+" Then .text4.Text = " Plus.    Majuscule.    Auriculaire droit.    2 rangées au-dessus de M, nettement à droite.  Voir leçon 16B."
If Mid(.text1.Text, pp, 1) = "-" Then .text4.Text = " Tiret.    Minuscule.    Index gauche.    2 rangées au-dessus de F, et en extension à droite.  Voir leçon 16B."
If Mid(.text1.Text, pp, 1) = "*" Then .text4.Text = " Astérisque.    Minuscule.   Auriculaire droit.    Rangée de départ, 2 touches à droite de M.  Voir leçon 16B."
If Mid(.text1.Text, pp, 1) = "/" Then .text4.Text = " Barre-Oblique.    Majuscule.    Annulaire droit.      Rangée au-dessous de L, et à droite.  Voir leçon 16B."
If Mid(.text1.Text, pp, 1) = "." Then .text4.Text = " Point.    Majuscule.    Majeur droit.    Rangée au-dessous de K, et à droite.  Voir leçon 16B."

' caractères commandés par AltGr
If numpad <= 0 Then
    If Mid(.text1.Text, pp, 1) = "~" Then .text4.Text = " Tilde.  Pouce droit sur AltGr tenu, puis avec l'auriculaire gauche, touche du chiffre 2.  Voir leçon 13E."
    If Mid(.text1.Text, pp, 1) = "#" Then .text4.Text = " Dièse.  Pouce droit sur AltGr tenu, puis avec l'annulaire gauche, tapez sur le guillemet.  Voir leçon 13E."
    If Mid(.text1.Text, pp, 1) = "{" Then .text4.Text = " Accolade gauche.  Pouce droit sur AltGr tenu, puis avec le majeur gauche, touche du chiffre 4.  Voir leçon 13F."
    If Mid(.text1.Text, pp, 1) = "[" Then .text4.Text = " Crochet gauche.  Pouce droit sur AltGr tenu, puis avec l'index gauche, touche du chiffre 5.  Voir leçon 13F."
    If Mid(.text1.Text, pp, 1) = "|" Then .text4.Text = " Barre-Verticale.  Pouce droit sur AltGr tenu, puis avec l'index gauche, touche du chiffre 6." & "                 "
    If Mid(.text1.Text, pp, 1) = "`" Then .text4.Text = " Accent grave.  Pouce droit sur AltGr tenu, puis tapez sur la touche du chiffre 7." & "                 "
    If Mid(.text1.Text, pp, 1) = "\" Then .text4.Text = " Barre-Oblique-Inversée.  Pouce droit sur AltGr tenu, puis tapez sur le souligné, touche chiffre 8.  Voir leçon 13E."
    If Mid(.text1.Text, pp, 1) = "@" Then .text4.Text = " A Commercial.  Pouce droit sur AltGr tenu, puis avec l'annulaire droit, tapez sur le a grave.  Voir leçon 13E."
    If Mid(.text1.Text, pp, 1) = "]" Then .text4.Text = " Crochet droit.  Pouce droit sur AltGr tenu, puis avec l'auriculaire droit, touche de la parenthèse.  Voir leçon 13F."
    If Mid(.text1.Text, pp, 1) = "}" Then .text4.Text = " Accolade droite.  Pouce droit sur AltGr tenu, puis avec l'auriculaire droit, touche du signe égal.  Voir leçon 13F."
End If

' caractères commandés par Alt+nombre-Ascii-Ansi
If numpad >= 1 Then
    If Mid(.text1.Text, pp, 1) = "#" Then .text4.Text = " Dièse.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 35.  Voir leçon 16C."
    If Mid(.text1.Text, pp, 1) = "@" Then .text4.Text = " A Commercial.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 64.  Voir leçon 16C."
    If Mid(.text1.Text, pp, 1) = "\" Then .text4.Text = " Barre-Oblique-Inversée.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 92.  Voir leçon 16C."
    If Mid(.text1.Text, pp, 1) = "~" Then .text4.Text = " Tilde.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 126.  Voir leçon 16C."
End If
If Mid(.text1.Text, pp, 1) = "É" Then .text4.Text = " E aigu Majuscule.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 144.  Voir leçon 16C."
If Mid(.text1.Text, pp, 1) = "€" Then .text4.Text = " €.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 0128.  Voir leçon 16C."
If Mid(.text1.Text, pp, 1) = "±" Then .text4.Text = " ±.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 0177.  Voir leçon 16C."
If Mid(.text1.Text, pp, 1) = "½" Then .text4.Text = " ½.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 0189.  Voir leçon 16C."
If Mid(.text1.Text, pp, 1) = "œ" Then .text4.Text = " œ.  Pouce gauche sur Alt tenu, puis au pavé numérique tapez le nombre 0156.  Voir leçon 16C."

' Lettres accentuées
If Mid(.text1.Text, pp, 1) = "à" Then .text4.Text = " a grave.   Minuscule.   Annulaire droit.    2 rangées au-dessus de L.  Voir leçon 8H."
If Mid(.text1.Text, pp, 1) = "é" Then .text4.Text = " e aigu.   Minuscule.   Auriculaire gauche.    2 rangées au-dessus de Q, et légèrement à droite.  Voir leçon 8H."
If Mid(.text1.Text, pp, 1) = "è" Then .text4.Text = " e grave.   Minuscule.     Index droit.    2 rangées au-dessus de J, et vers la gauche.  Voir leçon 8H."
If Mid(.text1.Text, pp, 1) = "ù" Then .text4.Text = " u grave.   Minuscule.   Auriculaire droit.    Rangée de départ, à droite de M.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "â" Then .text4.Text = " a circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et à droite du M, avant le a.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ê" Then .text4.Text = " e circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et à droite du M, avant le e.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "î" Then .text4.Text = " i circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et à droite du M, avant le i.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ô" Then .text4.Text = " o circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et à droite du M, avant le o.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "û" Then .text4.Text = " u circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et à droite du M, avant le u.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ä" Then .text4.Text = " a tréma.    a minuscule, tréma majuscule.   Tapez le tréma au-dessus et à droite du M, avant le a.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ë" Then .text4.Text = " e tréma.    e minuscule, tréma majuscule.   Tapez le tréma au-dessus et à droite du M, avant le e.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ï" Then .text4.Text = " i tréma.    i minuscule, tréma majuscule.   Tapez le tréma au-dessus et à droite du M, avant le i.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ö" Then .text4.Text = " o tréma.    o minuscule, tréma majuscule.   Tapez le tréma au-dessus et à droite du M, avant le o.  Voir leçon 8E."
If Mid(.text1.Text, pp, 1) = "ü" Then .text4.Text = " u tréma.    u minuscule, tréma majuscule.   Tapez le tréma au-dessus et à droite du M, avant le u.  Voir leçon 8E."

End With

' Lancer la suite dans une autre procedure (sinon message procedure too large)
help_f2_suite leçon
End Sub


' ************  HELP_F2_SUITE  ******  TRADUIRE SEULEMENT à DROITE DE .text4.Text =  *********
Public Sub help_f2_suite(leçon)
With leçon

' Raccourcis
If UCase(Left(.text1.Text, Len(vvMaj) + 1)) = UCase(vvMaj) & "+" Then
    .text4.Text = " MAJUSCULE tenue enfoncée. Auriculaire. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvMaj) - 1) & ".  Voir leçon 13D."
End If
If UCase(Left(.text1.Text, Len(vvControl) + 1)) = UCase(vvControl) & "+" Then
    .text4.Text = " CONTROL tenu enfoncé. Auriculaire. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvControl) - 1) & ".  Voir leçon 13D."
End If
If UCase(Left(.text1.Text, Len(vvAlt) + 1)) = UCase(vvAlt) & "+" Then
    .text4.Text = " ALT tenu enfoncé. Pouce gauche. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvAlt) - 1) & ".  Voir leçon 13D."
End If
If UCase(Left(.text1.Text, Len(vvControl) + Len(vvMaj) + 2)) = UCase(vvControl) & "+" & UCase(vvMaj) & "+" Then
    .text4.Text = " CONTROL et MAJ tenus enfoncés. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvControl) - Len(vvMaj) - 2) & ".  Voir leçon 13D."
End If
If UCase(Left(.text1.Text, Len(vvControl) + Len(vvAlt) + 2)) = UCase(vvControl) & "+" & UCase(vvAlt) & "+" Then
    .text4.Text = " CONTROL et ALT tenus enfoncés. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvControl) - Len(vvAlt) - 2) & ".  Voir leçon 13D."
End If
If UCase(Left(.text1.Text, Len(vvMaj) + Len(vvAlt) + 2)) = UCase(vvMaj) & "+" & UCase(vvAlt) & "+" Then
    .text4.Text = " MAJUSCULE et ALT tenus enfoncés. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvMaj) - Len(vvAlt) - 2) & ".  Voir leçon 13D."
End If

' Textes de plusieurs caractères, attention, placer ces lignes de code après celles dédiées à 1 caractère et après celles des raccourcis
If UCase(Left(.text1.Text, lt1)) = UCase(vvEspace) Then .text4.Text = " ESPACE.   Pouce gauche ou pouce droit.    Grande barre au-devant du clavier principal.  Voir leçon 1A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvControl) Then .text4.Text = " CONTROL. Auriculaire gauche ou droit.   Coins gauche et droit en bas du clavier principal.  Voir leçon 1C."
If Left(.text1.Text, lt1) = vvControlGauche Then .text4.Text = " CONTROL-GAUCHE.  Auriculaire gauche.  Touche du coin gauche en bas du clavier principal.  Voir leçon 1C."
If Left(.text1.Text, lt1) = vvControlDroit Then .text4.Text = " CONTROL-DROIT.   Auriculaire droit.   Touche du coin droit en bas du clavier principal.  Voir leçon 1C."
If Left(.text1.Text, lt1) = vvWindowsGauche Then .text4.Text = " WINDOWS-GAUCHE.   Auriculaire gauche.   2 rangées en-dessous de Q.  Voir leçon 13A."
If Left(.text1.Text, lt1) = vvWindowsDroit Then .text4.Text = " WINDOWS-DROIT.    Auriculaire droit.    2 rangées en-dessous de M.  Voir leçon 13A."
If Left(.text1.Text, lt1) = vvMenuContextuel Then .text4.Text = " MENU-CONTEXTUEL.   Auriculaire droit.   2 rangées en-dessous de M, en extension à droite.  Voir leçon 13A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvAlt) Then .text4.Text = " ALT.   Pouce gauche.    Touche à gauche de la barre ESPACE.  Voir leçon 13A."
If Left(.text1.Text, lt1) = vvAltGr Then .text4.Text = " AltGr.   Pouce droit.    Touche à droite de la barre ESPACE.  Voir leçon 13A."
If Left(.text1.Text, lt1) = vvAltOuAltGr Then .text4.Text = " ALT Pouce gauche.   AltGr Pouce droit.   A gauche ou à droite de la barre ESPACE.  Voir leçon 13A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvÉchap) Then .text4.Text = " ÉCHAP.    Auriculaire gauche.     Touche au coin à gauche en haut du clavier.  Voir leçon 1A."
If Left(.text1.Text, lt1) = vvVerrouillageMajuscules Then .text4.Text = " VERROUILLAGE-MAJUSCULES.   Auriculaire gauche.   Rangée de départ à gauche de Q.  Voir leçon 8A."
If Left(.text1.Text, lt1) = vvVerrouillageNumérique Then .text4.Text = " VERROUILLAGE-NUMÉRIQUE. Index droit. Coin en haut et à gauche, dans le pavé numérique.  Voir leçon 16A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvMaj) Then .text4.Text = vvMaj
If Left(.text1.Text, lt1) = vvMajGauche Then .text4.Text = " MAJ-GAUCHE.   Auriculaire gauche.     Rangée au-dessous de Q et très à gauche.  Voir leçon 8A."
If Left(.text1.Text, lt1) = vvMajDroit Then .text4.Text = " MAJ-DROIT.   Auriculaire droit.     Rangée au-dessous de M en extension à droite.  Voir leçon 8A."
If LCase(Left(.text1.Text, lt1)) = LCase(vvRetourArrière) Then .text4.Text = " RETOUR-Arrière.    Annulaire droit.    Clavier principal, coin en haut et à droite.  Voir leçon 13B."
If UCase(Left(.text1.Text, lt1)) = UCase(vvTab) Then .text4.Text = " TABULATION.   Auriculaire gauche.   Au-dessus du Q, et très à gauche.  Voir leçon 13B."
If Left(.text1.Text, lt1) = vvTabulationAvant Then .text4.Text = " TABULATION-AVANT. Minuscule. Auriculaire gauche. Au-dessus du Q, et très à gauche.  Voir leçon 13B."
If Left(.text1.Text, lt1) = vvTabulationArrière Then .text4.Text = " TABULATION-Arrière. Majuscule. Auriculaire gauche. Au-dessus du Q, et très à gauche.  Voir leçon 13B."
If Left(.text1.Text, lt1) = vvDébut Then .text4.Text = " DÉBUT, ou HOME.     Majeur droit.      Milieu rangée supérieure, groupe des 6.  Voir leçon 12B."
If Left(.text1.Text, lt1) = vvFin Then .text4.Text = " FIN, ou END.    Majeur droit.     Milieu rangée inférieure, groupe des 6.  Voir leçon 12B."
If Left(.text1.Text, lt1) = vvImpression Then .text4.Text = " IMPRESSION.   Majeur droit.   Rangée au-dessus du groupe des 6, au-dessus de INSERTION.  Voir leçon 12D."
If Left(.text1.Text, lt1) = vvArrêtDéfil Then .text4.Text = " ArrêtDéfil.   Index droit.   Rangée au-dessus du groupe des 6, au-dessus de DÉBUT.  Voir leçon 12D."
If UCase(Left(.text1.Text, lt1)) = vvPause Then .text4.Text = " PAUSE.   Annulaire droit.   Rangée au-dessus du groupe des 6, et de PAGE-PRÉCÉDENTE.  Voir leçon 12D."

' Touches de fonction
If Left(.text1.Text, lt1) = "F1" Then .text4.Text = "  F1.   Annulaire gauche.   3 rangées au-dessus de S, nettement à gauche.  Voir leçon 1B."
If Left(.text1.Text, lt1) = "F2" Then .text4.Text = "  F2.   Annulaire gauche.   3 rangées au-dessus de S, légèrement à droite.  Voir leçon 1B."
If Left(.text1.Text, lt1) = "F3" Then .text4.Text = "  F3.   Majeur gauche.   3 rangées au-dessus de D, légèrement à droite.  Voir leçon 1B."
If Left(.text1.Text, lt1) = "F4" Then .text4.Text = "  F4.   Index gauche.   3 rangées au-dessus de F, légèrement à droite.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F5" Then .text4.Text = "  F5.   Index gauche.   3 rangées au-dessus de F, en extension à droite.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F6" Then .text4.Text = "  F6.   Index droit.   3 rangées au-dessus de J, en extension.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F7" Then .text4.Text = "  F7.   Majeur droit.   3 rangées au-dessus de K, en extension.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F8" Then .text4.Text = "  F8.   Annulaire droit.   3 rangées au-dessus de L, en extension.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F9" Then .text4.Text = "  F9.   Auriculaire droit.   3 rangées au-dessus de M, en extension.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F10" Then .text4.Text = "  F10.   Auriculaire droit.   3 rangées au-dessus de M, en extension à droite.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F11" Then .text4.Text = "  F11.   Auriculaire droit.   3 rangées au-dessus de M, en extension très à droite.  Voir leçon 13C."
If Left(.text1.Text, lt1) = "F12" Then .text4.Text = "  F12.   Auriculaire droit.   3 rangées au-dessus de M, en extension extrème à droite.  Voir leçon 13C."

' Touches de même nom au clavier principal et au pavé numérique
If numpad = 0 Then
    If UCase(Left(.text1.Text, lt1)) = UCase(vvEntrée) Then .text4.Text = " Entrée.   Auriculaire droit.    Grande touche sur le bord droit du clavier principal.  Voir leçon 1A."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvInsertion) Then .text4.Text = " INSERTION REMPLACEMENT BASCULE. Index droit. Coin en haut à gauche, groupe des 6.  Voir leçon 12A."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvSuppression) Then .text4.Text = " SUPPRESSION.     Index droit.     Coin en bas à gauche, groupe des 6.  Voir leçon 12A."
    If Left(.text1.Text, lt1) = vvFlecheGauche Then .text4.Text = " FLECHE-GAUCHE. Index droit. Première touche en bas à droite du clavier principal.  Voir leçon 1B."
    If Left(.text1.Text, lt1) = vvFlecheBas Then .text4.Text = " FLECHE-BAS. Majeur droit. Deuxième touche en bas à droite du clavier principal.  Voir leçon 1B."
    If Left(.text1.Text, lt1) = vvFlecheDroite Then .text4.Text = " FLECHE-DROITE. Annulaire droit. Troisième touche en bas à droite du clavier principal.  Voir leçon 1B."
    If Left(.text1.Text, lt1) = vvFlecheHaut Then .text4.Text = " FLECHE-HAUT.    Majeur droit.     Au-dessus de la touche FLECHE-BAS.  Voir leçon 1B."
    If Left(.text1.Text, lt1) = vvPagePrécédente Then .text4.Text = " PAGE-PRÉCÉDENTE.   Annulaire droit.    Coin en haut à droite, groupe des 6.  Voir leçon 12C."
    If Left(.text1.Text, lt1) = vvPageSuivante Then .text4.Text = " PAGE-SUIVANTE.    Annulaire droit.    Coin en bas à droite, groupe des 6.  Voir leçon 12C."
End If
If numpad = 1 Or numpad = -1 Then
    If UCase(Left(.text1.Text, lt1)) = UCase(vvEntrée) Then .text4.Text = " Entrée.   Pavé Numérique.    Auriculaire droit.   Coin à droite et en bas.  Voir leçon 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvInsertion) Then .text4.Text = " INSERTION REMPLACEMENT BASCULE. Pavé Numérique. Index droit. Coin en bas à gauche.  Voir leçon 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvSuppression) Then .text4.Text = " SUPPRESSION.  Pavé Numérique. Annulaire droit. 2 rangées au-dessous et à droite du 5.  Voir leçon 16D."
    If Left(.text1.Text, lt1) = vvFlecheGauche Then .text4.Text = " FLECHE-GAUCHE. Pavé Numérique, mode flèche. Index droit. Première touche à gauche du 5.  Voir leçon 16D."
    If Left(.text1.Text, lt1) = vvFlecheBas Then .text4.Text = " FLECHE-BAS. Pavé Numérique, mode flèche. Majeur droit. En-dessous du 5.  Voir leçon 16D."
    If Left(.text1.Text, lt1) = vvFlecheDroite Then .text4.Text = " FLECHE-DROITE. Pavé Numérique, mode flèche. Annulaire droit. Première touche à droite du 5.  Voir leçon 16D."
    If Left(.text1.Text, lt1) = vvFlecheHaut Then .text4.Text = " FLECHE-HAUT.  Pavé numérique, mode flèche.  Majeur droit.  Au-dessus du 5.  Voir leçon 16D."
    If Left(.text1.Text, lt1) = vvPagePrécédente Then .text4.Text = " PAGE-PRÉCÉDENTE. Pavé Numérique, mode flèche. Annulaire droit. Au-dessus et à droite du 5.  Voir leçon 16D."
    If Left(.text1.Text, lt1) = vvPageSuivante Then .text4.Text = " PAGE-SUIVANTE. Pavé Numérique, mode flèche. Annulaire droit. Au-dessous et à droite du 5.  Voir leçon 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvDébut) Then .text4.Text = " DÉBUT.   Pavé Numérique, mode flèche.   Index droit.   Au-dessus et à gauche du 5.  Voir leçon 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvFin) Then .text4.Text = " FIN.   Pavé Numérique, mode flèche.   Index droit.   Au-dessous et à gauche du 5.  Voir leçon 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvPlus) Then .text4.Text = " PLUS.   Pavé Numérique.    Auriculaire droit.   Au-dessus et à droite du 6.  Voir leçon 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvTiret) Then .text4.Text = " TIRET ou MOINS.   Pavé Numérique.    Auriculaire droit.   Coin en haut à droite.  Voir leçon 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvMoins) Then .text4.Text = " TIRET ou MOINS.   Pavé Numérique.    Auriculaire droit.   Coin en haut à droite.  Voir leçon 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvBarreOblique) Then .text4.Text = " BARRE-OBLIQUE.   Pavé Numérique.    Majeur droit.   2 rangées au-dessus du 5.  Voir leçon 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvAstérisque) Then .text4.Text = " ASTÉRISQUE.   Pavé Numérique.    Annulaire droit.   2 rangées au-dessus du 6.  Voir leçon 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvPoint) Then .text4.Text = " POINT.   Pavé Numérique.    Annulaire droit.   2 rangées en-dessous du 6.  Voir leçon 16B."
End If

' Cas particulier "à la ligne" (Alt255 devant pour visibilité)
If .text3.Text = " " & vvAlaligne Then .text4.Text = " à la ligne, ou Entrée.   Auriculaire droit.    Grande touche sur le bord droit du clavier principal.  Voir leçon 1A."

' SUPPRIMER la MENTION "VOIR leçonxxA" sauf en mode F3 aide-mémoire (mais variable avecf3 inutilisable)
If avecf2 = 1 Then
    On Error Resume Next
    .text4.Text = Left(.text4.Text, Len(.text4.Text) - Len("Voir leçon 19A."))  'élimine les mentions voir leçon xx
End If

' Set
.text4.SelStart = 0
.text4.SelLength = Len(.text4.Text)
.text4.Visible = True

' En dehors du mode aide-mémoire
If avecf3 = 0 Then
    keyinhibit = 2
    f2link = 1
End If
End With
End Sub
