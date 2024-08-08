Attribute VB_Name = "Module_global"
'Ce logiciel libre est disponible sous licence GNU/GPL,
'dont une copie se trouvera dans le fichier gpl.txt,
'avec une traduction fran�aise non officielle gpl-fr.txt.

Option Explicit
' ***************  CONTIENT DECLARATIONS, MAIN, puis ROUTINES � TRADUIRE *********************
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
Global nivo, nivoRep, nom, nom_temp, vpath, vfile, vfileresults As String 'niveau_le�ons, niveau_nom_r�pertoire_immuable, nom_utilisateur, nom_temporaire, chemin_programme, chemin_avec_nom_fichier, chemin_avec_r�sultats
Global bannerThanks, bannerNosell, bannerFunction, bannerVersion, bannerCopyright, bannerAuthorAddress As String  'remerciements, non_vendu, fonction_du_logiciel, banni�re_app_version, banni�re_copyright, banni�re_auteur
Global bannerPrincipal, bannerMenu, bannerLe�on As String  'mot_principal, mot_menu, mot_le�on

Global biplevel, debexplilevel, debexplivalue, msgtext0 As String 'niveau_sonore_des_bips, d�bit_des_explications, string_rep�re_d�bit, texte_msg_msgform
Global msgtext1(150), msgtext2(150), datatext1(150) As String 'texte_msg_text1(iter), texte_msg_boite_de_dialogue(iter), donn�es_text1(ligne)
Global repj(9), repjexe(9) As String 'rep_Config_jaws_si_plusieurs_versions_jaws, rep_Exe_Jaws
Global fautesur(150), fautecourante, fauteprec As String 'cara_demand�_faut�(iwrong)
Global currentline, currentmenuline, scorecourant As String 'ligne_compl�te_en_cours_incluant_retour-chariot, ligne_courante_menu, score_courant
Global country, repjawscountry, repjaws, repUsers, repNVDA, repjawsjsb, ujaws(9) As String 'langue, rep_jaws_settings, test version Windows, r�p_jaws_pour_jcf, idem_valid, idem_r�p_name, idem_fichier_jss, unit�_r�p_jaws
Global perso_methode, pressez, pressez_entr�e, pressez_quit, pressez_suivant, pressez_Le�onSuivante, pressez_pr�c�dent, pressez_F2, pressez_touche As String 'msg_methode_pour_cr�er_perso_files, msg_pressez_espace_ou_Entr�e, msg_pressez_entr�e, msg_pressez_quit, msg_pressez_suivant, msg_pressez_pr�c�dent, msg_si_F2_inutile, msg_minimum
Global pressez_basic, pressez_ligne As String 'msg_pressez_de_base, msg_pressez_ligne_par_ligne
Global repjawsnames, repjawsfra As String 'nom_reps_jaws_type_c:\jaws401, nom_rep_settings_fra_courant
Global exo_courant As String 'exo_courant.txt (attention pas d'autre variable sur cette ligne d�claration)
Global debgenlevel, debgenvalue As String 'd�bit_reper�_par_string, d�bit_session_info

Global msgAide, msgAideF3, msgAtt, msgWith, msgBienvenue, msgEnter, msgLevel, msgNofic, msgSpeedExp, msgSpeedGen, msgUser As String
Global msgBienvFaudra, msgPage, msgPressEnter, msgSonori, msgTapez, msgTapez2, msgTapezTouche As String
Global msgBienvUsername, msgBienvRedo, msgBienvRep, msgBienvRepeat, msgBienvRetape, msgRelaunch As String
Global msgMots, msgCommandes, msgCommandesEn, msgExoFautes, msgExoIdem, msgExoSuivant, msgLevelIs, msgStandard, msgPersonnalis� As String
Global msgPrincPour, msgPrincDansniveau, msgPrincContenu, msgPrincTermin� As String
Global msgFormPressez, msgFormRecommencer, msgFormVous�tiez As String
Global msgAvec, msgMotsEn, msgSecondes, msgFautesSur, msgPressezF1, msgR�ussi�, msgR�ussi�100, msgPourcent, msgPourcentSeulement, msgTranslator As String
Global msgTypeClavier, msgScore, msgCommandesDispo, msgF1Aide, msgF2Loc, msgF3AM, msgEspace, msgCtrlEspace, msgAltEspace, msgMajEspace, msgSortir, msgSortir2, msgSortir3, msgAltF4 As String
Global msgAurevoir, msgNoficSono, msgEntr�eContinuer, msgD�tect�, msgUserIs, msgSpeedExpIs, msgSpeedGenIs, msgBipsAre, msgChoisir, msgChoisissez, msgOptions As String
Global msgBip, msgBipComment 'avril 2008
Global msg�chap, msg�chap2, msg�chap3, msgContinuer, msgPr�c�dent, msgSuivant, msgQuitter, msgQuitterAM, msgQuitterMP, msgQuitterVers As String
Global msgF1F2F3, msgLaLe�on, msgEstTermin�e, msgR�sultats, msgMotsMinute, msgSes, msgR�ussite As String
Global msgClassique, msgDiff�rent, msgDict�e, msgD�bit, msgExpli, msgLent, msgMoyen, msgVite, msgNormal, msgRapide As String
Global msgNiveauStandard, msgNiveauPersonnalis�, msgConseils, msgHit, msgNoSono, msgKeyboard As String
Global msgRestartTitle, msgRestart, msgRestartCmd, meRestart, msgResetTitle, msgReset, meReset As String

Global vvAccentGrave, vvAccoladeDroite, vvAccoladeGauche, vvAlaligne, vvAlt, vvAltGr, vvAltOuAltGr, vvApostrophe, vvArr�tD�fil, vvAst�risque As String
Global vvBarreOblique, vvBarreObliqueInvers�e, vvBarreVerticale, vvControl, vvControlDroit, vvControlGauche, vvCrochetDroit, vvCrochetGauche, vvCtrl, vvD�but, vvDeuxPoints, vvDiviser As String
Global vvEchap, vv�chap, vvEntr�e, vvEspace, vv�toile, vvFin, vvFlecheBas, vvFlecheDroite, vvFlecheGauche, vvFlecheHaut, vvGuillemet, vvImpression, vvInf�rieur, vvInsertion As String
Global vvMaj, vvMajDroit, vvMajGauche, vvMajuscule, vvMenuContextuel, vvMinuscule, vvMoins, vvMultiplier As String
Global vvPagePr�c�dente, vvPageSuivante, vvParenth�seDroite, vvParenth�seGauche, vvPause, vvPlus, vvPoint, vvPointExclamation, vvPointInterrogation, vvPointVirgule As String
Global vvRetourArriere, vvRetourArri�re, vvSansNom, vvSoulign�, vvSup�rieur, vvSuppression, vvTab, vvTabulationAvant, vvTabulationArri�re, vvTilde, vvTiret As String
Global vvVerrouillageMajuscules, vvVerrouillageNum�rique, vvVirgule, vvWindowsDroit, vvWindowsGauche As String

Global nboccur(150) As Integer 'nb_occurences_sur_cara_demand�_faut�(iwrong)
Global rrs(9), rrt(9) As Integer 'variables_selon_le_type_de_version_Jaws
Global ii, iiold, ll, llold, nbli, zz As Integer 'indice_courant_text2, old_ii, length_ligne_text1, old_length, nb_lignes_text1, length_currentline
Global iistart, iistartp, iistop, iistop0, iistop1, iistop2, iistop3, iistop9, iistopf As Integer 'indices_d�b_phrase, indices_fin_phrase (d�tecte le "." ou "!" ou "?")
Global jj, kk, mm, nn, pp, qq As Integer 'variables_de_boucle
Global iwrong, iwrongbis, iwrongbismax, iwrongl, irecur As Integer 'nb_fautes, nb_fautes_biss�es, nb_max, nb_fautes_dans_ligne, nb_fois_qu'on_recommence
Global iwrongCR, iwrongCRmax As Integer 'nb_fautes_pour_retour_charriot_�_la_ligne
Global nbcaras, nbonscaras, lt1 As Integer 'nb_caras_exo, nbonscaras, length_r�elle_text1
Global iter, iiante, iiprec As Integer 'indice_avancement, indice_ant�rieur, indice_pr�c�dent
Global cadencecara, cadencemot, cadenceligne As Integer 'cadences_demande_suivante
Global elapsed, elapsedtot As Integer 'temps_pass�_pour_ligne_courante, temps_total
Global nbmots, keyforce As Integer 'nombre_mots, force_keycode
Global echapbis, echapbismax, echapoff As Integer 'echap_biss�e, max_de_echapbis, echap_offset
Global fsize, fsizedefault, le�onfontsize, le�onfontsize5 As Integer 'font_size, font_size_default_msgform, fontsize_normale_pour_le�on, idem_le�on5
Global numpad As Integer 'pav�_num�rique:-1=�vite_lock__0=non__1=uniquement__2=oui_ainsi_que_chiffres_clavier_principal
Global firstmove As Integer 'first_cursor_move_in_msgform

Global indif, KeyAscii, Keycode, KeyExpect, ShiftExpect As Byte 'indiff�rent_majuscule_minuscule, r�ponse_KeyPress, r�ponse_KeyUp_or_down, r�ponse_attendue_code_touche, r�ponse_attendue_shift_touche
Global KeyFirst, KeySecond, KeyThird As Byte 'Codes successifs g�n�rs par Jaws pour les touches Echap, Alt, Control
Global derligne, nextle�on, typele�on, stopscroll As Byte 'derni�re_ligne, passer_�_le�on_suivante, typele�on_1_2_3_7_14, stop_scroll_results
Global f2link As Byte 'Enchaine_apr�s_f2
Global keyinhibit, t2inhibit, mcinhibit, fullscreeninhibit As Byte 'inhibit_keyup_after_msgbox, inhibit_event_change_text2, inhibit_after_menu-contextuel, inhibit_fullscreen_display
Global concatf, timevalid, winstop As Byte 'mode_o�_text1_s'ajoute, autorise_affich&calcul_vitesse, renforce_stop_touche_windows
Global notab, sonocara, timein, timeover, timeout As Byte 'touche_tab_sans_effet, sel_cara_pour_sonoriser, compteur_temps, temps_de_r�ponse_d�pass�, timeout=1_pour_quitter
Global menucount, numle�on, numexo, nbexo, nblines As Byte 'nb_choix_menu_courant, num�ro_le�on, num�ro_exercice, nb_d'exercices_pour_la_le�on, nb_lignes_le�on
Global f1msgform, inexo, iwait, ff, incomplet As Byte 'F1_valid_msgform, entr�e_le�on, var_wait, flag_faute, le�on_o�_manque_au_moins_un_r�sultat_pctok_d'un_exercice
Global numindex, tempnum As Byte 'num�ro_index_courant_menu, temporaire
Global pctok(50, 10) As Integer '25li_Standard_puis_25_Personnalis� ; 1�re_col_pct_r�ussite_moyenne_numle�on ; cols_r�ussite_exos ; der_col_visualise_numle�on
Global vitok(50, 10) As Integer '25li_Standard_puis_25_Personnalis� ; 1�re_col_vitesse_moyenne_numle�on ; cols_vitesse_exos ; der_col_visualise_numle�on
Global pctt, pct1, erepeat, lrepeat, wrepeat As Byte 'pourcentage_r�ussite_exercice, pourcent_limite_pour passer_au_suivant, �pelle, line_repeat, word_repeat
Global nfree As Byte 'num�ros_libres_pour_ouvrir_fichiers
Global msgf As Byte 'sortie_de_msgform_pour_la_fonction_msghb_avec_0=�chap,1=Entr�e,2=R�p�ter
Global passb As Byte 'mode_de_sortie_de_la_ligne_avec_0=Suivante_en_fin_de_ligne,1=Suivante_d�s_1�re_erreur,2=R�p�ter_la_ligne_si_2erreurs
Global avecf2, avecf3, nobip, noF1, noechapF1, pasdepoint As Byte 'en_mode_aidef2, en_mode_aidef3, pas_de_bip, pas_aide_F1_fautes, pas_echap_msgform_de_F1_vers_menu, chargement_non_limit�_par_un_"."
Global espacevalid, nodoublesono, noalt As Byte 'accepte_espace_pour_r�p�ter_dans_le�on1, �vite_double_sono_cas_particuliers, �vite_r�p�ter_fin_phrase_par_Alt
Global pagenum, pagemax As Byte 'num�ro_page (0_interdit_num�ros_page), derni�re_page
Global forcepause, consult, quitactive, altf4 As Byte '�vite_de_skipper_score, mode_consulter_r�sultats, quitquit_en_cours, quit_par_Alt+F4
Global bascule, bipinhibit, quitF2, FullScreenSwitch As Byte 'bascule, bip_inhibit, quit_msg_F2_inutile, passer_en_plein_�cran

Global le�on_courante, menu_courant, menu_suivant As Object 'exemples: le�on3, menu_le�on3, menu_le�on4
Global emax, startline, starttop As Date 'elapsed_max_allowed, temps_d�part_nouvelle_ligne, temps_top_d�part
Global bNumLockState, bCapsLockState, bScrollLockState As Boolean
Global scrw, frmw, scrh, frmh, vcomp, zfactor As Variant 'd�tecteurs de r�solution d'�cran
Global fbc, fbc_default, fbc_quit, fbc_f1, f_orange, f_rouge, f_rougevif, f_gris, f_violet, f_bleuclair, f_bleufonc�, f_vert As Variant 'font_back_colors
Global ffc, ffc_default, ffc_quit, ffc_f1, f_noirfonc�, f_noirgris, f_noirnoir As Variant 'font_fore_color

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
Global meFichier, meQuitter_bm, meOptions, meStandard, mePersonnalis�, meDebExpliNormal, meDebExpliRapide, meDebGenLent, meDebGenMoyen, meDebGenVite As String
Global meBipClassique, meBipDiff�rent, meAide, meAideG�n�rale, meAideM�moire, meEnseignant, meSonorisation, meAproposde As String

'12/2011 zoom et couleurs
Global zoomfactor, zoomvalue As Variant
Global zoomlevel, msgNoZoom, msgWithZoom, msgDisplay As String
Global colorslevel, msgBasicColors, msgOtherColors As String
Global meNoZoom, meWithZoom, meBasicColors, meOtherColors As String
Global f_grisp�le, f_grisfonc�, f_vertsombre, f_marronsombre, f_jaunevif, f_jaunetr�svif, f_violetsombre, f_bleuvif, f_violetvif, f_vertp�le, f_vertvif, f_noirpresque, f_blanc, f_orangeclair As Variant

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





'Touches � pb, voir les le�ons 1 et 13 lorsque Jaws rajoute des codes parasites sur ces touches :
'Control 17
'Alt-Droit 17+18
'Alt-Gauche 18
'Win-Gauche 91, traitement sp�cifique pour casser la priorit� donn�e par Windows � la touche Win
'Win-Droit 92, traitement sp�cifique pour casser la priorit� donn�e par Windows � la touche Win
'Menu-Contextuel 93, n�cessite de faire suivre par Echap
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


' **************** MAIN1 **** TRADUIRE � DROITE DU SIGNE �GAL ******************************
Public Sub main1()

' Variables messages
bannerAuthorAddress = "herve.beranger@neuf.fr"
bannerCopyright = "  Copyleft 2008-2019 GNU/GPL"
bannerFunction = "logiciel d'apprentissage du clavier"
bannerLe�on = "Le�on"
bannerMenu = "Menu"
bannerNosell = "Ce logiciel libre est disponible sous licence GNU/GPL," & CRLF & "dont une copie se trouvera dans le fichier c:\ApprentiClavier\gpl.txt," & CRLF & "avec une traduction fran�aise non officielle gpl-fr.txt." & CRLF2 & "ApprentiClavier est diffus� gratuitement sur les sites" & CRLF & "www.apprenticlavier.com, http://apprenticlavier.wifeo.com, http://www.winaide.net"
bannerPrincipal = "Principal"
bannerThanks = "L'auteur adresse ses vifs remerciements" & CRLF & "aux pr�curseurs b�n�voles de l'association Club Micro Son," & CRLF & "Robert Agro, Thierry Bertrand, Alain Rousseau."
bannerVersion = " ApprentiClavier Version 1.10"
clavierType = "AZERTY fran�ais (FRANCE)"
country = "Langue : fran�aise."
msgAide = "AIDE."
msgAideF3 = "Mode Aide-M�moire, �chap pour sortir"
msgAltEspace = "Alt+ESPACE : pour �PELER."
msgAltF4 = "Alt+F4 : pour QUITTER IMM�DIATEMENT."
msgAtt = "Attention�"
msgAurevoir = "AU REVOIR."
msgAvec = "avec  "
msgBienvenue = "Bienvenue dans "
msgBienvFaudra = "Il faudra souvent appuyer sur la touche Entr�e :"
msgBienvRedo = "  ERREUR.  Recommen�ons."
msgBienvRep = " Une deuxi�me fois, TAPEZ votre NOM."
msgBienvRepeat = "Par s�curit�, une deuxi�me fois, TAPEZ votre NOM :"
msgBienvRetape = " TAPEZ votre NOM, ou sur Entr�e."
msgBienvUsername = "Tapez votre NOM, ou tapez simplement sur Entr�e :"
msgBip = "Bip sur fautes : " 'avril 2008
msgBipComment = "Voir ""Options"" dans la barre des menus." 'avril 2008
msgBipsAre = "Les BIPS sur les fautes sont r�gl�s sur "
msgChoisir = "Fl�ches et Entr�e : pour CHOISIR."
msgChoisissez = "  Choisissez."
msgClassique = "Classique"
msgCommandes = " commandes.       "
msgCommandesEn = "  commandes en  "
msgCommandesDispo = "Apr�s ce MESSAGE, les COMMANDES DISPONIBLES seront :"
msgConseils = "CONSEILS PRATIQUES."
msgContinuer = "&Continuer (Entr�e)"
msgCtrlEspace = "CONTROL+ESPACE : pour R�P�TER le MOT."
msgD�bit = " Le d�bit g�n�ral sera r�gl� sur "
msgD�tect� = "D�TECT�."
msgDict�e = "Dict�e "
msgDiff�rent = "Diff�rent"
msg�chap = "(�chap)"
msg�chap2 = "(2 fois �chap)"
msg�chap3 = "(3 fois �chap)"
msgEnter = "  Remontez le long du clavier principal par la droite." + CRLF + "  Vous rencontrez une grande touche verticale." + CRLF + "  C'est la touche Entr�e."
msgEntr�eContinuer = "Touche Entr�e : pour CONTINUER."
msgEspace = "ESPACE : pour R�P�TER la TOUCHE demand�e."
msgEstTermin�e = "  est termin�e."
msgExoFautes = "Exercice autour des fautes."
msgExoIdem = "Voulez-vous REFAIRE L'EXERCICE ? "
msgExoSuivant = "Voulez-vous FAIRE L'EXERCICE SUIVANT ? "
msgExpli = " Les explications seront donn�es avec un d�bit "
msgF1F2F3 = " F1=Aide g�n�rale       F2=Description de la touche       F3=Aide-M�moire "
msgF1Aide = "Touche F1 : pour l'AIDE."
msgF2Loc = "Touche F2 : pour LOCALISER la TOUCHE demand�e."
msgF3AM = "Touche F3 : pour l'AIDE-M�MOIRE."
msgFautesSur = " FAUTES SUR "
msgFormPressez = " Pressez " + CRLF + " ESPACE pour R�P�TER la PAGE d'explications," + CRLF + " puis les fl�ches pour R�P�TER chaque LIGNE," + CRLF2 + " ou �chap pour SORTIR," + CRLF + " ou Entr�e pour CONTINUER."
msgFormRecommencer = "Veuillez RECOMMENCER."
msgFormVous�tiez = " Vous �tiez dans une page d'explications."
msgHit = "CONSIGNES DE FRAPPE."
msgKeyboard = "    Clavier : "
msgLaLe�on = "La le�on  "
msgLent = "Lent"
msgLevel = "         Niveau : "
msgLevelIs = "Le niveau des le�ons est "
msgMajEspace = "MAJ+ESPACE : pour R�P�TER la FIN de la PHRASE."
msgMots = " mots.   "
msgMotsEn = "  mots en  "
msgMotsMinute = "mots-minute"
msgMoyen = "Moyen"
msgNofic = "Pas de fichier  "
msgNoficSono = "Pas de fichier de vocalisation JAWS."
msgNoSono = "Consultez ""Aide"" dans la barre sup�rieure."
msgNormal = "Normal"
msgOptions = "Touche Alt, puis O, OPTIONS : NIVEAU, D�BIT, BIP, ZOOM, COULEURS."  ' 12/2011
msgPage = "Page "
msgPersonnalis� = "Personnalis�"
msgPourcent = " pourcent"
msgPourcentSeulement = " pourcent seulement"
msgPr�c�dent = "&Pr�c�dent"
msgPressEnter = " Appuyez maintenant sur la touche Entr�e ! "
msgPressezF1 = "Pressez F1 pour vous exercer sur les fautes."
msgPrincContenu = "Voici le contenu du fichier "
msgPrincDansniveau = "  dans le niveau "
msgPrincPour = "  pour  "
msgPrincTermin� = "  TERMIN�.  "
msgQuitter = "&Quitter "
msgQuitterAM = "&Quitter l'aide-m�moire "
msgQuitterMP = "&Quitter vers Menu Principal  "
msgQuitterVers = "&Quitter vers "
msgRapide = "Rapide"
msgRelaunch = "Il va falloir RELANCER ApprentiClavier."
msgReset = "Pour effacer tous les r�sultats d'un utilisateur et red�marrer � la premi�re le�on, vous devrez entrer dans le menu principal." + CRLF2 + "Puis vous taperez Alt+O Options, puis I pour ""Red�marrer.""" + CRLF2 + "Appuyez maintenant sur la touche Entr�e."
msgResetTitle = "Information."
msgRestart = "Voulez-vous Red�marrer � la premi�re le�on, et effacer tous les r�sultats de l'utilisateur?"
msgRestartCmd = CRLF2 + "Pressez sur Entr�e pour red�marrer, ou sur �chap pour annuler."
msgRestartTitle = "Confirmer le red�marrage."
msgR�sultats = "    R�SULTATS DE "
msgR�ussi� = "Cet exercice est r�ussi � "
msgR�ussi�100 = "nous enregistrons ici un taux de r�ussite de 100 pourcent. "
msgR�ussite = " exercices r�ussis en moyenne � "
msgScore = "Score : "
msgSecondes = " secondes."
msgSes = "Ses "
msgSonori = "   Vocalisation : "
msgSortir = "Alt+Q ou Touche �chap : pour SORTIR."
msgSortir2 = "Alt+Q ou Touche �chap 2 fois : pour SORTIR."
msgSortir3 = "Alt+Q ou Touche �chap 3 fois : pour SORTIR."
msgSpeedExp = "    Avec d�bits : explications="
msgSpeedExpIs = "Le D�BIT des EXPLICATIONS est "
msgSpeedGen = ", g�n�ral="
msgSpeedGenIs = "Le D�BIT G�N�RAL est r�gl� sur "
msgStandard = "Standard"
msgSuivant = "&Suivant"
msgTapez = "Tapez"
msgTapez2 = "Tapez :"
msgTapezTouche = "Tapez la touche :"
msgTranslator = "La traduction en anglais est due � Herv� B�ranger."
msgTypeClavier = "Cette version est pour clavier " & clavierType & ", " & country
msgUser = "   Utilisateur : "
msgUserIs = "L'utilisateur est "
msgVite = "Vite"
msgWith = "Vous allez taper des mots�"
pressez_basic = "      Pressez ESPACE pour R�P�TER,      " + CRLF + "       ou Entr�e pour CONTINUER.        " + CRLF
pressez = CRLF2 + pressez_basic
pressez_ligne = CRLF2 + "   Pressez Fl�che-Bas pour R�P�TER LIGNE PAR LIGNE, " + CRLF + pressez_basic
pressez_entr�e = CRLF3 + "      Pressez Entr�e pour CONTINUER.        " + CRLF
pressez_quit = CRLF2 + BLANCS12 + "  ATTENTION.             " + CRLF + BLANCS6 + "    Vous allez QUITTER         " + CRLF + BLANCS6 + "     ApprentiClavier.          " + CRLF2 + BLANCS12 + "   Pressez               " + CRLF + "      Entr�e pour NE PAS QUITTER,    " + CRLF + "        ou �CHAP  pour QUITTER.      "
pressez_Le�onSuivante = CRLF2 + "     Pressez ESPACE pour R�P�TER ce message," + CRLF + "     ou Entr�e pour PASSER � la le�on SUIVANTE."
pressez_suivant = CRLF2 + "     Pressez ESPACE pour R�P�TER ce message," + CRLF + "     ou Entr�e pour PASSER � l'EXERCICE SUIVANT."
pressez_pr�c�dent = CRLF2 + "     Pressez ESPACE pour R�P�TER ce message," + CRLF + "     ou Entr�e pour REFAIRE l'EXERCICE."
pressez_F2 = CRLF + " F2 ne peut pas vous renseigner," + CRLF + " car aucune TOUCHE n'est DEMAND�E." + CRLF2
pressez_touche = "     Appuyez sur une touche."
perso_methode = CRLF2 + "L'enseignant peut cr�er chaque fichier texte," + CRLF + "en copiant celui du sous-dossier le�ons\standard" + CRLF + "vers le sous-dossier le�ons\personnalis�," + CRLF + "puis en le modifiant � son gr�." + CRLF2 + "L'utilisateur choisira le niveau Standard, ou Personnalis�," + CRLF + "en se pla�ant dans n'importe quel menu," + CRLF + "puis en tapant sur Alt, puis sur O, Options ;" + CRLF + "puis S, Standard, ou P, Personnalis�."

msgNoZoom = "Sans zoom" ' 12/2011
msgWithZoom = "Avec zoom"
msgBasicColors = "Couleurs basiques"
msgOtherColors = "Autres couleurs"
msgDisplay = "      Affichage : "

' Variables noms de touches
vvAccentGrave = "ACCENT GRAVE"
vvAccoladeDroite = "ACCOLADE DROITE"
vvAccoladeGauche = "ACCOLADE GAUCHE"
vvAlaligne = "� la ligne"
vvAlt = "ALT"
vvAltGr = "AltGr"
vvAltOuAltGr = "ALT ou AltGr"
vvApostrophe = "APOSTROPHE"
vvArr�tD�fil = "Arr�tD�fil"
vvAst�risque = "AST�RISQUE"
vvBarreOblique = "BARRE-OBLIQUE"
vvBarreObliqueInvers�e = "BARRE-OBLIQUE-INVERS�E"
vvBarreVerticale = "BARRE-VERTICALE"
vvControl = "CONTROL"
vvControlDroit = "CONTROL-DROIT"
vvControlGauche = "CONTROL-GAUCHE"
vvCrochetDroit = "CROCHET DROIT"
vvCrochetGauche = "CROCHET GAUCHE"
vvCtrl = "CTRL"
vvD�but = "D�BUT"
vvDeuxPoints = "DEUX-POINTS"
vvDiviser = "DIVISER"
vvEchap = "ECHAP"  'Pr�f�rer vv�chap avec l'accent
vv�chap = "�CHAP"
vvEntr�e = "Entr�e"
vvEspace = "ESPACE"
vv�toile = "�TOILE"
vvFin = "FIN"
vvFlecheBas = "FLECHE-BAS"
vvFlecheDroite = "FLECHE-DROITE"
vvFlecheGauche = "FLECHE-GAUCHE"
vvFlecheHaut = "FLECHE-HAUT"
vvGuillemet = "GUILLEMET"
vvImpression = "IMPRESSION"
vvInf�rieur = "INF�RIEUR"
vvInsertion = "INSERTION"
vvMaj = "MAJ"
vvMajDroit = "MAJ-DROIT"
vvMajGauche = "MAJ-GAUCHE"
vvMajuscule = "  MAJUSCULE"
vvMenuContextuel = "MENU-CONTEXTUEL"
vvMinuscule = "  MINUSCULE"
vvMoins = "MOINS"
vvMultiplier = "MULTIPLIER"
vvPagePr�c�dente = "PAGE-PR�C�DENTE"
vvPageSuivante = "PAGE-SUIVANTE"
vvParenth�seDroite = "Parenth�se DROITE"
vvParenth�seGauche = "Parenth�se GAUCHE"
vvPause = "PAUSE"
vvPlus = "PLUS"
vvPoint = "POINT"
vvPointExclamation = "POINT d'EXCLAMATION"
vvPointInterrogation = "POINT d'INTERROGATION"
vvPointVirgule = "POINT-VIRGULE"
vvRetourArriere = "RETOUR-ARRIERE"
vvRetourArri�re = "RETOUR-Arri�re"
vvSansNom = "SansNom"
vvSoulign� = "SOULIGN�"
vvSup�rieur = "SUP�RIEUR"
vvSuppression = "SUPPRESSION"
vvTab = "TAB"
vvTabulationAvant = "TABULATION-AVANT"
vvTabulationArri�re = "TABULATION-Arri�re"
vvTilde = "TILDE"
vvTiret = "TIRET"
vvVerrouillageMajuscules = "VERROUILLAGE-MAJUSCULES"
vvVerrouillageNum�rique = "VERROUILLAGE-NUM�RIQUE"
vvVirgule = "VIRGULE"
vvWindowsDroit = "WINDOWS-DROIT"
vvWindowsGauche = "WINDOWS-GAUCHE"

' Variables pages d'explications
' Pr�sentation
pgia1 = bannerVersion & ", " + CRLF + bannerCopyright + ", Herv� B�ranger," + CRLF + "   " + bannerAuthorAddress + "." + CRLF2 + "ATTENTION." + CRLF + "� la fin de chaque page d'explications, veuillez utiliser :" + CRLF2 + " - la touche Entr�e pour CONTINUER," + CRLF + " - la touche ESPACE pour R�P�TER," + CRLF + " - la touche �chap  pour SORTIR." + CRLF2 + "La le�on 1 expliquera ces 3 touches." + CRLF2 + "Pour sortir, la touche �chap se trouve" + CRLF + "dans le coin en haut et � gauche du clavier."
'pgia2 = "ApprentiClavier permet l'apprentissage " + CRLF + "du clavier d'un ordinateur." + CRLF2 + "Il est d�riv� du logiciel CLAVSON, mis au point par" + CRLF + "l'association Club Micro Son, notamment par" + CRLF + "Robert Agro, Thierry Bertrand, Alain Rousseau." + CRLF2 + msgTypeClavier + CRLF2 + msgTranslator
pgia2 = "ApprentiClavier permet l'apprentissage d'un clavier classique" + CRLF + "d'ordinateur,type 105 touches avec pav� num�rique." + CRLF2 + msgTypeClavier
pgia3 = "Les enseignants peuvent modifier les exercices." + CRLF2 + "Ils peuvent �diter directement les fichiers textes" + CRLF + "du sous-dossier le�ons\Personnalis�." + CRLF2 + "Puis, dans n'importe quel menu," + CRLF + "on pourra alors basculer le niveau," + CRLF + "en tapant sur Alt, puis sur O, Options." + CRLF + "Puis S, Standard, ou P, Personnalis�."
pgia4 = "Cette pr�sentation est TERMIN�E." + CRLF2 + "Puisque vous l'avez suivie enti�rement," + CRLF + msgR�ussi�100

' Pour qui, Pourquoi ?
pgib1 = "ApprentiClavier est d�velopp� � l'intention des non-voyants." + CRLF + "Ils utiliseront un lecteur d'�cran tel que Jaws ou NVDA." + CRLF2 + "D�s que vous conna�trez la touche Alt et quelques lettres," + CRLF + "vous pourrez modifier le d�bit des explications." + CRLF + "Dans n'importe quel menu, vous frapperez Alt, puis O pour Options." + CRLF + "Alors, dans le sous-menu, vous choisirez N pour Normal, ou R pour Rapide." + CRLF2 + "ApprentiClavier est utilisable par les voyants, sans vocalisation." + CRLF2 + "Les exercices progressifs sont con�us pour un apprentissage autonome." + CRLF2 + "Pour cette version," + CRLF + "certaines indications sur les emplacements des touches" + CRLF + "seraient incorrectes pour les ordinateurs portables."
pgib2 = "Il vous sera propos� successivement 5 sortes de le�ons." + CRLF2 + "- l'entra�nement � la frappe des lettres, des mots, des phrases;" + CRLF + "- des exercices de r�gularit�;" + CRLF + "- des exercices de vitesse;" + CRLF + "- des dict�es." + CRLF + "- des exercices sur les raccourcis clavier, et le pav� num�rique."
pgib3 = "Il y a 3 objectifs :" + CRLF + "Premi�rement : permettre le rep�rage et la frappe des touches." + CRLF + "Deuxi�mement : faciliter l'utilisation d'une synth�se vocale." + CRLF + "Troisi�mement : utiliser les combinaisons de touches."
pgib4 = "Cette explication est TERMIN�E. " + CRLF + "Puisque vous l'avez suivie enti�rement," + CRLF + msgR�ussi�100

' Pour la frappe, des conseils
pgic1 = "INSTALLEZ-VOUS CONFORTABLEMENT." + CRLF + "Le dos est appuy� au dossier du si�ge." + CRLF + "Les bras sont en souplesse contre le corps." + CRLF + "Les poignets restent arrondis, d�tendus." + CRLF + "Chaque doigt se d�tache de la main pour frapper," + CRLF + "et revient � sa position de d�part," + CRLF + "avant de relancer une nouvelle frappe." + CRLF2 + "Il est recommand� de pratiquer chaque jour," + CRLF + "pendant 20 � 25 minutes." + CRLF + "L'utilisateur est fatigu� non seulement par la frappe," + CRLF + "mais aussi par l'�coute de la synth�se vocale."
pgic2 = "� PROPOS DES MENUS." + CRLF + "Le menu principal permet de choisir l'une des le�ons." + CRLF + "Chaque le�on comporte aussi un menu," + CRLF + "avec un choix des exercices." + CRLF2 + "Dans tous les menus," + CRLF + "on peut utiliser les touches Fl�che Haut ou Fl�che Bas," + CRLF + "puis la touche Entr�e pour valider votre choix." + CRLF2 + "La touche �chappement est tr�s importante." + CRLF + "�chap vous fait quitter l'exercice," + CRLF + "revenir au menu des exercices," + CRLF + "puis revenir au menu principal," + CRLF + "et m�me quitter compl�tement."
pgic3 = "� PROPOS DES ERREURS DE FRAPPE." + CRLF + "En g�n�ral, ApprentiClavier passe � la lettre suivante," + CRLF + "au bout de 5 fautes de frappe sur la lettre demand�e." + CRLF + "Il est pr�f�rable de ne pas s'acharner sur une lettre," + CRLF + "il vaut mieux refaire l'exercice." + CRLF2 + "Dans certains exercices tels que les phrases et les dict�es," + CRLF + "on passe � la lettre suivante d�s la deuxi�me erreur." + CRLF + "Pour les raccourcis clavier, le but doit �tre z�ro faute."
pgic4 = "� PROPOS DES R�SULTATS." + CRLF + "Le taux de r�ussite sera annonc�," + CRLF + "� la fin de chaque exercice," + CRLF + "sous la forme d'un pourcentage." + CRLF + "Il est aussi enregistr� et associ� � votre nom." + CRLF + "Il sera affich� � la fin de chaque ligne des menus." + CRLF2 + "Il faut obtenir au moins 85 pourcent." + CRLF + "M�me si le r�sultat est bon," + CRLF + "vous pourrez refaire l'exercice," + CRLF + "en vous d�pla�ant par les fl�ches dans les menus." + CRLF2 + "Vous pouvez consulter vos r�sultats," + CRLF + "en choisissant dans le menu principal" + CRLF + "l'avant-derni�re option, Consulter."
pgic5 = "Ces conseils de frappe sont TERMIN�S." + CRLF + "Puisque vous les avez suivis enti�rement," + CRLF + msgR�ussi�100

' Le�on1
pg1a1 = "Voici une DESCRIPTION du CLAVIER classique." + CRLF + "Le clavier principal," + CRLF + "c'est la plus grande partie du clavier" + CRLF + "qui se situe � gauche." + CRLF2 + "Ce clavier principal se compose de 5 rang�es." + CRLF + "Attention, au-dessus," + CRLF + "on trouve encore une sixi�me rang�e s�par�e." + CRLF + "C'est la rang�e des touches de fonction." + CRLF2 + "� droite du clavier principal," + CRLF + "on trouve des groupes de touches," + CRLF + "puis encore � droite un damier de 5 rang�es et 4 colonnes," + CRLF + "appel� pav� num�rique."
pg1a2 = "Voici les PREMIERES TOUCHES indispensables." + CRLF2 + "La TOUCHE ESPACE." + CRLF2 + "Partez du devant du clavier principal." + CRLF2 + "Au centre de la ligne du bas, il y a une grande barre," + CRLF + "c'est la TOUCHE ESPACE, ou barre d'espacement." + CRLF2 + "Utilisez le pouce gauche ou le pouce droit."
pg1a3 = "Si vous utilisez une synth�se vocale, ne soyez pas surpris." + CRLF + "Dans la plupart des exercices," + CRLF + "vous n'entendrez pas la touche que vous venez de taper," + CRLF + "car ApprentiClavier d�sactive l'�cho clavier." + CRLF2 + "Vous allez entendre deux notes de musique, " + CRLF + "car c'est le d�but de votre exercice."
pg1am1 = "La TOUCHE Entr�e." + CRLF2 + msgEnter + CRLF2 + "Entr�e vous permet de valider votre choix," + CRLF + "ou de progresser dans les pages d'explications." + CRLF2 + "Utilisez l'auriculaire droit."
pg1am2 = "La TOUCHE �chap." + CRLF + "Attendez d'avoir valid� ce message," + CRLF + "avant d'exercer la touche �chap." + CRLF2 + "Partez du devant, et avec l'auriculaire gauche," + CRLF + "contournez le clavier principal." + CRLF2 + "La touche d�tach�e qui fait le coin," + CRLF + "en haut � gauche du clavier, c'est la touche �chap." + CRLF2 + "�chap vous permet de quitter l'exercice" + CRLF + "ou l'action en cours," + CRLF + "et m�me de quitter progressivement ApprentiClavier." + CRLF2 + "Attention, exceptionnellement, il faudra frapper 2 fois �chap," + CRLF + "pour interrompre la le�on 1."
pg1b1 = "Les TOUCHES FL�CHES." + CRLF2 + "Partez du devant du clavier principal," + CRLF + "allez � droite de la barre d'espacement," + CRLF + "sautez quatre touches." + CRLF2 + "Vous rencontrez un groupe isol� de quatre touches," + CRLF + "dont trois sont horizontales et une au-dessus," + CRLF + "ce sont les FL�CHES." + CRLF2 + "De gauche � droite, on trouve :" + CRLF + "Fl�che Gauche, Fl�che Bas, Fl�che Droite," + CRLF + "et au-dessus on trouve Fl�che Haut." + CRLF2 + "Vous placerez l'index sur Fl�che Gauche."
pg1b2 = "Avec les fl�ches, vous pourrez vous d�placer dans les menus," + CRLF + "ou dans les pages d'explications." + CRLF2 + "Par exemple, dans un menu," + CRLF + "Fl�che Haut vous positionne sur la ligne au-dessus," + CRLF + "c'est-�-dire sur l'exercice pr�c�dent." + CRLF2 + "Dans toutes les pages d'explications," + CRLF + "si vous trouvez le d�bit des explications trop rapide," + CRLF + "vous pouvez relire imm�diatement ligne par ligne," + CRLF + "en tapant sur Fl�che Bas."
pg1bm1 = "La TOUCHE F1." + CRLF2 + "Partez du devant. Contournez par la gauche jusqu'en haut." + CRLF + "� droite de la touche �chap," + CRLF + "vous avez un groupe de 4 touches horizontales," + CRLF + "qui commen�e � gauche par F1." + CRLF2 + "F1 est une touche qui vous offre de l'aide," + CRLF + "dans ApprentiClavier, comme dans la plupart des logiciels." + CRLF2 + "Utilisez l'annulaire gauche."
pg1bm2 = "La TOUCHE F2." + CRLF2 + "C'est la touche situ�e � droite de la touche F1." + CRLF2 + "Dans ApprentiClavier, F2 est une touche" + CRLF + "qui vous indique l'emplacement et le doigt pr�vu" + CRLF + "pour chaque touche demand�e." + CRLF2 + "Utilisez l'annulaire gauche comme pour F1."
pg1bm3 = "La TOUCHE F3." + CRLF2 + "C'est la touche situ�e � droite de la touche F2." + CRLF2 + "Dans ApprentiClavier, F3 est la touche d'aide-m�moire," + CRLF + "qui vous informe sur chaque touche." + CRLF + "Vous pouvez utiliser F3 partout," + CRLF + "dans les le�ons, dans les menus," + CRLF + "et m�me d�s le lancement de la page de bienvenue." + CRLF2 + "Utilisez le majeur gauche."
pg1c0 = "Il y a trois touches � gauche de la barre ESPACE." + CRLF + "Ce sont les touches ALT, Windows, et CONTROL." + CRLF + "Attendez la le�on 13 pour le doigt� de la touche Windows." + CRLF2 + "Voici la TOUCHE gauche ALT." + CRLF + "En g�n�ral, cette touche lance la barre de menu." + CRLF + "Dans ApprentiClavier, apr�s vous �tre plac� dans le menu principal," + CRLF + "vous pourrez par exemple modifier le d�bit de la voix," + CRLF + "en tapant Alt, puis Fl�che Droite pour passer aux Options," + CRLF + "puis des Fl�ches Bas, puis Entr�e pour valider." + CRLF2 + "Avec le pouce gauche." + CRLF + "Partez de la barre ESPACE." + CRLF + "La touche Alt est la premi�re touche � sa gauche."
pg1cm0 = "Les TOUCHES CONTROL. On les appelle aussi C t r l." + CRLF2 + "Avec l'auriculaire gauche." + CRLF + "Allez au coin en bas et � gauche du clavier principal." + CRLF + "C'est la touche Control de gauche." + CRLF2 + "Semblablement avec l'auriculaire droit," + CRLF + "allez au coin droit en bas du clavier principal." + CRLF + "C'est aussi une touche CONTROL." + CRLF2 + "On peut utiliser indiff�remment" + CRLF + "la touche droite ou la touche gauche," + CRLF + "mais il vaut mieux s'entra�ner sur les 2 touches."
pg1d1 = "Maintenant vous allez frapper les touches," + CRLF + "que vous venez de d�couvrir." + CRLF2 + msgConseils + CRLF2 + "Mettez vos mains sur la rang�e de d�part," + CRLF + "qui est la troisi�me rang�e de touches" + CRLF + "en partant du devant." + CRLF2 + "L'index gauche est sur le point en relief de gauche, lettre F." + CRLF2 + "L'index droit est sur le point en relief de droite, lettre J."
pg1d2 = "Placez chaque pouce sur la barre d'espacement." + CRLF2 + "Gardez les poignets arrondis." + CRLF2 + "Tapez les touches demand�es d'un coup sec," + CRLF + "en ramenant le doigt � sa position de d�part."

' Le�on2
pg2a1 = msgConseils + CRLF2 + "Gardez vos mains sur la troisi�me ligne de touches" + CRLF + "en partant du devant du clavier." + CRLF2 + "L'index gauche doit �tre sur le point en relief � gauche," + CRLF + "lettre F." + CRLF2 + "L'index droit doit �tre sur le point en relief" + CRLF + "situ� trois touches plus loin � droite, lettre J."
pg2a2 = msgHit + CRLF2 + "Avec la main gauche, vous devrez taper :" + CRLF2 + "Le Q avec l'auriculaire." + CRLF + "Le S avec l'annulaire." + CRLF + "Le D avec le majeur." + CRLF + "Le F avec l'index."
pg2am1 = "Avec la main droite, vous devrez taper :" + CRLF2 + "Le J : avec l'index." + CRLF + "Le K : avec le majeur." + CRLF + "Le L : avec l'annulaire." + CRLF + "Le M : avec l'auriculaire."
pg2am2 = " Attention." + CRLF2 + " Reprenez avec les 2 mains."
pg2b1 = msgConseils + CRLF2 + "Gardez les mains sur la troisi�me rang�e de touches" + CRLF + "en partant du devant." + CRLF2 + "N'appuyez pas sur le clavier."
pg2b2 = msgHit + CRLF2 + "Tapez la touche pr�vue, mais seulement avec le doigt pr�vu." + CRLF + "� la cinqui�me erreur," + CRLF + "ApprentiClavier proposera la lettre qui suit." + CRLF2 + "Utilisez l'index pour les lettres G et H." + CRLF2 + "Pour le G, l'index gauche, en allant d'une touche vers la droite" + CRLF + "� partir du point en relief F." + CRLF2 + "Pour le H, l'index droit, en allant d'une touche vers la gauche" + CRLF + "� partir du deuxi�me point en relief J."
pg2bm1 = "Attention." + CRLF2 + "Avec la main gauche."
pg2bm2 = "Attention." + CRLF2 + "Avec la main droite."
pg2bm3 = "Attention." + CRLF2 + "Avec les deux mains."
pg2c1 = "Maintenant vous allez frapper les touches de la rang�e" + CRLF + "juste au-dessus de celle que vous venez d'�tudier." + CRLF2 + msgConseils + CRLF2 + "Frappez d'un coup sec et uniquement avec le doigt pr�vu." + CRLF2 + "Gardez les autres doigts � leur place" + CRLF + "et sans appuyer sur les touches."
pg2c2 = msgHit + CRLF + "Vous allez taper avec la main gauche." + CRLF2 + "Le A. Avec l'auriculaire." + CRLF + "Partez du Q, frappez nettement," + CRLF + "juste au-dessus, l�g�rement � gauche." + CRLF2 + "Le Z. Avec l'annulaire." + CRLF + "Partez du S. Frappez juste au-dessus, l�g�rement � gauche." + CRLF2 + "Le E. Avec le majeur. Partez du D." + CRLF + "Frappez juste au-dessus, l�g�rement � gauche." + CRLF2 + "Le R. Avec l'index. Partez du F." + CRLF + "Frappez juste au-dessus, l�g�rement � gauche."
pg2cm1 = "Maintenant, avec la main droite, vous devrez taper :" + CRLF2 + "Le U : avec l'index. Partez du J. Juste au-dessus, l�g�rement � gauche." + CRLF2 + "Le I : avec le majeur. Partez du K. Juste au-dessus, l�g�rement � gauche" + CRLF2 + "Le O : avec l'annulaire. Partez du L. Juste au-dessus, l�g�rement � gauche" + CRLF2 + "Le P : avec l'auriculaire. Partez du M. Juste au-dessus, l�g�rement � gauche."
pg2cm2 = "Attention." + CRLF2 + "Reprenez avec les 2 mains."
pg2cm3 = "C'est bient�t fini�"
pg2d1 = "Maintenant vous allez taper des groupes de plusieurs mots," + CRLF + "avec les lettres d�j� vues." + CRLF2 + "Les voyants doivent �viter de regarder le clavier." + CRLF2 + "Le caract�re ESPACE n'est pas sonoris� habituellement," + CRLF + "sauf ici pour la frappe."
pg2d2 = msgConseils + CRLF2 + "Gardez les mains sur la rang�e de d�part sans appuyer." + CRLF2 + "Pour faire r�p�ter la lettre, appuyez sur ESPACE." + CRLF2 + "Pour faire r�p�ter le mot," + CRLF + "appuyez sur CONTROL en le gardant enfonc�, et frappez ESPACE." + CRLF2 + "On vous proposera 2 fois chaque groupe de mots."
pg2d3 = msgHit + CRLF2 + "Ecoutez les mots. Tapez la lettre demand�e." + CRLF + "Tapez toujours un ESPACE entre les mots." + CRLF2 + "Pour faire un ESPACE," + CRLF + "tapez d'un coup sec sur la barre d'espacement." + CRLF + "Utilisez le pouce de la main," + CRLF + "qui n'a pas frapp� la derni�re lettre."
pg2e1 = "Maintenant vous allez utiliser l'index de chaque main."
pg2e2 = msgConseils + CRLF2 + "Gardez les poignets arrondis." + CRLF2 + "Ne bougez que le doigt pr�vu." + CRLF + "Ne faites aucune pression avec les autres doigts."
pg2e3 = msgHit + CRLF2 + "Tapez la lettre demand�e." + CRLF + "� la cinqui�me erreur, on passera � la lettre suivante." + CRLF2 + "Main gauche, le T." + CRLF + "L'index part du F. Allez en diagonale juste au-dessus � droite, pr�s de la lettre R." + CRLF2 + "Main droite, le Y." + CRLF + "L'index part du J. Allez en diagonale juste au-dessus � gauche, pr�s de la lettre U."
pg2em1 = "�coutez pour faire la diff�rence entre le T de Th�r�se, et le P de Patrick."
pg2f0 = "Maintenant vous allez taper � nouveau des groupes de 2 mots, s�par�s par un espace."
pg2g1 = "Maintenant vous allez taper des groupes de 3 mots." + CRLF2 + "Les mots sont courts, ils sont s�par�s par un espace."
pg2g2 = msgConseils + CRLF2 + "Tapez � votre rythme." + CRLF + "Les mains sont en place, elles sont souples." + CRLF2 + "Si vous �coutez une synth�se vocale," + CRLF + "ne la mettez pas trop fort, cela fatigue."
pg2g3 = msgHit + CRLF2 + "Ecoutez les mots. Tapez la lettre demand�e."
pg2h0 = "Maintenant vous allez taper des phrases courtes." + CRLF2 + "�coutez, puis frappez � votre rythme." + CRLF2 + "Tapez un espace entre les mots." + CRLF + "Mais ne tapez pas d'espace entre les phrases."

' Le�on3
pg3a1 = "Maintenant vous allez taper les lettres" + CRLF + "juste en-dessous de la rang�e de d�part." + CRLF2 + "RAPPEL." + CRLF2 + "F2 donne le nom et le doigt pr�vu" + CRLF + "pour la touche demand�e." + CRLF2 + "F3 est la touche d'aide-m�moire," + CRLF + "qui donne le nom de la touche frapp�e."
pg3a2 = msgConseils + CRLF2 + "Pour chaque touche, partez de la rang�e de d�part," + CRLF + "Allez � la rang�e inf�rieure vers la droite." + CRLF2 + "Frappez tranquillement. D�placez seulement le doigt pr�vu." + CRLF + "Il faut que le doigt soit ind�pendant de la main."
pg3a3 = msgHit + CRLF2 + "Frappez la lettre demand�e." + CRLF2 + "Ramenez le doigt � la position de d�part." + CRLF2 + "N'appuyez pas avec les autres doigts."
pg3a4 = "Avec la main gauche vous devrez taper :" + CRLF2 + "Le W." + CRLF + "Avec l'auriculaire. Partez de Q. Juste en-dessous � droite." + CRLF2 + "Le X." + CRLF + "Avec l'annulaire. Partez de S. Allez en bas � droite." + CRLF2 + "Le C." + CRLF + "Avec le majeur. Partez de D. Allez en bas � droite." + CRLF2 + "Le V." + CRLF + "Avec l'index. Partez de F. Allez en bas � droite."
pg3am1 = "Les 3 rang�es. Main gauche."
pg3am2 = "Maintenant, avec la main droite." + CRLF2 + "Le N. Avec l'index. Partez de J. Allez en bas, vers la gauche."
pg3b1 = msgConseils + CRLF2 + "Ecoutez et distinguez le B de Bernard et le V de V�ronique." + CRLF2 + "Ecoutez et distinguez aussi le B de Bernard et le P de Patrick."
pg3b2 = msgHit + CRLF2 + "Maintenant vous allez taper le B." + CRLF + "Avec l'index. Partez de F. Allez en bas nettement � droite."
pg3c0 = msgConseils + CRLF2 + "Ecoutez et distinguez les mots qui se ressemblent."

' Le�on4
pg4a1 = "Cette le�on approfondit l'�tude de l'alphabet." + CRLF2 + "Elle permet de reprendre des habitudes oubli�es." + CRLF2 + "Attention. Le rythme sera plus rapide." + CRLF + "On vous proposera la lettre suivante" + CRLF + "d�s que vous aurez frapp�."
pg4a2 = "RAPPEL." + CRLF2 + "F1 donne l'aide sur les commandes." + CRLF + "F2 donne le nom et le doigt pr�vu pour la touche demand�e." + CRLF + "F3 suivi d'une touche, donne le nom de la touche."
pg4a3 = msgConseils + CRLF2 + "Gardez les mains sur la troisi�me ligne de touches " + CRLF + "en partant du devant du clavier." + CRLF2 + "L'index gauche doit �tre sur le point en relief � gauche," + CRLF + "lettre F." + CRLF + "C'est la cinqui�me touche en partant du bord gauche du clavier." + CRLF2 + "L'index droit doit �tre sur le point en relief" + CRLF + "situ� trois touches plus loin � droite, lettre J."
pg4a4 = "Toutes les indications partent de cette troisi�me rang�e," + CRLF + "appel�e rang�e de d�part." + CRLF2 + "Apr�s 5 erreurs, on vous proposera la lettre suivante."
pg4am1 = "Attention." + CRLF2 + "Il faudra distinguer le F de Fran�ois et le S de Simone."
pg4b0 = "Maintenant vous allez encore taper des touches" + CRLF + "de la rang�e de d�part," + CRLF + "celles � utiliser avec la main droite."
pg4c0 = "Maintenant vous allez taper des lettres" + CRLF + "de la rang�e juste au-dessus de la rang�e de d�part." + CRLF2 + msgHit + CRLF2 + "Partez de la lettre de d�part." + CRLF + "D�placez le doigt." + CRLF + "Frappez d'un coup sec." + CRLF + "Ramenez le doigt � sa place." + CRLF + "Ne d�placez pas les autres doigts."
pg4d0 = msgConseils + CRLF2 + "Le G et le T se frappent avec l'index gauche." + CRLF2 + "Le G est juste � droite du F." + CRLF2 + "Le T est � droite du R, c'est-�-dire au-dessus et � droite du F." + CRLF2 + "Vous distinguerez le T de Thomas d'avec le D de Denis."
pg4e0 = msgConseils + CRLF2 + "Maintenant avec la main droite," + CRLF + "allez de la touche de d�part vers le haut l�g�rement � gauche."
pg4f0 = msgConseils + CRLF2 + "Le H et le Y se frappent avec l'index droit." + CRLF2 + "Le H est juste � gauche du J." + CRLF2 + "Le Y est � gauche de U, c'est-�-dire au-dessus" + CRLF + "et tr�s � gauche du J."
pg4g0 = msgConseils + CRLF2 + "Avec la main gauche, vous allez taper les lettres de la rang�e" + CRLF + "juste en-dessous de la rang�e de d�part." + CRLF2 + "Les 4 lettres sont nettement d�cal�es vers la droite."
pg4h0 = msgConseils + CRLF2 + "Maintenant, vous continuez dans la rang�e" + CRLF + "en-dessous de la rang�e de d�part, avec l'index." + CRLF2 + "Avec la main gauche." + CRLF + "Le B. En-dessous du F, en extension tr�s � droite." + CRLF2 + "Avec la main droite." + CRLF + "Le N. En-dessous du J. L�g�rement � gauche."

' Le�on5
pg5a1 = msgConseils + CRLF2 + "Gardez vos mains sur la troisi�me ligne de touches" + CRLF + "en partant du devant du clavier." + CRLF2 + "L'index gauche doit �tre sur le point en relief � gauche," + CRLF + "lettre F." + CRLF2 + "L'index droit doit �tre sur le point en relief" + CRLF + "situ� trois touches plus loin � droite, lettre J."
pg5a2 = msgHit + CRLF2 + "Avec la main gauche, vous devrez taper :" + CRLF2 + "Le Q avec l'auriculaire." + CRLF2 + "Le D avec le majeur."
pg5b1 = "Maintenant vous allez taper des proverbes connus." + CRLF2 + "La synth�se vocale ne dira la phrase qu'une fois," + CRLF + "puis chaque mot est prononc� une fois." + CRLF2 + "RAPPEL." + CRLF2 + "ESPACE r�p�te la lettre demand�e." + CRLF + "CONTROL+ESPACE r�p�te le mot demand�."
pg5b2 = msgHit + CRLF2 + "Ne tapez qu'avec le doigt pr�vu." + CRLF2 + "Utilisez la touche F2 pour rappeler le doigt pr�vu."
pg5c0 = "Maintenant vous allez taper des phrases" + CRLF + "reprenant toutes les lettres de l'alphabet."


' Suite (sinon : procedure too large)
Module_global.main2
End Sub


' **************** MAIN2 **** TRADUIRE � DROITE DU SIGNE �GAL ******************************
Public Sub main2()
' Le�on 6
pg6a1 = "Maintenant vous allez taper des mots courts," + CRLF + "qui vous seront envoy�s au hasard." + CRLF2 + "� chaque nouvel essai de cet exercice," + CRLF + "la s�quence sera diff�rente."
pg6a2a = "Attention." + CRLF2 + "Vous disposez de beaucoup de temps," + CRLF + "avec "
pg6a2b = " secondes pour chaque mot." + CRLF2 + "Vous pouvez faire �peler," + CRLF + "avec le pouce gauche enfonc� sur la touche ALT," + CRLF + "et un coup bref sur ESPACE," + CRLF + "mais le compteur de temps continuera."
pg6a3 = "On changera de mot � chaque erreur de frappe." + CRLF2 + "Si votre frappe est correcte, la synth�se vocale est muette." + CRLF2 + "D�s que le mot sera r�ussi," + CRLF + "ou bien si une frappe est incorrecte," + CRLF + "ou bien si le temps est �coul�," + CRLF + "on vous demandera de taper sur ESPACE."

pg6b1 = "Maintenant vous allez encore taper des mots courts," + CRLF + "qui vous seront envoy�s au hasard." + CRLF2 + "� chaque nouvel essai de cet exercice," + CRLF + "la s�quence sera diff�rente."
pg6b2a = "Attention." + CRLF2 + "Vous disposez seulement de "
pg6b2b = " secondes pour chaque mot." + CRLF2 + "Vous pouvez faire �peler," + CRLF + "avec le pouce gauche enfonc� sur la touche ALT," + CRLF + "et un coup bref sur ESPACE," + CRLF + "mais le compteur de temps continuera."
pg6b3 = "On changera de mot � chaque erreur de frappe." + CRLF2 + "Si votre frappe est correcte, la synth�se vocale est muette." + CRLF2 + "D�s que le mot sera r�ussi," + CRLF + "ou bien si une frappe est incorrecte," + CRLF + "ou bien si le temps est �coul�," + CRLF + "on vous demandera de taper sur ESPACE."

pg6c0a = "Maintenant vous allez taper des mots plus longs," + CRLF + "tr�s r�guli�rement." + CRLF2 + "Vous disposez de "
pg6c0b = " secondes pour chaque mot."
pg6d0a = "Maintenant vous allez taper des mots se terminant par : ation." + CRLF2 + "Vous disposez de "
pg6d0b = " secondes pour chaque mot."

' Le�on 7
pg7a1 = "Vous allez taper une phrase." + CRLF2 + msgConseils + CRLF2 + "Vous aurez 2 r�sultats." + CRLF2 + "     Le pourcentage de r�ussite." + CRLF2 + "     La vitesse de frappe."
pg7a2 = "Tapez la combinaison Control+ESPACE pour faire R�P�TER le MOT."
pg7b1 = msgConseils + CRLF2 + "N'augmentez votre vitesse que si vous obtenez " + CRLF + "au moins 80 pourcent de r�ussite." + CRLF2 + "RAPPEL." + CRLF2 + "Tapez en m�me temps sur CONTROL et ESPACE" + CRLF + "pour faire r�p�ter le mot demand�."
pg7b2 = msgHit + CRLF2 + "Frappe r�guli�re. Mains souples et sur la rang�e de d�part." + CRLF2 + "D�s la deuxi�me erreur, on vous proposera la lettre suivante."
pg7c0 = "Voici encore des phrases pour la vitesse."

' Le�on8
pg8a1 = "Maintenant vous allez frapper les trois touches" + CRLF + "pour les majuscules et les minuscules." + CRLF2 + "Attention. Vous remarquerez" + CRLF + "que la synth�se vocale n'indique MAJUSCULE ou MINUSCULE" + CRLF + "que lorsqu'on vient effectivement de changer d'�tat." + CRLF2 + "Par exemple," + CRLF + "un deuxi�me appel temporaire par la touche majuscule" + CRLF + "ne sera pas sonoris� en tant que MAJUSCULE."
pg8a2 = "Dans ApprentiClavier, on appelle VERROUILLAGE-MAJUSCULES" + CRLF + "ou FIXE-MAJUSCULES, en anglais CAPSLOCK," + CRLF + "la touche qui bloque le clavier en mode MAJUSCULES." + CRLF2 + "On appelle la touche MAJUSCULE en abr�g� MAJ," + CRLF + "en anglais SHIFT," + CRLF + "la touche qui fait passer temporairement en majuscule." + CRLF2 + "Il y a en r�alit� 2 touches MAJUSCULES," + CRLF + "avec la m�me fonction." + CRLF2 + "On appelle MAJ-GAUCHE la touche situ�e � gauche" + CRLF + "pour le passage temporaire en majuscule." + CRLF + "On appelle MAJ-DROIT la touche semblable situ�e � droite."
pg8a3 = "La touche VERROUILLAGE-MAJUSCULES est � gauche du Q." + CRLF + "Pressez-la une seule fois." + CRLF + "Elle permettra alors d'�crire tout un texte en majuscules." + CRLF2 + "Attention. Pour d�verrouiller," + CRLF + "c'est-�-dire pour revenir � des minuscules," + CRLF + "la touche � utiliser d�pend du r�glage dans Windows." + CRLF2 + "Parfois, ce sont les touches MAJ-GAUCHE et MAJ-DROIT" + CRLF + "qui suppriment le verrouillage, d�s qu'on les rel�che." + CRLF + "Il suffira alors de presser MAJ-GAUCHE ou MAJ-DROIT bri�vement." + CRLF2 + "Mais en g�n�ral, il faut appuyer � nouveau " + CRLF + "sur la touche VERROUILLAGE-MAJUSCULES elle-m�me."
pg8a4 = "Avec la main gauche." + CRLF2 + "Le VERROUILLAGE-MAJUSCULES. Avec l'auriculaire. Partez de Q." + CRLF + "Juste � gauche. C'est une touche plus grande." + CRLF2 + "Le MAJ-GAUCHE. Avec l'auriculaire. Partez de Q." + CRLF + "Allez en bas nettement � gauche en descendant d'une rang�e." + CRLF + "C'est une touche assez grande."
pg8am1 = "Attention." + CRLF + "Sur certains claviers, la touche MAJ-DROIT est raccourcie," + CRLF + "pour laisser la place � une touche � sa gauche." + CRLF + "Ici nous donnons l'emplacement le plus fr�quent." + CRLF2 + "Avec la main droite." + CRLF2 + "Le MAJ-DROIT. Avec l'auriculaire. Partez de M." + CRLF + "Descendez � la rang�e inf�rieure, tr�s � droite." + CRLF + "C'est une touche souvent tr�s grande."

pg8b1 = "Maintenant vous allez taper les signes de ponctuation" + CRLF + "accessibles en minuscules," + CRLF + "c'est-�-dire la virgule, le point-virgule, le deux-points," + CRLF + "et le point d'exclamation." + CRLF2 + msgConseils + CRLF + "On rajoute un espace avant les ponctuations" + CRLF + "constitu�es de 2 signes," + CRLF + "donc on rajoute toujours un espace" + CRLF + "avant point-virgule, deux-points," + CRLF + "point d'interrogation et point d'exclamation." + CRLF + "Mais on n'en rajoute pas avant une virgule ou un point." + CRLF2 + "Et on place toujours un espace derri�re," + CRLF + "si on continue sur la m�me ligne."
pg8b2 = msgHit + CRLF2 + "Maintenant vous allez taper avec la main droite." + CRLF2 + "La virgule." + CRLF + "Avec l'index. Partez de J. Allez en bas nettement � droite." + CRLF + "Le point-virgule." + CRLF + "Avec le majeur. Partez de K. Allez en bas nettement � droite." + CRLF + "Les deux-points." + CRLF + "Avec l'annulaire. Partez de L. Allez en bas bien � droite." + CRLF + "Le point d'exclamation." + CRLF + "Avec l'auriculaire. Partez de M. Allez en bas bien � droite."

pg8c1 = "Maintenant vous allez taper les signes de ponctuation" + CRLF + "accessibles en majuscules," + CRLF + "c'est-�-dire le point d'interrogation et le point." + CRLF2 + "Nous �tudierons aussi " + CRLF + "la Barre-Oblique," + CRLF + "et le signe Section, utilis� parfois " + CRLF + "pour indiquer une nouvelle section dans un document."
pg8c2 = msgHit + CRLF2 + "Mettez-vous en majuscules bloqu�es," + CRLF + "en appuyant une fois sur la touche VERROUILLAGE-MAJUSCULES." + CRLF2 + "Puis avec la main droite." + CRLF2 + "Le point d'interrogation." + CRLF + "Avec l'index. Partez de J. Allez en bas nettement � droite." + CRLF + "Le point." + CRLF + "Avec le majeur. Partez de K. Allez en bas nettement � droite." + CRLF + "La Barre-Oblique." + CRLF + "Avec l'annulaire. Partez de L. Allez en bas bien � droite." + CRLF + "La Section, appel�e aussi Chapitre ou Paragraphe." + CRLF + "Avec l'auriculaire. Partez de M. Allez en bas bien � droite."
pg8cm1 = "Attention." + CRLF + "D�verrouillez les majuscules, utilisez les touches MAJ-GAUCHE ou MAJ-DROIT." + CRLF + "Voici toutes les ponctuations de la rang�e."

pg8d1 = "Maintenant vous allez taper des groupes de mots" + CRLF + "contenant des majuscules et des minuscules, avec des ponctuations."
pg8d2 = msgHit + CRLF2 + "Tapez MAJ-GAUCHE pour taper une lettre avec la main droite." + CRLF2 + "Tapez MAJ-DROIT pour taper une lettre avec la main gauche."

pg8e1 = "Maintenant vous allez taper le U grave," + CRLF + "ainsi que 2 accents particuliers." + CRLF2 + "Attention." + CRLF + "L'accent circonflexe et le tr�ma ne sont prononc�s" + CRLF + "qu'apr�s la frappe de la voyelle � accentuer." + CRLF2 + "Vous pouvez appeler la touche F2 pour vous aider" + CRLF + "au moment de taper une lettre accentu�e."
pg8e2 = "Pour le circonflexe et le tr�ma, l'accent se frappe d'abord," + CRLF + "puis juste apr�s, on frappe la voyelle." + CRLF2 + "Le tr�ma se fait en appuyant d'abord sur MAJ-GAUCHE," + CRLF + "en maintenant l'appui enfonc�, et en tapant la touche tr�ma." + CRLF + "Puis on tape la voyelle demand�e."
pg8e3 = msgConseils + CRLF2 + "Avec la main droite." + CRLF2 + "Le  U grave." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Allez sur la m�me rang�e � droite."
pg8e4 = "L'accent circonflexe." + CRLF + "En minuscule. Avec l'auriculaire. Partez de M." + CRLF + "Allez � la rang�e au-dessus et � droite." + CRLF2 + "Le tr�ma." + CRLF + "En majuscule. D'abord pressez et maintenez Maj-Gauche." + CRLF + "Puis avec l'auriculaire" + CRLF + "allez � la m�me touche que l'accent circonflexe."

pg8f1 = "Maintenant vous allez taper des signes" + CRLF + "souvent employ�s en informatique." + CRLF2 + "Selon les synth�ses vocales," + CRLF + "le signe �toile est prononc� �toile ou Ast�risque." + CRLF + "Le signe  Inf�rieur �  est prononc� Inf�rieur." + CRLF + "Le signe  Sup�rieur �  est prononc� Sup�rieur."
pg8f2 = "Les signes Ast�risque et Inf�rieur se frappent en minuscules." + CRLF2 + "Le Sup�rieur se fait en majuscule avec MAJ-DROIT."
pg8f3 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Attention. Sur certains claviers," + CRLF + "la touche Ast�risque se trouve sous la touche Entr�e." + CRLF + "Ici nous donnons l'emplacement le plus fr�quent." + CRLF2 + "L'Ast�risque." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Allez 2 touches plus loin sur la m�me rang�e � droite."
pg8f4 = "Avec la main gauche." + CRLF2 + "L'Inf�rieur." + CRLF + "En minuscule. Avec l'auriculaire. Partez de Q." + CRLF + "Allez � la rang�e inf�rieure et � gauche, juste avant MAJ-GAUCHE." + CRLF2 + "Le Sup�rieur." + CRLF + "En majuscule. Allez � la m�me touche que pour l'Inf�rieur."

pg8g1 = "Maintenant vous allez taper 4 signes" + CRLF + "souvent employ�s dans les textes et en informatique." + CRLF2 + "Le pourcent, la lettre grecque Mu ou Micro, le dollar" + CRLF + "et la livre anglaise."
pg8g2 = msgHit + CRLF2 + "Le PourCent et le Mu se frappent en majuscules avec MAJ-GAUCHE." + CRLF2 + "Avec la synth�se vocale, le Mu est souvent prononc� MICRO."
pg8g3 = "Avec la main droite." + CRLF2 + "Le PourCent." + CRLF + "Avec l'auriculaire en majuscule. Partez de M." + CRLF + "Allez une touche � droite." + CRLF2 + "Attention. Sur certains claviers," + CRLF + "la touche Mu se trouve en dessous de la touche Entr�e." + CRLF + "Ici nous donnons l'emplacement le plus fr�quent." + CRLF + "Le Mu." + CRLF + "Avec l'auriculaire en majuscule. Partez de M." + CRLF + "Allez 2 touches plus loin � droite." + CRLF + "C'est donc en g�n�ral la m�me touche que l'ast�risque," + CRLF + "mais en majuscule."
pg8g4 = "Avec la main droite." + CRLF2 + "Le Dollar." + CRLF + "Avec l'auriculaire en minuscule. Partez de M." + CRLF + "Allez � la rang�e sup�rieure tr�s en extension � droite," + CRLF + "avant la grande touche." + CRLF2 + "La Livre." + CRLF + "Avec l'auriculaire en majuscule. Partez de M." + CRLF + "Allez � la m�me touche que le Dollar."

pg8h1 = "Maintenant vous allez taper les lettres accentu�es," + CRLF + "e aigu, e grave, a grave," + CRLF + "situ�es dans la rang�e la plus �loign�e du clavier principal." + CRLF + "C'est la deuxi�me rang�e au-dessus de la rang�e de d�part." + CRLF2 + msgConseils + CRLF2 + "Rep�rez la touche � frapper." + CRLF + "Le doigt se lib�re de la main sans l'entra�ner." + CRLF + "Il revient � sa position de d�part" + CRLF + "avant de frapper une autre touche."
pg8h2 = msgHit + CRLF2 + "Avec la main gauche." + CRLF + "Le �." + CRLF + "Avec l'auriculaire. Partez de Q." + CRLF + "Montez de 2 rang�es directement au-dessus." + CRLF + "C'est la troisi�me touche en partant de la gauche" + CRLF + "dans cette rang�e."
pg8h3 = "Avec la main droite." + CRLF2 + "Le �." + CRLF + "Avec l'index. Partez de J." + CRLF + "Montez de 2 rang�es au-dessus l�g�rement � gauche." + CRLF2 + "Le �." + CRLF + "Avec l'annulaire. Partez de L." + CRLF + "Montez de 2 rang�es au-dessus l�g�rement � droite."

' Le�on9
pg9a1 = "Maintenant vous allez taper le signe �" + CRLF + "qui correspond � l'exposant au carr�." + CRLF + "Vous taperez aussi le signe & prononc� �commercial," + CRLF + "qui remplace la pr�position ET." + CRLF2 + msgHit + CRLF2 + "Avec la main gauche, frappez sur la rang�e du haut," + CRLF + "2 rang�es au-dessus de la rang�e de d�part."
pg9a2 = "Le �." + CRLF + "C'est la touche la plus � gauche. Avec l'auriculaire." + CRLF + "En minuscule. Partez de Q. Montez 2 rang�es au-dessus." + CRLF + "En extension maximum � gauche." + CRLF2 + "Le &." + CRLF + "Avec l'auriculaire. En minuscule. Partez de Q." + CRLF + "Montez 2 rang�es au-dessus, l�g�rement � gauche."

pg9b1 = "Maintenant vous allez taper des signes de ponctuation" + CRLF + "de la rang�e sup�rieure." + CRLF + "Cet exercice �tudie les guillemets, l'apostrophe," + CRLF + "la parenth�se gauche et le tiret." + CRLF2 + msgConseils + CRLF + "Toutes les ponctuations de la rang�e sup�rieure " + CRLF + "se font en minuscules." + CRLF2 + "On ne rajoute pas d'espace, sauf dans les 4 cas suivants :" + CRLF + "avant d'ouvrir le guillemet," + CRLF + "avant d'ouvrir une parenth�se," + CRLF + "apr�s avoir ferm� le deuxi�me guillemet," + CRLF + "apr�s avoir ferm� la parenth�se droite."
pg9b2 = "PRONONCIATION." + CRLF2 + "Selon les synth�ses, la premi�re parenth�se se prononce : " + CRLF + "ouvre parenth�se, parenth�se gauche ou parenth�se ouverte." + CRLF2 + "La parenth�se qui termine se prononce : " + CRLF + "ferme parenth�se, parenth�se droite ou parenth�se ferm�e." + CRLF + "Il en sera de m�me pour les crochets et les accolades."
pg9b3 = msgHit + CRLF + "Avec la main gauche." + CRLF + "Le guillemet." + CRLF + "Avec l'annulaire. Partez de S." + CRLF + "Montez de 2 rang�es au-dessus directement." + CRLF + "L'apostrophe." + CRLF + "Avec le majeur. Partez de D." + CRLF + "Montez de 2 rang�es au-dessus directement." + CRLF + "La parenth�se gauche." + CRLF + "Avec l'index. Partez de F." + CRLF + "Montez de 2 rang�es au-dessus directement." + CRLF + "Le tiret." + CRLF + "Avec l'index en extension. Partez de F." + CRLF + "Montez de 2 rang�es au-dessus et � droite."

pg9c1 = "Maintenant vous allez taper des signes de ponctuation" + CRLF + "de la rang�e du haut � main droite." + CRLF2 + "Vous allez �tudier le soulign� et la parenth�se droite."
pg9c2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Le signe Soulign�." + CRLF + "C'est un caract�re qui est une sorte de tiret." + CRLF + "Avec l'index. Partez de J." + CRLF + "Montez de 2 rang�es au-dessus, et l�g�rement � droite." + CRLF2 + "La parenth�se droite." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Montez de 2 rang�es au-dessus directement."

pg9d1 = "Maintenant vous allez taper encore 4 signes" + CRLF + "de la rang�e sup�rieure." + CRLF2 + "Le �, le Degr�, le signe �gal, et le signe d'addition Plus." + CRLF2 + msgConseils + CRLF2 + "Le � se frappe uniquement en minuscule." + CRLF + "En majuscule, on tape simplement le C majuscule." + CRLF2 + "Le signe �gal se fait uniquement en minuscule," + CRLF + "m�me au milieu des chiffres." + CRLF2 + "Par contre, on tape en majuscules les signes Degr� et Plus."
pg9d2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Le �. Avec le majeur. Partez de K." + CRLF + "Montez de 2 rang�es au-dessus directement." + CRLF + "Le �. En majuscule. Avec l'auriculaire." + CRLF + "Partez de M. Montez de 2 rang�es au-dessus directement." + CRLF + "Le �gal. En minuscule. Avec l'auriculaire." + CRLF + "Partez de M. Montez de 2 rang�es au-dessus, et � droite." + CRLF + "Le Plus. En majuscule. Avec l'auriculaire." + CRLF + "Partez de M. Montez de 2 rang�es au-dessus, et � droite." + CRLF + "C'est la m�me touche que le signe �gal."

' Le�on10
pg10a1 = "Maintenant vous allez taper les chiffres." + CRLF2 + "PRONONCIATION." + CRLF2 + "Les nombres sont prononc�s globalement" + CRLF + "en prononciation fran�aise," + CRLF + "par exemple 10000 est prononc� dix mille." + CRLF2 + "Mais certaines synth�ses vocales ne prononcent que" + CRLF + "chiffre apr�s chiffre, soit 1 0 0 0 0." + CRLF2 + "Pour une prononciation globale," + CRLF + "�crivez le nombre complet sans espaces."
pg10a2 = msgConseils + CRLF + "Les chiffres se trouvent � 2 endroits." + CRLF + "D'une part sur la rang�e sup�rieure du clavier principal," + CRLF + "d'autre part dans le pav� num�rique de droite." + CRLF + "Ici nous �tudions seulement le clavier principal." + CRLF2 + "Les chiffres du clavier principal se tapent en majuscules." + CRLF2 + "En France, les d�cimales se marquent par une virgule," + CRLF + "et les milliers par un point ou un espace." + CRLF2 + "Dans beaucoup de pays anglo-saxons, c'est l'inverse." + CRLF + "Les d�cimales anglo-saxonnes se marquent par un point," + CRLF + "et les milliers par une virgule."
pg10a3 = msgHit + CRLF + "En majuscules. Avec la main gauche." + CRLF + "Le 1. Avec l'auriculaire." + CRLF + "Partez de Q. Montez de 2 rang�es au-dessus, nettement � gauche." + CRLF + "C'est la deuxi�me touche de cette rang�e." + CRLF + "Le 2. Avec l'auriculaire." + CRLF + "Partez de Q. Montez de 2 rang�es au-dessus, l�g�rement � droite." + CRLF + "Le 3. Avec l'annulaire." + CRLF + "Partez de S. Montez de 2 rang�es au-dessus, l�g�rement � droite." + CRLF + "Le 4. Avec le majeur." + CRLF + "Partez de D. Montez de 2 rang�es au-dessus, l�g�rement � droite."
pg10am1 = "Attention, nombres � 2 chiffres, tapez le nombre enti�rement."
pg10am2 = "Attention, nombres � 3 chiffres, tapez le nombre enti�rement."

pg10b1 = "Maintenant vous allez taper les trois chiffres 5, 6, 7." + CRLF2 + "PRONONCIATION." + CRLF2 + "Avec la synth�se vocale, �coutez la diff�rence entre 5 et 7." + CRLF2 + "Utilisez les touches F2 et F3 " + CRLF + "pour rappeler l'emplacement et le doigt pr�vu."
pg10b2 = msgHit + CRLF + "Avec la main gauche." + CRLF + "Le 5. Avec l'index." + CRLF + "Partez de F. Montez de 2 rang�es au-dessus l�g�rement � droite." + CRLF2 + "Le 6. Avec l'index." + CRLF + "Partez de F. Montez de 2 rang�es au-dessus et tr�s � droite."
pg10b3 = "Avec la main droite." + CRLF2 + "Le 7. Avec l'index." + CRLF + "Partez de J. Montez de 2 rang�es au-dessus l�g�rement � gauche."
pg10bm1 = "Attention tous les chiffres de 1 � 7 "
pg10bm2 = "Attention voici des nombres "

pg10c1 = "Maintenant vous allez taper le 8, le 9, et le 0." + CRLF2 + "PRONONCIATION." + CRLF2 + "Avec la synth�se vocale, le 8 peut se confondre avec le 6," + CRLF + "quand il est pris isol�ment."
pg10c2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "Le 8." + CRLF + "Avec l'index. Partez de J." + CRLF + "Montez de 2 rang�es au-dessus, et l�g�rement � droite." + CRLF + "Le 9." + CRLF + "Avec le majeur. Partez de K." + CRLF + "Montez de 2 rang�es au-dessus, l�g�rement � droite." + CRLF + "Le 0." + CRLF + "Avec l'annulaire. Partez de L. Montez de 2 rang�es au-dessus, l�g�rement � droite."

' Le�on11
pg11a0 = "Vous allez taper une phrase." + CRLF2 + msgConseils + CRLF2 + "Vous aurez 2 r�sultats." + CRLF2 + "     Le pourcentage de r�ussite." + CRLF2 + "     La vitesse de frappe."

pg11b1 = msgConseils + CRLF2 + "N'augmentez votre vitesse que si vous obtenez au moins" + CRLF + "80 pourcent de r�ussite." + CRLF2 + "RAPPEL." + CRLF2 + "Tapez en m�me temps sur CONTROL et ESPACE pour r�p�ter" + CRLF + "le mot demand�."
pg11b2 = msgHit + CRLF2 + "Frappe r�guli�re. Mains souples et sur la rang�e de d�part." + CRLF2 + "ATTENTION." + CRLF + "D�s la deuxi�me erreur, on vous proposera la lettre suivante."

pg11c0 = "Voici encore des phrases pour la vitesse."

' Le�on12
pg12a1 = "Maintenant vous allez frapper 2 touches" + CRLF + "qui agissent sur le texte." + CRLF2 + "Ce sont les touches INSERTION, en anglais INSERT, ou INS," + CRLF + "et SUPPRESSION, en anglais DELETE, ou DEL." + CRLF2 + "Attention." + CRLF2 + "Le pilote de synth�se peut modifier le comportement" + CRLF + "de certaines touches." + CRLF2 + "Par exemple Jaws peut �tre configur�" + CRLF + "pour supprimer l'effet bascule de la touche INSERTION."
pg12a2 = "Normalement, la touche INSERTION est une bascule." + CRLF2 + "Quand vous tapez du texte par-dessus un texte existant," + CRLF + "en g�n�ral, par d�faut il s'ins�re," + CRLF + "c'est-�-dire il s'ajoute au texte existant." + CRLF2 + "En appuyant sur la touche INSERTION," + CRLF + "vous basculez en mode REMPLACEMENT," + CRLF + "c'est-�-dire que votre future frappe �crasera l'ancien texte," + CRLF + "si votre curseur se trouve par-dessus un texte." + CRLF2 + "Si vous appuyez � nouveau plus tard sur la touche INSERTION," + CRLF + "vous reviendrez au mode INSERTION."
pg12a3 = "La touche SUPPRESSION supprime imm�diatement" + CRLF + "le caract�re courant de votre texte," + CRLF + "celui qui se trouve juste apr�s le curseur." + CRLF2 + "PRONONCIATION." + CRLF + "Pour la touche SUPPRESSION," + CRLF + "la synth�se prononce g�n�ralement le caract�re qui se pr�sente" + CRLF + "� droite du caract�re qui vient d'�tre supprim�."
pg12a4 = msgConseils + CRLF2 + "A droite du clavier principal," + CRLF + "dans le prolongement des 2 rang�es sup�rieures," + CRLF + "on trouve un ensemble de 2 ou 3 lignes de 3 touches." + CRLF2 + "Voici la touche SUPPRESSION." + CRLF + "Avec la main droite." + CRLF + "Partez de M. Allez � droite hors du clavier principal." + CRLF + "Placez l'index sur la premi�re touche" + CRLF + "de la petite rang�e rencontr�e."
pg12am1 = "Touche INSERTION." + CRLF + "En g�n�ral, la synth�se vocale prononce le mot INSERTION" + CRLF + "ou REMPLACEMENT." + CRLF + "Parfois, au lieu de REMPLACEMENT, le logiciel indique" + CRLF + "SURFRAPPE ou REFRAPPE." + CRLF2 + "C'est la premi�re touche � gauche" + CRLF + "dans la petite rang�e au-dessus." + CRLF2 + "Elle se trouve donc au-dessus de la touche SUPPRESSION."
pg12am2 = "Attention, avec des lettres du clavier."

pg12b1 = "Maintenant vous allez frapper les touches" + CRLF + "qui vous mettent en d�but ou en fin de ligne." + CRLF2 + "Ceci est vrai en g�n�ral pour les traitements de texte" + CRLF + "tels que WORD." + CRLF2 + msgConseils + CRLF2 + "La touche D�BUT, en anglais HOME," + CRLF + "vous place au d�but de la ligne courante." + CRLF2 + "La touche FIN, en anglais END," + CRLF + "vous place � la fin de la ligne courante."
pg12b2 = msgHit + CRLF2 + "Avec la main droite." + CRLF2 + "La touche FIN." + CRLF + "Partez de M. Placez la main sur la rang�e du bas" + CRLF + "du groupe des 2 ou 3 rang�es de 3 touches." + CRLF + "Le majeur frappe la deuxi�me touche de cette rang�e." + CRLF2 + "La touche D�BUT." + CRLF + "Partez de M. C'est le m�me groupe de touches." + CRLF + "Avec le majeur, frappez la deuxi�me touche de la rang�e au-dessus."
pg12bm1 = "Avec des lettres�"

pg12c1 = "Maintenant vous allez frapper les touches" + CRLF + "permettant de changer de page." + CRLF2 + msgConseils + CRLF2 + "PAGE-PR�C�DENTE, en anglais Page-Up," + CRLF + "permet de passer � la page pr�c�dente ou � l'�cran pr�c�dent." + CRLF2 + "PAGE-SUIVANTE, en anglais Page-Down," + CRLF + "permet de passer � la page suivante ou � l'�cran suivant." + CRLF2 + "Par exemple avec PAGE-PR�C�DENTE, dans ApprentiClavier," + CRLF + "vous pouvez revenir en arri�re dans les pages d'explications."
pg12c2 = msgHit + CRLF2 + "Partez du M. Allez � droite hors du clavier principal." + CRLF2 + "La touche PAGE-SUIVANTE." + CRLF + "Avec l'annulaire." + CRLF + "Frappez la troisi�me et derni�re touche du petit groupe," + CRLF + "dans sa rang�e inf�rieure." + CRLF2 + "La touche PAGE-PR�C�DENTE." + CRLF + "C'est la touche au-dessus de PAGE-SUIVANTE."

pg12d1 = "Maintenant vous allez frapper 3 touches" + CRLF + "situ�es au-dessus du groupe des 6 touches." + CRLF2 + "Il s'agit des touches IMPRESSION, Arr�tD�fil, et PAUSE." + CRLF2 + "IMPRESSION, en anglais PrintScreen, permet d'imprimer l'�cran" + CRLF + "ou le document." + CRLF2 + "Arr�tD�fil, en anglais ScrollLock, supprime la possibilit�" + CRLF + "de faire d�filer les pages." + CRLF2 + "PAUSE, en anglais ATTENTION, stoppe l'ex�cution" + CRLF + "du logiciel en cours." + CRLF2 + "Ces commandes n'agissent que si le logiciel actif le permet."
pg12d2 = msgHit + CRLF2 + "Touche IMPRESSION." + CRLF + "Avec la main droite, au-dessus de la touche INSERTION." + CRLF + "Partez de M. Allez � droite hors du clavier principal." + CRLF + "Placez l'index sur la premi�re touche de la rang�e sup�rieure." + CRLF2 + "A droite d'IMPRESSION, le majeur trouve Arr�tD�fil." + CRLF2 + "A sa droite, l'annulaire trouve la PAUSE."

' Suite (sinon procedure too large)
Module_global.main3
End Sub


' **************** MAIN3 **** TRADUIRE � DROITE DU SIGNE �GAL ******************************
Public Sub main3()

' Le�on13
pg13a1 = "Maintenant vous allez apprendre � utiliser les touches" + CRLF + "pour les menus ou les bo�tes de dialogue de Windows." + CRLF2 + "Ces touches sont situ�es � gauche et � droite" + CRLF + "de la grande barre ESPACE." + CRLF2 + "Les touches CONTROL, WINDOWS, ALT sont doubl�es." + CRLF + "On les trouve � gauche et � droite." + CRLF2 + "La touche Menu-Contextuel se trouve seulement � droite." + CRLF2 + "Attention." + CRLF + "L'objectif est de ne faire aucune erreur."
pg13a2 = "PRONONCIATION." + CRLF2 + "Ces touches ne sont pas prononc�es par la synth�se vocale," + CRLF + "sauf MENU-CONTEXTUEL qui est parfois prononc� APPLICATION." + CRLF2 + "C'est seulement le r�sultat de leur action qui est prononc�." + CRLF2 + "Attention. En g�n�ral," + CRLF + "la touche CONTROL stoppe la prononciation du message en cours."
pg13a3 = msgConseils + CRLF2 + "Avec la main gauche." + CRLF2 + "La touche CONTROL." + CRLF + "Avec l'auriculaire. Partez de Q. Descendez tr�s en bas � gauche." + CRLF + "C'est la touche d'angle du clavier." + CRLF2 + "La touche ALT." + CRLF + "Avec le pouce. Partez de la barre ESPACE." + CRLF + "C'est la premi�re touche � sa gauche."
pg13a4 = "Les touches Windows n'existaient pas sur les claviers tr�s anciens." + CRLF + "On les appelle aussi touches LOGO." + CRLF2 + "Attention, normalement elles lancent le menu D�marrer de Windows." + CRLF + "Il faudrait alors presser �chap pour annuler cette action." + CRLF2 + "La touche WINDOWS de gauche." + CRLF + "Avec l'auriculaire. Partez de Q. Descendez  de 2 rang�es." + CRLF + "Elle se trouve entre CONTROL et ALT."
pg13am1 = "Avec la main droite." + CRLF2 + "La touche AltGr." + CRLF + "Elle peut avoir une action diff�rente de celle de la touche ALT." + CRLF + "Avec le pouce. Partez de la barre ESPACE." + CRLF + "C'est la premi�re touche � sa droite." + CRLF2 + "La touche WINDOWS de droite." + CRLF + "Avec l'auriculaire. Partez de M. Descendez de 2 rang�es." + CRLF + "C'est juste � droite de AltGr." + CRLF2 + "La touche CONTROL de droite." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Descendez de 2 rang�es en extension � droite." + CRLF + "C'est la touche d'angle du clavier."
pg13am2 = "Le MENU-CONTEXTUEL." + CRLF + "MENU-CONTEXTUEL lance normalement un menu li� au contexte." + CRLF + "On l'annule par �chap." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Descendez de 2 rang�es et � droite." + CRLF + "C'est juste � gauche du CONTROL-DROIT."

pg13b1 = "Maintenant voici 2 touches qui agissent dans le texte." + CRLF2 + "La touche TABULATION-AVANT, appel�e TAB," + CRLF + "d�place le curseur texte vers la droite sur la m�me ligne." + CRLF + "C'est un d�placement d�fini de quelques caract�res." + CRLF + "La touche TAB n'est pas prononc�e." + CRLF2 + "En mode Majuscule, cette TAB devient une TABULATION-Arri�re" + CRLF + "qui recule le curseur de quelques caract�res vers la gauche." + CRLF2 + "Le RETOUR-Arri�re, en anglais BACKSPACE," + CRLF + "recule le curseur d'un seul caract�re," + CRLF + "effa�ant ainsi le caract�re que vous venez de taper." + CRLF + "En g�n�ral, la synth�se vocale prononce le caract�re effac�."
pg13b2 = msgHit + CRLF2 + "Avec la main gauche." + CRLF + "La touche TAB." + CRLF + "Avec l'auriculaire. Partez de Q." + CRLF + "Allez � la rang�e au-dessus et tr�s � gauche." + CRLF + "C'est une touche plus grande." + CRLF2 + "Avec la main droite." + CRLF + "Le RETOUR-Arri�re." + CRLF + "Avec l'auriculaire. Partez de M." + CRLF + "Montez de 2 rang�es au-dessus, en extension � droite." + CRLF + "C'est la touche plus grande dans le coin du clavier principal."

pg13c1 = "Maintenant vous allez frapper les touches de fonction F1 � F12." + CRLF2 + msgConseils + CRLF2 + "Rep�rez les touches sur la rang�e la plus �loign�e de vous," + CRLF + "dans le clavier principal." + CRLF + "Partez de la rang�e de d�part, en extension maximum." + CRLF + "Frappez uniquement avec le doigt pr�vu."
pg13c2 = "CONSIGNES DE FRAPPE. Avec la main gauche." + CRLF2 + "F1." + CRLF + "Avec l'annulaire. En extension maximum, nettement � gauche." + CRLF + "F2." + CRLF + "Avec l'annulaire. En extension maximum, l�g�rement � droite." + CRLF + "F3." + CRLF + "Avec le majeur. En extension maximum l�g�rement � droite." + CRLF + "F4." + CRLF + "Avec l'index. En extension maximum, l�g�rement � droite." + CRLF + "F5." + CRLF + "Avec l'index. En extension maximum et nettement � droite."
pg13c3 = "Attention. Avec la main droite." + CRLF2 + "F6." + CRLF + "Avec l'index. En extension maximum, directement." + CRLF + "C'est la deuxi�me touche du deuxi�me groupe de 4 touches." + CRLF2 + "F7." + CRLF + "Avec le majeur. En extension maximum, directement." + CRLF2 + "F8." + CRLF + "Avec l'annulaire. En extension maximum, directement." + CRLF2 + "F9 � F12." + CRLF + "Avec l'annulaire. En extension maximum, vers la droite."

pg13d1 = "Un raccourci-clavier est une combinaison de 2 ou 3 touches." + CRLF2 + "Par exemple, si vous tenez une touche CONTROL enfonc�e" + CRLF + "pendant que vous frappez bri�vement une autre touche," + CRLF + "vous ex�cutez une combinaison qui peut agir" + CRLF + "dans l'application que vous utilisez." + CRLF2 + "Par exemple, dans Word, Control tenu F, not� Ctrl+F," + CRLF + "lance la recherche d'une cha�ne de caract�res."
pg13d2 = "Les touches MAJ, Ctrl, Windows, Alt," + CRLF + "suivies par exemple d'une lettre," + CRLF + "sont utilis�es dans les raccourcis clavier." + CRLF2 + "Chaque combinaison a une action diff�rente." + CRLF2 + "Certains raccourcis clavier demandent 2 touches enfonc�es," + CRLF + "avant de frapper bri�vement la touche finale."
pg13d3 = "La touche AltGr est diff�rente." + CRLF2 + "Quand on appuie sur AltGr, � droite de ESPACE," + CRLF + "c'est comme si on appuyait � la fois sur CONTROL et ALT." + CRLF2 + "Pour la combinaison Control+Alt+V," + CRLF + "Il suffit donc de maintenir enfonc� AltGr, puis de taper V."

pg13e1 = "Maintenant vous allez taper 3 caract�res," + CRLF + "� l'aide d'une combinaison de touches," + CRLF + "d�marrant par l'enfoncement de la touche AltGr." + CRLF2 + "Cette frappe ressemble � celle d'un raccourci." + CRLF + "Pourtant il s'agit seulement de caract�res moins accessibles." + CRLF2 + "Dans une autre le�on, on verra une autre m�thode," + CRLF + "avec le pav� num�rique."
pg13e2 = "PRONONCIATION." + CRLF2 + "Le Di�se est prononc� Di�se." + CRLF2 + "La Barre-Oblique-Invers�e, ou Contre-Oblique," + CRLF + "ou encore Antislash, en anglais Backslash," + CRLF + "est parfois prononc�e Contre-Oblique." + CRLF2 + "La Barre-Oblique-Invers�e est souvent utilis�e pour pr�ciser" + CRLF + "l'emplacement hi�rarchique d'un dossier ou d'un fichier." + CRLF2 + "Le Acommercial, en anglais arobace," + CRLF + "est parfois prononc� at, mais plus souvent, arobase." + CRLF2 + "Le Acommercial est souvent utilis� dans les adresses internet."
pg13e3 = msgConseils + CRLF2 + "Appuyez et maintenez d'abord AltGr," + CRLF + "� droite de la barre ESPACE, avec le pouce droit." + CRLF2 + "Maintenez l'appui" + CRLF + "et frappez alors bri�vement la touche souhait�e." + CRLF2 + "Enfin vous pouvez rel�cher l'appui de AltGr." + CRLF2 + "Peu importe la position majuscule ou minuscule."
pg13e4 = "Le Di�se." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "et frappez bri�vement avec l'annulaire gauche" + CRLF + "sur la touche du chiffre 3 ou guillemet." + CRLF2 + "La Barre-Oblique-Invers�e." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "et frappez bri�vement avec l'index sur le chiffre 8 ou soulign�." + CRLF2 + "Le Arobase." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "et frappez bri�vement avec l'annulaire sur le chiffre 0 ou A grave."

pg13f1 = "Maintenant vous allez taper des crochets ou des accolades," + CRLF + "� l'aide d'une combinaison de touches," + CRLF + "d�marrant par l'enfoncement de la touche AltGr." + CRLF2 + "Cette frappe ressemble � celle d'un raccourci." + CRLF + "Pourtant il s'agit seulement de caract�res moins accessibles."
pg13f2 = "PRONONCIATION." + CRLF2 + "Les crochets sont prononc�s crochet gauche ou crochet droit." + CRLF2 + "Les accolades sont prononc�es accolade gauche ou accolade droite." + CRLF2 + "Les crochets et les accolades ressemblent aux parenth�ses," + CRLF + "mais en math�matiques," + CRLF + "ils expriment la hi�rarchie des regroupements."
pg13f3 = msgConseils + CRLF2 + "Appuyez et maintenez d'abord AltGr," + CRLF + "� droite de la barre ESPACE, avec le pouce droit." + CRLF2 + "Maintenez l'appui" + CRLF + "et frappez alors bri�vement la touche souhait�e." + CRLF2 + "Enfin vous pouvez rel�cher l'appui de AltGr." + CRLF2 + "Peu importe la position majuscule ou minuscule."
'12/2011 texte plus court pour pg13f4
pg13f4 = "Le crochet gauche." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez bri�vement avec l'auriculaire droit sur la touche du �." + CRLF2 + "Le crochet droit." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez bri�vement avec l'auriculaire droit sur le tr�ma." + CRLF2 + "L'accolade gauche." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez bri�vement avec l'auriculaire droit sur la touche du �." + CRLF2 + "L'accolade droite." + CRLF + "Maintenez le pouce droit sur AltGr," + CRLF + "frappez bri�vement avec l'auriculaire droit sur la touche du Dollar."

pg13g0 = "Dans les bo�tes de dialogue," + CRLF + "certains champs sont difficiles � remplir." + CRLF2 + "Par exemple, les chemins des fichiers ou des dossiers " + CRLF + "exigent les deux-points et la Barre-Oblique-Invers�e." + CRLF2 + "Certains noms comportent un ESPACE � respecter," + CRLF + "un caract�re tiret ou un caract�re soulign�." + CRLF2 + "Voici un exercice d'entra�nement."

' Le�on14
pg14a0 = "La dict�e sera lue une seule fois." + CRLF + "Puis, la dict�e sera faite phrase par phrase." + CRLF + "On passe � la lettre suivante d�s la deuxi�me erreur." + CRLF2 + "Le mot suivant � taper sera prononc�." + CRLF2 + "Quand on vous demande ""� la ligne""," + CRLF + "appuyez simplement sur la touche Entr�e." + CRLF2 + "Tapez la combinaison CONTROL+ESPACE pour R�P�TER le MOT." + CRLF2 + "Tapez Alt+ESPACE pour �PELER." + CRLF2 + "Tapez MAJ+ESPACE pour R�P�TER la FIN de la PHRASE."
pg14b0 = "Voici une deuxi�me dict�e."
pg14c0 = "Voici une troisi�me dict�e."

' Le�on15
pg15a0 = "La synth�se vocale va parler plus vite." + CRLF2 + "La dict�e sera lue une seule fois." + CRLF + "Puis, la dict�e sera faite phrase par phrase." + CRLF2 + "Le mot suivant � taper sera prononc�." + CRLF2 + "Tapez la combinaison CONTROL+ESPACE pour R�P�TER le MOT." + CRLF2 + "Tapez MAJ+ESPACE pour R�P�TER la PHRASE."
pg15b0 = "Voici une deuxi�me dict�e." + CRLF2 + "La synth�se vocale va parler plus vite."
pg15c0 = "Voici une troisi�me dict�e." + CRLF2 + "La synth�se vocale va parler encore plus vite."

' Le�on16
pg16a1 = "Maintenant vous allez apprendre � utiliser les touches" + CRLF + "du pav� num�rique." + CRLF2 + "Le pav� num�rique, c'est le groupe de 17 touches," + CRLF + "situ� le plus � droite du clavier principal."
pg16a2 = msgConseils + CRLF2 + "Partez du bas compl�tement � droite du clavier." + CRLF + "Montez � la troisi�me rang�e." + CRLF2 + "Une touche porte un point en relief, c'est le chiffre 5." + CRLF + "Placez le majeur sur cette touche 5." + CRLF2 + "Placez l'index � gauche sur le 4," + CRLF + "et placez l'annulaire � droite sur le 6." + CRLF2 + "Pour les 7, 8, 9, c'est la rang�e au-dessus." + CRLF + "Pour 1, 2, 3, c'est la rang�e au-dessous." + CRLF + "Pour le 0, c'est une grande touche encore plus bas sous le 1."
pg16a3 = "ATTENTION." + CRLF2 + "La touche dans le coin en haut et � gauche du pav�" + CRLF + "est une bascule." + CRLF2 + "C'est-�-dire, � chaque appui de cette touche sp�ciale," + CRLF + "on bascule toutes les touches du pav�" + CRLF + "du mode num�rique au mode fl�che," + CRLF + "ou du mode fl�che au mode num�rique." + CRLF2 + "Si la touche 5 ne r�pond pas, par exemple," + CRLF + "appuyez sur la bascule VERROUILLAGE-NUM�RIQUE," + CRLF + "situ�e en haut et � gauche du pav�."
pg16a4 = "RAPPEL." + CRLF + "Utilisez les aides F2 et F3." + CRLF2 + "Utilisez la touche bascule du coin en haut et � gauche du pav�," + CRLF + "si n�cessaire."

pg16b1 = "Maintenant vous allez taper les 5 signes d'op�rations," + CRLF + "avec le pav� num�rique." + CRLF2 + "C'est-�-dire :" + CRLF + "la touche PLUS," + CRLF + "la touche MOINS prononc�e TIRET," + CRLF + "la touche MULTIPLIER prononc�e �TOILE ou AST�RISQUE," + CRLF + "la touche DIVISER prononc�e SLASH ou BARRE-OBLIQUE," + CRLF + "et la touche POINT."
pg16b2 = msgHit + CRLF2 + "Le majeur reste sur le point en relief du 5." + CRLF2 + "La BARRE-OBLIQUE." + CRLF + "Partez du 5. Avec le majeur en extension," + CRLF + "montez � la deuxi�me rang�e au-dessus." + CRLF2 + "L'AST�RISQUE." + CRLF + "Partez du 6. Avec l'annulaire," + CRLF + "montez � la deuxi�me rang�e au-dessus."
pg16b3 = "Le TIRET." + CRLF + "Partez du 6. Avec l'auriculaire en extension," + CRLF + "montez � la deuxi�me rang�e au-dessus et � droite." + CRLF2 + "Le PLUS." + CRLF + "Avec l'annulaire." + CRLF + "C'est la touche � droite du 6." + CRLF2 + "Le POINT." + CRLF + "Partez du 6. Avec l'annulaire," + CRLF + "descendez � la deuxi�me rang�e au-dessous."

pg16c1 = "Chaque caract�re poss�de un code chiffr� pour le repr�senter." + CRLF + "Ce code est appel� nombre Ascii (prononcez ASKI)" + CRLF + "jusqu'� 3 chiffres, ou nombre ANSI � 4 chiffres." + CRLF + "Ce code peut d�pendre des options linguistiques et r�gionales" + CRLF + "s�lectionn�es sur votre ordinateur," + CRLF + "notamment s'il d�passe la valeur 127." + CRLF + "Ceci permet de taper des caract�res d'acc�s malcommode," + CRLF + "ou qui n'existent pas sur le clavier." + CRLF2 + "Il faudra tenir d'abord la touche ALT enfonc�e," + CRLF + "avec le pouce gauche," + CRLF + "et taper le nombre Ascii ou ANSI," + CRLF + "avec le pav� num�rique." + CRLF2 + "Le caract�re appara�tra quand vous rel�cherez la touche ALT."
pg16c2 = "Maintenant vous allez taper les caract�res suivants." + CRLF2 + "Le Di�se (#) : ALT tenu avec nombre 35." + CRLF + "La Barre-Oblique-Invers�e (\): ALT tenu avec nombre 92." + CRLF + "Le Acommercial (@) : ALT tenu avec nombre 64." + CRLF + "Le Tilde (~): ALT tenu avec nombre 126." + CRLF + "Le � (euro) : ALT tenu avec nombre � 4 chiffres 0128."
pg16cm1 = "Maintenant vous allez taper les autres caract�res suivants." + CRLF2 + "Le E aigu majuscule (�) : ALT tenu avec nombre 144." + CRLF + "Le � (ligature du e coll� dans l'o): ALT tenu avec nombre 0156." + CRLF + "Le � (plus ou moins): ALT tenu avec nombre 0177." + CRLF + "Le � : ALT tenu avec nombre 0189."

pg16d1 = "Lorsque la touche VERROUILLAGE-NUM�RIQUE," + CRLF + "dans le coin en haut et � gauche du pav� num�rique," + CRLF + "est bascul�e sur le mode Fl�che," + CRLF + "les touches du pav� deviennent des touches" + CRLF + "de direction du curseur." + CRLF2 + "Par exemple," + CRLF + "le 2 devient FLECHE-BAS," + CRLF + "le 4 devient FLECHE-GAUCHE," + CRLF + "le 6 devient FLECHE-DROITE," + CRLF + "le 8 devient FLECHE-HAUT."
pg16d2 = "Les autres touches de direction sont les suivantes :" + CRLF2 + "D�BUT. Avec l'index. Au-dessus du 4." + CRLF2 + "FIN. Avec l'index. En-dessous du 4." + CRLF2 + "PAGE-PR�C�DENTE. Avec l'annulaire. Au-dessus du 6." + CRLF2 + "PAGE-SUIVANTE. Avec l'annulaire. Au-dessous du 6."
pg16dm1 = "Il y a aussi une touche bascule INSERTION-REMPLACEMENT," + CRLF + "une touche SUPPRESSION, et une touche Entr�e." + CRLF2 + "INSERTION. Avec l'index." + CRLF + "Partez du 4. Descendez de 2 rang�es en-dessous." + CRLF2 + "SUPPRESSION. Avec l'annulaire." + CRLF + "Partez du 6. Descendez de 2 rang�es en-dessous." + CRLF2 + "Entr�e. Avec l'auriculaire. Allez au coin en bas � droite." + CRLF2 + "Pour cet exercice, n'utilisez pas les touches �quivalentes," + CRLF + "situ�es � gauche du pav� num�rique."

' Le�on17
pg17a0 = "Maintenant vous allez utiliser les lettres, les chiffres," + CRLF + "et les ponctuations du clavier principal."
pg17b0 = "Maintenant vous allez utiliser toutes les touches" + CRLF + "du clavier g�n�ral."
pg17c0 = "Maintenant vous allez taper les caract�res et les combinaisons," + CRLF + "en allant plus vite."
pg17d0 = "Maintenant vous allez utiliser tout le pav� num�rique," + CRLF + "en allant plus vite."

' Le�on18
pg18a0 = "Maintenant vous allez taper des mots accentu�s." + CRLF2 + "Attention." + CRLF + "Les substantifs sont au singulier." + CRLF + "Les verbes sont � l'infinitif."
pg18b0 = "Maintenant vous allez taper des mots avec doubles consonnes." + CRLF2 + "Il y aura des majuscules," + CRLF + "et une prononciation difficile."
pg18c0 = "Maintenant vous allez taper des mots dont les terminaisons" + CRLF + "sont courantes." + CRLF2 + "Attention." + CRLF + "La synth�se vocale ne va pas �peler les mots," + CRLF + "sauf si vous faites une erreur." + CRLF2 + "Si vous faites 2 erreurs dans un mot," + CRLF + "ce mot sera � nouveau propos�."
pg18d0 = "Maintenant vous allez taper rapidement des mots," + CRLF + "dont la prononciation est voisine." + CRLF2 + "La synth�se vocale parlera vite, sans �peler."
pg18e0 = "Maintenant vous allez taper des instructions," + CRLF + "dont se servent les programmeurs."

' Le�on19
pg19a0 = "La synth�se vocale va parler plus vite." + CRLF2 + "Le texte sera lu une seule fois." + CRLF + "Puis, la dict�e sera faite phrase par phrase." + CRLF2 + "Le mot suivant � taper sera prononc�." + CRLF2 + "Tapez la combinaison CONTROL+ESPACE pour R�P�TER le MOT." + CRLF2 + "Tapez sur Alt+ESPACE pour �PELER." + CRLF2 + "Tapez MAJ+ESPACE pour R�P�TER la FIN de la PHRASE."
pg19b0 = "Voici un deuxi�me texte." + CRLF2 + "La synth�se vocale va parler plus vite."
pg19c0 = "Voici un troisi�me texte." + CRLF2 + "La synth�se vocale va parler tr�s vite."
pg19d0 = "Voici un quatri�me texte." + CRLF2 + "La synth�se vocale va parler extr�mement vite."

' Menu Editor
meFichier = "&Fichier"
meQuitter_bm = "&Quitter"
meOptions = "&Options"
meStandard = "Niveau &Standard"
mePersonnalis� = "Niveau &Personnalis�"
meDebExpliNormal = "D�bit des explications &Normal"
meDebExpliRapide = "D�bit des explications &Rapide"
meDebGenLent = "D�bit g�n�ral &Lent"
meDebGenMoyen = "D�bit g�n�ral &Moyen"
meDebGenVite = "D�bit g�n�ral &Vite"
meBipClassique = "Bip &Classique"
meBipDiff�rent = "Bip &Diff�rent"
meAide = "&Aide"
meAideG�n�rale = "Aide g�n�rale"
meAideM�moire = "Aide-M�moire"
meEnseignant = "Aide pour l'&Enseignant"
meSonorisation = "Aide pour la &Vocalisation"
meAproposde = "� &Propos de"
meReset = "Red�marrer � la prem&i�re le�on"
meRestart = "Red�marrer � la prem&i�re le�on"

meNoZoom = "Sans z&oom"  ' 12/2011
meWithZoom = "Avec &zoom"
meBasicColors = "Couleurs &basiques"
meOtherColors = "A&utres couleurs"

' Suite
repjawscountry = "\settings\fra\"
Module_routines.inits
End Sub


' *************  HELP_F2  ******  TRADUIRE SEULEMENT � DROITE DE .text4.Text =  ************
Public Sub help_f2(le�on)
With le�on

' D�tecter le Alt255 final �ventuel
If Right(.text1.Text, 1) = "�" Then
    lt1 = Len(.text1.Text) - 1
Else
    lt1 = Len(.text1.Text)
End If

' Reset
.text4.Visible = False
Call Sleep(10) 'attention, pas trop long sinon pb quand AltGr reste enfonc�
pp = ii + 1 - ff
If pp <= 0 Then pp = 1

' Lettres d'1 seul caract�re minuscule
If Mid(.text1.Text, pp, 1) = "a" Then .text4.Text = " a.  Minuscule.   Auriculaire gauche.   Rang�e au-dessus de Q, et l�g�rement � droite.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "b" Then .text4.Text = " b.  Minuscule.   Index gauche.        Rang�e au-dessous de G, et � droite.  Voir le�on 3B."
If Mid(.text1.Text, pp, 1) = "c" Then .text4.Text = " c.  Minuscule.   Majeur gauche.       Rang�e au-dessous de D, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "d" Then .text4.Text = " d.  Minuscule.   Majeur gauche.       Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "e" Then .text4.Text = " e.  Minuscule.   Majeur gauche.       Rang�e au-dessus de D, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "f" Then .text4.Text = " f.  Minuscule.   Index gauche.        Rang�e de d�part, point en relief.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "g" Then .text4.Text = " g.  Minuscule.   Index gauche.         Rang�e de d�part, � droite du F.  Voir le�on 2B."
If Mid(.text1.Text, pp, 1) = "h" Then .text4.Text = " h.  Minuscule.   Index droit.         Rang�e de d�part, � gauche du J.  Voir le�on 2B."
If Mid(.text1.Text, pp, 1) = "i" Then .text4.Text = " i.  Minuscule.   Majeur droit.        Rang�e au-dessus de K, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "j" Then .text4.Text = " j.  Minuscule.   Index droit.         Rang�e de d�part, point en relief.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "k" Then .text4.Text = " k.  Minuscule.   Majeur droit.        Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "l" Then .text4.Text = " l.  Minuscule.   Annulaire droit.     Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "m" Then .text4.Text = " m.  Minuscule.   Auriculaire droit.   Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "n" Then .text4.Text = " n.  Minuscule.   Index droit.         Rang�e au-dessous de J, et � gauche.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "o" Then .text4.Text = " o.  Minuscule.   Annulaire droit.     Rang�e au-dessus de L, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "p" Then .text4.Text = " p.  Minuscule.   Auriculaire droit.   Rang�e au-dessus de M, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "q" Then .text4.Text = " q.  Minuscule.   Auriculaire gauche.  Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "r" Then .text4.Text = " r.  Minuscule.   Index gauche.        Rang�e au-dessus de F, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "s" Then .text4.Text = " s.  Minuscule.   Annulaire gauche.    Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "t" Then .text4.Text = " t.  Minuscule.   Index gauche.        Rang�e au-dessus de F, et � droite.  Voir le�on 2E."
If Mid(.text1.Text, pp, 1) = "u" Then .text4.Text = " u.  Minuscule.   Index droit.         Rang�e au-dessus de J, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "v" Then .text4.Text = " v.  Minuscule.   Index gauche.        Rang�e au-dessous de F, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "w" Then .text4.Text = " w.  Minuscule.   Auriculaire gauche.  Rang�e au-dessous de Q, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "x" Then .text4.Text = " x.  Minuscule.   Annulaire gauche.    Rang�e au-dessous de S, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "y" Then .text4.Text = " y.  Minuscule.   Index droit.         Rang�e au-dessus de H, et � gauche.  Voir le�on 2E."
If Mid(.text1.Text, pp, 1) = "z" Then .text4.Text = " z.  Minuscule.   Annulaire gauche.    Rang�e au-dessus de S, et � gauche.  Voir le�on 2C."

' Lettres d'1 seul caract�re majuscule
If Mid(.text1.Text, pp, 1) = "A" Then .text4.Text = " A.  Majuscule.   Auriculaire gauche.  Rang�e au-dessus de Q, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "B" Then .text4.Text = " B.  Majuscule.   Index gauche.        Rang�e au-dessous de G, et � droite.  Voir le�on 3B."
If Mid(.text1.Text, pp, 1) = "C" Then .text4.Text = " C.  Majuscule.   Majeur gauche.       Rang�e au-dessous de D, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "D" Then .text4.Text = " D.  Majuscule.   Majeur gauche.       Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "E" Then .text4.Text = " E.  Majuscule.   Majeur gauche.       Rang�e au-dessus de D, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "F" Then .text4.Text = " F.  Majuscule.   Index gauche.        Rang�e de d�part, point en relief.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "G" Then .text4.Text = " G.  Majuscule.   Index gauche.         Rang�e de d�part, � droite du F.  Voir le�on 2B."
If Mid(.text1.Text, pp, 1) = "H" Then .text4.Text = " H.  Majuscule.   Index droit.         Rang�e de d�part, � gauche du J.  Voir le�on 2B."
If Mid(.text1.Text, pp, 1) = "I" Then .text4.Text = " I.  Majuscule.   Majeur droit.        Rang�e au-dessus de K, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "J" Then .text4.Text = " J.  Majuscule.   Index droit.         Rang�e de d�part, point en relief.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "K" Then .text4.Text = " K.  Majuscule.   Majeur droit.        Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "L" Then .text4.Text = " L.  Majuscule.   Annulaire droit.     Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "M" Then .text4.Text = " M.  Majuscule.   Auriculaire droit.   Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "N" Then .text4.Text = " N.  Majuscule.   Index droit.         Rang�e au-dessous de J, et � gauche.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "O" Then .text4.Text = " O.  Majuscule.   Annulaire droit.     Rang�e au-dessus de L, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "P" Then .text4.Text = " P.  Majuscule.   Auriculaire droit.   Rang�e au-dessus de M, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "Q" Then .text4.Text = " Q.  Majuscule.   Auriculaire gauche.  Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "R" Then .text4.Text = " R.  Majuscule.   Index gauche.        Rang�e au-dessus de F, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "S" Then .text4.Text = " S.  Majuscule.   Annulaire gauche.    Rang�e de d�part.  Voir le�on 2A."
If Mid(.text1.Text, pp, 1) = "T" Then .text4.Text = " T.  Majuscule.   Index gauche.        Rang�e au-dessus de F, et � droite.  Voir le�on 2E."
If Mid(.text1.Text, pp, 1) = "U" Then .text4.Text = " U.  Majuscule.   Index droit.         Rang�e au-dessus de J, et � gauche.  Voir le�on 2C."
If Mid(.text1.Text, pp, 1) = "V" Then .text4.Text = " V.  Majuscule.   Index gauche.        Rang�e au-dessous de F, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "W" Then .text4.Text = " W.  Majuscule.   Auriculaire gauche.  Rang�e au-dessous de Q, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "X" Then .text4.Text = " X.  Majuscule.   Annulaire gauche.    Rang�e au-dessous de S, et � droite.  Voir le�on 3A."
If Mid(.text1.Text, pp, 1) = "Y" Then .text4.Text = " Y.  Majuscule.   Index droit.         Rang�e au-dessus de H, et � gauche.  Voir le�on 2E."
If Mid(.text1.Text, pp, 1) = "Z" Then .text4.Text = " Z.  Majuscule.   Annulaire gauche.    Rang�e au-dessus de S, et � gauche.  Voir le�on 2C."

' Chiffres
If Module_routines.IsNumLockOn() = "False" Then
    If Mid(.text1.Text, pp, 1) = "1" Then .text4.Text = " Chiffre 1.  Majuscule.  Auriculaire gauche.  2 rang�es au-dessus de Q, et l�g�rement � gauche.  Voir le�on 10A."
    If Mid(.text1.Text, pp, 1) = "2" Then .text4.Text = " Chiffre 2.  Majuscule.  Auriculaire gauche.  2 rang�es au-dessus de Q, et l�g�rement � droite.  Voir le�on 10A."
    If Mid(.text1.Text, pp, 1) = "3" Then .text4.Text = " Chiffre 3.  Majuscule.  Annulaire gauche.   2 rang�es au-dessus de S, et l�g�rement � droite.  Voir le�on 10A."
    If Mid(.text1.Text, pp, 1) = "4" Then .text4.Text = " Chiffre 4.  Majuscule.  Majeur gauche.   2 rang�es au-dessus de D, et l�g�rement � droite.  Voir le�on 10A."
    If Mid(.text1.Text, pp, 1) = "5" Then .text4.Text = " Chiffre 5.  Majuscule.  Index gauche.   2 rang�es au-dessus de F, et l�g�rement � droite.  Voir le�on 10B."
    If Mid(.text1.Text, pp, 1) = "6" Then .text4.Text = " Chiffre 6.  Majuscule.  Index gauche.   2 rang�es au-dessus de F, et en extension � droite.  Voir le�on 10B."
    If Mid(.text1.Text, pp, 1) = "7" Then .text4.Text = " Chiffre 7.  Majuscule.  Index droit.   2 rang�es au-dessus de J, et nettement � gauche.  Voir le�on 10B."
    If Mid(.text1.Text, pp, 1) = "8" Then .text4.Text = " Chiffre 8.  Majuscule.  Index droit.   2 rang�es au-dessus de J, et l�g�rement � droite.  Voir le�on 10C."
    If Mid(.text1.Text, pp, 1) = "9" Then .text4.Text = " Chiffre 9.  Majuscule.  Majeur droit.   2 rang�es au-dessus de K, et l�g�rement � droite.  Voir le�on 10C."
    If Mid(.text1.Text, pp, 1) = "0" Then .text4.Text = " Chiffre 0.  Majuscule.  Annulaire droit.   2 rang�es au-dessus de L, et l�g�rement � droite.  Voir le�on 10C."
Else
    If Mid(.text1.Text, pp, 1) = "1" Then .text4.Text = " Chiffre 1.  Pav� Mode Num�rique.  Index droit.   En-dessous du 5 et � gauche.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "2" Then .text4.Text = " Chiffre 2.  Pav� Mode Num�rique.  Majeur droit.   En-dessous du 5.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "3" Then .text4.Text = " Chiffre 3.  Pav� Mode Num�rique.  Annulaire droit.   En-dessous du 5 et � droite.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "4" Then .text4.Text = " Chiffre 4.  Pav� Mode Num�rique.  Index droit.   A gauche du 5.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "5" Then .text4.Text = " Chiffre 5.  Pav� Mode Num�rique.  Majeur droit.   Au centre du pav�, touche avec relief.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "6" Then .text4.Text = " Chiffre 6.  Pav� Mode Num�rique.  Annulaire droit.   A droite du 5.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "7" Then .text4.Text = " Chiffre 7.  Pav� Mode Num�rique.  Index droit.    Au-dessus du 5 et � gauche.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "8" Then .text4.Text = " Chiffre 8.  Pav� Mode Num�rique.  Index droit.   Au-dessus du 5.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "9" Then .text4.Text = " Chiffre 9.  Pav� Mode Num�rique.  Annulaire droit.   Au-dessus du 5 et � droite.  Voir le�on 16A."
    If Mid(.text1.Text, pp, 1) = "0" Then .text4.Text = " Chiffre 0.  Pav� Mode Num�rique.  Index droit.   2 rang�es en-dessous du 5, et � gauche.  Voir le�on 16A."
End If

' Ponctuations et signes
If Mid(.text1.Text, pp, 1) = " " Then .text4.Text = " ESPACE.   Pouce gauche, ou pouce droit.  Grande barre au-devant du clavier principal.  Voir le�on 1A."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " Au Carr�.  Minuscule.  Auriculaire gauche.  2 rang�es au-dessus de Q, en extension � gauche.  Voir le�on 9A."
If Mid(.text1.Text, pp, 1) = "&" Then .text4.Text = " Et Commercial.  Minuscule.  Auriculaire gauche.  2 rang�es au-dessus de Q, et � gauche.  Voir le�on 9A."
If Mid(.text1.Text, pp, 1) = """" Then .text4.Text = "Guillemet.  Minuscule.   Annulaire gauche.    2 rang�es au-dessus de S, et l�g�rement � droite.  Voir le�on 9B."
If Mid(.text1.Text, pp, 1) = "'" Then .text4.Text = " Apostrophe.  Minuscule.  Majeur gauche.   2 rang�es au-dessus de D, et l�g�rement � droite.  Voir le�on 9B."
If Mid(.text1.Text, pp, 1) = "(" Then .text4.Text = " Parenth�se gauche. Minuscule. Index gauche.  2 rang�es au-dessus de F, et l�g�rement � droite.  Voir le�on 9B."
If Mid(.text1.Text, pp, 1) = ")" Then .text4.Text = " Parenth�se droite.    Minuscule.    Auriculaire droit.     2 rang�es au-dessus de M.  Voir le�on 9C."
If Mid(.text1.Text, pp, 1) = "_" Then .text4.Text = " Soulign�.    Minuscule.     Index droit.      2 rang�es au-dessus de J.  Voir le�on 9C."
If Mid(.text1.Text, pp, 1) = "," Then .text4.Text = " Virgule.  Minuscule.    Index droit.       Rang�e au-dessous de J, et � droite.  Voir le�on 8B."
If Mid(.text1.Text, pp, 1) = "?" Then .text4.Text = " Point d'interrogation.    Majuscule.    Index droit.    Rang�e au-dessous de J, et � droite.  Voir le�on 8C."
If Mid(.text1.Text, pp, 1) = ";" Then .text4.Text = " Point-Virgule.    Minuscule.    Majeur droit.    Rang�e au-dessous de K, et � droite.  Voir le�on 8B."
If Mid(.text1.Text, pp, 1) = ":" Then .text4.Text = " Deux-Points.    Minuscule.    Annulaire droit.     Rang�e au-dessous de L, et � droite.  Voir le�on 8B."
If Mid(.text1.Text, pp, 1) = "!" Then .text4.Text = " Point d'exclamation.   Minuscule.   Auriculaire droit.   Rang�e au-dessous de M, et � droite.  Voir le�on 8B."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " Section.   Majuscule.   Auriculaire droit.   Rang�e au-dessous de M, et � droite.  Voir le�on 8C."
If Mid(.text1.Text, pp, 1) = "<" Then .text4.Text = " Inf�rieur.   Minuscule.   Auriculaire gauche.   Rang�e au-dessous de Q, et l�g�rement � gauche.  Voir le�on 8F."
If Mid(.text1.Text, pp, 1) = ">" Then .text4.Text = " Sup�rieur.   Majuscule.   Auriculaire gauche.   Rang�e au-dessous de Q, et l�g�rement � gauche.  Voir le�on 8F."
If Mid(.text1.Text, pp, 1) = "%" Then .text4.Text = " PourCent.    Majuscule.    Auriculaire droit.    Rang�e de d�part, � droite de M.  Voir le�on 8G."
If Mid(.text1.Text, pp, 1) = "$" Then .text4.Text = " Dollar.   Minuscule.   Auriculaire droit.   Rang�e au-dessus de M, et en extension � droite.  Voir le�on 8G."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " Livre.    Majuscule.    Auriculaire droit.    Rang�e au-dessus de M, et en extension � droite.  Voir le�on 8G."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " Mu, ou Micro.    Majuscule.   Auriculaire droit.    Rang�e de d�part, 2 touches � droite de M.  Voir le�on 8G."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " c c�dille.   Minuscule.    Majeur droit.    2 rang�es au-dessus de K.  Voir le�on 9D."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " Degr�.    Majuscule.    Auriculaire droit.    2 rang�es au-dessus de M.  Voir le�on 9D."
If Mid(.text1.Text, pp, 1) = "=" Then .text4.Text = " �gal.    Minuscule.    Auriculaire droit.    2 rang�es au-dessus de M, nettement � droite.  Voir le�on 9D."
If Mid(.text1.Text, pp, 1) = "^" Then .text4.Text = " Circonflexe.  Minuscule.  Auriculaire droit.  Au-dessus de M, et � droite.  Tapez ensuite la voyelle.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " Tr�ma.  Majuscule.  Auriculaire droit.  Au-dessus de M, et � droite.  Tapez ensuite la voyelle.  Voir le�on 8E."

' Caract�res 4 op�rations et POINT selon numpad ?
If Mid(.text1.Text, pp, 1) = "+" Then .text4.Text = " Plus.    Majuscule.    Auriculaire droit.    2 rang�es au-dessus de M, nettement � droite.  Voir le�on 16B."
If Mid(.text1.Text, pp, 1) = "-" Then .text4.Text = " Tiret.    Minuscule.    Index gauche.    2 rang�es au-dessus de F, et en extension � droite.  Voir le�on 16B."
If Mid(.text1.Text, pp, 1) = "*" Then .text4.Text = " Ast�risque.    Minuscule.   Auriculaire droit.    Rang�e de d�part, 2 touches � droite de M.  Voir le�on 16B."
If Mid(.text1.Text, pp, 1) = "/" Then .text4.Text = " Barre-Oblique.    Majuscule.    Annulaire droit.      Rang�e au-dessous de L, et � droite.  Voir le�on 16B."
If Mid(.text1.Text, pp, 1) = "." Then .text4.Text = " Point.    Majuscule.    Majeur droit.    Rang�e au-dessous de K, et � droite.  Voir le�on 16B."

' caract�res command�s par AltGr
If numpad <= 0 Then
    If Mid(.text1.Text, pp, 1) = "~" Then .text4.Text = " Tilde.  Pouce droit sur AltGr tenu, puis avec l'auriculaire gauche, touche du chiffre 2.  Voir le�on 13E."
    If Mid(.text1.Text, pp, 1) = "#" Then .text4.Text = " Di�se.  Pouce droit sur AltGr tenu, puis avec l'annulaire gauche, tapez sur le guillemet.  Voir le�on 13E."
    If Mid(.text1.Text, pp, 1) = "{" Then .text4.Text = " Accolade gauche.  Pouce droit sur AltGr tenu, puis avec le majeur gauche, touche du chiffre 4.  Voir le�on 13F."
    If Mid(.text1.Text, pp, 1) = "[" Then .text4.Text = " Crochet gauche.  Pouce droit sur AltGr tenu, puis avec l'index gauche, touche du chiffre 5.  Voir le�on 13F."
    If Mid(.text1.Text, pp, 1) = "|" Then .text4.Text = " Barre-Verticale.  Pouce droit sur AltGr tenu, puis avec l'index gauche, touche du chiffre 6." & "                 "
    If Mid(.text1.Text, pp, 1) = "`" Then .text4.Text = " Accent grave.  Pouce droit sur AltGr tenu, puis tapez sur la touche du chiffre 7." & "                 "
    If Mid(.text1.Text, pp, 1) = "\" Then .text4.Text = " Barre-Oblique-Invers�e.  Pouce droit sur AltGr tenu, puis tapez sur le soulign�, touche chiffre 8.  Voir le�on 13E."
    If Mid(.text1.Text, pp, 1) = "@" Then .text4.Text = " A Commercial.  Pouce droit sur AltGr tenu, puis avec l'annulaire droit, tapez sur le a grave.  Voir le�on 13E."
    If Mid(.text1.Text, pp, 1) = "]" Then .text4.Text = " Crochet droit.  Pouce droit sur AltGr tenu, puis avec l'auriculaire droit, touche de la parenth�se.  Voir le�on 13F."
    If Mid(.text1.Text, pp, 1) = "}" Then .text4.Text = " Accolade droite.  Pouce droit sur AltGr tenu, puis avec l'auriculaire droit, touche du signe �gal.  Voir le�on 13F."
End If

' caract�res command�s par Alt+nombre-Ascii-Ansi
If numpad >= 1 Then
    If Mid(.text1.Text, pp, 1) = "#" Then .text4.Text = " Di�se.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 35.  Voir le�on 16C."
    If Mid(.text1.Text, pp, 1) = "@" Then .text4.Text = " A Commercial.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 64.  Voir le�on 16C."
    If Mid(.text1.Text, pp, 1) = "\" Then .text4.Text = " Barre-Oblique-Invers�e.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 92.  Voir le�on 16C."
    If Mid(.text1.Text, pp, 1) = "~" Then .text4.Text = " Tilde.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 126.  Voir le�on 16C."
End If
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " E aigu Majuscule.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 144.  Voir le�on 16C."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " �.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 0128.  Voir le�on 16C."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " �.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 0177.  Voir le�on 16C."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " �.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 0189.  Voir le�on 16C."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " �.  Pouce gauche sur Alt tenu, puis au pav� num�rique tapez le nombre 0156.  Voir le�on 16C."

' Lettres accentu�es
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " a grave.   Minuscule.   Annulaire droit.    2 rang�es au-dessus de L.  Voir le�on 8H."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " e aigu.   Minuscule.   Auriculaire gauche.    2 rang�es au-dessus de Q, et l�g�rement � droite.  Voir le�on 8H."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " e grave.   Minuscule.     Index droit.    2 rang�es au-dessus de J, et vers la gauche.  Voir le�on 8H."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " u grave.   Minuscule.   Auriculaire droit.    Rang�e de d�part, � droite de M.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " a circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et � droite du M, avant le a.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " e circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et � droite du M, avant le e.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " i circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et � droite du M, avant le i.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " o circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et � droite du M, avant le o.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " u circonflexe.   Minuscule.   Tapez le circonflexe au-dessus et � droite du M, avant le u.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " a tr�ma.    a minuscule, tr�ma majuscule.   Tapez le tr�ma au-dessus et � droite du M, avant le a.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " e tr�ma.    e minuscule, tr�ma majuscule.   Tapez le tr�ma au-dessus et � droite du M, avant le e.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " i tr�ma.    i minuscule, tr�ma majuscule.   Tapez le tr�ma au-dessus et � droite du M, avant le i.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " o tr�ma.    o minuscule, tr�ma majuscule.   Tapez le tr�ma au-dessus et � droite du M, avant le o.  Voir le�on 8E."
If Mid(.text1.Text, pp, 1) = "�" Then .text4.Text = " u tr�ma.    u minuscule, tr�ma majuscule.   Tapez le tr�ma au-dessus et � droite du M, avant le u.  Voir le�on 8E."

End With

' Lancer la suite dans une autre procedure (sinon message procedure too large)
help_f2_suite le�on
End Sub


' ************  HELP_F2_SUITE  ******  TRADUIRE SEULEMENT � DROITE DE .text4.Text =  *********
Public Sub help_f2_suite(le�on)
With le�on

' Raccourcis
If UCase(Left(.text1.Text, Len(vvMaj) + 1)) = UCase(vvMaj) & "+" Then
    .text4.Text = " MAJUSCULE tenue enfonc�e. Auriculaire. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvMaj) - 1) & ".  Voir le�on 13D."
End If
If UCase(Left(.text1.Text, Len(vvControl) + 1)) = UCase(vvControl) & "+" Then
    .text4.Text = " CONTROL tenu enfonc�. Auriculaire. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvControl) - 1) & ".  Voir le�on 13D."
End If
If UCase(Left(.text1.Text, Len(vvAlt) + 1)) = UCase(vvAlt) & "+" Then
    .text4.Text = " ALT tenu enfonc�. Pouce gauche. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvAlt) - 1) & ".  Voir le�on 13D."
End If
If UCase(Left(.text1.Text, Len(vvControl) + Len(vvMaj) + 2)) = UCase(vvControl) & "+" & UCase(vvMaj) & "+" Then
    .text4.Text = " CONTROL et MAJ tenus enfonc�s. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvControl) - Len(vvMaj) - 2) & ".  Voir le�on 13D."
End If
If UCase(Left(.text1.Text, Len(vvControl) + Len(vvAlt) + 2)) = UCase(vvControl) & "+" & UCase(vvAlt) & "+" Then
    .text4.Text = " CONTROL et ALT tenus enfonc�s. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvControl) - Len(vvAlt) - 2) & ".  Voir le�on 13D."
End If
If UCase(Left(.text1.Text, Len(vvMaj) + Len(vvAlt) + 2)) = UCase(vvMaj) & "+" & UCase(vvAlt) & "+" Then
    .text4.Text = " MAJUSCULE et ALT tenus enfonc�s. Puis tapez la touche " & Right(.text1.Text, Len(.text1.Text) - Len(vvMaj) - Len(vvAlt) - 2) & ".  Voir le�on 13D."
End If

' Textes de plusieurs caract�res, attention, placer ces lignes de code apr�s celles d�di�es � 1 caract�re et apr�s celles des raccourcis
If UCase(Left(.text1.Text, lt1)) = UCase(vvEspace) Then .text4.Text = " ESPACE.   Pouce gauche ou pouce droit.    Grande barre au-devant du clavier principal.  Voir le�on 1A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvControl) Then .text4.Text = " CONTROL. Auriculaire gauche ou droit.   Coins gauche et droit en bas du clavier principal.  Voir le�on 1C."
If Left(.text1.Text, lt1) = vvControlGauche Then .text4.Text = " CONTROL-GAUCHE.  Auriculaire gauche.  Touche du coin gauche en bas du clavier principal.  Voir le�on 1C."
If Left(.text1.Text, lt1) = vvControlDroit Then .text4.Text = " CONTROL-DROIT.   Auriculaire droit.   Touche du coin droit en bas du clavier principal.  Voir le�on 1C."
If Left(.text1.Text, lt1) = vvWindowsGauche Then .text4.Text = " WINDOWS-GAUCHE.   Auriculaire gauche.   2 rang�es en-dessous de Q.  Voir le�on 13A."
If Left(.text1.Text, lt1) = vvWindowsDroit Then .text4.Text = " WINDOWS-DROIT.    Auriculaire droit.    2 rang�es en-dessous de M.  Voir le�on 13A."
If Left(.text1.Text, lt1) = vvMenuContextuel Then .text4.Text = " MENU-CONTEXTUEL.   Auriculaire droit.   2 rang�es en-dessous de M, en extension � droite.  Voir le�on 13A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvAlt) Then .text4.Text = " ALT.   Pouce gauche.    Touche � gauche de la barre ESPACE.  Voir le�on 13A."
If Left(.text1.Text, lt1) = vvAltGr Then .text4.Text = " AltGr.   Pouce droit.    Touche � droite de la barre ESPACE.  Voir le�on 13A."
If Left(.text1.Text, lt1) = vvAltOuAltGr Then .text4.Text = " ALT Pouce gauche.   AltGr Pouce droit.   A gauche ou � droite de la barre ESPACE.  Voir le�on 13A."
If UCase(Left(.text1.Text, lt1)) = UCase(vv�chap) Then .text4.Text = " �CHAP.    Auriculaire gauche.     Touche au coin � gauche en haut du clavier.  Voir le�on 1A."
If Left(.text1.Text, lt1) = vvVerrouillageMajuscules Then .text4.Text = " VERROUILLAGE-MAJUSCULES.   Auriculaire gauche.   Rang�e de d�part � gauche de Q.  Voir le�on 8A."
If Left(.text1.Text, lt1) = vvVerrouillageNum�rique Then .text4.Text = " VERROUILLAGE-NUM�RIQUE. Index droit. Coin en haut et � gauche, dans le pav� num�rique.  Voir le�on 16A."
If UCase(Left(.text1.Text, lt1)) = UCase(vvMaj) Then .text4.Text = vvMaj
If Left(.text1.Text, lt1) = vvMajGauche Then .text4.Text = " MAJ-GAUCHE.   Auriculaire gauche.     Rang�e au-dessous de Q et tr�s � gauche.  Voir le�on 8A."
If Left(.text1.Text, lt1) = vvMajDroit Then .text4.Text = " MAJ-DROIT.   Auriculaire droit.     Rang�e au-dessous de M en extension � droite.  Voir le�on 8A."
If LCase(Left(.text1.Text, lt1)) = LCase(vvRetourArri�re) Then .text4.Text = " RETOUR-Arri�re.    Annulaire droit.    Clavier principal, coin en haut et � droite.  Voir le�on 13B."
If UCase(Left(.text1.Text, lt1)) = UCase(vvTab) Then .text4.Text = " TABULATION.   Auriculaire gauche.   Au-dessus du Q, et tr�s � gauche.  Voir le�on 13B."
If Left(.text1.Text, lt1) = vvTabulationAvant Then .text4.Text = " TABULATION-AVANT. Minuscule. Auriculaire gauche. Au-dessus du Q, et tr�s � gauche.  Voir le�on 13B."
If Left(.text1.Text, lt1) = vvTabulationArri�re Then .text4.Text = " TABULATION-Arri�re. Majuscule. Auriculaire gauche. Au-dessus du Q, et tr�s � gauche.  Voir le�on 13B."
If Left(.text1.Text, lt1) = vvD�but Then .text4.Text = " D�BUT, ou HOME.     Majeur droit.      Milieu rang�e sup�rieure, groupe des 6.  Voir le�on 12B."
If Left(.text1.Text, lt1) = vvFin Then .text4.Text = " FIN, ou END.    Majeur droit.     Milieu rang�e inf�rieure, groupe des 6.  Voir le�on 12B."
If Left(.text1.Text, lt1) = vvImpression Then .text4.Text = " IMPRESSION.   Majeur droit.   Rang�e au-dessus du groupe des 6, au-dessus de INSERTION.  Voir le�on 12D."
If Left(.text1.Text, lt1) = vvArr�tD�fil Then .text4.Text = " Arr�tD�fil.   Index droit.   Rang�e au-dessus du groupe des 6, au-dessus de D�BUT.  Voir le�on 12D."
If UCase(Left(.text1.Text, lt1)) = vvPause Then .text4.Text = " PAUSE.   Annulaire droit.   Rang�e au-dessus du groupe des 6, et de PAGE-PR�C�DENTE.  Voir le�on 12D."

' Touches de fonction
If Left(.text1.Text, lt1) = "F1" Then .text4.Text = "  F1.   Annulaire gauche.   3 rang�es au-dessus de S, nettement � gauche.  Voir le�on 1B."
If Left(.text1.Text, lt1) = "F2" Then .text4.Text = "  F2.   Annulaire gauche.   3 rang�es au-dessus de S, l�g�rement � droite.  Voir le�on 1B."
If Left(.text1.Text, lt1) = "F3" Then .text4.Text = "  F3.   Majeur gauche.   3 rang�es au-dessus de D, l�g�rement � droite.  Voir le�on 1B."
If Left(.text1.Text, lt1) = "F4" Then .text4.Text = "  F4.   Index gauche.   3 rang�es au-dessus de F, l�g�rement � droite.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F5" Then .text4.Text = "  F5.   Index gauche.   3 rang�es au-dessus de F, en extension � droite.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F6" Then .text4.Text = "  F6.   Index droit.   3 rang�es au-dessus de J, en extension.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F7" Then .text4.Text = "  F7.   Majeur droit.   3 rang�es au-dessus de K, en extension.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F8" Then .text4.Text = "  F8.   Annulaire droit.   3 rang�es au-dessus de L, en extension.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F9" Then .text4.Text = "  F9.   Auriculaire droit.   3 rang�es au-dessus de M, en extension.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F10" Then .text4.Text = "  F10.   Auriculaire droit.   3 rang�es au-dessus de M, en extension � droite.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F11" Then .text4.Text = "  F11.   Auriculaire droit.   3 rang�es au-dessus de M, en extension tr�s � droite.  Voir le�on 13C."
If Left(.text1.Text, lt1) = "F12" Then .text4.Text = "  F12.   Auriculaire droit.   3 rang�es au-dessus de M, en extension extr�me � droite.  Voir le�on 13C."

' Touches de m�me nom au clavier principal et au pav� num�rique
If numpad = 0 Then
    If UCase(Left(.text1.Text, lt1)) = UCase(vvEntr�e) Then .text4.Text = " Entr�e.   Auriculaire droit.    Grande touche sur le bord droit du clavier principal.  Voir le�on 1A."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvInsertion) Then .text4.Text = " INSERTION REMPLACEMENT BASCULE. Index droit. Coin en haut � gauche, groupe des 6.  Voir le�on 12A."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvSuppression) Then .text4.Text = " SUPPRESSION.     Index droit.     Coin en bas � gauche, groupe des 6.  Voir le�on 12A."
    If Left(.text1.Text, lt1) = vvFlecheGauche Then .text4.Text = " FLECHE-GAUCHE. Index droit. Premi�re touche en bas � droite du clavier principal.  Voir le�on 1B."
    If Left(.text1.Text, lt1) = vvFlecheBas Then .text4.Text = " FLECHE-BAS. Majeur droit. Deuxi�me touche en bas � droite du clavier principal.  Voir le�on 1B."
    If Left(.text1.Text, lt1) = vvFlecheDroite Then .text4.Text = " FLECHE-DROITE. Annulaire droit. Troisi�me touche en bas � droite du clavier principal.  Voir le�on 1B."
    If Left(.text1.Text, lt1) = vvFlecheHaut Then .text4.Text = " FLECHE-HAUT.    Majeur droit.     Au-dessus de la touche FLECHE-BAS.  Voir le�on 1B."
    If Left(.text1.Text, lt1) = vvPagePr�c�dente Then .text4.Text = " PAGE-PR�C�DENTE.   Annulaire droit.    Coin en haut � droite, groupe des 6.  Voir le�on 12C."
    If Left(.text1.Text, lt1) = vvPageSuivante Then .text4.Text = " PAGE-SUIVANTE.    Annulaire droit.    Coin en bas � droite, groupe des 6.  Voir le�on 12C."
End If
If numpad = 1 Or numpad = -1 Then
    If UCase(Left(.text1.Text, lt1)) = UCase(vvEntr�e) Then .text4.Text = " Entr�e.   Pav� Num�rique.    Auriculaire droit.   Coin � droite et en bas.  Voir le�on 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvInsertion) Then .text4.Text = " INSERTION REMPLACEMENT BASCULE. Pav� Num�rique. Index droit. Coin en bas � gauche.  Voir le�on 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvSuppression) Then .text4.Text = " SUPPRESSION.  Pav� Num�rique. Annulaire droit. 2 rang�es au-dessous et � droite du 5.  Voir le�on 16D."
    If Left(.text1.Text, lt1) = vvFlecheGauche Then .text4.Text = " FLECHE-GAUCHE. Pav� Num�rique, mode fl�che. Index droit. Premi�re touche � gauche du 5.  Voir le�on 16D."
    If Left(.text1.Text, lt1) = vvFlecheBas Then .text4.Text = " FLECHE-BAS. Pav� Num�rique, mode fl�che. Majeur droit. En-dessous du 5.  Voir le�on 16D."
    If Left(.text1.Text, lt1) = vvFlecheDroite Then .text4.Text = " FLECHE-DROITE. Pav� Num�rique, mode fl�che. Annulaire droit. Premi�re touche � droite du 5.  Voir le�on 16D."
    If Left(.text1.Text, lt1) = vvFlecheHaut Then .text4.Text = " FLECHE-HAUT.  Pav� num�rique, mode fl�che.  Majeur droit.  Au-dessus du 5.  Voir le�on 16D."
    If Left(.text1.Text, lt1) = vvPagePr�c�dente Then .text4.Text = " PAGE-PR�C�DENTE. Pav� Num�rique, mode fl�che. Annulaire droit. Au-dessus et � droite du 5.  Voir le�on 16D."
    If Left(.text1.Text, lt1) = vvPageSuivante Then .text4.Text = " PAGE-SUIVANTE. Pav� Num�rique, mode fl�che. Annulaire droit. Au-dessous et � droite du 5.  Voir le�on 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvD�but) Then .text4.Text = " D�BUT.   Pav� Num�rique, mode fl�che.   Index droit.   Au-dessus et � gauche du 5.  Voir le�on 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvFin) Then .text4.Text = " FIN.   Pav� Num�rique, mode fl�che.   Index droit.   Au-dessous et � gauche du 5.  Voir le�on 16D."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvPlus) Then .text4.Text = " PLUS.   Pav� Num�rique.    Auriculaire droit.   Au-dessus et � droite du 6.  Voir le�on 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvTiret) Then .text4.Text = " TIRET ou MOINS.   Pav� Num�rique.    Auriculaire droit.   Coin en haut � droite.  Voir le�on 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvMoins) Then .text4.Text = " TIRET ou MOINS.   Pav� Num�rique.    Auriculaire droit.   Coin en haut � droite.  Voir le�on 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvBarreOblique) Then .text4.Text = " BARRE-OBLIQUE.   Pav� Num�rique.    Majeur droit.   2 rang�es au-dessus du 5.  Voir le�on 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvAst�risque) Then .text4.Text = " AST�RISQUE.   Pav� Num�rique.    Annulaire droit.   2 rang�es au-dessus du 6.  Voir le�on 16B."
    If UCase(Left(.text1.Text, lt1)) = UCase(vvPoint) Then .text4.Text = " POINT.   Pav� Num�rique.    Annulaire droit.   2 rang�es en-dessous du 6.  Voir le�on 16B."
End If

' Cas particulier "� la ligne" (Alt255 devant pour visibilit�)
If .text3.Text = "�" & vvAlaligne Then .text4.Text = " � la ligne, ou Entr�e.   Auriculaire droit.    Grande touche sur le bord droit du clavier principal.  Voir le�on 1A."

' SUPPRIMER la MENTION "VOIR le�onxxA" sauf en mode F3 aide-m�moire (mais variable avecf3 inutilisable)
If avecf2 = 1 Then
    On Error Resume Next
    .text4.Text = Left(.text4.Text, Len(.text4.Text) - Len("Voir le�on 19A."))  '�limine les mentions voir le�on xx
End If

' Set
.text4.SelStart = 0
.text4.SelLength = Len(.text4.Text)
.text4.Visible = True

' En dehors du mode aide-m�moire
If avecf3 = 0 Then
    keyinhibit = 2
    f2link = 1
End If
End With
End Sub
