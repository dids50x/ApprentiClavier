; Script de janvier 2006
Include "hjconst.jsh" ; DEFAULT HJ constants
;Include "common.jsm" ; message file

globals
	int BasicVoiceSpeed,
	int MinVoiceRate,
	int MaxVoiceRate,
	int Offset,
	int UserSpeed,
	int OffExpli,
	int OffGen,
	int GlobalVoiceChanges,
	int iWin95,
	int WindowClosed,
	int MemoSpeed


; ********************************************************************************************
void Function AutoStartEvent ()
; UserSpeed vitesse de base réglée à 45 par défaut (réglable entre 30 et 60).
let UserSpeed=45

; Offset-OffsetExpli-OffsetGen sont utilisés par les options "Débit Explications" et "Débit Général".
let Offset=0
let Offexpli=9
let OffGen=6

; Eviter l'autostart dû à un appel de la touche Windows pendant une leçon
if (GetWindowName(GetRealWindow (GetFocus())) == "Leçon") then
	; rien
else
	; Echo clavier seulement pour Bienvenue
	SetJcfOption(OPT_TYPING_ECHO,1)
	; Peu de ponctuations en entrant dans ApprentiClavier
	SetJcfOption(OPT_PUNCTUATION,1)
	; Sauvegarder les réglages de base de la voix en entrant dans ApprentiClavier
	InitializeGlobalVoiceSettings(V_RATE)
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset)
endif
EndFunction


; ********************************************************************************************
Void Function WindowCreatedEvent (handle hWindow, int nLeft, int nTop, int nRight, int nBottom)
; Vitesse sono, et ponctuation, selon la mise en forme du titre de la fenêtre (caption)

; Pour "Erreur-Recommençons" dans Bienvenue
if (GetWindowName(GetRealWindow (GetFocus())) == "Bienvenue") then
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,1)

; Pour "débit des explications normal" amenant le titre commençant à gauche sans blancs
elif (GetWindowName(GetRealWindow (GetFocus())) == "") then
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)

; Pour "débit des explications rapide" amenant le titre commençant par 3 blancs
elif (GetWindowName(GetRealWindow (GetFocus())) == "   ") then
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset + Offexpli)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)

; Pour "débit général LENT"
elif (GetWindowName(GetRealWindow (GetFocus())) == "Lent") then
	let Offset=0 - OffGen
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)

; Pour "débit général MOYEN"
elif (GetWindowName(GetRealWindow (GetFocus())) == "Moyen") then
	let Offset=0
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)

; Pour "débit général VITE"
elif (GetWindowName(GetRealWindow (GetFocus())) == "Vite") then
	let Offset=OffGen
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)

; Pour débit ajusté par la mise en forme du titre MENU PRINCIPAL, selon le "Débit général"
elif (GetWindowName(GetRealWindow (GetFocus())) == "Menu Principal") then
	let Offset=0 - OffGen
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Menu  Principal") then
	let Offset=0
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Menu    Principal") then
	let Offset=OffGen
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)

; Pour débit ajusté par la mise en forme du titre MENU Leçon, selon le "Débit général"
elif (GetWindowName(GetRealWindow (GetFocus())) == "Menu Leçon") then
	let Offset=0 - OffGen
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Menu  Leçon") then
	let Offset=0
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Menu    Leçon") then
	let Offset=OffGen
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,1 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)

; Pour la leçon 13 F
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 13 F.") then
	SetJcfOption(OPT_PUNCTUATION,4)

; Pour les dictées 14 et 15
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 14") then
	ValueVoiceSetting(V_RATE,False,-10 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 14") then
	let MemoSpeed=-10 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,-10 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 15 A") then
	ValueVoiceSetting(V_RATE,False,-5 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 15 A") then
	let MemoSpeed=-5 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,-5 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 15 B") then
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 15 B") then
	let MemoSpeed=UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 15 C") then
	ValueVoiceSetting(V_RATE,False,5 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 15 C") then
	let MemoSpeed=5 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,5 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)

; Pour la leçon 18D
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 18 D") then
	let MemoSpeed=5 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,5 + UserSpeed + Offset)

; Pour les dictées 19
elif (GetWindowName(GetRealWindow(GetFocus())) == "Dictée 19 A") then
	ValueVoiceSetting(V_RATE,False,8 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow(GetFocus())) == "Leçon 19 A") then
	let MemoSpeed=8 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,8 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 19 B") then
	ValueVoiceSetting(V_RATE,False,10 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 19 B") then
	let MemoSpeed=10 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,10 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 19 C") then
	ValueVoiceSetting(V_RATE,False,12 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 19 C") then
	let MemoSpeed=12 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,12 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Dictée 19 D") then
	ValueVoiceSetting(V_RATE,False,14 + UserSpeed + Offset)
elif (GetWindowName(GetRealWindow (GetFocus())) == "Leçon 19 D") then
	let MemoSpeed=14 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,14 + UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,4)

; Pour les autres cas, vitesse, ponctuation STANDARD (attention memoSpeed = setting+1)
else
	let MemoSpeed=1 + UserSpeed + Offset
	ValueVoiceSetting(V_RATE,False,UserSpeed + Offset)
	SetJcfOption(OPT_PUNCTUATION,1)
	SetJcfOption(OPT_TYPING_ECHO,0)
endif
EndFunction


; ********************************************************************************************
Function WindowDestroyedEvent (handle hWindow)
; Pour reprendre la vitesse des leçons-dictées, malgré un appel de F1
if (MemoSpeed!=UserSpeed + Offset) then
	ValueVoiceSetting(V_RATE,False,MemoSpeed + Offset)
endif
PcCursor()
EndFunction


; ********* GRAB current voice parameter setting, usual DEFAULT ***************
; *** Int Function GetSettingInformation (int Setting, string ContextName, 
; ***	     Int ByRef MinSetting, Int ByRef MaxSetting)


; ************** SET current voice parameters, usual DEFAULT  *******************
; *** Void Function SetVoiceSetting (int ParameterToSet, int Setting, string ContextName, 
; ***       int UpOrDown, int InSayAll)


; *********** VALUE (entre 0 et 100) for current voice parameters, modified  *********
Void Function ValueVoiceSetting (int iSetting, int InSayAll, int ValidValue)
var
	string ContextName,
	int CurrentSetting,
	int MinSetting,
	int MaxSetting,
	int iDirection
If IsJAWSCursor () then
	let ContextName = VCTX_JAWSCURSOR
else
	let ContextName = VCTX_PCCURSOR
endIf
let CurrentSetting = GetSettingInformation (iSetting, ContextName, MinSetting, MaxSetting)
;if MaxSetting<=50 then
;	saystring("50")
;endif
let CurrentSetting = ValidValue * (MaxSetting - MinSetting) / 100
if (CurrentSetting < MinSetting) then
	let CurrentSetting = MinSetting
endif
if (CurrentSetting > MaxSetting) then
	let CurrentSetting = MaxSetting
endif
if (ValidValue >= 50) then
	let iDirection = V_UP
else
	let iDirection = V_DOWN
endif
SetVoiceSetting (iSetting, CurrentSetting, ContextName, iDirection, InSayAll)
EndFunction


; *********** JawsBACKSPACE modified, pour éviter réponse "Espace" en Jaws4.01  *********
Script JawsBackspace()
;TypeKey (cKs3)
TypeKey ("Backspace")
EndScript

