#SingleInstance force 			; only one instance of this script may run at a time!
#NoEnv  						; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Warn  							; Enable warnings to assist with detecting common errors.
#Persistent
SendMode Input  				; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%		; Ensures a consistent starting directory.

Menu, Tray,Icon, % A_SCriptDir . "\vh.ico" 



;------------------ SECTION OF GLOBAL VARIABLES: BEGINNING ---------------------------- 
global English_USA 		:= 0x0409   ; see AutoHotkey help: Language Codes https://www.autohotkey.com/docs/misc/Languages.htm
, PolishLanguage 		:= 0x0415	; https://www.autohotkey.com/docs/misc/Languages.htm 
, TransFactor 			:= 255
, WordTrue 				:= -1 ; ComObj(0xB, -1) ; 0xB = VT_Bool || -1 = true
, WordFalse 			:= 0 ; ComObj(0xB, 0) ; 0xB = VT_Bool || 0 = false
, OurTemplateEN 		:= "S:\OrgFirma\Szablony\Word\OgolneZmakrami\TQ-S440-en_UserDoc.dotm"
, OurTemplatePL 		:= "S:\OrgFirma\Szablony\Word\OgolneZmakrami\TQ-S440-pl_DokUzyt.dotm"
, OurTemplate 			:= ""
;---------------- Zmienne do funkcji autozapisu ----------------
, flag_as 				:= 0
, size 					:= 0
, size_org 				:= 0
, table					:= []
, AutosaveFilePath		:= "C:\temp1\KopiaZapasowaPlikowWord\"
, interval 				:= 10*60*1000	;10 min.
;--------------- Zmienne do przełączania okienek ---------------
, cntWnd 				:= 0
, cntWnd2 				:= 0
, id					:= []
, order 				:= []
;---------------------------------------------------------------
, MyTemplate 			:= ""
, template 				:= ""
, ToRemember 			:= "", OldClipBoard := ""
;------------------- SECTION OF GLOBAL VARIABLES: END---------------------------- 


#Include, *i ..\Otagle3\WarstwaWord\MakraOgolne\SetHeadersAndFooters.ahk
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\Wypunktowania.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\UsunWielokrotneSpacje.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\Refresh.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\TwardaSpacja.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\Hiperlacza.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\FindBlad.ahk
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\FindDeadLinks.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\ResizeImages.ahk 
#Include, *i ..\Otagle3\WarstwaWord\UstawieniaDokumentu\CheckingMacro.ahk 


SetTimer, AutoSave, % interval	;positive, the timer will automatically repeat until it is explicitly disabled by the script; counts down
TrayTip, %A_ScriptName%, % "Autozapis dokumentów w MS Word włączony.`nAby wyłączyć tę funkcję, naciśnij kombinację klawiszy Ctrl+LewyAlt+Q.", 5, 0x1	;0x1 = Info icon

;/////////////////////////////// - INI SECTION - //////////////////////////////////////




IniRead, ParenthesiIni, 		VariousFunctions.ini, 	Menu memory, Parenthesis, 			NO
IniRead, BrowserIni, 			VariousFunctions.ini, 	Menu memory, Browser, 				NO
IniRead, SetEnglishKeyboardIni, VariousFunctions.ini, 	Menu memory, Set English Keyboard, 	NO
IniRead, AltGrIni, 				VariousFunctions.ini, 	Menu memory, AltGr, 				NO
IniRead, WindowSwitcherIni, 	VariousFunctions.ini, 	Menu memory, Window `Switcher, 		NO
IniRead, PrintScreenIni, 		VariousFunctions.ini, 	Menu memory, Print`Screen, 			NO
IniRead, CapitalizeIni, 		VariousFunctions.ini, 	Menu memory,`Capitalize, 			NO
IniRead, MicrosoftWordIni, 		VariousFunctions.ini, 	Menu memory, Microsoft `Word, 		NO
IniRead, TotalCommanderIni, 	VariousFunctions.ini, 	Menu memory, Total Commander, 		NO
IniRead, PaintIni, 				VariousFunctions.ini, 	Menu memory, Paint, 				NO
IniRead, RebootIni, 			VariousFunctions.ini, 	Menu memory, Reboot, 				NO
IniRead, ShutdownIni, 			VariousFunctions.ini, 	Menu memory, Shutdown, 				NO
IniRead, TranspIni, 			VariousFunctions.ini, 	Menu memory, Transparency, 			NO
IniRead, F13Ini, 				VariousFunctions.ini, 	Menu memory, F13, 					NO
IniRead, F14Ini, 				VariousFunctions.ini, 	Menu memory, F14, 					NO
IniRead, F15Ini, 				VariousFunctions.ini, 	Menu memory, F15, 					NO
IniRead, TopIni,				VariousFunctions.ini, 	Menu memory, Always on top,			NO
IniRead, KeePassIni,			VariousFunctions.ini, 	Menu memory, KeePass,				NO
IniRead, HyperIni,				VariousFunctions.ini, 	Menu memory, Hyperlink,				NO
IniRead, HideIni,				VariousFunctions.ini, 	Menu memory, Hidetext,				NO
IniRead, ShowIni,				VariousFunctions.ini, 	Menu memory, Showtext,				NO
IniRead, AddTemplateIni,		VariousFunctions.ini, 	Menu memory, Add Template,			NO
IniRead, TemplateOffIni,		VariousFunctions.ini, 	Menu memory, Template Off,			NO
IniRead, StrikethroIni, 		VariousFunctions.ini,   Menu memory, Strikethrough Text,    NO
IniRead, DeleteLineIni, 		VariousFunctions.ini,   Menu memory, Delete Line,    		NO
IniRead, AlignLeftIni, 			VariousFunctions.ini,   Menu memory, Align Left,    		NO
IniRead, ApplyStyleIni, 		VariousFunctions.ini,   Menu memory, Apply Styles,    		NO
IniRead, OpenPathIni, 			VariousFunctions.ini,   Menu memory, Open and Show Path,    NO
IniRead, TableIni, 				VariousFunctions.ini,   Menu memory, Table,    				NO
IniRead, SuspendIni,			VariousFunctions.ini,	Menu memory, Suspend,				NO
IniRead, VolumeIni,				VariousFunctions.ini,	Menu memory, Volume Up & Down,		NO
IniRead, BroWinSwiIni,			VariousFunctions.ini,	Menu memory, Browser Win Switcher,	NO
IniRead, TranspMouIni, 			VariousFunctions.ini,	Menu memory, Transparency Mouse,	NO
IniRead, AutosaveIni, 			VariousFunctions.ini,	Menu memory, Autosave,				NO
;/////////////////////////////// - TRAY LABEL - //////////////////////////////////////

;~ Tray:
;--------------------------------------------------	
Menu, Tray, NoStandard
;--------------------------------------------------	
Menu, SubmenuTop, Add, Yes, TopYes
	Menu, SubmenuTop, Add, No, TopNo
	Menu, SubmenuTop, Add, Description, TopDesc
Menu, Tray, Add, &Always on top (Ctrl + Windows + F8), :SubmenuTop
if (TopIni = "NO")
		{
			Menu, SubmenuTop, Check, No 
			Menu, SubmenuTop, UnCheck, Yes 
			F_Top("No")	
		}
		if (TopIni = "YES")
		{
			Menu, SubmenuTop, Check, Yes 
			Menu, SubmenuTop, UnCheck, No
			F_Top("Yes")
		}
;--------------------------------------------------	
Menu, SubmenuCapitalize, Add, Yes, CapitalizeYES 
	Menu, SubmenuCapitalize, Add, No, CapitalizeNO 
	Menu, SubmenuCapitalize, Add, Description, CapitalizeDesc
Menu, Tray, Add, &Capitalization switcher (Shift+F3), :SubmenuCapitalize
		if(CapitalizeIni = "NO")
		{
			Menu, SubmenuCapitalize, Check, No 
			Menu, SubmenuCapitalize, UnCheck, Yes 
			F_Capitalize("No")	
		}
		if (CapitalizeIni = "YES")
		{
			Menu, SubmenuCapitalize, Check, Yes
			Menu, SubmenuCapitalize, UnCheck, No
			F_Capitalize("Yes")
		}
;--------------------------------------------------	
	Menu, Subbrowinswitcher, Add, Yes, Browinswiyes
	Menu, Subbrowinswitcher, Add, No, Browinswino
	Menu, Subbrowinswitcher, Add, Description, BrowinswiDesc
Menu, Tray, Add, C&hrome tab switcher (Xbutton1/Xbutton2), :Subbrowinswitcher
		if (BroWinSwiIni = "NO")
		{
			Menu, Subbrowinswitcher, Check, No 
			Menu, Subbrowinswitcher, UnCheck, Yes 
			F_BroWinSwi("No")	
		}
		if (BroWinSwiIni = "YES")
		{
			Menu, Subbrowinswitcher, Check, Yes 
			Menu, Subbrowinswitcher, UnCheck, No
			F_BroWinSwi("Yes")
		}
;--------------------------------------------------
Menu, SubF13, Add, Yes, F13yes 
	Menu, SubF13, Add, No, F13no
	Menu, SubF13, Add,  Description, F13Desc 
	Menu, SubF14, Add, Yes, F14yes 
	Menu, SubF14, Add, No, F14no
	Menu, SubF14, Add,  Description, F14Desc 
	Menu, SubF15, Add, Yes, F15yes 
	Menu, SubF15, Add, No, F15no
	Menu, SubF15, Add,  Description, F15Desc 
	Menu, SubmenuPedals, Add, F13, :SubF13 
	Menu, SubmenuPedals, Add, F14, :SubF14 
	Menu, SubmenuPedals, Add, F15, :SubF15	
Menu, Tray, Add, &Foot switch (F13/F14/F15), :SubmenuPedals
		if(F13Ini = "NO")
		{
			Menu, SubF13, Check, No 
			Menu, SubF13, UnCheck, Yes 
			F_F13("No")	
		}
		if (F13Ini = "YES")
		{
			Menu, SubF13, Check, Yes 
			Menu, SubF13, UnCheck, No
			F_F13("Yes")
		}
		if(F14Ini = "NO")
		{
			Menu, SubF14, Check, No 
			Menu, SubF14, UnCheck, Yes 
			F_F14("No")	
		}
		if (F14Ini = "YES")
		{
			Menu, SubF14, Check, Yes 
			Menu, SubF14, UnCheck, No
			F_F14("Yes")
		}

		if(F15Ini = "NO")
		{
			Menu, SubF15, Check, No 
			Menu, SubF15, UnCheck, Yes 
			F_F15("No")	
		}
		if (F15Ini = "YES")
		{
			Menu, SubF15, Check, Yes 
			Menu, SubF15, UnCheck, No
			F_F15("Yes")
		}
;--------------------------------------------------	
;TABLE 
	Menu, Subtable, Add, Yes, Tableyes
	Menu, Subtable, Add, No, Tableno 
	Menu, Subtable, Add, Description, TableDesc
		if(TableIni = "NO")
			{
				Menu, Subtable, Check, No 
				Menu, Subtable, UnCheck, Yes 
				F_Table("No")	
			}
		if (TableIni = "YES")
			{
				Menu, Subtable, Check, Yes 
				Menu, Subtable, UnCheck, No
				F_Table("Yes")
			}

;OPEN AND SHOW PARTH
	Menu, Subopenpath, Add, Yes, Openpathyes
	Menu, Subopenpath, Add, No, Openpathno
	Menu, Subopenpath, Add, Description, OpenpathDesc
		if(OpenPathIni = "NO")
			{
				Menu, Subopenpath, Check, No 
				Menu, Subopenpath, UnCheck, Yes 
				F_OpenPath("No")	
			}
		if (OpenPathIni = "YES")
			{
				Menu, Subopenpath, Check, Yes 
				Menu, Subopenpath, UnCheck, No
				F_OpenPath("Yes")
			}

;APPY STYLES
	Menu, Subapply, Add, Yes, Applystylesyes
	Menu, Subapply, Add, No, Applystylesno 
	Menu, Subapply, Add, Description, ApplystylesDesc
		if(ApplyStyleIni = "NO")
			{
				Menu, Subapply, Check, No 
				Menu, Subapply, UnCheck, Yes 
				F_ApplyStyles("No")	
			}
		if (ApplyStyleIni = "YES")
			{
				Menu, Subapply, Check, Yes 
				Menu, Subapply, UnCheck, No
				F_ApplyStyles("Yes")
			}

;ALIGN LEFT 
	Menu, Subalign, Add, Yes, Alignleftyes
	Menu, Subalign, Add, No, Alignleftno 
	Menu, Subalign, Add, Description, AlignleftDesc
		if(AlignLeftIni = "NO")
			{
				Menu, Subalign, Check, No 
				Menu, Subalign, UnCheck, Yes 
				F_AlignLeft("No")	
			}
		if (AlignLeftIni = "YES")
			{
				Menu, Subalign, Check, Yes 
				Menu, Subalign, UnCheck, No
				F_AlignLeft("Yes")
			}

;DELETE LINE
	Menu, Subdeleteline, Add, Yes, Deletelineyes
	Menu, Subdeleteline, Add, No, Deletelineno 
	Menu, Subdeleteline, Add, Description, DeletelineDesc
		if(DeleteLineIni = "NO")
			{
				Menu, Subdeleteline, Check, No 
				Menu, Subdeleteline, UnCheck, Yes 
				F_DeleteLine("No")	
			}
		if (DeleteLineIni = "YES")
			{
				Menu, Subdeleteline, Check, Yes 
				Menu, Subdeleteline, UnCheck, No
				F_DeleteLine("Yes")
			}

;SHOW 
	Menu, Subshow, Add, Yes, Showyes
	Menu, Subshow, Add, No,  Showno
	Menu, Subshow, Add, Description,  ShowDesc
		if(HyperIni = "NO")
			{
				Menu, Subshow, Check, No 
				Menu, Subshow, UnCheck, Yes 
				F_Show("No")	
			}
		if (HyperIni = "YES")
			{
				Menu, Subshow, Check, Yes 
				Menu, Subshow, UnCheck, No
				F_Show("Yes")
			}
;HIDE
	Menu, Subhide, Add, Yes, Hideyes
	Menu, Subhide, Add, No, Hideno 
	Menu, Subhide, Add, Description, HideDesc
		if(HideIni = "NO")
		{
			Menu, Subhide, Check, No 
			Menu, Subhide, UnCheck, Yes 
			F_Hide("No")	
		}
		if (HideIni = "YES")
		{
			Menu, Subhide, Check, Yes 
			Menu, Subhide, UnCheck, No
			F_Hide("Yes")
		}

;HYPERLINKS
	Menu, Subhyper, Add, Yes, Hyperyes 
	Menu, Subhyper, Add, No, Hyperno
	Menu, Subhyper, Add,  Description, HyperDesc 
		if(HyperIni = "NO")
		{
			Menu, Subhyper, Check, No 
			Menu, Subhyper, UnCheck, Yes 
			F_Hyper("No")	
		}
		if (HyperIni = "YES")
		{
			Menu, Subhyper, Check, Yes 
			Menu, Subhyper, UnCheck, No
			F_Hyper("Yes")
		}

	Menu, Subhidden, Add, Show, :Subshow
	Menu, Subhidden, Add, Hide, :Subhide

;ADD TEMPLATE
	Menu, Subaddtemplate, Add, Yes, AddTemplateyes 
	Menu, Subaddtemplate, Add, No, AddTemplateno
	Menu, Subaddtemplate, Add,  Description, AddTemplateDesc 
		if(AddTemplateIni = "NO")
		{
			Menu, Subaddtemplate, Check, No 
			Menu, Subaddtemplate, UnCheck, Yes 
			F_AddTemplate("No")	
		}
		if (AddTemplateIni = "YES")
		{
			Menu, Subaddtemplate, Check, Yes 
			Menu, Subaddtemplate, UnCheck, No
			F_AddTemplate("Yes")
		}

;TEMPLATE OFF
	Menu, Subtemplateoff, Add, Yes, Templateoffyes 
	Menu, Subtemplateoff, Add, No, Templateoffno
	Menu, Subtemplateoff, Add,  Description, TemplateoffDesc 
		if(TemplateOffIni = "NO")
		{
			Menu, Subtemplateoff, Check, No 
			Menu, Subtemplateoff, UnCheck, Yes 
			F_TemplateOff("No")	
		}
		if (TemplateOffIni = "YES")
		{
			Menu, Subtemplateoff, Check, Yes 
			Menu, Subtemplateoff, UnCheck, No
			F_TemplateOff("Yes")
		}

;AUTOSAVE
	Menu, Subautosave, Add, Yes, Autosaveyes
	Menu, Subautosave, Add, No, Autosaveno
	Menu, Subautosave, Add, Description, AutosaveDesc
		if (AutosaveIni = "NO")
		{
			Menu, Subautosave, Check, No 
			Menu, Subautosave, UnCheck, Yes 
			F_Autosave("No")	
		}
		if (AutosaveIni = "YES")
		{
			Menu, Subautosave, Check, Yes 
			Menu, Subautosave, UnCheck, No
			F_Autosave("Yes")
		}


	Menu, Subtemplate, Add, Add Template, :Subaddtemplate
	Menu, Subtemplate, Add, Template Off, :Subtemplateoff


	Menu, Substrikethro, Add, Yes, Strikethroyes
	Menu, Substrikethro, Add, No, Strikethrono 
	Menu, Substrikethro, Add, Description, StrikethroDesc 
	if(StrikethroIni = "NO")
		{
			Menu, Substrikethro, Check, No 
			Menu, Substrikethro, UnCheck, Yes 
			F_Strikethrough("No")	
		}
		if (StrikethroIni = "YES")
		{
			Menu, Substrikethro, Check, Yes 
			Menu, Substrikethro, UnCheck, No
			F_Strikethrough("Yes")
		}

	Menu, SubmenuWord, Add, Align Left, :Subalign
	Menu, SubmenuWord, Add, Apply Styles, :Subapply
	Menu, SubmenuWord, Add, Autosave, :Subautosave
	Menu, SubmenuWord, Add, Delete Line, :Subdeleteline
	Menu, SubmenuWord, Add, Hidden text, :Subhidden 
	Menu, SubmenuWord, Add, Hyperlink, :Subhyper
	Menu, SubmenuWord, Add, Open and Show Path, :Subopenpath
	Menu, SubmenuWord, Add, Strikethrough Text, :Substrikethro
	Menu, SubmenuWord, Add, Table, :Subtable
	Menu, SubmenuWord, Add, Template, :Subtemplate
Menu, Tray, Add, F&unctions in MSWord, :SubmenuWord
;--------------------------------------------------	
Menu, SubmenuBrowser, Add, Yes, BrowserYes
	Menu, SubmenuBrowser, Add, No, BrowserNo
	Menu, SubmenuBrowser, Add, Description, BrowserDesc 	
Menu, Tray, Add, &Open tabs in Chrome, :SubmenuBrowser
		if(BrowserIni = "NO")
		{
			Menu, SubmenuBrowser, Check, No 
			Menu, SubmenuBrowser, UnCheck, Yes 
			F_BrowserMenu("No")
		}
		if (BrowserIni = "YES")
		{
			Menu, SubmenuBrowser, Check, Yes
			Menu, SubmenuBrowser, UnCheck, No
			F_BrowserMenu("Yes")
		}
;--------------------------------------------------	
Menu, SubmenuParenthesis, Add, Yes, ParenthesisYes 
	Menu, SubmenuParenthesis, Add, No, ParenthesisNo
	Menu, SubmenuParenthesis, Add, Description, ParenthesisDesc 
Menu, Tray, Add, &Parenthesis watcher, :SubmenuParenthesis 
		if(ParenthesiIni = "NO")
		{
			Menu, SubmenuParenthesis, Check, No 
			Menu, SubmenuParenthesis, UnCheck, Yes 
			F_ParenthesisMenu("No")	
		}
		if (ParenthesiIni = "YES")
		{
			Menu, SubmenuParenthesis, Check, Yes 
			Menu, SubmenuParenthesis, UnCheck, No
			F_ParenthesisMenu("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuSetEngKeyboard, Add, Yes, SetEngKeyboardYES
	Menu, SubmenuSetEngKeyboard, Add, No, SetEngKeyboardNO
	Menu, SubmenuSetEngKeyboard, Add, Description, SetEngKeyboardDesc				
Menu, Tray, Add, Po&lish/US keyboard switcher, :SubmenuSetEngKeyboard
		if(SetEnglishKeyboardIni = "NO")
		{
			Menu, SubmenuSetEngKeyboard, Check, No 
			Menu, SubmenuSetEngKeyboard, UnCheck, Yes 
			F_SetEngKeyboardMenu("No")	
		}
		if (SetEnglishKeyboardIni = "YES")
		{
			Menu, SubmenuSetEngKeyboard, Check, Yes
			Menu, SubmenuSetEngKeyboard, UnCheck, No
			F_SetEngKeyboardMenu("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuSuspend, Add, Yes, SuspendYes
	Menu, SubmenuSuspend, Add, No, SuspendNo
	Menu, SubmenuSuspend, Add, Description, SuspendDesc
Menu, submenupower, Add, &Suspend (Ctrl + Shift + F1), :SubmenuSuspend
		if(SuspendIni = "NO")
		{
			Menu, SubmenuSuspend, Check, No 
			Menu, SubmenuSuspend, UnCheck, Yes 
			F_Suspend("No")	
		}
		if (SuspendIni = "YES")
		{
			Menu, SubmenuSuspend, Check, Yes 
			Menu, SubmenuSuspend, UnCheck, No
			F_Suspend("Yes")
		}
;--------------------------------------------------	
	Menu, SubmenuReboot, Add, Yes, RebootYES
	Menu, SubmenuReboot, Add, No, RebootNO 
	Menu, SubmenuReboot, Add, Description, RebootDesc
Menu, submenupower, Add, R&eboot (Ctrl+Volume_Up or Ctr+Shift+F2), :SubmenuReboot
		if(RebootIni = "NO")
		{
			Menu, SubmenuReboot, Check, No 
			Menu, SubmenuReboot, UnCheck, Yes 
			F_Reboot("No")	
		}
		if (RebootIni = "YES")
		{
			Menu, SubmenuReboot, Check, Yes
			Menu, SubmenuReboot, UnCheck, No
			F_Reboot("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuShutdown, Add, Yes, ShutdownYES
	Menu, SubmenuShutdown, Add, No, ShutdownNO 
	Menu, SubmenuShutdown, Add, Description, ShutdownDesc 
Menu, submenupower, Add, Sh&utdown and Power down (Ctrl+Volume_Mute or Ctrl+Shift+F3), :SubmenuShutdown
		if(ShutdownIni = "NO")
		{
			Menu, SubmenuShutdown, Check, No 
			Menu, SubmenuShutdown, UnCheck, Yes 
			F_Shutdown("No")	
		}
		if (ShutdownIni = "YES")
		{
			Menu, SubmenuShutdown, Check, Yes
			Menu, SubmenuShutdown, UnCheck, No
			F_Shutdown("Yes")
		}
;--------------------------------------------------
Menu, Tray, Add, Po&wer PC, :submenupower
;--------------------------------------------------
	Menu, SubmenuAltGr, Add, Yes, AltGrYES
	Menu, SubmenuAltGr, Add, No, AltGrNO
	Menu, SubmenuAltGr, Add, Description, AltGrDesc
Menu, Tray, Add, &Right-click context menu (AltGr), :SubmenuALtGr
		if(AltGrIni = "NO")
		{
			Menu, SubmenuAltGr, Check, No 
			Menu, SubmenuAltGr, UnCheck, Yes 
			F_AltGr("No")	
		}
		if (AltGrIni = "YES")
		{
			Menu, SubmenuAltGr, Check, Yes
			Menu, SubmenuAltGr, UnCheck, No
			F_AltGr("Yes")
		}
;--------------------------------------------------
Menu, SubmenuKeepass, Add, Yes, KeepassYes
	Menu, SubmenuKeepass, Add, No, KeepassNo
	Menu, SubmenuKeepass, Add, Description, KeepassDesc
Menu,submenurun, Add, &KeePass (Ctrl + Shift + k), :SubmenuKeepass
		if(KeePassIni = "NO")
		{
			Menu, SubmenuKeepass, Check, No 
			Menu, SubmenuKeepass, UnCheck, Yes 
			F_KeePass("No")	
		}
		if (KeePassIni = "YES")
		{
			Menu, SubmenuKeepass, Check, Yes 
			Menu, SubmenuKeepass, UnCheck, No
			F_KeePass("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuMcsWord, Add, Yes, McsWordYES
	Menu, SubmenuMcsWord, Add, No, McsWordNO 
	Menu, SubmenuMcsWord, Add, Description, McsWordDesc 
Menu, submenurun, Add, &Microsoft `Word (Media_Next), :SubmenuMcsWord 
		if(MicrosoftWordIni = "NO")
		{
			Menu, SubmenuMcsWord, Check, No 
			Menu, SubmenuMcsWord, UnCheck, Yes 
			F_McsWord("No")	
		}
		if (MicrosoftWordIni = "YES")
		{
			Menu, SubmenuMcsWord, Check, Yes
			Menu, SubmenuMcsWord, UnCheck, No
			F_McsWord("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuPaint, Add, Yes, PaintYES
	Menu, SubmenuPaint, Add, No, PaintNO 
	Menu, SubmenuPaint, Add, Description, PaintDesc 
Menu, submenurun, Add, `Pai&nt (Media_Play_Pause), :SubmenuPaint
		if(PaintIni = "NO")
		{
			Menu, SubmenuPaint, Check, No 
			Menu, SubmenuPaint, UnCheck, Yes 
			F_Paint("No")	
		}
		if (PaintIni = "YES")
		{
			Menu, SubmenuPaint, Check, Yes
			Menu, SubmenuPaint, UnCheck, No
			F_Paint("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuTotCom, Add, Yes, TotComYES
	Menu, SubmenuTotCom, Add, No, TotComNO 
	Menu, SubmenuTotCom, Add, Description, TotComDesc 
Menu, submenurun, Add, Total `Comman&der (Media_Prev), :SubmenuTotCom
		if(TotalCommanderIni = "NO")
		{
			Menu, SubmenuTotCom, Check, No 
			Menu, SubmenuTotCom, UnCheck, Yes 
			F_TotalCommander("No")	
		}
		if (TotalCommanderIni = "YES")
		{
			Menu, SubmenuTotCom, Check, Yes
			Menu, SubmenuTotCom, UnCheck, No
			F_TotalCommander("Yes")
		}
;--------------------------------------------------
	Menu, SubmenuPrintScreen, Add, Yes, PrintScreenYES 
	Menu, SubmenuPrintScreen, Add, No, PrintScreenNO 
	Menu, SubmenuPrintScreen, Add, Description, PrintScreenDesc 
Menu, submenurun, Add, Pr&int`Screen (PrintScreen/Volume Down) , :SubmenuPrintScreen
		if(PrintScreenIni = "NO")
		{
			Menu, SubmenuPrintScreen, Check, No 
			Menu, SubmenuPrintScreen, UnCheck, Yes 
			F_PrintScreen("No")	
		}
		if (PrintScreenIni = "YES")
		{
			Menu, SubmenuPrintScreen, Check, Yes
			Menu, SubmenuPrintScreen, UnCheck, No
			F_PrintScreen("Yes")
		}
Menu, Tray, Add, Ru&n..., :submenurun
;--------------------------------------------------		
	Menu, Submenukeys, Add, Yes, TranspYes
	Menu, Submenukeys, Add, No, TranspNo
	Menu, Submenukeys, Add, Description, TranspDesc
		if(TranspIni = "NO")
		{
			Menu, Submenukeys, Check, No 
			Menu, Submenukeys, UnCheck, Yes 
			F_Transparency("No")	
		}
		if (TranspIni = "YES")
		{
			Menu, Submenukeys, Check, Yes 
			Menu, Submenukeys, UnCheck, No
			F_Transparency("Yes")
		}

	Menu, Submenumouse, Add, Yes, Mouseyes
	Menu, Submenumouse, Add, No, Mouseno
	Menu, Submenumouse, Add, Description, MouseDesc 
		if(TranspMouIni = "NO")
		{
			Menu, Submenumouse, Check, No 
			Menu, Submenumouse, UnCheck, Yes 
			F_TransparencyMou("No")	
		}
		if (TranspMouIni = "YES")
		{
			Menu, Submenumouse, Check, Yes 
			Menu, Submenumouse, UnCheck, No
			F_TransparencyMou("Yes")
		}
		
	Menu, SubmenuTransp, Add, Mouse, :Submenumouse
	Menu, SubmenuTransp, Add, Keys, :Submenukeys 
Menu, Tray, Add, &Transparency switcher, :SubmenuTransp
;--------------------------------------------------	
	Menu, Subvolume, Add, Yes, Volumeyes
	Menu, Subvolume, Add, No, Volumeno
	Menu, Subvolume, Add, Description, VolumeDesc
Menu, Tray, Add, &Volume Up And Down (MouseWheel), :Subvolume
		if (VolumeIni = "NO")
		{
			Menu, Subvolume, Check, No 
			Menu, Subvolume, UnCheck, Yes 
			F_Volume("No")	
		}
		if (VolumeIni = "YES")
		{
			Menu, Subvolume, Check, Yes 
			Menu, Subvolume, UnCheck, No
			F_Volume("Yes")
		}
;--------------------------------------------------	
	Menu, SubmenuWindowSwitcher, Add, Yes, WindowSwitcherYES
	Menu, SubmenuWindowSwitcher, Add, No, WindowSwitcherNO 
	Menu, SubmenuWindowSwitcher, Add, Description, WindowSwitcherDesc
Menu, Tray, Add, Win&dow Switcher (LWin and LAlt), :SubmenuWindowSwitcher
		if(WindowSwitcherIni = "NO")
		{
			Menu, SubmenuWindowSwitcher, Check, No 
			Menu, SubmenuWindowSwitcher, UnCheck, Yes 
			F_WindowSwitcher("No")	
		}
		if (WindowSwitcherIni = "YES")
		{
			Menu, SubmenuWindowSwitcher, Check, Yes
			Menu, SubmenuWindowSwitcher, UnCheck, No
			F_WindowSwitcher("Yes")
		}
;--------------------------------------------------
Menu, Tray, Add 
Menu, Tray, Add, About
;--------------------------------------------------
Menu, Tray, Add 
Menu, Tray, Standard
return
;///////////////////////////////// - MENU LABELS - //////////////////////////////////////
About:
MsgBox, Authors: Maciej Slojewski, Hanna Zietak, Jakub Masiak    		Version: 1.1.1
return
;--------------------------------------------------
TurnOffTooltip:
	ToolTip ,
return
;--------------------------------------------------

Autosaveyes:
	Menu,Subautosave, Check, Yes 
	Menu, Subautosave, Uncheck, No 
	F_Autosave("Yes")
return

Autosaveno:
	Menu, Subautosave, Check, No 
	Menu, Subautosave, Uncheck, Yes
	F_Autosave("No")
return

AutosaveDesc:
	Msgbox, Toggle autosave function, by pressing (LCtrl + LAlt + q)
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subautosave, Check, Yes 
		Menu, Subautosave, UnCheck, No 
		F_Autosave("Yes")
		}
		else, 
		{
		Menu, Subautosave, Check, No
		Menu, Subautosave, UnCheck, Yes  
		F_Autosave("No")
		}
return

AutoSave:
{
	init := InitAutosaveFilePath(AutosaveFilePath)
	
	if WinExist("ahk_class OpusApp")
		oWord := ComObjActive("Word.Application")
		
	else
		return
	try
	{
		Loop, % oWord.Documents.Count
		{
			doc := oWord.Documents(A_Index)
			path := doc.Path
			if (path = "")
				return
			fullname := doc.FullName
			
			SplitPath, fullname, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
			doc.Save
			FileGetSize, size_org, % fullname
			size := table[fullname]
			if (size_org != size)
			{
				FormatTime, TimeString, , yyyyMMddHHmmss
				copyname := % AutosaveFilePath . OutNameNoExt . "_" . TimeString . "." . OutExtension
				FileCopy, % fullname, % copyname
				FileGetSize, size, % copyname
				table[fullname] := size
			}
			
		}
	}
	catch
	{
		; try again in 5 seconds
		SetTimer, AutoSave, 5000
		return
	}
	; reset the timer in case it was changed by catch
	SetTimer, AutoSave, % interval
	oWord := ""
	doc := ""
	return
}

InitAutosaveFilePath(path)
{
	if !FileExist(path)
		FileCreateDir, % path
	return true
}
;--------------------------------------------------
Mouseyes:
	Menu,Submenumouse, Check, Yes 
	Menu, Submenumouse, Uncheck, No 
	F_TransparencyMou("Yes")
return

Mouseno:
	Menu, Submenumouse, Check, No 
	Menu, Submenumouse, Uncheck, Yes
	F_TransparencyMou("No")
return

MouseDesc:
	Msgbox, Toggle window tranparency, by moving mouse wheel and pressing {Ctrl}+{Shift}. 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Submenumouse, Check, Yes 
		Menu, Submenumouse, UnCheck, No 
		F_TransparencyMou("Yes")
		}
		else, 
		{
		Menu, Submenumouse, Check, No
		Menu, Submenumouse, UnCheck, Yes  
		F_TransparencyMou("No")
		}
return
;-------------------------------------------------
Browinswiyes:
	Menu,Subbrowinswitcher, Check, Yes 
	Menu, Subbrowinswitcher, Uncheck, No 
	F_BroWinSwi("Yes")
return

Browinswino:
	Menu, Subbrowinswitcher, Check, No 
	Menu, Subbrowinswitcher, Uncheck, Yes
	F_BroWinSwi("No")
return

BroWinSwiDesc:
	Msgbox, Switche tabs in Google Chrome Browser, by pressing {Xbutton1} and {Xbutton2}
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subbrowinswitcher, Check, Yes 
		Menu, Subbrowinswitcher, UnCheck, No 
		F_BroWinSwi("Yes")
		}
		else, 
		{
		Menu, Subbrowinswitcher, Check, No
		Menu, Subbrowinswitcher, UnCheck, Yes  
		F_BroWinSwi("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
Volumeyes:
	Menu,Subvolume, Check, Yes 
	Menu, Subvolume, Uncheck, No 
	F_Volume("Yes")
return

Volumeno:
	Menu, Subvolume, Check, No 
	Menu, Subvolume, Uncheck, Yes
	F_Volume("No")
return

VolumeDesc:
	Msgbox, Turn the volume up and down, by moving a mouse wheel. Works only when a caret is over the system tray. 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subvolume, Check, Yes 
		Menu, Subvolume, UnCheck, No 
		F_Volume("Yes")
		}
		else, 
		{
		Menu, Subvolume, Check, No
		Menu, Subvolume, UnCheck, Yes  
		F_Volume("No")
		}
return

;--------------------------------------------------
;-------------------------------------------------
Suspendyes:
	Menu,SubmenuSuspend, Check, Yes 
	Menu, SubmenuSuspend, Uncheck, No 
	F_Suspend("Yes")
return

Suspendno:
	Menu, SubmenuSuspend, Check, No 
	Menu, SubmenuSuspend, Uncheck, Yes
	F_Suspend("No")
return

SuspendDesc:
	Msgbox, Suspend, by pressing {Ctrl} + {shift} + F1
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuSuspend, Check, Yes 
		Menu, SubmenuSuspend, UnCheck, No 
		F_Suspend("Yes")
		}
		else, 
		{
		Menu, SubmenuSuspend, Check, No
		Menu, SubmenuSuspend, UnCheck, Yes  
		F_Suspend("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
Tableyes:
	Menu,Subtable, Check, Yes 
	Menu, Subtable, Uncheck, No 
	F_Table("Yes")
return

Tableno:
	Menu, Subtable, Check, No 
	Menu, Subtable, Uncheck, Yes
	F_Table("No")
return


TableDesc:
	Msgbox, Hotstring: After typing "tabela`t", you receive | | |{Enter} - table in Microsoft Word 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subtable, Check, Yes 
		Menu, Subtable, UnCheck, No 
		F_Table("Yes")
		}
		else, 
		{
		Menu, Subtable, Check, No
		Menu, Subtable, UnCheck, Yes  
		F_Table("No")
		}
	return
;--------------------------------------------------
;-------------------------------------------------
Openpathyes:
	Menu, Subopenpath, Check, Yes 
	Menu, Subopenpath, Uncheck, No 
	F_OpenPath("Yes")
return

Openpathno:
	Menu, Subopenpath, Check, No 
	Menu, Subopenpath, Uncheck, Yes
	F_OpenPath("No")
return


OpenpathDesc:
	Msgbox, Open "Open" window and show the path of a document in Microsoft Word, by pressing {Ctrl} + o
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subopenpath, Check, Yes 
		Menu, Subopenpath, UnCheck, No 
		F_OpenPath("Yes")
		}
		else, 
		{
		Menu, Subopenpath, Check, No
		Menu, Subopenpath, UnCheck, Yes  
		F_OpenPath("No")
		}
	return

;--------------------------------------------------
;-------------------------------------------------
Applystylesyes:
	Menu,Subapply, Check, Yes 
	Menu, Subapply, Uncheck, No 
	F_ApplyStyles("Yes")
return

Applystylesno:
	Menu, Subapply, Check, No 
	Menu, Subapply, Uncheck, Yes
	F_ApplyStyles("No")
return


ApplystylesDesc:
	Msgbox, Open and close the Apply Styles window in Microsoft Word, by pressing {Ctrl} + {shift} + s
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subapply, Check, Yes 
		Menu, Subapply, UnCheck, No 
		F_ApplyStyles("Yes")
		}
		else, 
		{
		Menu, Subapply, Check, No
		Menu, Subapply, UnCheck, Yes  
		F_ApplyStyles("No")
		}
		return
;--------------------------------------------------
;-------------------------------------------------
Alignleftyes:
	Menu,Subalign, Check, Yes 
	Menu, Subalign, Uncheck, No 
	F_AlignLeft("Yes")
return

Alignleftno:
	Menu, Subalign, Check, No 
	Menu, Subalign, Uncheck, Yes
	F_AlignLeft("No")
return

AlignleftDesc:
	Msgbox, Align your content with the left margin in Microsoft Word, by pressing {Ctrl} + {Shift} + l
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subalign, Check, Yes 
		Menu, Subalign, UnCheck, No 
		F_AlignLeft("Yes")
		}
		else, 
		{
		Menu, Subalign, Check, No
		Menu, Subalign, UnCheck, Yes  
		F_AlignLeft("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
Deletelineyes:
	Menu,Subdeleteline, Check, Yes 
	Menu, Subdeleteline, Uncheck, No 
	F_DeleteLine("Yes")
return

Deletelineno:
	Menu, Subdeleteline, Check, No 
	Menu, Subdeleteline, Uncheck, Yes
	F_DeleteLine("No")
return

DeletelineDesc:
	Msgbox, Delete whole text line in Microsoft Word, by pressing {Ctrl} + l
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subdeleteline, Check, Yes 
		Menu, Subdeleteline, UnCheck, No 
		F_DeleteLine("Yes")
		}
		else, 
		{
		Menu, Subdeleteline, Check, No
		Menu, Subdeleteline, UnCheck, Yes  
		F_DeleteLine("No")
		}
		return
;--------------------------------------------------
;-------------------------------------------------
Strikethroyes:
	Menu,Substrikethro, Check, Yes 
	Menu, Substrikethro, Uncheck, No 
	F_Strikethrough("Yes")
return

Strikethrono:
	Menu, Substrikethro, Check, No 
	Menu, Substrikethro, Uncheck, Yes
	F_Strikethrough("No")
return

StrikethroDesc:
	Msgbox, Stike selected text through, by pressing {Ctrl} + {Shift} + x
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Substrikethro, Check, Yes 
		Menu, Substrikethro, UnCheck, No 
		F_Strikethrough("Yes")
		}
		else, 
		{
		Menu, Substrikethro, Check, No
		Menu, Substrikethro, UnCheck, Yes  
		F_Strikethrough("No")
		}
return

;--------------------------------------------------
;-------------------------------------------------
Hyperyes:
	Menu,Subhyper, Check, Yes 
	Menu, Subhyper, Uncheck, No 
	F_Hyper("Yes")
return

Hyperno:
	Menu, Subhyper, Check, No 
	Menu, Subhyper, Uncheck, Yes
	F_Hyper("No")
return

HyperDesc:
	Msgbox, Add hyperlink in selected text in Microsoft Word, by pressing {Ctrl} + k
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subhyper, Check, Yes 
		Menu, Subhyper, UnCheck, No 
		F_Hyper("Yes")
		}
		else, 
		{
		Menu, Subhyper, Check, No
		Menu, Subhyper, UnCheck, Yes  
		F_Hyper("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------

Hideyes:
	Menu, Subhide, Check, Yes 
	Menu, Subhide, Uncheck, No 
	F_Hide("Yes")
return

Hideno:
	Menu, Subhide, Check, No 
	Menu, Subhide, Uncheck, Yes
	F_Hide("No")
return

HideDesc:
	Msgbox, Hide selected text in Microsoft Word, by pressing {Shift} + {Ctrl} + h
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subhide, Check, Yes 
		Menu, Subhide, UnCheck, No 
		F_Hide("Yes")
		}
		else, 
		{
		Menu, Subhide, Check, No
		Menu, Subhide, UnCheck, Yes  
		F_Hide("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
Showyes:
	Menu, Subshow, Check, Yes 
	Menu, Subshow, Uncheck, No 
	F_Show ("Yes")
return

Showno:
	Menu, Subshow, Check, No 
	Menu, Subshow, Uncheck, Yes
	F_Show("No")
return

ShowDesc:
	Msgbox, Show hidden text and special signs in Microsoft Word, by pressing {Ctrl} + *
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subshow, Check, Yes 
		Menu, Subshow, UnCheck, No 
		F_Show("Yes")
		}
		else, 
		{
		Menu, Subshow, Check, No
		Menu, Subshow, UnCheck, Yes  
		F_Show("No")
		}
	return
;--------------------------------------------------
;-------------------------------------------------
Templateoffyes:
	Menu, Subtemplateoff, Check, Yes 
	Menu, Subtemplateoff, Uncheck, No 
	F_TemplateOff("Yes")
return

Templateoffno:
	Menu, Subtemplateoff, Check, No
	Menu, Subtemplateoff, Uncheck, Yes
	F_TemplateOff("No")
return

TemplateoffDesc:
	Msgbox, Switch off added template, by pressing {Ctrl} + {Shift} + t
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subtemplateoff, Check, Yes 
		Menu, Subtemplateoff, UnCheck, No 
		F_TemplateOff("Yes")
		}
		else, 
		{
		Menu, Subtemplateoff, Check, No
		Menu, Subtemplateoff, UnCheck, Yes  
		F_TemplateOff("No")
		}
return

;--------------------------------------------------
;-------------------------------------------------
Addtemplateyes:
	Menu, Subaddtemplate, Check, Yes 
	Menu, Subaddtemplate, Uncheck, No 
	F_AddTemplate("Yes")
return
Addtemplateno:
	Menu, Subaddtemplate, Check, No
	Menu, Subaddtemplate, Uncheck, Yes
	F_AddTemplate("No")
return

AddtemplateDesc:
	Msgbox, Add Polish or English template in Microsoft Word, by pressing {Ctrl} + t
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Subaddtemplate, Check, Yes 
		Menu, Subaddtemplate, UnCheck, No 
		F_AddTemplate("Yes")
		}
		else, 
		{
		Menu, Subaddtemplate, Check, No
		Menu, Subaddtemplate, UnCheck, Yes  
		F_AddTemplate("No")
		}
return

AutoTemplate:
	oWord := ComObjActive("Word.Application")
	try
		template := oWord.ActiveDocument.CustomDocumentProperties["PopSzab"].Value
	catch
	{
		oWord.ActiveDocument.CustomDocumentProperties.Add("PopSzab",0,4," ")
		template := oWord.ActiveDocument.CustomDocumentProperties["PopSzab"].Value
	}
	if ((template == "PL") or (template == "EN"))
	{
		gosub, AddTemplate
	}
	else
		gosub, ChooseTemplate
	return

AddTemplate:
	if !(FileExist("S:\"))
	{
		MsgBox,16,, Unable to add template. To continue, connect to voestalpine servers and try again.
		oWord := ""
		return
	}
	OurTemplate := oWord.ActiveDocument.AttachedTemplate.FullName
	if (template == "PL")
	{
		if (OurTemplate == OurTemplatePL)
		{
			oWord := ""
			
		}
		else
		{
			oWord.ActiveDocument.AttachedTemplate := OurTemplatePL
			oWord.ActiveDocument.UpdateStylesOnOpen := WordTrue
			oWord.ActiveDocument.UpdateStyles
			;MsgBox, 64, Informacja, % MsgText("Dołączono szablon!`nDołączono domyślny szablon dokumentu: `n") oWord.ActiveDocument.AttachedTemplate.FullName, 5
			MsgBox, 64,, Dołączono szablon! 
			OurTemplate := OurTemplatePL
		}
	}
	else if (template == "EN")
	{
		if (OurTemplate == OurTemplateEN)
		{
			oWord := ""
			
		}
		else
		{
			oWord.ActiveDocument.AttachedTemplate := OurTemplateEN
			oWord.ActiveDocument.UpdateStylesOnOpen :=  WordTrue
			oWord.ActiveDocument.UpdateStyles
			;MsgBox, 64, Informacja, % MsgText("Dołączono szablon!`nDołączono domyłlny szablon dokumentu: `n") oWord.ActiveDocument.AttachedTemplate.FullName, 5
			MsgBox, 64,, The template is added!
			OurTemplate := OurTemplateEN
		}
	}
	oWord.ActiveDocument.CustomDocumentProperties["PopSzab"] := template
	MsgBox, 36,, Do you want to set size of the margins?
	IfMsgBox, Yes
	{
		oWord := ComObjActive("Word.Application")
		oWord.Run("!Wydruk")
	}
	MsgBox, 36,, Do you want to add some building blocks to your document?
	IfMsgBox, Yes
		gosub, AddBB
	oWord := ""
	return

AddBB:
	Gui, BB:New
	Gui, BB:Add, Text,, Choose building blocks you want to add:
	Gui, BB:Add, Checkbox, vFirstPage, First Page
	Gui, BB:Add, Checkbox, vID, ID
	Gui, BB:Add, Checkbox, vChangeLog, Change Log
	Gui, BB:Add, Checkbox, vTOC, Table of Contents
	Gui, BB:Add, Checkbox, vLOT, List of Tables
	Gui, BB:Add, Checkbox, vLOF, List of Figures
	Gui, BB:Add, Checkbox, vIntro, Introduction
	Gui, BB:Add, Checkbox, vLastPage, Last Page
	Gui, BB:Add, Button, w200 gBBOK Default, OK
	oWord.Run("AddDocProperties")
	Gui, BB:Show,, Add Building Blocks
return

ChooseTemplate:
	MsgBox, 36,, Do you want to add a template to this document?
	IfMsgBox, Yes
	{
		Gui, Temp:New
		Gui, Temp:Add, Text,, Choose template:
		Gui, Temp:Add, Radio, vMyTemplate Checked, Polish template
		Gui, Temp:Add, Radio,, English template
		Gui, Temp:Add, Button, w200 gTempOK Default, OK
		Gui, Temp:Show,, Add Template
	}
	return
	
TempOK:
	Gui, Temp:Submit, +OwnDialogs
	if (MyTemplate == 1)
	{
		template := "PL"
	}
	else if (MyTemplate == 2)
	{
		template := "EN"
	}
	gosub, AddTemplate
	return


BBOK:
	Gui, BB:Submit, +OwnDialogs
	Gui, BB:Destroy
	if (FirstPage == 1)
		BB_Insert("Strona ozdobna", "")
	if (ID == 1)
		BB_Insert("identyfikator", "")
	if (ChangeLog == 1)
		BB_Insert("Lista zmian", "")
	if (TOC == 1)
	{
		BB_Insert("Spis treści", "")
		Send, {Right}{Enter}{Enter}
	}
	if (LOT == 1)
	{
		BB_Insert("Spis tabel", "")
		Send, {Right}{Enter}{Enter}
	}
	if (LOF == 1)
	{
		BB_Insert("Spis rysunków", "")
		Send, {Right}{Enter}{Enter}
	}
	if (Intro == 1)
	{
		oWord := ComObjActive("Word.Application")
		oWord.ActiveDocument.Bookmarks.Add("intro", oWord.Selection.Range)
		Send, {Enter}{Enter}
	}
	if (LastPage == 1)
	{
		oWord := ComObjActive("Word.Application")
		oWord.Selection.InsertBreak(wdSectionBreakNextPage := 2)
		BB_Insert("OstatniaStronaObrazek", "")
	if (Intro == 1)
	{
		oWord := ComObjActive("Word.Application")
		oWord.Selection.GoTo(-1,,,"intro")
		oWord.Selection.Find.ClearFormatting
		oWord.ActiveDocument.Bookmarks("intro").Delete
	}
	}


	BB_Insert(Name_BB, AdditionalText)
	{
	global 
	Name_BB := MsgText(Name_BB)
	Base(AdditionalText)
	oWord := ComObjActive("Word.Application")
	;~ MsgBox, % oWord.ActiveDocument.AttachedTemplate.FullName
	if  !( InStr(OurTemplate, "TQ-S402-pl_OgolnyTechDok.dotm") or InStr(OurTemplate, "TQ-S402-en_OgolnyTechDok.dotm") )
		{
		MsgBox, 16, % MsgText("Próba wstawienia bloku z szablonu"), % MsgText("Próbujesz wstawić blok konstrukcyjny przypisany do szablonu, ale szablon nie zostać jeszcze dołączony do tego pliku.`nNajpierw dołącz szablon, a nastepnie wywołaj ponownie tę funkcję.")
		}
	else
		{
		OurTemplate := oWord.ActiveDocument.AttachedTemplate.FullName
		oWord.Templates(OurTemplate).BuildingBlockEntries(Name_BB).Insert(oWord.Selection.Range, WordTrue)
		}
	oWord :=  "" ; Clear global COM objects when done with them
	}
return
;--------------------------------------------------
;-------------------------------------------------

KeepassYes:
	Menu, SubmenuKeepass, Check, Yes 
	Menu, SubmenuKeepass, Uncheck, No 
	F_KeePass("Yes")
return

KeepassNo:
	Menu, SubmenuKeepass, Check, Yes 
	Menu, SubmenuKeepass, Uncheck, No 
	F_KeePass("No")
return

KeepassDesc:
	Msgbox, Run the KeePass 2 application, by pressing {Shift} + {Ctrl} + k 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuKeepass, Check, Yes 
		Menu, SubmenuKeepass, UnCheck, No 
		F_KeePass("Yes")
		}
		else, 
		{
		Menu, SubmenuKeepass, Check, No
		Menu, SubmenuKeepass, UnCheck, Yes  
		F_KeePass("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
TopYes:
	Menu, SubmenuTop, Check, Yes 
	Menu, SubmenuTop, Uncheck, No 
	F_Top("Yes")
return

TopNo:
	Menu, SubmenuTop, Check, Yes 
	Menu, SubmenuTop, Uncheck, No 
	F_Top("No")
return

TopDesc:
	Msgbox, Toggle window parameter always on top, by pressing {Ctrl} + {Windows} + {F8}
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuTop, Check, Yes 
		Menu, SubmenuTop, UnCheck, No 
		F_Top("Yes")
		}
		else, 
		{
		Menu, SubmenuTop, Check, No
		Menu, SubmenuTop, UnCheck, Yes  
		F_Top("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
F13yes:
	Menu, SubF13, check, Yes 
	Menu, SubF13, uncheck, No
	F_F13("Yes")
return

F13no:
	Menu, SubF13, check, No 
	Menu, SubF13, uncheck, Yes
	F_F13("No")
return

F13Desc:
	MsgBox, Switch windows in the system tray, by pressing {F13} 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubF13, Check, Yes 
		Menu, SubF13, UnCheck, No 
		F_F13("Yes")
		}
		else, 
		{
		Menu, SubF13, Check, No
		Menu, SubF13, UnCheck, Yes  
		F_F13("No")
		}
return


F14yes:
	Menu, SubF14, check, Yes 
	Menu, SubF14, uncheck, No
	F_F14("Yes")
return

F14no:
	Menu, SubF14, check, No 
	Menu, SubF14, uncheck, Yes
	F_F14("No")
return

F14Desc:
	Msgbox, Immediately resets the hotstring recognizer, by pressing {F13}
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubF14, Check, Yes 
		Menu, SubF14, UnCheck, No 
		F_F14("Yes")
		}
		else, 
		{
		Menu, SubF14, Check, No
		Menu, SubF14, UnCheck, Yes  
		F_F14("No")
		}
return


F15yes:
	Menu, SubF15, check, Yes 
	Menu, SubF15, uncheck, No
	F_F15("Yes")
Return

F15no:
	Menu, SubF15, check, No 
	Menu, SubF15, uncheck, Yes
	F_F15("No")
return

F15Desc:
	Msgbox, Make a beep sound, by pressing {F15}
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubF15, Check, Yes 
		Menu, SubF15, UnCheck, No 
		F_F15("Yes")
		}
		else, 
		{
		Menu, SubF15, Check, No
		Menu, SubF15, UnCheck, Yes  
		F_F15("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
TranspYes:
	Menu, Submenukeys, Check, Yes
	Menu, Submenukeys, UnCheck, No 
	F_Transparency("Yes")
return

TranspNo:
	Menu, Submenukeys, Check, No
	Menu, Submenukeys, UnCheck, Yes 
	F_Transparency("No")
Return

TranspDesc:
	MsgBox, Toggle window transparency by prssing {Ctr} + {Windows} + {F9}
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, Submenukeys, Check, Yes 
		Menu, Submenukeys, UnCheck, No 
		F_Transparency("Yes")
		}
		else, 
		{
		Menu, Submenukeys, Check, No
		Menu, Submenukeys, UnCheck, Yes  
		F_Transparency("No")
		}
return
;--------------------------------------------------
;-------------------------------------------------
ShutdownYES:
	Menu,SubmenuShutdown, Check, Yes 
	Menu, SubmenuShutdown, UnCheck, No
	F_Shutdown("Yes")
return

ShutdownNO:
	Menu,SubmenuShutdown, Check, No
	Menu, SubmenuShutdown, UnCheck, Yes
	F_Shutdown("No")
return 

ShutdownDesc:
	MsgBox, `Shutdown system, by pressing a Multimedia key - Volume Mute or Ctrl + Shift + F3 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuShutdown, Check, Yes 
		Menu, SubmenuShutdown, UnCheck, No 
		F_Shutdown("Yes")
		}
		else, 
		{
		Menu, SubmenuShutdown, Check, No
		Menu, SubmenuShutdown, UnCheck, Yes  
		F_Shutdown("No")
		}
return

;--------------------------------------------------
;--------------------------------------------------
RebootYES:
	Menu,SubmenuReboot, Check, Yes 
	Menu, SubmenuReboot, UnCheck, No
	F_Reboot("Yes")
return

RebootNO:
	Menu,SubmenuReboot, Check, No
	Menu, SubmenuReboot, UnCheck, Yes
	F_Reboot("No")
return 

RebootDesc:
	MsgBox, Reboot by pressing a Multimedia key - Volume Up or Ctrl + Shift + F2 
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuReboot, Check, Yes 
		Menu, SubmenuReboot, UnCheck, No 
		F_Shutdown("Yes")
		}
		else, 
		{
		Menu, SubmenuReboot, Check, No
		Menu, SubmenuReboot, UnCheck, Yes  
		F_Shutdown("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
PaintYES:
	Menu,SubmenuPaint, Check, Yes 
	Menu, SubmenuPaint, UnCheck, No
	F_Paint("Yes")
return

PaintNO:
	Menu,SubmenuPaint, Check, No
	Menu, SubmenuPaint, UnCheck, Yes
	F_Paint("No")
return 

PaintDesc:
	MsgBox, Run Paint application by pressing a Multimedia key - Media_Play_Pause
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuPaint, Check, Yes 
		Menu, SubmenuPaint, UnCheck, No 
		F_Paint("Yes")
		}
		else, 
		{
		Menu, SubmenuPaint, Check, No
		Menu, SubmenuPaint, UnCheck, Yes  
		F_Paint("No")
		}
return

;--------------------------------------------------
;--------------------------------------------------
TotComYES:
	Menu,SubmenuTotCom, Check, Yes 
	Menu, SubmenuTotCom, UnCheck, No
	F_TotalCommander("Yes")
return

TotComNO:
	Menu,SubmenuTotCom, Check, No
	Menu, SubmenuTotCom, UnCheck, Yes
	F_TotalCommander("No")
return 

TotComDesc:
	MsgBox, Run Total Commander application by pressing a Multimedia key - Media_Prev
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuTotCom, Check, Yes 
		Menu, SubmenuTotCom, UnCheck, No 
		F_TotalCommander("Yes")
		}
		else, 
		{
		Menu, SubmenuTotCom, Check, No
		Menu, SubmenuTotCom, UnCheck, Yes  
		F_TotalCommander("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
McsWordYES:
	Menu,SubmenuMcsWord, Check, Yes 
	Menu, SubmenuMcsWord, UnCheck, No
	F_McsWord("Yes")
return

McsWordNO:
	Menu,SubmenuMcsWord, Check, No
	Menu, SubmenuMcsWord, UnCheck, Yes
	F_McsWord("No")
return 

McsWordDesc:
	MsgBox, Run Microsoft Word application by pressing a Multimedia key - Media_Next
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuMcsWord, Check, Yes 
		Menu, SubmenuMcsWord, UnCheck, No 
		F_McsWord("Yes")
		}
		else, 
		{
		Menu, SubmenuMcsWord, Check, No
		Menu, SubmenuMcsWord, UnCheck, Yes  
		F_McsWord("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
CapitalizeYES:
	Menu, SubmenuCapitalize, Check, Yes 
	Menu, SubmenuCapitalize, UnCheck, No
	F_Capitalize("Yes")
return

CapitalizeNO:
	Menu, SubmenuCapitalize, UnCheck, Yes
	Menu, SubmenuCapitalize, Check, No 
	F_Capitalize("No")
return

CapitalizeDesc:
	MsgBox,
	( 
Author: Jakub Masiak`n
Refactored by: Hanna Ziętak on 2022-01-05`n
	Input: one or more words (eg. Dog/Dog is jumping)`n
Shift+{F3} and selected sentence/word or a caret in the middle of a word`n
	Output: Alters the first or all letters`n
EXAMPLES:`n
	Dog is jumping -> DOG IS JUMPING`n
	Dog -> DOG`n
	DOG IS JUMPING -> dog is jumping`n
	DOG -> dog`n
	dog is jumping -> Dog is jumping`n
	dog -> Dog`n
Shifts caret to the end.`n
It works everywhere exept Word, because in Word Application this function already exists.`n
	)
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuCapitalize, Check, Yes 
		Menu, SubmenuCapitalize, UnCheck, No 
		F_Capitalize("Yes")
		}
		else, 
		{
		Menu, SubmenuCapitalize, Check, No
		Menu, SubmenuCapitalize, UnCheck, Yes  
		F_Capitalize("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
PrintScreenYES:
	Menu, SubmenuPrintScreen, Check, Yes 
	Menu, SubmenuPrintScreen, UnCheck, No 
	F_PrintScreen("Yes")
return 

PrintScreenNO:
	Menu, SubmenuPrintScreen, Check, No 
	Menu, SubmenuPrintScreen, UnCheck, Yes 
	F_PrintScreen("No")
return 

PrintScreenDesc:
	MsgBox, 
	(
Run applications: SnippingTool.exe (by pressing {Volume down} key → Multimedia keys)`n						
Windows Printscreen application (by pressing {PrintScreen} key (https://support.microsoft.com/pl-pl/help/4488540/how-to-take-and-annotate-screenshots-on-windows-10)    
	)    
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuPrintScreen, Check, Yes 
		Menu, SubmenuPrintScreen, UnCheck, No 
		F_PrintScreen("Yes")
		}
		else, 
		{
		Menu, SubmenuPrintScreen, Check, No
		Menu, SubmenuPrintScreen, UnCheck, Yes  
		F_PrintScreen("No")
		}
return   
;--------------------------------------------------
;--------------------------------------------------
WindowSwitcherYes:
	Menu, SubmenuWindowSwitcher, Check, Yes 
	Menu, SubmenuWindowSwitcher, UnCheck, No 
	F_WindowSwitcher("Yes")
return 

WindowSwitcherNO:
	Menu, SubmenuWindowSwitcher, Check, No
	Menu, SubmenuWindowSwitcher, UnCheck, Yes  
	F_WindowSwitcher("No")
return

WindowSwitcherDesc: 
	MsgBox, Switches between windows by pressing {Left Windows} key and {Left Alt} key, then you can move between windows by using left, right, up and down arrows  
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuWindowSwitcher, Check, Yes 
		Menu, SubmenuWindowSwitcher, UnCheck, No 
		F_WindowSwitcher("Yes")
		}
		else, 
		{
		Menu, SubmenuWindowSwitcher, Check, No
		Menu, SubmenuWindowSwitcher, UnCheck, Yes  
		F_WindowSwitcher("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
AltGrYES:
	Menu, SubmenuAltGr, Check, Yes  
	Menu, SubmenuAltGr, Uncheck, No 
	F_AltGr("Yes")
return

AltGrNO:
	Menu, SubmenuAltGr, Uncheck, Yes
	Menu, SubmenuAltGr, Check, No  
	F_AltGr("No")
return

AltGrDesc: 
	Msgbox, Redirects AltGr -> context menu    (only in English keyboardLayout)
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuAltGr, Check, Yes 
		Menu, SubmenuAltGr, UnCheck, No 
		F_AltGr("Yes")
		}
		else, 
		{
		Menu, SubmenuAltGr, Check, No
		Menu, SubmenuAltGr, UnCheck, Yes  
		F_AltGr("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
SetEngKeyboardYES:
	Menu, SubmenuSetEngKeyboard, Check, Yes 
	Menu, SubmenuSetEngKeyboard, UnCheck, No
	F_SetEngKeyboardMenu("Yes")
return

SetEngKeyboardNO:
	Menu, SubmenuSetEngKeyboard, Check, No
	Menu, SubmenuSetEngKeyboard, UnCheck, Yes
	F_SetEngKeyboardMenu("No")
return

SetEngKeyboardDesc:
	MsgBox, Change keyboard settings (from Polish keyboard to English keyboard)
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuSetEngKeyboard, Check, Yes 
		Menu, SubmenuSetEngKeyboard, UnCheck, No 
		F_SetEngKeyboardMenu("Yes")
		}
		else, 
		{
		Menu, SubmenuSetEngKeyboard, Check, No
		Menu, SubmenuSetEngKeyboard, UnCheck, Yes  
		F_SetEngKeyboardMenu("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
BrowserYes:
	Menu, SubmenuBrowser, Check, Yes 
	Menu, SubmenuBrowser, UnCheck, No 
	F_BrowserMenu("Yes")
return 

BrowserNo:
	Menu, SubmenuBrowser, UnCheck, Yes 
	Menu, SubmenuBrowser, Check, No
	F_BrowserMenu("No")
return 

BrowserDesc:
	MsgBox,
	(
	 Runs links:
chrome.exe https://translate.google.com/`n
https://www.linkedin.com/feed/`n
https://mail.google.com/mail/u/0/#inbox`n
http://www.meteo.pl/`n 
https://trello.com/b/5h4R58KL/organizacyjne`n
https://team.voestalpine.net/SitePages/Home.aspx`n
https://helpdesk.tens.pl/helpdesk/`n
https://portal-signaling-poland.voestalpine.net/synergy/docs/Portal.aspx`n
https://solidsystemteamwork.voestalpine.root.local/internalprojects/vaSupp/CPS/SitePages/Home.aspx`n
	https://solidsystemteamwork.voestalpine.root.local/Processes/custprojects/780MDSUpgradeKit/SitePages/Home.aspx`n
	)
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuBrowser, Check, Yes 
		Menu, SubmenuBrowser, UnCheck, No 
		F_BrowserMenu("Yes")
		}
		else, 
		{
		Menu, SubmenuBrowser, Check, No
		Menu, SubmenuBrowser, UnCheck, Yes  
		F_BrowserMenu("No")
		}
return
;--------------------------------------------------
;--------------------------------------------------
ParenthesisYes:
	Menu, SubmenuParenthesis, Check, Yes 
	Menu, SubmenuParenthesis, UnCheck, No 
	F_ParenthesisMenu("Yes")
return

ParenthesisNo:
	Menu, SubmenuParenthesis, UnCheck, Yes
	Menu, SubmenuParenthesis, Check, No
	F_ParenthesisMenu("No")
return 

ParenthesisDesc:
	MsgBox, After pressing keys like: {  [  (  `" , the closing symbols }  ]  ) `" will also appear. Aditionally a caret will jump between the parenthesis/quotation marks. It works also, when you have already written a text and want to put it between parenthesis/quotation marks. You have to select words and press parenthesis/quotation marks.
	MsgBox, 4, ,Would you like to activate it?
		IfMsgBox, Yes
		{
		Menu, SubmenuParenthesis, Check, Yes 
		Menu, SubmenuParenthesis, UnCheck, No 
		F_ParenthesisMenu("Yes")
		}
		else, 
		{
		Menu, SubmenuParenthesis, Check, No
		Menu, SubmenuParenthesis, UnCheck, Yes  
		F_ParenthesisMenu("No")
		}
return
;--------------------------------------------------

; End of Auto-execute Section: https://www.autohotkey.com/docs/Scripts.htm#auto

#z::Menu, Tray, Show
return
;/////////////////////////////// - SECTION OF FUNCTIONS: BEGINNING - ////////////////////////////////////////
Base(AdditionalText := "")
{
	AdditionalText := MsgText(AdditionalText)
	tooltip, [F24]  %A_thishotKey% %AdditionalText%
	SetTimer, TurnOffTooltip, -5000
	return
}
;*************************************************************************
MsgText(string)
{
    vSize := StrPut(string, "CP0")
    VarSetCapacity(vUtf8, vSize)
    vSize := StrPut(string, &vUtf8, vSize, "CP0")
    Return StrGet(&vUtf8, "UTF-8") 
}
;*************************************************************************
F_ParenthesisMenu(MyArg)
{
    if (MyArg = "Yes")
	{
		Hotkey, ~{ , F_Parenthesis, On
		Hotkey, ~" , F_Parenthesis, On
		Hotkey, ~( , F_Parenthesis, On
		Hotkey, ~[ , F_Parenthesis, On
		;~ Hotkey, ~LButton Up, F_StoreClipboard(0.1), On	; for MS Word: 0.01 is too short, 0.1 worksS
		Hotkey,     ~+Right Up,     F_Parenthesis, On	;;events related to keyboard; order matters!
		Hotkey,     ~+Left Up,      F_Parenthesis, On
		Hotkey,     ~^+Left Up,     F_Parenthesis, On
		Hotkey,     ~^+Right Up,    F_Parenthesis, On
        Hotkey,     ~LButton Up,    F_Parenthesis, On
		IniWrite,   YES, VariousFunctions.ini, Menu memory, Parenthesis
	}

	if (MyArg = "No")
	{
		Hotkey, ~{ , F_Parenthesis, Off
		Hotkey, ~" , F_Parenthesis, Off
		Hotkey, ~( , F_Parenthesis, Off
		Hotkey, ~[ , F_Parenthesis, Off
		;~ Hotkey, ~LButton Up, F_Parenthesis(0.1), On	; for MS Word: 0.01 is too short, 0.1 works
		Hotkey,      ~+Right Up,    F_Parenthesis, Off	;;events related to keyboard; order matters!
		Hotkey,      ~+Left Up,     F_Parenthesis, Off
		Hotkey,      ~^+Left Up,    F_Parenthesis, Off
		Hotkey,      ~^+Right Up,   F_Parenthesis, Off 
        Hotkey,      ~LButton Up,   F_Parenthesis, Off 
		IniWrite, NO, VariousFunctions.ini, Menu memory, Parenthesis
    }

}
F_StoreClipboard(IsMouse*)
{
	global	;assume-global mode
	local OldClipboard := ""
	
	OldClipBoard := ClipboardAll
	Clipboard := ""
	Send, ^c
	if (IsMouse[1] != 0)
		ClipWait, % IsMouse[1]
	else
		ClipWait
	ToRemember := Clipboard
	OutputDebug, % "ToRemember:" . A_Tab . ToRemember
	Clipboard := OldClipBoard
	OldClipBoard := ""
}
F_Parenthesis()
{	
    global 
	local ThisHotkey := A_ThisHotkey, SelectedWords:= true, f_Parenthesis := false
        ,  f_Cliboard := false
        ,  OldClipboard := ""
    static ToRemember := ""

	; OutputDebug, tu jestem
    if (InStr(ThisHotkey, "{"))
        f_Parenthesis   := true
    if (InStr(ThisHotkey, """"))
        f_Parenthesis   := true
    if (InStr(ThisHotkey, "("))
        f_Parenthesis   := true
    if (InStr(ThisHotkey, "["))
        f_Parenthesis   := true

    if (InStr(ThisHotkey, "+Right Up"))
        f_Cliboard      := true
    if (InStr(ThisHotkey, "+Left Up"))
        f_Cliboard      := true
    if (InStr(ThisHotkey, "^+Right Up"))
       f_Cliboard       := true
    if (InStr(ThisHotkey, "^+Left Up"))
        f_Cliboard      := true
    if (InStr(ThisHotkey, "LButton Up"))
        f_Cliboard      := true

	if (f_Parenthesis)
    {
	    ThisHotkey := SubStr(ThisHotkey, 0) ;extract last character
        Switch ThisHotkey
        {
            Case "(":   
                if (ToRemember)
                {
                    Send, % ToRemember . ")"
                    ToRemember := ""
                }
                else
                    Send, % ")" . "{Left}"

            Case "[":
                if (ToRemember)
                {
                    Send, % ToRemember . "]"
                    ToRemember := ""
                }
                else
                    Send, % "]" . "{Left}"

            Case "{":   
                if (ToRemember)
                {
                    Send, % ToRemember . "}"
                    ToRemember := ""
                }
                else
                    Send, % "{}}" . "{Left}"

            Case """":   
                if (ToRemember)
                {
                    Send, % ToRemember . """"
                    ToRemember := ""
                }
                else
                    Send, % """" . "{Left}"
        }

    }
    f_Parenthesis := false

    if (f_Cliboard)
    {
	    OldClipBoard := ClipboardAll
	    Clipboard := ""
	    Send, ^c
	    ClipWait, 0.1
	    ToRemember := Clipboard
	    ; OutputDebug, % "ToRemember:" . A_Tab . ToRemember
	    Clipboard := OldClipBoard
	    OldClipBoard := ""
    }
}
;*************************************************************************
F_BrowserMenu(MyBroArg)
{
	if (MyBroArg = "Yes")
	{
		Run, chrome.exe https://translate.google.com/ https://www.linkedin.com/feed/ https://mail.google.com/mail/u/0/#inbox http://www.meteo.pl/ https://trello.com/b/5h4R58KL/organizacyjne https://team.voestalpine.net/SitePages/Home.aspx https://helpdesk.tens.pl/helpdesk/ https://portal-signaling-poland.voestalpine.net/synergy/docs/Portal.aspx https://solidsystemteamwork.voestalpine.root.local/internalprojects/vaSupp/CPS/SitePages/Home.aspx https://solidsystemteamwork.voestalpine.root.local/Processes/custprojects/780MDSUpgradeKit/SitePages/Home.aspx 
		IniWrite, YES, VariousFunctions.ini, Menu memory, Browser
	}
	if (MyBroArg = "No")
	{
		IniWrite, NO, VariousFunctions.ini, Menu memory, Browser
	}
	return 
}
;*************************************************************************
F_SetEngKeyboardMenu(MySetKeyArg)
{
	global	;assume-global mode
	if (MySetKeyArg = "Yes")
		{
			SetDefaultKeyboard(English_USA)
			TrayTip, VariousFunctions.ahk, Keyboard style: English_USA, 5, 0x1 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Set English Keyboard
		}	
	
	if (MySetKeyArg = "No")
		{
			SetDefaultKeyboard(PolishLanguage)
			TrayTip, VariousFunctions.ahk, Keyboard style: PolishLanguage, 5, 0x1
			IniWrite, NO, VariousFunctions.ini, Menu memory, Set English Keyboard
		}			
}
;~ https://docs.microsoft.com/pl-pl/windows/win32/api/winuser/nf-winuser-systemparametersinfoa?redirectedfrom=MSDN
SetDefaultKeyboard(LocaleID)
{
	static SPI_SETDEFAULTINPUTLANG := 0x005A, SPIF_SENDWININICHANGE := 2
	WM_INPUTLANGCHANGEREQUEST := 0x50
	
	Language := DllCall("LoadKeyboardLayout", "Str", Format("{:08x}", LocaleID), "Int", 0)
	VarSetCapacity(binaryLocaleID, 4, 0)
	NumPut(LocaleID, binaryLocaleID)
	DllCall("SystemParametersInfo", UINT, SPI_SETDEFAULTINPUTLANG, UINT, 0, UPTR, &binaryLocaleID, UINT, SPIF_SENDWININICHANGE)
	
	WinGet, windows, List
	Loop % windows
		{
		PostMessage WM_INPUTLANGCHANGEREQUEST, 0, % Language, , % "ahk_id " windows%A_Index%
		}
}
;*************************************************************************
F_AltGr(MyAltArg)  ; redirects AltGr -> context menu; only in English keyboardLayout
;Ralt::AppsKey ; redirects AltGr -> context menu
{
	if (MyAltArg = "Yes")
	{
		Hotkey, RAlt, F_JustAlt, On
		IniWrite, YES, VariousFunctions.ini, Menu memory, AltGr
	}
		
	if (MyAltArg = "No")	
	{
		Hotkey, RAlt, F_JustAlt, Off
		IniWrite, NO, VariousFunctions.ini, Menu memory, AltGr
	}

}
F_JustAlt()
{   
Send, {AppsKey}
}
;*************************************************************************
F_WindowSwitcher(MyWinSWi) ; calls for windows switcher
{
	if (MyWinSwi = "Yes")
	{
		Hotkey,	LWin & LAlt, F_windowswitch, On 
		Hotkey,	LAlt & LWin, F_windowswitch, On 
		IniWrite, YES, VariousFunctions.ini, Menu memory, Window Switcher
	}	
		
	if (MyWinSwi = "No")
	{
		 Hotkey, LAlt & LWin, F_windowswitch, Off  
		 Hotkey, LWin & LAlt, F_windowswitch, Off  
		 IniWrite, NO, VariousFunctions.ini, Menu memory, Window Switcher
	}
	
}
F_windowswitch()
{
		Send, {Ctrl Down}{LAlt Down}{Tab}{LAlt Up}{Ctrl Up}
	return
}
;*************************************************************************
F_PrintScreen(Myprtscn)  ;https://support.microsoft.com/pl-pl/help/4488540/how-to-take-and-annotate-screenshots-on-windows-10
{
	if (Myprtscn = "Yes")
	{
		Hotkey, PrintScreen, F_prtscn, on 
		Hotkey, Volume_Down, F_voldown, on 
		IniWrite, YES, VariousFunctions.ini, Menu memory, Print Screen 
	}
	if (Myprtscn = "No")
	{
		Hotkey, PrintScreen, F_prtscn, off 
		Hotkey, Volume_Down, F_voldown, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Print Screen 
	}
}	
F_prtscn()
{ 
	Send, {Shift Down}{LWin Down}s{Shift Up}{LWin Up}
}

F_voldown()  ; run Snipping Tool (Microsoft Windows operating system tool)
{
	tooltip, [%A_thishotKey%] Run system tool Snipping Tool
	SetTimer, TurnOffTooltip, -5000
	Run, %A_WinDir%\system32\SnippingTool.exe
}
;*************************************************************************
F_Capitalize(MyCapslock)
{
#IfWinNotActive, ahk_exe WINWORD.EXE
	if (MyCapslock = "Yes")
	{
		Hotkey, IfWinNotActive, ahk_exe WINWORD.EXE
		Hotkey, +F3, ForceCapitalize
		IniWrite, YES, VariousFunctions.ini, Menu memory, Capitalize
	}

	if (MyCapslock = "No")
	{
		Hotkey, +F3, ForceCapitalize, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Capitalize
	}

}
;*************************************************************************
	;~ Author: Jakub Masiak
				;~ Refactored by: Hanna Ziętak on 2022-01-05
 ;~ * 				Input: one or more words (eg. Dog/Dog is jumping)
 ;~ * 					Shift+{F3} and selected sentence/word or a caret in the middle of a word
 ;~ * 				Output: Alters the first or all letters 
 ;~ * 					EXAMPLES:
 ;~ * 						Dog is jumping -> DOG IS JUMPING 
 ;~ * 						Dog -> DOG 
 ;~ * 						DOG IS JUMPING -> dog is jumping 
 ;~ * 						DOG -> dog 
 ;~ * 						dog is jumping -> Dog is jumping 
 ;~ * 						dog -> Dog  
				;~ shifts caret to the end 
				;~ it works everywhere exept Word, because in Word Application this function already exists 
 ;~ */

ForceCapitalize()	; by Jakub Masiak, revised by Hania Ziętak on 2021-12-20
{
	SelectWord := false												;Select Word := no (false)
	OldClipboard := ClipboardAll									;save content of clipboard to variable OldClipboard 
	Clipboard := ""													;clear content of clipboard
	Send, ^c														;copy to clipboard (ctrl + c)
	if (Clipboard = "")												;if clipboard is still empty, mark and copy word where caret is located
	{
		SelectWord := true											;Select Word := yes (true)
		Send, {Ctrl Down}{left}{Shift Down}{Right}{Shift Up}{Ctrl Up}  ;mark entire word ; do przerobienia
		Send, ^c														
	}
		ClipWait, 0														;wait until clipboard is full with anything
	state := "FirstState"											;Initial state
	Loop, Parse, Clipboard											;each character of Clipboard will be treated as a separate substring.
	{
		if A_LoopField is upper
		{
			if (state = "FirstState")
			{
				state := "UpperCaseState"							; "UpperCaseState" - a considered letter is uppercase 
			}
		}
		else if A_LoopField is lower
		{
			if (state = "FirstState")
			{
				state := "LowerCaseState"							; "LowerCaseState" - a considered letter is lowercase
			}
		}
		if (state = "UpperCaseState")
		{
			if A_Loopfield is lower
			{
				state := "AfterUpperCaseState"  					; "AfterUpperCaseState" - a considered letter is after a uppercase letter 
			}
		}
		if (state = "LowerCaseState")
		{
			if A_Loopfield is upper
			{
				state := "AfterUpperCaseState"
			}
		}
	}																;end of loop  ; the script is exiting the loop with the last letter status 

	if (state = "AfterUpperCaseState")
	{
		StringUpper, Clipboard, Clipboard
	}
	if (state = "UpperCaseState")
	{
		StringLower, Clipboard, Clipboard
	}
	if (state = "LowerCaseState")
	{
		FirstLetter := ""											; exit state of the loop  ; a previous letter (in a word) in second loop  
		NotAgain  := true											; flag: preventing capitalizing next letters  
		Loop, Parse, Clipboard										; this loop is for the case that we have a word or sentence with all small letters (eg. dog/dog is jumping) and the next case is one capital then all small letters (eg. Dog/Dog is jumping) 
		{
			WhoAmI := A_LoopField									; what is first or after the first letter: space, dot, end of line or next letter (which has to be small)
			if (WhoAmI = A_Space)
			{
				FirstLetter := % FirstLetter . A_Space
			}
			else if (WhoAmI = ".") or (WhoAmI = "`n") 
			{
				NotAgain := true	
				FirstLetter := FirstLetter . WhoAmI
			}
			else if (NotAgain = true) and (WhoAmI != A_Space)
			{
				StringUpper, WhoAmI, WhoAmI
				NotAgain := false
				FirstLetter := FirstLetter . WhoAmI
			}
			else
			{
				FirstLetter := FirstLetter . WhoAmI
			}
		}
		Clipboard := FirstLetter			
	}
	StringLen, Length, Clipboard								; counts letters in selected sentence/ word 
	Send, % "{Text}" . Clipboard
 	Sleep, 100
	Clipboard := OldClipboard
return
}
;*************************************************************************
F_McsWord(MyWord)
{ 
	if (MyWord="Yes")
	{
	 Hotkey, Media_Next, F_MediaNext, On 
	 IniWrite, YES, VariousFunctions.ini, Menu memory, Microsoft Word
	}
	
	if (MyWord="No")
	{
	 Hotkey, Media_Next, F_MediaNext, Off
	 IniWrite, NO, VariousFunctions.ini, Menu memory, Microsoft Word

	}

}

F_MediaNext()
{
	 tooltip, [%A_thishotKey%] Run text processor Microsoft Word  
	 SetTimer, TurnOffTooltip, -5000
	 Run, WINWORD.EXE
}
;*************************************************************************
F_TotalCommander(MyTotalCommander) ;run Total Commander app
{
	if (MyTotalCommander="Yes")
	{
		Hotkey, Media_Prev, F_MediaPrev, on
		IniWrite, YES, VariousFunctions.ini, Menu memory, Total Commander
	}

	if (MyTotalCommander="No")
	{
		Hotkey, Media_Prev, F_MediaPrev, Off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Total Commander
	}
	
}

F_MediaPrev()
{
	tooltip, [%A_thishotKey%] Run twin-panel file manager Total Commander
	SetTimer, TurnOffTooltip, -5000
	Run, c:\totalcmd\TOTALCMD64.EXE 
}
;*************************************************************************
F_Paint(MyPaint) ;run Paint app
{
	if (MyPaint="Yes")
	{
		Hotkey, Media_Play_Pause, F_MediaPlayPause, on
		IniWrite, YES, VariousFunctions.ini, Menu memory, Paint
	}

	if (MyPaint="No")
	{
		Hotkey, Media_Play_Pause, F_MediaPlayPause, Off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Paint
	}
	
}

F_MediaPlayPause()
{
	tooltip, [%A_ThisHotKey%] Run basic graphic editor Paint
	SetTimer, TurnOffTooltip, -5000
	Run, %A_WinDir%\system32\mspaint.exe
}
;*************************************************************************
F_Reboot(MyReboot) ; Reboot
{
	if (MyReboot="Yes")
	{
		Hotkey, ^Volume_Up, F_volup, on
		Hotkey, +^F2, F_volup, on 
		IniWrite, YES, VariousFunctions.ini, Menu memory, Reboot
	}

	if (MyReboot="No")
	{
		Hotkey, ^Volume_Up, F_volup, Off
		Hotkey, +^F2, F_volup, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Reboot
	}
	
}
F_volup()
{
 Shutdown, 2
 Msgbox, reboot!
}
;*************************************************************************
F_Shutdown(Myshutdown) ; Shutdown 
{
	if (Myshutdown="Yes")
	{
		Hotkey, ^Volume_Mute, F_volmute, On
		Hotkey, +^F3, F_volmute, On
		IniWrite, YES, VariousFunctions.ini, Menu memory, Shutdown
	}

	if (Myshutdown="No")
	{
		Hotkey, ^Volume_Mute, F_volmute, Off
		Hotkey, +^F3, F_volmute, Off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Shutdown
	}
	
}
F_volmute()
{
Shutdown, 1 + 8
Msgbox, power down!
}
;*************************************************************************
F_Transparency(MyTransp) ;toggle window transparency
{
		if (MyTransp = "Yes")
		{
			Hotkey, ^#F9, F_transp, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, Transparency 
		}
		if (MyTransp = "No")
		{
			Hotkey, ^#F9, F_transp, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Transparency
		}

}
F_transp()
{ 
	global 
	static WindowTransparency := false
	if (WindowTransparency = false)
		{
		WinSet, Transparent, 125, A
		WindowTransparency := true
		ToolTip, This window atribut Transparency was changed to semi-transparent ;, % A_CaretX, % A_CaretY - 20
		SetTimer, TurnOffTooltip, -2000
		return
		}
	else
		{
		WinSet, Transparent, 255, A
		WindowTransparency := false
		ToolTip, This window atribut Transparency was changed to opaque ;, % A_CaretX, % A_CaretY - 20
		;SetTimer, TurnOffTooltip, -2000
		return
		}
}
;*************************************************************************
F_F13(myf13)
{
	if (myf13="Yes")
	{
		Hotkey, F13, F_f13key, on 
		IniWrite, YES, VariousFunctions.ini, Menu memory, F13
	}
	if (myf13="No")
	{
		Hotkey, F13, F_f13key, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, F13
	}

}
;0000000000000000
F_f13key()
{
	Send, #t
	SoundBeep, 1000, 200
}
;00000000000000000000000000000000000000
F_F14(myf14)
{
	if (myf14="Yes")
	{
		Hotkey, F14, F_f14key, on 
		IniWrite, YES, VariousFunctions.ini, Menu memory, F14
	}
		if 	(myf14="No")
	{
		Hotkey, F14, F_f14key, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, F14
	}
}
;0000000000000000
F_f14key()
{
	msgbox, tu jestem 
	Hotstring("Reset")
	SoundBeep, 1500, 200 ; freq = 100, duration = 200 ms
	ToolTip, [%A_thishotKey%] reset of AutoHotkey string recognizer, % A_CaretX, % A_CaretY - 20
	SetTimer, TurnOffTooltip, -2000
}
;00000000000000000000000000000000000000
F_F15(myf15)
{
	if (myf15="Yes")
	{
		Hotkey, F15, F_f15key, on 
		IniWrite, YES, VariousFunctions.ini, Menu memory, F15
	}
	if 	(myf15="No")
	{
		Hotkey, F15, F_f15key, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, F15
	}
}
F_f15key()
{
	SoundBeep, 2000, 200
}
;*************************************************************************
F_Top(myalways)
{
	if (myalways= "Yes")
	{
		Hotkey, ^#F8, F_always, on 
		IniWrite, YES, VariousFunctions.ini, Menu memory, Always on top
	}
	if (myalways= "No")
	{
		Hotkey, ^#F8, F_always, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Always on top
	}
}
F_always()
{
	WinSet, AlwaysOnTop, toggle, A 
	ToolTip, This window atribut "Always on Top" is toggled ;, % A_CaretX, % A_CaretY - 20
	SetTimer, TurnOffTooltip, -2000
return
}
;*************************************************************************
F_KeePass(mykeepass)
{
	if (mykeepass= "Yes")
	{
		Hotkey, +^k, F_keepass2, on
		IniWrite, YES, VariousFunctions.ini, Menu memory, KeePass
	}
	if (mykeepass= "No")
	{
		Hotkey, +^k, F_keepass2, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, KeePass
	}
}
F_keepass2()
{
	Run, C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe 
}
;*************************************************************************
F_Hyper(myhyper)
{
	if (myhyper= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^k, F_hiper, on  
		IniWrite, YES, VariousFunctions.ini, Menu memory, Hyperlink
	
	}
	if (myhyper= "No")
	{
		Hotkey, ^k, F_hiper, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Hyperlink
	
	}
}
F_hiper()
{
	Send, {LAlt Down}{Ctrl Down}h{Ctrl Up}{LAlt Up}
}
;*************************************************************************
F_Hide(myhide)
{
	if (myhide= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, +^h, HideSelectedText, on  
		IniWrite, YES, VariousFunctions.ini, Menu memory, Hidetext
	
	}
	if (myhide= "No")
	{
		Hotkey, +^h, HideSelectedText, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Hidetext
	
	}
}
HideSelectedText() ; 2019-10-22 2019-11-08
{
	global oWord
	global  WordTrue, WordFalse

	oWord := ComObjActive("Word.Application")
	OurTemplate := oWord.ActiveDocument.AttachedTemplate.FullName
	if (InStr(OurTemplate, "TQ-S402-pl_OgolnyTechDok.dotm") or InStr(OurTemplate, "TQ-S402-en_OgolnyTechDok.dotm"))
	{
		nazStyl := oWord.Selection.Style.NameLocal
		if (nazStyl = "Ukryty ms")
			Send, ^{Space}
		else
		{
			language := oWord.Selection.Range.LanguageID
			oWord.Selection.Paragraphs(1).Range.LanguageID := language
			TemplateStyle("Ukryty ms")
		}
	}
	else
	{
		StateOfHidden := oWord.Selection.Font.Hidden
		oWord.Selection.Font.Hidden := WordTrue
		If (StateOfHidden == WordFalse)
		{
			oWord.Selection.Font.Hidden := WordTrue	
			}
		else
		{
			oWord.Selection.Font.Hidden := WordFalse
		}
	}
	
	oWord := "" ; Clear global COM objects when done with them
}
TemplateStyle(StyleName)
{
	global OurTemplateEN, OurTemplatePL, oWord
	StyleName := MsgText(StyleName)
	Base(StyleName)
	oWord := ComObjActive("Word.Application") 
	;~ SoundBeep, 750, 500 ; to fajnie dzia�a
	if  !(InStr(OurTemplate, "TQ-S402-pl_OgolnyTechDok.dotm") or InStr(OurTemplate, "TQ-S402-en_OgolnyTechDok.dotm"))
		{
		;~ MsgBox, % oWord.ActiveDocument.AttachedTemplate.FullName
		MsgBox, 16, % MsgText("Próba wywołania stylu z szablonu"), % MsgText("Próbujesz wywołać styl przypisany do szablonu, ale szablon nie został jeszcze dołączony do tego pliku.`nNajpierw dołącz szablon, a następnie wywołaj ponownie tę funkcję.")
		oWord := "" ; Clear global COM objects when done with them
		return
		}
	else
		{
		oWord.Selection.Style := StyleName
		oWord := "" ; Clear global COM objects when done with them
		return
		}
}
;*************************************************************************
F_Show(myshow)
{
	if (myshow= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^*, ShowHiddenText, on  
		IniWrite, YES, VariousFunctions.ini, Menu memory, Showtext
	
	}
	if (myshow= "No")
	{
		Hotkey, ^*, ShowHiddenText, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Showtext
	
	}
}
ShowHiddenText(AdditionalText := "")
;~ by Jakub Masiak
{
	global oWord
	Base(AdditionalText)
	oWord := ComObjActive("Word.Application")
	HiddenTextState := oWord.ActiveWindow.View.ShowHiddenText
	if (oWord.ActiveWindow.View.ShowAll = -1)
	{
		oWord.ActiveWindow.View.ShowAll := 0
		oWord.ActiveWindow.View.ShowTabs := 0
		oWord.ActiveWindow.View.ShowSpaces := 0
		oWord.ActiveWindow.View.ShowParagraphs := 0
		oWord.ActiveWindow.View.ShowHyphens := 0
		oWord.ActiveWindow.View.ShowObjectAnchors := 0
		oWord.ActiveWindow.View.ShowHiddenText := 0
	}
	else
	{
		oWord.ActiveWindow.View.ShowAll := -1
		oWord.ActiveWindow.View.ShowTabs := -1
		oWord.ActiveWindow.View.ShowSpaces := -1
		oWord.ActiveWindow.View.ShowParagraphs := -1
		oWord.ActiveWindow.View.ShowHyphens := -1
		oWord.ActiveWindow.View.ShowObjectAnchors := -1
		oWord.ActiveWindow.View.ShowHiddenText := -1
	}
	oWord := ""
	return
}
;*************************************************************************
F_TemplateOff(mytemplateff)
{
	if (mytemplateff= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^+t, F_mytemplateoff, on  
		IniWrite, YES, VariousFunctions.ini, Menu memory, Template Off
	
	}
	if (mytemplateff= "No")
	{
		Hotkey, ^+t,  F_mytemplateoff, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Template Off
	
	}
}
F_mytemplateoff()
{
oWord := ComObjActive("Word.Application")
OurTemplateOff := oWord.ActiveDocument.AttachedTemplate.FullName

if (InStr(OurTemplateOff, "TQ-S440-pl_DokUzyt.dotm") or InStr(OurTemplateOff, "TQ-S440-en_UserDoc.dotm"))
{
	oWord.ActiveDocument.AttachedTemplate := ""
	oWord.ActiveDocument.UpdateStylesOnOpen := -1
	MsgBox,0x40,, % MsgText("Szablon został odłączony.")
}
oWord := ""
return
}
;********************************************************************
F_AddTemplate(myaddtemplate)
{
	if (myaddtemplate= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^t, F_myaddtemplate, on  
		IniWrite, YES, VariousFunctions.ini, Menu memory, Add Template 
	
	}
	if (myaddtemplate= "No")
	{
		Hotkey, ^t,  F_myaddtemplate, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Add Template
	
	}
}
F_myaddtemplate()
{
	gosub, AutoTemplate
}
;*************************************************************************
F_Strikethrough(mystrikethro)
{
	if (mystrikethro= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^+x, StrikeThroughText, on  
		IniWrite, YES, VariousFunctions.ini, Menu memory, Strikethrough Text
	
	}
	if (mystrikethro= "No")
	{
		Hotkey, ^+x,  StrikeThroughText, off
		IniWrite, NO, VariousFunctions.ini, Menu memory, Strikethrough Text
	
	}
}
StrikeThroughText()
{
	global oWord
	global  WordTrue, WordFalse	

	oWord := ComObjActive("Word.Application")
	StateOfStrikeThrough := oWord.Selection.Font.StrikeThrough ; := wdToggle := 9999998 
	if (StateOfStrikeThrough == WordFalse)
		{
		oWord.Selection.Font.StrikeThrough := wdToggle := 9999998
		}
	else
		{
		oWord.Selection.Font.StrikeThrough := 0
		}
	oWord :=  "" ; Clear global COM objects when done with them
}
;*************************************************************************
F_DeleteLine(mydelete)
{
	if (mydelete= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^l, DeleteLineOfText, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Delete Line
	}
	if (mydelete= "No")
	{
		Hotkey, ^l, DeleteLineOfText, off 
		IniWrite, No, VariousFunctions.ini, Menu memory, Delete Line 
	}
}
DeleteLineOfText() ; 2019-10-03
{
	global oWord
	oWord := ComObjActive("Word.Application")
	oWord.Selection.HomeKey(Unit := wdLine := 5)
	oWord.Selection.EndKey(Unit := wdLine := 5, Extend := wdExtend := 1)
	oWord.Selection.Delete(Unit := wdCharacter := 1, Count := 1)
	oWord :=  "" ; Clear global COM objects when done with them
}
;*************************************************************************
F_AlignLeft(myalign)
{
	if (myalign= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, +^l, F_myalignleft, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Align Left
	}
	if (myalign= "No")
	{
		Hotkey, +^l, F_myalignleft, off 
		IniWrite, No, VariousFunctions.ini, Menu memory, Align Left
	}
}
F_myalignleft()
{
	Send, ^l
}
;*************************************************************************
F_ApplyStyles(mystyle)
{
	if (mystyle= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, +^s, ToggleApplyStylesPane, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Apply Styles 
	}
	if (mystyle= "No")
	{
		Hotkey, +^s, ToggleApplyStylesPane, off 
		IniWrite, No, VariousFunctions.ini, Menu memory, Apply Styles 
	}
}
ToggleApplyStylesPane()
{
	global oWord
	global  WordTrue, WordFalse	
	
	oWord := ComObjActive("Word.Application")
	; ApplyStylesTaskPane := oWord.CommandBars("Apply styles").Visible
	ApplyStylesTaskPane := oWord.Application.TaskPanes(17).Visible
	try
	{	
	If (ApplyStylesTaskPane = WordFalse)
		oWord.Application.TaskPanes(17).Visible := WordTrue
	Else If (ApplyStylesTaskPane = WordTrue)
		oWord.CommandBars("Apply styles").Visible := WordFalse
	}
		catch
	{
		MsgBox,48,, % MsgText("Aby wywołać panel ""Stosowanie stylów"", zaznaczenie nie powinno zawierać kanwy rysunku.")
		return
	}
	
	oWord := ""
}
;*************************************************************************
F_OpenPath(mypath)
{
	if (mystyle= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotkey, ^o, FullPath, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Open and Show Path
	}
	if (mystyle= "No")
	{
		Hotkey, ^o,  FullPath, off 
		IniWrite, No, VariousFunctions.ini, Menu memory, Open and Show Path
	}
}
FullPath(AdditionalText := "") ; display full path to a file in window title bar 
;~ by Jakub Masiak
{
	global oWord
    Base(AdditionalText)
	oWord := ComObjActive("Word.Application")
    oWord.ActiveWindow.Caption := oWord.ActiveDocument.FullName
    oWord := ""
	Send, ^{o down}{o up}
}
;*************************************************************************
F_Table(mytable)
{
	if (mytable= "Yes")
	{
		Hotkey, IfWinActive, ahk_exe WINWORD.EXE
		Hotstring(":*:tabela`t", "| | |", "on")
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Table
	}

	if (mytable= "No")
	{
		Hotstring(":*:tabela`t", "| | |`n", "off") 
		IniWrite, No, VariousFunctions.ini, Menu memory, Table
	}
}
;*************************************************************************
F_Suspend(mysuspend)
{
	if (mysuspend= "Yes")
	{
		Hotkey, +^F1, F_mysuspend1, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Suspend
	}
	if (mysuspend= "No")
	{
		Hotkey, +^F1,  F_mysuspend1, off 
		IniWrite, No, VariousFunctions.ini, Menu memory, Suspend
	}
}

F_mysuspend1()
{
	DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
}
;***************** ********************************************************
F_Volume(myvolume)
{
	FuncObj := func("MouseIsOver").bind("ahk_class Shell_TrayWnd")
	if (myvolume= "Yes") 
	{
		Hotkey, If, % FuncObj
		; Hotkey, IfWinActive, ahk_class Shell_TrayWnd
		Hotkey, WheelUp, 	F_mywheelup, 	on 
		Hotkey, WheelDown, 	F_mywheeldown, 	on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Volume Up & Down
	}
	if (myvolume= "No")
	{
		Hotkey, WheelUp, 	F_mywheelup, 	off
		Hotkey, WheelDown, 	F_mywheeldown, 	off
		IniWrite, No, VariousFunctions.ini, Menu memory, Volume Up & Down
	}
	Hotkey, If	;call off Hotkey if
}
MouseIsOver(WinTitle)
{
	MouseGetPos,,, Win
	return WinExist(WinTitle . " ahk_id " . Win)
}
F_mywheelup()
{
	Send {Volume_up}
}
F_mywheeldown()
{
	Send {Volume_down}
}
;***************** ********************************************************
F_BroWinSwi(mybrowinswi)
{
	if (mybrowinswi= "Yes")
	{
		Hotkey, Xbutton1, F_mybutton1, on 
		Hotkey, Xbutton2, F_mybutton2, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Browser Win Switcher 				
	}
	if (mybrowinswi= "No")
	{
		Hotkey, Xbutton1, F_mybutton1, off
		Hotkey, Xbutton2, F_mybutton2, off
		IniWrite, No, VariousFunctions.ini, Menu memory, Browser Win Switcher
	}
}

F_mybutton1()
{
	if !WinExist("ahk_class Chrome_WidgetWin_1")
		{
		Run, chrome.exe
		}
	if WinActive("ahk_class Chrome_WidgetWin_1") or WinActive("ahk_class TTOTAL_CMD")
		{
		Send, ^+{Tab}
		}
	else
		{
		WinActivate ahk_class Chrome_WidgetWin_1
		}
return	
}

F_mybutton2()
{
		if !WinExist("ahk_class Chrome_WidgetWin_1")
		{
		Run, chrome.exe
		}
	if WinActive("ahk_class Chrome_WidgetWin_1")  or WinActive("ahk_class TTOTAL_CMD")
		{
		Send, ^{Tab}
		}
	else
		{
		WinActivate ahk_class Chrome_WidgetWin_1
		}
return
}
;***************** ********************************************************
F_TransparencyMou(mytransmouse)
{
	if (mytransmouse= "Yes")
	{
		Hotkey, ^+WheelDown, F_MouseTranspdown, on 
		Hotkey, ^+WheelUp, F_MouseTranspup, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Tranparency Mouse
	}
	if (mytransmouse= "No")
	{
		Hotkey, ^+WheelDown, F_MouseTranspdown, off
		Hotkey, ^+WheelUp, F_MouseTranspup, off
		IniWrite, No, VariousFunctions.ini, Menu memory, Tranparency Mouse
	}
}

F_MouseTranspdown()
{
 TransFactor := TransFactor - 25.5
    if (TransFactor < 0)
        TransFactor := 0
    WinSet, Transparent, %TransFactor%, A
    TransProc := Round(100*TransFactor/255)
    ToolTip, Transparency set to %TransProc%`%
    SetTimer, TurnOffTooltip, -500
    Return 
}

F_MouseTranspup()
{
	TransFactor := TransFactor + 25.5
    if (TransFactor > 255)
        TransFactor := 255
    WinSet, Transparent, %TransFactor%, A  
    TransProc := Round(100*TransFactor/255)
    ToolTip, Transparency set to %TransProc%`%
    SetTimer, TurnOffTooltip, -500
    Return
}

F_Autosave(myautosave)
{
	if (myautosave= "Yes")
	{
		Hotkey,<!^q, F_myautosave, on 
		IniWrite, Yes, VariousFunctions.ini, Menu memory, Autosave
	}
	if (myautosave= "No")
	{
		Hotkey,<!^q, F_myautosave, off
		IniWrite, No, VariousFunctions.ini, Menu memory, Autosave
	}
}


F_myautosave()
{
	if (flag_as = 0)
	{
		SetTimer, AutoSave, Off
		TrayTip, %A_ScriptName%, Autozapis został wyłączony!, 5, 0x1
		flag_as := 1
	}
	else if (flag_as = 1)
	{
		SetTimer, AutoSave, On
		TrayTip, %A_ScriptName%, Autozapis został ponownie włączony!, 5, 0x1
		flag_as := 0
		return
	}
}




; ;~ https://jacks-autohotkey-blog.com/2020/03/09/auto-capitalize-the-first-letter-of-sentences/#more-41175
; CapitalizeFirstLetters()  ; czy ta funkcja jest jeszcze potrzebna?


- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
