#Requires AutoHotkey v1.1.33+ 	; Displays an error and quits if a version requirement is not met. 
#SingleInstance force 			; only one instance of this script may run at a time!
#NoEnv  						; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent
SendMode Input  				; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%		; Ensures a consistent starting directory.

Menu, Tray,Icon, % A_SCriptDir . "\vh.ico"  

#include %A_ScriptDir%\Lib\UIA_Interface.ahk

;------------------ SECTION OF GLOBAL VARIABLES: BEGINNING ---------------------------- 
global English_USA 		:= 0x0409   ; see AutoHotkey help: Language Codes https://www.autohotkey.com/docs/misc/Languages.htm
, PolishLanguage 		:= 0x0415	; https://www.autohotkey.com/docs/misc/Languages.htm 
, TransFactor 			:= 255
, WordTrue 				:= -1 ; ComObj(0xB, -1) ; 0xB = VT_Bool || -1 = true
, WordFalse 			:= 0 ; ComObj(0xB, 0) ; 0xB = VT_Bool || 0 = false
, OurTemplateEN 		:= "S:\OrgFirma\Szablony\Word\OgolneZmakrami\TQ-S440-en_UserDoc.dotm"
, OurTemplatePL 		:= "S:\OrgFirma\Szablony\Word\OgolneZmakrami\TQ-S440-pl_DokUzyt.dotm"
, OurTemplateOldPL		:= "S:\OrgFirma\Szablony\Word\OgolneZmakrami\TQ-S402-pl_OgolnyTechDok.dotm"
, OurTemplateOldEN		:= "S:\OrgFirma\Szablony\Word\OgolneZmakrami\TQ-S402-en_OgolnyTechDok.dotm"
, OurTemplate 			:= ""
;---------------- Zmienne do funkcji autozapisu ----------------
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

;/////////////////////////////// - INI SECTION - //////////////////////////////////////
F_IniRead()

;//////////////////////////////////////////////////////////////// - PREPARATION OF GUI ELEMENTS - //////////////////////////////////////////////////////////////////////////////////////////////////////////
F_PrepareGuiElements()
F_PrepareOpenTabsGui()

;//////////////////////////////////////////////////////////////// - ACTIVATION OF SELECTED FUNCTIONS - //////////////////////////////////////////////////////////////////////////////////////////////////////////
F_CB1_AlwaysOnTop() ;Always on top
F_CB2_AutomatePaint() ;Automate Paint
F_CB3_ChromeTabSwitcher() ;Chrome tab switcher
F_CB4_OpenTabs() ;Open tabs in Chrome
F_CB5_ParenthesisWatcher() ;Parenthesis watcher
F_CB6_USKeyboard() ;US Keyboard
F_CB7_RightClick() ;Right click context menu
F_CB8_VolumeUpDown() ;Volume up and down
F_CB9_WindowSwitcher() ;Window switcher
F_CB10_CAPSLOCK() ;Capitalization switcher CAPS LOCK
F_CB11_SHIFTF3() ;Capitalization switcher SHIFT +F3
F_CB12_FootswitchF13() ;FOOTSWITCH F13
F_CB13_FootswitchF14() ;FOOTSWITCH F14
F_CB14_FootswitchF15() ;FOOTSWITCH F15
F_CB15_TranslateEN_PL() ;Google translate en - pl
F_CB16_TranslatePL_EN() ;Google translate pl - en
F_CB17_Suspend() ;Power PC suspend 
F_CB18_Reboot() ;Power PC reboot
F_CB19_Shutdown() ;Powe PC shutdown
F_CB20_TransparencyMouse() ;Transparency switcher - mouse
F_CB21_TransparencyKeys() ;Transparency switcher - keys
F_CB22_AlignLeft() ;Align left
F_CB23_ApplyStyles() ;Apply styles
F_CB24_Autosave() ;Autosave
F_CB25_DeleteLine() ;Delete line
F_CB26_Hide() ;Hide
F_CB27_Show() ;Show
F_CB28_HyperLink() ;Hyperlink
F_CB29_OpenAndShowPath() ;Open and show path    
F_CB30_StrikethroughText() ;Strikethrough text
F_CB31_Table() ;Table
F_CB32_AddTemplate() ;Add template
F_CB33_TemplateOff() ;Template off
F_CB34_KeePass() ;Run KeePass
F_CB35_MSWord() ;Run MS Word
F_CB36_Paint() ;Run Paint
F_CB37_TotalCommander() ;Run Total Commander
F_CB38_PrintScreen() ;Run print screen

if (AutosaveIni)
	F_AutoSave()

;//////////////////////////////////////////////////////////////// - MENU TRAY - //////////////////////////////////////////////////////////////////////////////////////////////////////////
	Menu, Tray, NoStandard
	Menu, Tray, add, 		ChooseFunctions	
	Menu, Tray, Add 
	Menu, Tray, Standard
	Menu, Tray, Default, 	ChooseFunctions	;Menu, MenuName, Default [, MenuItemName]
return	;End of initialization

;////////////////////////////////////////////////////////////////////////// - GUI - /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
~+#v::
ChooseFunctions:
	GUI, submit, Hide
	Gui Show, w833 h512, Various functions
Return  ;end of auto-execute section


GuiEscape:
	GUI, submit, Hide
	Gui, Hide
Return

;////////////////////////////////////////////////////////////////// - ABOUT - ///////////////////////////////////////////
F_About1()
{
	Gui, submit, NoHide
	MsgBox,
		(
Authors:`nMaciej Słojewski, Hanna Ziętak, Jakub Masiak, Kasandra Krajewska`n`nInterface Author:`nKasandra Krajewska`n`nVersion: 1.1.2
		)
Return
}

;//////////////////////////////////////////////////////////////////// - ALWAYS ON TOP (1) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button1() 
{
	Gui, submit, NoHide
	Msgbox, 
		(
Toggle window parameter always on top, by pressing {Ctrl} + {Windows} + {F8}.
		)
Return
}


F_CB1_AlwaysOnTop()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check1 ;0, 1, -1
	{
		Case 0:
			Hotkey, ^#F8, F_always, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Always on top

        Case 1:
			Hotkey, ^#F8, F_always, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Always on top
	}
}

F_always()
{
    static OnOffToggle := false

	WinSet, AlwaysOnTop, toggle, A
    OnOffToggle := !OnOffToggle
    WinGetTitle, OutputVar, A
    MsgBox, 4096, % A_ScriptName . ":" . A_Space . "information", % "Changed AlwaysOnTop feature for the following window:" . "`n" . OutputVar . A_Space . "to:" . A_Space . (OnOffToggle ? "true" : "false")
    Return
}

;/////////////////////////////////////////////////////////////////// - AUTOMATE PAINT (2)- //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button2()
{
	Gui, Submit, NoHide
	Msgbox, 
		(
Rotate image by pressing F2, 
resize image by pressing F3, 
choose red rectangle by pressing F4, 
"save as" by pressing F5.
		)
}

F_CB2_AutomatePaint()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check2 ;0, 1, -1
		{
			Case 0:
				Hotkey, IfWinActive, AHK_class MSPaintApp
				Hotkey, F2, F_rotate, off
				Hotkey, F3, F_resize, off
				Hotkey, F4, F_choose, off 
				Hotkey, F5, F_saveas, off
				Hotkey, +/, F_GUI, off
				Hotkey, IfWinActive
				IniWrite, NO, VariousFunctions.ini, Menu memory, AutomatePaint

			Case 1:
				Hotkey, IfWinActive, AHK_class MSPaintApp
				Hotkey, F2, F_rotate, on 
				Hotkey, F3, F_resize, on 
				Hotkey, F4, F_choose, on 
				Hotkey, F5, F_saveas, on  
				Hotkey, +/, F_GUI, on
				Hotkey, IfWinActive
				IniWrite, YES, VariousFunctions.ini, Menu memory, AutomatePaint	
		}
}

	F_rotate()
	{
		UIA 		:= UIA_Interface() ; Initialize UIA interface
		WinWaitActive, ahk_exe mspaint.exe
		paint1 		:= UIA.ElementFromHandle(WinExist("ahk_exe mspaint.exe"))

		paint1.FindFirstByNameAndType("Obróć", "SplitButton").Click() 
		paint1.WaitElementExistByName("Obrót w prawo o 90°").Click()
		UIA 		:= ""
	,	paint1 		:= ""
	}

	F_resize()
	{
		UIA 		:= UIA_Interface() ; Initialize UIA interface
		WinWaitActive, ahk_exe mspaint.exe
		paint1 		:= UIA.ElementFromHandle(WinExist("ahk_exe mspaint.exe"))

		paint1.FindFirstByNameAndType("Zmień rozmiar", "Button").Click() 
		paint1.WaitElementExistByNameandType("Piksele","RadioButton").Click()
		paint1.FindFirstByNameandType("Zmień rozmiar w poziomie","Edit").SetValue("800")
		paint1.FindFirstByNameandType("OK","Button").Click()
		UIA 		:= ""
	,	paint1 		:= ""
	}

	F_choose()
	{
		UIA 		:= UIA_Interface() ; Initialize UIA interface
		WinWaitActive, ahk_exe mspaint.exe
		paint1 		:= UIA.ElementFromHandle(WinExist("ahk_exe mspaint.exe"))

		paint1.FindFirstByNameAndType("Kształty", "Button").Click() ;in polish language version 
		paint1.WaitElementExistByName("Zaokrąglony prostokąt").Click()
		Sleep,50
		paint1.FindFirstByNameAndType("Edytuj kolory", "Button").Click() 
		paint1.FindFirstByNameandType("Odc.:","Edit").SetValue(0)
		paint1.FindFirstByNameandType("Nas.:","Edit").SetValue(240)
		paint1.FindFirstByNameandType("Jaskr.:","Edit").SetValue(120)
		paint1.FindFirstByNameandType("Czerw.:","Edit").SetValue(255)
		paint1.FindFirstByNameandType("OK","Button").Click()
		UIA 		:= ""
	,	paint1 		:= ""	
	}

	F_saveas()
	{
		UIA 		:= UIA_Interface() ; Initialize UIA interface
		WinWaitActive, ahk_exe mspaint.exe
		paint1 		:= UIA.ElementFromHandle(WinExist("ahk_exe mspaint.exe"))

		paint1.FindFirstByNameAndType("Karta Plik", "Button").Click() 
		paint1.WaitElementExistByNameandType("Zapisz jako", "SplitButton").Click()
		UIA 		:= ""
	,	paint1 		:= ""
	}

	F_GUI()
	{
		Gui, PaintHelp: New
		Gui, PaintHelp: Font, s11, Arial
		Gui, PaintHelp: Add, Text, , F2: Right image rotation
		Gui, PaintHelp: Add, Text, , F3: Change image size to 800 px
		Gui, PaintHelp: Add, Text, , F4: Choose red rectangle
		Gui, PaintHelp: Add, Text, , F5: Save as
		Gui, PaintHelp: Add, Button, x300 gOK, OK
		Gui, PaintHelp: Show, w350, Shortcuts hints for MS Paint 
	}

OK:
	Gui, Submit, Hide  
Return

;/////////////////////////////////////////////////////////////////// - CHROME TAB SWITCHER (3)- //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button3()
{
	Gui, Submit, NoHide
	Msgbox, 
				(
ReturnSwitch tabs in Google Chrome Browser, by pressing {Xbutton1} and {Xbutton2}.
				)
Return
}

F_CB3_ChromeTabSwitcher()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check3 ;0, 1, -1
		{
			Case 0:
			Hotkey, Xbutton1, F_mybutton1, off
			Hotkey, Xbutton2, F_mybutton2, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Browser Win Switcher
		

			Case 1:
			Hotkey, Xbutton1, F_mybutton1, on 
			Hotkey, Xbutton2, F_mybutton2, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Browser Win Switcher
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

;/////////////////////////////////////////////////////////////////// - OPEN TABS IN CHROME (4) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button4()
{
	global 
	Gui OpenTabs: Show, w318 h395, Open tabs in Chrome
}

F_SAVE()
{
	global
	Gui, OpenTabs: Submit, Hide
	IniWrite, % Tab1,	VariousFunctions.ini, OpenTabs,	Tab1
	IniWrite, % Tab2,	VariousFunctions.ini, OpenTabs,	Tab2
	IniWrite, % Tab3,	VariousFunctions.ini, OpenTabs,	Tab3
	IniWrite, % Tab4,	VariousFunctions.ini, OpenTabs,	Tab4
	IniWrite, % Tab5,	VariousFunctions.ini, OpenTabs,	Tab5
	IniWrite, % Tab6,	VariousFunctions.ini, OpenTabs,	Tab6
	IniWrite, % Tab7,	VariousFunctions.ini, OpenTabs,	Tab7
	IniWrite, % Tab8,	VariousFunctions.ini, OpenTabs,	Tab8
	IniWrite, % Tab9,	VariousFunctions.ini, OpenTabs,	Tab9
	IniWrite, % Tab10,	VariousFunctions.ini, OpenTabs,	Tab10 
}

F_CB4_OpenTabs()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check4 ;0, 1, -1
		{
			Case 0:
			IniWrite, NO, VariousFunctions.ini, Menu memory, Browser
		
			Case 1:
			Run, % "chrome.exe" . A_space . Temp1
			WinWait, ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe
			WinMaximize, 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Browser
		}
}

;/////////////////////////////////////////////////////////////////// - PARENTHESIS WATCHER (5)- //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button5()
{
	Gui, Submit, NoHide
	Msgbox, 
				(
After pressing keys like: {  [  (  `" , the closing symbols }  ]  ) `" will also appear. Aditionally a caret will jump between the parenthesis/quotation marks. It works also, when you have already written a text and want to put it between parenthesis/quotation marks. You have to select words and press parenthesis/quotation marks.
				)
Return
}


F_CB5_ParenthesisWatcher()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide
	Switch Check5 ;0, 1, -1
		{
			Case 0:
			Hotkey, 	~{ , 			F_Parenthesis, Off
			Hotkey, 	~" , 			F_Parenthesis, Off
			Hotkey, 	~( , 			F_Parenthesis, Off
			Hotkey, 	~[ , 			F_Parenthesis, Off
			Hotkey,      ~+Right Up,    F_Parenthesis, Off	;events related to keyboard; order matters!
			Hotkey,      ~+Left Up,     F_Parenthesis, Off
			Hotkey,      ~^+Left Up,    F_Parenthesis, Off
			Hotkey,      ~^+Right Up,   F_Parenthesis, Off 
			IniWrite, NO, VariousFunctions.ini, Menu memory, Parenthesis

			Case 1:
			Hotkey, 	~{ , 			F_Parenthesis, On
			Hotkey, 	~" ,			F_Parenthesis, On
			Hotkey, 	~( , 			F_Parenthesis, On
			Hotkey, 	~[ , 			F_Parenthesis, On		
			Hotkey,     ~+Right Up,     F_Parenthesis, On	;events related to keyboard; order matters!
			Hotkey,     ~+Left Up,      F_Parenthesis, On
			Hotkey,     ~^+Left Up,     F_Parenthesis, On
			Hotkey,     ~^+Right Up,    F_Parenthesis, On
			IniWrite,   YES, VariousFunctions.ini, Menu memory, Parenthesis
		}
}

F_Parenthesis()
{	
    global 
	local ThisHotkey := A_ThisHotkey, f_Parenthesis := false
		, LastPressedKey := A_PriorKey
		, PreviousHotkey := A_PriorHotkey
        ,  f_Cliboard := false
        ,  OldClipboard := ""
    static ToRemember := "", PreviousClipboard := ""

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
	{
		Send, ^c
		ClipWait, 0	;wait until clipboard is full with anything
		PreviousClipboard := Clipboard
		OutputDebug, % "PreviousClipboard:" . A_Tab . PreviousClipboard . "`n"
	}
       
	if (PreviousHotkey = "~LButton Up") and (InStr(LastPressedKey, Shift))
	{
	    ToRemember := PreviousClipboard
		PreviousClipboard := ""
		f_Parenthesis := true
	}

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
					F_tooltip()
                }
                else
                    Send, % ")" . "{Left}"
					F_tooltip()

            Case "[":
                if (ToRemember)
                {
                    Send, % ToRemember . "]"
                    ToRemember := ""
					F_tooltip()
                }
                else
                    Send, % "]" . "{Left}"
					F_tooltip()

            Case "{":   
                if (ToRemember)
                {
                    Send, % ToRemember . "{}}"
                    ToRemember := ""
					F_tooltip()
                }
                else
				{
		        	Send, % "{}}" . "{Left}"
					F_tooltip()
				}
            Case """":   
                if (ToRemember)
                {
                    Send, % ToRemember . """"
                    ToRemember := ""
					F_tooltip()
                }
                else
				{
                    Send, % """" . "{Left}"
					F_tooltip()
				}
        }

    }
    f_Parenthesis := false

    if (f_Cliboard)
    {
	    OldClipBoard := ClipboardAll
	    Clipboard := ""
	    Send, ^c
	    ClipWait, 0	;wait until clipboard is full with anything
	    ToRemember := Clipboard
	    Clipboard := OldClipBoard
	    OldClipBoard := ""
    }
}

F_tooltip() 
{
	ToolTip, Parenthesis watcher (VariousFunctions.ahk), A_CaretX, A_CaretY - 20
	SetTimer, TurnOffTooltip, -1000 
}

;/////////////////////////////////////////////////////////////////// - US KEYBORD (6) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button6()
{
Gui, Submit, NoHide
	Msgbox, 
			(
Change keyboard settings (from Polish keyboard to English keyboard)
			)
Return
}

F_CB6_USKeyboard()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide
	Switch Check6 ;0, 1, -1
		{
			Case 0:
			SetDefaultKeyboard(PolishLanguage)
			TrayTip, VariousFunctions.ahk, Keyboard style: PolishLanguage, 5, 0x1
			IniWrite, NO, VariousFunctions.ini, Menu memory, Set English Keyboard			

			Case 1:
			SetDefaultKeyboard(English_USA)
			TrayTip, VariousFunctions.ahk, Keyboard style: English_USA, 5, 0x1 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Set English Keyboard
		}
}

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

;/////////////////////////////////////////////////////////////////// - RIGHT-CLICK CONTEXT MENU (7) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button7()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Redirects AltGr -> context menu`n
(only in English keyboard layout)
				)
Return
}

F_CB7_RightClick()
{
	global 
	Gui, Submit, NoHide 
	Switch Check7 ;0, 1, -1
		{
			Case 0:
			Hotkey, RAlt, F_JustAlt, Off
			IniWrite, NO, VariousFunctions.ini, Menu memory, AltGr
		
			Case 1:
			Hotkey, RAlt, F_JustAlt, On
			IniWrite, YES, VariousFunctions.ini, Menu memory, AltGr
		}
}

F_JustAlt()
{   
	Send, {AppsKey}
}

;/////////////////////////////////////////////////////////////////// - VOLUME UP AND DOWN (8) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button8()
{
	Gui, Submit, NoHide
	Msgbox, 
				(
ReturnTurn the volume up and down, by moving a mouse wheel. Works only when a caret is over the system tray.
				)
Return
}

#If MouseIsOver("ahk_class Shell_TrayWnd")
#If

F_CB8_VolumeUpDown()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check8 ;0, 1, -1
		{
			Case 0:
				Hotkey, If, MouseIsOver("ahk_class Shell_TrayWnd")
				Hotkey, WheelUp, 	F_mywheelup, 			off
				Hotkey, WheelDown, 	F_mywheeldown, 			off
				IniWrite, NO, VariousFunctions.ini, Menu memory, Volume Up & Down
				Hotkey, If
		
			Case 1:
				Hotkey, If, MouseIsOver("ahk_class Shell_TrayWnd")
				Hotkey, WheelUp, 	F_mywheelup, 			on 
				Hotkey, WheelDown, 	F_mywheeldown, 			on 
				IniWrite, YES, VariousFunctions.ini, Menu memory, Volume Up & Down
				Hotkey, If
		}
}


MouseIsOver(WinTitle)
{
	MouseGetPos,,, Win
	return WinExist(WinTitle . " ahk_id " . Win)
}

F_mywheelup()
{
	Send {Volume_Up}
}

F_mywheeldown()
{
	Send {Volume_Down}
}

;/////////////////////////////////////////////////////////////////// - WINDOW SWITCHER (9)- //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
F_Button9()
{
	Gui, Submit, NoHide
	Msgbox, 
				(
Switches between windows by pressing {Left Windows} key and {Left Alt} key, then you can move between windows by using ← → ↑ ↓ 
				)
Return
}

F_CB9_WindowSwitcher()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check9 ;0, 1, -1
		{
			Case 0:
			Hotkey,	LWin & LAlt, 	F_windowswitch, 	Off 
			Hotkey,	LAlt & LWin, 	F_windowswitch, 	Off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Window Switcher
		
			Case 1:
			Hotkey,	LWin & LAlt, 	F_windowswitch, 	On 
			Hotkey,	LAlt & LWin, 	F_windowswitch,		On 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Window Switcher
		}
}

F_windowswitch()
{
		Send, {Ctrl Down}{LAlt Down}{Tab}{LAlt Up}{Ctrl Up}
	return
}
;/////////////////////////////////////////////////////////////////// - CAPITALIZATION SWITCHER (10) (11)- //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;CAPSLOCK
F_Button10()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Press Capslock to change capitalization of the letters.
	EXAMPLES:
	Dog is jumping -> DOG IS JUMPING
	Dog -> DOG
	DOG IS JUMPING -> dog is jumping
	DOG -> dog
	dog is jumping -> Dog is jumping
	dog -> Dog
It works everywhere exept Word, because in Word Application this function already exists.
				)
Return
}

F_CB10_CAPSLOCK()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check10 ;0, 1, -1
		{
			Case 0:
			Hotkey, Capslock, ForceCapitalize, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Capitalize Capslock
		
			Case 1:
			Hotkey, IfWinNotActive, ahk_exe WINWORD.EXE
			Hotkey, Capslock, ForceCapitalize, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Capitalize Capslock
		}
}

;SHIFT + F3
F_Button11()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Press {Shift} + {F3} to change capitalization of the letters.
	EXAMPLES:
	Dog is jumping -> DOG IS JUMPING
	Dog -> DOG
	DOG IS JUMPING -> dog is jumping
	DOG -> dog
	dog is jumping -> Dog is jumping
	dog -> Dog
It works everywhere exept Word, because in Word Application this function already exists.
				)
Return
}

F_CB11_SHIFTF3()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check11 ;0, 1, -1
		{
			Case 0:
			Hotkey, +F3, ForceCapitalize, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Capitalize Shift
		
			Case 1:
			Hotkey, IfWinNotActive, ahk_exe WINWORD.EXE
			Hotkey, +F3, ForceCapitalize, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, Capitalize Shift
		}
}

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
		ClipWait, 0													;wait until clipboard is full with anything
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
		TooltipCap()
		
	}
	if (state = "UpperCaseState")
	{
		StringLower, Clipboard, Clipboard
		TooltipCap()
		
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
				TooltipCap()
			}
			else if (WhoAmI = ".") or (WhoAmI = "`n") 
			{
				NotAgain := true	
				FirstLetter := FirstLetter . WhoAmI
				TooltipCap()
			}
			else if (NotAgain = true) and (WhoAmI != A_Space)
			{
				StringUpper, WhoAmI, WhoAmI
				NotAgain := false
				FirstLetter := FirstLetter . WhoAmI
				TooltipCap()
			}
			else

			{
				FirstLetter := FirstLetter . WhoAmI
				TooltipCap()
			}
		}
		Clipboard := FirstLetter			
	}
	Send, % "{Text}" . Clipboard
 	Sleep, 100
	Clipboard := OldClipboard
	OldClipboard := ""
return
}

TooltipCap()
{
	ToolTip, Capitalization switcher (VariousFunctions.ahk), A_CaretX, A_CaretY - 20
	SetTimer, TurnOffTooltip, -1000 
}

;/////////////////////////////////////////////////////////////////// - FOOT SWITCH (12) (13) (14) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;F13
F_Button12()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Switch windows in the system tray, by pressing {F13}. 
				)
Return
}

F_CB12_FootswitchF13()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check12 ;0, 1, -1
		{
			Case 0:
			Hotkey, F13, F_f13key, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, F13
		
			Case 1:
			Hotkey, F13, F_f13key, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, F13
		}
}

F_f13key()
{
	Send, #t
	SoundBeep, 1000, 200
}

;F14
F_Button13()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Immediately resets the hotstring recognizer, by pressing {F14}. 
				)
Return
}

F_CB13_FootswitchF14()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check13 ;0, 1, -1
		{
			Case 0:
			Hotkey, F14, F_f14key, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, F14
		
			Case 1:
			Hotkey, F14, F_f14key, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, F14
		}
}

F_f14key()
{
	msgbox, Tu jestem! 
	Hotstring("Reset")
	SoundBeep, 1500, 200 ; freq = 100, duration = 200 ms
	ToolTip, [%A_thishotKey%] reset of AutoHotkey string recognizer, % A_CaretX, % A_CaretY - 20
	SetTimer, TurnOffTooltip, -2000
}

;F15
F_Button14()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Make a beep sound, by pressing {F15}.
				)
Return
}

F_CB14_FootswitchF15()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check14 ;0, 1, -1
		{
			Case 0:
			Hotkey, F15, F_f15key, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, F15
		
			Case 1:
			Hotkey, F15, F_f15key, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, F15
		}
}

F_f15key()
{

	SoundBeep, 2000, 200
}

;/////////////////////////////////////////////////////////////////// - GOOGLE TRANSLATE (15) (16) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;ENGLISH → POLISH
F_Button15()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Translate from English to Polish, by selecting text and pressing {Win} + {Ctrl} + t
				)
Return
}

F_CB15_TranslateEN_PL()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check15 ;0, 1, -1
		{
			Case 0:
			Hotkey, #^t, TranslationENtoPL, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, ENtoPL
		
			Case 1:
			Hotkey, #^t, TranslationENtoPL, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, ENtoPL
		}
}

;POLISH → ENGLISH
F_Button16()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Translate from Polish to English, by selecting text and pressing {Win} + t.
				)
Return
}

F_CB16_TranslatePL_EN()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check16 ;0, 1, -1
		{
			Case 0:
			Hotkey, #t, TranslationPLtoEN, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, PLtoEN
		
			Case 1:
			Hotkey, #t, TranslationPLtoEN, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, PLtoEN
		}
}

TranslationPLtoEN()
{											
	OldClipboard := ClipboardAll								 
	Clipboard := ""												
	Send, ^c																										
    ClipWait, 0	
    MsgBox, % OldClipboard  GoogleTranslate(Clipboard, "pl", "en")
	; TooltipTrans()	
    Clipboard:=  % GoogleTranslate(Clipboard, "pl", "en")
}

TranslationENtoPL()
{
	OldClipboard := ClipboardAll									
	Clipboard := ""													
	Send, ^c														
	ClipWait, 0
    MsgBox, % GoogleTranslate(Clipboard, "en", "pl")
	; TooltipTrans()
    Clipboard:= GoogleTranslate(Clipboard, "en", "pl")
}

;Author: https://www.autohotkey.com/boards/viewtopic.php?t=63835
GoogleTranslate(str, from := "auto", to := "en") {
   static JS := CreateScriptObj(), _ := JS.( GetJScript() ) := JS.("delete ActiveXObject; delete GetObject;")
   
   json := SendRequest(JS, str, to, from, proxy := "")
   oJSON := JS.("(" . json . ")")

   if !IsObject(oJSON[1]) {
      Loop % oJSON[0].length
         trans .= oJSON[0][A_Index - 1][0]
   }
   else {
      MainTransText := oJSON[0][0][0]
      Loop % oJSON[1].length {
         trans .= "`n+"
         obj := oJSON[1][A_Index-1][1]
         Loop % obj.length {
            txt := obj[A_Index - 1]
            trans .= (MainTransText = txt ? "" : "`n" txt)
         }
      }
   }
   if !IsObject(oJSON[1])
      MainTransText := trans := Trim(trans, ",+`n ")
   else
      trans := MainTransText . "`n+`n" . Trim(trans, ",+`n ")

   from := oJSON[2]
   trans := Trim(trans, ",+`n ")
   Return trans
}

SendRequest(JS, str, tl, sl, proxy) {
   static http
   ComObjError(false)
   if !http
   {
      http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
      ( proxy && http.SetProxy(2, proxy) )
      http.open("GET", "https://translate.google.com", true)
      http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0")
      http.send()
      http.WaitForResponse(-1)
   }
   http.open("POST", "https://translate.googleapis.com/translate_a/single?client=gtx"
      . "&sl=" . sl . "&tl=" . tl . "&hl=" . tl
      . "&dt=at&dt=bd&dt=ex&dt=ld&dt=md&dt=qca&dt=rw&dt=rm&dt=ss&dt=t&ie=UTF-8&oe=UTF-8&otf=0&ssel=0&tsel=0&pc=1&kc=1"
      . "&tk=" . JS.("tk").(str), true)

   http.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=utf-8")
   http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0")
   http.send("q=" . URIEncode(str))
   http.WaitForResponse(-1)
   Return http.responsetext
}

URIEncode(str, encoding := "UTF-8")  {
   VarSetCapacity(var, StrPut(str, encoding))
   StrPut(str, &var, encoding)

   while code := NumGet(Var, A_Index - 1, "UChar")  {
      bool := (code > 0x7F || code < 0x30 || code = 0x3D)
      UrlStr .= bool ? "%" . Format("{:02X}", code) : Chr(code)
   }
   Return UrlStr
}

GetJScript()
{
   script =
   (
      var TKK = ((function() {
        var a = 561666268;
        var b = 1526272306;
        return 406398 + '.' + (a + b);
      })());

      function b(a, b) {
        for (var d = 0; d < b.length - 2; d += 3) {
            var c = b.charAt(d + 2),
                c = "a" <= c ? c.charCodeAt(0) - 87 : Number(c),
                c = "+" == b.charAt(d + 1) ? a >>> c : a << c;
            a = "+" == b.charAt(d) ? a + c & 4294967295 : a ^ c
        }
        return a
      }

      function tk(a) {
          for (var e = TKK.split("."), h = Number(e[0]) || 0, g = [], d = 0, f = 0; f < a.length; f++) {
              var c = a.charCodeAt(f);
              128 > c ? g[d++] = c : (2048 > c ? g[d++] = c >> 6 | 192 : (55296 == (c & 64512) && f + 1 < a.length && 56320 == (a.charCodeAt(f + 1) & 64512) ?
              (c = 65536 + ((c & 1023) << 10) + (a.charCodeAt(++f) & 1023), g[d++] = c >> 18 | 240,
              g[d++] = c >> 12 & 63 | 128) : g[d++] = c >> 12 | 224, g[d++] = c >> 6 & 63 | 128), g[d++] = c & 63 | 128)
          }
          a = h;
          for (d = 0; d < g.length; d++) a += g[d], a = b(a, "+-a^+6");
          a = b(a, "+-3^+b+-f");
          a ^= Number(e[1]) || 0;
          0 > a && (a = (a & 2147483647) + 2147483648);
          a `%= 1E6;
          return a.toString() + "." + (a ^ h)
      }
   )
   Return script
}

CreateScriptObj() {
   static doc, JS, _JS
   if !doc {
      doc := ComObjCreate("htmlfile")
      doc.write("<meta http-equiv='X-UA-Compatible' content='IE=9'>")
      JS := doc.parentWindow
      if (doc.documentMode < 9)
         JS.execScript()
      _JS := ObjBindMethod(JS, "eval")
   }
   Return _JS
}

	Clipboard := OldClipboard
return

TooltipTrans()
{
	ToolTip, Google Translate (VariousFunctions.ahk), A_CaretX, A_CaretY - 20
	SetTimer, TurnOffTooltip, -1000
}

;/////////////////////////////////////////////////////////////////// - POWER PC (17) (18) (19) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;SUSPEND
F_Button17()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Suspend, by pressing {Ctrl} + {shift} + F1
				)
Return
}

F_CB17_Suspend()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check17 ;0, 1, -1
		{
			Case 0:
			Hotkey, +^F1,  F_mysuspend1, off 
			IniWrite, NO, VariousFunctions.ini, Menu memory, Suspend
		
			Case 1:
			Hotkey, +^F1, F_mysuspend1, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Suspend
		}
}

F_mysuspend1()
{
	DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
}

;REBOOT
F_Button18()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Reboot by pressing a Multimedia key - {Ctrl}+{Volume Up} or {Ctrl} + {Shift} + {F2}
				)
Return
}

F_CB18_Reboot()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check18 ;0, 1, -1
		{
			Case 0:
			Hotkey, ^Volume_Up, F_volup, 		Off
			Hotkey, +^F2, 		F_volup, 		off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Reboot
		
			Case 1:
			Hotkey, ^Volume_Up, F_volup, 		on
			Hotkey, +^F2, 		F_volup, 		on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Reboot
		}
}

F_volup()
{
	Shutdown, 2
}

;SHUTDOWN
F_Button19()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Shutdown system, by pressing a Multimedia key - {Ctrl}+{Volume Mute} or {Ctrl} + {Shift} + {F3}
				)
Return
}

F_CB19_Shutdown()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check19 ;0, 1, -1
		{
			Case 0:
			Hotkey, ^Volume_Mute, 	F_volmute, 		Off
			Hotkey, +^F3, 			F_volmute, 		Off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Shutdown
		
			Case 1:
			Hotkey, ^Volume_Mute, 	F_volmute, 		On
			Hotkey, +^F3, 			F_volmute, 		On
			IniWrite, YES, VariousFunctions.ini, Menu memory, Shutdown
		}
}

F_volmute()
{
	Shutdown, 1 + 8
}

;/////////////////////////////////////////////////////////////////// - TRANSPARENCY SWITCHER (20) (21) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;MOUSE
F_Button20()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Smooth toggle window tranparency, by moving mouse wheel and pressing {Ctrl}+{Shift}.
				)
Return
}

F_CB20_TransparencyMouse()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check20 ;0, 1, -1
		{
			Case 0:
			Hotkey, ^+WheelDown, 	F_MouseTranspdown, 	off
			Hotkey, ^+WheelUp, 		F_MouseTranspup, 	off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Tranparency Mouse
		
			Case 1:
			Hotkey, ^+WheelDown,	 F_MouseTranspdown, 	on 
			Hotkey, ^+WheelUp,		 F_MouseTranspup, 		on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Tranparency Mouse
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

;KEYS
F_Button21()
{
	Gui, Submit, NoHide
	Msgbox, 
				(
Toggle window transparency by pressing {Ctr} + {Windows} + {F9} by half.
				)
Return
}

F_CB21_TransparencyKeys()
{
	global		;assume-global mode of operation
	Gui, Submit, NoHide 
	Switch Check21 ;0, 1, -1
		{
			Case 0:
			Hotkey, ^#F9, F_transp, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Transparency
		
			Case 1:
			Hotkey, ^#F9, F_transp, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, Transparency
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

;//////////////////////////////////////////// - FUNCTIONS IN MS WORD (22) (23) (24) (25) (26) (27) (28) (29) (30) (31) (32) (32) (33) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;ALIGN LEFT (22)
F_Button22()
{
	Gui, Submit, NoHide
	Msgbox, 
				(
Align your content with the left margin in Microsoft Word, by pressing {Ctrl} + {Shift} + l.
				)
Return
}

F_CB22_AlignLeft()
{
	global
	Gui, Submit, NoHide 
	Switch Check22 ;0, 1, -1
		{
			Case 0:
			Hotkey, +^l, F_myalignleft, off 
			IniWrite, NO, VariousFunctions.ini, Menu memory, Align Left
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, +^l, F_myalignleft, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Align Left
		}
}

F_myalignleft()
{
	Send, ^l
}

;APPLY STYLES (23)
F_Button23()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Open and close the Apply Styles window in Microsoft Word, by pressing {Ctrl} + {shift} + s.
				)
Return
}

F_CB23_ApplyStyles()
{
	global
	Gui, Submit, NoHide 
	Switch Check23 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, +^s, ToggleApplyStylesPane, off 
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Apply Styles
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, +^s, ToggleApplyStylesPane, on 
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Apply Styles 
		}
}

ToggleApplyStylesPane()
{
	global oWord
	global  WordTrue, WordFalse	
	
	oWord := ComObjActive("Word.Application")
	ApplyStylesTaskPane := oWord.Application.TaskPanes(17).Visible

	If (ApplyStylesTaskPane = WordFalse)
		oWord.Application.TaskPanes(17).Visible 	:= WordTrue
	If (ApplyStylesTaskPane = WordTrue)
		oWord.CommandBars("Apply styles").Visible 	:= WordFalse

	oWord := ""
}


;AUTOSAVE (24)
F_Button24()
{
Gui, Submit, NoHide
	Msgbox, 
				(
The function starts autosave of word documents, every 10 min, if the file size has changed. The copy is saved in the path: C:\temp1\KopiaZapasowaPlikowWord.
				)
Return
}

F_CB24_Autosave()
{
	global
	Gui, Submit, NoHide 
	Switch Check24 ;0, 1, -1
		{
			Case 0:
			IniWrite, NO, VariousFunctions.ini, Menu memory, Autosave
			AutosaveIni := 0
			SetTimer, F_AutoSave, Off
			TrayTip, %A_ScriptName%, Autozapis został wyłączony!, 5, 0x1

			Case 1:
			IniWrite, YES, VariousFunctions.ini, Menu memory, Autosave
			AutosaveIni := 1
			SetTimer, F_AutoSave, % interval
			TrayTip, %A_ScriptName%, Autozapis został włączony!, 5, 0x1
		}
}

F_AutoSave()
{
	InitAutosaveFilePath(AutosaveFilePath)
	
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
		SetTimer, F_AutoSave, 5000	; try again in 5 seconds
		return
	}
	; reset the timer in case it was changed by catch
	SetTimer, F_AutoSave, % interval
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

;DELETE LINE (25)
F_Button25()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Delete whole text line in Microsoft Word, by pressing {Ctrl} + l.
				)
Return
}

F_CB25_DeleteLine()
{
	global
	Gui, Submit, NoHide 
	Switch Check25 ;0, 1, -1
		{
			Case 0:
			Hotkey, ^l, DeleteLineOfText, off 
			IniWrite, NO, VariousFunctions.ini, Menu memory, Delete Line 
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^l, DeleteLineOfText, on 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Delete Line
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

;HIDDEN TEXT (26) (27)
;HIDE (26)
F_Button26()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Set "Ukryty ms" style of selected text in Microsoft Word, by pressing {Shift} + {Ctrl} + h (in out template only).
				)
Return
}

F_CB26_Hide()
{
	global
	Gui, Submit, NoHide 
	Switch Check26 ;0, 1, -1
		{
			Case 0:
			Hotkey, +^h, HideSelectedText, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Hidetext 
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, +^h, HideSelectedText, on  
			IniWrite, YES, VariousFunctions.ini, Menu memory, Hidetext
}
}

HideSelectedText() ; 2019-10-22 2019-11-08
{
	global oWord
	global  WordTrue, WordFalse

	oWord 		:= ComObjActive("Word.Application")
,	OurTemplate := oWord.ActiveDocument.AttachedTemplate.FullName
	if (InStr(OurTemplate, OurTemplateEN) or InStr(OurTemplate, OurTemplatePL) or InStr(OurTemplate, OurTemplateOldPL) or InStr(OurTemplate, OurTemplateOldEN)) ;if template is attached
	{
		nazStyl := oWord.Selection.Style.NameLocal	;nazStyl = set style of currently selected text
		if (nazStyl = "Ukryty ms")					;if style of selected text is "Ukryty ms", give this text the default formatting (keyboard shortcut Ctrl + Space bar)
			Send, ^{Space}
		else
		{
			language := oWord.Selection.Range.LanguageID
			oWord.Selection.Paragraphs(1).Range.LanguageID := language	;set currently language for selected text
			oWord.Selection.Style := "Ukryty ms"
		}
	}
	else	;if template is not attached
	{
		StateOfHidden 					:= oWord.Selection.Font.Hidden
,		oWord.Selection.Font.Hidden 	:= WordTrue
		if (StateOfHidden == WordFalse)
			oWord.Selection.Font.Hidden := WordTrue	
		else
			oWord.Selection.Font.Hidden := WordFalse
	}
	oWord := "" ; Clear global COM objects when done with them
}

;HIDDENTEXT (26) (27)
;SHOW (27)
F_Button27()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Enables and disables hidden text and non-printing characters in Microsoft Word, by pressing {Ctrl} + *.
				)
Return
}

F_CB27_Show()
{
	global
	Gui, Submit, NoHide 
	Switch Check27 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^*, ShowHiddenText, off
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Showtext 
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^*, ShowHiddenText, on  
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Showtext
		}
}

ShowHiddenText(AdditionalText := "")
;~ by Jakub Masiak
{
	global oWord
	oWord 			:= ComObjActive("Word.Application")
,	HiddenTextState := oWord.ActiveWindow.View.ShowHiddenText

	if (oWord.ActiveWindow.View.ShowAll = WordTrue)
	{
		oWord.ActiveWindow.View.ShowAll 			:= WordFalse
		oWord.ActiveWindow.View.ShowTabs 			:= WordFalse
		oWord.ActiveWindow.View.ShowSpaces 			:= WordFalse
		oWord.ActiveWindow.View.ShowParagraphs 		:= WordFalse
		oWord.ActiveWindow.View.ShowHyphens 		:= WordFalse
		oWord.ActiveWindow.View.ShowObjectAnchors 	:= WordFalse
		oWord.ActiveWindow.View.ShowHiddenText 		:= WordFalse
	}
	else
	{
		oWord.ActiveWindow.View.ShowAll 			:= WordTrue
		oWord.ActiveWindow.View.ShowTabs 			:= WordTrue
		oWord.ActiveWindow.View.ShowSpaces 			:= WordTrue
		oWord.ActiveWindow.View.ShowParagraphs 		:= WordTrue
		oWord.ActiveWindow.View.ShowHyphens 		:= WordTrue
		oWord.ActiveWindow.View.ShowObjectAnchors 	:= WordTrue
		oWord.ActiveWindow.View.ShowHiddenText 		:= WordTrue
	}
	oWord := ""
}

;HYPERLINK (28)
F_Button28()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Add hyperlink in selected text in Microsoft Word, by pressing {Ctrl} + k.
				)
Return
}

F_CB28_HyperLink()
{
	global
	Gui, Submit, NoHide 
	Switch Check28 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^k, F_hiper, off
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Hyperlink
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^k, F_hiper, on 
			Hotkey, IfWinActive 
			IniWrite, YES, VariousFunctions.ini, Menu memory, Hyperlink
		}
}

F_hiper()
{
	Send, {LAlt Down}{Ctrl Down}h{Ctrl Up}{LAlt Up}
}

;OPEN AND SHOW PATH (29)
F_Button29()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Open "Open" window and show the path of a document on the top bar of the document in Microsoft Word, by pressing {Ctrl} + o and esc.
				)
Return
}

F_CB29_OpenAndShowPath()
{
	global
	Gui, Submit, NoHide 
	Switch Check29 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^o,  FullPath, off 
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Open and Show Path
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^o, FullPath, on 
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Open and Show Path
		}
}

FullPath(AdditionalText := "") ; display full path to a file in window title bar 
;~ by Jakub Masiak
{
	global oWord
    ; Base(AdditionalText)
	oWord := ComObjActive("Word.Application")
    oWord.ActiveWindow.Caption := oWord.ActiveDocument.FullName
    oWord := ""
	Send, ^{o down}{o up}
}

;STRIKETHROUGH TEXT (30)
F_Button30()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Strike selected text through, by pressing {Ctrl} + {Shift} + x.
				)
Return
}

F_CB30_StrikethroughText()
{
	global
	Gui, Submit, NoHide 
	Switch Check30 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^+x,  StrikeThroughText, off
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Strikethrough Text
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^+x, StrikeThroughText, on  
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Strikethrough Text
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

;TABLE (31)
F_Button31()
{
Gui, Submit, NoHide
	Msgbox, 
				(
After typing "tabela" + tab, you receive | | | + {Enter}. You recive table in Microsoft Word. 
				)
Return
}

F_CB31_Table()
{
	global
	Gui, Submit, NoHide 
	Switch Check31 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotstring(":*:tabela`t", "| | |`n", "off")
			Hotkey, IfWinActive 
			IniWrite, NO, VariousFunctions.ini, Menu memory, Table
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotstring(":*:tabela`t", "| | |", "on")
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Table
		}
}

;TEMPLATE (32) (33)
;ADD (32)
F_Button32()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Add Polish or English template in Microsoft Word, by pressing {Ctrl} + t.
				)
Return
}

F_CB32_AddTemplate()
{
	global
	Gui, Submit, NoHide 
	Switch Check32 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^t,  F_myaddtemplate, off
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Add Template
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^t, F_myaddtemplate, on  
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Add Template 
		}
}

F_myaddtemplate()
{
	gosub, AutoTemplate
}
AutoTemplate:
	oWord := ComObjActive("Word.Application")
	try
		var_PopSzab := oWord.ActiveDocument.CustomDocumentProperties["PopSzab"].Value ;PopSzab Sets the value depending on the selected template (File - information - properties - advanced properties).
	catch
	{
		oWord.ActiveDocument.CustomDocumentProperties.Add("PopSzab",0,4," ")
		var_PopSzab := oWord.ActiveDocument.CustomDocumentProperties["PopSzab"].Value
	}
	if ((var_PopSzab == "PL") or (var_PopSzab == "EN") or (var_PopSzab == "OldEN") or (var_PopSzab == "OldPL")) ;if there was already a template plugged in the file (we know this from PopSzab), it automatically sets the last template
	{
		gosub, AddTemplate
	}
	else
		gosub, ChooseTemplate ;if the file did not already have a template attached allows you to select a template
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
			MsgBox, 64,, Dołączono szablon! 
			OurTemplate := OurTemplatePL
		}
	}
	if (template == "OldPL")
	{
		if (OurTemplate == OurTemplateOldPL)
		{
			oWord := ""
			
		}
		else
		{
			oWord.ActiveDocument.AttachedTemplate := OurTemplateOldPL
			oWord.ActiveDocument.UpdateStylesOnOpen := WordTrue
			oWord.ActiveDocument.UpdateStyles
			MsgBox, 64,, Dołączono szablon! 
			OurTemplate := OurTemplateOldPL
		}
	}
	if (template == "EN")
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
			MsgBox, 64,, The template is added!
			OurTemplate := OurTemplateEN
		}
	}
	if (template == "OldEN")
	{
		if (OurTemplate == OurTemplateOldEN)
		{
			oWord := ""
			
		}
		else
		{
			oWord.ActiveDocument.AttachedTemplate := OurTemplateOldEN
			oWord.ActiveDocument.UpdateStylesOnOpen :=  WordTrue
			oWord.ActiveDocument.UpdateStyles
			MsgBox, 64,, The template is added!
			OurTemplate := OurTemplateOldEN
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
		Gui, Temp:Add, Radio, vMyTemplate Checked, Polish template User Doc
		Gui, Temp:Add, Radio,, English template User Doc
		Gui, Temp:Add, Radio,, Polish template OgólnyTechDok
		Gui, Temp:Add, Radio,, English template OgólnyTechDok
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
	if (MyTemplate == 2)
	{
		template := "EN"
	}
	if (MyTemplate == 3)
	{
		template := "OldPL"
	}
	if (MyTemplate == 4)
	{
		template := "OldEN"
	}
	gosub, AddTemplate
	return


BBOK:
	Gui, BB:Submit, +OwnDialogs
	Gui, BB:Destroy
	if (FirstPage == 1)
		BB_Insert("Pierwsza strona zwykła", "")
	if (ID == 1)
		BB_Insert("identyfikator", "")
	if (ChangeLog == 1)
		BB_Insert("Change log", "")
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
		BB_Insert("Okładka tył", "")
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
	oWord := ComObjActive("Word.Application")
	if  !(InStr(OurTemplate, OurTemplateEN) or InStr(OurTemplate, OurTemplatePL) or InStr(OurTemplate, OurTemplateOldPL) or InStr(OurTemplate, OurTemplateOldEN))
		{
		MsgBox, 16, Próba wstawienia bloku z szablonu, Próbujesz wstawić blok konstrukcyjny przypisany do szablonu, ale szablon nie zostać jeszcze dołączony do tego pliku.`nNajpierw dołącz szablon, a nastepnie wywołaj ponownie tę funkcję.
		}
	else
		{
		OurTemplate := oWord.ActiveDocument.AttachedTemplate.FullName
		oWord.Templates(OurTemplate).BuildingBlockEntries(Name_BB).Insert(oWord.Selection.Range, WordTrue)
		}
	oWord :=  "" ; Clear global COM objects when done with them
	}
return

;TEMPLATE (32) (33)
;OFF (33)
F_Button33()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Switch off added template, by pressing {Ctrl} + {Shift} + t.
				)
Return
}

F_CB33_TemplateOff()
{
	global
	Gui, Submit, NoHide 
	Switch Check33 ;0, 1, -1
		{
			Case 0:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^+t, F_mytemplateoff, off
			Hotkey, IfWinActive
			IniWrite, NO, VariousFunctions.ini, Menu memory, Template Off
		
			Case 1:
			Hotkey, IfWinActive, ahk_exe WINWORD.EXE
			Hotkey, ^+t, F_mytemplateoff, on  
			Hotkey, IfWinActive
			IniWrite, YES, VariousFunctions.ini, Menu memory, Template Off
		}
}

F_mytemplateoff()
{
oWord := ComObjActive("Word.Application")
OurTemplateOff := oWord.ActiveDocument.AttachedTemplate.FullName

if (InStr(OurTemplate, OurTemplateEN) or InStr(OurTemplate, OurTemplatePL) or InStr(OurTemplate, OurTemplateOldPL) or InStr(OurTemplate, OurTemplateOldEN))
{
	oWord.ActiveDocument.AttachedTemplate := ""
	oWord.ActiveDocument.UpdateStylesOnOpen := -1
	MsgBox,0x40,, Szablon został odłączony.
}
oWord := ""
return
}
;/////////////////////////////////////////////////////////////////// - RUN (34) (35) (36) (37) (38) - //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;KEEPASS (34)
F_Button34()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Run the KeePass 2 application, by pressing {Shift} + {Ctrl} + k 
				)
Return
}

F_CB34_KeePass()
{
	global
	Gui, Submit, NoHide 
	Switch Check34 ;0, 1, -1
		{
			Case 0:
			Hotkey, +^k, F_keepass2, off
			IniWrite, NO, VariousFunctions.ini, Menu memory, KeePass
		
			Case 1:
			Hotkey, +^k, F_keepass2, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, KeePass
		}
}

F_keepass2()
{
	Run, C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe 
}

;MS WORD (35)
F_Button35()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Run Microsoft Word application by pressing {Media_Next} → Multimedia keys
				)
Return
}

F_CB35_MSWord()
{
	global
	Gui, Submit, NoHide 
	Switch Check35 ;0, 1, -1
		{
			Case 0:
			Hotkey, Media_Next, F_MediaNext, Off
	 		IniWrite, NO, VariousFunctions.ini, Menu memory, Microsoft Word
		
			Case 1:
			Hotkey, Media_Next, F_MediaNext, On 
	 		IniWrite, YES, VariousFunctions.ini, Menu memory, Microsoft Word
		}
}

F_MediaNext()
{
	 tooltip, [%A_thishotKey%] Run text processor Microsoft Word  
	 SetTimer, TurnOffTooltip, -5000
	 Run, WINWORD.EXE
}

;PAINT (36)
F_Button36()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Run Paint application by pressing {Media_Play_Pause} → Multimedia keys
				)
Return
}

F_CB36_Paint()
{
	global
	Gui, Submit, NoHide 
	Switch Check36 ;0, 1, -1
		{
			Case 0:
			Hotkey, Media_Play_Pause, F_MediaPlayPause, Off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Paint
		
			Case 1:
			Hotkey, Media_Play_Pause, F_MediaPlayPause, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, Paint
		}
}

F_MediaPlayPause()
{
	tooltip, [%A_ThisHotKey%] Run basic graphic editor Paint
	SetTimer, TurnOffTooltip, -5000
	Run, %A_WinDir%\system32\mspaint.exe
}

;TOTAL COMMANDER (37)
F_Button37()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Run Total Commander application by pressing - {Media_Prev} → Multimedia keys
				)
Return
}

F_CB37_TotalCommander()
{
	global
	Gui, Submit, NoHide 
	Switch Check37 ;0, 1, -1
		{
			Case 0:
			Hotkey, Media_Prev, F_MediaPrev, Off
			IniWrite, NO, VariousFunctions.ini, Menu memory, Total Commander
		
			Case 1:
			Hotkey, Media_Prev, F_MediaPrev, on
			IniWrite, YES, VariousFunctions.ini, Menu memory, Total Commander
		}
}

F_MediaPrev()
{
	tooltip, [%A_thishotKey%] Run twin-panel file manager Total Commander
	SetTimer, TurnOffTooltip, -5000
	Run, C:\Program Files\totalcmd\TOTALCMD64.EXE
}

;PRINT SCREEN (38)
F_Button38()
{
Gui, Submit, NoHide
	Msgbox, 
				(
Run Printscreen application, by pressing {PrintScreen} key.
				)
Return
}

F_CB38_PrintScreen()
{
	global
	Gui, Submit, NoHide 
	Switch Check38 ;0, 1, -1
		{
			Case 0:
				Hotkey, PrintScreen,	F_prtscn, 		off 
				IniWrite, NO, VariousFunctions.ini, Menu memory, Print Screen 
		
			Case 1:
				Hotkey, PrintScreen, 	F_prtscn, 		on 
				IniWrite, YES, VariousFunctions.ini, Menu memory, Print Screen
		}
}


F_prtscn()
{ 
	Send, {Shift Down}{LWin Down}s{Shift Up}{LWin Up}
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - SECTION OF FUNCTIONS - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

F_IniRead()
{
	global	;assume-global mode of operation
	Temp1 := ""
	IniRead, ParenthesiIni, 		VariousFunctions.ini, 	Menu memory, 	Parenthesis, 			NO
	IniRead, BrowserIni, 			VariousFunctions.ini, 	Menu memory, 	Browser, 				NO
	IniRead, SetEnglishKeyboardIni, VariousFunctions.ini, 	Menu memory, 	Set English Keyboard, 	NO
	IniRead, AltGrIni, 				VariousFunctions.ini, 	Menu memory, 	AltGr, 					NO
	IniRead, WindowSwitcherIni, 	VariousFunctions.ini, 	Menu memory, 	Window Switcher, 		NO
	IniRead, PrintScreenIni, 		VariousFunctions.ini, 	Menu memory, 	Print Screen, 			NO
	IniRead, CapitalShiftIni, 		VariousFunctions.ini, 	Menu memory,	Capitalize Shift, 		NO
	IniRead, CapitalCapslIni, 		VariousFunctions.ini, 	Menu memory,	Capitalize Capslock, 	NO
	IniRead, MicrosoftWordIni, 		VariousFunctions.ini, 	Menu memory,	Microsoft `Word, 		NO
	IniRead, TotalCommanderIni, 	VariousFunctions.ini, 	Menu memory,	Total Commander, 		NO
	IniRead, PaintIni, 				VariousFunctions.ini, 	Menu memory,	Paint, 					NO
	IniRead, RebootIni, 			VariousFunctions.ini, 	Menu memory,	Reboot, 				NO
	IniRead, ShutdownIni, 			VariousFunctions.ini, 	Menu memory,	Shutdown, 				NO
	IniRead, TranspIni, 			VariousFunctions.ini, 	Menu memory,	Transparency, 			NO
	IniRead, F13Ini, 				VariousFunctions.ini, 	Menu memory,	F13, 					NO
	IniRead, F14Ini, 				VariousFunctions.ini, 	Menu memory,	F14, 					NO
	IniRead, F15Ini, 				VariousFunctions.ini, 	Menu memory,	F15, 					NO
	IniRead, TopIni,				VariousFunctions.ini, 	Menu memory,	Always on top,			NO
	IniRead, KeePassIni,			VariousFunctions.ini, 	Menu memory,	KeePass,				NO
	IniRead, HyperIni,				VariousFunctions.ini, 	Menu memory,	Hyperlink,				NO
	IniRead, HideIni,				VariousFunctions.ini, 	Menu memory,	Hidetext,				NO
	IniRead, ShowIni,				VariousFunctions.ini, 	Menu memory,	Showtext,				NO
	IniRead, AddTemplateIni,		VariousFunctions.ini, 	Menu memory,	Add Template,			NO
	IniRead, TemplateOffIni,		VariousFunctions.ini, 	Menu memory,	Template Off,			NO
	IniRead, StrikethroIni, 		VariousFunctions.ini,   Menu memory,	Strikethrough Text,    	NO
	IniRead, DeleteLineIni, 		VariousFunctions.ini,   Menu memory,	Delete Line,    		NO
	IniRead, AlignLeftIni, 			VariousFunctions.ini,   Menu memory,	Align Left,    			NO
	IniRead, ApplyStyleIni, 		VariousFunctions.ini,   Menu memory,	Apply Styles,    		NO
	IniRead, OpenPathIni, 			VariousFunctions.ini,   Menu memory,	Open and Show Path,    	NO
	IniRead, TableIni, 				VariousFunctions.ini,   Menu memory,	Table,    				NO
	IniRead, SuspendIni,			VariousFunctions.ini,	Menu memory,	Suspend,				NO
	IniRead, VolumeIni,				VariousFunctions.ini,	Menu memory,	Volume Up & Down,		NO
	IniRead, BroWinSwiIni,			VariousFunctions.ini,	Menu memory,	Browser Win Switcher,	NO
	IniRead, TranspMouIni, 			VariousFunctions.ini,	Menu memory,	Transparency Mouse,		NO
	IniRead, AutosaveIni, 			VariousFunctions.ini,	Menu memory,	Autosave,				NO
	IniRead, ENtoPLini, 			VariousFunctions.ini,	Menu memory,	ENtoPL,					NO
	IniRead, PLtoENini, 			VariousFunctions.ini,	Menu memory,	PLtoEN,					NO
	IniRead, AutomatePaintini, 		VariousFunctions.ini,	Menu memory,	AutomatePaint,			NO

	;[OpenTabs]
	IniRead, URL1  ,				VariousFunctions.ini,	OpenTabs,		Tab1,					Unknown link
	IniRead, URL2  ,				VariousFunctions.ini,	OpenTabs,		Tab2,					Unknown link
	IniRead, URL3  ,				VariousFunctions.ini,	OpenTabs,		Tab3,					Unknown link
	IniRead, URL4  ,				VariousFunctions.ini,	OpenTabs,		Tab4,					Unknown link
	IniRead, URL5  ,				VariousFunctions.ini,	OpenTabs,		Tab5,					Unknown link
	IniRead, URL6  ,				VariousFunctions.ini,	OpenTabs,		Tab6,					Unknown link
	IniRead, URL7  ,				VariousFunctions.ini,	OpenTabs,		Tab7,					Unknown link
	IniRead, URL8  ,				VariousFunctions.ini,	OpenTabs,		Tab8,					Unknown link
	IniRead, URL9  ,				VariousFunctions.ini,	OpenTabs,		Tab9,					Unknown link
	IniRead, URL10 ,				VariousFunctions.ini,	OpenTabs,		Tab10,					Unknown link
	Temp1 := URL1 . A_space . URL2 . A_space . URL3 . A_space . URL4 . A_space . URL5 . A_space . URL6 . A_space . URL7 . A_space . URL8 . A_space . URL9 . A_space . URL10
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - 

F_PrepareGuiElements()
{
	global	;assume-global mode of operation
	Gui Font, s9, Segoe UI

	Gui Add, Picture, 	x8 y8 w47 h47, C:\Repozytoria\companyTemplates\AutoHotKeyScripts\VariousFunctions\vh.ico
	Gui Add, Text, 		x96 y16 w91 h15 +0x200 				 , CHOOSE YOUR
	Gui Add, Text, 		x104 y32 w72 h17 +0x200 			 , FUNCTIONS
	Gui Add, Button, 	x8 y64 w220 h41 ClassButton gF_About1 , About

	Gui Add, GroupBox, 	x8 y104 w220 h303
	Gui Add, CheckBox, 	x24 y120 w170 h23 vCheck1 gF_CB1_AlwaysOnTop, Always on top
	Switch TopIni
	{
		Case "YES": GuiControl, , Check1, 1
		Case "NO":	GuiControl, , Check1, 0
	}

	Gui Add, Button, 	x200 y120 w22 h24 gF_Button1,	?

	Gui Add, CheckBox, x24 y152 w170 h23 vCheck2 gF_CB2_AutomatePaint, Automate Paint
	Switch AutomatePaintIni
	{
		Case "YES": GuiControl, , Check2, 1
		Case "NO":	GuiControl, , Check2, 0
	}

	Gui Add, Button, x200 y152 w22 h24 gF_Button2, ?

	Gui Add, CheckBox, x24 y184 w170 h23 vCheck3 gF_CB3_ChromeTabSwitcher, Chrome tab switcher
	Switch BroWinSwiIni
	{
		Case "YES": GuiControl, , Check3, 1
		Case "NO":	GuiControl, , Check3, 0
	}

	Gui Add, Button, x200 y184 w22 h24 gF_Button3, ?

	Gui Add, CheckBox, x24 y216 w170 h23 vCheck4 gF_CB4_OpenTabs, Open tabs in Chrome
	Switch BrowserIni
	{
		Case "YES": GuiControl, , Check4, 1
		Case "NO":	GuiControl, , Check4, 0
	}

	Gui Add, Button, x200 y216 w22 h24 gF_Button4, ?

	Gui Add, CheckBox, x24 y248 w170 h23 vCheck5 gF_CB5_ParenthesisWatcher, Parenthesis watcher
	Switch ParenthesisIni
	{
		Case "YES": GuiControl, , Check5, 1
		Case "NO":	GuiControl, , Check5, 0
	}

	Gui Add, Button, x200 y248 w22 h24 gF_Button5, ?

	Gui Add, CheckBox, x24 y280 w170 h23 vCheck6 gF_CB6_USKeyboard, US keybord
	Switch SetEnglishKeyboardIni
	{
		Case "YES": GuiControl, , Check6, 1
		Case "NO":	GuiControl, , Check6, 0
	}

	Gui Add, Button, x200 y280 w22 h24 gF_Button6, ?

	Gui Add, CheckBox, x24 y312 w170 h23 vCheck7 gF_CB7_RightClick, Right-click context menu
	Switch AltGrIni
	{
		Case "YES": GuiControl, , Check7, 1
		Case "NO":	GuiControl, , Check7, 0
	}

	Gui Add, Button, x200 y312 w22 h24 gF_Button7, ?

	Gui Add, CheckBox, x24 y344 w170 h23 vCheck8 gF_CB8_VolumeUpDown, Volume Up and Down
	Switch VolumeIni
	{
		Case "YES": GuiControl, , Check8, 1
		Case "NO":	GuiControl, , Check8, 0
	}

	Gui Add, Button, x200 y344 w22 h24 gF_Button8, ?

	Gui Add, CheckBox, x24 y376 w170 h23 vCheck9 gF_CB9_WindowSwitcher, Window switcher
	Switch WindowSwitcherIni
	{
		Case "YES": GuiControl, , Check9, 1
		Case "NO":	GuiControl, , Check9, 0
	}

	Gui Add, Button, x200 y376 w22 h24 gF_Button9, ?

	Gui Add, GroupBox, x8 y416 w219 h89, Capitalization switcher

	Gui Add, CheckBox, x24 y440 w169 h23 vCheck10 gF_CB10_CAPSLOCK, Capslock
	Switch CapitalCapsIni
	{
		Case "YES": GuiControl, , Check10, 1
		Case "NO":	GuiControl, , Check10, 0
	}

	Gui Add, Button, x200 y440 w22 h24 gF_Button10, ?

	Gui Add, CheckBox, x24 y472 w169 h23 vCheck11 gF_CB11_SHIFTF3, Shift + F3
	Switch CapitalShiftIni
	{
		Case "YES": GuiControl, , Check11, 1
		Case "NO":	GuiControl, , Check11, 0
	}

	Gui Add, Button, x200 y472 w22 h24 gF_Button11, ?

	Gui Add, GroupBox, x240 y32 w177 h128, Foot switch

	Gui Add, CheckBox, x256 y56 w129 h23 vCheck12 gF_CB12_FootswitchF13, F13
	Switch F13Ini
	{
		Case "YES": GuiControl, , Check12, 1
		Case "NO":	GuiControl, , Check12, 0
	}

	Gui Add, Button, x392 y56 w22 h24 gF_Button12, ?

	Gui Add, CheckBox, x256 y88 w129 h23 vCheck13 gF_CB13_FootswitchF14, F14
	Switch F14Ini
	{
		Case "YES": GuiControl, , Check13, 1
		Case "NO":	GuiControl, , Check13, 0
	}

	Gui Add, Button, x392 y88 w22 h24 gF_Button13, ?

	Gui Add, CheckBox, x256 y120 w129 h23 vCheck14 gF_CB14_FootswitchF15, F15
	Switch F15Ini
	{
		Case "YES": GuiControl, , Check14, 1
		Case "NO":	GuiControl, , Check14, 0
	}

	Gui Add, Button, x392 y120 w22 h24 gF_Button14, ?

	Gui Add, GroupBox, x240 y168 w178 h100, Google translate

	Gui Add, CheckBox, x256 y192 w129 h23 vCheck15 gF_CB15_TranslateEN_PL, English → Polish
	Switch ENtoPLIni
	{
		Case "YES": GuiControl, , Check15, 1
		Case "NO":	GuiControl, , Check15, 0
	}

	Gui Add, Button, x392 y192 w22 h24 gF_Button15, ?

	Gui Add, CheckBox, x256 y224 w129 h23 vCheck16 gF_CB16_TranslatePL_EN, Polish → English
	Switch PLtoENIni
	{
		Case "YES": GuiControl, , Check16, 1
		Case "NO":	GuiControl, , Check16, 0
	} 

	Gui Add, Button, x392 y224 w22 h24 gF_Button16, ?

	Gui Add, GroupBox, x240 y280 w177 h127, Power PC

	Gui Add, CheckBox, x256 y312 w129 h23 vCheck17 gF_CB17_Suspend, Suspend
	Switch SuspendIni
	{
		Case "YES": GuiControl, , Check17, 1
		Case "NO":	GuiControl, , Check17, 0
	} 

	Gui Add, Button, x392 y312 w22 h24 gF_Button17, ?

	Gui Add, CheckBox, x256 y344 w129 h23 vCheck18 gF_CB18_Reboot, Reboot
	Switch RebootIni
	{
		Case "YES": GuiControl, , Check18, 1
		Case "NO":	GuiControl, , Check18, 0
	} 

	Gui Add, Button, x392 y344 w22 h24 gF_Button18, ?

	Gui Add, CheckBox, x256 y376 w129 h23 vCheck19 gF_CB19_Shutdown, Shutdown
	Switch ShutdownIni
	{
		Case "YES": GuiControl, , Check19, 1
		Case "NO":	GuiControl, , Check19, 0
	} 

	Gui Add, Button, x392 y376 w22 h24 gF_Button19, ?

	Gui Add, GroupBox, x240 y416 w178 h90, Transparency switcher

	Gui Add, CheckBox, x256 y440 w131 h23 vCheck20 gF_CB20_TransparencyMouse, Mouse
	Switch TransparencyMouIni
	{
		Case "YES": GuiControl, , Check20, 1
		Case "NO":	GuiControl, , Check20, 0
	} 

	Gui Add, Button, x392 y440 w22 h24 gF_Button20, ?

	Gui Add, CheckBox, x256 y472 w130 h23 vCheck21 gF_CB21_TransparencyKeys, Keys
	Switch TransparencyIni
	{
		Case "YES": GuiControl, , Check21, 1
		Case "NO":	GuiControl, , Check21, 0
	} 

	Gui Add, Button, x392 y472 w22 h24 gF_Button21, ?

	Gui Add, GroupBox, x432 y32 w188 h474, Functions in MS WORD

	Gui Add, CheckBox, x440 y56 w138 h23 vCheck22 gF_CB22_AlignLeft, Align Left
	Switch AlignLeftIni
	{
		Case "YES": GuiControl, , Check22, 1
		Case "NO":	GuiControl, , Check22, 0
	}

	Gui Add, Button, x584 y56 w22 h24 gF_Button22, ?  

	Gui Add, CheckBox, x440 y88 w138 h23 vCheck23 gF_CB23_ApplyStyles, Apply styles
	Switch ApplyStyleIni
	{
		Case "YES": GuiControl, , Check23, 1
		Case "NO":	GuiControl, , Check23, 0
	} 

	Gui Add, Button, x584 y88 w22 h24 gF_Button23, ?

	Gui Add, CheckBox, x440 y120 w138 h23 vCheck24 gF_CB24_Autosave, Autosave
	Switch AutosaveIni
	{
		Case "YES": GuiControl, , Check24, 1
		Case "NO":	GuiControl, , Check24, 0
	} 

	Gui Add, Button, x584 y120 w22 h24 gF_Button24, ?

	Gui Add, CheckBox, x440 y152 w138 h23 vCheck25 gF_CB25_DeleteLine, Delete Line
	Switch DeleteLineIni
	{
		Case "YES": GuiControl, , Check25, 1
		Case "NO":	GuiControl, , Check25, 0
	} 

	Gui Add, Button, x584 y152 w22 h24 gF_Button25, ?

	Gui Add, Text, x440 y184 w138 h23 +0x200  , Hidden text

	Gui Add, CheckBox, x480 y216 w99 h22 vCheck26 gF_CB26_Hide, Hide
	Switch HideIni
	{
		Case "YES": GuiControl, , Check26, 1
		Case "NO":	GuiControl, , Check26, 0
	} 

	Gui Add, Button, x584 y216 w22 h24 gF_Button26, ?

	Gui Add, CheckBox, x480 y248 w98 h22 vCheck27 gF_CB27_Show, Show
	Switch ShowIni
	{
		Case "YES": GuiControl, , Check27, 1
		Case "NO":	GuiControl, , Check27, 0
	}  

	Gui Add, Button, x584 y248 w22 h24 gF_Button27, ?

	Gui Add, CheckBox, x440 y280 w138 h23 vCheck28 gF_CB28_Hyperlink, Hyperlink
	Switch HyperIni
	{
		Case "YES": GuiControl, , Check28, 1
		Case "NO":	GuiControl, , Check28, 0
	} 

	Gui Add, Button, x584 y280 w22 h24 gF_Button28, ?

	Gui Add, CheckBox, x440 y312 w138 h23 vCheck29 gF_CB29_OpenAndShowPath, Open and show path
	Switch OpenPathIni
	{
		Case "YES": GuiControl, , Check29, 1
		Case "NO":	GuiControl, , Check29, 0
	} 

	Gui Add, Button, x584 y312 w22 h24 gF_Button29, ?

	Gui Add, CheckBox, x440 y344 w138 h23 vCheck30 gF_CB30_StrikethroughText, Strikethrough text
	Switch OpenPathIni
	{
		Case "YES": GuiControl, , Check30, 1
		Case "NO":	GuiControl, , Check30, 0
	} 

	Gui Add, Button, x584 y344 w22 h24 gF_Button30, ?

	Gui Add, CheckBox, x440 y376 w138 h23 vCheck31 gF_CB31_Table, Table
	Switch TableIni
	{
		Case "YES": GuiControl, , Check31, 1
		Case "NO":	GuiControl, , Check31, 0
	} 

	Gui Add, Button, x584 y376 w22 h24 gF_Button31, ?

	Gui Add, Text, x440 y408 w138 h23 +0x200  , Template

	Gui Add, CheckBox, x480 y440 w100 h24 vCheck32 gF_CB32_AddTemplate, Add template
	Switch AddTemplateIni
	{
		Case "YES": GuiControl, , Check32, 1
		Case "NO":	GuiControl, , Check32, 0
	} 

	Gui Add, Button, x584 y440 w22 h24 gF_Button32, ?

	Gui Add, CheckBox, x480 y472 w102 h24 vCheck33 gF_CB33_TemplateOff, Template Off
	Switch TemplateOffIni
	{
		Case "YES": GuiControl, , Check33, 1
		Case "NO":	GuiControl, , Check33, 0
	} 

	Gui Add, Button, x584 y472 w22 h24 gF_Button33, ?    

	Gui Add, GroupBox, x632 y32 w188 h183, Run . . .

	Gui Add, CheckBox, x648 y56 w138 h23 vCheck34 gF_CB34_KeePass, KeePass
	Switch KeePassIni
	{
		Case "YES": GuiControl, , Check34, 1
		Case "NO":	GuiControl, , Check34, 0
	} 

	Gui Add, Button, x792 y56 w22 h24 gF_Button34, ?

	Gui Add, CheckBox, x648 y88 w138 h23 vCheck35 gF_CB35_MSWord, MS Word
	Switch MicrosoftWordIni
	{
		Case "YES": GuiControl, , Check35, 1
		Case "NO":	GuiControl, , Check35, 0
	} 

	Gui Add, Button, x792 y88 w22 h24 gF_Button35, ?

	Gui Add, CheckBox, x648 y120 w138 h24 vCheck36 gF_CB36_Paint, Paint
	Switch PaintIni
	{
		Case "YES": GuiControl, , Check36, 1
		Case "NO":	GuiControl, , Check36, 0
	} 

	Gui Add, Button, x792 y120 w22 h24 gF_Button36, ?

	Gui Add, CheckBox, x648 y152 w138 h24 vCheck37 gF_CB37_TotalCommander, Total Commander
	Switch TotalCommanderIni
	{
		Case "YES": GuiControl, , Check37, 1
		Case "NO":	GuiControl, , Check37, 0
	} 

	Gui Add, Button, x792 y152 w22 h24 gF_Button37, ?

	Gui Add, CheckBox, x648 y184 w138 h23 vCheck38 gF_CB38_PrintScreen, Print Screen
	Switch PrintScreenIni
	{
		Case "YES": GuiControl, , Check38, 1
		Case "NO":	GuiControl, , Check38, 0
	} 

	Gui Add, Button, x792 y184 w22 h24 gF_Button38, ?

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - 

F_PrepareOpenTabsGui()
{
	global	;assume-global mode of operation
	Gui Font, s9, Segoe UI
	Gui OpenTabs: Add, Text, x16 y8 w287 h27 +0x200, Type up to 10 links to open in Chrome and click "Save".
			Gui OpenTabs: Add, Edit, x8 y48 w302 h21 vTab1
				GuiControl, OpenTabs:, Tab1, % URL1
			Gui OpenTabs: Add, Edit, x8 y80	 w302 h21 vTab2
				GuiControl, OpenTabs:, Tab2, % URL2
			Gui OpenTabs: Add, Edit, x8 y112 w302 h21 vTab3
				GuiControl, OpenTabs:, Tab3, % URL3
			Gui OpenTabs: Add, Edit, x8 y144 w302 h21 vTab4
				GuiControl, OpenTabs:, Tab4, % URL4
			Gui OpenTabs: Add, Edit, x8 y176 w302 h21 vTab5
				GuiControl, OpenTabs:, Tab5, % URL5
			Gui OpenTabs: Add, Edit, x8 y208 w302 h21 vTab6
				GuiControl, OpenTabs:, Tab6, % URL6
			Gui OpenTabs: Add, Edit, x8 y240 w302 h21 vTab7
				GuiControl, OpenTabs:, Tab7, % URL7
			Gui OpenTabs: Add, Edit, x8 y272 w302 h21 vTab8
				GuiControl, OpenTabs:, Tab8, % URL8
			Gui OpenTabs: Add, Edit, x8 y304 w302 h21 vTab9
				GuiControl, OpenTabs:, Tab9, % URL9
			Gui OpenTabs: Add, Edit, x8 y336 w302 h21 vTab10
				GuiControl, OpenTabs:, Tab10, % URL10
			Gui OpenTabs: Add, Button, x232 y368 w80 h23 gF_SAVE, &SAVE	
}


TurnOffTooltip()
{
	Tooltip,
	return
}
