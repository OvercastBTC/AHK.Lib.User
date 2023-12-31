﻿/************************************************************************
 * @description 
 * @file AHK Script.v2.ahk
 * @author 
 * @date 2023/12/28
 * @version 0.0.0
 ***********************************************************************/

#Requires AutoHotkey v2+
#Include <Directives\__AE.v2>
; --------------------------------------------------------------------------------
TraySetIcon("shell32.dll","16", true) ; this changes the icon into a little laptop thing.
; --------------------------------------------------------------------------------
myScript := A_TrayMenu
; TrayMenu := MenuBar()
myScript.Delete() ; V1toV2: not 100% replacement of NoStandard, Only if NoStandard is used at the beginning
myScript.Add("Made by OvercastBTC",madeBy)
myScript.Add()
; Tray.Add("Run at startup", "runAtStartup")
; Tray.ToggleCheck(fileExist(startup_shortcut) ? "check" : "unCheck", Run at startup ; update the tray menu status on startup
; TODO Finish updating Run at Startup
; Tray.Add("Presentation mode {Win+Shift+P}", togglePresentationMode) ; => does not exist
; Tray.Add("Keyboard shortcuts {Ctrl+Shift+Alt+\}", "viewKeyboardShortcuts")
; Tray.Add("Open file location", "openFileLocation")
; Tray.Add()
; Tray.Add("Run GUI_FE", "GUIFE")
myScript.Add('Run WindowsListMenu.ahk', WindowListMenu)
myScript.Add()
myScript.Add('Run GUI_ListofFiles.ahk', GUI_ListofFiles)
myScript.Add("Run WindowProbe.ahk", WindowProbe)
; Tray.Add("Run Windows_Data_Types_offline.ahk", Windows_Data_Types_offline)
myScript.Add()
; Tray.Add("View in GitHub", "viewInGitHub")
; Tray.Add("See AutoHotKey documentation", "viewAHKDoc") ; => does not exist
; Tray.Add()
myScript.AddStandard()
; Tray.Show
; --------------------------------------------------------------------------------
madeBy(madeBy,*){
}
WindowListMenu(*){
	Run("WindowListMenu.ahk")
}
; --------------------------------------------------------------------------------
; #Include <Abstractions\Script>
; Script.startup.CheckStartupStatus()
#Include <Common\Common_Include>
; #Include <WINDOWS.v2>
#Include <Common\Misc Scripts.V2>
#Include <Tools\Info>
#Include <System\Web>
#Include <Tools\CleanInputBox>
#Include <App\Git>
#Include <System\UIA>
#Include <GetNearestMonitorInfo().v2>
#Include <Utils\ClipSend>
; #Include <WindowSpyDpi>
; ---------------------------------------------------------------------------
SetNumLockState("AlwaysOn")
SetCapsLockState("AlwaysOff")
SetScrollLockState("AlwaysOff")
; --------------------------------------------------------------------------------
#HotIf WinActive("Chrome River - Google Chrome")
*^s::SaveCR()
; ---------------------------------------------------------------------------
SaveCR(){
	expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
	; Sleep(100)
	expRpt.WaitElement({Type: '50000 (Button)', Name: "Save", LocalizedType: "button", AutomationId: "save-btn"}, 10000).Highlight(100).Invoke()
}

#HotIf
~n::
{
	If (A_PriorKey = '``'){
		Sleep(100)
		Send('{Right}')
	}
}
#HotIf WinActive("ahk_exe hznHorizon.exe")

^+Down::
{
	bak := ClipboardAll()
	d:=0, A_Clipboard := '', d := 15, text := '', aText := []
	; if A_Clipboard != '' {
	; 	loop {
	; 		Sleep(d)
	; 	} until A_Clipboard := ''
	; }
	cutline()
	cutline()
	cutline() {
		Send('{End}{Shift Down}{Home}{Down}{Shift Up}')
		Send('^x')
		text := A_Clipboard
		Sleep(d)
		aText.Push(text)
		Send('{Home}{Shift Down}{End}{Down}{Shift Up}')
		Send('{Del}')
		return aText
	}
	Send(aText[2])
	Send(aText[1])
	; Infos(atext[1] '`n' aText[2])
	return
	text := ''
	loop text.Length {
		Sleep(d)
	} until text = ''
	text .= (aText[2] '`n' aText[1])
	loop text.Length {
		Sleep(d)
	} until text != ''
	; Send('^x')
	; Sleep(d)
	Send('{Del 2}')
	Sleep(d)
	send('{End}')
	Sleep(d)
	Send('{Enter}')
	Sleep(d)
	send('^v')
	Sleep(100)
	A_Clipboard := bak
}
^+Up::
{
	bak := ClipboardAll()
	d:=0, A_Clipboard := '', d := 15, text := ''
	Send('{Home}+{End}')
	Sleep(d)
	Send('^x')
	Sleep(d)
	Send('{Del}')
	Sleep(d)
	send('{Home}{Up}')
	Sleep(d)
	Send('{Enter}')
	Sleep(d)
	send('^v')
	Sleep(100)
	A_Clipboard := bak
}
^+d::
{
	bak := ClipboardAll()
	d:=0,text := '', txt := '', A_Clipboard := '', d := 15
	; Sleep(d)
	Send('{Home}+{End}')
	; @step: Copy
	; SndMsgCopy()
	; Sleep(d)
	; text := A_Clipboard
	; Sleep(d)
	; @step: Cut
	SndMsgCut()
	; SendMessage('0x300',0,0,ControlGetFocus('A'),'A') ; Send('^x')
	; Send('{End}{Enter}')
	; @step Paste
	; SendMessage('0x302',0,0,ControlGetFocus('A'),'A')
	Send('^v')
	Sleep(d)
	Send('{Enter}')
	Sleep(d)
	Send('^v')
	Sleep(d)
	A_Clipboard := bak
}
#HotIf
;! ---------------------------------------------------------------------------
;! ---------------------------------------------------------------------------
;! ---------------------------------------------------------------------------
; ---------------------------------------------------------------------------
SndMsgCopy(*) {
	static WM_COPY := 0x0301, fCtl := 0
	try fCtl := ControlGetFocus('A')
	(fCtl = 0) ? SendEvent('^c') : SendMessage(WM_COPY,0,0,fCtl,'A')
}
; ---------------------------------------------------------------------------
SndMsgCut(*) {
	static WM_CUT := 0x0300, fCtl := 0
	try fCtl := ControlGetFocus('A')
	(fCtl = 0) ? SendEvent('^x') : SendMessage(WM_CUT,0,0,fCtl,'A')
}
; ---------------------------------------------------------------------------
SndMsgPaste(*) {
	static WM_PASTE := 0x0302, fCtl := 0
	try fCtl := ControlGetFocus('A')
	(fCtl = 0) ? SendEvent('^v') : SendMessage(WM_PASTE,0,0,fCtl,'A')
}
which_brackets(str, lChar, rChar, MapObj) {
	bkts := []
	for k, v in MapObj {
		If (str ~= '\' k){
			lChar := k, rChar := v
		}
	}
	bkts.SafePush(lChar)
	bkts.SafePush(rChar)
	; Infos([brackets]`n' 'str: ' bkts.str := str '`n' 'lChar: ' bkts.lChar := lChar '`n' 'rChar: ' bkts.rChar := rChar)
	return (bkts, {lChar:lChar, rChar:rChar})
}
; split_text(text) => (stext := [], stext := StrSplit(text))
; ; ---------------------------------------------------------------------------
; do_brackets_exist(text, lChar, rChar) => ((lChar.length > 0) ? (((text ~= ('\' lChar) || (('\' lChar) && ('\' rChar)))) ? true : false) : false)
; ; ---------------------------------------------------------------------------
; brackets_at_ends(text, lChar, rChar){
; 	stext := '', stext := split_text(text), ltext := stext.Length
; 	; ---------------------------------------------------------------------------
; 	(bL := (stext[1] ~= '\' lChar) 		? true : false)
; 	(bR := (stext[ltext] ~= '\' rChar) 	? true : false)
; 	; ---------------------------------------------------------------------------
; 	(bL = 1 ? 'true' : 'false'), (bR = 1 ? 'true' : 'false')
; 	; ---------------------------------------------------------------------------
; 	; Infos('true: ' true '`n' 'false: ' false '`n' 'bL: ' bL ' : ' stext[1] '`n' 'bR: ' bR ' : ' stext[ltext])
; 	return ((bL || bR = true) ? true : false)
; }
; ---------------------------------------------------------------------------
clip_empty(mode := false){
	if A_Clipboard = '' {
		while (A_Clipboard = '') {
			Sleep(10)
		}
	} else {
		while !(A_Clipboard = '') {
			Sleep(10)
		}
	}
}
; ---------------------------------------------------------------------------
#HotIf !WinActive(" - Visual Studio Code")
/************************************************************************
 * function ...: Mirror the VS Code hotkeys for the below params
 * @file AHK Script.v2.ahk
 * @author OvercastBTC
 * @date 2023/12/28
 * @version 0.5.0
 * @example By using the * => the key is intercepted and not sent to the window/control
 * 		Params:
 * 			1. ''
 * 			2. ""
 * 			3. ()
 * 			4. {}
 * 			5. []
 ***********************************************************************/
*'::
*"::
*(::
*{::
*[::
{
	CoordMode('Caret','Window')
	Infos.DestroyAll()
	_AE_bInpt_sLvl(1)
	A_Clipboard := '', text := ''
	SetTimer(clip_empty, -500)
	t1 := '', t2 := '', t3 := '', t4 := '', t5 := '', t6 := '', t7 := '',
	c1 := '', c2 := '', c3 := '', c4 := '', c5 := '', c6 := '', c7 := ''
	bLeft := '', bRight := '', lChar := '', rChar := '',
	; s := '', c := '',
	bK := '', bV := '',
	eK := '', eV := '', 
	x  := 0, y  := 0,	xC := 0, yC := 0, w := 0, p := ''
	control  := unset, fCtl := unset, WinA := unset, ClassNN := unset
	Selected := ''
	bChar := '', bChars := '', bV1 := '', bV2 := '', b1 := '', b2 := ''
	eChar := '', eChars := '', eV1 := '', eV2 := '', e1 := '', e2 := ''
	cLen := 0, sLen := 0
	; toggle := 0
	; ---------------------------------------------------------------------------
	at := [], bkts := [], eChr := [], bChr := [], mbChr := [], meChr := []
	; ---------------------------------------------------------------------------
	eMap := Map('"','"', "'", "'", '[', ']', '{', '}', '(', ')')
	; ---------------------------------------------------------------------------
	; @step ...: Try and get the focused control, caret and mouse positions
	; @step ...: Try and get caret position
	; @step ...: Try and get mouse position (includes window handle, and control handle)
	; ---------------------------------------------------------------------------
	try if WinActive('ahk_exe hznHorizon.exe'){
		static fCtl := ControlGetFocus(), ClassNN := ControlGetClassNN(fCtl)
	}
	teams()
	return
	try (CaretGetPos(&xC, &yC))
	try MouseGetPos(&x, &y, &w, &control)
	; Infos(
	; 	'xC[' xC '] yC[' yC ']'
	; 	'`n'
	; 	'x[' x '] y[' y ']'
	; 	'`n'
	; 	'control: ' ClassNN ' Title: ' WinT 
	; )
	; ---------------------------------------------------------------------------
	; @step: get the key pressed, removing the hotkey modifier, and trim whitespaces
	; ---------------------------------------------------------------------------
	str := StrSplit(A_ThisHotkey,'*', '*', 1), str := Trim(str[1])
	; ---------------------------------------------------------------------------
	; @info ...: brackets identified => b := Array() 
	; ---------------------------------------------------------------------------
	bkts := which_brackets(str, lChar,rChar, eMap)
	; ---------------------------------------------------------------------------
	bLeft := bkts.lChar, bRight := bkts.rChar
	; ---------------------------------------------------------------------------
	; @step: Copy the selected text into the clipboard for evaluation
	; ---------------------------------------------------------------------------
	; WinActive('ahk_exe hznHorizon.exe') ? Send('+{Left}') : Click(2)
	c1 := A_Clipboard.length
	SndMsgCopy()
	; c2 := A_Clipboard.length
	; ---------------------------------------------------------------------------
	; @step: store the clipboard into the variable text, see if text is empty
	; @info: don't technically need to do that to see if the text is empty
	; @info: however, two birds with one stone => set text variable
	; ---------------------------------------------------------------------------
	text := A_Clipboard, t1 := text
	c2 := t1.length
	; ---------------------------------------------------------------------------
	; @info ...: set a short sleep cycle until the text = clipboard
	; ---------------------------------------------------------------------------
	Loop {
		Sleep(15)
	} until A_Clipboard = text
	; ---------------------------------------------------------------------------
	; @step ...: Set cLen (clipboard length, or initial cLen length)
	; @step ...: Set sLen (string length, or initial sLen length)
	; @why  ...: Used to select the whole word or sentance for toggling backets.
	; ---------------------------------------------------------------------------
	cLen := text.length
	; ---------------------------------------------------------------------------
	t2 := text, c3 := t2.length
	;! ---------------------------------------------------------------------------
	; @step ...: get stats from text
	; @why  ...: Can be used to determine the location of the text (line, total lines, etc.)
	;! ---------------------------------------------------------------------------
	try stats := AE_GetStats()
	; stats := AE_GetStats()
	; ---------------------------------------------------------------------------
	; (text.length = 0) ? l := 1 : l := 0
	; needle:= "D)\s+$"
	; ---------------------------------------------------------------------------
	; @step ...: Is there a space at the end?
	; @step ...: If so, remove it and store it in s (s := ' ', or s := A_Space)
	; @why  ...: If its a single line, no need to re-add
	; @why  ...: If it's part of a sentance, need to add it back at the end
	;! Moved to next section
	; ---------------------------------------------------------------------------
	; needle:= ('\s+$')
	; spaceNeedle:= ('\s+\Q([^\s]+)\E\s+')
	; spacereplace := "[$1]"
	; spacereplace := ''
	; if ((text ~= spaceNeedle) != 0) {
	; 	text := RegExReplace(text, spaceNeedle, spacereplace)
	; 	s := ' '
	; } else {
	; 	text := text
	; 	s := ''
	; }
	t3 := text, c4 := t3.length
	; ((text ~= needle) != 0) ? (text := RegExReplace(text, needle, ''), s := '') : s := ' '
	; ---------------------------------------------------------------------------
	; @step ...: Is there a dash, dash space, or space (or combo of) at the beginning?
	; @step ...: If so, remove it/them and store it in bChars 
	; @step ...: Is there a space, comma, question mark, semi-colon, colon, (?more?) (or combo of) at the end?
	; //@step ...: If so, remove it and store it in s (s := ' ', or s := A_Space)
	; @step ...: If so, remove it/them and store it in eChars 
	; @why  ...: If its a single word on the line, no need to re-add ; fix [identify if it's a single word on the line]
	; @why  ...: If it's part of a sentance, need to add it back at the end
	; ---------------------------------------------------------------------------
	bChrNeedle := '^[ -]+'
	bChrReplace := '$1'
	; ---------------------------------------------------------------------------
	eChrNeedle := '[,.\s?:;]+$'
	eChrReplace := '$1'
	; ---------------------------------------------------------------------------
	; cp := [], bcp := [], cpt := ''
	b1 := bChars
	if (text ~= bChrNeedle) {
		RegExMatch(text, bChrNeedle,&mbChr)
		; text := RegExReplace(text, bChrNeedle,'')
		text := RegExReplace(text, bChrNeedle,bChrReplace)
		try for each, bV1 in mbChr {
			bChr.SafePush(bV1)
			cpt .= bV1 '`n' ;? testing purposes only
		}
		try for each, bV2 in bChr {
			bChars .= bV2
		}
	}
	b2 := bChars
	t4 := text, c5 := t4.length
	; ---------------------------------------------------------------------------
	e1 := eChars
	if (text ~= eChrNeedle) {
		RegExMatch(text, eChrNeedle,&meChr)
		; text := RegExReplace(text, eChrNeedle,'')
		text := RegExReplace(text, eChrNeedle,eChrReplace)
		try for each, eV1 in meChr {
			eChr.SafePush(eV1)
			cpt .= eV1 '`n'
		}
		try for each, eV2 in eChr {
			eChars .= eV2
		}
	}
	e2 := eChars
	t5 := text, c6 := t5.length
	; ---------------------------------------------------------------------------
	; @step if text.length != 0 => do brackets exist? If so, remove them, if not, add them
	; ---------------------------------------------------------------------------
	bK := bLeft, bV := bRight
	bktPattern := "\Q" bK "\E([^" bV "]+)\Q" bV "\E"
	bktRplc := '$1'
	((text ~= eMap.Get(bLeft)) ? (text := RegExReplace(text, bktPattern, bktRplc)) : (text := bK text bV ))
	; ---------------------------------------------------------------------------
	; @step Add the bChars and eChars back to the text
	; ---------------------------------------------------------------------------
	text := bChars text eChars
	; (text ~= eMap.Get(bLeft)) ? (text := RegExReplace(text, pattern, replacement)
	; 							; , text := text c s
	; 						) 
	; 	: (text := k text v c s)
	; ---------------------------------------------------------------------------
	A_Clipboard := text
	t6 := A_Clipboard, c7 := t6.length
	WinActive('ahk_exe hznHorizon.exe') ? SndMsgPaste() : Send('^v')
	Infos(
		'c1[' c1 ']' 
		'`n'
		'c2[' c2 ']' ' t1: ' t1
		'`n'
		'c3[' c3 ']' ' t2: ' t2
		'`n'
		'c4[' c4 ']' ' t3: ' t3
		'`n'
		'b1[' b1 ']'
		'`n'
		'b2[' b2 ']'
		'`n'
		'e1[' e1 ']'
		'`n'
		'e2[' e2 ']'
		'`n'
		'c5[' c5 ']' ' t4: ' t4
		'`n'
		'c6[' c6 ']' ' t5: ' t5
		'`n'
		'c7[' c7 ']' ' t6: ' t6
		'`n'
		't7: ' t7
		'RegEx:'
		'`n'
		cpt
	)
	; Sleep(100)
	
	; try Send('^v')
	; A_Clipboard := text
	; Send('^v')

	; ---------------------------------------------------------------------------
	; SendEvent(l '}')
	Send('+{Left ' text.length '}')
	; WinActive('ahk_exe hznHorizon.exe') ? SendEvent('^{left}+^{right}') : SendEvent('^+{left 3}')
	; Send('^{left}+^{right}')
	; _AE_bInpt_sLvl(0)
	; Run(A_ScriptName)
}
teams(*){
	x := 0, y := 0, w := 0, control := '', hwnd := '', WinA := 0, WinC := '', WinT := '', fCtl := 0
	; el := UIA.ElementExist({Type: '50004 (Edit)', Name: "Type a message"})
	WinA := WinActive('A'), text := '', aedit := []
	WinC := WinGetClass(WinA)
	WinT := WinGetTitle(WinA)
	; fCtl := ControlGetFocus(WinA)
	; ctls := WinGetControls(WinA)
	; for each, value in ctls {
	; 	text .= value '(' ControlGetHwnd(value) ')' '`n'
	; }
	MouseGetPos(&x, &y, &w, &control, 2)
	classNN := ControlGetClassNN(control)
	text .= 'Col: ' EditGetCurrentCol(control, WinA) '`n'
	text .= 'Lin: ' EditGetCurrentLine(control, WinA) '`n'
	; text .= EditGetLine(1,control, WinA) '`n'
	text .= 'Lines: ' EditGetLineCount(control, WinA) '`n'
	try text .= 'eTxt: ' sel := EditGetSelectedText(control, WinA) '`n'
	try text .= 'cTxt: ' sel := ControlGetText(control, WinA) '`n'

	el := UIA.ElementFromHandle('A')
	edit := el.WaitElement({Type: '50004 (Edit)', Name: "Type a message"})
	; aedit := edit.Children
	; for each, value in aedit {
	; 	text .= value '`n'
	; }
	Infos(
		'el: '
		'`n'
		; edit.value
		; '`n'
		; edit.Children
		'`n'
		'edit.classname: ' edit.ClassName
		'`n'
		'x[' x ']' ' ' 'y[' y ']'
		'`n'
		'w(' w ')'
		'`n'
		'control: ' ClassNN ' (' control ')'
		'`n'

		'`n'
		'WinA: ' WinA
		'`n'
		'WinC: ' WinC
		'`n'
		'WinT: ' WinT
		'`n'
		; 'fCtl: ' fCtl
		'Controls List:`n' text
	)
	; Type: 50033 (Pane) Name: "Keatts, Terry | Microsoft Teams classic, Main Window" LocalizedType: "pane" ClassName: "Chrome_WidgetWin_1"
	; 1: Type: 50030 (Document) Name: "Keatts, Terry | Microsoft Teams classic" Value: "https://teams.microsoft.com/_#/conversations/19:928764bb-35cc-44cc-9b42-7c8b166f9a8e_d5552339-b24f-492a-89b2-2a6b9889197f@unq.gbl.spaces?groupId=554f7f6a-aad5-49e4-b796-09580ac5f32c&ctx=chat" LocalizedType: "document" AutomationId: "27687008" ClassName: "Chrome_RenderWidgetHostHWND"
	; 1,1: Type: 50006 (Image) LocalizedType: "image"
	; 1,2: Type: 50000 (Button) Name: "Go back, hover to see history" LocalizedType: "button"
	; 1,3: Type: 50000 (Button) Name: "Forward" LocalizedType: "button"
	; 1,4: Type: 50026 (Group) LocalizedType: "search"
	; 1,4,1: Type: 50003 (ComboBox) LocalizedType: "combo box" AutomationId: "controlbox-input-group"
	; 1,4,1,1: Type: 50004 (Edit) Name: "Search" LocalizedType: "edit" AutomationId: "control-input"
	; 1,5: Type: 50011 (MenuItem) Name: "Settings and more" LocalizedType: "menu item" AutomationId: "settings-menu-button"
	; 1,6: Type: 50011 (MenuItem) Name: "Profile." LocalizedType: "menu item" AutomationId: "personDropdown"
	; 1,7: Type: 50026 (Group) Name: "App bar, use up or down arrow keys to navigate through the controls" LocalizedType: "navigation" AutomationId: "teams-app-bar"
	; 1,7,1: Type: 50008 (List) LocalizedType: "list"
	; 1,7,1,1: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,1,1: Type: 50000 (Button) Name: "Activity Toolbar" LocalizedType: "button" AutomationId: "app-bar-14d6962d-6eeb-4f48-8890-de55454bb136"
	; 1,7,1,2: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,2,1: Type: 50000 (Button) Name: "Chat Toolbar more options" LocalizedType: "button" AutomationId: "app-bar-86fcd49b-61a2-4701-b771-54728cd291fb"
	; 1,7,1,3: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,3,1: Type: 50000 (Button) Name: "Teams Toolbar" LocalizedType: "button" AutomationId: "app-bar-2a84919f-59d8-4441-a975-2a8c2643b741"
	; 1,7,1,4: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,4,1: Type: 50000 (Button) Name: "Calendar Toolbar" LocalizedType: "button" AutomationId: "app-bar-ef56c0de-36fc-4ef8-b417-3d82ba9d073c"
	; 1,7,1,5: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,5,1: Type: 50000 (Button) Name: "Calls Toolbar" LocalizedType: "button" AutomationId: "app-bar-20c3440d-c67e-4420-9f80-0e50c39693df"
	; 1,7,1,6: Type: 50007 (ListItem) Name: "More added apps" LocalizedType: "list item"
	; 1,7,1,6,1: Type: 50011 (MenuItem) Name: "More added apps Toolbar" LocalizedType: "menu item" AutomationId: "apps-button"
	; 1,7,1,7: Type: 50007 (ListItem) Name: "More added apps" LocalizedType: "list item"
	; 1,7,1,8: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,8,1: Type: 50000 (Button) Name: "Apps Toolbar" LocalizedType: "button" AutomationId: "discover-apps-button"
	; 1,7,1,9: Type: 50007 (ListItem) LocalizedType: "list item"
	; 1,7,1,9,1: Type: 50011 (MenuItem) Name: "Help Toolbar. Press F1 to open help menu." LocalizedType: "menu item" AutomationId: "app-bar-help-button"
	; 1,8: Type: 50011 (MenuItem) Name: "Opens card" LocalizedType: "menu item"
	; 1,8,1: Type: 50006 (Image) Name: "Profile picture of Keatts, Terry." LocalizedType: "image"
	; 1,9: Type: 50011 (MenuItem) Name: "Opens card" LocalizedType: "menu item" AutomationId: "chat-header-title"
	; 1,10: Type: 50018 (Tab) LocalizedType: "tab" AutomationId: "messages-header-v2-tab-group"
	; 1,10,1: Type: 50019 (TabItem) Name: "Chat" LocalizedType: "tab item"
	; 1,10,2: Type: 50019 (TabItem) Name: "4 more tabs. Press Enter to expand list." LocalizedType: "tab item" AutomationId: "show-more-tabs-button"
	; 1,10,2,1: Type: 50019 (TabItem) Name: "4 more" LocalizedType: "tab item"
	; 1,10,2,2: Type: 50000 (Button) LocalizedType: "button"
	; 1,11: Type: 50000 (Button) Name: "Add a tab" LocalizedType: "button" AutomationId: "add-tab-button-v2"
	; 1,12: Type: 50021 (ToolBar) LocalizedType: "tool bar"
	; 1,12,1: Type: 50000 (Button) Name: "Video call" LocalizedType: "button"
	; 1,12,2: Type: 50000 (Button) Name: "Audio call" LocalizedType: "button"
	; 1,13: Type: 50000 (Button) Name: "Screen sharing" LocalizedType: "button"
	; 1,14: Type: 50000 (Button) Name: "Add people" LocalizedType: "button"
	; 1,15: Type: 50000 (Button) Name: "Pop out chat" LocalizedType: "button" AutomationId: "chat-header-popout-button"
	; 1,16: Type: 50021 (ToolBar) LocalizedType: "tool bar"
	; 1,17: Type: 50026 (Group) LocalizedType: "group" AutomationId: "aria-live-assertive"
	; 1,18: Type: 50026 (Group) LocalizedType: "group" AutomationId: "aria-live-polite"
	; 1,19: Type: 50026 (Group) LocalizedType: "group"
	; 1,19,1: Type: 50030 (Document) Name: "Microsoft Teams" Value: "https://teams.microsoft.com/multi-window/?agent=electron&version=23110224705" LocalizedType: "document"
	; 1,19,1,1: Type: 50026 (Group) Name: "Message List" LocalizedType: "main"
	; 1,19,1,1,1: Type: 50020 (Text) Name: "11:27 AM" LocalizedType: "text"
	; 1,19,1,1,2: Type: 50020 (Text) Name: "Bacon, Adam JCD" LocalizedType: "text"
	; 1,19,1,1,3: Type: 50026 (Group) Name: "UTF-8 is not required. Removed. Bacon, Adam JCD Today at 11:27 AM." LocalizedType: "group" AutomationId: "message-body-1703964463954"
	; 1,19,1,1,3,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,3,2: Type: 50020 (Text) Name: "UTF-8 is not required. Removed." LocalizedType: "text"
	; 1,19,1,1,4: Type: 50020 (Text) Name: "11:28 AM" LocalizedType: "text"
	; 1,19,1,1,5: Type: 50020 (Text) Name: "Bacon, Adam JCD" LocalizedType: "text"
	; 1,19,1,1,6: Type: 50026 (Group) Name: "P.S. I'm on my phone, not at my computer Bacon, Adam JCD Today at 11:28 AM." LocalizedType: "group" AutomationId: "message-body-1703964486402"
	; 1,19,1,1,6,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,6,2: Type: 50020 (Text) Name: "P.S. I'm on my phone, not at my computer" LocalizedType: "text"
	; 1,19,1,1,7: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,8: Type: 50020 (Text) Name: "11:28 AM" LocalizedType: "text"
	; 1,19,1,1,9: Type: 50026 (Group) Name: "I only seem to get that error when the directory or json file is missing. Keatts, Terry Today at 11:28 AM." LocalizedType: "group" AutomationId: "message-body-1703964505612"
	; 1,19,1,1,9,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,9,2: Type: 50020 (Text) Name: "I only seem to get that error when the directory or json file is missing." LocalizedType: "text"
	; 1,19,1,1,10: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,11: Type: 50020 (Text) Name: "11:28 AM" LocalizedType: "text"
	; 1,19,1,1,12: Type: 50026 (Group) Name: "added to the startup to create a directory if one doesn't exist and that fixed that Keatts, Terry Today at 11:28 AM." LocalizedType: "group" AutomationId: "message-body-1703964521128"
	; 1,19,1,1,12,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,12,2: Type: 50020 (Text) Name: "added to the startup to create a directory if one doesn't exist and that fixed that" LocalizedType: "text"
	; 1,19,1,1,13: Type: 50020 (Text) Name: "11:29 AM" LocalizedType: "text"
	; 1,19,1,1,14: Type: 50020 (Text) Name: "Bacon, Adam JCD" LocalizedType: "text"
	; 1,19,1,1,15: Type: 50026 (Group) Name: "Did you use what I just wrote? Or not, and you were getting the error? Bacon, Adam JCD Today at 11:29 AM." LocalizedType: "group" AutomationId: "message-body-1703964592396"
	; 1,19,1,1,15,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,15,2: Type: 50020 (Text) Name: "Did you use what I just wrote? Or not, and you were getting the error?" LocalizedType: "text"
	; 1,19,1,1,16: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,17: Type: 50020 (Text) Name: "11:37 AM" LocalizedType: "text"
	; 1,19,1,1,18: Type: 50026 (Group) Name: "My error is not the UTF-8, it is that I keep making invalid directorys.... Keatts, Terry Today at 11:37 AM." LocalizedType: "group" AutomationId: "message-body-1703965038543"
	; 1,19,1,1,18,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,18,2: Type: 50020 (Text) Name: "My error is not the UTF-8, it is that I keep making invalid directorys...." LocalizedType: "text"
	; 1,19,1,1,19: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,20: Type: 50020 (Text) Name: "11:41 AM" LocalizedType: "text"
	; 1,19,1,1,21: Type: 50026 (Group) Name: "Now have integrated RecLib into Omnibar Keatts, Terry Today at 11:41 AM." LocalizedType: "group" AutomationId: "message-body-1703965296018"
	; 1,19,1,1,21,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,21,2: Type: 50020 (Text) Name: "Now have integrated RecLib into Omnibar" LocalizedType: "text"
	; 1,19,1,1,22: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,23: Type: 50020 (Text) Name: "11:42 AM" LocalizedType: "text"
	; 1,19,1,1,24: Type: 50026 (Group) Name: "Also, migrated the former managed Rec library content onto the JSON file Keatts, Terry Today at 11:42 AM." LocalizedType: "group" AutomationId: "message-body-1703965329126"
	; 1,19,1,1,24,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,24,2: Type: 50020 (Text) Name: "Also, migrated the former managed Rec library content onto the JSON file" LocalizedType: "text"
	; 1,19,1,1,25: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,26: Type: 50020 (Text) Name: "12:14 PM" LocalizedType: "text"
	; 1,19,1,1,27: Type: 50026 (Group) Name: "BTW, confirmed that the compiled code reads and write to JSON just fine Keatts, Terry Today at 12:14 PM." LocalizedType: "group" AutomationId: "message-body-1703967258351"
	; 1,19,1,1,27,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,27,2: Type: 50020 (Text) Name: "BTW, confirmed that the compiled code reads and write to JSON just fine" LocalizedType: "text"
	; 1,19,1,1,28: Type: 50020 (Text) Name: "12:55 PM" LocalizedType: "text"
	; 1,19,1,1,29: Type: 50020 (Text) Name: "Bacon, Adam JCD" LocalizedType: "text"
	; 1,19,1,1,30: Type: 50026 (Group) Name: "Sent Begin Reference, Also, migrated the former managed Rec library content onto the JSON file, Keatts, Terry, 12/30/2023, 11:42 AM, End reference into a separate json file? or the same json file as you are using for startup stuff? Bacon, Adam JCD Today at 12:55 PM." LocalizedType: "group" AutomationId: "message-body-1703969732247"
	; 1,19,1,1,30,1: Type: 50000 (Button) Name: "More message options" LocalizedType: "button"
	; 1,19,1,1,30,2: Type: 50026 (Group) Name: "Begin Reference, Also, migrated the former managed Rec library content onto the JSON file, Keatts, Terry, 12/30/2023, 11:42 AM, End reference" LocalizedType: "group"
	; 1,19,1,1,30,2,1: Type: 50020 (Text) Name: "Keatts, Terry" LocalizedType: "text"
	; 1,19,1,1,30,2,2: Type: 50020 (Text) Name: "12/30/2023, 11:42 AM" LocalizedType: "text"
	; 1,19,1,1,30,2,3: Type: 50020 (Text) Name: "Also, migrated the former managed Rec library content onto the JSON file" LocalizedType: "text"
	; 1,19,1,1,30,3: Type: 50020 (Text) Name: "into a separate json file? or the same json file as you are using for startup stuff?" LocalizedType: "text"
	; 1,19,1,2: Type: 50026 (Group) Name: "Compose" LocalizedType: "region"
	; 1,19,1,2,1: Type: 50020 (Text) Name: "Keatts, Terry is out of office and may not respond" LocalizedType: "text"
	; 1,19,1,2,2: Type: 50000 (Button) Name: "Close Keatts, Terry is out of office and may not respond" LocalizedType: "button" AutomationId: "banner-close-button-id"
	; 1,19,1,2,3: Type: 50004 (Edit) Name: "Type a message" Value: "Inspections: [see below for details on Inspections]
	; - Hoist and head frame: Annual. MSHA regulations require adequate levels of equipment maintenance be conducted.
	; - Frame/Shaft/Rails: W
	; - Other: D
	; - Subsidence audits are performed continuously." LocalizedType: "edit" AutomationId: "new-message-81d9c242-1978-48eb-aee7-f654c1bae43e"
	; 1,19,1,2,3,1: Type: 50020 (Text) Name: "Inspections: [see below for details on Inspections]" LocalizedType: "text"
	; 1,19,1,2,3,2: Type: 50020 (Text) Name: "- Hoist and head frame: Annual. MSHA regulations require adequate levels of equipment maintenance be conducted." LocalizedType: "text"
	; 1,19,1,2,3,3: Type: 50020 (Text) Name: "- Frame/Shaft/Rails: W" LocalizedType: "text"
	; 1,19,1,2,3,4: Type: 50020 (Text) Name: "- Other: D" LocalizedType: "text"
	; 1,19,1,2,3,5: Type: 50020 (Text) Name: "- Subsidence audits are performed continuously." LocalizedType: "text"
	; 1,19,1,2,4: Type: 50021 (ToolBar) Name: "Actions for new message" LocalizedType: "tool bar"
	; 1,19,1,2,4,1: Type: 50000 (Button) Name: "Format (Ctrl+Shift+X)" LocalizedType: "button"
	; 1,19,1,2,4,2: Type: 50011 (MenuItem) Name: "Set delivery options" LocalizedType: "menu item"
	; 1,19,1,2,4,3: Type: 50011 (MenuItem) Name: "Attach files" LocalizedType: "menu item"
	; 1,19,1,2,4,4: Type: 50026 (Group) LocalizedType: "group"
	; 1,19,1,2,4,4,1: Type: 50011 (MenuItem) Name: "Messaging extensions" LocalizedType: "menu item"
	; 1,19,1,2,5: Type: 50021 (ToolBar) Name: "Video clip" LocalizedType: "tool bar"
	; 1,19,1,2,5,1: Type: 50000 (Button) Name: "Record a video clip" LocalizedType: "button"
	; 1,19,1,2,6: Type: 50021 (ToolBar) Name: "Send" LocalizedType: "tool bar"
	; 1,19,1,2,6,1: Type: 50000 (Button) Name: "Send (Ctrl+Enter)" LocalizedType: "button"
	; 1,19,1,3: Type: 50026 (Group) LocalizedType: "group"
	; 1,19,1,4: Type: 50026 (Group) LocalizedType: "group"
	; 1,19,1,5: Type: 50017 (StatusBar) LocalizedType: "status" AutomationId: "a11y-status-message"
	; 2: Type: 50033 (Pane) LocalizedType: "pane" ClassName: "Intermediate D3D Window"
	; 3: Type: 50033 (Pane) LocalizedType: "pane"
	; 3,1: Type: 50033 (Pane) LocalizedType: "pane"
	; 3,1,1: Type: 50033 (Pane) LocalizedType: "pane"
	; 3,1,1,1: Type: 50033 (Pane) LocalizedType: "pane"
	; 3,1,1,1,1: Type: 50033 (Pane) LocalizedType: "pane"
}
; :*:`  ::{bs 1}{right}{Space} ;? testing purposes only
#HotIf
#HotIf !WinActive('ahk_exe hznHorizon.exe')
^!v::NotHznPaste()
NotHznPaste(*) {
	Static Msg := WM_PASTE := 770, wParam := 0, lParam := 0
	; WinActive('ahk_exe WINWORD.exe') ? Send('+{Insert}') : hCtl := ControlGetFocus('A')
	; Run('WINWORD.exe /x /q /a /w')

	WinActive('ahk_exe Code.exe') ? Send('^v') : ((WinActive('ahk_exe WINWORD.exe') || WinActive('ahk_exe Teams.exe')) ? Send('+{Insert}') : tryDll())

	tryDll() {
		hCtl := tryHwnd()
		try {
			DllCall('SendMessage', 'Ptr', hCtl , 'UInt', Msg, 'UInt', wParam, 'UIntP', lParam)
		} 
		catch {
			Send('+{Insert}')
		}
	}

}

#HotIf
; --------------------------------------------------------------------------------
:*:;---::
{
	Send('+{Home}')
	A_Clipboard := '; ---------------------------------------------------------------------------'
	; Send('; {- 75}')
	Send('^v')
}
#HotIf WinActive(" - Visual Studio Code")
; :CB0:function::
:*C1:function...::
{
	BlockInput(1)
	Send('^+{Left}')
	; Sleep(10)
	; Send('{Del}')
	; Sleep(100)
	Send('function ...:')
	Sleep(100)
	Send(A_Tab)
	BlockInput(0)

}
; --------------------------------------------------------------------------------
:?C1*:{Ins::
{
	BlockInput(1)
	Send('^+{Left}')
	Send('Insert')
	Send('{Esc}')
	BlockInput(0)
}
:?C1*:{Ent::
{
	BlockInput(1)
	Send('^+{Left}')
	SendText('Enter')
	Send('{Esc}')
	BlockInput(0)

}
:*:sleep::
{
	BlockInput(1)
	Send('^+{Left}')
	Send('Sleep(100)')
	Sleep(200)
	Send('{Left}')
	Send('+{Left 3}')
	BlockInput(0)
}
#HotIf
; ^#g::git_InstallAHKLibrary('https://github.com/Axlefublr/lib-v2/blob/1d1940095902e483a56bbdd8af6ab0642d8a9fe6/Scr/Keys/Trinity.ahk','Scr\Keys\')

; --------------------------------------------------------------------------------
; #HotIf WinActive(A_ScriptName)
#c::try CenterWindow("A")

CenterWindow(winTitle*) {
    hwnd := WinExist(winTitle*)
    WinGetPos( ,, &W, &H, hwnd)
    mon := GetNearestMonitorInfo(hwnd)
    WinMove(mon.WALeft + mon.WAWidth // 2 - W // 2, mon.WATop + mon.WAHeight // 2 - H // 2,,, hwnd)
}
; --------------------------------------------------------------------------------
; Section .....: Functions
; Function ....: Run scripts selection from the Script Tray Icon
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
; ^+#1::
GUIFE(*){
	Run("GUI_FE.ahk")
}
; list

; ^+#3::
WindowProbe(*){
	Run("WindowProbe.ahk", "C:\Users\bacona\OneDrive - FM Global\3. AHK\")
}
; ^+#4::
GUI_ListofFiles(*){
	Run("GUI_ListofFiles.ahk")
}
; ^+#5::
{
Windows_Data_Types_offline(*){
	Run("Windows_Data_Types_offline.ahk", "C:\Users\bacona\OneDrive - FM Global\3. AHK\AutoHotkey_MSDN_Types-master\src\v1.1_deprecated\")
}
}
; #o::
Detect_Window_Info(*){
	Run("Detect_Window_Info.ahk")
}
; ^+#6::
Detect_Window_Update(*){
	Edit()
}
; ^+#7::
test_script(*){
	Run("test_script.ahk")
}
;---------------------------------------------------------------------------
;      Shift+WIN+m Button to Click on Window Anywhere to Drag
;---------------------------------------------------------------------------
; This script is from the Automator

!LButton::
{
CoordMode("Mouse", "Screen")
MouseGetPos(&x, &y)
WinMove(x, y, , , "A")
return
;---------------------------------------------------------------------------
;                       Time Stamp Code
;---------------------------------------------------------------------------
}
#HotIf WinActive('ahk_exe hznhorizon.exe') || WinActive('ahk_exe git.exe')
:*:ts::
; format month and year
{
	date := FormatTime(A_Now, "yyyy.MM")
	Send("(AJB - " date ")")
	return
}
#HotIf
:*:tsf::
; format month and year
{
	date := FormatTime(A_Now, "yyyy.MM.dd")
	Send("(AJB - " date ")")
	return
}
;---------------------------------------------------------------------------
;                      Helpful Stuff
;---------------------------------------------------------------------------
:*:attext::	; Timestamp
{
Date := FormatTime(A_Now, "MM/dd/yyyy")	; format month, day and year
IB := InputBox("Nameplate Head Thickness (after 0.)", "Air Tank", "w300 h125"), npht := IB.Value
IB := InputBox("Nameplate Shell Thickness (after 0.)", "Air Tank", "w300 h125"), npst := IB.Value
IB := InputBox("Actual Head Thickness (after 0.)", "Air Tank", "w300 h125"), aht := IB.Value
IB := InputBox("Actual Shell Thickness (after 0.)", "Air Tank", "w300 h125"), ast := IB.Value
Send("Nameplate HD: 0." npht " / SH: 0." npst " // " Date " TJK HD: 0." aht " / SH: 0." ast)
Return
} ; Added bracket before function

Join(sep, params*) {
	for index,param in params
		str .= param . sep
	return SubStr(str, 1, -StrLen(sep))
}
;MsgBox % Join("`n", "one", "two", "three")

;---------------------------------------------------------------------------
;                          General Abbreviations
;---------------------------------------------------------------------------

^!0::
{
	static var := "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔŒÕÖØÙÚÛÜßàáâãäåæçèéêëìíîïñòóôœõöøùúûüÿ¿¡«»§¶†‡•-–—™©®¢€¥£₤¤αβγδεζηθικλμνξοπρσςτυφχψωΓΔΘΛΞΠΣΦΨΩ∫∑∏√−±∞≈∝≡≠≤≥×·÷∂′″∇‰°∴ø∈∩∪⊂⊃⊆⊇¬∧∨∃∀⇒⇔→↔↑ℵ∉°₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹"
	Global arr
	arr := strsplit(var)

	static w:=20, cnt := 14,
	myGui := Gui()
	myGui.Opt("-caption")
	myGui.MarginX := "0", myGui.MarginY := "0"
	myGui.SetFont("s10")
	Loop arr
		{
		x := mod((a_index - 1),cnt) * w, y := floor((a_index - 1)/ cnt) * w
		ogctextz := myGui.add("text", "x" . x . " y" . y . " w" . w . " h" . w . " center vz" . a_index, arr[a_index])
		ogctextz.OnEvent("Click", myGui.insert.Bind("Normal"))
		}
	myGui.show()
	return
}

insert(A_GuiEvent, GuiCtrlObj, Info, *)
{
	myGui := 
	oSaved := myGui.submit()
	Send(arr[SubStr(A_GuiEvent, 2)])
}

;---------------------------------------------------------------------------
;      Alt+Left Mouse Button to Click on Window Anywhere to Drag
;---------------------------------------------------------------------------
; This script modified from the original: http://www.autohotkey.com/docs/scripts/EasyWindowDrag.htm

Alt & ~LButton::EWD_MoveWindow()
; ~MButton & LButton::
; CapsLock & LButton::
; EWD_MoveWindow(*)
; {
;     CoordMode "Mouse"  ; Switch to screen/absolute coordinates.
;     MouseGetPos &EWD_MouseStartX, &EWD_MouseStartY, &EWD_MouseWin
;     WinGetPos &EWD_OriginalPosX, &EWD_OriginalPosY,,, EWD_MouseWin
;     if !WinGetMinMax(EWD_MouseWin)  ; Only if the window isn't maximized 
;         SetTimer EWD_WatchMouse, 10 ; Track the mouse as the user drags it.

;     EWD_WatchMouse()
;     {
;         if !GetKeyState("LButton", "P")  ; Button has been released, so drag is complete.
;         {
;             SetTimer , 0
;             return
;         }
;         if GetKeyState("Escape", "P")  ; Escape has been pressed, so drag is cancelled.
;         {
;             SetTimer , 0
;             WinMove EWD_OriginalPosX, EWD_OriginalPosY,,, EWD_MouseWin
;             return
;         }
;         ; Otherwise, reposition the window to match the change in mouse coordinates
;         ; caused by the user having dragged the mouse:
;         CoordMode "Mouse"
;         MouseGetPos &EWD_MouseX, &EWD_MouseY
;         WinGetPos &EWD_WinX, &EWD_WinY,,, EWD_MouseWin
;         SetWinDelay -1   ; Makes the below move faster/smoother.
;         WinMove EWD_WinX + EWD_MouseX - EWD_MouseStartX, EWD_WinY + EWD_MouseY - EWD_MouseStartY,,, EWD_MouseWin
;         EWD_MouseStartX := EWD_MouseX  ; Update for the next timer-call to this subroutine.
;         EWD_MouseStartY := EWD_MouseY
;     }
; }

#HotIf WinExist(A_ScriptName)
; ; Current date and time
FormatDateTime(format, datetime:="") {
	if (datetime = "") {
		datetime := A_Now
	}
	CurrentDateTime := FormatTime(datetime, format)
	SendInput(CurrentDateTime)
	return
}
; ---------------------------------------------------------------------------
;?			Hotstrings
; ---------------------------------------------------------------------------
:X:/datetime::FormatDateTime("dddd, MMMM dd, yyyy, HH:mm")
:X:/datetimett::FormatDateTime("dddd, MMMM dd, yyyy hh:mm tt")
:X:/cf::FormatDateTime("yyyy.MM.dd HH:mm")
:X:/time::FormatDateTime("HH:mm")
:X:/timett::FormatDateTime("hh:mm tt")
:X:/date::FormatDateTime("MMMM dd, yyyy")
:X:/daten::FormatDateTime("MM/dd/yyyy")
:X:/datet::FormatDateTime("yy.MM.dd")
:X:/week::FormatDateTime("dddd")
:X:/day::FormatDateTime("dd")
:X:/month::FormatDateTime("MMMM")
:X:/monthn::FormatDateTime("MM")
:X:/year::FormatDateTime("yyyy")
::wtf::Wow that's fantastic
:X:/paste::Send(A_Clipboard)
; ::/cud::
;     ; useful for WSLs
; {
;     SendInput("/mnt/c/Users/" A_UserName "/")
; Return
; }
::/nrd::npm run dev
::/gm::Good morning
::/ge::Good evening
::/gn::Good night
::/ty::Thank you
::/tyvm::Thank you very much
::/wc::Welcome
::/mp::My pleasure
::/lorem::Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
::/plankton::Plankton are the diverse collection of organisms found in water that are unable to propel themselves against a current. The individual organisms constituting plankton are called plankters. In the ocean, they provide a crucial source of food to many small and large aquatic organisms, such as bivalves, fish and whales.

#HotIf