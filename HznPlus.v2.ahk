/************************************************************************
 * function ......: Horizon Button ==> A Horizon function library
 * @description ..: This library is a collection of functions and buttons that deal with missing interfaces with Horizon.
 * @file HznPlus.v2.ahk
 * @author OvercastBTC
 * @date 2023.09.15
 * @version 3.0.2
 * @ahkversion v2+
 * @Section .....: Auto-Execution
 ***********************************************************************/
; --------------------------------------------------------------------------------
/************************************************************************
 * function ...........: Resource includes for .exe standalone
 * @author OvercastBTC
 * @date 2023.09.28
 * @version 3.0.2
 ***********************************************************************/
; --------------------------------------------------------------------------------
Persistent(1)
#Include <Directives\__AE.v2>
#Include <Directives\__HznToolbar>
#Include <System\DPI>
; --------------------------------------------------------------------------------
; --------------------------------------------------------------------------------
#Include <EnumAllMonitorsDPI.v2>
#Include <GetNearestMonitorInfo().v2>
#Include <Class\RTE\RichEdit>
#Include <Class\RTE\RichEditDlgs>
#Include <__receiver>
#Include <Tools\Info>
; #Include <Misc\Out>
#Include <System\UIA>
#Include <RTE.v2\Project Files\RichEdit_Editor_v2>

; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
/************************************************************************
* ;Description ...: Create Tray Icon
* @description ...: Create the tray icon using the embedded B64 via Create_HznHorizon_ico()
* @author OvercastBTC
* @date 2023.08.15
* @version 2.0.1
 ***********************************************************************/
TraySetIcon('HICON:' Create_HznHorizon_ico())
; ---------------------------------------------------------------------------
; Check_Startup_Status()
; ---------------------------------------------------------------------------
/************************************************************************
* @Title .........: Create Tray Menu
* @Description ...: Create options (//TODO add functions here) for use with the script (e.g., run at startup, open horizon, etc.)
* @file HznPlus.v2.ahk
* @author OvercastBTC
* @date 2023.08.09
* @version 0.0.1
* @TODO ;[ ] TODO	.: Convert from v1 to v2
 ***********************************************************************/
; -------------------------------------------------------------------------------------------------

; HznTray := A_TrayMenu
; see at the bottom for the functions
; return
; --------------------------------------------------------------------------------------------------
/************************************************************************
; function ...: Horizon Hotkeys
@description: Hotkeys (shortcuts) for normal Windows hotkeys that should exist
@author OvercastBTC
@notes
@function Italics 		(CTRL+I)
@function Bold			(CTRL+B)
@function Underline		(CTRL+U)
@function SelectAll		(CTRL+A)
@function Save			(CTRL+S)
@function Save			(CTRL+S)
@function HznGetText 	(CTRL+SHIFT+C) (like Ctrl+c & displays the text)
@function Find			(CTRL+SHIFT+F)
@function Replace		(CTRL+SHIFT+R)
 ***********************************************************************/

; --------------------------------------------------------------------------------
#Hotif WinActive(A_ScriptName) && WinExist('__A_Process.v2.ahk')
	$^s::
	{
		Run(WinExist('__A_Process.v2.ahk'))
		Run(A_ScriptName)
	}
#HotIf
#HotIf WinActive('ahk_exe hznhorizon.exe')

tab::
{
	try {
		ClassNN := ControlGetClassNN(ControlGetFocus('A'))
		if ((ClassNN ~= 'i).*TextBox.*') || (ClassNN ~= 'i).*TX11.*')) {
		Suspend(1)
		SendEvent('^i')
		Suspend(0)
		}
		Else {
			SendEvent('{sc0f}')
		}
	}
	return
}
hznButtonCount(hTb := HznToolbar._hTb()) {
	Static Msg := TB_BUTTONCOUNT := 1048, wParam := 0, lParam := 0
	BUTTONCOUNT := SendMessage(TB_BUTTONCOUNT, wParam, lParam, , hTb)
	OutputDebug('Used hznButtonCount()`n # of buttons is ' BUTTONCOUNT)
	return BUTTONCOUNT
}
; --------------------------------------------------------------------------------
F5::button(120) 			; find (focused tab) and find/replace
^f::button(120)				; find (focused tab) and find/replace
^h::HznFindReplace()	; find/replace (focused tab) and find
; --------------------------------------------------------------------------------
^a::HznSelectAll()
; ^+a::
; {
; 	value := AE_Get_TEXTLIMIT()
; 	MsgBox(value)
; }
^End::AE_Select_End()
^Home::AE_Select_Beginning()

^b::button(100)			; bold
^i::button(101)			; italics
^u::button(102)			; underline
; button(103) = Separator
; ^l::button(104) 		; Align Left
; ^r::button(105) 		; Align Right ; fix => cut???
; ^e::button(106) 		; Align Center
; ^j::button(107) 		; Justitfied
; button(108) = Separator
; ^x::button(109) 		; cut
; ^c::button(110) 		; copy
; ^v::button(111) 		; paste
; ^z::button(113) 		; undo
; ^y::button(114) 		; redo
; idCommand (115) unknown but does something?
+F12::button(116) 		; Bulleted List

F7::button(117)			; spell check
^F12::button(118)		; Insert Table ; fix (drop down)
^=::
{
	button(119) 		; [super script]
	SendEvent('{Space}')
	SendEvent('!o')
	; return
}	
^+=::					; [sub script]
{
	button(119) 		
	SendEvent('{Down}')
	SendEvent('{Space}')
	SendEvent('!o')
	; return
}

; --------------------------------------------------------------------------------

^+v::HznPaste()
^+c::HznGetText()
*^s::HznSave()
^F4::HznClose() 		; fix [] => need to only be on certain screen(s).
; ^+9::HznEnableButtons(hTb := HznToolbar._hTb()) ; works!!!
^+8::HznTbCustomize() 	; works!!! enables and shows all buttons on the toolbar
; ^+7:: 					; ! test hotkey
; {
; 	hCtl := HznToolbar.hCtl()
; 	hTb := HznToolbar.hTb()
; 	; ControlSetEnabled(1,nCtl)
; 	n:=SendMessage(WM_ENABLE := 0x000A, true,,hTb,hTb)
; 	n1:=DllCall('EnableWindow', 'Ptr', hTb, 'Int', true)
; 	; DllCall ("CreateWindowEx"
; 	; 				, "Uint", 0
; 	; 				, "str",  'Toolbar32'         		; ClassName
; 	; 				, "str",  "msvb_lib_toolbar"  		; WindowName
; 	; 				, "Uint", 0x40000000 | 0x10000000 	; WS_CHILD, WS_VISIBLE
; 	; 				, "int",  0                       	; Left
; 	; 				, "int",  0                       	; Top
; 	; 				, "int",  200                     	; Width
; 	; 				, "int",  200                     	; Height
; 	; 				, "Uint", hCtl	            			; hWndParent
; 	; 				, "Uint", 0                       	; hMenu
; 	; 				, "Uint", fCtl_I                    ; hInstance
; 	; 				, "Uint", 0)
; 	transp := WinGetTransparent(hTb)
; 	; Dllcall ("SetLayeredWindowAttributes", "ptr", hTb, "ptr", 0, "char", 255, "uint", 2)
; 	SendMessage(0x421,,,,hTb)
; 	SendMessage(0x421,,,,hTb)
; 	n2:=DllCall('ShowWindow', 'Ptr', hTb, 'Int', 1)
; 	; n3:=DllCall('WinMain', 'UInt', hTb,'UInt',0, 'Int', 1, 'Str','ahk_exe hznhorizon.exe')
; 	n3:=DllCall('IsWindowVisible', 'Ptr', hTb)
; 	MsgBox(hCtl '`n'
; 			. fCtl '`n'
; 			. transp '`n'
; 			. fCtlI '`n'
; 			. nCtl '`n'
; 			. hTb '`n'
; 			. n '`n'
; 			. n1 '`n'
; 			. n2 '`n'
; 			. n3 '`n'
; 			)
; 	return		
; }
#HotIf
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
/************************************************************************
* Function ...: HznFindReplace() (Ctrl+Shift+f || Ctrl+r)
* @author OvercastBTC 
***********************************************************************/

HznFindReplace()
{
	SendLevel(5)
	Send('^f')
	; Sleep(100)
	SendEvent('{LAlt down}')
	SendEvent('p')
	; Sleep(100)
	Send('{LAlt Up}')
	; return
}
; --------------------------------------------------------------------------------
/************************************************************************
* Function ...: HznSave() (Ctrl-s)
* @author OvercastBTC 
***********************************************************************/

HznSave()
{
	SendLevel((A_SendLevel+1))
	BlockInput(1)
	pvTxt := A_DetectHiddenText, 	DetectHiddenText(0)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(0)
	fCtl := ControlGetFocus('A')
	; pvKey := A_KeyDelay, 			SetKeyDelay(20)
	; Sleep(200)
	idWin := WinGetID('ahk_exe hznHorizon.exe')
	; Sleep(200)
	WinActivate(idWin)
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin) ;, SetKeyDelay(pvKey)
	; CaretGetPos(&cX:=0, &xY:=0)
	SendEvent('{LAlt down}')
	SendEvent('f') ; SendEvent required
	; Sleep(50)
	; SendEvent('s')
	SendEvent('{LAlt Up}')
	Sleep(100)
	; Send('{Down}' . '{Up}')
	; try {
	; 	hCtl := ControlGetFocus('A')
	; 	nCtl := ControlGetClassNN(hCtl)
	; 	hWnd := WinExist('A')
	; 	win_title := WinGetTitle(hWnd)
	; }
	; ControlChooseString('Save',hCtl,hWnd)
	; ControlChooseIndex(1, hCtl, hWnd)
	; ToolTip(win_title . ' (' . hWnd . ')' . '`n' . nCtl . ' (' . hCtl . ')')
	SendEvent('{Enter}')
	Sleep(100)
	ControlFocus(fCtl, 'A')
	BlockInput(0)
	SendLevel(0)
	return
}
; --------------------------------------------------------------------------------
/************************************************************************
 * @function HznClose() (Ctrl-F4)
 * @author OvercastBTC 
 * @author Descolada and his UIA library and functions
 * @description Horizon Close Button (Ctrl-F4) ==> ≠ Alt-F4
***********************************************************************/
; --------------------------------------------------------------------------------

HznClose(*)
{
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	; prevents the user from doing anything to screw it up by blocking all input
	BlockInput(1)
	; --------------------------------------------------------------------------------
	/** @function Alt + - (Hzn Hotkey for hidden menuitems )						*/
	pvTxt := A_DetectHiddenText, 	DetectHiddenText(0)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(0)
	idWin := WinGetID('ahk_exe hznHorizon.exe')
	WinActivate(idWin)
	Sleep(100)
	; ; --------------------------------------------------------------------------------
	; Send('{LAlt down}')
	; SendEvent('-') ; SendEvent required
	; Sleep(100)
	; Send('{LAlt Up}')
	; Sleep(300)
	; Send('c')
	; --------------------------------------------------------------------------------
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	; MenuSelect(idWin,,'2&')
	SendLevel(0)
	BlockInput(0)
	return
}

; -------------------------------------------------------------------------------------------------
/************************************************************************
 * Desc ......: Horizon Hotkey - Select-All (Ctrl-A)
 * @author ...: Descolada, OvercastBTC
 * Function ..: Select-All() (Ctrl-A)
 * @param Msg - EM_SETSEL := 177 - the Windows API message for "Set Selection"
 * @param wParam - := 0
 * @param lParam := -1
***********************************************************************/

HznSelectAll(*)
{
	Static Msg := EM_SETSEL := 177, wParam := 0, lParam := -1
	hCtl := ControlGetFocus('A')
	DllCall('SendMessage', 'UInt', hCtl, 'UInt', Msg, 'UInt', wParam, 'UIntP', lParam)
	return
}
; --------------------------------------------------------------------------------

/************************************************************************
 * Function ..: HznGetText() (Ctrl+Shift+c)
 * @author ...: OvercastBTC
 * @param hCtl ........: Gets the handle (hwnd) to the control in focus
 * @param A_Clipboard .: The AHK builtin clipboard
 * @returns {text in the control (RT6TextBox or TX11)}
***********************************************************************/

HznGetText(*)
{
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	BlockInput(1) ; 1 = On, 0 = Off
	rect := Map()
	hRect := []
	static hCtl := ControlGetFocus('A')
	static nCtl := ControlGetClassNN(hCtl)
	static hCtl_title := WinGetTitle(hCtl)
	static hHzn := WinExist('ahk_exe hznHorizon.exe')
	; hznUIA := UIA.ElementFromHandle(hCtl).BoundingRectangle
	; Infos(
	; 	hznUIA.l . '`n' .
	; 	hznUIA.r . '`n' .
	; 	hznUIA.t . '`n' .
	; 	hznUIA.b
	; )
	;// rect := WindowGetRect(hCtl)
	;// rect := GetClientSize(hCtl)
	;// rect := GetRect(hCtl)
	;// mRect := GetClientRect(hCtl, hCtl_title)
	;// width := mRect['width']
	;// height := mRect['height']
	ControlGetPos(&x, &y, &w, &h, nCtl, 'A')
	mRect := WindowGetRect(hCtl)
	; width := mRect.width
	; height := mRect.height
	width := Round((w - (w/3)))
	height := h
	rect.Set('width',width, 'height',height, 'x', x, 'y', y)
	; infos('Title: ' . hCtl_title . '`n' 
	; 	. 'ClassNN: ' . nCtl . ' (' . hCtl . ')' . '`n' 
	; 	. 'width: ' . width . '`n' 
	; 	. ' w: ' . w . '`n'
	; 	. 'height: ' . height . '`n'
	; 	. 'h: ' . h . '`n'
	; 	. 'x: ' . x . ' y: ' . y . '`n'
	; 	; hznUIA.l . '`n' .
	; 	; hznUIA.r . '`n' .
	; 	; hznUIA.t . '`n' .
	; 	; hznUIA.b
	; )
	MouseMove(x, y, -1)
	RTE_Title := 'hznRTE - A Rich Text Editor for Horizon'
	file_name := '__receiver.ahk'
	file_line := '', TX11 := '', match:='', aLine := ''
	aTX11 := [], new_map := []
	width_needle := 'i)width', height_needle := 'i)height',	x_needle := 'i)x', y_needle := 'i)y'
	match_array := [width_needle, height_needle, x_needle, y_needle]
	; --------------------------------------------------------------------------------
	HznSelectAll(hCtl)
	clip_it()
	; --------------------------------------------------------------------------------
	hzntx11 := FileOpen(file_name,'rw','UTF-8')
	Sleep(300)
	; --------------------------------------------------------------------------------
	Update_Receiver_Rect_Map(file_name)
	; --------------------------------------------------------------------------------
	Update_Receiver_Rect_Map(file_name){
		loop read file_name {
			aTX11.Push(A_LoopReadLine)
			file_line .= A_LoopReadLine . '`n'
		}
		; Sleep(500)
		; --------------------------------------------------------------------------------
		for each, aLine in aTX11 {
			; --------------------------------------------------------------------------------
			for each, match in match_array {
				if ((aLine ~= match)) {
					str_match := StrSplit(match,')','i ) "')
					rMatch := str_match[2]
					; Infos(rMatch)
					rect_match := rect[rMatch]
					; rect_match := dpiRect.%str_match[2]%
					new_str := RegExReplace(aLine, ':= ([0-9].*)', ':= ' rect_match . ',' )
					aLine := new_str
					aTX11.RemoveAt(A_Index)
					aTX11.InsertAt(A_Index, aLine)
				}
			}
			; --------------------------------------------------------------------------------
			/**
			 * function: write each value in the array to @param TX11 (string variable) 
			 */
			TX11 .= aLine . '`n'
			; --------------------------------------------------------------------------------
		}
		; Infos(file_line . '`n' . TX11)
		
		; --------------------------------------------------------------------------------
		; hzntx11 := FileOpen(file_name,'rw','UTF-8')
		hzntx11.Write(TX11)
		hzntx11 := ''
		return 0
	}
	
	; -----------------------------------------------------------------------------
	; Info('Loading Rich Text Editor...', 3000)
	; Sleep(500)
	
	RTE()
	hRTE := WinWaitActive('hznRTE ')
	pid_RTE := WinGetPID(hRTE)
	Sleep(100)
	
	clip_it()
	; Clip_it(1)
	Send('^v')
	; --------------------------------------------------------------------------------
	; MsgBox(text '`n`n' '[This text has been copied to the clipboard. Use Ctrl+v to paste, or right-click and select Paste in your window of choice.]`n`n[This message will autoclose in 30 seconds]','Copy of Horizon Text', 'T30')
	; MsgBox('[This text has been copied to the clipboard. Use Ctrl+v, or right-click and select paste, to paste in your window of choice.]`n`n[This message will autoclose in 3 seconds]','Copy of Horizon Text', 'T3')
	Info('[This text has been copied to the clipboard. Use Ctrl+v, or right-click and select paste, to paste in your window of choice.]`n`n[This message will autoclose in 3 seconds]', 3000)
	; --------------------------------------------------------------------------------
	; return text
	ProcessWaitClose(pid_RTE)
	; FileCloseFN()
	WinActivate(hCtl)
	ControlFocus(hCtl, hCtl_title)
	HznSelectAll(hCtl)
	; Send('+{Insert}') ;! <==== DON'T USE THIS, IT PASTES THE WINDOWS DEFAULT FONT WHICH IS ARIAL 12, HORIZON IS TIMES NEW ROMAN 11.
	Send('^v')
	BlockInput(0)
	SendLevel(0)
	; return
}

; --------------------------------------------------------------------------------
/************************************************************************
 * Function ..: HznPaste() (Ctrl+Shift+v)
 * @author ...: OvercastBTC
 * @source ...: https://learn.microsoft.com/en-us/windows/win32/dataxchg/wm-paste
 * @sourceparam .......: #define WM_PASTE                        0x0302
 * @param hCtl ........: Gets the control in focus hwnd, then the class name
 * @param Msg .........: Windows API message => WM_PASTE
 * @param WM_PASTE ....: Decimal number (versus hex = 0x0302) for WM_PASTE
 * @param wParam ......: This parameter is not used and must be zero.
 * @param lParam ......: This parameter is not used and must be zero.
 * @returns {null} ....: This message does not return a value.
***********************************************************************/

HznPaste(*) {
	Static Msg := WM_PASTE := 770, wParam := 0, lParam := 0
	hCtl := ControlGetFocus('A')
	DllCall('SendMessage', 'Ptr', hCtl, 'UInt', Msg, 'UInt', wParam, 'UIntP', lParam)
	return
}
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
; HznEnableButtons(hTb?, *)
HznEnableButtons(hTb) {
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	BlockInput(1) ; 1 = On, 0 = Off
	; hTb := HznToolbar._hTb()
	; --------------------------------------------------------------------------------
	Static   WM_COMMAND := 273, TB_GETBUTTON:= 1047, TB_BUTTONCOUNT:= 1048,TB_COMMANDTOINDEX := 1049,TB_GETITEMRECT := 1053,MEM_PHYSICAL := 4,MEM_RELEASE := 32768,TB_GETSTATE := 1042,TB_GETBUTTONSIZE := 1082, TB_ENABLEBUTTON := 0x0401, TB_SETSTATE := 0x0411,TBSTATE_ENABLED := 4,TBSTATE_DISABLED := 0, TBSTATE_HIDDEN := 8 
	; --------------------------------------------------------------------------------

	; Step: count and load all the msvb_lib_toolbar buttons into memory
	; --------------------------------------------------------------------------------
	buttonCount := SendMessage(TB_BUTTONCOUNT, 0, 0,, hTb)
	; Infos('Btn Ct: ' buttonCount)
	; --------------------------------------------------------------------------------
	; Step: Use the @params to enable the button
	; --------------------------------------------------------------------------------
	Loop buttonCount {
		idCommand := A_Index +99
		If (idCommand <= 102) {
			Msg := TB_SETSTATE, wParam := idCommand, lParam_HI := 4, lParam_LO := TBSTATE_ENABLED, control := hTb
		; SendMessage(TB_SETSTATE, idCommand, 0|TBSTATE_ENABLED,,hTb)
			SendMessage(Msg, wParam, lParam_HI|lParam_LO,control,hTb)
			tryGetItemList(idCommand)
		}
		If (idCommand > 108) {
			Msg := TB_SETSTATE, wParam := idCommand, lParam_HI := 4, lParam_LO := TBSTATE_ENABLED, control := hTb
		; SendMessage(TB_SETSTATE, idCommand, 0|TBSTATE_ENABLED,,hTb)
			SendMessage(Msg, wParam, lParam_HI|lParam_LO,control,hTb)
		}
		tryGetItemList(control) {
			try {
				button_items_rows := ControlGetItems(control)
				item := ''
				item_list := ''
				; try {
					for item in button_items_rows {
						item_list .= item . '`n'
						FileAppend(item_list, '_button_items_rows.ahk', 'UTF-8')
					}
				; }
			}
		}
	}
	; Infos(item_list)
	BlockInput(0)
	SendLevel(0)
	; return
}
/************************************************************************
 * function: HznTbCustomize()
 * @description : enables and shows all buttons on the toolbar
 * @hotkey 		: ^+8::
 * @param 		hTb 
 * @returns 	{void} 
 ***********************************************************************/

HznTbCustomize(hTb?)
{
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	BlockInput(1) ; 1 = On, 0 = Off
	hTb := HznToolbar._hTb()
	Static TB_CUSTOMIZE := 1051, wParam := 0, lParam := 0, control := hTb
	try SendMessage(TB_CUSTOMIZE, wParam, lParam, control, hTb) ;wow, this works!!!
	BlockInput(0)
	SendLevel(0)
	return
}
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
/************************************************************************
 * @function button()
 * @author ....: Descolada, OvercastBTC
 * @Desc ......:  Call the Horizon msvb_lib_toolbar buttons using the button() function
 * @param A_ThisHotkey - AHK's built in variable.
 * @param idCommand .: optional => idCommand := 0 (preset value, else => idCommand?)
***********************************************************************/
button(idCommand:=0)
{
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	BlockInput(1) ; 1 = On, 0 = Off
	Static  WM_COMMAND := 273, TB_GETBUTTON := 1047, TB_BUTTONCOUNT := 1048
	Static	TB_COMMANDTOINDEX := 1049, TB_GETITEMRECT := 1053 
	Static	MEM_PHYSICAL := 4, MEM_RELEASE := 32768, TB_GETSTATE := 1042
	Static	TB_GETBUTTONSIZE := 1082, TB_ENABLEBUTTON := 0x0401
	hTb := HznToolbar._hTb()
	nCtl := HznToolbar._nCtl()
	; --------------------------------------------------------------------------------
	; HznButton(hTb, idCommand, nCtl)
	; function: !!! ===> Programatically "Click" the button!!! <=== !!!
	Msg := WM_COMMAND, wParam_hi := 0, wParam_lo := idCommand, lParam := control := hTb
	; DllCall('SendMessage', 'UInt', hTb, 'UInt', Msg, 'UInt', wParam_hi | wParam_lo, 'UIntP', lParam)
	SendMessage(Msg, wParam_hi | wParam_lo, lParam, , hTb)
	; --------------------------------------------------------------------------------
	SendLevel(0) ; restore normal SendLevel
	BlockInput(0) ; 1 = On, 0 = Off
	return
}
; --------------------------------------------------------------------------------
/************************************************************************
 * @function idButton()
 * @author ....: Descolada, OvercastBTC
 * @Desc ......:  Call the Horizon msvb_lib_toolbar buttons using the button() function
 * @param A_ThisHotkey - AHK's built in variable.
 * @param idBtn .: optional => idCommand := 0 (preset value, else => idCommand?)
***********************************************************************/

; idButton(buttonhotkey?)
; {
; 	try {
; 		(buttonhotkey >= 100) ? idBtn := (buttonhotkey - 99) : buttonhotkey
; 	}
; 	catch{
; 		idBtn:= (A_ThisHotkey = "^b")  ? 1 	: 	;.........: bold
; 				(A_ThisHotkey = "^i")  ? 2 	: 	;.........: italic
; 				(A_ThisHotkey = "^u")  ? 3 	: 	; ........: underline
; 				(A_ThisHotkey = '^+b') ? 15	:	; ........: Bulleted List
; 				(A_ThisHotkey = 'F5')  ? 19	:	; ........: Find/Replace
; 				(A_ThisHotkey = '^F5') ? 19	:	; ........: Find/Replace
; 				(A_ThisHotkey = 'F7')  ? 16	:	; ........: Spell Check
; 				(A_ThisHotkey = '^F7') ? 16	: OutputDebug('idBtn: ' idBtn '`n')	; ........: Spell Check
; 				; (A_ThisHotkey = "^x")  ? 8 	:	; ........: cut
; 				; (A_ThisHotkey = "^c")  ? 9 	:	; ........: copy
; 				; (A_ThisHotkey = "^v")  ? 10	:	; ........: paste
; 				; (A_ThisHotkey = "^z")  ? 12	:	; ........: undo
; 				; (A_ThisHotkey = "^y")  ? 13	:	; ........: redo
; 				; (A_ThisHotkey = '^+s') ? 18	:	; ........: super|sub script
; 			}
; 	OutputDebug('idBtn: ' idBtn '`n')
; 	return idBtn
; }
; -------------------------------------------------------------------------------------------------
/************************************************************************
 * Function .....: HznButton()
 * @Desc Clicks the nth item in a Win32 application toolbar.
 * @author ......: Descolada, OvercastBTC
 * @param hTb - The handle of the toolbar control.
 * @param hTb - same as hTb
 * @param fCtl* - [OPTIONAL] The ClassNN of the "focused" control.
 * @param fCtlInstance - The ClassNN [instance] of the "focused" control.
 * @param nCtl - The [ClassNN and instance] of the toolbar control.
 * @function param:=0 => in AHK v2 => optional
 ***********************************************************************/
; --------------------------------------------------------------------------------
; HznButton(hTb, nCtl, idCommand, n?, pID?, fCtl?, hTx?, fCtlInstance?) {
HznButton(hTb, idCommand, nCtl?) {
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	BlockInput(1) ; 1 = On, 0 = Off
	; --------------------------------------------------------------------------------
	Static  WM_COMMAND := 273, TB_GETBUTTON := 1047, TB_BUTTONCOUNT := 1048
	Static	TB_COMMANDTOINDEX := 1049, TB_GETITEMRECT := 1053 
	Static	MEM_PHYSICAL := 4, MEM_RELEASE := 32768, TB_GETSTATE := 1042
	Static	TB_GETBUTTONSIZE := 1082, TB_ENABLEBUTTON := 0x0401
	; --------------------------------------------------------------------------------
	; Step: count and load all the msvb_lib_toolbar buttons into memory
	; --------------------------------------------------------------------------------
	buttonCount := SendMessage(TB_BUTTONCOUNT, 0, 0, , hTb)
	; --------------------------------------------------------------------------------
	; Step: Use the @params to press the button
	; --------------------------------------------------------------------------------
	
	try if (idCommand >= 100 && idCommand <= (buttonCount + 99)) {
		; * Get the toolbar "thread" process ID (PID)
		DllCall("GetWindowThreadProcessId", "Ptr", hTb, "UInt*", &tpID := 0)
		; --------------------------------------------------------------------------------
		; * Open the process with PROCESS_VM_OPERATION, PROCESS_VM_READ, and PROCESS_VM_WRITE access
		hProcess := DllCall('OpenProcess', 'UInt', 8 | 16 | 32, "Int", 0, "UInt", tpID, "Ptr")
		; --------------------------------------------------------------------------------
		; * Identify if the process is 32 or 64 bit (efficiency step)
		Is32bit := Win32_64_Bit(hProcess)
		; --------------------------------------------------------------------------------
		; * Allocate memory for the TBBUTTON structure in the target process's address space
		remoteMemory := remote_mem_buff(hProcess, Is32bit, &TBBUTTON_SIZE)
		; --------------------------------------------------------------------------------
		; DllCall("VirtualFreeEx", "Ptr", hProcess, "Ptr", remoteMemory, "UPtr", 0, "UInt", MEM_RELEASE)
		; DllCall("CloseHandle", "Ptr", hProcess)
		; --------------------------------------------------------------------------------
		; Step: Store previous and set min delay
		; --------------------------------------------------------------------------------
		pCD := A_ControlDelay, pMD := A_MouseDelay, pWD := A_WinDelay
		SetControlDelay(-1), SetMouseDelay(-1), SetWinDelay(-1)
		; --------------------------------------------------------------------------------
		; try (idCommand < 100) ? idCommand := ((n - 1) + 100) : idCommand
		; btnstate := GETBUTTONSTATE(idCommand, hTb)
		; If (!btnstate = 4) || (!btnstate = 6) ;! note: (AJB - 09/2023) verified
		; 	return
		
		; --------------------------------------------------------------------------------
		; ; function: !!! ===> Programatically "Click" the button!!! <=== !!!
		; Msg := WM_COMMAND, wParam_hi := 0, wParam_lo := idCommand, lParam := control := hTb
		; ; DllCall('SendMessage', 'UInt', hTb, 'UInt', Msg, 'UInt', wParam_hi | wParam_lo, 'UIntP', lParam)
		; SendMessage(Msg, wParam_hi | wParam_lo, lParam, , hTb)
		; --------------------------------------------------------------------------------
		; Step: Restore previous and set delay
		; --------------------------------------------------------------------------------
		SetControlDelay(pCD), SetMouseDelay(pMD), SetWinDelay(pWD)
		; --------------------------------------------------------------------------------
		BlockInput(0) ; 1 = On, 0 = Off
		; --------------------------------------------------------------------------------
	}
	catch
		throw ValueError("The specified toolbar " nCtl " was not found. Please ensure the edit field has been selected and try again.", -1)
	; Return 0
}
; --------------------------------------------------------------------------------
/**
 * @function ControlClick()
 * @note All of these are equivalent as of 09.11.23
 */
HznControlClick(X, Y)
{
; ControlClick("x" ((X+(W/2))*DPIsc) " y" ((Y+(H/2))*DPIsc), hTb,,,, "NA")
; ControlClick("x" (X2*DPIsc) " y" (Y2*DPIsc), hTb,,,, "NA")
; ControlClick("x" X " y" Y, hTb,,,, "NA") ; <===== this one is the best usable
}
; --------------------------------------------------------------------------------
; Callback definition for EVENT_SYSTEM_CAPTURESTART
; msvb_lib_toolbar_created := CallbackCreate(CallBack_TB_CREATED)
; ^+!q::
; EnumWindowsProc(&nCtl:='', &hTb:=0)
; {
; 	nCtl := ''
; 	win_get_controls := WinGetControls('A')
; 	list_controls := DisplayObj(win_get_controls)
; 	OutputDebug(list_controls)
; 	bak_TitleMatchMode := A_TitleMatchMode
; 	OutputDebug(bak_TitleMatchMode)
; 	SetTitleMatchMode('RegEx')
; 	RegExMatch(list_controls, 'msvb_lib_toolbar\d', &m)
; 	OutputDebug(m)
; 	SetTitleMatchMode(bak_TitleMatchMode)
; 	&nCtl := m[]
; 	&hTb := ControlGethWnd(nCtl, "A")
; 	MsgBox(nCtl ' ' hTb . '`n')
; 	return
; }	
; --------------------------------------------------------------------------------
remote_mem_buff(hProcess, Is32bit?, &CTRL_SIZE?)
{
	Static MEM_PHYSICAL := 4 ; 0x04 ; 0x00400000, ; via MSDN Win32
	Static MEM_COMMIT := 4096
	Is32bit := 0
	RPtrSize := Is32bit ? 4 : 8
	CTRL_SIZE := 8 + (RPtrSize * 3) 
	remoteMemory := DllCall("VirtualAllocEx", "Ptr", hProcess, "Ptr", 0, "UPtr", CTRL_SIZE, "UInt", MEM_COMMIT, "UInt", MEM_PHYSICAL, "Ptr")
	return remoteMemory
}
; --------------------------------------------------------------------------------
/**
 * function: Open the target process with PROCESS_VM_OPERATION, PROCESS_VM_READ, and PROCESS_VM_WRITE access
 * @param tpID
 * @param PROCESS_VM_READ
 * 		@param Int32
 * 		@Dec_Value 8
 * 		@Hex_Value 0x00000008 ??? or 0x0018 ???
 * @param PROCESS_VM_READ
 * 		@param Int32
 * 		@Dec_Value 16
 * 		@Hex_Value 0x00000010 ??? or 0x0010
 * @param PROCESS_VM_WRITE
 * 		@param Int32
 * 		@Dec_Value 32
 * 		@Hex_Value 0x00000020 ??? or 0x0020
 * @example hProcess := DllCall('OpenProcess', 'UInt', 0x0018 | 0x0010 | 0x0020, 'Int', 0, 'UInt', tpID, 'Ptr')
 * @returns {hProcess}
 */
_hProcess(tpID)
{
	Static PROCESS_VM_OPERATION := 8, PROCESS_VM_READ := 16, PROCESS_VM_WRITE := 32
	; hProcess := DllCall( 'OpenProcess', 'UInt'
						; , PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE
						; , "Int", 0
						; , "UInt", tpID
						; , "Ptr")
	; return hProcess
	return DllCall( 'OpenProcess', 'UInt', PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, "Int", 0, "UInt", tpID, "Ptr")
}
; --------------------------------------------------------------------------------
; rect := WindowGetRect("window title etc.")
; MsgBox(rect.width "`n" rect.height)
hIntersectRect(l1, t1, r1, b1, l2, t2, r2, b2) {
	rect1 := Buffer(16), rect2 := Buffer(16), rectOut := Buffer(16)
	NumPut("int", l1, "int", t1, "int", r1, "int", b1, rect1)
	NumPut("int", l2, "int", t2, "int", r2, "int", b2, rect2)
	if DllCall("user32\IntersectRect", "Ptr", rectOut, "Ptr", rect1, "Ptr", rect2)
		return {l:NumGet(rectOut, 0, "Int"), t:NumGet(rectOut, 4, "Int"), r:NumGet(rectOut, 8, "Int"), b:NumGet(rectOut, 12, "Int")}
}
WindowGetRect(hwnd) {
    try {
        rect := Buffer(64, 0) ; V1toV2: if 'rect' is a UTF-16 string, use 'VarSetStrCapacity(&rect, 16)'
        DllCall("GetClientRect", "Ptr", hwnd, "Ptr", rect)
        ; return {width: NumGet(rect, 8, "Int"), height: NumGet(rect, 12, "Int")}
		return {Left: NumGet(rect, 0, "Int"), Top: NumGet(rect, 4, "Int"), Right: NumGet(rect, 8, "Int"), Bottom: NumGet(rect, 12, "Int"), Height: ((NumGet(rect, 12, "Int")) - (NumGet(rect, 4, "Int"))), Width: ((NumGet(rect, 8, "Int") - (NumGet(rect, 0, "Int")))) }
    }
}
GetRect(hwnd, &RC := "", *) { ; Retrieves the rich edit control's formatting rectangle
; Returns an object with keys L (eft), T (op), R (ight), and B (ottom).
; If a variable is passed in the Rect parameter, the complete RECT structure will be stored in it.
RC := Buffer(32, 0)
; SendMessage(0x00B2, 0, RC.Ptr, HWND)
SendMessage(178, 0, RC.Ptr, HWND)
; Return {L: NumGet(RC, 0, "Int"), T: NumGet(RC, 4, "Int"), R: NumGet(RC, 8, "Int"), B: NumGet(RC, 12, "Int")}
Return {Left: NumGet(RC, 0, "Int"), Top: NumGet(RC, 4, "Int"), Right: NumGet(RC, 8, "Int"), Bottom: NumGet(RC, 12, "Int"), Height: ((NumGet(RC, 12, "Int")) - (NumGet(RC, 4, "Int"))), Width: ((NumGet(RC, 8, "Int") - (NumGet(RC, 0, "Int")))) }
}
GetClientSize(hCtl)
{
	BtnStructSize := 32
	; rc := Buffer(BtnStructSize, 0)
	rc := Buffer(BtnStructSize, 16)
	; Buffer(rc:=0,16)
    ; DllCall("GetClientRect", "Ptr", hCtl, "Ptr", rc)
	tbRECT := GETITEMRECT(hCtl)
	l:=tbRECT.Left
	b:=tbRECT.Bottom
	; SendMessage(Msg, wParam, lParam,, hTb)
	; l := NumGet(rc, 0, "Int")
	; t := NumGet(rc, 4, "Int")
    ; r := NumGet(rc, 8, "int")
    ; b := NumGet(rc, 12, "int")
	; w := l - r
	; h := b-t

	; MsgBox('w:' . w . ' ' . 'h: ' . h . '`n')
	; return {Left:l,Top:t,Right:r,Bottom:b,Width:w, Height:h}
	return {Left:l,Bottom:b}
}
GetClientRect(hCtl, hCtl_title){
	
	mTX11 := Map()
	dpi_hzn := DPI.MonitorFromWindow('hznHorizon.exe')
	Info('dpi_hzn: ' . dpi_hzn, 30000)
	RECT := Buffer(32, 0)
	DllCall("MapWindowPoints", "Ptr", hCtl, "Ptr", 0, "Ptr", RECT, "UInt", 2)
	Left:= NumGet(RECT, 0, "Int")
	Top:= NumGet(RECT, 4, "Int"), 
	Right:= NumGet(RECT, 8, "Int")
	Bottom:= NumGet(RECT, 12, "Int")
	Width := (Left + (Right/2))
	Height := (Top + (Bottom/2))
	mTX11.Set('Left', Left, 'Right', Right, 'Top', Top, 'Bottom', Bottom, 'height', height, 'width', width)
	; Info(
	; 	'Left: ' . Left . ' ' . 'Right: ' . Right
	; 	. '`n'
	; 	'Top: ' . Top . '`n' . 'Bottom: ' . Bottom
	; 	. '`n'
	; 	'Height: ' . Height . '`n' . 'Width: ' . Width
	; 	, 30000)
	return mTX11
}
; fix
^+7::somestufffortoolbar()

somestufffortoolbar(idButton?, hTb?){
	buttons := hznButtonCount()
    hTb := HznToolbar._hTb()
    hTx := ControlGetFocus('A')
    pID := HznToolbar._pID()
    tpID := HznToolbar._tpID()
	hTx := ControlGetFocus('A')
	Text := Buffer(128, 0)
	RECT := Buffer(32, 0)
	a_idButton := []
	aStrings := []
	a_text := []
	a_Rect := []
	vButton := ''
	idButton := ''
	vText := ''
	a_bText := ''
	b_bText := ''
	dText := ''
	vRect := ''
	text := ''
    TB_GETBUTTON := 1047
	Static Msg := TB_GETITEMRECT := 1053
	; Static Msg := EM_GETRECT := 178
	static wParam := 1
	Static PROCESS_VM_OPERATION := 8, PROCESS_VM_READ := 16, PROCESS_VM_WRITE := 32
	hCtl := hTb
	DllCall("GetWindowThreadProcessId", "Ptr", hCtl, "UInt*", &tpID := 0)
	hProcess := DllCall( 'OpenProcess', 'UInt', 8 | 16 | 32, "Int", 0, "UInt", tpID, "Ptr")
	Static MEM_PHYSICAL := 4 ; 0x04 ; 0x00400000, ; via MSDN Win32
	Static MEM_COMMIT := 4096
	Is32bit := 0, RPtrSize := Is32bit ? 4 : 8, CTRL_SIZE := 8 + (RPtrSize * 3) 
	remoteMemory := DllCall("VirtualAllocEx", "Ptr", hProcess, "Ptr", 1, "UPtr", CTRL_SIZE, "UInt", MEM_COMMIT, "UInt", MEM_PHYSICAL, "Ptr")
	; remoteMemory := remote_mem_buff(hProcess, , &CTRL_SIZE)
	lParam := remoteMemory
	SendMessage(Msg, wParam, lParam, hCtl, hCtl)
	RECT := Buffer(CTRL_SIZE, 0)
	; Note : Winapi TBBUTTON struct(32 bytes on x64, 20 bytes on x86)
	Is32bit := Win32_64_Bit(remoteMemory)
	BtnStructSize := Is32bit ? 20 : 32
	; RECT := BtnStruct := Buffer(BtnStructSize, 0)
	; BtnStruct := Buffer(BtnStructSize, 0)
	; --------------------------------------------------------------------------------
	; * Read the button information stored in the RECT (remoteMemory)
	DllCall("ReadProcessMemory", "Ptr", hProcess, "Ptr", remoteMemory, "Ptr", RECT, "UPtr", BtnStructSize, "UInt*", &bytesRead:=32, "Int")
	Loop buttons {
		button := (A_Index + 99)
		idButton := String(button)
		; Get_Button := _GETBUTTON(idButton)
		hProcess := DllCall('OpenProcess', 'UInt', 8 | 16 | 32, "Int", 0, "UInt", tpID, "Ptr")
		remoteMemory := remote_buffer := remote_mem_buff(hProcess,0,&TBBUTTON_SIZE)
		GETBUTTON := SendMessage(TB_GETBUTTON, idButton, remoteMemory, hTb, hTb)

		a_nText 	:= NumGet(Text, 0, 	"UInt")
		b_nText 	:= NumGet(Text, 4, 	"UInt")
		c_nText 	:= NumGet(Text, 8, 	"Int")
		d_nText 	:= NumGet(Text, 12, "Int")
		fsState 	:= NumGet(Text, 16, "UChar")
		fsStyle 	:= NumGet(Text, 17, "UChar")
		cxWORD 		:= NumGet(Text, 18, "UShort")

		Info(
			'[' button . '] ' 
			; '1 ' . txt_strcap . ' ' 
			; '2 ' . txt_buff . ' ' 
			'3 ' . a_bText   . ' ' 
			'4 ' . b_bText . ' ' 
			'5 ' . vText . ' ' 
			'6 ' . a_nText . ' '
			'7 ' . b_nText . ' ' 
			'8 ' . c_nText . ' '
			'9 ' . d_nText . ' ' 
			'10 ' . fsState . ' '
			'11 ' . fsStyle . ' '	
			'12 ' . cxWORD . ' '
			, 30000
		)
		; dText .= '[' . button . '] ' . 'a_bText: ' . a_bText . ' ' 
		; . 'b_bText: ' . b_bText . ' ' . 'vText: ' . vText . ' ' . 'a_nText: ' 
		; . a_nText . '`n' . 'b_nText: ' . b_nText . '`n' 
		; . 'c_nText: ' . c_nText . ' ' . 'd_nText: ' . d_nText . '`n'
		; dText := 'idButton: ' . button . '`n' . 'a_bText: ' . a_bText . '`n' . 'b_bText: ' . b_bText . '`n' . 'vText: ' . vText . '`n' . 'a_nText: ' 
		; ; . a_nText . '`n' . 'b_nText: ' . b_nText . '`n' 
		; . 'c_nText: ' . c_nText . '`n' . 'd_nText: ' . d_nText . '`n'
	}
	; Info(dText, 30000)
	; for each, idButton in a_idButton {
	; 	a_text.Push(text)
		; --------------------------------------------------------------------------------
	; SendMessage(0x433, idButton, RECT, hTx , hTx)	; TB_GETRECT
	DllCall("MapWindowPoints", "Ptr", hTx, "Ptr", 0, "Ptr", RECT, "UInt", 2)
	Left := NumGet(RECT, 0, "Int")
	Bottom := NumGet(RECT, 12, "Int")
	Info('Left: ' . Left . '`n' . 'Bottom: ' . Bottom, 30000)
	; 	a_Rect.Push(Left)
	; 	a_Rect.Push(Bottom)
	; }
	; for each, vButton in a_idButton {
	; 	idButton .= vButton . '`n'
	; 	for each, vText in a_text {
	; 		bText .= vText . '`n'
	; 	}
	; }
	; for each, vRect in a_Rect {
	; 	bRect := ''
	; 	bRect .= 'L: ' Left . ' ' . 'B: ' Bottom 
	; }
	; info(
	; 	'Buttond ID: ' 	. idButton 	. '`n'
	; 	'Text: ' 		. text 		. '`n'
	; 	'Left: ' 		. Left 		. '`n'
	; 	'Bottom: ' 		. Bottom 	. '`n'
	; 	, 30000
	; )
}
_GETBUTTON(n:=1, hTb?, pID?, hProcess?)
{
	Static 	TB_GETBUTTON := 1047 ; hex = 0x417
	OutputDebug('n: ' n '`n')
	; pID := WinGetPID(hTb)
	pID := HznToolbar._pID()
	tpID := HznToolbar._tpID()
	hTb := HznToolbar._hTb()
	hProcess := _hProcess(tpID)
	; hProcess = 0 ? hProcess := hp_Remote(pID) : hpRemote := hpRemote
	remoteMemory := remote_buffer := remote_mem_buff(hProcess,0,&TBBUTTON_SIZE)
	GETBUTTON := SendMessage(TB_GETBUTTON, n-1, remoteMemory, hTb, hTb)
	; MsgBox(GETBUTTON) ; ===> displays a zero (0)
	return GETBUTTON
}
GETBUTTONINFO(hTb?){
	hTb := HznToolbar._hTb()
	; wParam := idButton, lParam := struct(), 
	; SendMessage(0x43F, wParam, lParam, , hTb) ; TB_GETBUTTONINFOW
	;;TBBUTTONINFO=48:32

	; typedef struct {
	; 	0: 0,
	; 	"UInt" UINT cbSize ;
	; 	4: 4,
	; 	"UInt" DWORD dwMask ;
	; 	8: 8,
	; 	"Int" int idCommand ;
	; 	12: 12,
	; 	"Int" int iImage ;
	; 	16: 16,
	; 	"UChar" BYTE fsState ;
	; 	17: 17,
	; 	"UChar" BYTE fsStyle ;
	; 	18: 18,
	; 	"UShort" WORD cx ;
	; 	24: 20,
	; 	"UPtr" DWORD_PTR lParam ;
	; 	32: 24,
	; 	"Ptr" LPTSTR pszText ;
	; 	40: 28,
	; 	"Int" int cchText ;
	; } TBBUTTONINFO, *LPTBBUTTONINFO ;
	; 48: 32
}
GETBUTTONSTATE(idButton, hTb := HznToolbar._hTb())
{
	Static TB_GETSTATE := 1042 ; 0x0412
	; hTb := HznToolbar._hTb()
	btnCt := hznButtonCount()
	; Infos(btnCt)
	Msg := TB_GETSTATE, wParam := idButton, lParam := 0, control := hTb
	GETSTATE := SendMessage(TB_GETSTATE, wParam, lParam, hTb, hTb)
	btnstate := SubStr(GETSTATE,1,1)
	btnname := idButton = 100 ? 'Bold' : idButton = 101 ? 'Italic' : idButton = 102 ? 'Underline' : ''
	; If (btnstate = 4) || (btnstate = 6) || (btnstate = 1)
	; 	{
	; 	return btnstate
	; 	}
	; MsgBox(   'The ' btnname
	; 		. ' button is not available.' '`n'
	; 		. 'idButton: ' idButton '`n'
	; 		. 'btnstate: ' btnstate)
	; OutputDebug("btnstate: " . btnstate . "`n")
	return btnstate
}
; --------------------------------------------------------------------------------
/**
 * @function COMMANDBUTTON()
 * @param Msg := WM_COMMAND
 * @param wParam_hi := control defined notification code => := 0 or !btnstate (opposite of current position => 4 || 6)
 * @param wParam_lo := control identifier => idCommand from above
 * @param lParam 	:= handle to the control => hTb
 */
COMMANDBUTTON(idCommand, hTb, btnstate:=0)
{
	Static WM_COMMAND := 273 ; hex = 0x111
	Static Msg := WM_COMMAND, wParam_hi := !btnstate , wParam_lo := idCommand, lParam := control := hTb := hTb
	; --------------------------------------------------------------------------------
	SendMessage(Msg, wParam_hi | wParam_lo,lParam,, hTb)
	; --------------------------------------------------------------------------------
	; if(btnstate = 4) ? SendMessage(TB_SETSTATE, idCommand, 6 | 0, hTb, hTb) : SendMessage(TB_SETSTATE, idCommand, 4 | 0, hTb, hTb)
	; WM_NCLBUTTONDOWN := 0x00A1
	; WM_NCLBUTTONUP := 0x00A2
	; SendMessage(WM_LBUTTONDOWN,,X1|Y1,hTb,hTb)
	; Sleep(100)
	; SendMessage(WM_LBUTTONUP,,X1|Y1,hTb,hTb)
}
; --------------------------------------------------------------------------------
Win32_64_Bit(hpRemote)
; Win32_64_Bit(hpRemote, &Is32:='False')
{
	A_Is64bitOS ? DllCall("IsWow64Process", "Ptr", hpRemote, "Int*", Is32bit := 0) : DllCall("IsWow64Process", "Ptr", hpRemote, "Int*", Is32bit := 1)
	; If (A_Is64bitOS)
	; 	{
	; 		Static Is32bit := 0
	; 	Try
	; 		DllCall("IsWow64Process", "Ptr", hpRemote, "Int*", Is32bit)
	; 	catch
	; 		Is32bit := 1
		Is32bit = 0 ? Is32:='False' : is32:='True'
	OutputDebug("Is32bit: " Is32 "`n")
	return Is32bit
}
; --------------------------------------------------------------------------------
; /*
HznDPI(hCtl?,hCtl_title?, &arHznDPI:='', &DPIsc:=0)
{
	hCtl := ControlGetFocus('A')
	hCtl_title := WinGetTitle(hCtl)
	arHznDPI := Array()
	try {
		nmHwnd 	:= GetNearestMonitorInfo(hCtl, hCtl_title).Handle
		nmName 	:= GetNearestMonitorInfo(hCtl, hCtl_title).Name
		nmNum 	:= GetNearestMonitorInfo(hCtl, hCtl_title).Number
		nmPri 	:= GetNearestMonitorInfo(hCtl, hCtl_title).Primary
		mDPIx 	:= GetNearestMonitorInfo(hCtl, hCtl_title).x
		mDPIy 	:= GetNearestMonitorInfo(hCtl, hCtl_title).y
		mDPIw 	:= GetNearestMonitorInfo(hCtl, hCtl_title).WinDPI
		DPImw 	:= DPI.GetForWindow(hCtl, hCtl_title)
		DPIsc 	:= DPI.GetScaleFactor(DPImw) ; <====== this one
		; DPIsc1 	:= DPI.GetScaleFactor(mDPIx)
	} catch Error as e {
		; OutputDebug('nmHwnd: '	nmHwnd '`n'
		; 		.	'nmName: '	nmName '`n'
		; 		.	'nmNum: '	nmNum '`n'
		; 		.	'nmPri: '	nmPri '`n'
		; 		.	'mDPIx: '	mDPIx '`n'
		; 		.	'mDPIy: '	mDPIy '`n'
		; 		.	'mDPIw: '	mDPIw '`n'
		; 		.	'DPImw: '	DPImw '`n'
		; 		.	'DPIsc: '	DPIsc '`n'
		; 		; .	'DPIsc1: '	DPIsc1 '`n'
		; 		.	'PriDPI: '	A_ScreenDPI '`n'
		; 		)		
		throw e		
	}
	arHznDPI.Push(
		{  	  Number: '(' nmNum ')'
			, 1_Handle:nmHwnd
			, 2_Name:nmName
			, 3_x:mDPIx
			, 4_y:mDPIy
			, 5_WinDPI:mDPIw
			, 6_MainWinDPI:DPImw
			, 7_ScreenDPI:DPIsc
		})
	; --------------------------------------------------------------------------------
	return {arHznDPI:arHznDPI,DPIsc:DPIsc}
}
; */
; --------------------------------------------------------------------------------
; TODO need to validate that this works, but not high priority
/**
 * function .: for use in a button call that requires ControlCLick() and DPI adjustments
 * @description Get the bounds of each button (Get Item Rectangle)
 * @param GETITEMRECT( hProcess,n,remoteMemory,hTb,TBBUTTON_SIZE,Is32bit,&RECT,&BtnStructSize,&BtnStruct,&bytesRead,&Left,&Top,&Right,&Bottom, &X, &Y)
*/	
; GETITEMRECT(hProcess, n,remoteMemory,hTb, TBBUTTON_SIZE, Is32bit, &RECT, &BtnStructSize, &BtnStruct, &bytesRead, &Left, &Top, &Right, &Bottom, &X, &Y)
GETITEMRECT(hCtl)
{
	; wParam := n, lParam := remoteMemory, control := ''
	Static Msg := TB_GETITEMRECT := 1053
	RECT := Buffer(32, 0)
	DllCall("GetClientRect", "Ptr", hCtl, "Ptr", 0, "Ptr", RECT, "UInt", 2)
	Left := NumGet(RECT, 0, "Int")
	Bottom := NumGet(RECT, 12, "Int")
	Info('Left: ' . Left . '`n' . 'Bottom: ' . Bottom, 30000)
	Return {Left:Left, Bottom:Bottom}
	; --------------------------------------------------------------------------------
	; Note: Updated 09.11.23
	; hCtl_title := WinGetTitle(hCtl)
	; DPIsc := HznDPI(hCtl,hCtl_title).DPIsc
	; X1 			:= Left
	; Y1 			:= Top
	; Left 		:= X1
	; Top			:= Y1
	; W1 			:= Right-Left
	; H1 			:= Bottom-Top
	; ; W 			:= W1/2
	; W 			:= W1
	; ; H 			:= H1/2
	; H 			:= H1
	; X2 			:= X1+W
	; Y2 			:= Y1+H
	; dpiX 		:= X2*=DPIsc
	; dpiY 		:= Y2*=DPIsc
	; dpiWidth	:= Round((W*=DPIsc),0)
	; dpiHeight 	:= Round((H*=DPIsc),0)
	; --------------------------------------------------------------------------------
	; OutputDebug(
	; MsgBox(	
	; 	'X1:' X1 . ' ' . 'Y1:' . Y1 .  ' ' . 'W1:' . W1 . ' ' . 'H1:' . H1 . '`n'
	; 	. 'W:' . W . " " . 'H:' . H . '`n'
	; 	. 'X2:' X2 . ' ' . 'Y2:' . Y2 . '`n'
	; 	. 'X:' . X . ' ' . 'Y:' . Y . '`n'
	; 	)
}
; --------------------------------------------------------------------------------
/**
 * Installs the script to the user startup folder
 * @function Check_Startup_Status()
 * @param hWndToolbar - The handle of the toolbar control.
 * @param n - The index of the toolbar item to click (1-based). Note: Separators are considered items as well.
 * @returns {void} 
 */
; --------------------------------------------------------------------------------
Check_Startup_Status(){
arrStartup := Array(
	startup := A_Startup '\',
	script := StrSplit(A_Startup, '.'),
	scriptnoext := script[1],
	scriptlnk := script[1] . '.lnk',
	scriptSC := script[1] . script[2] . ' - Shortcut',
	)

Startup_Shortcut := arrStartup() 
; Startup_Shortcut := A_Startup "\" A_ScriptName
If (!FileExist(Startup_Shortcut))
    {
        
        myGui := Gui(,"Add Script to Startup Folder",)
        myGui.Opt("AlwaysOnTop")
        myGui.SetFont("s12")
        myGui.Add("Text",, "It is recommended you add this script to your Startup folder so`nthat it is active every time your computer starts.`n`nMake sure that the script is in the location you want to save`nit in before adding it to the startup folder.`n`nWould you like to add this script to the startup folder now?")
        myGui.Add("Button","x15 +default", "Add to Startup").OnEvent("Click", ClickedAdd)
        myGui.Add("Button","x+m", "Cancel").OnEvent("Click", ClickedCancel)
        myGui.Show("w500")
        
        ClickedAdd(*)
        {
            myGui.Destroy()
            FileCreateShortcut(A_ScriptDir "\" A_ScriptName, Startup_Shortcut)
            Run( "shell:startup")
            MsgBox("Shortcut added to your Startup folder at:`n`n" Startup_Shortcut "`n`nYour Startup folder has been opened for you. Please delete any old shortcuts.")
            Return
        }

        ClickedCancel(*)
        {
            myGui.Destroy()
            MsgBox("Please move the file to the location you want to save it, then run it again.")
            Return
        }
    } 
Else 
    {
		; MsgBox("Shortcut Exists") ; Test message for debugging. Leave commented out for normal operation.
        Return
    }
}
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
/************************************************************************
 * @function Create_HznHorizon_ico()
 * function: Create the .ico file for use in as the Tray Icon
 * @author 2.0.00.00 - 2023-07-31 - iPhilip - converted script to AutoHotkey v2
 * @param B64 - A picture converted to a string and encoded via b64
 * @source ;-> https://www.autohotkey.com/boards/viewtopic.php?f=83&t=119966
***********************************************************************/
; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_HznHorizon_ico(NewHandle := False) {
	Static hBitmap := 0
	If (NewHandle)
		hBitmap := 0
	If (hBitmap)
		Return hBitmap
	B64 := "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAgBAAAGAbAABgGwAAAAAAAAAAAAD7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////////////////////////////////////f/////RERE//tcQv/7XEL////////////////////////////////////////////////3/////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL///////////////////////////f/////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC////////////////////////////////90RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC////////////////////////////////////////////////////////////////////////////////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL///////////////////////////////////////////f////3////9/////f///////////////////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC/////////////////0RERP/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/////////////////////////////////RERE//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//////////////////////////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC////////////////////////////////////////////RERE//tcQv/7XEL/+1xC//tcQv////////////////////////////////////////////////9ERET/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/+1xC//tcQv/7XEL/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", StrPtr(B64), "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", &DecLen := 0, "Ptr", 0, "Ptr", 0)
		Return False
	Dec := Buffer(DecLen, 0)
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", StrPtr(B64), "UInt", 0, "UInt", 0x01, "Ptr", Dec, "UIntP", &DecLen, "Ptr", 0, "Ptr", 0)
		Return False
	hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
	pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
	DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", Dec, "UPtr", DecLen)
	DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
	DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream := ComValue(13, 0))
	hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
	SI := Buffer(8 + 2 * A_PtrSize, 0), NumPut("UInt", 1, SI, 0)
	DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", &pToken := 0, "Ptr", SI, "Ptr", 0)
	DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", &pBitmap := 0)
	DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", &hBitmap := 0)
	DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
	DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
	DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
	Return hBitmap
	}
; --------------------------------------------------------------------------------
/*
@resource ...: https://www.autohotkey.com/boards/viewtopic.php?t=73851
@ahkversion .: v1
*/
; CreateTrayMenu()
; addTrayMenuOption(label := "", command := "")
; {
;   Static Tray
;   Tray:= A_TrayMenu
; 	if (label = "" && command = "") {
; 		Tray.Add()
; 	} else {
; 		Tray.Add(label, command)
; 	}
; }
; shortcut(){
;     FileExist(Startup_Shortcut) ? ShortcutExist() : NoShortcut()
; }
; ShortcutExist(){
;     Tray.Check("Run at startup")
; }

; NoShortcut(){
;     Tray.Uncheck(Run at startup)
; }
; CreateTrayMenu()
; {
;     ;Menu, Tray, Icon, % hICON ; this changes the icon into a little Horizon thing.
;     ;Tray.Delete() ; V1toV2: not 100% replacement of NoStandard, Only if NoStandard is used at the beginning
;     addTrayMenuOption("Made with nerd by Adam Bacon and Terry Keatts", "madeBy")
;     addTrayMenuOption()
;     addTrayMenuOption("Run at startup", "runAtStartup")
;     shortcut()
;     ; update the tray menu status on startup
;     addTrayMenuOption()
;     Tray.AddStandard()
; }
; runAtStartup() 
; {
;     FileExist(startup_shortcut) ? rmstartup() : addstartup()
; }

; addstartup(){
;     FileCreateShortcut(A_ScriptName, A_Startup . "\" . A_ScriptName . ".lnk", icostr)
;     Tray.Check(Run at startup) ; update the tray menu status on change
;     trayNotify("Startup shortcut added", "This script will now automatically run when your turn on your computer.")
; }

; rmstartup(){
;     FileDelete(startup_shortcut)
;     Tray.Uncheck(Run at startup) ; update the tray menu status on change
;     trayNotify("Startup shortcut removed", "This script will not run when you turn on your computer.")
; }

; madeBy()
; {
; 	MsgBox("This was made by nerds, for nerds. Regular people are ok too, lol.`nModified by OvercastBTC")
; }
; HideTrayTip()
; {
;     (SubStr(A_OSVersion, 1, 3) = "10.")   ? IconTrayTip()
;                                         : TrayTip
; }

; IconTrayTip()
; {
;     Tray.NoIcon()
;     Sleep(200)
;     Tray.Icon()
; }

; trayNotify(title, message, options := 0)
; {
;     TrayTip(title, message, options)
; }