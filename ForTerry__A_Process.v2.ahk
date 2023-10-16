#Requires Autohotkey v2+
#Include <Directives\__AE.v2>
#Include <HznPlus.v2>
Persistent(1) ; Keeps script permanently running
CoordMode("ToolTip", "Screen")
; ProcessSetPriority('High')
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
ICON := '.\ICON\DllCall.ico'
TraySetIcon(ICON) ; this changes the icon into a little dllcall thing.
; --------------------------------------------------------------------------------
#HotIf WinActive('ahk_exe hznhorizon.exe')
global gOpenWindows := Map()
global initial_controls_map := Map()
initial_list_ClassNN := []
initial_list_ClassNN_hwnd := []
WinGetListList := []

hwnd := WinWaitActive('ahk_exe hznhorizon.exe')
process_name := WinGetProcessName(hwnd)
try {
	GetAncestor_hwnd := DllCall("GetAncestor", "UInt", hwnd, "Uint", 2)
	GetAncestor_process_name := WinGetProcessName(GetAncestor_hwnd)
	if (GetAncestor_process_name == process_name) && (GetAncestor_hwnd == hwnd){
		OutputDebug('It`'s Horizon: ' . process_name . ' (' . hwnd . ')`n')
		initial_list_ClassNN := WinGetControls('A')
		for each, value in initial_list_ClassNN {
			try {
				initial_list_ClassNN_hwnd := ControlGetHwnd(initial_list_ClassNN, 'A')
				initial_controls_map.Push({initial_list_ClassNN:initial_list_ClassNN_hwnd})
				OutputDebug('test' . {initial_list_ClassNN:initial_list_ClassNN_hwnd})
			}
		}
	} else {
		OutputDebug(GetAncestor_process_name . ' (' . GetAncestor_hwnd . ')`n')
		WinGetListList := WinGetList(hwnd)
		for each, value in WinGetListList
			OutputDebug('value:' . value . '`n')
	}
}


; If (WinActive('A') ~= 'i).*h.*z.*n.*'){
; 	OutputDebug('First Try: ' . 'True' . '`n')
; }
; for hWinGetList in WinGetList() { ; TODO this is way way cooler and a more extensive list of stuff
for hWinGetList in WinGetList(hwnd) {
	SendLevel(((A_SendLevel < 100) ? (A_SendLevel + 1):(A_SendLevel)))
	BlockInput(1)
	pvTxt := A_DetectHiddenText, DetectHiddenText(1)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(1)
	txneedle := 'i)\wt.*x.*'
	try {
		gOpenWindows[hWinGetList] := ControlGetClassNN(hWinGetList)
		tryWinGetList := gOpenWindows[hWinGetList] := ControlGetClassNN(hWinGetList)
		OutputDebug('Pre(' . A_Index . ') ' . gOpenWindows[hWinGetList] . '`n') ; cool, but not useful
		; if ((gOpenWindows[hWinGetList] ~= 'i)t.*x.*'))
		if ((tryWinGetList ~= txneedle))
			OutputDebug('Post(' . A_Index . ') ' . gOpenWindows[hWinGetList]) ; cool, but not useful
	} catch {
		gOpenWindows[hWinGetList] := DllCall("GetAncestor", "UInt", hWinGetList, "Uint", 2)
		while (gOpenWindows[hWinGetList] = 'Default IME'){
			return
		}
		OutputDebug('Post1(' . A_Index . ') ' . gOpenWindows[hWinGetList] '`n') ; cool, but not useful
		while (gOpenWindows[hWinGetList] = 'Default IME'){
			return
		}
		gOpenWindows[hWinGetList] := WinGetTitle(hWinGetList)
		OutputDebug('Post2(' . A_Index . ') ' . gOpenWindows[hWinGetList] '`n') ; cool, but not useful
		while (gOpenWindows[hWinGetList] = 'Default IME'){
			return
		}
	} finally {
		gOpenWindows[hWinGetList] := WinGetClass(hWinGetList)
		OutputDebug('Post3(' . A_Index . ') ' . gOpenWindows[hWinGetList] '`n') ; cool, but not useful
	}
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	BlockInput(0)
	SendLevel(0)
}
hook := SetWinEventHook(HandleWinEvent)

SetWinEventHook(callbackFunc, winTitle := 0, eventMin := 0x8000, eventMax := 0x8001, flags := 0x0002) {
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel) : (A_SendLevel + 1)))
	BlockInput(1)
	pvTxt := A_DetectHiddenText, DetectHiddenText(1)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(1)
	EVENT_OBJECT_ACCELERATORCHANGE := 0x8012, 	EVENT_OBJECT_STATECHANGE := 32778, 
	EVENT_OBJECT_CLOAKED := 0x8017, 			EVENT_OBJECT_CONTENTSCROLLED := 0x8015, 
	EVENT_OBJECT_CONTENTSCROLLED := 0x8015, 	EVENT_OBJECT_DEFACTIONCHANGE := 0x8011, 
	EVENT_OBJECT_DESCRIPTIONCHANGE := 0x800D, 	EVENT_OBJECT_CREATE := 32768, 
	EVENT_OBJECT_DESTROY := 32769, 				EVENT_OBJECT_DRAGSTART := 0x8021, 
	EVENT_OBJECT_DRAGCANCEL := 0x8022, 			EVENT_OBJECT_DRAGCOMPLETE := 0x8023, 
	EVENT_OBJECT_DRAGENTER := 0x8024, 			EVENT_OBJECT_DRAGLEAVE := 0x8025, 
	EVENT_OBJECT_DRAGDROPPED := 0x8026, 		EVENT_OBJECT_END := 0x80FF, 
	EVENT_OBJECT_FOCUS := 0x8005, 				EVENT_OBJECT_HELPCHANGE := 0x8010, 
	EVENT_OBJECT_HIDE := 0x8003, 				EVENT_OBJECT_HOSTEDOBJECTSINVALIDATED := 0x8020, 
	EVENT_OBJECT_IME_HIDE := 0x8028, 			EVENT_OBJECT_IME_SHOW := 0x8027, 
	EVENT_OBJECT_IME_CHANGE := 0x8029, 			EVENT_OBJECT_INVOKED := 0x8013, 			EVENT_OBJECT_LIVEREGIONCHANGED := 0x8019, 	EVENT_OBJECT_LOCATIONCHANGE := 0x800B, 		
	EVENT_OBJECT_NAMECHANGE := 0x800C, 			EVENT_OBJECT_PARENTCHANGE := 0x800F, 
	EVENT_OBJECT_REORDER := 0x8004, 			EVENT_OBJECT_SELECTION := 0x8006, 
	EVENT_OBJECT_SELECTIONADD := 0x8007, 		EVENT_OBJECT_SELECTIONREMOVE := 0x8008, 	
	EVENT_OBJECT_SELECTIONWITHIN := 0x8009, 	EVENT_OBJECT_SHOW := 0x8002, 
	EVENT_OBJECT_TEXTEDIT_CONVERSIONTARGETCHANGED := 0x8030, 
	EVENT_OBJECT_TEXTSELECTIONCHANGED := 0x8014, 
	EVENT_OBJECT_UNCLOAKED := 0x8018, 
	EVENT_OBJECT_VALUECHANGE := 0x800E, 		EVENT_SYSTEM_ALERT := 0x0002, 
	EVENT_SYSTEM_ARRANGMENTPREVIEW := 0x8016, 	EVENT_SYSTEM_CAPTUREEND := 0x0009, 
	EVENT_SYSTEM_CAPTURESTART := 0x0008, 		EVENT_SYSTEM_CONTEXTHELPEND := 0x000D, 
	EVENT_SYSTEM_CONTEXTHELPSTART := 0x000C, 	EVENT_SYSTEM_DESKTOPSWITCH := 0x0020, 
	EVENT_SYSTEM_DIALOGEND := 0x0011, 			EVENT_SYSTEM_DIALOGSTART := 0x0010, 
	EVENT_SYSTEM_DRAGDROPEND := 0x000F, 		EVENT_SYSTEM_DRAGDROPSTART := 0x000E, 
	EVENT_SYSTEM_END := 0x00FF, 				EVENT_SYSTEM_FOREGROUND := 0x0003, 
	EVENT_SYSTEM_MENUSTART := 0x0004, 			EVENT_SYSTEM_MENUEND := 0x0005, 
	EVENT_SYSTEM_MENUPOPUPSTART := 0x0006, 		EVENT_SYSTEM_MENUPOPUPEND := 0x0007, 
	EVENT_SYSTEM_MINIMIZEEND := 0x0017, 		EVENT_SYSTEM_MINIMIZESTART := 0x0016, 
	EVENT_SYSTEM_MOVESIZEEND := 0x000B, 		EVENT_SYSTEM_MOVESIZESTART := 0x000A, 
	EVENT_SYSTEM_SCROLLINGEND := 0x0013, 		EVENT_SYSTEM_SCROLLINGSTART := 0x0012, 
	EVENT_SYSTEM_SOUND := 0x0001, 				EVENT_SYSTEM_SWITCHEND := 0x0015, 
	EVENT_SYSTEM_SWITCHSTART := 0x0014
	local PID := 0, callbackFuncType, 
	hook := { winTitle: winTitle, flags: flags, eventMin: eventMin, eventMax: eventMax, threadId: 0 }
	callbackFuncType := Type(callbackFunc)
	if !(callbackFuncType = "Func" || callbackFuncType = "Closure")
		throw ValueError("The callbackFunc argument must be a function", -1)
	if winTitle != 0 {
		if !(hook.winTitle := WinExist(winTitle))
			throw TargetError("Window not found", -1)
		hook.threadId := DllCall("GetWindowThreadProcessId", "Int", hook.winTitle, "UInt*", &PID)
	}
	OutputDebug('PID: ' . PID . '`n')
	hook.hHook := DllCall("SetWinEventHook",
		"UInt", eventMin, 
		"UInt", eventMax, 
		"Ptr", 	0, 
		"Ptr", 	hook.callback := CallbackCreate(callbackFunc, "C Fast", 7), 
		"UInt", hook.PID := PID, "UInt", hook.threadId, "UInt", flags)
	
	hook.DefineProp("__Delete", 
	{ call: (this) => (DllCall("UnhookWinEvent", "Ptr", this.hHook), CallbackFree(this.callback)) })

	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	BlockInput(0)
	SendLevel(0)
	return hook
}
HandleWinEvent(hWinEventHook, event, ghwnd, idObject, idChild, idEventThread, dwmsEventTime) {
	initialSendLevel := A_SendLevel
	SendLevel(((A_SendLevel < 100) && (initialSendLevel >= 1) ? (A_SendLevel):(A_SendLevel + 1)))
	BlockInput(1)
	pvTxt := A_DetectHiddenText, DetectHiddenText(1)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(1)
	static EVENT_OBJECT_CREATE := 32768, EVENT_OBJECT_DESTROY := 32769, OBJID_WINDOW := 0, INDEXID_CONTAINER := 0
	global gOpenWindows
	static hznCtrls_map := Map()
	static hznCtrlsarray_ClassNN := []
	static hznCtrlsarray_hwnd := []
	static hznCtrlsList := []
	static hznmsvbTbClassNN := []
	static hznmsvbTbHwnd := []
	static needle := 'i)m.*bar\d'
	hznCtrls := ''
	matchTbClassNN := ''
	matchTbHwnd := ''
	matchTb := ''
	; needle := 'im)(\w+\sbar\w+)'
	static hwnd := WinExist('ahk_exe hznhorizon.exe')
	static process_name := WinGetProcessName(hwnd)
	win_active := WinActive('A')
	win_active_title := WinGetTitle(win_active)
	OutputDebug('win_active: ' . win_active_title . ' (' . win_active . ')' . '`n')
	if !(process_name ~= 'i).*hznhorizon.*'){
		return
	}
	if (process_name ~= 'i).*hznhorizon.*') && (WinGetProcessName(WinActive('A')) ~= 'i).*hznhorizon.*'){
		; Sleep(300)
		try {
			
			hznCtrlsarray_ClassNN := WinGetControls('A')
			for each, hznCtrlsarray_ClassNN in hznCtrlsarray_ClassNN {
				a_ClassNN := hznCtrlsarray_ClassNN
				hznCtrlsarray_hwnd := WinGetControlsHwnd(a_ClassNN, 'A')
				a_hWnd := hznCtrlsarray_hwnd
				hznCtrls_map[a_ClassNN] := {aClassNN:hznCtrlsarray_ClassNN, aHwnd:a_hWnd}
				; OutputDebug(hznCtrlsarray_ClassNN . A_Space . initial_list_ClassNN_hwnd . '`n')
				OutputDebug(hznCtrls_map[a_ClassNN].aClassNN . ' ' . hznCtrls_map[a_ClassNN].aHwnd .  '`n')
				if (hznCtrls_map.Has(initial_list_ClassNN_hwnd)){
					OutputDebug('already exists`n')
					hznCtrls_map.RemoveAt(A_Index)
					return
				}
			}
		}
		for each, hznCtrls in hznCtrlsarray_ClassNN {
			; a_match := []
			; u_match_c := []
			; u_match_h := []
			; a_match_Class := ''
			; a_match_hWnd := ''
			; a_match_k := ''
			; a_match_v := 0
			; hznCtrlsList.Push(hznCtrls)
			; OutputDebug('hznCtrls: ' . hznCtrls . '`n')
			if (hznCtrls ~= needle) {
				; a_match_hWnd := ControlGetHwnd(hznCtrls, 'A')
				; ahWnd := a_match_hWnd
				; a_match_Class := hznCtrls
				; OutputDebug('hznCtrlsneedle: ' . hznCtrls . A_Space . ahWnd . '`n')
				; u_match_c.Push(a_match_Class) ; ' ' a_match_hWnd)
				; u_match_h.Push(a_match_hWnd) ; ' ' a_match_hWnd)
				; for a_match_k, a_match_v in u_match_c {
				; 	OutputDebug('We have a match! ' . a_match_v . '`n')
				; }
				if !(hznmsvbTbClassNN.Has(hznCtrls)){
					OutputDebug('!hznmsvbTbClassNN.Has(hznCtrls) ' . hznCtrls . '`n')
					hznmsvbTbClassNN.Push(hznCtrls)
					matchTbHwnd := ControlGetHwnd(hznCtrls, 'A')
					hznmsvbTbHwnd.Push(matchTbHwnd)
				} else if (hznmsvbTbClassNN.Has(hznCtrls)) && !(hznmsvbTbHwnd.Has(matchTbHwnd)){
					OutputDebug('!(hznmsvbTbClassNN.Has(hznCtrls)) && !(hznmsvbTbHwnd.Has(matchTbHwnd))' . hznCtrls . '`n')
					matchTbHwnd := ControlGetHwnd(hznCtrls, 'A')
					hznmsvbTbHwnd.Push(matchTbHwnd)
					; continue
				} else {
					return
				}
			}
		}
		; }
	}
	if (event == EVENT_OBJECT_CREATE) {
		try {
			for each, matchTb in hznmsvbTbHwnd {
				if !(hznmsvbTbHwnd.Has(matchTb)){
					try {
						HznEnableButtons(matchTb)
						; for hTb in Enabled_Toolbar_Map
						; 	if !(hTb == matchTb){
						; 		HznEnableButtons(matchTb)
						; 	}
						OutputDebug('Enabled: ' . ControlGetClassNN(matchTb) . ' (' . matchTb . ')' . '`n')
						Enabled_Toolbar_Map := Map()
						Enabled_Toolbar_Map := {ClassNN:ControlGetClassNN(matchTb), hTb:matchTb}
						; hznmsvbTbHwnd.RemoveAt(matchTb)
					} ;catch {
					; 	try {
					; 		Sleep(500)
					; 		cArray := [1, 2, 3, 4, 5]
					; 		value_cArray := ''
					; 		for each, value_cArray in cArray {
					; 			nTB := 'msvb_lib_toolbar' . value_cArray
					; 			hTb := ControlGetHwnd(nTB, 'A')
					; 			HznEnableButtons(hTb)
					; 			OutputDebug('Catch Enabled: ' . nTB . ' (' . hTb . ')' . '`n')
					; 		}
					; 	}
					; }
				} else {
					return
				}
							; OutputDebug(nTB . ' (' . hTb . ')')
						; nTB := 'msvb_lib_toolbar2'
						; hTb := ControlGetHwnd(nTB, 'A')
						; HznEnableButtons(hTb)
						; OutputDebug(nTB . ' (' . hTb . ')')
						; nTB := 'msvb_lib_toolbar3'
						; hTb := ControlGetHwnd(nTB, 'A')
						; HznEnableButtons(hTb)
						; OutputDebug(nTB . ' (' . hTb . ')')
						; nTB := 'msvb_lib_toolbar4'
						; hTb := ControlGetHwnd(nTB, 'A')
						; HznEnableButtons(hTb)
						; OutputDebug(nTB . ' (' . hTb . ')')
						; nTB := 'msvb_lib_toolbar5'
						; hTb := ControlGetHwnd(nTB, 'A')
						; HznEnableButtons(hTb)
						; OutputDebug(nTB . ' (' . hTb . ')')
					}	
		}
	} else if (event = EVENT_OBJECT_DESTROY) {
		; for ghwnd in gOpenWindows{
		; 	if gOpenWindows.Has(ghwnd) {
		; 		; OutputDebug('destroy? ' . gOpenWindows[ghwnd] . '`n')
		; 		if (gOpenWindows[ghwnd] ~= needle) {
		; 			OutputDebug('destroyed: ' . gOpenWindows[ghwnd] . '`n')
		; 			gOpenWindows.Delete(ghwnd)
		; 		}
		; 			; ToolTip("Notepad window destroyed")
		; 		; else if gOpenWindows[ghwnd].title = "Calculator"
		; 		; 	ToolTip("Calculator window destroyed")
		; 		; Delete info about windows that have been destroyed to avoid unnecessary memory usage
		; 	}
		; }
		; for each, matchTb in hznmsvbTbHwnd {
		; 	try {
		; 		OutputDebug('matchTb (destroy): ' . matchTb . '`n')
		; 		hznmsvbTbClassNN.RemoveAt(A_Index)
		; 	}
		; }
		for ghwnd in gOpenWindows{
			try{
				destroy_list_ClassNN := ControlGetClassNN(ghwnd)
				if gOpenWindows.Has(ghwnd) {
					OutputDebug('destroyed: ' destroy_list_ClassNN . ' (' . gOpenWindows[ghwnd] . ')' . '`n')
					; if (gOpenWindows[ghwnd] ~= needle) {
					; 	OutputDebug('destroy? ' . gOpenWindows[ghwnd] . '`n')
					; 	gOpenWindows.Delete(ghwnd)
					; }
					if (destroy_list_ClassNN ~= needle) {
						OutputDebug('destroy? ' . gOpenWindows[ghwnd] . '`n')
						gOpenWindows.Delete(ghwnd)
					}
					; ToolTip("Notepad window destroyed")
					; else if gOpenWindows[ghwnd].title = "Calculator"
					; 	ToolTip("Calculator window destroyed")
					; Delete info about windows that have been destroyed to avoid unnecessary memory usage
				}
			}
		}
	}
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	BlockInput(0)
	SendLevel(0)
	; SetTimer(ToolTip, -3000) ; Remove created ToolTip in 3 seconds
}
ShellHook_AProcess()
{
	; Global HSHELL_RUDEAPPACTIVATED := 32772
	hWnd := WinActive('A')
	DllCall("RegisterShellHookWindow", "UInt", A_ScriptHwnd)
	OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), A_Process_GetInfo)
	; OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), MyFunc)
}
A_Process_GetInfo(event, hwnd,*) {
	; Shell Hook Events:
	HSHELL_WINDOWCREATED := 1, HSHELL_WINDOWDESTROYED := 2, HSHELL_ACTIVATESHELLWINDOW := 3, HSHELL_WINDOWACTIVATED := 4, HSHELL_WINDOWACTIVATED := 32772, HSHELL_GETMINRECT := 5, HSHELL_REDRAW := 6, HSHELL_TASKMAN := 7, HSHELL_LANGUAGE := 8, HSHELL_SYSMENU := 9, HSHELL_ENDTASK := 10, HSHELL_ACCESSIBILITYSTATE := 11, HSHELL_APPCOMMAND := 12, HSHELL_WINDOWREPLACED := 13, HSHELL_WINDOWREPLACING := 14, HSHELL_HIGHBIT := 32768, HSHELL_RUDEAPPACTIVATED := 32772, HSHELL_FLASH := 32774
	; 32768 = 0x8000 = HSHELL_HIGHBIT := 32768
	; 32772 = 0x8000 + 4 = 0x8004 = HSHELL_RUDEAPPACTIVATED := 32772 (HSHELL_HIGHBIT + HSHELL_WINDOWACTIVATED)
	; 32774 = 0x8000 + 6 = 0x8006 = HSHELL_FLASH := 32774 (HSHELL_HIGHBIT + HSHELL_REDRAW)
	; EVENT_OBJECT_CREATE := 32768 (0x8000)
	; --------------------------------------------------------------------------------
	EVENT_OBJECT_ACCELERATORCHANGE := 0x8012, 	EVENT_OBJECT_STATECHANGE := 32778, 
	EVENT_OBJECT_CLOAKED := 0x8017, 			EVENT_OBJECT_CONTENTSCROLLED := 0x8015, 
	EVENT_OBJECT_CONTENTSCROLLED := 0x8015, 	EVENT_OBJECT_DEFACTIONCHANGE := 0x8011, 
	EVENT_OBJECT_DESCRIPTIONCHANGE := 0x800D, 	EVENT_OBJECT_CREATE := 32768, 
	EVENT_OBJECT_DESTROY := 0x8001, 			EVENT_OBJECT_DRAGSTART := 0x8021, 
	EVENT_OBJECT_DRAGCANCEL := 0x8022, 			EVENT_OBJECT_DRAGCOMPLETE := 0x8023, 
	EVENT_OBJECT_DRAGENTER := 0x8024, 			EVENT_OBJECT_DRAGLEAVE := 0x8025, 
	EVENT_OBJECT_DRAGDROPPED := 0x8026, 		EVENT_OBJECT_END := 0x80FF, 
	EVENT_OBJECT_FOCUS := 0x8005, 				EVENT_OBJECT_HELPCHANGE := 0x8010, 
	EVENT_OBJECT_HIDE := 0x8003, 				EVENT_OBJECT_HOSTEDOBJECTSINVALIDATED := 0x8020, 
	EVENT_OBJECT_IME_HIDE := 0x8028, 			EVENT_OBJECT_IME_SHOW := 0x8027, 
	EVENT_OBJECT_IME_CHANGE := 0x8029, 			EVENT_OBJECT_INVOKED := 0x8013, 			EVENT_OBJECT_LIVEREGIONCHANGED := 0x8019, 	EVENT_OBJECT_LOCATIONCHANGE := 0x800B, 		
	EVENT_OBJECT_NAMECHANGE := 0x800C, 			EVENT_OBJECT_PARENTCHANGE := 0x800F, 
	EVENT_OBJECT_REORDER := 0x8004, 			EVENT_OBJECT_SELECTION := 0x8006, 
	EVENT_OBJECT_SELECTIONADD := 0x8007, 		EVENT_OBJECT_SELECTIONREMOVE := 0x8008, 	
	EVENT_OBJECT_SELECTIONWITHIN := 0x8009, 	EVENT_OBJECT_SHOW := 0x8002, 
	EVENT_OBJECT_TEXTEDIT_CONVERSIONTARGETCHANGED := 0x8030, 
	EVENT_OBJECT_TEXTSELECTIONCHANGED := 0x8014, 
	EVENT_OBJECT_UNCLOAKED := 0x8018, 
	EVENT_OBJECT_VALUECHANGE := 0x800E, 		EVENT_SYSTEM_ALERT := 0x0002, 
	EVENT_SYSTEM_ARRANGMENTPREVIEW := 0x8016, 	EVENT_SYSTEM_CAPTUREEND := 0x0009, 
	EVENT_SYSTEM_CAPTURESTART := 0x0008, 		EVENT_SYSTEM_CONTEXTHELPEND := 0x000D, 
	EVENT_SYSTEM_CONTEXTHELPSTART := 0x000C, 	EVENT_SYSTEM_DESKTOPSWITCH := 0x0020, 
	EVENT_SYSTEM_DIALOGEND := 0x0011, 			EVENT_SYSTEM_DIALOGSTART := 0x0010, 
	EVENT_SYSTEM_DRAGDROPEND := 0x000F, 		EVENT_SYSTEM_DRAGDROPSTART := 0x000E, 
	EVENT_SYSTEM_END := 0x00FF, 				EVENT_SYSTEM_FOREGROUND := 0x0003, 
	EVENT_SYSTEM_MENUSTART := 0x0004, 			EVENT_SYSTEM_MENUEND := 0x0005, 
	EVENT_SYSTEM_MENUPOPUPSTART := 0x0006, 		EVENT_SYSTEM_MENUPOPUPEND := 0x0007, 
	EVENT_SYSTEM_MINIMIZEEND := 0x0017, 		EVENT_SYSTEM_MINIMIZESTART := 0x0016, 
	EVENT_SYSTEM_MOVESIZEEND := 0x000B, 		EVENT_SYSTEM_MOVESIZESTART := 0x000A, 
	EVENT_SYSTEM_SCROLLINGEND := 0x0013, 		EVENT_SYSTEM_SCROLLINGSTART := 0x0012, 
	EVENT_SYSTEM_SOUND := 0x0001, 				EVENT_SYSTEM_SWITCHEND := 0x0015, 
	EVENT_SYSTEM_SWITCHSTART := 0x0014

	nEvent := (event = 1) ? 'HSHELL_WINDOWCREATED' : (event = 2) ? 'HSHELL_WINDOWDESTROYED' : (event = 3) ? 'HSHELL_ACTIVATESHELLWINDOW' : (event = 4) ? 'HSHELL_WINDOWACTIVATED' : (event = 32772) ? 'HSHELL_RUDEAPPACTIVATED' : (event = 5) ? 'HSHELL_GETMINRECT' : (event = 6) ? 'HSHELL_REDRAW' : (event = 7) ? 'HSHELL_TASKMAN' : (event = 8) ? 'HSHELL_LANGUAGE' : (event = 9) ? 'HSHELL_SYSMENU' : (event = 10) ? 'HSHELL_ENDTASK' : (event = 11) ? 'HSHELL_ACCESSIBILITYSTATE' : (event = 12) ? 'HSHELL_APPCOMMAND' : (event = 13) ? 'HSHELL_WINDOWREPLACED' : (event = 14) ? 'HSHELL_WINDOWREPLACING' : (event = 32768) ? 'HSHELL_HIGHBIT' : (event = 32774) ? 'HSHELL_FLASH' : 'No Event'
	; //0x800A = EVENT_OBJECT_STATECHANGE := 32778
	; (hookevent = 0x8012) ? 'EVENT_OBJECT_ACCELERATORCHANGE' : 'Fail'
	; = 'EVENT_OBJECT_CLOAKED'
	; if (event != HSHELL_RUDEAPPACTIVATED || !hwnd || !1){
	; 	return
	; }
	; if ((event = 1) || (event = 2) || (event = 3) || (event = 4) || (event = 32772)) {
	if ((event = 1) || (event = 3) || (event = 4) || (event = 32772)) {
	SendLevel((A_SendLevel + 1))
	BlockInput(1)
	pvTxt := A_DetectHiddenText, DetectHiddenText(1)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(1)
	hznHwnd := 0
	matches := []
	match := ''
	hWnd := WinActive('A')
	
	; hznGUI := Gui('+LastFound') ;, A_Process)
	
	; hznGUI.Opt('LastFound')
	; OutputDebug(hznGUI.Opt('LastFound'))
	; GuiControlObj := GuiCtrlFromHwnd(Hwnd)
	DllCall("GetWindowThreadProcessId", "Int", hwnd, "Int*", &tpID := 0)
	name := WinGetProcessName(hwnd)
	A_Process := name
	; if (A_Process ~= 'i).*hznHorizon.*'){
	; 	hznGui.OnEvent('32768'||'32778'||'0x8002',OtherToolTip)
	; }
	; hznGUI.Show()

	HznToolTip()
	/*
	if (A_Process ~= 'i).*hznHorizon.*')
		static hznHwnd := hWnd
		list := []
		test1 := WinGetTitle('A')
		test2 := WinActive(test1)
		OutputDebug(test1 . '`n')
		OutputDebug(test2 . '`n')
		try {
			list := WinGetControls(hznHwnd)
			ClassNN := ''
			hTb := 0
			for , ClassNN in list {
				if (ClassNN ~= 'i)m.*bar.*') {
					OutputDebug('ClassNN.Length' . ClassNN.Length . '`n')
					If (ClassNN.Length > 1){
						matches.Push(ClassNN)
						for each, match in matches
							hTb := ControlGetHwnd(match, 'A')
							HznEnableButtons(hTb)
					}
					Else {
						hTb := ControlGetHwnd(ClassNN, 'A')
						HznEnableButtons(hTb)
					}

					; MsgBox(hznHwnd . hTb . ClassNN)
				}
			}
		}
		*/
; ------------------------------------------

; hWinEventHook := DllCall("SetWinEventHook"
; 	, "UInt", eventMin ;  UINT eventMin
; 	, "UInt", eventMax ;  UINT eventMax
; 	, "Ptr", 0x0 ;  HMODULE hmodWinEventProc
; 	, "Ptr", CB_WinEventProc ;  WINEVENTPROC lpfnWinEventProc
; 	, "UInt", idProcess ;  DWORD idProcess
; 	, "UInt", 0x0 ;  DWORD idThread
; 	, "UInt", 0x0 | 0x2) ;  UINT dwflags, OutOfContext|SkipOwnProcess
	; ------------------------------------------
HznToolTip() {
	ToolTip('A_Process: ' . A_Process . ' '
		. tpID . ' (tPID) ' . hWnd . ' (hWnd)' . '`n'
		. 'Event: '		. event . '`n'
		. 'nEvent: '		. nEvent . '`n'
		; . 'C_Focus_Hwnd:  '	. fCtl . '`n'
		; . 'C_Focus_Class:  '	. hCtl . '`n'
		; . 'List: '		. list . '`n'
		, 0, 0
	)
}
OtherToolTip() {
	ToolTip('A_Process: ' . A_Process . ' '
		; . 'C_Focus_Hwnd:  '	. fCtl . '`n'
		; . 'C_Focus_Class:  '	. hCtl . '`n'
		; . 'List: '		. list . '`n'
		, 0, 0
	)
}
; ------------------------------------------
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	BlockInput(0)
	SendLevel(0)
	return A_Process
}
}
class WinEventHook
{
	__New(eventMin, eventMax, hookProc, options := '', idProcess := 0, idThread := 0, dwFlags := 0) {
		this.pCallback := CallbackCreate(hookProc, options, 7)
		this.hHook := DllCall('SetWinEventHook', 'UInt', eventMin, 'UInt', eventMax, 'Ptr', 0, 'Ptr', this.pCallback
			, 'UInt', idProcess, 'UInt', idThread, 'UInt', dwFlags, 'Ptr')
	}
	__Delete() {
		DllCall('UnhookWinEvent', 'Ptr', this.hHook)
		CallbackFree(this.pCallback)
	}
}
; #Requires AutoHotkey v2

; Map out all open windows so we can keep track of their names when they're closed.
; After the window close event the windows no longer have their titles, so we can't do it afterwards.

; Set up our hook. Putting it in a variable is necessary to keep the hook alive, since once it gets
; rewritten (for example with hook := "") the hook is automatically destroyed.
; hook := SetWinEventHook(HandleWinEvent)
; We have no hotkeys, so Persistent is required to keep the script going.
; Persistent()

/**
 * Our event handler which needs to accept 7 arguments. To ignore some of them use the * character,
 * for example HandleWinEvent(hWinEventHook, event, hwnd, idObject, idChild, *)
 * @param hWinEventHook Handle to an event hook function. This isn't useful for our purposes 
 * @param event Specifies the event that occurred. This value is one of the event constants (https://learn.microsoft.com/en-us/windows/win32/winauto/event-constants).
 * @param hwnd Handle to the window that generates the event, or NULL if no window is associated with the event.
 * @param idObject Identifies the object associated with the event.
 * @param idChild Identifies whether the event was triggered by an object or a child element of the object.
 * @param idEventThread Id of the thread that triggered this event.
 * @param dwmsEventTime Specifies the time, in milliseconds, that the event was generated.
 */


/**
 * Sets a new WinEventHook and returns on object describing the hook. 
 * When the object is released, the hook is also released.
 * @param callbackFunc The function that will be called, which needs to accept 7 arguments:
 *    hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime
 * @param winTitle Optional: WinTitle of a certain window to hook to. Default is system-wide hook.
 * @param eventMin Optional: Specifies the event constant for the lowest event value in the range of events that are handled by the hook function.
 *  Default is EVENT_OBJECT_CREATE = 0x8000
 *  See more about event constants: https://learn.microsoft.com/en-us/windows/win32/winauto/event-constants
 * @param eventMax Optional: Specifies the event constant for the highest event value in the range of events that are handled by the hook function.
 *  Default is EVENT_OBJECT_DESTROY = 0x8001
 * @param flags Flag values that specify the location of the hook function and of the events to be skipped.
 *  Default is WINEVENT_OUTOFCONTEXT = 0 and WINEVENT_SKIPOWNPROCESS = 2. 
 * @returns {Object} 
 */
/*
find_bar() {
	SendLevel((A_SendLevel + 1))
	BlockInput(1)
	pvTxt := A_DetectHiddenText, DetectHiddenText(1)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(1)
	; idWin := WinGetID('ahk_exe hznHorizon.exe')
	; ; Sleep(200)
	; WinActivate(idWin)
	; hznhWnd := WinExist('ahk_exe hznHorizon.exe')
	list := WinGetControls('A')
	; clist := WinGetControlsHwnd('A')
	
	; fCtl := ControlGetFocus('A')
	; hCtl := ControlGetClassNN(fCtl)
	static ControlClassNN := ''
	; static ControlHwnd := ''
	; static enumlist := ''
	static matchlist := ''
	static matches := []
	; static CtlList := []
	static match := ''
	needle := 'i)m*.bar*.'
	; needle := 'im)(\w+\sbar\w+)'
	; ? KEEP FROM HERE
	for , ControlClassNN in list {
		; enumlist .= '[' . A_Index . ']' . A_Space . ControlClassNN '`n'
		if (ControlClassNN ~= needle) {
			; if InStr(ControlClassNN , needle) {
			hTb := ControlGetHwnd(ControlClassNN, 'A')
			OutputDebug('[' . A_Index . '] ' . ControlClassNN . ' hCtl: ' hTb '`n')
			; matches.Push(ControlClassNN . A_Space . '(' . hTb . ')')
			matches.Push(hTb)
		}
	}

	try {
		if (matches) {
			for , match in matches {
				; OutputDebug('[' . A_Index . '] ' . match . '`n')
				matchlist .= '[' . A_Index . '] ' . match . '`n'
			}
		}
		; FileAppend(matchlist, '_enumlist_after_regex.txt', '`n UTF-8')
		; Else {
		; 	MsgBox('No ' . 'matches' . ' found.')
		; }
	}
	; MsgBox(matchlist)
	; MsgBox(listctl)
	; return
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	BlockInput(0)
	SendLevel(0)
	return hTb
}
*/
; --------------------------------------------------------------------------------
; WinShowHook(true)
; WinShowHook(doHook) {
; 	static CBA := CallbackCreate(GetHznControls)
; 	static hHook := 0
; 	; static EVENT_OBJECT_SHOW := 0x8002
; 	static EVENTmin := EVENT_OBJECT_CREATE := 0x8000
; 	static EVENTmax := EVENT_OBJECT_DESTROY := 0x8001
; 	;				0x8000 = EVENT_OBJECT_CREATE := 0x8000
; 	;				0x8011 = EVENT_OBJECT_DEFACTIONCHANGE
; 	;				0x800D = EVENT_OBJECT_DESCRIPTIONCHANGE
; 	;				0x8001 = EVENT_OBJECT_DESTROY := 0x8001
; 	if (doHook && hHook = 0) {
; 		; MsgBox("Installing WinEvent Hook")
; 		OutputDebug("Installing WinEvent Hook")
; 		hHook := DllCall(
; 			"SetWinEventHook",
; 			"Int", EVENTmin,
; 			"Int", EVENTmax,
; 			"Ptr", 0,
; 			"Ptr", CBA,
; 			"Int", 0,
; 			"Int", 0,
; 			"Int", 0,
; 			"Ptr")
; 		list := GetHznControls()
; 		info := 'CBA: ' CBA '`n' 'hHook: ' hHook '`n' 'list: ' list
; 		ToolTip(info, 0, 200)
; 	} else if ( not doHook && hHook != 0) {
; 		DllCall("UnhookWinEvent", "Ptr", hHook)
; 		hHook := 0
; 	}
; }
GetHznControls(hwnd := 'A', &list_controls := '')
{
	win_get_controls := WinGetControls(hwnd)
	list_controls := DisplayObj(win_get_controls)
	OutputDebug(list_controls '`n')
	return list_controls
	; TODO toolbarwindow325????? and RegExMatch for msvb_lib_toolbar
}
/*

; --------------------------------------------------------------------------------
; Sub-Section .....: WinActiveChange2()
; Description .....: removed WinGet, A_Process, ProcessName, A in place of getActiveProcessName()
; --------------------------------------------------------------------------------
*/
WinActiveChange2(wParam, hwnd?, &A_Process:=0) ; hwnd)
{
	; static HSHELL_RUDEAPPACTIVATED := 32772, ;pId
	if (wParam != 32772)
	{
		Return ; only listen for HSHELL_RUDEAPPACTIVATED
	}
	; handle := DllCall("GetForegroundWindow", "Ptr")
	hWnd := WinActive('A')
	DllCall("GetWindowThreadProcessId", "Int", hwnd, "Int*", &tpID:=0)
	pname:= getActiveProcessName()
	
	callback := CallbackCreate(enumChildCallback, "Fast")
	DllCall('EnumChildWindows', 'Int', hwnd, 'Ptr', callback, 'Int', tpId)
	handle := DllCall("OpenProcess", "Int", 0x0400, "Int", 0, "Int", tpID)
	length := 259 ;max path length in windows
	VarSetStrCapacity(&name, length) ; V1toV2: if 'name' is NOT a UTF-16 string, use 'name := Buffer(length)'
	DllCall("QueryFullProcessImageName", "Int", handle, "Int", 0, "Ptr", name, "int*", &length)
	DllCall("CloseHandle", "Ptr", handle) ; added => not in original script
	; name = full path
	; pname = program (.exe)
	SplitPath(name, &A_Process)
	; ToolTip('tpID: ' . tpID . '`n'
	; 	; . 'pID: ' 	. pID . '`n'
	; 	. 'hWnd: ' . hWnd . '`n'
	; 	. 'name: ' . name . '`n'
	; 	. 'name: ' . pname . '`n'
	; 	. 'name: ' . name . '`n'
	; 	; . 'List: '	. list . '`n'
	; 	, 0, 0
	; )
	; ToolTip(A_Process, 0, 0, 1) ;"`npname: " getActiveProcessName(),0,0,1
	return A_Process
}
; --------------------------------------------------------------------------------
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; --------------------------------------------------------------------------------
; Sub-Section .....: getActiveProcessName()
; Description .....: maybe a better version of WinActiveChange()
; --------------------------------------------------------------------------------
getActiveProcessName(*)
{
	handle := DllCall("GetForegroundWindow", "Ptr")
	DllCall("GetWindowThreadProcessId", "Int", handle, "int*", &wtpID:=0)
	static true_tpID := wtpID
	callback := CallbackCreate(enumChildCallback, "Fast")
	DllCall("EnumChildWindows", "int", handle, "ptr", callback, "int", wtpID)
	handle := DllCall("OpenProcess", "Int", 0x0400, "Int", 0, "Int", true_tpID)
	length := 259 ;max path length in windows
	VarSetStrCapacity(&name, length) ; V1toV2: if 'name' is NOT a UTF-16 string, use 'name := Buffer(length)'
	; DllCall("QueryFullProcessImageName", "Int", handle, "Int", 0, "Ptr", name, "int*", &length)
	DllCall("CloseHandle", "Ptr", handle) ; added => not in original script
	; name = full path
	; pname = program (.exe)
	SplitPath(name, &pname)
	return pname
}

enumChildCallback(hwnd, tpID, &child_tpID:=0)
{
	DllCall("GetWindowThreadProcessId", "Int", hwnd, "int*", &child_tpID:=0)
	if (child_tpID != tpID)
		static true_tpID := child_tpID
	DllCall("CloseHandle", "Ptr", hwnd) ; added => not in original script
	; return 1
}
; --------------------------------------------------------------------------------
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; --------------------------------------------------------------------------------
; Sub-Section .....: Script Name, Startup Path, and icon
; --------------------------------------------------------------------------------
; ICON := 'C:\Users\bacona\OneDrive - FM Global\Documents\AutoHotkey\Lib\ICON\DllCall.ico'
; --------------------------------------------------------------------------------
; Sub-Section .....: Create Tray Menu
; --------------------------------------------------------------------------------

; Tray:= A_TrayMenu
; TraySetIcon(ICON) ; this changes the icon into a little dllcall thing.
; Tray.Delete() ; V1toV2: not 100% replacement of NoStandard, Only if NoStandard is used at the beginning
; addTrayMenuOption("Made with nerd by Adam Bacon and Terry Keatts", "madeBy")
; addTrayMenuOption()
; addTrayMenuOption("Run at startup", "runAtStartup")
; Tray.fileExist(Startup_Shortcut) ? "check" : "unCheck"("Run at startup") ; update the tray menu status on startup
; addTrayMenuOption()
; Tray.AddStandard()

; 	; GUI =======================================================================================================================================================================

; 	Main := Gui("+Resize +MinSize854x480")
; 	Main.MarginX := 0
; 	Main.MarginY := 0

; 	Main.SetFont("s9", "Segoe UI")
; 	LV := Main.AddListView("xm+10 ym+10 w830 r10", ["Event", "Time", "ProcessID", "ProcessName", "ParentProcessID", "ExecutablePath", "CommandLine"])
; 	for k, v in ["80", "80", "80", "120", "80", "150"]
; 		LV.ModifyCol(k, v)

; 	Main.OnEvent("Size", GuiSize)
; 	Main.OnEvent("Close", GuiClose)
; 	Main.Show("AutoSize")


; 	; WMI =======================================================================================================================================================================

; 	WMI := ComObjGet("winmgmts:")

; 	ComObjConnect(Sink := ComObject("WbemScripting.SWbemSink"), "SINK_")

; 	Command := "WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
; 	WMI.ExecNotificationQueryAsync(Sink, "SELECT * FROM __InstanceCreationEvent " . Command)
; 	WMI.ExecNotificationQueryAsync(Sink, "SELECT * FROM __InstanceDeletionEvent " . Command)


; 	SINK_OnObjectReady(Obj, *)
; 	{
; 		TI := Obj.TargetInstance
; 		switch Obj.Path_.Class
; 		{
; 			case "__InstanceCreationEvent":
; 				{
; 					EventType := "Created"
; 					Time := FormatTime(SubStr(TI.CreationDate, 1, 14), "HH:mm:ss")
; 				}
; 			case "__InstanceDeletionEvent":
; 				{
; 					EventType := "Terminated"
; 					Time := FormatTime(A_Now, "HH:mm:ss")
; 				}
; 		}
; 		LV.Insert(1, , EventType, Time, TI.ProcessID, TI.Name, TI.ParentProcessId, TI.ExecutablePath, TI.CommandLine)
; 	}


; 	; WINDOW EVENTS =============================================================================================================================================================

; 	GuiSize(thisGui, MinMax, Width, Height)
; 	{
; 		if (MinMax = -1)
; 			return
; 		LV.Move(, , Width - 21, Height - 21)
; 		LV.Redraw()
; 	}


; 	GuiClose(*)
; 	{
; 		ComObjConnect(Sink)
; 		ExitApp()
; 	}


; 	; ===========================================================================================================================================================================


; 	; ===========================================================================================================================================================================
; }


; return

; #2::
