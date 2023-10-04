#Requires AutoHotkey v2
;@include-winapi
; --------------------------------------------------------------------------------
#Include <Directives\__AE.v2>
#Include <Directives\__HznToolbar>
#include <Directives\WinHook.v2>
; --------------------------------------------------------------------------------
#HotIf WinActive('ahk_exe hznHorizon.exe') ;&& WinActive('ahk_exe hznHorizon.exe')
; Callback definition for EVENT_SYSTEM_CAPTURESTART
; msvb_lib_toolbar_created := CallbackCreate(CallBack_TB_CREATED)
; /*
; Persistent(1)
; DetectHiddenText(1)
; DetectHiddenWindows(1)
WinShowHook(true)
~*Esc::Exit()
; Return
CoordMode('ToolTip', 'Screen')
HookProcedure() {
	global A_Process
	hWnd := WinActive('A')
	pID := WinGetPID(hWnd)
	Name := WinGetProcessName(hWnd)
	DllCall("GetWindowThreadProcessId", "Ptr", hWnd, "UInt*", &tpID := 0)
	; list := GetHznControls(hWnd)
	ToolTip(  'tpID: '	. tpID . '`n' 
			. 'pID: ' 	. pID . '`n' 
			. 'hWnd: '	. hWnd . '`n' 
			. 'name: '	. name . '`n'
			; . 'List: '	. list . '`n'
			, 0, 0
		)
	A_Process := name
		; hwnd := WinActive(hzntitle)
		; pID := WInGetID()
	GetHznControls(hwnd:=0)
	{
		win_get_controls := WinGetControls('A')
		list_controls := DisplayObj(win_get_controls)
		OutputDebug(list_controls '`n')
		return list_controls
	}
	; bak_TitleMatchMode := A_TitleMatchMode
	; OutputDebug('pMatch: ' bak_TitleMatchMode '`n')
	; SetTitleMatchMode('RegEx')
	; OutputDebug('nMatch' A_TitleMatchMode '`n')
	; ; Pause()
	; RegExMatch(list_controls, 'imJS)msvb_lib_toolbar\d', &m)
	; SetTitleMatchMode(bak_TitleMatchMode)
	; MsgBox(m)
	; nCtl := m[]
	; instance := SubStr(nCtl, -1, 1)
	; hTb := ControlGethWnd(nCtl, "A")
	; MsgBox(nCtl ' ' hTb . '`n')		
	SoundBeep(555,1000)
	return A_Process
}

WinShowHook(doHook) {
	static CBA := CallbackCreate(HookProcedure)
	static hHook := 0
	; static EVENT_OBJECT_SHOW := 0x8002
	static EVENTmin := HSHELL_RUDEAPPACTIVATED := 1
	static EVENTmax := HSHELL_RUDEAPPACTIVATED := 32772
	if (doHook && hHook = 0) {
		; MsgBox("Installing WinEvent Hook")
		OutputDebug("Installing WinEvent Hook")
		hHook := DllCall(
			"SetWinEventHook",
			"Int", EVENTmin,
			"Int", EVENTmax,
			"Ptr", 0,
			"Ptr", CBA,
			"Int", 0,
			"Int", 0,
			"Int", 0,
			"Ptr")
	} else if ( not doHook && hHook != 0) {
		DllCall("UnhookWinEvent", "Ptr", hHook)
		hHook := 0
	}
}
; }
return
/*
{
DetectHiddenText(1)
DetectHiddenWindows(1)
OnExit(ObjBindMethod(WinHook.Event, "UnHookAll"))

start := false
MyGui := Gui("+AlwaysOnTop", "WinHook Monitor")
MyGui.Add("Tab3", "w1020" " h480 x5 y5 vTabs +0x8 +Theme", ["Shell", "Events"])
MyGui["Tabs"].UseTab(1)
Btn3 := MyGui.Add("Button", "", "Start")
Btn4 := MyGui.Add("Button", "x+10", "Stop")
MyGui.Add("Text", "x+10 yp+4", "WinHook.Shells")
Btn3.OnEvent("Click", (*) => WinHook.Shell.Add(AllShell))
Btn4.OnEvent("Click", (*) => WinHook.Shell.UnHookAll())
LV2 := MyGui.Add("ListView", "xs+10 yp+25 r25 w1000", ["Win_Hwnd", "Win_Title", "Win_Class", "Win_Exe", "Win_Event"])
MyGui["Tabs"].UseTab(2)
Btn1 := MyGui.Add("Button", "", "Start")
Btn2 := MyGui.Add("Button", "x+10", "Stop")
Btn1.OnEvent("Click", (*) => WinHook.Event.Add(0x0000, 0xFFFF, AllEvent))
Btn2.OnEvent("Click", (*) => WinHook.Event.UnHookAll())
MyGui.Add("Text", "x+10 yp+4", "WinHook.Events")
LV1 := MyGui.Add("ListView", "xs+10 yp+25 r25 w1000", ["hWinEventHook", "Event", "hWnd", "idObject", "idChild", "dwEventThread", "dwmsEventTime", "wClass", "wTitle"])
MyGui.Show()

AllShell(Win_Hwnd, Win_Title, Win_Class, Win_Exe, Win_Event) {
	n := LV2.add(, Win_Hwnd, Win_Title, Win_Class, Win_Exe, Win_Event)
	LV2.Modify(n, "Vis")
	Loop 5
		LV2.ModifyCol(A_Index, "AutoHdr")
}

AllEvent(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
	try wTitle := WinGetTitle("ahk_id " Hwnd)
	try wClass := WinGetClass("ahk_id " Hwnd)
	n := LV1.add(, hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime, wClass, wTitle)
	LV1.Modify(n, "Vis")
	Loop 8
		LV1.ModifyCol(A_Index, "AutoHdr")
}
}
return
*/
/*
^+!a::
{
EnumWindowsProc(*)
{
match := Array()
controls := Array()
hTb := ''
v := ''
list_controls := ''
; Haystack := v
Matchposition := 0
needle := 'imJS)msvb_lib_toolbar\d'
controls := WinGetControls('A')
list_controls := DisplayObj(controls)
OutputDebug(list_controls)
controls.Push(controls)
for k, v in controls
	OutputDebug(v '`n')
	match := RegExMatch(v, needle, &match, 1)
	OutputDebug('Match: ' match)
return
}
}
/*
; max := controls.Capacity()
	; OutputDebug(max)
	loop controls.Length
		for k,v in controls
			FoundPos := InStr(v,needle,1)
			OutputDebug(FoundPos)
			; OutputDebug(' ' A_Index ' ' v)
			match.Push(v)
	return match
	list_ctrls := ''
	for , v in controls
		loop {
			OutputDebug(v)
			match := InStr(v,needle,,5)
			OutputDebug(match)
			RegExMatch(v, needle, &match, Matchposition+1)
			match .= match
		}
		; if(v = RegExMatch(v, needle))
		; 	nCtl .= v
		; 	hwnd := ControlGethWnd(v, 'A')
		; 	hTb .= hwnd
		; 	OutputDebug(nCtl ' ' hTb)
		ctrls .= v ' '
		; ctrls .= v
	; Loop parse ctrls, , A_Space
		list_ctrls := ctrls
	OutputDebug(list_ctrls '`n')
	try Matchposition := RegExMatch(list_ctrls, needle, &match, Matchposition + 1)
		OutputDebug(Matchposition '`n')
	; RegExMatch(Haystack, needle, &match, Matchposition + 1)
	MsgBox(list_ctrls)
	return
	list_controls := DisplayObj(controls)
	OutputDebug(list_controls)
	Haystack := list_controls
	Matchposition := 0
	Loop
	{
		Matchposition := RegExMatch(Haystack, needle, &match, Matchposition + 1) ; Use last match position + 1 as starting position and put the next match in "CurrentMatch"
		If !Matchposition ; if no more matches are found, exit the loop
			Break
		; AllMatches .= Match ; concatenate matches (same as AllMatches := AllMatches . CurrentMatch)
		OutputDebug(Match)
	}

	return
	; SetTitleMatchMode(bak_TitleMatchMode)
	nCtl := m[]
	mtc := DisplayObj(m)
	OutputDebug(mtc)
	return
	hTb := ControlGethWnd(nCtl, "A")


	OutputDebug(mtc)
	return
	Loop
		{
			RegExMatch(A_LoopField, 'imJS)msvb_lib_toolbar\d', &m)
			match.Push(m)
		}
	list_matches := DisplayObj(match)
	OutputDebug(list_matches)
	return
	; list_controls := DisplayObj(win_get_controls)
	; OutputDebug(list_controls '`n')
	; ; bak_TitleMatchMode := A_TitleMatchMode
	; ; OutputDebug(bak_TitleMatchMode '`n')
	; SetTitleMatchMode(3)
	; ; OutputDebug(A_TitleMatchMode '`n')
	; RegExMatch(list_controls, 'imJS)^msvb_lib_toolbar', &m)
	; ; SetTitleMatchMode(bak_TitleMatchMode)
	; nCtl := DisplayObj(m)
	; Loop parse m.Count
	; 	try hTb := ControlGethWnd(nCtl, "A")
	; 	OutputDebug('Control: ' . nCtl . '`n' . 'hWnd: ' . hTb . '`n')
	; return
	; win_get_controls := WinGetControls('A')
	; list_controls := DisplayObj(win_get_controls)
	; ; bak_TitleMatchMode := A_TitleMatchMode
	; ; SetTitleMatchMode('RegEx')
	; RegExMatch(list_controls, 'msvb_lib_toolbar\d', &m)
	; ; SetTitleMatchMode(bak_TitleMatchMode)
	; nCtl := m[]
	; instance := SubStr(nCtl, -1, 1)
	; hTb := ControlGethWnd(nCtl, "A")
	; MsgBox(nCtl ' ' hTb . '`n')
	; return
}
}
; while !WinActive('ahk_exe hznHorizon.exe')
; 	return
; YourMonitoringCallback()
; GetToolbar(instance)
; {
; 	
; 	win_get_controls := WinGetControls('A')
; 	list_controls := DisplayObj(win_get_controls)
; 	bak_TitleMatchMode := A_TitleMatchMode
; 	SetTitleMatchMode('RegEx')
; 	RegExMatch(list_controls, 'msvb_lib_toolbar\d', &m)
; 	SetTitleMatchMode(bak_TitleMatchMode)
; 	&nCtl := m[]
; 	&instance := SubStr(nCtl, -1, 1)
; 	&hTb := ControlGethWnd(nCtl, "A")
; 	MsgBox(nCtl ' ' hTb . '`n')
; 	return
; }
; YourMonitoringCallback(wParam?, lParam?, Message?, hWnd?) {
; 	static instances := []
; 	try {
; 		instance := GetToolbar(instances.Length + 1)
; 		instances.Push(instance)
; 	} catch {
; 		; Handle the edge case where the toolbar info couldn't be retrieved.
; 	}
; 	MsgBox(
; 		"Toolbar hWnd: " instance.hTb
; 		"`n"
; 		"Textbox hWnd: " instance
; 	)

; }
#HotIf