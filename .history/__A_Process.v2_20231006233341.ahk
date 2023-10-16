#Requires Autohotkey v2+
#Include <Directives\__AE.v2>
#Include <HznPlus.v2>
; Persistent(1) ; Keeps script permanently running
CoordMode("ToolTip", "Screen")
; ProcessSetPriority('High')
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
ICON := '.\ICON\DllCall.ico'
TraySetIcon(ICON) ; this changes the icon into a little dllcall thing.
; --------------------------------------------------------------------------------
ShellHook_AProcess()
ShellHook_AProcess()
{
	Global HSHELL_RUDEAPPACTIVATED := 32772
	hWnd := WinActive('A')
	DllCall("RegisterShellHookWindow", "UInt", A_ScriptHwnd)
	OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), A_Process_GetInfo)
}
A_Process_GetInfo(event, hwnd,*) {
	; static HSHELL_RUDEAPPACTIVATED := 32772
	if (event != HSHELL_RUDEAPPACTIVATED || !hwnd)
		return
	else
		try {
			hWnd := WinActive('A')
			DllCall("GetWindowThreadProcessId", "Int", hwnd, "Int*", &tpID := 0)
			name := WinGetProcessName(hwnd)
			A_Process := name
		try {
			if (A_Process = 'hznHorizon.exe')
				try
					fCtl := ControlGetFocus('A')
					hCtl := ControlGetClassNN(fCtl)
					list := GetHznControls(hWnd)
					try {
						if (hctl ~= 'i)TX11*.') {
							Sleep(1500)
							HznEnableButtons()
						} Else {
							Sleep(3000)
							HznEnableButtons()
						}
					} catch Error as e {
						Sleep(5000)
						HznEnableButtons()
					}
			HznToolTip()
		} catch Error as e {
			fCtl := ''
			hCtl := ''
			list := ''
			OtherToolTip()
		}

		return A_Process
}
}
HznToolTip() {
	ToolTip('tpID:          ' . tpID . '`n'
		; 'pID: ' 	. pID . '`n'
		. 'hWnd:       ' . hWnd . '`n'
		. 'A_Process: ' . A_Process . '`n'
		. 'C_Focus_Hwnd:  ' . fCtl . '`n'
		. 'C_Focus_Class:  ' . hCtl . '`n'
		; . 'List: '		. list . '`n'
		, 0, 0
	)
}
OtherToolTip() {
	ToolTip('tpID:          ' . tpID . '`n'
		; 'pID: ' 	. pID . '`n'
		. 'hWnd:       ' . hWnd . '`n'
		. 'A_Process: ' . A_Process . '`n'
		; . 'C_Focus_Hwnd:  '	. fCtl . '`n'
		; . 'C_Focus_Class:  '	. hCtl . '`n'
		; . 'List: '		. list . '`n'
		, 0, 0
	)
}
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
aProcess(*)
{
	Static HSHELL_RUDEAPPACTIVATED := 32772
	; Global A_Process	; set a super-global variable that can be accessible within all functions by default.
	try hwnd := WinActive('A')
    DllCall("RegisterShellHookWindow", "UInt", Hwnd)
    MsgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")
    OnMessage(MsgNum, WinActiveChange)
	; OnMessage(DllCall("RegisterWindowMessage", 'int*' "SHELLHOOK"), "WinActiveChange2")
	; OnMessage(Msg, "WinActiveChange2")
	; ToolTip(A_Process)
	; Return A_Process
}
; OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), WinActiveChange) ; original
; 	Msg := DllCall('RegisterWindowMessage', 'Str', 'SHELLHOOK', HSHELL_RUDEAPPACTIVATED, 'Int')
	
	
	; OnMessage( Msg, WinActiveChange2) ;, HSHELL_RUDEAPPACTIVATED,hwnd := DllCall("GetForegroundWindow", "Ptr")))
; }
; --------------------------------------------------------------------------------
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; --------------------------------------------------------------------------------
; Sub-Section .....: WinActiveChange()
; --------------------------------------------------------------------------------

WinActiveChange(wParam?, hwnd?)
{
	Global A_Process
	static HSHELL_RUDEAPPACTIVATED := 32772	
	; if (wParam != HSHELL_RUDEAPPACTIVATED) ; only listen for HSHELL_RUDEAPPACTIVATED
	if !(HSHELL_RUDEAPPACTIVATED) ; only listen for HSHELL_RUDEAPPACTIVATED
		return
	pvTxt := A_DetectHiddenText, DetectHiddenText(0)
	pvWin := A_DetectHiddenWindows, DetectHiddenWindows(0)
	Try
		hWnd := WinActive('A')
		; pID := WinGetPID(hWnd)
		Name := WinGetProcessName(hWnd)
		DllCall("GetWindowThreadProcessId", "Ptr", hWnd, "UInt*", &tpID := 0)
	DetectHiddenText(pvTxt), DetectHiddenWindows(pvWin)
	A_Process := Name
	ToolTip(  'tpID: ' . tpID . '`n'
			. 'hWnd: ' . hWnd . '`n'
			. 'name: ' . name . '`n'
			; . 'List: '	. list . '`n'
			, 0, 0
	)
	; MsgBox(A_Process)
	; return A_Process
}
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
; Sub-Section .....: WinActiveChange1()
; Description .....: removed WinGet, A_Process, ProcessName, A in place of getActiveProcessName()
; Comment: [x] NOT IN USE
; --------------------------------------------------------------------------------

; WinActiveChange1(wParam, hwnd*) {
; 	HSHELL_RUDEAPPACTIVATED := 32772
; 	global A_Process
	
; 	if (wParam != HSHELL_RUDEAPPACTIVATED) ; only listen for HSHELL_RUDEAPPACTIVATED
; 		return
; 	getActiveProcessName()
; 	;WinGet,  a_process,  ProcessName,  A
; 	ToolTip("Original: " a_process "`npname: " getActiveProcessName(),0,0,1)
; 	return A_Process
; }
; --------------------------------------------------------------------------------

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
