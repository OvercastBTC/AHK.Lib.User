; --------------------------------------------------------------------------------
; Section .....: Auto-Execution
; --------------------------------------------------------------------------------
#SingleInstance Force
Persistent(1) ; Keeps script permanently running
; #MaxThreads 255 ; Default: 10, Max: 255.
SendMode("Input")  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.
SetTitleMatchMode(2) ; sets title matching to search for "containing" instead of "exact"
DetectHiddenText(true)
DetectHiddenWindows(true)
CoordMode("ToolTip", "Screen")
ProcessSetPriority('H')
#Requires Autohotkey v2
; --------------------------------------------------------------------------------
; Sub-Section .....: Initiation Section
; --------------------------------------------------------------------------------
#HotIf WinActive("ahk_exe Code.exe")
~^s::Run(A_ScriptName)
#HotIf
ICON := '.\ICON\DllCall.ico'
TraySetIcon(ICON) ; this changes the icon into a little dllcall thing.
; --------------------------------------------------------------------------------
; #Persistent
; SetBatchLines, -1
; Process, Priority, , High
#1::
; _A_Process(&A_Process:=0)
_A_Process(*)
{
	Static HSHELL_RUDEAPPACTIVATED := 32772
	apGui := Gui()
	apGui.Opt('+LastFound')
	hWnd := WinActive('A')

	; DllCall("RegisterShellHookWindow", UInt, hWnd)
	DllCall('RegisterShellHookWindow', 'UInt', hWnd)
	MsgNum := DllCall("RegisterWindowMessage", 'Str', "SHELLHOOK")
	; Msg := DllCall('RegisterWindowMessage', 'Str', 'SHELLHOOK', HSHELL_RUDEAPPACTIVATED, 'Int')
	OnMessage(MsgNum, "ShellMessage")
	; OnMessage(Msg, "ShellMessage")
	Return

	ShellMessage(wParam, hwnd)
	{
		; If (wParam = 1) || (wParam = 32772) ;  HSHELL_WINDOWCREATED := 1
		If (wParam = 32772) ;  HSHELL_WINDOWCREATED := 1
		{
			A_Process := WinGetProcessName(hwnd)
			ToolTip(A_Process, 0, 0)
			Return A_Process
		}
	}
}
return
; #1::
; A_Process(A_Process)
; {
; 	Static HSHELL_RUDEAPPACTIVATED := 32772
; 	; Global A_Process	; set a super-global variable that can be accessible within all functions by default.
; 	apGui := Gui()
; 	apGui.Opt('+LastFound')
; 	DllCall('RegisterShellHookWindow', 'UInt')
; 	; OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), WinActiveChange) ; original
; 	;OnMessage(DllCall("RegisterWindowMessage", Str, "SHELLHOOK"), "WinActiveChange1")
; 	Msg := DllCall('RegisterWindowMessage', 'Str', 'SHELLHOOK', HSHELL_RUDEAPPACTIVATED, 'Int')
	
	
; 	OnMessage( Msg, WinActiveChange2) ;HSHELL_RUDEAPPACTIVATED,hwnd := DllCall("GetForegroundWindow", "Ptr")))
; 	ToolTip(A_Process)
; 	Return A_Process
; }
; --------------------------------------------------------------------------------
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; --------------------------------------------------------------------------------
; Sub-Section .....: WinActiveChange()
; Comment: [x] NOT IN USE
; --------------------------------------------------------------------------------

WinActiveChange(wParam?, hwnd?)
{
	
	static HSHELL_RUDEAPPACTIVATED := 32772	
	if (wParam != HSHELL_RUDEAPPACTIVATED) ; only listen for HSHELL_RUDEAPPACTIVATED
		return
	
	A_Process := WinGetProcessName('A')
	; ToolTip(a_process, 0, 0, 1)
	MsgBox(A_Process)
	return A_Process
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
WinActiveChange2(wParam, hwnd*) ; hwnd)
{
	HSHELL_RUDEAPPACTIVATED := 32772, ;pId
	
	wParam != HSHELL_RUDEAPPACTIVATED
	{
		Return ; only listen for HSHELL_RUDEAPPACTIVATED
	}
	; handle := DllCall("GetForegroundWindow", "Ptr")

	DllCall("GetWindowThreadProcessId", "Int", hwnd, "Int*", &pId)
	; pId := DllCall('GetWindowThreadProcessId'
	; 			 , 'Int', hwnd
	; ; , "Int*", &pId
	; 				)
	getActiveProcessName()
	callback := CallbackCreate(enumChildCallback, "Fast")
	DllCall('EnumChildWindows'
			, 'Int', hwnd, 'Ptr', callback, 'Int', pId)
	handle := DllCall("OpenProcess", "Int", 0x0400, "Int", 0, "Int", pId)
	length := 259 ;max path length in windows
	VarSetStrCapacity(&name, length) ; V1toV2: if 'name' is NOT a UTF-16 string, use 'name := Buffer(length)'
	DllCall("QueryFullProcessImageName", "Int", handle, "Int", 0, "Ptr", name, "int*", &length)
	DllCall("CloseHandle", "Ptr", handle) ; added => not in original script
	; name = full path
	; pname = program (.exe)
	SplitPath(name, &A_Process)
	ToolTip(A_Process, 0, 0, 1) ;"`npname: " getActiveProcessName(),0,0,1
	return A_Process
}
; --------------------------------------------------------------------------------
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; --------------------------------------------------------------------------------
; Sub-Section .....: getActiveProcessName()
; Description .....: maybe a better version of WinActiveChange()
; --------------------------------------------------------------------------------
getActiveProcessName()
{
	handle := DllCall("GetForegroundWindow", "Ptr")
	DllCall("GetWindowThreadProcessId", "Int", handle, "int*", &pid)
	global true_pid := pid
	callback := CallbackCreate(enumChildCallback, "Fast")
	DllCall("EnumChildWindows", "int", handle, "ptr", callback, "int", pid)
	handle := DllCall("OpenProcess", "Int", 0x0400, "Int", 0, "Int", true_pid)
	length := 259 ;max path length in windows
	VarSetStrCapacity(&name, length) ; V1toV2: if 'name' is NOT a UTF-16 string, use 'name := Buffer(length)'
	DllCall("QueryFullProcessImageName", "Int", handle, "Int", 0, "Ptr", name, "int*", &length)
	DllCall("CloseHandle", "Ptr", handle) ; added => not in original script
	; name = full path
	; pname = program (.exe)
	SplitPath(name, &pname)
	return pname
}

enumChildCallback(hwnd, pid)
{
	DllCall("GetWindowThreadProcessId", "Int", hwnd, "int*", &child_pid)
	if (child_pid != pid)
		global true_pid := child_pid
	DllCall("CloseHandle", "Ptr", hwnd) ; added => not in original script
	return 1
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