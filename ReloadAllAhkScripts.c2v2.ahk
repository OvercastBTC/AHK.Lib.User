; -=================================================================================
; Section .....: Auto-Execution
; ----------------------------------------------------------------------------------
; -=================================================================================
#Warn All, OutputDebug
#SingleInstance Force
SendMode("Input")  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.
SetTitleMatchMode(2) ; sets title matching to search for "containing" instead of "exact"
DetectHiddenText(true)
DetectHiddenWindows(true)
#Requires Autohotkey v2.0+

; ReloadAllAhkScripts()
; ;Reload
; ;return
; ;#IfWinExist
; return

/*
reloadall:
DetectHiddenWindows, On
SetTitleMatchMode, 2
WinGet, allAhkExe, List, ahk_class AutoHotkey
Loop, % allAhkExe {
	hwnd := allAhkExe%A_Index%
	if (hwnd = A_ScriptHwnd or hwnd = Windows.ahk or hwnd = QuickAccessPopUp)  ; ignore the current window for reloading
	{
		Sleep 100
		continue
	}
	PostMessage, 0x0111, 65303,,, % "ahk_id" . hwnd
}	
Reload
return
*/
;~^s:
;#IfWinActive ahk_exe AutoHotkeyU32.exe

;DetectHiddenWindows, On
;SetTitleMatchMode, 2
;Loop, % allAhkExe {
		;hwnd := allAhkExe%A_Index%
		;if (hwnd = A_ScriptHwnd)  ; ignore the current window for reloading
		;{
			;continue
		;}
		;PostMessage, 0x0111, 65303,,, % "ahk_id" . hwnd
	;}	
;Reload
;#IfWinActive
;return
/*
#Requires AutoHotkey v1.1.33
skip := "|FindText|SimpleSpy|"
SoundBeep 1500
/*
~^s::
{ ; V1toV2: Added bracket
DetectHiddenWindows On
WinGet running, List, ahk_class AutoHotkey
Loop % running
	{
		WinGetTitle title, % "ahk_id" running%A_Index%
		scriptPath := SubStr(title, 1, InStr(title, " - AutoHotkey") - 1)
		SplitPath scriptPath,,,, fnBare
		If !InStr(skip, "|" fnBare "|") && scriptPath != A_LineFile
			Run % scriptPath " /r"
}
Reload
Return
*/

/*
arr := []
for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where name = 'Autohotkey.exe'")
 		if !instr(strsplit(trim(process.CommandLine," """), """").pop(), A_ScriptName) 
			{
				arr.push(strsplit(trim(process.CommandLine," """), """").pop())
			process, close, % process.ProcessId
		}
		;sleep, 10000					; remove ; to test that the running scripts really close
		for each, scrpt in arr
			run, % scrpt
		*/
		;DetectHiddenWindows, On
		;SetTitleMatchMode, 2
		/*
		WinGet running, List, ahk_class AutoHotkey
		Loop % running
			{
				WinGetTitle title, % "ahk_id " running%A_Index%
				scriptPath := SubStr(title, 1, InStr(title, " - AutoHotkey") - 1)
				SplitPath scriptPath,,,, fnBare
				If !InStr(skip, "|" fnBare "|") && scriptPath != A_LineFile
					;Run % scriptPath " /r"
				PostMessage, 0x0111, 65303,,, % "ahk_id " . hwnd
			}
			;Reload
			;*/
ReloadAllAhkScripts()
{
	DetectHiddenWindows(true)
	oList := WinGetList("ahk_class AutoHotkey",,,)
	aList := Array()
	List := oList.Length
	For v in oList
	{
	   aList.Push(v)
	}
	scripts := ""
	Loop aList.Length {
	   title := WinGetTitle("ahk_id" aList[A_Index])
	   scripts .=  (scripts ? "`r`n" : "") . RegExReplace(title, " - AutoHotkey v[\.0-9]+$")
	}
	If true ;( targetProcessID = DllCall('GetCurrentProcessId', 'Ptr') ) 
		PostMessage(0x111,65405,0,,"ahk_id " List%A_Index%)
}
return
/*
	oallAhkExe := WinGetList("ahk_class AutoHotkey",,,)
	aallAhkExe := Array()
	allAhkExe := oallAhkExe.Length
	For v in oallAhkExe
	{
		aallAhkExe.Push(v)
	}
	;OutputDebug % allAhkExe
	Loop aallAhkExe.Length
		{
		hwnd := aallAhkExe[A_Index]
		for item in hwnd
			aTitle := WinGetTitle("ahk_id " hwnd)
		OutputDebug(aTitle . "`n")
		 ; ignore the current window for reloading
		if (aTitle = "Scriptlet_Library_v4.ahk")
			{
				Continue
			}
			PostMessage(0x0111, 65303, , , "ahk_id" . hwnd)
		}
}
;#HotIf