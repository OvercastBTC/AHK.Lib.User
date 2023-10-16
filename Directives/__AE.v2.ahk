/************************************************************************
 * function ......: Auto Execution (AE)
 * @description ..: A work in progress (WIP) of standard AE setup(s)
 * @file AE.v2.ahk
 * @author OvercastBTC
 * @date 2023.09.18
 * @version 1.0.0
 * @ahkversion v2+
 ***********************************************************************/
; --------------------------------------------------------------------------------
/************************************************************************
 * function ...........: Resource includes for .exe standalone
 * @author OvercastBTC
 * @date 2023.08.15
 * @version 3.0.2
 ***********************************************************************/
;@Ahk2Exe-Obey U_V, = "%A_PriorLine~U)^(.+")(.*)".*$~$2%" ? "SetVersion" : "Nop"
;@Ahk2Exe-%U_V%        %A_AhkVersion%%A_PriorLine~U)^(.+")(.*)".*$~$2%
;@include-winapi
; --------------------------------------------------------------------------------
#Warn All, OutputDebug
#SingleInstance Force
#WinActivateForce
#Requires AutoHotkey v2
; --------------------------------------------------------------------------------
#MaxThreads 255 ; Allows a maximum of 255 instead of default threads.
#MaxThreadsBuffer true
A_MaxHotkeysPerInterval := 1000
; --------------------------------------------------------------------------------
SendMode("Input")
SetWorkingDir(A_ScriptDir)
SetTitleMatchMode(2)
; --------------------------------------------------------------------------------
_AE_DetectHidden(true)
_AE_SetDelays(-1)
_AE_PerMonitor_DPIAwareness()
; --------------------------------------------------------------------------------
#HotIf WinActive(A_ScriptName ' - Visual Studio Code')
	$~^s:: Run(A_ScriptName)
#HotIf
; --------------------------------------------------------------------------------
_AE_DetectHidden(bool?)
{
    bool := true
	DetectHiddenText(bool)
    DetectHiddenWindows(bool)
}
; --------------------------------------------------------------------------------
_AE_SetDelays(n?) {
	n := -1
	SetControlDelay(n)
	SetMouseDelay(n)
	SetWinDelay(n)
}
; --------------------------------------------------------------------------------
_AE_PerMonitor_DPIAwareness()
{
	MaximumPerMonitorDpiAwarenessContext := VerCompare(A_OSVersion, ">=10.0.15063") ? -4 : -3
	Global DefaultDpiAwarenessContext := MaximumPerMonitorDpiAwarenessContext, A_DPI := DefaultDpiAwarenessContext
	try
		DllCall("SetThreadDpiAwarenessContext", "ptr", MaximumPerMonitorDpiAwarenessContext, "ptr")
	catch 
		DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
	else
		DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
	return A_DPI
}
__AE_CopyLib() {
	local Lib := A_MyDocuments '\AutoHotkey\Lib'
	If !(DirExist(Lib)) {
		DirCreate(Lib)
	}
	FileCopy(A_ScriptName, Lib A_ScriptName, 1)
}
Class AE {

	; --------------------------------------------------------------------------------
	_AE_DetectHidden(bool?){
        bool := true
        DetectHiddenText(bool)
        DetectHiddenWindows(bool)
	}
	; --------------------------------------------------------------------------------

	_AE_SetDelays(n?) {
		n := -1
		SetControlDelay(n)
		SetMouseDelay(n)
		SetWinDelay(n)
	}
	; --------------------------------------------------------------------------------
	_AE_PerMonitor_DPIAwareness() {
        MaximumPerMonitorDpiAwarenessContext := VerCompare(A_OSVersion, ">=10.0.15063") ? -4 : -3
        Global DefaultDpiAwarenessContext := MaximumPerMonitorDpiAwarenessContext, A_DPI := DefaultDpiAwarenessContext
        try
            DllCall("SetThreadDpiAwarenessContext", "ptr", MaximumPerMonitorDpiAwarenessContext, "ptr")
        catch 
            DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
        else
            DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
        return A_DPI
	}

}