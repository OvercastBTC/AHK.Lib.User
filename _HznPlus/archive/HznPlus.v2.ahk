/************************************************************************
 * ;Desc ..: Horizon Button ==> A Horizon function library
 * @description ..: This library is a collection of functions and buttons that deal with missing interfaces with Horizon.
 * @file HznPlus.v2.ahk
 * @author OvercastBTC
 * @date 2023.08.15
 * @version 2.0.5
 * @ahkversion v2+
 * @Section .....: Auto-Execution
 ***********************************************************************/
; --------------------------------------------------------------------------------
/************************************************************************
 * ;Description ..: Resource includes for .exe standalone
 * @author OvercastBTC
 * @date 2023.08.15
 * @version 2.0.8
 ***********************************************************************/
;@Ahk2Exe-SetMainIcon 512.png
;@Ahk2Exe-AddResource 512.png, 160  ; Replaces 'H on blue'
;@Ahk2Exe-AddResource 512.png, 206  ; Replaces 'S on green'
;@Ahk2Exe-AddResource 512.png, 207  ; Replaces 'H on red'
;@Ahk2Exe-AddResource 512.png, 208  ; Replaces 'S on red'
;@include-winapi
; --------------------------------------------------------------------------------
#MaxThreads 255 ; Allows a maximum of 255 instead of default threads.
#Warn All, OutputDebug
#SingleInstance Force
SendMode("Input") ;* Superior speed and reliability.
SetWorkingDir(A_ScriptDir) ; A consistent starting directory.
SetTitleMatchMode(2) ; Match = "containing" instead of "exact"
DetectHiddenText(true)
DetectHiddenWindows(true)
SetMouseDelay(-1)
; CoordMode("Mouse", "Client")
#Requires Autohotkey v2
#Include <gdi_plus_plus>
#Include <Gdip_All>
#Include <Class_Toolbar.c2v2>
#Include <Tools\Hider>
#Include <Lib.v2\UIA>
; #Include <Lib.v2\UIA_Browser>
; --------------------------------------------------------------------------------

; //TODO: 2023.07.17 ...: Work on the below but consider if needed due to new FreeLibraryAndExitThread
; //TODO: Library_Load(winuser.dll)
; //TODO: Library_Load(processthreadsapi.dll)
; //TODO: Library_Load(memoryapi.dll)

; Startup_Shortcut := A_Startup "\" A_ScriptName ".lnk"
Startup_Shortcut := A_Startup "\" A_ScriptName
If !(FileExist(Startup_Shortcut)){
	FileCopy(A_ScriptName,Startup_Shortcut)
	MsgBox("Shortcut added to your Startup folder at " Startup_Shortcut)
} 
; Else {
; 	MsgBox("Shortcut Exists")
; }
; --------------------------------------------------------------------------------
/************************************************************************
* @Title .........: Create Tray Icon
* @description ...: Create the tray icon using the embedded B64 via Create_HznHorizon_ico()
* @author OverastBTC
* @date 2023.08.15
* @version 2.0.1
 ***********************************************************************/
TraySetIcon('HICON:' Create_HznHorizon_ico())
; --------------------------------------------------------------------------------

#HotIf WinActive(A_ScriptName " - Visual Studio Code")
#Include <Abstractions\Script>
~^s::
{
	Script.Reload()
}
#HotIf
#HotIf !ProcessGetName("hznhorizon.exe")
	SetWinDelay(-1)
	SetControlDelay(-1)
#HotIf

; -------------------------------------------------------------------------------------------------
/************************************************************************
* @Title .........: Create Tray Menu
* @Description ...: Create options (//TODO add functions here) for use with the script (e.g., run at startup, open horizon, etc.)
* @file HznPlus.v2.ahk
* @author OverastBTC
* @date 2023.08.09
* @version 0.0.1
* @TODO ;[ ] TODO	.: Convert from v1 to v2
 ***********************************************************************/
; -------------------------------------------------------------------------------------------------
; Sub-Section .....: Create Tray Menu Functions
; Description .....: addTrayMenuOption() ; madeBy() ; runAtStartup() ; trayNotify()


; HznTray := A_TrayMenu
; see at the bottom for the functions
return
; --------------------------------------------------------------------------------------------------
/************************************************************************
;Description: Horizon Hotkeys
;Description: Hotkeys (shortcuts) for normal Windows hotkeys that should exist
@author OvercastBTC
@notes
@function Italics - (CTRL+I)
@function Bold - (CTRL+B)
@function Underline - (CTRL+U) (where applicable)
@function SelectAll - (CTRL+A)
@function Save - (CTRL+S)
***********************************************************************/
; -------------------------------------------------------------------------------------------------
/************************************************************************
* ;desc ......: Horizon Save Button (Ctrl-S)
* @desc ......: Use UIAutomation (UIA) to find the element where the save button exists
* @author OvercastBTC, Descolada and his UIA librarys
* @function HznSave() - (Ctrl-S)
* @param Msg - EM_SETSEL := 177 - the Windows API message for "Set Selection"
* @param wParam - := 0
* @param lParam := -1
***********************************************************************/
; -------------------------------------------------------------------------------------------------

#HotIf WinActive("ahk_exe hznhorizon.exe")

^i::button()
^b::button()
^u::button()
^a::HznSelectAll()
^s::HznSave()
#HotIf

; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
/************************************************************************
* @desc ......: Horizon Save Button (Ctrl-S)
* @Desc ......: Use UIAutomation (UIA) to find the element where the save button exists
* @Author OvercastBTC 
* @Credit Descolada and his UIA library and functions
* @Function HznSave() (Ctrl-S)
* @param Msg - EM_SETSEL := 177 - the Windows API message for "Set Selection"
* @param wParam - := 0
* @param lParam := -1
***********************************************************************/
; --------------------------------------------------------------------------------

HznSave()
{
	global IUIAutomationActivateScreenReader:=0
	global IUIAutomationMaxVersion:=0
	idWin := WinGetID("[Main] ahk_exe hznHorizon.exe")
	idControl := ControlGetHwnd("msvb_lib_toolbar1")
	hznWin := UIA.ElementFromHandle(idWin)
	hznMenuBar := HznWin.FindElement({
		LocalizedType: "menu bar",
		Name: "System"})
	hznSave := hznMenuBar.Highlight().ControlClick()
	return
}

; -------------------------------------------------------------------------------------------------
/**
 * Desc ......: Horizon Hotkey - Select-All (Ctrl-A)
 * @author ...: Descolada, OvercastBTC
 * Function ..: Select-All() (Ctrl-A)
 * @param Msg - EM_SETSEL := 177 - the Windows API message for "Set Selection"
 * @param wParam - := 0
 * @param lParam := -1
 */
; -------------------------------------------------------------------------------------------------
HznSelectAll()
{
	Static Msg := EM_SETSEL := 177, wParam := 0, lParam := -1
	fCtl := ControlGetClassNN(ControlGetFocus("A"))
	hCtl := ControlGethWnd(fCtl, "A")
	DllCall('SendMessage', 'Ptr', hCtl, 'UInt', Msg, 'UInt', wParam, 'UIntP', lParam)
}
; -------------------------------------------------------------------------------------------------
/**
 * Desc ......:  Call the Horizon msvb_lib_toolbar buttons using the button() function
 * @author ....: Descolada, OvercastBTC
 * @function button()
 * @param A_ThisHotkey - AHK's built in variabl.
 */
; -------------------------------------------------------------------------------------------------
button(*) {
	SendLevel(5)
	BlockInput(1) ; 1 = On, 0 = Off
	fCtl := ControlGetClassNN(ControlGetFocus("A"))
	bID := SubStr(fCtl, -1, 1)
	nCtl := "msvb_lib_toolbar" bID
	hTb := ControlGethWnd(nCtl, "A")
	hTx := ControlGethWnd(fCtl, "A")
	hIDx:= A_ThisHotkey = "^i" ? 2 ; .........: italic = 2
		:  A_ThisHotkey = "^b" ? 1 ; .........: bold = 1
		:  A_ThisHotkey = "^u" ? 10 ; ........: (AJB - 08/2023) FE Notepad, Comments, and COPE => u = 10; 3 worked somewhere? (???9 and 10???) (else i or b)
		:  A_ThisHotkey = "^x" ? 11 ; ........: cut = 11 and 12
		:  A_ThisHotkey = "^c" ? 13 ; ........: copy
		:  A_ThisHotkey = "^v" ? 16 ; ........: paste
		:  A_ThisHotkey = "^z" ? 17 ; ........: undo = 17 and 18
		:  A_ThisHotkey = "^y" ? 20 : 21 ; ...: redo
	HznButton(hTb,hIDx,nCtl,fCtl,hTx, bID)
	SendLevel(0)
	BlockInput(0) ; 1 = On, 0 = Off
	return
}
; --------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------
/**
 * Desc:  Clicks the nth item in a Win32 application toolbar.
 * @author ......: Descolada, OvercastBTC
 * Function .....: HznButton()
 * @param hToolbar - The handle of the toolbar control.
 * @param hTb - The handle of the toolbar control.
 * @param fCtl* - [OPTIONAL] The ClassNN of the "focused" control.
 * @param bID - The ClassNN instance of the "focused" control.
 * @param nCtl - The ClassNN and instance of the toolbar control.
 * @param n - The index (hIDx) of the toolbar item to click (1-based). Note: Separators are considered items as well.
 * @param HznButton(hToolbar, n, nCtl, fCtl:=0, hTx:=0, bID:=0)
 * @param param:=0 => in AHK v2 => optional
 * @example
 * fCtl := ControlGetClassNN(ControlGetFocus("A"))
 * bID := SubStr(fCtl, -1, 1)
 * nCtl := "msvb_lib_toolbar" bID
 * hTb := ControlGethWnd(nCtl, "A")
 * hTx := ControlGethWnd(fCtl, "A")
 * HznButton(hToolbar, 3) ; Clicks the third item
 */
; --------------------------------------------------------------------------------

HznButton(hToolbar, n, nCtl, fCtl:=0, hTx:=0, bID:=0) {
	; --------------------------------------------------------------------------------
	; Step: [OPTIONAL] Block all input while the function running
	; --------------------------------------------------------------------------------
	BlockInput(1) ; 1 = On, 0 = Off
	; --------------------------------------------------------------------------------
	static TB_BUTTONCOUNT 		:= 1048
	Static TB_GETBUTTON 		:= 1047
	Static WM_COMMAND			:= 273 ; 0x111
	Static TB_PRESSBUTTON		:= 1027 ; 0x403
	Static TB_BUTTONCOUNT  		:= 1048 ; 0x0418
	Static TB_GETBUTTON    		:= 1047 ; 0x417,
	Static TB_GETITEMRECT  		:= 1053 ; 0x41D,
	Static MEM_COMMIT      		:= 4096 ; 0x1000, ; 0x00001000, ; via MSDN Win32 
	Static MEM_RESERVE     		:= 8192 ; 0x2000, ; 0x00002000, ; via MSDN Win32
	Static MEM_PHYSICAL    		:= 4 ; 0x04    ; 0x00400000, ; via MSDN Win32
	Static MEM_PROTECT     		:= 64 ; 0x40 ;  
	Static MEM_RELEASE     		:= 32768 ; 0x8000 ; 
	Static WM_USER				:= 1024 ; 0x400
	Static TB_Something			:= WM_USER+3 ; (1042)
	Static TB_GETSTATE			:= WM_USER+18 ; (1042)
	Static TB_GETBITMAP			:= WM_USER+44 ; (1068)
	Static TB_GETBUTTONSIZE		:= WM_USER+58 ; 1082
	Static TB_GETBUTTON			:= WM_USER+23 ; 1047
	Static TB_GETBUTTONTEXTA	:= WM_USER+45 ; (1069)
	Static TB_GETBUTTONTEXTW	:= WM_USER+75 ; (1099)
	Static TB_GETITEMRECT		:= WM_USER+29 ; (1053)
	Static TB_BUTTONCOUNT		:= WM_USER+24 ; (1048)
	Static WM_GETDLGCODE		:= 135
	Static WM_NEXTDLGCTL		:= 40
	Static TB_COMMANDTOINDEX	:= 1049
	Static TB_GETBUTTONINFOW	:= 1087
	Static TB_SETSTATE 			:= 1041 ; 0x0411
	Static TBSTATE_PRESSED		:= 2 ; 0x02 
	Static WM_LBUTTONDOWN 		:= 513 ; 0x201
	Static WM_LBUTTONUP 		:= 515 ; 0x202
	Static BM_CLICK				:= 245 ; 0x000000F5
	Static TB_ISBUTTONPRESSED 	:= 1035 ; 0x040B
	; --------------------------------------------------------------------------------
	; try ControlGetPos(&ctrlx:=0, &ctrly:=0,&ctrlw,&ctrlh, hToolbar, hToolbar)
	; OutputDebug("&ctrlx: " . ctrlx " &ctrly: " . ctrly " &ctrlw: " . ctrlw " &ctrlh: " . ctrlh . "`n")
	; --------------------------------------------------------------------------------
	; Step: count and load all the msvb_lib_toolbar buttons into memory
	; --------------------------------------------------------------------------------
	buttonCount := SendMessage(TB_BUTTONCOUNT, 0, 0, , Integer(hToolbar))
	; --------------------------------------------------------------------------------
	; Step: Use the @params to press the button
	; --------------------------------------------------------------------------------
	if (n >= 1 && n <= buttonCount)
	{
		; --------------------------------------------------------------------------------
		; Step: Get the toolbar "thread" process ID (PID) 
		DllCall("GetWindowThreadProcessId", "Ptr", hToolbar, "UInt*", &targetProcessID:=0)
		; --------------------------------------------------------------------------------
		; Step: Open the target process with PROCESS_VM_OPERATION, PROCESS_VM_READ, and PROCESS_VM_WRITE access
        hProcess := DllCall("OpenProcess", "UInt", 0x0018 | 0x0010 | 0x0020, "Int", 0, "UInt", targetProcessID, "Ptr")
		; --------------------------------------------------------------------------------
		; Step: [OPTIONAL] Identify if the process is 32 or 64 bit (efficiency step)
		; --------------------------------------------------------------------------------
		A_Is64bitOS ? DllCall("IsWow64Process", "Ptr", hProcess, "Int*", Is32bit := true) : Is32bit := True
		If (A_Is64bitOS) {
			Try DllCall("IsWow64Process", "Ptr", hProcess, "Int*", Is32bit := true)
		} Else {
			Is32bit := True
		}
		; --------------------------------------------------------------------------------
		; Step: Allocate memory for the TBBUTTON structure in the target process's address space
		; --------------------------------------------------------------------------------
		RPtrSize := Is32bit ? 4 : 8
		TBBUTTON_SIZE := 8 + (RPtrSize * 3) 
		remoteMemory := DllCall("VirtualAllocEx", "Ptr", hProcess, "Ptr", 0, "UPtr", TBBUTTON_SIZE, "UInt", 0x1000, "UInt", MEM_PHYSICAL, "Ptr")
		; --------------------------------------------------------------------------------
		; Step: Get the bounds of each button (Get Item Rectangle)
		; --------------------------------------------------------------------------------
		SendMessage(TB_GETITEMRECT, n-1, remoteMemory,, hToolbar)
		RECT := Buffer(TBBUTTON_SIZE, 0)
		; Note : Winapi TBBUTTON struct(32 bytes on x64, 20 bytes on x86)
		BtnStructSize := Is32bit ? 20 : 32
		BtnStruct := Buffer(BtnStructSize, 0)
		; --------------------------------------------------------------------------------
		; Step: Read the button information stored in the RECT (remoteMemory)
		; --------------------------------------------------------------------------------
        DllCall("ReadProcessMemory", "Ptr", hProcess, "Ptr", remoteMemory, "Ptr", RECT, "UPtr", BtnStructSize, "UInt*", &bytesRead:=0, "Int")
        DllCall("VirtualFreeEx", "Ptr", hProcess, "Ptr", remoteMemory, "UPtr", 0, "UInt", 0x8000)
        DllCall("CloseHandle", "Ptr", hProcess)
		; --------------------------------------------------------------------------------
		; Step: Store previous and set min delay
		; --------------------------------------------------------------------------------
		prevDelay := A_ControlDelay
		prevMDelay := A_MouseDelay
		prevWDelay := A_WinDelay
		SetControlDelay(-1)
		SetMouseDelay(-1)
		SetWinDelay(-1)
		; --------------------------------------------------------------------------------
		Left 	:= NumGet(RECT, 0, "int")
		Top 	:= NumGet(RECT, 4, "int")
		Right 	:= NumGet(RECT, 8, "int")
		Bottom 	:= NumGet(RECT, 12, "int")
		X 		:= Left
		Y 		:= Top
		W 		:= Right-Left
		H 		:= Bottom-Top
		; --------------------------------------------------------------------------------
		ControlClick("x" ((X+W)//2) " y" ((Y+H)//2), hToolbar,,,, "NA")
		Sleep(500)
		; --------------------------------------------------------------------------------
		; --------------------------------------------------------------------------------
		; Step: Restore previous and set delay
		; --------------------------------------------------------------------------------
        SetControlDelay(prevDelay)
		SetMouseDelay(prevMDelay)
		SetWinDelay(prevWDelay)
		; --------------------------------------------------------------------------------
		OutputDebug("ButtonCount: " buttonCount "`n")
		OutputDebug("pID: " targetProcessID "`n")
		OutputDebug("remoteMemory: " remoteMemory "`n")   
		OutputDebug("hProcess: " hProcess "`nbytesRead: " bytesRead "`n")
		OutputDebug("X: " X " Y: " Y " W: " W " H: " H "`n")
		OutputDebug("x" ((X+W)//2) " y" ((Y+H)//2) "`n")
		; --------------------------------------------------------------------------------
		BlockInput(0) ; 1 = On, 0 = Off
		; --------------------------------------------------------------------------------
    } else
        ; throw ValueError("The specified index " n " is out of range. Please specify a valid index between 1 and " buttonCount ".", -1)
        throw ValueError("The specified toolbar " nCtl " was not found. Please ensure the edit field has been selected and try again.", -1)
    Return 0
}
; --------------------------------------------------------------------------------
#HotIf WinActive("ahk_exe hznhorizon.exe")
; --------------------------------------------------------------------------------------------------
/**
@SubSection .....: Horizon {Enter} ==> Select Option
@Description ....: Hotkeys (shortcuts) sending {Enter} in leu of "Double Click"
@Author OvercastBTC
; [ ] fix: .........: Still doesn't work. Might have to not use the WinActive() function
*/
; --------------------------------------------------------------------------------------------------
#HotIf WinActive("ahk_exe hznhorizonmgr.exe") or WinActive("ahk_exe hznhorizon.exe")
{
	fCtl := ControlGetClassNN(ControlGetFocus("A"))
	hCtl := ControlGethWnd(fCtl, "A")
	; Shift & Enter::DllCall("PostMessage", "Ptr", &hCtl, "UInt", 0x0203, "Ptr", 0, "Ptr", 0)
	Shift & Enter::DllCall("PostMessage", "Str", &hCtl, "UInt", 0x0203, "Ptr", 0, "Ptr", 0)
}
return
#HotIf
; --------------------------------------------------------------------------------
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