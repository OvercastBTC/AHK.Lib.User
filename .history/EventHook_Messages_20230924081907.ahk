#Include <Directives\__AE.v2>
#Requires AutoHotkey v2
; #Requires AutoHotKey v2.0-beta.3
; #SingleInstance force
ProcessSetPriority("High")

; Created by AHK_user
; Based on WinEventHook Messages v0.3 by Serenity  https://www.autohotkey.com/board/topic/32662-tool-wineventhook-messages/
; But with extra features:
; - Switching between window messages and shell messages
; - Copy example code to get started (Rightclick on row to see the options)
; --------------------------------------------------------------------------------
; order of messagelist is important, do not change
winMessageList := "SOUND,ALERT,FOREGROUND,MENUSTART,MENUEND,MENUPOPUPSTART,MENUPOPUPEND,CAPTURESTART,CAPTUREEND,MOVESIZESTART,MOVESIZEEND,CONTEXTHELPSTART,CONTEXTHELPEND,DRAGDROPSTART,DRAGDROPEND,DIALOGSTART,DIALOGEND,SCROLLINGSTART,SCROLLINGEND,SWITCHSTART,SWITCHEND,MINIMIZESTART,MINIMIZEEND"
shellMessageList := "WINDOWCREATED,WINDOWDESTROYED,ACTIVATESHELLWINDOW,WINDOWACTIVATED,GETMINRECT,REDRAW,TASKMAN,LANGUAGE,SYSMENU,ENDTASK,ACCESSIBILITYSTATE,APPCOMMAND,WINDOWREPLACED,WINDOWREPLACING,HIGHBIT,FLASH,RUDEAPPACTIVATED"

ExcludeScriptMessages := 1 ; 0 to include
Title := "WinEventHook Messages", Filters := "", PauseStatus := 0
GuiWinEventHook()
WM_VSCROLL := 0x115, SB_BOTTOM := 7

HookWinEvent()
; mygui := Gui()
; #HotIf WinActive( "ahk_class Notepad")
; #HotIf WinActive(mygui.Name)
Hotkey("~RButton", RClick_LV)

FilterMessageSetAll(Status := 1) {
    global myGui, winMessageList, shellMessageList
    PreString := mygui.mode = "WinEventHook" ? "EVENT_SYSTEM_" : "HSHELL_"
    MessageList := mygui.mode = "ShellEventHook" ? shellMessageList : winMessageList
    loop parse MessageList, ","
    {
        myGui.filterMsg[PreString A_LoopField] := Status
        if Status {
            FilterMenu.Check(A_LoopField)
        } else {
            FilterMenu.UnCheck(A_LoopField)
        }
    }
    UpdateMessages()
    return
}

FilterMessage(Message, Status := 0) {
    global myGui
    PreString := mygui.mode = "WinEventHook" ? "EVENT_SYSTEM_" : "HSHELL_"
    myGui.filterMsg[Message] := Status
    if Status {
        FilterMenu.Check(StrReplace(Message, PreString))
    } else {
        FilterMenu.UnCheck(StrReplace(Message, PreString))
    }
    UpdateMessages()
}

UpdateMessages() {
    global myGui, ogLV_WMessages, ogLV_ShellMessages
    ActiveListview := mygui.mode = "WinEventHook" ? ogLV_WMessages : ogLV_ShellMessages
    ColMessages := mygui.mode = "WinEventHook" ? 7 : 4
    ColPhwnd := mygui.mode = "WinEventHook" ? 9 : 1
    ListViewCount := ActiveListview.GetCount()
    ActiveListview.Opt("-Redraw")
    Loop ListViewCount {
        RowMessage := ActiveListview.GetText(ListViewCount - A_Index + 1, ColMessages)
        if (myGui.filterMsg.Has(RowMessage) and myGui.filterMsg[RowMessage]) {
            ActiveListview.Delete(ListViewCount - A_Index + 1) ; Select each row whose first field contains the filter-text.
        } else if (myGui.ExcludeOwnMessages) {
            RowphWnd := ActiveListview.GetText(ListViewCount - A_Index + 1, ColPhwnd)
            if (myGui.phwnd + 0 = RowphWnd) {
                ActiveListview.Delete(ListViewCount - A_Index + 1) ; Select each row whose first field contains the filter-text.
            }
        }
    }
    ActiveListview.Opt("+Redraw")
}

RClick_LV(p*) {
    MouseGetPos(, , , &ControlHwnd, 2)
    if (ogLV_WMessages.Hwnd = ControlHwnd) {
        selMessage := ogLV_WMessages.GetText(ogLV_WMessages.GetNext(, "F"), 7)
        selEvent := ogLV_WMessages.GetText(ogLV_WMessages.GetNext(, "F"), 6)
        selMessageText := StrReplace(selMessage, "EVENT_SYSTEM_")
        myMenu := Menu()

        omi_Filter_Process := myMenu.Add("Filter process", (*) => (myGui.HookSelected := !myGui.HookSelected, NewHook()))
        if (myGui.HookSelected) {
            myMenu.Check("Filter process")
        }
        if (selMessageText != "") {
            myMenu.Add("Filter message [" selMessageText "]", (*) => (selMessage, FilterMessageSetAll(), FilterMessage(selMessage)))
            myMenu.Add()
            myMenu.Add("Exclude message [" selMessageText "]", (*) => (selMessage, FilterMessage(selMessage, 1)))
        }
        myMenu.Add()
        myMenu.Add("Copy Event [" selEvent "]", (*) => (A_Clipboard := selEvent, Tooltip2("Copied [" A_Clipboard "]")))
        myMenu.Add("Copy message [" selMessage "]", (*) => (A_Clipboard := selMessage, Tooltip2("Copied [" A_Clipboard "]")))
        myMenu.Add()
        myMenu.Add("Copy example code", (*) => (A_Clipboard := CreateWinEventCode(selEvent, selMessage), Tooltip2("Copied:`n" A_Clipboard)))
        myMenu.Show()
    } else if (ogLV_ShellMessages.Hwnd = ControlHwnd) {
        selEvent := ogLV_ShellMessages.GetText(ogLV_ShellMessages.GetNext(, "F"), 3)
        selMessage := ogLV_ShellMessages.GetText(ogLV_ShellMessages.GetNext(, "F"), 4)
        selMessageText := StrReplace(selMessage, "EVENT_SYSTEM_")
        myMenu := Menu()

        omi_Filter_Process := myMenu.Add("Filter process", (*) => (myGui.HookSelected := !myGui.HookSelected))
        if (myGui.HookSelected) {
            myMenu.Check("Filter process")
        }
        if (selMessageText != "") {
            myMenu.Add("Filter message [" selMessageText "]", (*) => (selMessage, FilterMessageSetAll(), FilterMessage(selMessage)))
            myMenu.Add()
            myMenu.Add("Exclude message [" selMessageText "]", (*) => (selMessage, FilterMessage(selMessage, 1)))
        }
        myMenu.Add()
        myMenu.Add("Copy Event [" selEvent "]", (*) => (A_Clipboard := selEvent, Tooltip2("Copied [" A_Clipboard "]")))
        if selMessage != "" {
            myMenu.Add("Copy message [" selMessage "]", (*) => (A_Clipboard := selMessage, Tooltip2("Copied [" A_Clipboard "]")))
            myMenu.Add()
        }
        selMessage := selMessage = "" ? "ShellEvent" : selMessage
        myMenu.Add("Copy example code", (*) => (A_Clipboard := CreateShellEventCode(selEvent, selMessage), Tooltip2("Copied:`n" A_Clipboard)))
        myMenu.Show()
    }

}

GuiWinEventHook() {
    global ogLV_WMessages, ogLV_ShellMessages, winMessageList, shellMessageList, myGui, PauseStatus, FilterMenu, MyMenuBar
    title := "WinEventHook Messages"
    myGui := Gui()
    myGui.Title := Title
    myGui.Opt("+LastFound +AlwaysOnTop +Resize") ; +ToolWindow
    myGui.MarginX := "0", myGui.MarginY := "0"
    myGui.SetFont("s8", "Microsoft Sans Serif")
    myGui.HookSelected := 0
    myGui.ExcludeOwnMessages := 0
    myGui.BackColor := ""
    myGui.filterMsg := Map()
    myGui.Mode := "WinEventHook"
    myGui.phwnd := WinGetPID(myGui.hwnd)
    ogLV_WMessages := myGui.Add("ListView", "w600 r10 +Grid +NoSort", ["Hwnd", "idObject", "idChild", "Title", "Class", "Event", "Message", "Process", "idProcess", "WinTitle"])
    ogLV_WMessages.Opt("Count1000 -Multi")
    ogLV_WMessages.ModifyCol(1, 60), ogLV_WMessages.ModifyCol(2, 40), ogLV_WMessages.ModifyCol(3, 40)
    ogLV_WMessages.ModifyCol(4, 100), ogLV_WMessages.ModifyCol(5, 100), ogLV_WMessages.ModifyCol(7, 190)
    ogLV_WMessages.visible := 1

    ogLV_ShellMessages := myGui.Add("ListView", "x0 y0 w400 r10 +Grid +NoSort", ["lParam (hWnd)", "Process", "wParam (Event)", "Message", "WinTitle"])
    ogLV_ShellMessages.ModifyCol(1, 85), ogLV_ShellMessages.ModifyCol(2, 100), ogLV_ShellMessages.ModifyCol(3, 60), ogLV_ShellMessages.ModifyCol(4, 180), ogLV_ShellMessages.ModifyCol(5, 180)
    ogLV_ShellMessages.visible := 0

    FilterMenu := Menu()

    ; FilterMenu.Add("Filter All", (*) => (FilterMessageSetAll(1)))
    ; FilterMenu.Add("Filter None", (*) => (FilterMessageSetAll(0)))
    ; FilterMenu.Add()
    ; FilterMenu.Add("Exclude Own Messages", (ItemName, ItemPos, ItemMenu) => (FilterMenu.ToggleCheck(ItemName), myGui.ExcludeOwnMessages := !myGui.ExcludeOwnMessages, UpdateMessages()))
    ; FilterMenu.Add()
    ; loop parse winMessageList, ","
    ; {
    ;     myGui.filterMsg["EVENT_SYSTEM_" A_LoopField] := 0
    ;     FilterMenu.Add(A_LoopField,(ItemName, ItemPos, ItemMenu) => (FilterMessage("EVENT_SYSTEM_" ItemName,!myGui.filterMsg["EVENT_SYSTEM_" ItemName]) ))
    ; }

    ListMenu := Menu()
    ListMenu.Add("Clear All", (*) => (ogLV_WMessages.Delete()))
    SettingsMenu := Menu()

    MyMenuBar := MenuBar()

    MyMenuBar.Add("Filter", FilterMenu)
    MyMenuBar.Add("List", ListMenu)
    Gui_BuildMenu()
    MyMenuBar.Add("Pause", (ItemName, ItemPos, ItemMenu) => (Gui_Pause()))

    MyMenuBar.Add("&Reload", (*) => Reload())
    MyMenuBar.Add("ShellEventHook", (*) => ToggleModus())
    MyGui.MenuBar := MyMenuBar

    MyGui.OnEvent("Size", Gui_Size)
    MyGui.OnEvent("Close", Gui_Close)
    myGui.Show()

    Gui_Size(thisGui, MinMax, Width, Height) {
        WinGetClientPos(&XWin, &YWin, &wWin, &hWin, thisGui)
        ogLV_WMessages.Move(, , wWin, hWin)
        ogLV_ShellMessages.Move(, , wWin, hWin)
    }

    ToggleModus() {

        if (mygui.mode = "ShellEventHook") {
            MyMenuBar.Rename("WinEventHook", "ShellEventHook")
            myGui.Mode := "WinEventHook"
            ogLV_WMessages.visible := 1
            ogLV_ShellMessages.visible := 0
            HookWinEvent()
            UnhookShellEvent()
            Gui_BuildMenu()
        } else {
            MyMenuBar.Rename("ShellEventHook", "WinEventHook")
            myGui.Mode := "ShellEventHook"
            ogLV_WMessages.visible := 0
            ogLV_ShellMessages.visible := 1
            UnhookWinEvent()
            HookShellEvent()
            Gui_BuildMenu()
        }
        myGui.Title := myGui.Mode " messages"
    }

    Gui_Pause() {
        PauseStatus := !PauseStatus
        if (PauseStatus) {
            MyMenuBar.Rename("Pause", "Start")
            if (mygui.mode = "ShellEventHook") {
                UnhookShellEvent()
            } else {
                UnhookWinEvent()
            }
            myGui.Title := myGui.Mode " messages - PAUSE"
        } else {
            MyMenuBar.Rename("Start", "Pause")
            if (mygui.mode = "ShellEventHook") {
                HookShellEvent()
            } else {
                HookWinEvent()
            }
            myGui.Title := myGui.Mode " messages"
        }
    }

    Gui_BuildMenu() {
        global MyMenuBar
        try MyMenuBar.Delete("Filter")
        FilterMenu.Delete()
        FilterMenu.Add("Filter All", (*) => (FilterMessageSetAll(1)))
        FilterMenu.Add("Filter None", (*) => (FilterMessageSetAll(0)))
        FilterMenu.Add()
        FilterMenu.Add("Exclude Own Messages", (ItemName, ItemPos, ItemMenu) => (FilterMenu.ToggleCheck(ItemName), myGui.ExcludeOwnMessages := !myGui.ExcludeOwnMessages, UpdateMessages()))
        FilterMenu.Add()
        MyMenuBar.Insert("List", "Filter", FilterMenu)
        MessageList := mygui.mode = "ShellEventHook" ? shellMessageList : winMessageList
        PreString := mygui.mode = "WinEventHook" ? "EVENT_SYSTEM_" : "HSHELL_"
        loop parse MessageList, ","
        {
            if (!myGui.filterMsg.Has(PreString A_LoopField)) {
                myGui.filterMsg[PreString A_LoopField] := 0
            }
            FilterMenu.Add(A_LoopField, (ItemName, ItemPos, ItemMenu) => (FilterMessage(PreString ItemName, !myGui.filterMsg[PreString ItemName])))
        }

    }
}

CaptureWinEvent(hWinEventHook, Event, hWnd, idObject, idChild, dwEventThread, dwmsEventTime) {
    Global PauseStatus, WM_VSCROLL, SB_BOTTOM, ogLV_WMessages, winMessageList

    Event += 0
    message := ""

    if (Event = 1)
        Message := "EVENT_SYSTEM_SOUND"
    else if (Event = 2)
        Message := "EVENT_SYSTEM_ALERT"
    else if (Event = 3)
        Message := "EVENT_SYSTEM_FOREGROUND"
    else if (Event = 4)
        Message := "EVENT_SYSTEM_MENUSTART"
    else if (Event = 5)
        Message := "EVENT_SYSTEM_MENUEND"
    else if (Event = 6)
        Message := "EVENT_SYSTEM_MENUPOPUPSTART"
    else if (Event = 7)
        Message := "EVENT_SYSTEM_MENUPOPUPEND"
    else if (Event = 8)
        Message := "EVENT_SYSTEM_CAPTURESTART"
    else if (Event = 9)
        Message := "EVENT_SYSTEM_CAPTUREEND"
    else if (Event = 10)
        Message := "EVENT_SYSTEM_MOVESIZESTART"
    else if (Event = 11)
        Message := "EVENT_SYSTEM_MOVESIZEEND"
    else if (Event = 12)
        Message := "EVENT_SYSTEM_CONTEXTHELPSTART"
    else if (Event = 13)
        Message := "EVENT_SYSTEM_CONTEXTHELPEND"
    else if (Event = 14)
        Message := "EVENT_SYSTEM_DRAGDROPSTART"
    else if (Event = 15)
        Message := "EVENT_SYSTEM_DRAGDROPEND"
    else if (Event = 16)
        Message := "EVENT_SYSTEM_DIALOGSTART"
    else if (Event = 17)
        Message := "EVENT_SYSTEM_DIALOGEND"
    else if (Event = 18)
        Message := "EVENT_SYSTEM_SCROLLINGSTART"
    else if (Event = 19)
        Message := "EVENT_SYSTEM_SCROLLINGEND"
    else if (Event = 20)
        Message := "EVENT_SYSTEM_SWITCHSTART"
    else if (Event = 21)
        Message := "EVENT_SYSTEM_SWITCHEND"
    else if (Event = 22)
        Message := "EVENT_SYSTEM_MINIMIZESTART"
    else if (Event = 23)
        Message := "EVENT_SYSTEM_MINIMIZEEND"

    Sleep(50)
    EventHex := Event

    if (message != "") {
        try {
            if (myGui.filterMsg.Has(message) and myGui.filterMsg[message] = 1) {
                return
            }
            ; give a little time for WinGetTitle/WinGetActiveTitle functions, otherwise they return blank
            WinhWnd := WinGetTitle(hWnd) = "" ? DllCall("user32\GetAncestor", "Ptr", hWnd, "UInt", 1, "Ptr") : hWnd
            phWnd := WinGetPID(WinhWnd)
            if (myGui.ExcludeOwnMessages and myGui.phwnd + 0 = phWnd) {
                return
            }
            WinClass := WinGetClass(hWnd)
            ogLV_WMessages.Add("", format("0x{:x}", hWnd), format("0x{:x}", idObject), format("0x{:x}", idChild), WinGetTitle(hWnd), WinClass, format("0x{:x}", EventHex), Message, WinGetProcessName(hWnd), format("0x{:x}", phWnd), WinGetTitle(WinhWnd))

            if (!WinActive(myGui)) {
                SendMessage(WM_VSCROLL, SB_BOTTOM, 0, "SysListView321", "ahk_id " myGui.Hwnd)
            }
        }
    }
}

HookWinEvent() {
    global
    HookProcAdr := CallbackCreate(CaptureWinEvent, "F")
    dwFlags := (ExcludeScriptMessages = 1 ? 0x1 : 0x0)
    hWinEventHook := SetWinEventHook(0x1, 0x17, 0, HookProcAdr, 0, 0, 0)
}

SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags) {
    DllCall("ole32\CoInitialize", "Uint", 0)
    return DllCall("SetWinEventHook", "Uint", eventMin, "Uint", eventMax, "Uint", hmodWinEventProc, "Uint", lpfnWinEventProc, "Uint", idProcess, "Uint", idThread, "Uint", dwFlags)
}

NewHook(p*) {
    Global

    if (myGui.HookSelected = 1) {
        SelHwnd := ogLV_WMessages.GetText(ogLV_WMessages.GetNext(0, "Focused"))
        idProcess := WinGetPID("ahk_id " SelHwnd)
        if (idProcess = "") {
            myGui.HookSelected := !myGui.HookSelected
            Return
        }
    } Else {
        idProcess := "0" ; hook all
    }
    UnhookWinEvent()
    HookProcAdr := CallbackCreate(CaptureWinEvent, "F") ; new hook
    hWinEventHook := SetWinEventHook(0x1, 0x17, 0, HookProcAdr, idProcess, 0, dwFlags)
}

UnhookWinEvent() {
    Global
    DllCall("UnhookWinEvent", "Uint", hWinEventHook)
    DllCall("GlobalFree", "UInt", HookProcAdr) ; free up allocated memory for RegisterCallback
}


Gui_Close(p*) {
    UnhookWinEvent()
    loop {
        Exitapp()
    }
}


CaptureShellMessage(wParam, lParam, msg, hwnd) {
    Global PauseStatus, WM_VSCROLL, SB_BOTTOM, ogLV_WMessages, winMessageList, myGui

    try {
        pname := WinGetProcessName("ahk_id " lParam)
    } Catch {
        pname := ""
    }
    msg := ""
    if (wParam = 1)
        msg := "HSHELL_WINDOWCREATED"
    if (wParam = 2)
        msg := "HSHELL_WINDOWDESTROYED"
    if (wParam = 3)
        msg := "HSHELL_ACTIVATESHELLWINDOW"
    if (wParam = 4)
        msg := "HSHELL_WINDOWACTIVATED"
    if (wParam = 5)
        msg := "HSHELL_GETMINRECT"
    if (wParam = 6)
        msg := "HSHELL_REDRAW"
    if (wParam = 7)
        msg := "HSHELL_TASKMAN"
    if (wParam = 8)
        msg := "HSHELL_LANGUAGE"
    if (wParam = 9)
        msg := "HSHELL_SYSMENU"
    if (wParam = 10)
        msg := "HSHELL_ENDTASK"
    if (wParam = 11)
        msg := "HSHELL_ACCESSIBILITYSTATE"
    if (wParam = 12)
        msg := "HSHELL_APPCOMMAND"
    if (wParam = 13)
        msg := "HSHELL_WINDOWREPLACED"
    if (wParam = 14)
        msg := "HSHELL_WINDOWREPLACING"
    if (wParam = 15)
        msg := "HSHELL_HIGHBIT"
    if (wParam = 16)
        msg := "HSHELL_FLASH"
    if (wParam = 17)
        msg := "HSHELL_RUDEAPPACTIVATED"

    if (myGui.filterMsg.Has(msg) and myGui.filterMsg[msg] = 1) {
        return
    }
    ; give a little time for WinGetTitle/WinGetActiveTitle functions, otherwise they return blank
    ; WinhWnd := WinGetTitle(hWnd) = "" ? DllCall("user32\GetAncestor", "Ptr", hWnd, "UInt", 1, "Ptr") : hWnd
    phWnd := WinGetPID("ahk_id" hwnd)
    WinTitle := WinGetTitle("ahk_id " hwnd)

    if (myGui.ExcludeOwnMessages and myGui.phwnd + 0 = phWnd) {
        return
    }

    ogLV_ShellMessages.Add("", format("0x{:x}", lParam), pname, format("0x{:x}", wParam), msg, Wintitle)

    if (!WinActive(myGui)) {
        SendMessage(WM_VSCROLL, SB_BOTTOM, 0, "SysListView322", "ahk_id " myGui.Hwnd)
    }

}

HookShellEvent() {
    global
    DllCall("RegisterShellHookWindow", "UInt", myGui.Hwnd)
    MsgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")
    OnMessage(MsgNum, CaptureShellMessage)
}

UnHookShellEvent() {
    global
    OnMessage(MsgNum, CaptureShellMessage, 0)
}

CreateWinEventCode(Event, Message) {
    Code := '; Callback definition for ' Message '`n'
    Code .= 'Hook_' StrReplace(Message, "EVENT_SYSTEM_", "ES_") ' := DllCall("SetWinEventHook", "UInt", ' Event ', "UInt", ' Event ', "UInt", 0, "UInt", CallbackCreate(CallBack_' StrReplace(Message, "EVENT_SYSTEM_", "ES_") '), "UInt", 0, "UInt", 0, "UInt", 0)`n`n'
    Code .= 'CallBack_' StrReplace(Message, "EVENT_SYSTEM_", "ES_") '(hWinEventHook, Event, hWnd, idObject, idChild, dwEventThread, dwmsEventTime){`n`t;Enter code here to trigger on the message`n}`n`n'
    Code .= '; Code to unhook the callback`n'
    Code .= 'DllCall("UnhookWinEvent", "Ptr", Hook_' StrReplace(Message, "EVENT_SYSTEM_", "ES_") ')'
    return Code
}
CreateShellEventCode(Event, Message) {
    CallBackName := "CallBack_" Message
    Code := '; Callback definition for ' Message '`n'
        . 'DllCall("RegisterShellHookWindow", "UInt", myGui.Hwnd)`n'
        . 'MsgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")`n'
        . 'OnMessage(MsgNum, ' CallBackName ')`n`n'
        . CallBackName '(wParam, lParam, msg, hwnd){`n'
        . '   if (wParam = ' Event '){ `; ' Message '`n'
        . '        `; Enter your code here`n'
        . '   } `n'
        . '}`n`n'
        . '; Code to unhook the callback`n'
        . 'OnMessage(MsgNum, ' CallBackName ', 0)`n'

    return Code
}

Tooltip2(Text := "") {
    ToolTip(Text)
    SetTimer () => ToolTip(), -5000
}
