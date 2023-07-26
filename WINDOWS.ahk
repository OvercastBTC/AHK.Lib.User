; ===================================================================================================
; .............: Begin Section
; Section .....: Auto-Execution
; ===================================================================================================
;SetWinDelay 0 ; ..... (AJB - 06/2023) - comment out for testing
;SetControlDelay 0 ; . (AJB - 06/2023) - comment out for testing
;SetBatchLines, 0 ; .. (AJB - 06/2023)  - comment out for testing
;SetWinDelay, -1 ; ... (AJB - 06/2023) re-enabled 06.15.23 
;SetControlDelay, -1 ; (AJB - 06/2023) re-enabled 06.15.23 
;SetBatchLines, -1 ; . (AJB - 06/2023) re-enabled 06.15.23
;#MaxMem 4095 ; Allows the maximum amount of MB per variable.
;#MaxThreads 255 ; Allows a maximum of 255 instead of default threads.
;ListHotkeys
#Warn All, OutputDebug
#NoEnv  ; Recommended for performance and compatibility
#Persistent ; Keeps script permanently running
#SingleInstance,Force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; sets title matching to search for "containing" instead of "exact"
DetectHiddenText,On
DetectHiddenWindows, On
#Requires AutoHotkey 1.1+

; ----------------------------------------------------------------------------------------------------------------------
; .............: End Sub-Section
;=======================================================================================================================
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; Menu, Tray, Icon, shell32.dll, 16 ; this changes the icon into a little laptop thing.
; Create the tray menu with the custom icon

; grouping explorers
GroupAdd, ExplorerGroup, ahk_class CabinetWClass
GroupAdd, ExplorerGroup, ahk_class #32770 ; This is for all the Explorer-based "save" and "load" boxes, from any program!

; grouping browsers
GroupAdd, BrowserGroup, ahk_exe firefox.exe
GroupAdd, BrowserGroup, ahk_exe chrome.exe
GroupAdd, BrowserGroup, ahk_exe brave.exe
GroupAdd, BrowserGroup, ahk_exe msedge.exe
GroupAdd, BrowserGroup, ahk_exe opera.exe
GroupAdd, BrowserGroup, ahk_exe iexplore.exe

; grouping terminals (WSLs as well)
GroupAdd, TerminalGroup, ahk_exe WindowsTerminal.exe
GroupAdd, TerminalGroup, ahk_exe powershell.exe
GroupAdd, TerminalGroup, ahk_exe cmd.exe
GroupAdd, TerminalGroup, ahk_exe debian.exe
GroupAdd, TerminalGroup, ahk_exe kali.exe
GroupAdd, TerminalGroup, ahk_exe ubuntu.exe

; ; for dev only
; #IfWinActive, ahk_exe code.exe
;     ~^s::
;         TrayTip, Reloading updated script, %A_ScriptName%, 1, 1
;         Sleep, 1750
;         Reload
;     return
; #IfWinActive

; variables
; directory paths
;userdir := "C:\Users\" . A_UserName . "\"
;pc := "This PC"
;desktop := userdir . "Desktop\"
;documents := userdir . "Documents\"
;downloads := userdir . "Downloads\"
;music := userdir . "Music\"
;pictures := userdir . "Pictures\"
;pictures := userdir . ""
;videos := userdir . "Videos\"
;c := "C:\"
;arlbibek := documents . A_UserName
;screenshot := userdir . "Pictures\Screenshots\"
userdir:= "C:\Users\" . A_UserName . "\",
OneDrive := "OneDrive - FM Global\",
ahk := "AHK.Main\AHK-Projects.v1.v1\",
pc:= "This PC",
desktop:= userdir . "Desktop\",
documents:= userdir . "Documents\",
downloads:= userdir . "Downloads\",
music:= userdir . "Music\",
pictures:= userdir . "Pictures\",
videos:= userdir . "Videos\",
c:= "C:\"
;arlbibek := documents . "arlbibek\"
arlbibek := documents
screenshot:= userdir . "Pictures\Screenshots\"
defaultfolder := % userdir . OneDrive . ahk ;C:\Users\bacona\OneDrive - FM Global\AHK.Main\AHK-Projects.v1\"

; script name and startup path
;splitPath, A_ScriptFullPath, , , script_ext, script_name
;global script_full_name := script_name "." script_ext
global startup_shortcut := A_Startup "\" A_ScriptName ".lnk"
; FUNCTIONS
; Run, % defaultfolder . "lib\hznbutton.ahk"
Menu, Tray, Icon, %A_ScriptDir%\assets\windows-ahk.ico
trayNotify(title, message, seconds = "", options = 0) {
    TrayTip, %title%, %message%, %seconds%, %options%
}
;WinWaitActive, AutoHotkey Dash,,3
WinActivate, AutoHotkey Dash
Send {Esc}
get_default_browser() {
    ; return ahk_exe name of the users default browser (e.g. firefox.exe)
    ; referenced from: https://www.autohotkey.com/board/topic/67330-how-to-open-default-web-browser/

    ; Find the Registry key name for the default browser
    RegRead, BrowserKeyName, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice, Progid

    ; Find the executable command associated with the above Registry key
    RegRead, BrowserFullCommand, HKEY_CLASSES_ROOT, %BrowserKeyName%\shell\open\command

    ; The above RegRead will return the path and executable name of the browser contained within quotes and optional parameters
    ; We only want the text contained inside the first set of quotes which is the path and executable
    ; Find the ending quote position (we know the beginning quote is in position 0 so start searching at position 1)
    StringGetPos, pos, BrowserFullCommand, ",,1

    ; Decrement the found position by one to work correctly with the StringMid function
    pos := --pos

    ; Extract and return the path and executable of the browser
    StringMid, BrowserPathandEXE, BrowserFullCommand, 2, %pos%
    splitPath, BrowserPathandEXE, BrowserClassEXE, , script_ext, script_name

    Return BrowserClassEXE
}

activate(program, action:="minimize", arguments:=""){
    ahk_type := "ahk_exe"
    try {
        IfWinNotExist %ahk_type% %program%
            Run %program% %arguments%
    } catch e {
        err := e.extra
        FileNotFound := "The system cannot find the file specified."
        if InStr(err, FileNotFound)
        {
            MsgBox, % "Error: " e.extra "`n" e.message "`n`nNote: Please install and/or consider adding the respective program (folder) to the PATH of the System Variables. `n(This may need system restart to take effect)"
            return
        }
        else
            MsgBox, 16,, % "Error: " e.extra "`n" e.message
    }
    if (action == "cycle"){
        program_name := StrSplit(program, ".")[1]
        group_name = %program_name%Group
        GroupAdd, %group_name%, %ahk_type% %program%
    }
    ahk_program = %ahk_type% %program%
    if WinActive(ahk_program) {
        if (action == "minimize"){
            WinMinimize, %ahk_type% %program%
        } else if (action == "cycle"){
            GroupActivate, %group_name%, r
        } else {
            MsgBox % "Error: Wrong parameter value: " action "`nThe parameter 'action' should be either 'minimize' or 'cycle'."
            return
        }
    } else {
        WinActivate, %ahk_type% %program%
    }
}

get_selected() {
    ; Save the current clipboard contents
    ClipSave := ClipboardAll
    clipboard := ""
    ; Send the correct key combination based on the active window
    if WinActive("ahk_group TerminalGroup")
        Send, ^+c
    else
        Send, ^c
    ; Wait for the clipboard to be updated
    ClipWait, 1
    copied := clipboard
    ; Restore the previous clipboard contents
    clipboard := ClipSave
    return copied
}

win_search(search_str) {
    ; Removes all CR+LF's (? next line) and extra spaces
    search_str := RegExReplace(search_str, "(\r|\n|\s{2,})")
    ; Only search if something has been selected
    If (StrLen(search_str) == 0) {
        ; Resend the command if nothing is selected
        Send, #s
        return
    }
    ; Check if the active window is a browser
    if WinActive("ahk_group BrowserGroup") {
        Send, ^t
        SendRaw, %search_str%
        Send, {Enter}
    }
    ; Check if the search string is a URL
    else if (RegExMatch(search_str, "^(https?:\/\/|www\.)"))
        Run, % search_str
    else {
        ; Replace spaces with pluses and escape special characters
        search_str := RegExReplace(search_str, " ", "+")
        search_str := RegExReplace(search_str, "[!*'();:@&=+$,\/?#[\]\\]", "\\$1")
        Run, https://duckduckgo.com/?t=ffab&q=%search_str%&atb=v292-4&ia=web
    }
}

exploreTo(path) {
    ; navigate to a path using ctrl + l in file explorer
    Send, ^l
    Sleep, 50
    SendInput, %path%
    Sleep, 50
    Send, {Enter}
}

changeCaseTo(case) {
    selected_text := get_selected()
    if (case = "lower")
        StringLower, selected_text, selected_text
    else if (case = "titled")
        StringUpper selected_text, selected_text, T
    else if (case = "upper")
        StringUpper, selected_text, selected_text
    else
        MsgBox % "Invalid parameter value: " case "`nThe parameter should be either 'lower', 'titled' or 'upper'."
    Sleep, 10
    Send, {Text}%selected_text%
    str_len := StrLen(selected_text)
    Send, +{left %str_len%}
    Send, {CapsLock}
}
/*
; For google docs/sheets
docSheetWr(text){
    ; for google docs/sheets only
    gdoc := " - Google Docs"
    gsheet := " - Google Sheets"
    WinGetActiveTitle, ActiveTitle
    if InStr(ActiveTitle, gdoc, True) || InStr(ActiveTitle, gsheet, True) {
        Send, %text%
    }
}

docWr(text){
    ; for google docs
    gdoc := " - Google Docs"
    WinGetActiveTitle, ActiveTitle
    if InStr(ActiveTitle, gdoc, True) {
        Send, %text%
    }
}

sheetWr(text){
    ; for google sheets
    gsheet := " - Google Sheets"
    WinGetActiveTitle, ActiveTitle
    if InStr(ActiveTitle, gsheet, True) {
        Send, %text%
    }
}
*/

runAtStartup() {
    if (FileExist(startup_shortcut)) {
        FileDelete, % startup_shortcut
        Menu, Tray, % "unCheck", Run at startup ; update the tray menu status on change
        trayNotify("Startup shortcut removed", "This script will not run when you turn on your computer.")
    } else {
        FileCreateShortcut, % A_ScriptFullPath, % startup_shortcut
        Menu, Tray, % "check", Run at startup ; update the tray menu status on change
        trayNotify("Startup shortcut added", "This script will now automatically run when your turn on your computer.")
    }

}

togglePresentationMode(){
    ; Toggle presentation mode
    Run, presentationsettings.exe
    WinWait ahk_exe presentationsettings.exe
    ControlGet, presentationMode, Checked, , Button1, Presentation Settings, , ,

    If (presentationMode == 1){
        ; presentationMode is on, turning it off
        Control, UnCheck , , Button1, Presentation Settings, , ,
        Menu, Tray, % "unCheck", Presentation mode {Win+Shift+P}
        trayNotify("Presentation mode: off", "Presentation mode has been toggled off.")

    } Else {
        ; presentationMode is off, turning it on
        Control, Check , , Button1, Presentation Settings, , ,
        Menu, Tray, % "check", Presentation mode {Win+Shift+P}
        trayNotify("Presentation mode: on", "Presentation mode has been toggled on, your computer will stay awake indefinitely.")
    }
    Control, Check, , Button7, Presentation Settings, , ,
}

madeBy(){
    ; visit authors website
    ;Run, https://bibeka.com.np/
    MsgBox This was made by nerds, for nerds. Regular people are ok too, lol.`nModified by Adam Bacon`nCredit:Made with ❤️ by Bibek Aryal
}

viewInGitHub(){
    ; visit source code on github
    Run, https://github.com/arlbibek/windows-ahk
}

viewAHKDoc(){
    ; view official AHK documentation
    Run, https://www.autohotkey.com/docs/AutoHotkey.htm
}

viewAHKVariables(){
    ; view official AHK built in variables documentation
    Run, https://www.autohotkey.com/docs/v1/Variables.htm
}
viewGuiDoc(){
    ; view official AHK built in Gui documentation
    Run, https://www.autohotkey.com/docs/v1/lib/Gui.htm#Labels
}
runScriptlet(){
	; run Scriptlet_Library_v4
	Run, Scriptlet_Library_v4.v1.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.Main\AHK.Projects.v1
}
runAHKScript(){
	; run the main AHK Script.ahk
	Run, AHK Script.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\AHK-Projects.v1
}
runWinspector(){
	; run WinSpectorU.exe
	Run, WinspectorU.exe, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\Winspector
}
runUIAutomation(){
	; run UIAutomation.ahk
	Run, UIAViewer.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\AHK-Projects.v1\UIAutomation
}
runtxtundrmouse(){
	; run GetTextUnderMouse.ahk
	Run, GetTextUnderMouse.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\AHK-Projects.v1\UIA
}
runshellhookmessages(){
	; run Messages_Shell.ahk
	Run, Messages_ShellHook.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\AHK-Projects.v1
}
runconstatine(){
	; run Constantine.ahk
	Run, Constantine.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\AHK-Projects.v1\AutoGUI\Tools
}
runHznbutton(){
	; run Hznbutton.ahk
	Run, Hznbutton.ahk, C:\Users\bacona\OneDrive - FM Global\AHK.MAIN\AHK-Projects.v1\lib
}

openFileLocation(){
    Run % A_ScriptDir
}

download(url, filename) {
    UrlDownloadToFile, %url%, %filename%
    if (ErrorLevel != 0) {
        trayNotify(script_full_name, "File '" . filename . "' couldn't be downloaded from " . url . ". Maybe the system is offline?")
    } else {
        trayNotify(script_full_name, "File '" . filename . "' downloaded.")
    }
}

viewKeyboardShortcuts(){
    hotkey_pdf_url := "https://github.com/arlbibek/windows-ahk/raw/master/hotkeys.pdf"
    hotkey_pdf := "hotkeys.pdf"
    hotkey_pdf_path := % A_ScriptDir . "\" . hotkey_pdf
    While, True {
        if not FileExist(hotkey_pdf_path){

            MsgBox, 4, File not found: would like to download?, The %hotkey_pdf% file doesn't exist. `nThis pdf file contains detailed the list of keyboard shortcuts for %script_full_name%. `n`nWould you like to download and open the file? `nURL: %hotkey_pdf_url%

            IfMsgBox, Yes
                download(hotkey_pdf_url, hotkey_pdf)
else
    Break
        } else {
            Run, % hotkey_pdf_path
            Break
        }
    }
}

addTrayMenuOption(label = "", command = "") {
    if (label = "" && command = "") {
        Menu, Tray, Add
    } else {
        Menu, Tray, Add, % label, % command
    }
}
;#[WINDOWS]
updateTrayMenu() {
	Menu, Tray, NoStandard
	addTrayMenuOption("Modified with Nerd by Adam Bacon and Terry Keatts.`nCredit:Made with ❤️ by Bibek Aryal", "madeBy")
	addTrayMenuOption()
	addTrayMenuOption("Run at startup", "runAtStartup")
	Menu, Tray, % fileExist(startup_shortcut) ? "check" : "unCheck", Run at startup ; update the tray menu status on startup
	addTrayMenuOption("Presentation mode {Win+Shift+P}", "togglePresentationMode")
	addTrayMenuOption("Keyboard shortcuts {Ctrl+Shift+Alt+\}", "viewKeyboardShortcuts")
	addTrayMenuOption("Open file location", "openFileLocation")
	addTrayMenuOption()
	addTrayMenuOption("View in GitHub", "viewInGitHub")
	addTrayMenuOption("See AutoHotKey documentation", "viewAHKDoc")
	addTrayMenuOption("See AHK GUI documentation", "viewGuiDoc")
	addTrayMenuOption("AHK Built in Functions (e.g., A_Function)", "viewAHKVariables")
	addTrayMenuOption()
	addTrayMenuOption("Run GetTextUnderMouse.ahk", "runtxtundrmouse")
	addTrayMenuOption("Run WinspectorU.exe", "runWinspector")
	addTrayMenuOption("Run UIAutomation.ahk", "runUIAutomation")
	addTrayMenuOption("Run Messages_ShellHook.ahk", "runshellhookmessages")
	addTrayMenuOption()
	addTrayMenuOption("Run Scriptlet_Library_v4.ahk", "runScriptlet")
	addTrayMenuOption("Run AHK Script.ahk", "runAHKScript")
	addTrayMenuOption("Run Hznbutton.ahk", "runHznbutton")
	addTrayMenuOption("Run Constantine.ahk", "runconstatine")
	addTrayMenuOption()
	Menu, Tray, Standard
}

updateTrayMenu()
;trayNotify(script_full_name . " started", "Open keyboard shortcuts with {Ctrl + Shift + Alt + \}`n`nModified with Nerd by Adam Bacon and Terry Keatts`nMade with ❤️ by Bibek Aryal."),.5
/*
;#[HOTKEYS]

; == HOTKEYS ==
; Remapping function keys

; F1 to Firefox
;F1::activate(get_default_browser(), "cycle")
;+F1::Run % get_default_browser()

; F2 is Rename
; F2::

; F3 to Spotify
;F3::activate("spotify.exe")

; F4 to VS Code
;F4::activate("code.exe", "cycle")

; F5 is Refresh
; F5::

; F6 is SumatraPDF
;F6::activate("SumatraPDF.exe")

; F7 to Microsoft
;F7::activate("winword.exe")

; F8 to Microsoft Excel
;F8::activate("excel.exe")

; F9 is
; F9::

; F10 is WindowsTerminal
;F10::activate("WindowsTerminal.exe", "minimize", "-d " . userdir) ; Launching windows terminal on current users directory
;F10::Run % "WindowsTerminal.exe" ;-d " . userdir

; F11 is Full Screen
; F11::

; F12 is
; F12::
*/
; Windows hotkeys

; file explorer
#e::
IfWinNotExist, ahk_class CabinetWClass
	Run, explorer.exe
GroupAdd, ActiveExplorers, ahk_class CabinetWClass
if WinActive("ahk_exe explorer.exe"){
        GroupActivate, ActiveExplorers, r
    }
    else
        WinActivate ahk_class CabinetWClass
return

; notepad
;#n::activate("notepad++.exe", "cycle")
;+#n::Run, notepad.exe

; search selected text/clipboard on the web
#s::win_search(get_selected())
+#s::win_search(Clipboard)

;  center window
#c::
    WinGetTitle, ActiveWindowTitle, A
    WinGetPos,,, Width, Height, %ActiveWindowTitle%
    TargetX := (A_ScreenWidth / 2) - (Width / 2)
    TargetY := (A_ScreenHeight / 2) - (Height / 2)
    WinMove, %ActiveWindowTitle%,, %TargetX%, %TargetY%
Return

; Other Hotkeys

^!c::
    Send, ^c
    ClipWait, 3
    Clipboard := StrReplace(Clipboard, "`r`n", " ")
    Clipboard := StrReplace(Clipboard, "- ", "")
Return

; Toggle presentation mode
^+`::activate("SyncTrayzor.exe")

; Toggle presentation mode
#+p::togglePresentationMode()

; Change the case of selected text(s)
CapsLock & 7::changeCaseTo("lower")
CapsLock & 8::changeCaseTo("titled")
CapsLock & 9::changeCaseTo("upper")

+Space::
    str := get_selected()
    str := StrReplace(str, " ", "_")
    SendRaw % str
    str_len := StrLen(str)
    Send, +{left %str_len%}
Return

$Escape::
    ; Long press (> 0.5 sec) on Esc closes window - but if you change your mind you can keep it pressed for 3 more seconds
    KeyWait, Escape, T0.5 ; Wait no more than 0.5 sec for key release (also suppress auto-repeat)
    If ErrorLevel ; timeout, so key is still down...
    {
        SoundPlay *64 ; Play an asterisk
        WinGet, X, ProcessName, A
        SplashTextOn,,150,,`nRelease button to close %x%`n`nKeep pressing it to NOT close window...
        KeyWait, Escape, T3 ; Wait no more than 3 more sec for key to be released
        SplashTextOff
        If !ErrorLevel ; No timeout, so key was released
        {
            PostMessage, 0x112, 0xF060,,, A ; ...so close window
            Return
        }
        ; Otherwise,
        SoundPlay *64
        KeyWait, Escape ; Wait for button to be released
        ; Then do nothing...
        Return
    }
    Send {Esc}
; REFERENCED FROM: https://www.autohotkey.com/board/topic/80697-long-keypress-hotkeys-wo-modifiers/
Return

^+!\::viewKeyboardShortcuts()
^+!s::
    Suspend
    if (A_IsSuspended = 1){
        trayNotify(script_full_name . " suspended", "All hotkeys will be suspended (paused). `n`nPress {Ctrl + Shift + Alt + S} or use the tray menu option to toggle back.")
    } else {
        trayNotify(script_full_name . " restored", "All hotkeys resumed (will work as intended). `n`nPress {Ctrl + Shift + Alt + S} to suspend.")
    }
Return

Return

; Hotkeys within file explorers

; navigating within the file explorer
#IfWinActive ahk_group ExplorerGroup
    ^+u::exploreTo(userdir)
    ^+e::exploreTo(pc)
    ^+h::exploreTo(desktop)
    ^+d::exploreTo(documents)
    ^+j::exploreTo(downloads)
    ^+m::exploreTo(music)
    ^+p::exploreTo(pictures)
    ^+v::exploreTo(videos)
    ^+a::exploreTo(arlbibek)
    ^+s::exploreTo(screenshot)
#IfWinActive

; opening programs via file explorer
#IfWinActive ahk_class CabinetWClass
    ^+t::exploreTo("wt") ; windows terminal
    ^+\::exploreTo("code .") ; VS code in current directory
#IfWinActive
/*
; Hotkeys within google docs and sheets

;^Insert::docSheetWr("^!m") ; open comment box

; Move the current sheet/docs to trash
;^Delete::
    ;docSheetWr("!/Move to trash")
    ;Sleep, 200
    ;docSheetWr("{Enter}")
;return

; Highlight selected text/shell
;!1::
    ;docSheetWr("!/Highlight color: yellow")
    ;Sleep, 200
    ;docSheetWr("{Enter}")
;return

; Remove Highlight for selected text/shell(s)
;!+1::
    ;docSheetWr("!/Highlight color: none")
    ;Sleep, 200
    ;docSheetWr("{Enter}")
;return

; Wrap selected shell (Sheets only)
!2::
    sheetWr("!/Wrap text{Down 2}{Enter}")
    Sleep, 200
    sheetWr("{Enter}")
return

; Trim whitespace (Sheets only)
!3::
    sheetWr("!/Trim whitespace")
    Sleep, 200
    docSheetWr("{Enter}")
return

; Spell check
!4::
    docSheetWr("!/Spell check")
    Sleep, 200
    docSheetWr("{Enter}")
return

;  Strike selected text
!8::docSheetWr("!+5")
*/
; ; == HOT STRINGS ==

; ; Current date and time
FormatDateTime(format, datetime="") {
    if (datetime = "") {
        datetime := A_Now
    }
    FormatTime, CurrentDateTime, %datetime%, %format%
    SendInput, %CurrentDateTime%
    return
}

; Hotstrings
::/datetime::
    FormatDateTime("dddd, MMMM dd, yyyy, HH:mm")
Return

::/datetimett::
    FormatDateTime("dddd, MMMM dd, yyyy hh:mm tt")
Return
:X:/commit::FormatDateTime("yyyy.MM.dd HH:mm")
::/time::
    FormatDateTime("HH:mm")
Return
::/timett::
    FormatDateTime("hh:mm tt")
Return
::/date::
    FormatDateTime("MMMM dd, yyyy")
Return
::/daten::
    FormatDateTime("MM/dd/yyyy")
Return
::/datet::
    FormatDateTime("yy.MM.dd")
Return
::/week::
    FormatDateTime("dddd")
Return
::/day::
    FormatDateTime("dd")
Return
::/month::
    FormatDateTime("MMMM")
Return
::/monthn::
    FormatDateTime("MM")
Return
::/year::
    FormatDateTime("yyyy")
Return

; Others
::wtf::Wow that's fantastic
::/paste::
    Send % Clipboard
Return
::/cud::
    ; useful for WSLs
    SendInput, /mnt/c/Users/%A_UserName%/
Return
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

; Made with ❤️ by Bibek Aryal.