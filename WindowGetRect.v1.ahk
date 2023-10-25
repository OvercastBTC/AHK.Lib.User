
rect := WindowGetRect("window title etc.")
MsgBox % rect.width "`n" rect.height

WindowGetRect(windowTitle*) {
    if hwnd := WinExist(windowTitle*) {
        VarSetCapacity(rect, 16, 0)
        DllCall("GetClientRect", "Ptr", hwnd, "Ptr", &rect)
        return {width: NumGet(rect, 8, "Int"), height: NumGet(rect, 12, "Int")}
    }
}