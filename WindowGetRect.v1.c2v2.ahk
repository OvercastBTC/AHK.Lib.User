
rect := WindowGetRect("window title etc.")
MsgBox(rect.width "`n" rect.height)

WindowGetRect(windowTitle*) {
    if hwnd := WinExist(windowTitle*) {
        rect := Buffer(16, 0) ; V1toV2: if 'rect' is a UTF-16 string, use 'VarSetStrCapacity(&rect, 16)'
        DllCall("GetClientRect", "Ptr", hwnd, "Ptr", rect)
        return {width: NumGet(rect, 8, "Int"), height: NumGet(rect, 12, "Int")}
    }
}