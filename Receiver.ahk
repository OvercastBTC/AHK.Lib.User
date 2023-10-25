#Requires AutoHotkey v2.0
#Include <Directives\__AE.v2>

OnMessage(0x004A, Receive_WM_COPYDATA)  ; 0x004A is WM_COPYDATA
Persistent(1)

Receive_WM_COPYDATA(wParam, lParam?, msg?, hwnd?)
{
    wParam_String := NumGet(wParam, 2*A_PtrSize, "Ptr")  ; Retrieves the CopyDataStruct's lpData member.
    ; lParam_String := NumGet(lParam, 2*A_PtrSize, "Ptr")  ; Retrieves the CopyDataStruct's lpData member.
    wParam_CopyOfData := StrGet(wParam_String)  ; Copy the string out of the structure.
    ; lParam_CopyOfData := StrGet(lParam_String)  ; Copy the string out of the structure.
    ; wParam_CopyOfData := wParam  ; Copy the string out of the structure.
    ; lParam_CopyOfData := lParam  ; Copy the string out of the structure.
    ; Show it with ToolTip vs. MsgBox so we can return in a timely fashion:
    ToolTip( A_ScriptName "`nReceived the following string:`n" 
            . 'wParam_CopyOfData: ' . wParam_CopyOfData . '`n' 
            ; . 'lParam_CopyOfData: ' . lParam_CopyOfData . '`n'
        )
    return true  ; Returning 1 (true) is the traditional way to acknowledge this message.
}