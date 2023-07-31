﻿/*
    Recupera el identificador del thread y el identificador del proceso que creó la ventana especificada.
    Parámetros:
        hWnd: El identificador de la ventana.
    Return:
            0 = Ha ocurrido un error.
        [obj] = Si tuvo éxito devuelve un objeto con las siguientes claves:
            ThreadID  = El identificador del thread que creó la ventana especificada.
            ProcessID = El identificador del proceso que creó la ventana especificada.
    Ejemplo:
        MsgBox(WinGetProcessName('ahk_pid' . GetWindowThreadProcessId(WinExist('A')).ProcessID))
*/
#Requires AutoHotkey v2
#Include <Class_Toolbar.c2v2>
Class GetWindowThreadProcessId extends Horizon {
    GetWindowThreadProcessId(hWnd) {
        Local ProcessID
        ; , ThreadID  := DllCall('User32.dll\GetWindowThreadProcessId', 'Ptr'  , hWnd      ;hWnd
                                                                    ; , 'UIntP', &ProcessID ;lpdwProcessId
                                                                    ; , 'UInt')            ;ReturnType    --> dwThreadId
        return    ProcessID := DllCall('User32.dll\GetWindowThreadProcessId', 'Ptr', hWnd, 'Ptr', 0, 'UInt')
        ; Return (ThreadID ? {ThreadID: ThreadID, ProcessID: ProcessID} : FALSE)
    } ;https://msdn.microsoft.com/en-us/library/windows/desktop/ms633522(v=vs.85).aspx
}
