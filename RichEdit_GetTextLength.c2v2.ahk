
#Requires Autohotkey v2.0
GetTextLength(hCtrl, Flags:=0, CodePage:="")
{
    static EM_GETTEXTLENGTHEX:=95,WM_USER:=0x400
    static GTL_DEFAULT:=0,GTL_USECRLF:=1,GTL_PRECISE:=2,GTL_CLOSE:=4,GTL_NUMCHARS:=8,GTL_NUMBYTES:=16

    hexFlags:=0
    Loop Parse, Flags, A_Tab "" A_Space
        hexFlags |= GTL_%A_LOOPFIELD%

    VarSetStrCapacity(&GETTEXTLENGTHEX, 4) ; V1toV2: if 'GETTEXTLENGTHEX' is NOT a UTF-16 string, use 'GETTEXTLENGTHEX := Buffer(4)'
    NumPut("UPtr", hexFlags, GETTEXTLENGTHEX, 0), NumPut("UPtr", (codepage="unicode"||codepage="u") ? 1200 : 1252, GETTEXTLENGTHEX, 4)
    ErrorLevel := SendMessage(EM_GETTEXTLENGTHEX | WM_USER, &GETTEXTLENGTHEX, 0, , hCtrl)
    ; if (ERRORLEVEL = 2147942487)
    ;     return "", errorlevel := A_ThisFunc "> Invalid combination of parameters."
    ; if (ERRORLEVEL = "FAIL")
    ;     return "", errorlevel := A_ThisFunc "> Invalid control handle."
    return ERRORLEVEL
}