
#requires AutoHotkey v1.1+
RichEdit_GetTextLength(hCtrl, Flags=0, CodePage="")  {
    static EM_GETTEXTLENGTHEX=95,WM_USER=0x400
    static GTL_DEFAULT=0,GTL_USECRLF=1,GTL_PRECISE=2,GTL_CLOSE=4,GTL_NUMCHARS=8,GTL_NUMBYTES=16
  
    hexFlags:=0
      Loop, parse, Flags, %A_Tab%%A_Space%
          hexFlags |= GTL_%A_LOOPFIELD%
  
    VarSetCapacity(GETTEXTLENGTHEX, 4)
    NumPut(hexFlags, GETTEXTLENGTHEX, 0), NumPut((codepage="unicode"||codepage="u") ? 1200 : 1252, GETTEXTLENGTHEX, 4)
    SendMessage, EM_GETTEXTLENGTHEX | WM_USER, &GETTEXTLENGTHEX,0,, ahk_id %hCtrl%
    IfEqual, ERRORLEVEL,0x80070057, return "", errorlevel := A_ThisFunc "> Invalid combination of parameters."
    IfEqual, ERRORLEVEL,FAIL      , return "", errorlevel := A_ThisFunc "> Invalid control handle."
    return ERRORLEVEL
  }