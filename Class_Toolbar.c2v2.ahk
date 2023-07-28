

Class HznToolbar
{
    
    ; Static TB := "TB_"
    ; Static BUTTONCOUNT := "BUTTONCOUNT"
    ; Static GETBUTTON := "GETBUTTON"
    ; Static GETITEMRECT := "GETITEMRECT"
    ; Static MEM := "MEM_"
    ; static COMMIT := "COMMIT"
    ; static RESERVE := "RESERVE"
    ; static PROTECT := "PROTECT"
    ; static RELEASE := "RELEASE"
    __New()
    {
        SendLevel(5)
        static fCtl := ControlGetClassNN(ControlGetFocus("A"))
        static tbID := SubStr(fCtl, -1, 1)
        Global tCtl := "msvb_lib_toolbar" tbID
    }

    ; hToolbar(&tCtl)
    ; {
    ;     Return cHwnd := ControlGetHwnd(&tCtl, "A")
    ; }

    BtnCount(Msg)
    {
        Msg := TB_BUTTONCOUNT := 1048
        wParam := 0
        lParam := 0
        Ctl := ""
        WinTitle := cHwnd := ControlGetHwnd(tCtl, "A")
        Return SendMessage(Msg, wParam, lParam, Ctl, "ahk_id " WinTitle)
    }

    static Msg := Map(

        "TB_BUTTONCOUNT", 1048,
        "TB_GETBUTTON", 1047,
        "TB_GETITEMRECT", 1053,
        "MEM_COMMIT", 4096
        "MEM_RESERVE", 8192
        "MEM_PHYSICAL", 4,
        "MEM_PROTECT", 0x40,  
        "MEM_RELEASE", 0x8000,
        "dwSize", 32

    )
}