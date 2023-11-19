;No dependencies

DarkMode(guiObj) {
	guiObj.BackColor := "171717"
	return guiObj
}
Gui.Prototype.DefineProp("DarkMode", {Call: DarkMode})

MakeFontNicer(guiObj, fontSize := 20) {
	guiObj.SetFont("s" fontSize " cC5C5C5", "Consolas") ;? cC5C5C5 = gray,gray,silver
	return guiObj
}
Gui.Prototype.DefineProp("MakeFontNicer", {Call: MakeFontNicer})

PressTitleBar(guiObj) {
	PostMessage(0xA1, 2,,, guiObj) ;? WM_NCLBUTTONDOWN
	return guiObj
}
Gui.Prototype.DefineProp("PressTitleBar", {Call: PressTitleBar})

NeverFocusWindow(guiObj) {
	WinSetExStyle("0x08000000L", guiObj) ;? WS_EX_NOACTIVATE
	WinSetExStyle('0x00000020L', guiObj) ;? WS_EX_TRANSPARENT
	WinSetExStyle('0x02000000L', guiObj) ;? WS_EX_COMPOSITED ; 
	WinSetExStyle('0x00000200L', guiObj) ;? WS_EX_CLIENTEDGE ; The window has a border with a sunken edge.
	WinSetExStyle('0x00040000L', guiObj) ;? WS_EX_APPWINDOW ; Forces a top-level window onto the taskbar when the window is visible.
	
	return guiObj
}
Gui.Prototype.DefineProp("NeverFocusWindow", {Call: NeverFocusWindow})

MakeClickthrough(guiObj) {
	WinSetTransparent(255, guiObj)
	guiObj.Opt("+E0x20")
	return guiObj
}
Gui.Prototype.DefineProp("MakeClickthrough", {Call: MakeClickthrough})