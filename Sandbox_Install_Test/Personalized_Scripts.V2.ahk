;---------------------------------------------------------------------------
;                           Autoexecute Section
;---------------------------------------------------------------------------
#Include <AE.V2>
#Requires AutoHotkey v2

;---------------------------------------------------------------------------
;                         Personal HotStrings
;---------------------------------------------------------------------------

/*  KEY SYMBOLS
	# = Windows
	! = Alt
	^ = Control
	+ = Shift
	& = An ampersand may be used between any two keys or mouse buttons to combine them into a custom hotkey.
*/

; For full Hotkey instructions see https://www.autohotkey.com/docs/v2/Hotkeys.htm
; For Hotstring options refer to https://www.autohotkey.com/docs/v2/Hotstrings.htm#Options


;---------------------------------------------------------------------------
;                  Personalized Hotstrings and Includes 
;---------------------------------------------------------------------------





;---------------------------------------------------------------------------
;                 			 Personalized Scripts 
;---------------------------------------------------------------------------




;-------------------------------------------------------------------------
;              		      Quick Launch Websites
;--------------------------------------------------------------------------

^+#f:: ; use ctrl+Shift+Win+f
{
Run("`"https://portal.fcm.travel/Account/Login`"")
return
}

^+#a:: ; use ctrl+Shift+Win+a
{
Run("`"https://www.approvalguide.com//`"")
return
}

^+#j:: ; use ctrl+Shift+Win+j
{
Run("`"https://www.jurisdictiononline.com/`"")
return
}

^+#r::Run("`"http://riskview/RiskView_1_0/Home/MainPage`"")

^+#e::Run("`"https://engnet/Engnet/engnet/engnet.asp`"")


;--------------------------------------------------------------------------
;     		    Win-W to Pop Open Common/Files Folders GUI
;--------------------------------------------------------------------------
MButton::
#w::{

	myGui := Gui(,"Quick Open Folder/File/App",)
	myGui.Opt("AlwaysOnTop")
	myGui.SetFont("s12")
	myGui.Add("Button","x15 +default", "Desktop").OnEvent("Click", ClickedDesktop)
	myGui.Add("Button","x15", "One Drive / My Docs").OnEvent("Click", ClickedOneDrive)
	myGui.Add("Button","x15", "Downloads").OnEvent("Click", ClickedDownloads)
	; myGui.Add("Button","x15", "B+M LE Guide").OnEvent("Click", ClickedBMLEG)
	myGui.Add("Button","x15", "Cancel").OnEvent("Click", ClickedCancel)
	myGui.Show("w350")

	ClickedDesktop(*)
	{
		myGui.Destroy()
		TakeAction("C:\Users\" A_UserName "\OneDrive - FM Global\Desktop")
	}

	ClickedOneDrive(*)
	{
		myGui.Destroy()
		TakeAction("C:\Users\" A_UserName "\OneDrive - FM Global\")
	}

	ClickedDownloads(*)
	{
		myGui.Destroy()
		TakeAction("C:\Users\" A_UserName "\OneDrive - FM Global\Downloads")
	}
	
	; ClickedBMLEG(*)
	; {
	; 	myGui.Destroy()
	; 	Run("C:\Users\" A_UserName "\FM Global\Operating Standards - Documents\general\B+M Loss Expectancy Guide - July 2023.xlsx")
	; }
		
	ClickedCancel(*)
	{
		myGui.Destroy()
	}

	TakeAction(Filepath)
	{
		IF WinActive("Save")	or WinActive("Open")
		{
			Send(Filepath "{Enter}")
		}

		Else
		{
			Run(Filepath)
		}
	}

}