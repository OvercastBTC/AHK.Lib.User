;---------------------------------------------------------------------------
;                           Autoexecute Section
;---------------------------------------------------------------------------
#Include AE.V2.ahk
#Requires AutoHotkey v2

;---------------------------------------------------------------------------
;                              CHANGELOG
;---------------------------------------------------------------------------

; Version 6, Released 11/3/2023
; Moved FMDS Titles into separate file for easier management.

; Version 5.0, Released 9/17/2023
; Converted script to V2. Moved Autocorrect stuff from this to the Autocorrect script
; as the Autocorrect script now detects duplicate items with the new Win+H hotkey script.
; This will prevent users from accidentally adding duplicate items. Also, moved OS/DS
; quick launch script from the main script to this script to make management for the 
; stand alone app easier as it will be able to share this script without modification.

; Version 4.0, Released 7/28/2023
; Reworked script based on new Horizon italics functionality so that standard control keys
; will work. Some items have been moved to Quick Access Pop-up. Updated name for OS 3-36.

; Version 3.4 - Released 5/30/2023
; Changed "sgsv" to "sgsvf" based on user feedback. Added "degf" for addeding a degree symbol.

; Version 3.3 - Released 5/3/2023
; Added case specific hot string areas to reduce rework on strange cases.

; Version 3.2 - Released 4/27/2023
; Updated OS 7-11 title.

; Version 3.1 - Released 4/11/2023
; Updated OS 10-1 title. Improved fast launch of OS GUI.

; Version 3 - 7 Mar 2023
; OS Titles now input in italics except for Horizon.

; Version 2.0 - Released 3/6/2023
; Added the Shift+Ctrl+WIN+O hot key for fast launch of a specific OS.;

; Version 1.0 - Released 2/24/2023
; Modified to correct a few standard titles and add the "densityf" hotstring with input for the varibles.
; Also modified to break this script out into a sub-script to make maintenance easier for users, so they
; are not forced to update the starter script with all the personal information, which likely won't change
; as much as the sub-scripts.

;@Ahk2Exe-IgnoreBegin
;---------------------------------------------------------------------------
;                         User Abbreviations 
;---------------------------------------------------------------------------

:x*:eml::Send(work_email)	    	; Email

:x*:peml::Send(personal_email)		; Personal Email

:x*:emp#::Send(employee_number)		; Employee Number

:x*:addrf::Send(office_address) 	; Office Address

:x*:paddrf::Send(personal_address)	; Personal Address

:x*:planeml::Send(manager_name ", " manager_title ", " manager_email "`n" your_name ", FM Global " your_title ", " work_email A_Space ) ; plan review email distribution addition

:x*:lettersig::Send(your_name "`n" your_title "`n" your_office "`n" work_email "`n" office_phone "`n")

:x*:usr::Send(A_UserName)	; workstation username

:*:plansig::
{
	myGui := Gui(,"Plan Review Office",)
	myGui.Opt("AlwaysOnTop")
	myGui.SetFont("s12")
	myGui.Add("Button","x15 +default", "San Francisco").OnEvent("Click", ClickedSF)
	myGui.Add("Button","x15", "Los Angeles").OnEvent("Click", ClickedLA)
	myGui.Add("Button","x15", "Dallas").OnEvent("Click", ClickedTX)
	myGui.Add("Button","x15", "Cancel").OnEvent("Click", ClickedCancel)
	myGui.Show("w350")

	ClickedSF(*)
	{
		myGui.Destroy()
		Send(your_name "`n" your_title "`nFM Global`nENGSanFranciscoPlanReviewSM@fmglobal.com`n(925) 287-4336`n") ; Plan Review Signature for SF Office
	}

	ClickedLA(*)
	{
		myGui.Destroy()
		Send(your_name "`n" your_title "`nFM Global`nENGLosAngelesPlanReview@fmglobal.com`n(818) 227-2263`n") ; Plan Review Signature for LA Office
	}

	ClickedTX(*)
	{
		myGui.Destroy()
		Send(your_name "`n" your_title "`nFM Global`nENGDallasPlanReview@fmglobal.com`n(972) 377-4808`n") ; Plan Review Signature for Dallas Office
	}
	
	ClickedCancel(*)
	{
		myGui.Destroy()
	}

	return
}

;@Ahk2Exe-IgnoreEnd

;---------------------------------------------------------------------------
;                          Common Abbreviations
;---------------------------------------------------------------------------
:?*X:microm::Send(chr(181) "m")
:?*X:3/4f::Send(chr(190))
:?*X:1/2f::Send(chr(189))
:?*X:1/4f::Send(chr(188))
:?*X:+-::Send(chr(177))
:?*X:circler::Send(chr(174)) ;®
:?*X:regtmf::Send(chr(174)) ;®
:?*:trademarkf::™
:?*:circlec::©
:?*:copywritef::©
:?*:greater than or equal to::≥
:?*:>=::≥ ; greater than or equal to
:?*X:degf::Send(chr(176) "F")
:?*X:degc::Send(chr(176) "C")
:?*X:prisecf::Send("1" chr(176) "/2" chr(176) A_Space "Inj Testing")
:?*:prisecf::1°/2° Inj Testing
:*:hrsgf::heat recovery steam generator (HRSG)
:*:mocf::management of change (MOC)
:*C:agf::Approval Guide
:*C:fmdsf::FM Global Property Loss Prevention Data Sheet
:*C:FMDSf::FM Global Property Loss Prevention Data Sheet
:*C:2oo3::two out of three-voting logic
:*C:fmrtps::FM Global Red Tag Permit System
:*C:fmhwps::FM Global Hot Work Permit System
:*C:fmdsl::FM Global Data Sheet
:*C:fma::FM Approved
:*C:FMA::FM Approved
:*C:fmilsc::FM Approved Ignitable Liquid Storage Cabinet
:*:sgsvf::seismic gas shutoff valve
:*:erpf::emergency response plan
:*:ferpf::flood emergency response plan
::wst::water supply tool
:*C:wstf::water supply tool
::efc::eFC
::wdt::water delivery time
:*:wpivf::wall post-indicator valve
:*:pivf::post-indicator valve
:*:ulf::Underwriters Laboratories
:*:uupf::uncartoned unexpanded plastic
:*:uepf::uncartoned expanded plastic
:*:cupf::cartoned unexpanded plastic
:*:cepf::cartoned expanded plastic
:*:sopf::standard operating procedure
:*:eopf::emergency operating procedure
::ooo::out of office
::OSY::OS&Y
:*:oemf::original equipment manufacturer
:*:ndef::nondestructive examination
:*:mehpf::minimum end head pressure
:*:mawpf::maximum allowable working pressure
:*:mipf::metal insulating panels
:*:impf::insulated metal panels
::lwco::low-water cutoff
::lwfco::low-water fuel cutoff
::lfpil::low flash point ignitable liquid
::hfpil::high flash point ignitable liquid
:*:lmgtfy::https://letmegooglethat.com/?q=
:*:itmf::inspection, testing and maintenance
:*:irasf::in-rack automatic sprinklers
:*:ilf::ignitable liquid
:*:hrlf::higher RelativeLikelihood
:*:htff::heat transfer fluid
::gp::generally protected
:*:frpf::fiber-reinforced plastic panels
:*:epof::emergency power off
:*:blrbf::black liquor recovery boiler
:C:efile::eFile
:*:icsf::industrial control system
:*:otf::operational technology network
:*:itf::information technology network
:*:mfaf::multifactor authentication
:*:bmf::boiler and machinery
:*:fnhf::fire and natural hazards
::ostt::overspeed trip test
::the majority of::most of
::in the near future::soon
::as a result of::because of
::recs::recommendations
:*:cbtf::computer-based training
:*:--f::--------------------------------------------------------------------------------------------------------------------`n{End}`n--------------------------------------------------------------------------------------------------------------------
:*C:ahkf::AutoHotkey
:*C:AHKf::AutoHotKey
::rsw::right sidewall
::lsw::left sidewall
::tpd::tons per day
::fpm::ft./min.
:*:bpvf::boiler and pressure vessel
:*:lmgtfy::https://letmegooglethat.com/?q=

:*:densityf::								; input density and area in statement
{
    IB := InputBox("Density in gpm/sq. ft.", "Sprinkler System Density", "w300 h125"), density := IB.Value
    IB := InputBox("Area in sq. ft.", "Demand Area", "w300 h125"), area := IB.Value
    Send(density " gpm/ft" chr(178) " over " area " ft" chr(178))
    return
}

;---------------------------------------------------------------------------
;                           NFPA Standards
;---------------------------------------------------------------------------

:*:nfpa 13f::
{ 
Send("NFPA 13, ")
Send("^i")
Sleep(100)
Send("Installation of Sprinkler Systems")
Sleep(100)
Send("^i")
return
}


:*:nfpa 25f::
{ 
Send("NFPA 25, ")
Send("^i")
Sleep(100)
Send("Standard for the Inspection, Testing, and Maintenance of Water-Based Fire Protection Systems")
Sleep(100)
Send("^i")
return
}