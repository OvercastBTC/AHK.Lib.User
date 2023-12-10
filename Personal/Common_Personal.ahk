﻿#Requires AutoHotkey v2
#Include <Directives\__AE.v2>
; --------------------------------------------------------------------------------

; --------------------------------------------------------------------------------
/**
 * @function Common_Personal_Abbreviations
 */
; --------------------------------------------------------------------------------

:*:peml::adam.bacon80@icloud.com	; Personal Email
:*:eml::adam.bacon@fmglobal.com		; Email
:*:emp#::52906						; Employee Number
:*:costcenter::908					; Employee Number
:*:officeaddys::					; Office Address (single line)
{
	Send('FM Global: 601 108th Ave NE')
	Send(', Suite 1400, Bellevue, WA 98004')
}	
:*:officeaddyml::					; Office Address (multi line)
{
	A_Clipboard := "
	(
		FM Global: 
		601 108th Ave NE 
		Suite 1400 
		Bellevue, WA 98004
	)"
	Send(A_Clipboard)
}	

:*:paddrf::1378 Fenmore Ave, Sanger, CA 93657	; My Address
; :*:pwnotes::
; =============================== + FM Stuff + =============================
:*:fmusr::bacona				; workstation username
:*:fmpw::**{!}980HDKBab			; fm 
; --------------------------------------------------------------------------------
#HotIf WinActive("EngNET")
	:*:usr::bacona
	; :*:pw::80ab**HD{!}9KB{!}2{!}9
	:*:pw::**{!}980HDKBab
#HotIf
#HotIf WinActive("Xfinity")
	:*:usr::harmony_bacon@comcast.net
	:*:pw::Ah{!}9801988
#HotIf
#HotIf WinActive("AT&T")
	:*:usr::bacon942882
	:*:pw::Ah_19801988
#HotIf
#HotIf WinActive("marriott") || WinActive('Marriott')
	:*:usr::063194266
	:*:pw::HDBacon{!}9801988
#HotIf
#HotIf WinActive("GitHub")
	:*:usr::OvercastBTC
	:*:pw::HDBacon{!}980{!}988
#HotIf
#HotIf WinActive("AutoHotkey")
	:*:usr::OvercastBTC
	:*:pw::HDBacon{!}980{!}988
#HotIf
#HotIf WinActive("ticketmaster.com") || WinActive('Seattle Mariners Account Manager')
	:*:usr::adam.bacon80@icloud.com
	:*:pw::HDBacon{!}980{!}988
#HotIf
;============================== Other =============================================
:?*:ppw::HDBacon{!}980{!}988								; other password
:?*:epw::HDBacon{!}9801988									; other password
; :?*:fpw::80**19HDabKB										; g password
:?*:fpw::{!}980ab**HD{!}9KB									; g password
:?*:wifionboard::Ah{!}980198812
; --------------------------------------------------------------------------------
:*:solusr::abacon01
:*:solpw::7jZuhYwi
:*:solvznpw::gill4-frank-dug
; - myQuest ----------------------------------------------------------------------
#HotIf WinActive('Login - CAS – Central Authentication Service - Google Chrome')
:*:usr::adam.bacon8088@gmail.com
:*:pw:: Ah{!}980198819
#HotIf
; - Reddit -----------------------------------------------------------------------
#HotIf WinActive("Reddit")
:*:usr::Comfortable_Kick_776
:*:pw:: 80ab**HD!9KB20!9
#HotIf
; --------------------------------------------------------------------------------
#HotIf WinActive('Code Project')
:*:usr::adam.bacon80@icloud.com
:*:pw:: HDBacon20{1}5
#HotIf
; --------------------------------------------------------------------------------
#HotIf WinActive('Carey')
:*:usr::adam.bacon@fmglobal.com
:*:pw:: Ah{!}9801988
#HotIf
; - fmjobapp ---------------------
:X?:joblog::Run("https://intranet-fmglobal.icims.com/jobs/candidate?back=dashboard")
#HotIf WinActive("intranet-fmglobal.icims.com")
:*:eml::adam.bacon@fmglobal.com
:*:usr::adambacon160807
:*:pw::80ab**HD{!}9KB{!}9
#HotIf
; - Marsh ---------------------
:X?:marshlog::Run("https://us-mmc-datatransfer.mmc.com/login?#/")
#HotIf WinActive("kiteworks")
:*:usr::adam.bacon@fmglobal.com
:*:pw::HDBacon{!}9801988
#HotIf
; - Delta---------------------
:*:deltalog::Run("https://www.delta.com/mydelta/walletMyReceiptSummary")

HotIf(WinActive("Delta"))
; :*:usr::9339385271
:X*:usr::Send(log.delta.login['u'])
; :*:pw::HDBacon{!}9801988
:X*:pw::Send(log.delta.login['p'])
HotIf()
; - Alaska--------------------
:X?:aklog::Run('https://www.alaskaair.com')

aTitle := 'Alaska Airlines'
; HotIf(WinActive(aTitle))
HotIf(WinActive('Alaska Airlines'))
; #HotIf WinActive("Alaska Airlines")
; :*:usr::72832631
:X*:usr::Send(log.alaksa.login['u'])
:X*:pw::Send(log.alaksa.login['p'])
; #HotIf
HotIf()

HotIf(WinActive('Alaska'))
:*X:pw::log.alaksa.login['u']
class log {
	class alaksa {
		static usr := 72832631
		static pw  := 'Ah{!}9801988'
		static login := Map(
			'u', this.usr,
			'p',this.pw,
		)
	}
	Class delta {
		static usr := '9339385271'
		static pw := 'Ah{!}9801988'
		static login := Map(
			'u', this.usr,
			'p', this.pw,
		)
	}
}
; --------------------------------------------------------------------------------
/** @airline United */
:X?:unitedlog::Run("https://www.united.com/")
#HotIf WinActive("United")
:*:usr::MP216222
:*:pw::HDBacon{!}9801988
#HotIf
; - Hilton---------------------------------------
:X?:hiltonlog::Run("https://www.hilton.com/")
#HotIf WinActive("Hilton")
:*:usr::942499799
:*:pw::HDBacon{!}9801988
#HotIf
; - the-automator.com----------------------------
:X?:automatorlog::Run("https://https://www.the-automator.com/wp-login.php")
#HotIf WinActive("the-Automator")
:*:usr::adam.bacon80@icloud.com
:*:pw::HDBacon{!}9801988
#HotIf
; - iCloud ---------------------------------------
:X?:icloudlog::Run("https://www.icloud.com/")
#HotIf WinActive("iCloud") || WinActive('Apple Music')
:*:usr::adamjcdbacon@hotmail.com
:*:pw::80ab**HD{!}9KB20{!}9
#HotIf
; - Discord ---------------------------------------
:X?:discordlog::Run("https://discord.com/channels/@me")
#HotIf WinActive("Discord")
:*:usr::adam.bacon8073@gmail.com
:*:pw::HDBacon{!}980{!}988
#HotIf
; - Avis ---------------------------------------
:X?:avislog::Run("https://www.avis.com/")
#HotIf WinActive("Avis")
:*:usr::S4K347
:*:pw::HDBacon{!}9801988
#HotIf
/*
MileIQ: 80ab**HD19KB
Marriott: 063194266 , HDBacon!9801988
ATT: Ah_19801988
UID: 1170ON
(VAHealth)idme: Ah!980198819
FCM: 80ab**HD19KB1
Xf: Ah!....
Hilton: 942499799 , HDBacon!9801988
United: MP216222 , HDBacon!9801988
Alaska: 72832631 , Ah!9801988
Delta: 9339385271, HDBacon!9801988
Hertz: 21082776 , HD
Carey: fm email, HD
Expense Report Format:

Business Travel - Airfare - 
Business Travel - Meal - 
Business Travel - Rental Car - 

FM Approvals: 80ab**HD19KB
*/