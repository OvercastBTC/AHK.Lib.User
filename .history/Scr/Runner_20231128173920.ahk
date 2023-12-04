; --------------------------------------------------------------------------------
#Requires AutoHotkey v2+
#Include <Directives\__AE.v2>
; --------------------------------------------------------------------------------
#Include <Environment>
#Include <Links>
#Include <App\Spotify>
#Include <App\Davinci>
#Include <Extensions\String>
#Include <Utils\ClipSend>
#Include <Extensions\String>
#Include <Utils\Win>
#Include <Paths>
#Include <Utils\Unicode>
#Include <Abstractions\Script>
#Include <Abstractions\Text>
#Include <Converters\DateTime>
#Include <Tools\CleanInputBox>
#Include <Misc\Meditate>
#Include <Misc\CountLibraries>
#Include <App\Gimp>
#Include <App\Shows>
#Include <Misc\Calculator>
#Include <App\Explorer>
#Include <App\Browser>
; --------------------------------------------------------------------------------
#Include <Scr\GeneralKeyChorder>
#Include <Scr\ExpenseReportKeyChorder>
#Include <Directives\_setup>
#Include <Tools\KeycodeGetter>
#Include <Tools\FileSystemSearch>
#include <System\UIA>
#Include <Scr\Keys\VimMode>
#Include <Utils\GetFilesSortedByDate>
; --------------------------------------------------------------------------------

SetCapsLockState("AlwaysOff")
; --------------------------------------------------------------------------------

toggleCapsLock()
{
    SetCapsLockState(!GetKeyState('CapsLock', 'T'))
}
; #j:: {
#HotIf WinActive("Chrome River - Google Chrome")
*^s::SaveCR()
SaveCR(){
	expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
	Sleep(100)
	expRpt.FindElement({Type: '50000 (Button)', Name: "Save", LocalizedType: "button", AutomationId: "save-btn"}).Highlight(100).Invoke()
	
}
#HotIf
CapsLock:: {
	
	if !input := CleanInputBox().WaitForInput() {
		return false
	}

	static runner_commands := Map(

		"libs?",	() => Infos(CountLibraries()),
		"drop",		() => Shows.DeleteShow(true),
		"finish",	() => Shows.DeleteShow(false),
		"show",		() => Shows.Run("episode"),
		; "down",		() => Shows.Run("downloads"),
		"down",		() => Run(Paths.Downloads),
		'start',	() => Run(A_StartUp),
		'setup',	() => Setups(),
		'horizon',	() => Run('C:\Program Files\FMGlobal\Horizon\hznHorizonMgr.exe'),
		'winspector',() => Run(Paths.OneDrive "AHK.Main\Winspector\WinspectorU.exe"),
		'scr4',		() => Run("C:\Users\bacona\AppData\Local\Programs\AutoHotkey\AHK.Projects.v1\Scriptlet_Library_v4.ahk"),
		; 'bmle', () => Run('C:\Users\' A_UserName '\FM Global\Operating Standards - Documents\general\B+M Loss Expectancy Guide - July 2023.xlsx'),
		'bmle', () => Run('excel.exe /e ' '"' Paths.FMGlobal '\Operating Standards - Documents\general\B+M Loss Expectancy Guide - July 2023.xlsx"'),
		'user', () => Infos(A_Username),
		; 'run', () => Infos(val := GetFilesSortedByDate('runner.ahk')),
		'code', () => Infos(val := GetFilesSortedByDate(Paths.Code)),
		'edit runner', () => VSCodeEdit(Paths.Lib '\Scr\Runner.ahk'),
		; 'edit runner', () => VSCodeEdit(, Paths.Lib '\Scr\runner.ahk'),
		; 'edit runner', () => VSCodeEdit(, 'runner.ahk'),
		; 'edit', (input) => VSCodeEdit(, input),
		; 'runner', () => VSCodeEdit('c:\Users\bacona\OneDrive - FM Global\Documents\AutoHotkey\Lib\Scr\Runner.ahk'),
		; 'runner', () => Run(Paths.Code ' "c:\Users\bacona\OneDrive - FM Global\Documents\AutoHotkey\Lib\Scr\Runner.ahk"'),
		; 'runner', () => EditFile(Paths.Lib '\Scr\Runner.ahk'),
		; 'runner', RunWithVSCode(Paths.code, Paths.Lib '\Scr\Runner.ahk'),
		'hznp', () => EditFile(Paths.Lib '\HznPlus.v2.ahk'),
		'aprocess', () => EditFile(Paths.Lib '\A_Process.v2.ahk'),
		'bdrive', () => Run('\\corp\data\San Francisco\Engineering\AutoHotkey'),
		'install', () => Run('edit ' '\\corp\data\San Francisco\Engineering\AutoHotkey\Install_Script.V2.ahk'),
		'visit', visitplanner,
		'get', (runner := GetFilesSortedByDate('Runner.ahk')) => Infos(runner), ;! test func ; finds and shows proper file and path
		; 'e', exitvp,
		; 'c', expensereport,
		; 'cc', CompanyCar,
		; 'cs', CarService,
		; 'ef', EmployeeMeal,
		; 'm', EmployeeMeal,
		; 'h', Hotel,
		; 'os', Office_Expenses,
		; 'af', AirlineFee,
		; 'rc', RentalCar,
		; 'ta', TravelAgency,
		; 'wc', WorkClothes,
		; 'sav', Save,
		; 'fee',FeeType,
		; 'sfee',SeatFee,
		; 'carservice',Car_Service,
		; 'tr',TripPurpose, 
		'bnm',() => bmle(),
		; 'bdrive', () => RunWait("Powershell.exe -WindowStyle hidden -Command -LiteralPath " '\\corp\data\San Francisco\Engineering\AutoHotkey' " -DestinationPath " "") ,
	
		; 'info test', Run(Paths.AppDataProgs),
		; "gimp", 	() => Gimp.winObj.RunAct(),
		; "davinci", 	() => Davinci.winObj.RunAct(),

		; fix RunAct_Folders() => dunno where that is
		; "ext",		() => Explorer.WinObjs.VsCodeExtensions.RunAct_Folders(), 
		; fix RunAct_Folders() => dunno where that is
		; "saved",	() => Explorer.WinObjs.SavedScreenshots.RunAct_Folders(), 
		; fix RunAct_Folders() => dunno where that is
		; "screenshots", () => Explorer.WinObjs.Screenshots.RunAct_Folders(), 
		'vim', () => Environment.VimMode := !Environment.VimMode,
		'yy', () => Environment.VimMode := !Environment.VimMode,
		; 'main',		() => Win.App(Paths.Prog '\AHK Script.v2.ahk'),
		; 'main',	_Install_Git,
		'main',	_CheckUpdate,
		; 'main', () => Infos(DriveGetFileSystem('https://fmglobal.sharepoint.com/:u:/r/teams/AutoHotKeyUserGroup/Shared Documents/General/Starter Script Files and Guide - V2/Lib')),
		; 'main',	bDriveStatus,
		; 'main',	() => Infos(bStatus := DriveGetStatus('\\corp\data\')),
		; 'note', (input) => _NoteOpen(input),
		'notes', () => Run(Paths.Lib '\Notes') ,
		'links', () => Run(Paths.Lib '\Links'),
		; 'n', _RunEnvNoter ,
		'Lib', () => Run(Paths.Lib),
		'lib', () => Run(Paths.Lib),
		'scr', () => Run(Paths.Lib '\Scr'),
		'lnchr', () => Run(Paths.lnchr),
		'run lnchr', () => Run(Paths.lnchr '\LNCHR-Main.ahk'),
		'test', () => uFile(3,0,1),
		'key', () => KeyCodeGetter,
		'uia', () => Run(Paths.Lib '\System\UIA.ahk'),
		; 'bmle', () => FileSystemSearch(Paths.FMGlobal) => FileSystemSearch.StartSearch(),
	)
	; --------------------------------------------------------------------------------

	; --------------------------------------------------------------------------------
		/**
	 * Syntax sugar: "Run *this* using *this program*"
	 * @param with *String* The program to use to run something with: either Program.exe format, or the full path to the executable
	 * @param runFile *String* The path to the file (or link!) you want to run
	 */
	static RunWithVSCode(with, runFile) => Run(with ' "' runFile '"')
	; --------------------------------------------------------------------------------
	static bmle(input:='') {
		; bmleg := '"' Paths.FMGlobal '\Operating Standards - Documents\general\B+M Loss Expectancy Guide - July 2023.xlsx"'
		; bmleg := "C:\Users\" A_UserName "\FM Global\Operating Standards - Documents\general\B+M Loss Expectancy Guide - July 2023.xlsx"
		bmleg := GetFilesSortedByDate(Paths.FMGlobal 'B+M Loss Expectancy Guide*')
		Infos(bmleg)
		; fR := FileRead(bnm)
		Loop Read bmleg {
			; if InStr(A_LoopReadLine,'Switchgear'){
				Infos(A_LoopReadLine)
			; }
		}
	}

	static visitplanner(search := '') {
		vpLink := 'https://app.fmglobal.com/polaris/assignments/'
		; vpLink := 'https://prodfmgidp.b2clogin.com/prodfmgidp.onmicrosoft.com/b2c_1a_main_signup_signin/oauth2/v2.0/authorize?client_id=c3115488-df7a-48b6-afdf-5c9bf9bb50a8&scope=c3115488-df7a-48b6-afdf-5c9bf9bb50a8%20openid%20profile%20offline_access&redirect_uri=https%3A%2F%2Fapp.fmglobal.com%2Fpolaris%2Fassignments%2F&client-request-id=e45f427d-e897-4aec-bcec-da88fdd76839&response_mode=fragment&response_type=code&x-client-SKU=msal.js.browser&x-client-VER=2.32.2&client_info=1&code_challenge=X2m2_jDyAxSZxhE1e5VKxk6f5B0EmKWVqrjHg3LJW2U&code_challenge_method=S256&nonce=22c37f1f-0dc0-4868-ba06-2a51b08df346&state=eyJpZCI6ImI2NmQyMzY0LTdiNjMtNDQwYi04MGU4LTYyNWNjNTI4OTM2MiIsIm1ldGEiOnsiaW50ZXJhY3Rpb25UeXBlIjoicmVkaXJlY3QifX0%3D'

		login()
		login(){
			RunWait(vpLink)
			; hWe := WinExist(pIDvp)
			; WinWaitActive(hWe)
			; Sleep(100)
			WinWaitActive('Sign In - Google Chrome') || WinWaitActive('Polaris - Assignments - Google Chrome')
			; WinWaitActive('Sign In - Google Chrome')
			vp := UIA.ElementFromChromium('A',false,1000)
			evp := vp.FindElement({AutomationId: 'signInName'})
			evp.Value := 'adam.bacon@fmglobal.com'
			; vvp := vp.FindElement({AutomationId: 'signInName'})
			; vvp.Value := 'adam.bacon@fmglobal.com'
			vp.FindElement({Name:'Continue', AutomationId:'next'}).Invoke()
			Sleep(1000)
			wE := WinExist('Polaris - Assignments - Google Chrome')
			; Infos(we '`n' WinGetTitle(wE))
			; WinActivate(wE)
			; WinWaitActive(wE)
			vpN := UIA.ElementFromChromium().FindElement({Type:'button', Name:'Load More'})
			vpN.Invoke()
		}
	}
	; --------------------------------------------------------------------------------
	static exitvp(){
		ehW := WinExist('Polaris ')
		WinActivate(ehW)
		exitV := UIA.ElementFromChromium(WinActive())
		vv := exitV.FindElement({Type:'MenuItem', Name: 'Profile'})
		vv.Invoke()
		Sleep(100)
		vvl := exitV.FindElement({Type:'MenuItem', Name: 'Log Out'})
		vvl.Invoke()
		Sleep(100)
		; Send('{Tab 3}{Enter}')
		Sleep(100)
		wE := WinExist('FM Global - Google Chrome')
		WinActivate(wE)
		wWwE := WinWaitActive(wE)
		wA := WinActive()
		iNFOS(wWwE '`n' WinGetTitle(Wa))
		Sleep(1000)
		ControlSend('{Ctrl down}{F4}{Ctrl up}',,WinGetTitle(winactive('A')))
		ControlSend('F4',,WinGetTitle(winactive('A')))
		ControlSend('{Ctrl up}',,WinGetTitle(winactive('A')))
		ControlSend('{Ctrl down}{F4}{Ctrl up}',,wA)
		ControlSend('F4',,wA)
		ControlSend('{Ctrl up}',,wA)
	}
	; --------------------------------------------------------------------------------
	static EditFile(filename, dir := '') {
		SplitPath(filename,&sFile,&sDir,&sExt,&sName,&sDrive)
		; Infos(filename '`n' sFile '`n' sDir '`n' sExt '`n' sName '`n' sDrive '`n' WinActive(sFile ' - ' 'Visual Studio Code'))
		; return
		if (!dir && we:= WinExist(sFile ' - ' 'Visual Studio Code')) {
			; WinActivate(,sFile ' - ' 'Visual Studio Code')
			WinActivate(we)
			Infos('[WinActive] ' sFile ' was already running!!!', 5000)
		} 
		; results := FileExist(filename)
		; Infos(results)
		else {
			Run(Paths.Code ' "' filename '"')
			Infos('Running ' Paths.Code ' ' '"' filename '"',3000)
		}
	}
	static VSCodeEdit(file := '') {
		; Infos('start`n' A_ScriptName)
		static folder := Paths.code
		; Infos(Paths.code)
		; Infos(folder)
		; return
		aMatch := []
		aFolder := [(Paths.Lib), (Paths.v2Proj), (Paths.Prog)]
		aIgnore := ['.git', '.history', '.vscode', '.Other']
		fNeedle := '\.*'
		lFNeedle := '\*.ahk'
		vF := ''
		vI := ''
		vS := ''
		vL := ''
		list := ''
		alist := []
		blist := []
		mAlist := Map()
		; RegExMatch(file, fNeedle, &fMatch)
		try {
			; Infos(
			; WinExist(file ' - Visual Studio Code') ? WinActivate(file ' - Visual Studio Code') : WinActivate(file)
			SplitPath(file,&sFile)
			tFile := WinGetTitle(sFile)
			; Infos(
				; sFile
				; '`n'
				; tFile
				; , 10000 )
			fWinE := WinExist(tFile) 
			; || 
			; fWinE := WinExist(sFile ' - Visual Studio Code')
			; fWinE := WinExist(sFile ' - Visual Studio Code') ? WinActivate(fWinA) : WinActivate(sFile ' - Visual Studio Code')
			If fWinE {
				WinActivate(fWinE)
				Infos('[WinActive] ' sFile ' was already running!!!', 5000)
			} else {
				WinActivate(sFile ' - Visual Studio Code')
				Infos('[WinActive] ' sFile ' was already running!!!', 5000)
			}
			; fWinA := WinExist(sFile ' - Visual Studio Code') ? WinActivate(fWinA) : WinActivate(sFile ' - Visual Studio Code')
			; WinActive(sFile ' - Visual Studio Code') ? Infos('[WinActive] ' sFile ' was already running!!!', 5000) : Infos('A: ' WinActivate(sFile))
		} catch {
			CoordMode('mouse', 'Window')
			Run(A_ScriptDir)
			number := WinWaitActive(A_ScriptDir)
			Sleep(10)
			Send('r')
			Sleep(100)
			Send('{AppsKey Down}')
			Sleep(50)
			Send('{AppsKey Up}')
			Sleep(100)
			Send('i')
			WinActivate(number)
			WinGetClientPos(&x, &y,&w,&h,'A')
			; x := A_ScreenDPI / A_DPI
			; y := A_ScreenDPI / A_DPI
			adj := A_DPI / A_ScreenDPI 
			x := adj + (w/2)
			y := adj + (h/2)
			MouseMove(x, y, -1)
			; hCtl := ControlGetFocus(number)
			; nCtl := ControlGetClassNN(hCtl)
			; Infos(nCtl)

			; Run('"' folder ' ' file '"')
		}
		return
		for vF in aFolder {
			count := aFolder.Length
			aList.Capacity := count
			Loop count {
				aName := 'aList' A_Index
				aList.Push(aName)
			}
		}
		for each, value in aFolder {
			mAlist.Set(value, each)

		}
	}

	static _RunEnvNoter() {
		SendLevel(5)
		Send('#h')
		Sleep(300)
		SendLevel(0)
		Send('n')
		return
	}
	static runner_regex := Map(

		"go",      (input) => _GitLinkOpenCopy(input),
		"gl",      (input) => ClipSend(Git.Link(input),, false),
		"cp",      (input) => (A_Clipboard := input, Info('"' input '" copied')),
		"rap",     (input) => Spotify.NewRapper(input),
		"fav",     (input) => Spotify.FavRapper(input),
		"disc",    (input) => Spotify.NewDiscovery(input),
		"link",    (input) => Shows.SetLink(input),
		"ep",      (input) => Shows.SetEpisode(input),
		"finish",  (input) => Shows._OperateConsumed(input, false),
		"dd",      (input) => Shows.SetDownloaded(input),
		"drop",    (input) => Shows._OperateConsumed(input, true),
		"relink",  (input) => Shows.UpdateLink(input),
		"ev",      (input) => Infos(Calculator(input)),
		"evp",     (input) => ClipSend(Calculator(input)),
		'o',       (input) => _LinkOpen(input),
		

	)

	; static runner_regex := Map(
	; 	"go",      	(input) => _GitLinkOpenCopy(input),
	; 	"gl",      	(input) => ClipSend(Git.Link(input),, false),
	; 	"p",       	(input) => _LinkPaste(input),
	; 	"o",       	(input) => _LinkOpen(input),
	; 	; "o",       	(input) => _NoteOpen(input),
	; 	"cp",      	(input) => (Infos((A_Clipboard := input) '" copied')),
	; 	"rap",     	(input) => Spotify.NewRapper(input),
	; 	"fav",     	(input) => Spotify.FavRapper(input),
	; 	"disc",    	(input) => Spotify.NewDiscovery(input),
	; 	"link",    	(input) => Shows.SetLink(input),
	; 	"ep",      	(input) => Shows.SetEpisode(input),
	; 	"finish",  	(input) => Shows._OperateConsumed(input, false),
	; 	"dd",      	(input) => Shows.SetDownloaded(input),
	; 	"drop",    	(input) => Shows._OperateConsumed(input, true),
	; 	"relink",  	(input) => Shows.UpdateLink(input),
	; 	"ev",      	(input) => Infos(Calculator(input)),
	; 	"evp",     	(input) => ClipSend(Calculator(input)),
	; 	'addlib'	(input) => CleanInputBox.SetInput(git_InstallAHKLibrary(input)),
	; )

	if runner_commands.Has(input) {
		runner_commands[input].Call()
		return
	}

	regex := "^("
	for key, _ in runner_regex {
		regex .= key "|"
	}
	regex .= ") (.+)"
	result := input.RegexMatch(regex)
	if runner_regex.Has(result[1]){
		runner_regex[result[1]].Call(result[2])
	}
	static _GitLinkOpenCopy(input) {
		link := Git.Link(input)
		Browser.RunLink(link)
		A_Clipboard := link
	}

	static _LinkPaste(input) {
		link := Environment.Links.Choose(input)
		if !link
			return
		ClipSend(link,, false)
	}

	static _LinkOpen(input?) {
		if (input) {
			Browser.RunLink(input)
		}

		if !input := CleanInputBox().WaitForInput()
			return

		link := Environment.Links.Choose(input)
		if (!link){
			return
		}
		Browser.RunLink(link)
	}
	static _NoteOpen(input) {
		note := Environment.Notes.Choose(input)
		if !note
			return
		Browser.RunLink(note)
	}

	static _bdrive(){
        ; static b_drive := '\\corp\data\San Francisco\Engineering\AutoHotkey'
        static b_drive := '\\corp\data'
        ; static b_drive := 'https://fmglobal.sharepoint.com/:u:/r/teams/AutoHotKeyUserGroup/Shared%20Documents/General/'
        ; bStatus := DirExist(b_drive)
        bExist := DirExist(b_drive)
		; bStatus := GetHtml(b_drive)
        bStatus := (bExist = 'D') ? 'Ready' : 'Offline'
        Infos(bStatus, 5000)
        return bStatus
    }

	static _getWord(){
		delimiter1 := A_Space '```(`)`,.`=/\+-*!@#$`%^&*?<>~`;`:`{`}`[`]'
		word := GetWordUnderMouse(delimiter1)
		Infos(word)


		GetWordUnderMouse(delim) 
		{ 
		; delim = character(s) specified to separate words
		; if no delimiter character(s) specified then the entire line of text 
		; under the mouse pointer will be returned.
		VarSetStrCapacity(&word, 2048) 
		; RetVal := DllCall("gwrd.dll\GetWord", "str", word, "str", delim, "Uint")
		; RetVal := DllCall("gwetw\GetWord", "str", word, "str", delim, "Uint")
		; The DLL returns the handle to the control the mouse is over 
		Return word
		}
	}

	static _Install_Git(){
		Git_Link := 'https://github.com/git-for-windows/git-for-windows.github.io/blob/main/latest-64-bit-installer.url'

		if GetHtml(Git_Link) {
			new_Link := git_InstallAHKLibrary(Git_Link)
		}
		;! Custom Setup & Silent install
		; https://github.com/git-for-windows/git/wiki/Silent-or-Unattended-Installation
	}
	static _Install_VSCode(){
		vscode_link := 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user'

		if GetHtml(vscode_link) {
			;fix ..: after verifying internet connectivity, perform a silent install of vscode
		}
	}
	static bDriveStatus() {
		; drive := '\\corp\data\'
		drive := 'https://fmglobal.sharepoint.com/:u:/r/teams/AutoHotKeyUserGroup/Shared%20Documents/General/'
		bStatus := DriveGetStatus(drive)
		Infos(bStatus)
		return bStatus
	}
	static _is_bdrive(){
		; static b_drive := '\\corp\data\San Francisco\Engineering\AutoHotkey'
		static b_drive := '\\corp\data\', driveType:='network'
		drive := ''
		tDrive := ''
		is_dir := DriveGetList()
		dArray := StrSplit(is_dir)
		try {
			for drive in dArray {
				drive := drive ':\'
				dType := DriveGetType(drive)
				tDrive .= '[' drive '] | ' dType ' | ' DriveGetStatus(drive) '`n'
				if (drive == 'B:\') || ((dType == driveType) && (DriveGetStatus(b_drive) == 'Ready')){
					bStatus := DriveGetStatus(drive)
					Infos('[B:\] ' bStatus)
					return bStatus
				}
			}
		}
		catch {
			bStatus := DriveGetStatus(b_drive)
			Infos('[B]' bStatus)
			return bStatus
		}
	}

	bDrive_NotConnected(fCounter := 5, fC_units:='sec', sCounter:= 2, sC_units := 'sec') {
		static cSec := 1000
		static cMin := (cSec * 60)
		static cHrs := (cMin * 60)
		static cDays := (cHrs * 24)
		; --------------------------------------------------------------------------------
		;? fMultiply = first multiplier => if = sec|min|hr|[days], then multiply (fCounter * fMultiply)
		fMultiply := (fC_units = 'sec') ? cSec : (fC_units = 'min') ? cMin : (fC_units = 'hr') ? cHrs : cDays
		;? sMultiply = second multiplier => if = sec|min|hr|[days], then multiply (sCounter * sMultiply)
		sMultiply := (sC_units = 'sec') ? cSec : (sC_units = 'min') ? cMin : (sC_units = 'hr') ? cHrs : cDays
		; --------------------------------------------------------------------------------
		fCounter := (fCounter * fMultiply)
		sCounter := (sCounter * sMultiply)
		counter := FMCounter()
		counter.Start()
		Sleep(fCounter)
		counter.Stop()
		Sleep(sCounter)
		return
	}

	; --------------------------------------------------------------------------------
	static _CheckUpdate() {
		static b_drive 		:= '\\corp\data\San Francisco\Engineering\AutoHotkey'
		static uFolder 		:= '\CheckUpdate'
		static uFile 		:= 'ScriptVersionMap.ahk'
		static diff 		:= ''
		static Lib 			:= Paths.Lib
		static lib_uPath 	:= Lib uFolder '\' uFile
		static vNeedle 		:= "i)m_ver := '\d.\d.\d'"
		static b_uPath 	:= b_drive uFolder '\' uFile
		;? set up the variables to simplify things visually, and setup a map() for simpler updating
		static m:='major', mi:='minor', p:='patch'
		vArray := []
		
		; open_file := FileOpen(lib_uPath,'rw','UTF-8'), (open_file ~= vNeedle) ? 'true' : 'false' 
		rFile := FileRead(lib_uPath,'UTF-8') ;? can do a regex match from this one
		vFLocal := vF_Check(rFile)
		; open_file := FileOpen(b_uPath,'rw','UTF-8'), (open_file ~= vNeedle) ? 'true' : 'false'
		rFile := FileRead(b_uPath,'UTF-8') ;? can do a regex match from this one
		vFRemote := vF_Check(rFile)
		; Infos(; 	vFRemote[m] ' : ' vFLocal[m] '`n' vFRemote[mi] ' : ' vFLocal[mi] '`n' vFRemote[p] ' : ' vFLocal[p])
		; --------------------------------------------------------------------------
		vArray := []
		major := (vFRemote[m] = vFLocal[m]) ? true : false
		minor := (vFRemote[mi] = vFLocal[mi]) ? true : false
		patch := (vFRemote[p] = vFLocal[p]) ? true : false
		; infos(major '`n' minor '`n' patch)
		vArray.Push(major, minor, patch)
		; --------------------------------------------------------------------------
		;? Use an array, combined with a switch - case, to simplify if/else if/else 
		vVer := ''
		for each, vVer in vArray {
			switch vVer {
				case  true:
					if ((major = true) && (minor = true) && (patch = true)) {
						Infos(
							'No need to update at this time.`n'
							'Local Version: ' vFLocal[m] '.' vFLocal[mi] '.' vFLocal[p] '`n'
							'Remote Version: ' vFRemote[m] '.' vFRemote[mi] '.' vFRemote[p], 10000)
						return
					}
				case false:
					if (!(major = true) || !(minor != true) || !(patch != true)) {
						Infos('false')
						; fix ..: if false, download crap
					}
			}
			; infos(major '`n' minor '`n' patch) ;? validation
		}
		; --------------------------------------------------------------------------
	}	
	static vF_Check(cFile){
		static m:='major', mi:='minor', p:='patch'
		version_map := Map(), vNeedle := "'\d.\d.\d'", ver := ''
		;! requires FileRead() for regexmatch; read file if not already done
		; rFile := FileRead(cFile,'UTF-8') 
		RegExMatch(cFile, vNeedle, &ver_match)
		;? Remove the '' and . ; place each ver in a variable
		for each, ver in ver_match {
			version := StrReplace(ver, "'", ''), vers := StrSplit(version, '.',,3)
			major := vers[1], minor := vers[2], patch := vers[3]
		}
		; Infos(ver '`n' version '`n' major '`n' minor '`n' patch) ;? validation
		;? create a map() for the variable values
		version_map.Set(m, major, mi, minor, p, patch)
		;* return the map() to the calling function
		return version_map
	}
	
	static uFile(m, mi, p) {
		static b_drive 		:= '\\corp\data\San Francisco\Engineering\AutoHotkey'
		static uFolder 		:= '\CheckUpdate'
		static uFile 		:= 'ScriptVersionMap.ahk'
		static diff 		:= ''
		static Lib 			:= Paths.Lib
		static bFile 		:= b_drive uFolder '\' uFile
		static lFile 		:= Lib uFolder '\' uFile
		static fNeedle := "'[0-9].[0-9].[0-9]'"
		; static m:='major', mi:='minor', p:='patch' ;? vMap.Set(m, major, mi, minor, p, patch)
		; m:='', mi:='', p:=''
		vMap := Map()
		ver := ''
		folder_array := [bFile, lFile]
		aFile := []
		aMatch := []
		mArray := []
		bArray := []
		lArray := []
		fLine := ''
		f := ''
		match:=''
		aLine := ''
		; --------------------------------------------------------------------------------
		bfo := FileOpen(bFile, 'rw', 'UTF-8')
		lfo := FileOpen(lFile, 'rw', 'UTF-8')
		bArray := vReadFile(bFile)
		lArray := vReadFile(lFile)
		; bArray := vReadFile(bfo)
		; lArray := vReadFile(lfo)
		vReadFile(vFile){
			; Infos(vFile)
			; rFile := FileRead(vFile,'UTF-8')
			Sleep(300)
			loop read vFile {
				aFile.Push(A_LoopReadLine)
				fLine .= A_LoopReadLine . '`n'
			}
			return aFile
		}
		; --------------------------------------------------------------------------------
		new_bFile := vMatchLine(bArray, bFile)
		vMatchLine(vArray, file_Open) {
			Infos(A_Index FileRead(file_open))
			RegExMatch(file_Open, fNeedle, &vMatch_Array)
			for each, value in vMatch_Array {
				Infos(value)
			}
			for each, aLine in aFile {
				; Infos(aLine '`n' fNeedle)
					if (aLine ~= fNeedle) {
						Infos(aLine)
						version := StrReplace(aLine, "'", ''), vers := StrSplit(version, '.',,3)
						major := vers[1], minor := vers[2], patch := vers[3]
					} 
				; }
				; str_match := StrSplit(aLine,')','i ) "')
				; rMatch := str_match[2]
				; Infos(rMatch)
				; ver_match := mArray[rMatch]
				; ;? Remove the '' and . ; place each ver in a variable
				; for each, ver in ver_match {
				; 	version := StrReplace(ver, "'", ''), vers := StrSplit(version, '.',,3)
				; 	major := vers[1], minor := vers[2], patch := vers[3]
				; }
				; Infos(ver '`n' version '`n' major '`n' minor '`n' patch) ;? validation
				;? create a map() for the variable values
				; vMap.Set(m, major, mi, minor, p, patch)
				nM  := RegExReplace(major, '([0-9]{1})', m)
				nMi := RegExReplace(minor, '([0-9]{1})', mi)
				nP  := RegExReplace(patch, '([0-9]{1})', p)
				aLine := "'" nM '.' nMi '.' nP "'"
				aFile.RemoveAt(A_Index)
				aFile.InsertAt(A_Index, aLine)
			}
			; --------------------------------------------------------------------------------
			; --------------------------------------------------------------------------------
			/**
			 * function: write each value in the array to @param f (string variable) 
			 */
			f .= aLine . '`n'
			; --------------------------------------------------------------------------------
			rf := file_Open.Write(f)
			Infos(f)
			file_Open := ''
			return rf
		}
	}
	; --------------------------------------------------------------------------------
	; hzntx11 := FileOpen(file_name,'rw','UTF-8')
	return 0
}
; }

; An example class for counting the seconds...
class FMCounter {
	__New() {
		this.interval := 1000
		this.count := 0
		; Tick() has an implicit parameter "this" which is a reference to
		; the object, so we need to create a function which encapsulates
		; "this" and the method to call:
		this.timer := ObjBindMethod(this, "Tick")
	}
	Start() {
		SetTimer( this.timer, this.interval)
		Infos("Counter started")
	}
	Stop() {
		; To turn off the timer, we must pass the same object as before:
		SetTimer( this.timer, 0)
		Infos("Counter stopped at " this.count)
	}
	; In this example, the timer calls this method:
	Tick() {
		Infos( ++this.count)
	}
	MyCallBack() {
		SetTimer( this.timer, 1000)
	}
}
