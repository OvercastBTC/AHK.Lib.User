; ---------------------------------------------------------------------------
#Requires AutoHotkey v2+
#Include <Directives\__AE.v2>
; ---------------------------------------------------------------------------
RTM := A_TrayMenu
RTM.Delete()
RTM.Add()
RTM.AddStandard()
; RTM.Show() ;? For a menu at the mouse
; ---------------------------------------------------------------------------
#Include <Includes\Includes_Runner>
; ---------------------------------------------------------------------------

; ---------------------------------------------------------------------------
SetCapsLockState("AlwaysOff")
; ---------------------------------------------------------------------------
; toggleCapsLock() {
; 	SetCapsLockState(!GetKeyState('CapsLock', 'T'))
; }
; ---------------------------------------------------------------------------



; ---------------------------------------------------------------------------
CapsLock:: {
	
	if !input := CleanInputBox().WaitForInput() {
		return false
	}

	static runner_commands := Map(

		"libs?",	() => Infos(CountLibraries()),
		; "drop",		() => Shows.DeleteShow(true),
		; "finish",	() => Shows.DeleteShow(false),
		; "show",		() => Shows.Run("episode"),
		'getkey', KeyCodeGetter,
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
		'runner', () => Run(Paths.Code ' "c:\Users\bacona\OneDrive - FM Global\Documents\AutoHotkey\Lib\Scr\Runner.ahk"'),
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
		'vim', 		() => Environment.VimMode := !Environment.VimMode,
		'yy', 		() => Environment.VimMode := !Environment.VimMode,
		'main',		() => Main.winObj.App(),
		; 'main',	_Install_Git,
		; 'main',	_CheckUpdate,
		'checkupdate',	_CheckUpdate,
		; 'main', () => Infos(DriveGetFileSystem('https://fmglobal.sharepoint.com/:u:/r/teams/AutoHotKeyUserGroup/Shared Documents/General/Starter Script Files and Guide - V2/Lib')),
		; 'main',	bDriveStatus,
		; 'main',	() => Infos(bStatus := DriveGetStatus('\\corp\data\')),
		; 'note', (input) => _NoteOpen(input),
		'notes', () => Run(Paths.Lib '\Notes') ,
		'links', () => Run(Paths.Lib '\Links'),
		; 'n', _RunEnvNoter ,
		'Lib', () => Run(Paths.Lib),
		'lib', () => Run(Paths.Lib),
		'Libv2', () => Run(Paths.v2Lib),
		'libv2', () => Run(Paths.v2Lib),
		'scr', () => Run(Paths.Lib '\Scr'),
		'lnchr', () => Run(Paths.lnchr),
		'run lnchr', () => Run(Paths.lnchr '\LNCHR-Main.ahk'),
		'test', () => uFile(3,0,1),
		'key', () => KeyCodeGetter,
		'uia', () => Run(Paths.v2Lib '\System\UIA.ahk'),
		; 'bmle', () => FileSystemSearch(Paths.FMGlobal) => FileSystemSearch.StartSearch(),
		'mytime',MyTime,
		'approvals',approvals,
		'map', makeCheatSheet,
		'os',OSGui,
		; 'grog',git_InstallAHKLibrary('https://github.com/GroggyOtter/ahkv2_definition_rewrite/blob/main/ahk2.d.ahk','C:\Users\bacona\.vscode\extensions\thqby.vscode-autohotkey2-lsp-2.2.8\syntaxes'),
		; 'grog',Install_Git_Lib('https://github.com/GroggyOtter/ahkv2_definition_rewrite/blob/main/ahk2.d.ahk','C:\Users\bacona\.vscode\extensions\thqby.vscode-autohotkey2-lsp-2.2.8\syntaxes'),
	)

	static runner_regex := Map(

		"go",      (input) => _GitLinkOpenCopy(input),
		"gl",      (input) => ClipSend(Git.Link(input),, false),
		"cp",      (input) => (A_Clipboard := input, Info('"' input '" copied')),
		"rap",     (input) => Spotify.NewRapper(input),
		"fav",     (input) => Spotify.FavRapper(input),
		"disc",    (input) => Spotify.NewDiscovery(input),
		; "link",    (input) => Shows.SetLink(input),
		; "ep",      (input) => Shows.SetEpisode(input),
		; "finish",  (input) => Shows._OperateConsumed(input, false),
		; "dd",      (input) => Shows.SetDownloaded(input),
		; "drop",    (input) => Shows._OperateConsumed(input, true),
		; "relink",  (input) => Shows.UpdateLink(input),
		"ev",      (input) => Infos(Calculator(input)),
		"evp",     (input) => ClipSend(Calculator(input)),
		'o',       (input) => _LinkOpen(input),
		

	)

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
		link := Environment.Links.Choose(input)
		if (input) {
			Browser.RunLink(link)
			return
		}

		if !input := CleanInputBox().WaitForInput()
			return

		if (!link){
			return
		}
		Browser.RunLink(link)
	}
	; ---------------------------------------------------------------------------
	static _NoteOpen(input) {
		note := Environment.Notes.Choose(input)
		if !note
			return
		Browser.RunLink(note)
	}
	; ---------------------------------------------------------------------------
	static _ShortcutOpen(input) {
		sc_text := Environment.shortcutkeys.Choose(input)
		if (!sc_text){
			Infos("No Match Found", 2000)
			return
		}
		Infos(sc_text)
		return
	}
	; ---------------------------------------------------------------------------
	static makeCheatSheet(){
		for key, value in runner_commands {
			Infos('key: ' key ' ' 'value: ' value)
		}
	}
	; --------------------------------------------------------------------------------
	static MyTime(){
		login()
		login(){
			aCtls := []
			activeWindow := 'EngNET - Work'
			nCtl := 'Internet Explorer_Server1'
			; wEx := WinExist()
			Edge.RunLink('https://engnet/EngNet/engnet/engnet.asp')
			WinWaitActive('EngNET - Work')
			WinActivate('ahk_exe msedge.exe')
			hCtl := ControlGetHwnd(nCtl, 'A')
			ControlFocus(hCtl)
			Sleep(1000)
			BlockInput(1)
			SendLevel(5)
			Sleep(1000)
			SendEvent('pw{Enter}')
			BlockInput(0)
			return
		}
	}
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

		login()
		login(){
			Run(vpLink)
			WaitElement_timeDelay := 30000
			; hWe := WinExist(pIDvp)
			; WinWaitActive(hWe)
			; Sleep(100)
			; WinWaitActive('Sign In - Google Chrome') || WinWaitActive('Polaris - Assignments - Google Chrome')
			WinWaitActive('Sign In - Google Chrome')
			vp := UIA.ElementFromChromium('A',false,WaitElement_timeDelay)
			vp.WaitElement({AutomationId: 'signInName'},WaitElement_timeDelay).Value := 'adam.bacon@fmglobal.com'
			vpC := UIA.ElementFromChromium('A',false,WaitElement_timeDelay)
			; vpC.WaitElement({Name:'Continue', AutomationId:'next'},WaitElement_timeDelay).Invoke()
			vpC.WaitElement({Name:'Continue', AutomationId:'next'},WaitElement_timeDelay).Click(,,,,true)
			; ---------------------------------------------------------------------------
			; WinWaitActive('Polaris - Assignments - Google Chrome')
			; vpN := UIA.ElementFromChromium('A',false,WaitElement_timeDelay)
			; vpN.WaitElement({Type:'button', Name:'Load More'},WaitElement_timeDelay).Invoke()
			; ---------------------------------------------------------------------------
			WinWaitActive('Polaris - Assignments - Google Chrome')
			try {
				fvL := vL.FindElement({Type: '50000 (Button)', Name: "All Loaded", LocalizedType: "button", AutomationId: "rds-button-8"})
			}
			if !fvL {
				Loop {
					vL := UIA.ElementFromChromium('A',false,WaitElement_timeDelay)
					try fvL := vL.FindElement({Type: '50000 (Button)', Name: "All Loaded", LocalizedType: "button", AutomationId: "rds-button-8"})
					vpN := UIA.ElementFromChromium('A',false,WaitElement_timeDelay)
					; vpN.WaitElement({Type:'button', Name:'Load More'},WaitElement_timeDelay).Invoke()
					vpN.WaitElement({Type:'button', Name:'Load More'},WaitElement_timeDelay).Click(,,,,true)
					; counter++
					Sleep(200)
				} until fvL := vL.FindElement({Type: '50000 (Button)', Name: "All Loaded", LocalizedType: "button", AutomationId: "rds-button-8"})
			} else {
				return
			}
			; vpN.Invoke()
			; HotIf()
		}
	}
	static approvals(search := '') {
		vpLink := 'https://www.approvalguide.com/'
		Title := 'FM Approvals - Approval Guide - '
		login()
		
		login(){
			Run(vpLink)
			WaitElement_timeDelay := 30000
			aG := UIA.ElementFromChromium('A',false,WaitElement_timeDelay).WaitElement({Type: '50000 (Button)', Name: "Log In", LocalizedType: "button"},WaitElement_timeDelay).Invoke()
			agE := UIA.ElementFromChromium('A',false,WaitElement_timeDelay).WaitElement({Type: '50004 (Edit)', Name: "Email Address", LocalizedType: "edit", AutomationId: "ContentPlaceHolder1_MFALoginControl1_UserIDView_txtUserid_UiInput"},WaitElement_timeDelay).Value := 'adam.bacon@fmglobal.com'
			; ---------------------------------------------------------------------------
			agL := UIA.ElementFromChromium('A',false,WaitElement_timeDelay).WaitElement({Type: '50005 (Link)', Name: "Submit", LocalizedType: "link"},WaitElement_timeDelay).Invoke()
			; ---------------------------------------------------------------------------
			pass := UIA.ElementFromChromium('A',false,WaitElement_timeDelay).WaitElement({Type: '50004 (Edit)', Name: "Password", LocalizedType: "edit", AutomationId: "ContentPlaceHolder1_MFALoginControl1_PasswordView_tbxPassword_UiInput"},WaitElement_timeDelay).Value := '80ab{*}{*}HD19KB'
			; ---------------------------------------------------------------------------
			; return
		}
	}
	
	; --------------------------------------------------------------------------------
	static exitvp(){
		ehW := WinExist('Polaris ')
		WaitElement_timeDelay := 30000
		WinActivate(ehW)
		exitV := UIA.ElementFromChromium('A',false,WaitElement_timeDelay).WaitElement({Type:'MenuItem', Name: 'Profile'}, WaitElement_timeDelay).Invoke()
		
		UIA.ElementFromChromium('A',false,WaitElement_timeDelay).WaitElement({Type:'MenuItem', Name: 'Log Out'}, WaitElement_timeDelay).Invoke()
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
		Sleep(500)
		; SendLevel(0)
		Send('n')
		; return
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
		static b_uPath 		:= b_drive uFolder '\' uFile
		; ---------------------------------------------------------------------------
		mainLib := 'C:\Users\' A_UserName '\AppData\Local\Programs\AutoHotkey\v2\Lib\TestingStuff'
		Infos('Does the directory exist???' DirExist(mainLib))
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
		static fNeedle := "([0-9]\.[0-9]\.[0-9])"
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
		return 0
	}
	; --------------------------------------------------------------------------------

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
