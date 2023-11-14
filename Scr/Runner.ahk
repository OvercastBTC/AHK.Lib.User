; --------------------------------------------------------------------------------
#Requires AutoHotkey v2
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
#Include <Converters\DateTime>
#Include <Tools\CleanInputBox>
#Include <Misc\Meditate>
#Include <Misc\CountLibraries>
#Include <App\Gimp>
#Include <App\Shows>
#Include <Misc\Calculator>
#Include <App\Explorer>
; --------------------------------------------------------------------------------
#Include <Scr\GeneralKeyChorder>
; --------------------------------------------------------------------------------


SetCapsLockState("AlwaysOff")
; --------------------------------------------------------------------------------

toggleCapsLock()
{
    SetCapsLockState(!GetKeyState('CapsLock', 'T'))
}
; #j:: {
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
		"down",		() => Explorer.WinObjs.Run(Paths.Downloads),

		; "gimp", 	() => Gimp.winObj.RunAct(),
		; "davinci", 	() => Davinci.winObj.RunAct(),

		; "ext",		() => Explorer.WinObjs.VsCodeExtensions.RunAct_Folders(), ; fix RunAct_Folders() => dunno where that is
		; "saved",	() => Explorer.WinObjs.SavedScreenshots.RunAct_Folders(), ; fix RunAct_Folders() => dunno where that is
		; "screenshots", () => Explorer.WinObjs.Screenshots.RunAct_Folders(), ; fix RunAct_Folders() => dunno where that is
		'vim', 		() => Environment.VimMode:=true,
		; 'main',		() => Win.App(Paths.Prog '\AHK Script.v2.ahk'),
		; 'main',	_Install_Git,
		'main',	_CheckUpdate,
		; 'main',	_is_bdrive,
	)

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
		link := Environment.Linoks.Choose(input)
		if !link
			return
		ClipSend(link,, false)
	}

	static _LinkOpen(input?) {
		if (input) {
			Browser.RunLink(link)
		}

		if !input := CleanInputBox().WaitForInput()
			return

		link := Environment.Links.Choose(input)
		if (!link){
			return
		}
		Browser.RunLink(link)
	}
	; static _NoteOpen(input) {
	; 	note := Environment.Notes.Choose(input)
	; 	if !note
	; 		return
	; 	Browser.RunLink(note)
	; }
	static _NoteOpen(input?) {
		if !input := CleanInputBox().WaitForInput()
			return
		note := Environment.Notes.Choose(input)
		if !(note) {
			return
		}
		A_Clipboard := note
		Infos(note)
	}

	static _Install_Git(){
		Git_Link := 'https://github.com/git-for-windows/git-for-windows.github.io/blob/main/latest-64-bit-installer.url'

		if GetHtml(Git_Link) {
			new_Link := git_InstallAHKLibrary(Git_Link)
		}
	}
	static _Install_VSCode(){
		vscode_link := 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user'

		if GetHtml(vscode_link) {
			;fix ..: after verifying internet connectivity, perform a silent install of vscode
		}
	}
	
	static _is_bdrive(){
		is_dir := DirExist('\\corp\data\San Francisco\Engineering\AutoHotkey')
		Infos(is_dir)
	}

	static _CheckUpdate() {
		static b_drive 			:= '\\corp\data\San Francisco\Engineering\AutoHotkey'
		static CheckUpdate 		:= '\CheckUpdate'
		static UpdateFile 		:= 'ScriptVersionMap.ahk'
		static diff_tool 		:= ''
		static Lib 				:= Paths.Lib
		static lib_UpdateFile 	:= Lib CheckUpdate '\' UpdateFile
		static ver_needle 			:= "i)m_ver := '\d.\d.\d'"
		static update_ver_file 	:= b_drive CheckUpdate '\' UpdateFile
		;? set up the variables to simplify things visually, and setup a map() for simpler updating
		static m:='major', mi:='minor', p:='patch'
		
		; open_file := FileOpen(lib_UpdateFile,'rw','UTF-8'), (open_file ~= ver_needle) ? 'true' : 'false' 
		read_file := FileRead(lib_UpdateFile,'UTF-8') ;? can do a regex match from this one
		vFLocal := file_vers_check(read_file)
		; open_file := FileOpen(update_ver_file,'rw','UTF-8'), (open_file ~= ver_needle) ? 'true' : 'false'
		read_file := FileRead(update_ver_file,'UTF-8') ;? can do a regex match from this one
		vFRemote := file_vers_check(read_file)
		; Infos(; 	vFRemote[m] ' : ' vFLocal[m] '`n' vFRemote[mi] ' : ' vFLocal[mi] '`n' vFRemote[p] ' : ' vFLocal[p])
		; --------------------------------------------------------------------------
		version_array := []
		major := (vFRemote[m] = vFLocal[m]) ? true : false
		minor := (vFRemote[mi] = vFLocal[mi]) ? true : false
		patch := (vFRemote[p] = vFLocal[p]) ? true : false
		; infos(major '`n' minor '`n' patch)
		version_array.Push(major, minor, patch)
		; --------------------------------------------------------------------------
		;? Use an array, combined with a switch - case, to simplify if/else if/else 
		vVer := ''
		for each, vVer in version_array {
			switch vVer {
				case  true:
					if ((major = true) && (minor = true) && (patch = true)) {
						Infos('true')
						return
					}
				case false:
					if (!(major = true) || !(minor != true) || !(patch != true)) {
						Infos('false')
						; fix ..: if false, download crap
					}
			}
			; infos(major '`n' minor '`n' patch)
		}
		; --------------------------------------------------------------------------
		file_vers_check(check_file){
			version_map := Map()
			ver_needle := "'\d.\d.\d'"
			; read_file := FileRead(check_file,'UTF-8') ;? requires FileRead() to do a regex match
			RegExMatch(check_file, ver_needle, &ver_match)
			ver := ''
			for each, ver in ver_match {
				; version := StrReplace(ver, "'", ''), vers := StrSplit(version, '.',,3)
				;? Expanded for edification
				;* Remove the '' around the numbers
				version := StrReplace(ver, "'", '')
				;* Remove the . between the numbers
				vers := StrSplit(version, '.',,3)
				; major := vers[1], minor := vers[2], patch := vers[3]
				;? Expanded for edification
				;* place each value in a variable
				major := vers[1]
				minor := vers[2]
				patch := vers[3]
			}
			; Infos(ver '`n' version '`n' major '`n' minor '`n' patch)
			;? create a map() for the variable values
			version_map.Set(m, major, mi, minor, p, patch)
			return version_map
		}
	}
}
