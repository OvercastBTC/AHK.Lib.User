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
#Include <.\Directives\__AE.v2>
#Requires AutoHotkey v2
; --------------------------------------------------------------------------------


SetCapsLockState("AlwaysOff")
; --------------------------------------------------------------------------------

toggleCapsLock()
{
    SetCapsLockState(!GetKeyState('CapsLock', 'T'))
}
; #j:: {
CapsLock:: {
	if !(input := CleanInputBox().WaitForInput()) {
		return false
		; return
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
		'main',		() => Win.App(Paths.Prog '\AHK Script.v2.ahk'),
	)

	static runner_regex := Map(
		"go",      	(input) => _GitLinkOpenCopy(input),
		"gl",      	(input) => ClipSend(Git.Link(input),, false),
		"p",       	(input) => _LinkPaste(input),
		"o",       (input) => _LinkOpen(input),
		; "o",       	(input) => _NoteOpen(input),
		"cp",      	(input) => (Infos((A_Clipboard := input) '" copied')),
		"rap",     	(input) => Spotify.NewRapper(input),
		"fav",     	(input) => Spotify.FavRapper(input),
		"disc",    	(input) => Spotify.NewDiscovery(input),
		"link",    	(input) => Shows.SetLink(input),
		"ep",      	(input) => Shows.SetEpisode(input),
		"finish",  	(input) => Shows._OperateConsumed(input, false),
		"dd",      	(input) => Shows.SetDownloaded(input),
		"drop",    	(input) => Shows._OperateConsumed(input, true),
		"relink",  	(input) => Shows.UpdateLink(input),
		"ev",      	(input) => Infos(Calculator(input)),
		"evp",     	(input) => ClipSend(Calculator(input)),
		'addlib'	(input) => CleanInputBox.SetInput(git_InstallAHKLibrary(input)),
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
	if runner_regex.Has(result[1])
		runner_regex[result[1]].Call(result[2])

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
}
