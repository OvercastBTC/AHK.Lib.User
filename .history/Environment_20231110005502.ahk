#Requires AutoHotkey v2
#Include <Utils\Choose>
; ---------------------------
#Include <RecLibs\Common_Rec_Texts>
#Include <Notes\Rust>
#Include <Notes\Vim>
#Include <Notes\Long>
#Include <Notes\Git>
#Include <Notes\Math>
#Include <Notes\Code>
#Include <Notes\Info>
#Include <Notes\Tech>
#Include <Notes\Terminal>
; ---------------------------
#Include <Links\Rust>
#Include <Links\Github>
#Include <Links\DiscordPins>
#Include <Links\General>
#Include <Links\Channel>
#Include <Links\Memes>
#Include <Links\Tools>
#Include <Links\Docs>
#Include <Links\AhkLib>
#Include <Links\Fonts>
; ---------------------------
#Include <Common_Rec_Texts>
#Include <Common_Abbrevations>
#Include <Common_DSPs>
#Include <Common_ExpenseReport>
#Include <Common_HumanElement>
#Include <Common_OSTitles>
#Include <Common_Personal>
; #Include <Common_Rec_Texts>
; ---------------------------
#Include <Extensions\Map>
#Include <Tools\StateBulb>
; ---------------------------
#Include <Directives\__AE.v2>

class Environment {

	static Notes => this._GenerateNotesMap()
	static Links => this._GenerateLinksMap()
	static RecLibs => this._GenerateRecLibsMap()
	static _vimMode := false
	static VimMode {
		get => this._vimMode
		set {
			this._vimMode := value
			if value
				StateBulb[1].Create()
			else {
				StateBulb[1].Destroy()
				this.WindowManagerMode := value
			}
		}
	}
	static _GenerateNotesMap() {
		Notes := Map()
		Notes.Set(Notes_Terminal*)
		Notes.Set(Notes_Code*)
		Notes.Set(Notes_Tech*)
		Notes.Set(Notes_Info*)
		Notes.Set(Notes_Long*)
		Notes.Set(Notes_Math*)
		Notes.Set(Notes_Git*)
		Notes.Set(Notes_Vim*)
		Notes.Set(Notes_Rust*)
		return Notes
	}
	static _GenerateLinksMap() {
		Links := Map()
		Links.Set(Links_General*)
		Links.Set(Links_Channel*)
		Links.Set(Links_Memes*)
		Links.Set(Links_Tools*)
		Links.Set(Links_Docs*)
		Links.Set(Links_AhkLib*)
		Links.Set(Links_DiscordPins*)
		Links.Set(Links_Github*)
		Links.Set(Links_Fonts*)
		Links.Set(Links_Rust*)
		return Links
	}
	static _GenerateRecLibsMap() {
		RecLibs := Map()
		; RecLibs.Set(understanding_the_risk*)
		; RecLibs.Set(map_understanding_the_risk*)
		; Common.Set(Common_Channel*)
		; Common.Set(Common_Memes*)
		; Common.Set(Common_Tools*)
		; Common.Set(Common_Docs*)
		; Common.Set(Common_AhkLib*)
		; Common.Set(Common_DiscordPins*)
		; Common.Set(Common_Github*)
		; Common.Set(Common_Fonts*)
		; Common.Set(Common_Rust*)
		return RecLibs
	}
	; static _GenerateCommonMap() {
		; Common := Map()
		; Common.Set(Common_Abbreviations*)
		; Common.Set(Common_Channel*)
		; Common.Set(Common_Memes*)
		; Common.Set(Common_Tools*)
		; Common.Set(Common_Docs*)
		; Common.Set(Common_AhkLib*)
		; Common.Set(Common_DiscordPins*)
		; Common.Set(Common_Github*)
		; Common.Set(Common_Fonts*)
		; Common.Set(Common_Rust*)
		; return Common
	; }
}