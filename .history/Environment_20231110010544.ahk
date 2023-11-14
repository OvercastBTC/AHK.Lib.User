#Requires AutoHotkey v2
#Include <Utils\Choose>
; ---------------------------
; #Include <RecLibs\Common_Rec_Texts>
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
; #Include <Directives\__AE.v2>

class Environment {

	static Notes => Environment._GenerateNotesMap()
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
		Notes.Set(this.kvMap(Notes_Git*))
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

	static kvMap(the_map){
		map_key:=''
		map_value:=''
		new_map := Map()
		for map_key, map_value in the_map {
			new_map.Set(map_key, map_value)
		}
		return new_map
	}
}