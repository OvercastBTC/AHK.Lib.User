; ---------------------------
#Include <Extensions\Map>
#Include <Tools\StateBulb>
#Include <Utils\Choose>
#Include <Tools\Info>
; ---------------------------
#Include <Includes\Notes>
; ---------------------------
#Include <Includes\Links>
; ---------------------------
#Include <RecLibs\Common_Rec_Texts>
; ---------------------------
; #Include <Common_Rec_Texts>
; #Include <Common_Abbrevations>
; #Include <Common_DSPs>
; #Include <Common_ExpenseReport>
; #Include <Common_HumanElement>
; #Include <Common_OSTitles>
; #Include <Common_Personal>
; #Include <Common_Rec_Texts>
; ---------------------------


class Environment {

	static Notes => this._GenerateNotesMap()
	static Links => this._GenerateLinksMap()
	static RecLibs => this._GenerateRecLibsMap()
	static _vimMode := false
	static VimMode {
		get => this._vimMode
		set {
			this._vimMode := value
			; if value{
			; 	StateBulb[1].Create()
			; }
			; else {
			; 	StateBulb[1].Destroy()
			; 	this.WindowManagerMode := value
			; }
		}
	}
	static _GenerateNotesMap() {
		Notes := Map()
		Notes.Set(Notes_Terminal*),
		Notes.Set(Notes_Code*),
		Notes.Set(Notes_Tech*),
		Notes.Set(Notes_Info*),
		Notes.Set(Notes_Long*),
		Notes.Set(Notes_Math*),
		Notes.Set(Notes_Git*),
		Notes.Set(Notes_Vim*),
		Notes.Set(Notes_Rust*)
		return Notes
	}
	static _GenerateLinksMap() {
		Links := Map()
		Links.Set(Links*),
		Links.Set(Links_AhkLib*),
		Links.Set(Links_Channel*),
		Links.Set(Links_Docs*),
		Links.Set(Links_DiscordPins*),
		Links.Set(Links_Fonts*),
		Links.Set(Links_General*),
		Links.Set(Links_Github*),
		Links.Set(Links_Memes*),
		Links.Set(Links_Rust*),
		Links.Set(Links_Tools*)
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