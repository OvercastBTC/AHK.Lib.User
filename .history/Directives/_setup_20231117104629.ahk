#Include <Extensions\Array>
#Include <Extensions\Gui>
#Include <Extensions\String>
#Include <Directives\__AE.v2>
#include <Tools\Info>
#Include <Directives\_user_info>
; ^+a::setups.initial_setups_gui()
^+a::setups()
class Setups {

	/**
	 * To use Setup, you just need to create an instance of it, no need to call any method after
	 * @param text *String*
	 * @param autoCloseTimeout *Integer* in milliseconds. Doesn't close automatically
	 */
	; __New(text, autoCloseTimeout := 0) {		
	__New(text:='', autoCloseTimeout := 0) {
		this.autoCloseTimeout := autoCloseTimeout
		this.text := text
		; this._CreateGui()
		Setups.initial_setups_gui()
		this.hwnd := Setups.gSetup.hwnd
		if !this._GetAvailableSpace() { 
			this._StopDueToNoSpace()
			return
		}
		this._SetupHotkeysAndEvents()
		this._SetupAutoclose()
		this._Show()
	}

	static fontSize     := 10
	static distance     := 2
	; static unit         := A_ScreenDPI / 96
	static unit         := A_ScreenDPI / 144
	static guiWidth     := Setups.fontSize * Setups.unit * Setups.distance
	; static guiWidth     := 30
	static maximumSetup := Floor(A_ScreenHeight / Setups.guiWidth)
	static spots        := Setups._GeneratePlacesArray()
	static foDestroyAll := (*) => Setups.DestroyAll()
	static maxNumberedHotkeys := 12
	; static maxWidthInChars := 110
	static maxWidthInChars := A_ScreenWidth
	; --------------------------------------------------------------------------------
	static initial_setup := false

	static aOffice := ['Seattle', 'Walnut Creek']
	
	; static _map_to_array() {
	; 	k := ''
	; 	v := ''
	; 	static ak := []
	; 	static av := []
	; 	at := ''
	; 	for k, v in user_info.info_map {
	; 		ak.Push(k)
	; 		av.Push(v)
	; 		at .= 'key: ' k ' value: ' v '`n'
	; 		info('key: ' k ' | value: ' v)
	; 	}
	; }
	
	static initial_setups_gui(*) {
		; static _map_to_array() {
		; 	k := ''
		; 	v := ''
		; 	static ak := []
		; 	static av := []
		; 	at := ''
		; 	for k, v in user_info.info_map {
		; 		; ak.Push(k)
		; 		; av.Push(v)
		; 		at .= 'key: ' k ' value: ' v '`n'
		; 		; info('key: ' k ' | value: ' v)
		; 	}
		; }
		static ak_map_to_array() {
			k := ''
			v := ''
			ak := []
			at := ''
			for k, v in user_info.info_map {
				ak.Push(k)
				; at .= 'key: ' A_Index ' value: ' v '`n'
				; info('key: ' k ' | value: ' v)
				; info(at)
			}
			return ak
		}
		; static av_map_to_array() {
		; 	k := ''
		; 	v := ''
		; 	av := []
		; 	at := ''
		; 	for k, v in user_info.info_map {
		; 		av.Push(v)
		; 		at .= 'key: ' A_Index ' value: ' v '`n'
		; 		; info('key: ' k ' | value: ' v)
		; 	}
		; 	return av
		; }
		
		this.gSetup := Gui("AlwaysOnTop +Caption +ToolWindow", 'FM Global New User Setup').DarkMode().MakeFontNicer(Setups.fontSize).NeverFocusWindow()
		; this.gSetup.SetFont( "s10 q5", "Fira Code") ;? added
		; this.gcText := Setups.gSetup.AddText(, Setups._FormatText())
		width := (this.maxWidthInChars/7)
		kk := ''
		kv := ''
		vk := ''
		vv := ''
		office_info := []
		ak_array := []
		av_array := []
		val_array := []
		ak_array := ak_map_to_array()
		; av_array := av_map_to_array()
		; for kk, kv in this._map_to_array.ak {
		for kv in ak_array {
			index := 1
			margin := 10
			; if (kv ~= 'i)office.*') {
			; 	office_info.Push(kv)
			; }
			;? add a text box with the map keys ; x+margin needed for readability
			this.gSetup.AddText('x' margin,kv ': `t`t')
			; this.gSetup.AddText(,'`t`t')
			;? remove any spaces in the map key so it can be used as a variabale
			val := RegExReplace(kv, A_Space, '_')
			;? pushes all vVals into an array so Gui.Submit() saves the value
			val_array.Push(val) ; fix => finish setting this up (store values, then update _user_info.ahk)
			this.gSetup.AddEdit('v' val ' x+1 w' width ' Section')
		}
		this.gSetup.AddText('x' margin,'Choose Office: `t`t')
		this.gSetup.AddComboBox('x+1 w' width,this.aOffice)
		; fix add the rest of the office_address_map (with dependancy on office selected)
		; fix organize???
		this.gSetup.Show()
	}

	static DestroyAll() {
		for index, SetupObj in Setups.spots {
			if !SetupObj
				continue
			SetupObj.Destroy()
		}
	}


	static _GeneratePlacesArray() {
		availablePlaces := []
		loop Setups.maximumSetup {
			availablePlaces.Push(false)
		}
		return availablePlaces
	}


	autoCloseTimeout := 0
	bfDestroy := this.Destroy.Bind(this)


	/**
	 * Will replace the text in the Setup
	 * If the window is destoyed, just creates a new Setup Otherwise:
	 * If the text is the same length, will just replace the text without recreating the gui.
	 * If the text is of different length, will recreate the gui in the same place
	 * (once again, only if the window is not destroyed)
	 * @param newText *String*
	 * @returns {Setup} the class object
	 */
	ReplaceText(newText) {

		try WinExist(Setups.gSetup)
		catch
			return Setup(newText, this.autoCloseTimeout)

		if StrLen(newText) = StrLen(this.gcText.Text) {
			this.gcText.Text := newText
			this._SetupAutoclose()
			return this
		}

		Setups.spots[this.spaceIndex] := false
		return Setup(newText, this.autoCloseTimeout)
	}

	Destroy(*) {
		try HotIfWinExist("ahk_id " this.gSetup.Hwnd)
		catch Any {
			return false
		}
		Hotkey("Escape", "Off")
		Hotkey("^Escape", "Off")
		if this.spaceIndex <= Setups.maxNumberedHotkeys
			Hotkey("F" this.spaceIndex, "Off")
		this.gSetup.Destroy()
		Setups.spots[this.spaceIndex] := false
		return true
	}


	_CreateGui() {
		this.gSetup  := Gui("AlwaysOnTop -Caption +ToolWindow").DarkMode().MakeFontNicer(Setups.fontSize).NeverFocusWindow()
		this.gcText := this.gSetup.AddText(, this._FormatText())
	}

	static _FormatText() {
		text := String(this.text)
		lines := text.Split("`n")
		if lines.Length > 1 {
			text := this._FormatByLine(lines)
		}
		else {
			text := this._LimitWidth(text)
		}
		return text.Replace("&", "&&")
	}

	_FormatByLine(lines) {
		newLines := []
		for index, line in lines {
			newLines.Push(this._LimitWidth(line))
		}
		text := ""
		for index, line in newLines {
			if index = newLines.Length {
				text .= line
				break
			}
			text .= line "`n"
		}
		return text
	}

	_LimitWidth(text) {
		if StrLen(text) < Setups.maxWidthInChars {
			return text
		}
		insertions := 0
		while (insertions + 1) * Setups.maxWidthInChars + insertions < StrLen(text) {
			insertions++
			text := text.Insert("`n", insertions * Setups.maxWidthInChars + insertions)
		}
		return text
	}

	_GetAvailableSpace() {
		spaceIndex := unset
		for index, isOccupied in Setups.spots {
			if isOccupied
				continue
			spaceIndex := index
			Setups.spots[spaceIndex] := this
			break
		}
		if !IsSet(spaceIndex)
			return false
		this.spaceIndex := spaceIndex
		return true
	}

	_CalculateYCoord() => Round(this.spaceIndex * Setups.guiWidth - Setups.guiWidth)

	_StopDueToNoSpace() => this.gSetup.Destroy()

	_SetupHotkeysAndEvents() {
		HotIfWinExist("ahk_id " Setups.gSetup.Hwnd)
		Hotkey("Escape", this.bfDestroy, "On")
		Hotkey("^Escape", Setups.foDestroyAll, "On")
		if this.spaceIndex <= Setups.maxNumberedHotkeys
			Hotkey("F" this.spaceIndex, this.bfDestroy, "On")
		; this.gcText.OnEvent("Click", this.bfDestroy)
		Setups.gSetup.OnEvent("Close", this.bfDestroy)
	}

	_SetupAutoclose() {
		if this.autoCloseTimeout {
			SetTimer(this.bfDestroy, -this.autoCloseTimeout)
		}
	}

	_Show() => Setups.gSetup.Show("AutoSize NA x0 y" this._CalculateYCoord())

}

; Setup(text, timeout?) => Setup(text, timeout ?? 2000)
Setup(text, timeout?) => Setups(text, timeout ?? 0)