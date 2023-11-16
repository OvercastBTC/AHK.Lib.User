; --------------------------------------------------------------------------------
#Requires AutoHotkey v2
#Include <Directives\__AE.v2>
; --------------------------------------------------------------------------------
#Include <Environment>
#Include <Includes\Links>
; --------------------------------------------------------------------------------
#Include <Tools\CleanInputBox>
#Include <App\Autohotkey>
#Include <Tools\KeycodeGetter>
#Include <Misc\EmojiSearch>
#Include <Utils\GetInput>
#Include <App\Browser>
#Include <App\Steam>
#Include <App\DS4>
#Include <Abstractions\Registers>
#Include <Converters\Layouts>
; --------------------------------------------------------------------------------
#Include <Utils\GetWeather>
#Include <Tools\Info>
; --------------------------------------------------------------------------------
#HotIf WinActive('Runner.ahk')
	*^s::Run('Runner.ahk')
#HotIf
; --------------------------------------------------------------------------------
#Include <RecLibs\Common_Rec_Texts>

#h:: {
	; Run(A_ScriptName)
	sValidKeys := Registers.ValidRegisters "[]\{}|-=_+;:'`",<.>/?"
	try key := Registers.ValidateKey(GetInput("L1", "{Esc}").Input, sValidKeys)
	catch UnsetItemError {
		Registers.CancelAction()
		return
	}

	static actions := Map(

		; "a", () => Browser.RunLink(Links["ahk v2 docs"]),
		'a', HznAutoComplete,
		; 'b', [func here],
		"c", () => Infos(A_Clipboard),
		; "d", () => DS4.winObj.App(),
		"d", () => Browser.RunLink(Links["ahk v2 docs"]),
		"e", () => Browser.RunLink(Links["gogoanime"]),
		"f", () => Browser.RunLink(Links["skill factory"]),
		"g", () => Browser.RunLink(Links["my github"]),
		"h", () => Browser.RunLink(Links["phind"]),
		"i", _ShowInInfo,
		"j", () => EmojiSearch(CleanInputBox().WaitForInput()),
		"k", KeyCodeGetter,
		; 'l', HznAutoComplete,
		'l', _ViewLinks(),
		; 'l', _ViewRecs,
		"m", () => Browser.RunLink(Links["gmail"]),
		'n', _ViewNote,
		"o", () => Browser.RunLink(Links["monkeytype"]),
		'p', () => Infos(A_MyDocuments),
		;q
		"r", () => Browser.RunLink(Links["reddit"]),
		"s", () => Steam.winObj.App(),
		"t", () => Browser.RunLink(Links["mastodon"]),
		"T", () => Browser.RunLink(Links["twitch"]),
		"u", () => Infos(GetWeather()),
		"v", () => Browser.RunLink(Links["vk"]),
		"w", () => Browser.RunLink(Links["wildberries"]),
		"x", () => Browser.RunLink(Links["regex"]),
		; 'y', ,
		; 'z', ,

	)
	; --------------------------------------------------------------------------------
	if (key) {
		try {
			actions[key].Call()
		}
	}
	; --------------------------------------------------------------------------------
	static _ViewNote(input?) {
		if !input := CleanInputBox().WaitForInput() {
			return
		}
		note := Environment.Notes.Choose(input)
		if !note{
			return
		}
		A_Clipboard := note
		Infos(note)
	}
	static _ViewLinks(input?) {
		if !input := CleanInputBox().WaitForInput()
			return
		link := Environment.Links.Choose(input)
		if !link
			return
		A_Clipboard := link
		Infos(link)
	}
	static _OpenLinks(input?) {
		if !input := CleanInputBox().WaitForInput()
			return
		link := Environment.Links.Choose(input)
		if (!link) {
			return
		}
		A_Clipboard := link
		Infos(link)
		Browser.RunLink(link)
	}

	static _ViewRecs() {
		if !input := CleanInputBox().WaitForInput()
			return
		RecLibs := Environment.RecLibs.Choose(input)
		if !RecLibs
			return
		A_Clipboard := RecLibs
		Infos(RecLibs)
	}

	static _ShowInInfo() {
		if !input := CleanInputBox().WaitForInput()
			return
		Infos(input)
	}
}
; ^+#l::HznAutoComplete()

HznAutoComplete() {
	; SetCapsLockState("Off")
	acInfos := Infos('AutoComplete enabled'
					'Press "Shift+{Enter}",to activate'
				)
	; acInfos := Infos('Press "ctrl + a" to activate, or press "Shift+Enter"')
	; Hotkey(" ", (*) => createGUI())
	; Hotkey("^a", (*) => createGUI())
	Hotkey('+Enter', (*) => createGUI() )
	; createGUI()
	createGUI() {
		initQuery := "Recommendation Library"
		initQuery := ""
		; global entriesList := ["Red", "Green", "Blue"]
		mList := []
		mlist := RecLibs.understanding_the_risk
		; mlist := [RecLibs.understanding_the_risk, RecLibs.HumanElement.electrical]
		; Infos(mlist)
		; entriesList := [mlist]
		; entries := []
		entries := ''
		entriesList := []
		m:=''
		for each, m in mList {
			entriesList.Push(m)
		}
		e:=''
		for each, e in entriesList {
			; entriesList := ''
			; entries := ''
			; entriesList .= value '`n'
			entries .= e '`n'
		}
		
		global QSGui, initQuery, entriesList
		global width := Round(A_ScreenWidth / 4)
		QSGui := Gui("AlwaysOnTop +Resize +ToolWindow Caption", "Recommendation Picker")
		QSGui.SetColor := 0x161821
		QSGui.BackColor := 0x161821
		QSGui.SetFont( "s10 q5", "Fira Code")
		; QSCB := QSGui.AddComboBox("vQSEdit w200", entriesList)
		QSCB := QSGui.AddComboBox("vQSEdit w" width ' h200' ' Wrap', entriesList)
		qEdit := QSGui.AddEdit('vqEdit w' width ' h200')
		; qEdit.OnEvent('Change', (*) => updateEdit(QSCB, entriesList))
		; QSGui_Change(QSCB) => qEdit.OnEvent('Change',qEdit)
		QSGui.Add('Text','Section')
		QSGui.Opt('+Owner ' QSGui.Hwnd)
		; QSCB := QSGui.AddComboBox("vQSEdit w" width ' h200', entriesList)
		QSCB.Text := initQuery
		QSCB.OnEvent("Change", (*) => AutoComplete(QSCB, entriesList))
		; QSCB.OnEvent('Change', (*) => updateEdit(QSCB, entriesList))
		QSBtn := QSGui.AddButton("default hidden yp hp w0", "OK")
		QSBtn.OnEvent("Click", (*) => processInput())
		QSGui.OnEvent("Close", (*) => QSGui.Destroy())
		QSGui.OnEvent("Escape", (*) => QSGui.Destroy())
		; QSGui.Show( "w222")
		; QSGui.Show("w" width ' h200')
		QSGui.Show( "AutoSize")
	}

	processInput() {
		QSSubmit := QSGui.Submit()    ; Save the contents of named controls into an object.
		if QSSubmit.QSEdit {
			; MsgBox(QSSubmit.QSEdit, "Debug GUI")
			initQuery := QSSubmit.QSEdit
			Infos.foDestroyAll()
			Sleep(100)
			updated_Infos := Infos(QSSubmit.QSEdit)
			
		}
		QSGui.Destroy()
		WinWaitClose(updated_Infos.hwnd)
		Run(A_ScriptName)
	}

	; https://github.com/Pulover/CbAutoComplete
	; Rewritten for AutoHotkey v2

	AutoComplete(ComboBox, entriesList) {
	; CB_GETEDITSEL = 0x0140, CB_SETEDITSEL = 0x0142
		currContent := ComboBox.Text
		QSGui['qEdit'].Value := currContent
		; QSGui['your name'].Value := currContent
		; QSGui.Add('Text','Section','Text')
		QSGui.Show("AutoSize")
		; QSGui.Show()
		if ((GetKeyState("Delete", "P")) || (GetKeyState("Backspace", "P"))){
			return
		}

		valueFound := false
		for index, value in entriesList {
		; for index, value in entries
			; Check if the current value matches the target value
			if (value = currContent)
			{
				valueFound := true
				break ; Exit the loop if the value is found
			}
		}
		if (valueFound){
			return ; Exit Nested request
		}
		Start :=0, End :=0
		MakeShort(0, &Start, &End)
		try {
			if (ControlChooseString(ComboBox.Text, ComboBox) > 0) {
				Start := StrLen(currContent)
				End := StrLen(ComboBox.Text)
				PostMessage( 0x0142, 0, MakeLong(Start, End), , "ahk_id" ComboBox.Hwnd)
			}
		} Catch as e {
			ControlSetText( currContent, ComboBox)
			PostMessage( 0x0142, 0, MakeLong(StrLen(currContent), StrLen(currContent)), , "ahk_id " ComboBox.Hwnd)
		}

		MakeShort(Long, &LoWord, &HiWord) {
			LoWord := Long & 0xffff
			, HiWord := Long >> 16
		}

		MakeLong(LoWord, HiWord) {
			return (HiWord << 16) | (LoWord & 0xffff)
		}
	}
}