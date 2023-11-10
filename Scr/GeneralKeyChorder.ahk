; #Include <Notes\Vim>; #Include <Notes\Code>; #Include <Notes\Git>; #Include <Notes\Info>; #Include <Notes\Long>; #Include Notes\Math>; #Include <Notes\Rust>; #Include <Notes\Tech>; #Include <Notes\Terminal>; #Include <Notes\Vim>
#Include <Tools\CleanInputBox>
#Include <App\Autohotkey>
#Include <Tools\KeycodeGetter>
#Include <Misc\EmojiSearch>
#Include <Utils\GetInput>
#Include <App\Browser>
#Include <App\Steam>
#Include <App\DS4>
#Include <Environment>
#Include <Abstractions\Registers>
#Include <Converters\Layouts>
#Include <Links>
#Include <Utils\GetWeather>
#Include <Directives\__AE.v2>
#Include <Common_OSTitles>
#Include <RecLibs\Common_Rec_Texts>

#h:: {
	; Run(A_ScriptName)
	sValidKeys := Registers.ValidRegisters "[]\{}|-=_+;:'`",<.>/?"
	try key := Registers.ValidateKey(GetInput("L1", "{Esc}").Input, sValidKeys)
	catch UnsetItemError {
		Registers.CancelAction()
		return
	}

	; static FilesInFolder(folder?){
	; 	folder := ''
	; 	files := ''
	; 	if !folder {
	; 		folder := Paths.Lib
	; 	}
	; 	for key, value in Paths.folder {
	; 		fkey .= key . '`n'
	; 	}
	; 	Infos(fkey)
	; 	folder := Choose(fkey)
	; 	mFilesInFolder := Map()
	; 	loop files (A_LoopFilePath '\*.ahk') {
	; 		FileName:=StrSplit(A_LoopFileName, '.ahk')
	; 		mFilesInFolder.Set(A_Index, FileName)
	; 	}
	; 	for each, value in mFilesInFolder {
	; 		files .= value . '`n'
	; 	}
	; 	Infos(mFilesInFolder.Get(A_Index))
	; 	file := mFilesInFolder.Choose(FileName)
	; 	return _ViewChoice(file)
	; }
	; static _GenerateChoiceMap() {
	; 	Choice := Map()
	; 	; Choice.Set(Choice_General*)

	; 	return Choice
	; }
	; static _ViewChoice(input?) {
	; 	; if !input := CleanInputBox().WaitForInput()
	; 	; 	return
	; 	; choice := Environment.Notes.Choose(input)
	; 	choice := FilesInFolder(input)
	; 	if !choice
	; 		return
	; 	A_Clipboard := choice
	; 	Infos(choice)
	; }
	static _ViewNote(input?) {
		if !input := CleanInputBox().WaitForInput()
			return
		note := Environment.Notes.Choose(input)
		if !note
			return
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
	static _ViewRecs(input?) {
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

	static actions := Map(

		"a", () => Browser.RunLink(Links["ahk v2 docs"]),
		"^a", () => Browser.RunLink(Links["ahk v2 docs"]),
		
		"c", () => Infos(A_Clipboard),
		"d", () => DS4.winObj.App(),
		"e", () => Browser.RunLink(Links["gogoanime"]),
		"f", () => Browser.RunLink(Links["skill factory"]),
		"g", () => Browser.RunLink(Links["my github"]),
		"h", () => Browser.RunLink(Links["phind"]),
		'i', _ViewLinks,
		"j", () => EmojiSearch(CleanInputBox().WaitForInput()),
		"k", KeyCodeGetter,
		'l', HznAutoComplete(),
		"m", () => Browser.RunLink(Links["gmail"]),
		"n", () => Browser.RunLink(Links["monkeytype"]),
		'o', _ViewNote,


		"r", () => Browser.RunLink(Links["reddit"]),
		"s", () => Steam.winObj.App(),
		"t", () => Browser.RunLink(Links["mastodon"]),
		"T", () => Browser.RunLink(Links["twitch"]),
		"u", () => Infos(GetWeather()),
		"x", () => Browser.RunLink(Links["regex"]),
		"v", () => Browser.RunLink(Links["vk"]),
		"w", () => Browser.RunLink(Links["wildberries"]),
		; 'l', _ViewRecs,
		; "i", _ShowInInfo,
		; "i", _ViewChoice('Notes'),

	)
	if key
		try actions[key].Call()
	
	
}
^+#l::HznAutoComplete()

HznAutoComplete(*) {
	
	SetCapsLockState( "Off")
	acInfos := Infos('Press "ctrl + a" to activate, or press "Shift+Enter"')
	Hotkey("^a", (*) => createGUI())
	Hotkey('+Enter', (*) => createGUI() )
	
	createGUI() {
		initQuery := "Recommendation Library"
		initQuery := ""
		; global entriesList := ["Red", "Green", "Blue"]
		mList := []
		mlist := RecLibs.understanding_the_risk
		; Infos(mlist)
		; entriesList := [mlist]
		; entries := []
		entriesList := []
		for each, value in mList {
			entriesList.Push(value)
			; entries.Push(value)
		}
		; for each, value in entries {
		for each, value in entriesList {
			; entriesList := ''
			entries := ''
			; entriesList .= value '`n'
			entries .= value '`n'
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
	for index, value in entriesList
	; for index, value in entries
	{
		; Check if the current value matches the target value
		if (value = currContent)
		{
			valueFound := true
			break ; Exit the loop if the value is found
		}
	}
	if (valueFound)
		return ; Exit Nested request

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