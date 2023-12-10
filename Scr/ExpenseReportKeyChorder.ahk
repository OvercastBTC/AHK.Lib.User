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
#Include <App\Horizon>
#Include <Abstractions\Registers>
#Include <Converters\Layouts>
; --------------------------------------------------------------------------------
#Include <Utils\GetWeather>
#Include <Tools\Info>
#include <System\UIA>
#Include <Utils\Win>
; --------------------------------------------------------------------------------
#Include <RecLibs\Common_Rec_Texts>
; --------------------------------------------------------------------------------

; #e:: {
#!e::expensereport()

expensereport() {
	expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
	; expRpt.WaitElement({Type:'Text', Name:'help'}, delay).Highlight(100).Invoke()
	; Run(A_ScriptName)
	sValidKeys := Registers.ValidRegisters "[]\{}|-=_+;:'`",<.>/?"
	try key := Registers.ValidateKey(GetInput("L1", "{Esc}").Input, sValidKeys)
	catch UnsetItemError {
		Registers.CancelAction()
		return
	}

	static actions := Map(
		
		; 'a', ,
		; 'b', ,
		; 'c', ,
		; 'd', ,
		; 'e', ,
		; 'f', ,
		; 'g', ,
		; 'h', ,
		; 'i', ,
		; 'j', ,
		; 'k', ,
		; 'l', ,
		; 'm', ,
		; 'n', ,
		; 'o', ,
		; 'p', ,
		; 'q', ,
		; 'r', ,
		; 's', ,
		; 't', ,
		; 'u', ,
		; 'v', ,
		; 'w', ,
		; 'x', ,
		; 'y', ,
		; 'z', ,
		; 'A', ,
		; 'B', ,
		; 'C', ,
		; 'D', ,
		; 'E', ,
		; 'F', ,
		; 'G', ,
		; 'H', ,
		; 'I', ,
		; 'J', ,
		; 'K', ,
		; 'L', ,
		; 'M', ,
		; 'N', ,
		; 'O', ,
		; 'P', ,
		; 'Q', ,
		; 'R', ,
		; 'S', ,
		; 'T', ,
		; 'U', ,
		; 'V', ,
		; 'W', ,
		; 'X', ,
		; 'Y', ,
		; 'Z', ,
		'A', _ShowInInfo, ;! nothing; testing func; type something in the CleanInputBox and it will be Infos
		'c', CompanyCar,
		; 'e', expensereport, ;! nothing; placeholder
		'f', AirlineFee,
		's', CarService,
		'E', EmployeeMeal,
		'm', EmployeeMeal,
		'h', Hotel,
		'o', Office_Expenses,
		'rc', RentalCar,
		'ta', TravelAgency,
		'wc', WorkClothes,
		'sav', Save,
		'fee',FeeType,
		'sfee',SeatFee,
		'carservice',Car_Service,
		'tr',TripPurpose, 
		; "a", () => Browser.RunLink(Links["ahk v2 docs"]),
		'a', hAutoComplete,
		; 'b', [func here],
		; "c", () => Infos(A_Clipboard),
		"c", () => Browser.RunLink(Links["chromeriver"]),
		; "d", () => DS4.winObj.App(),
		"d", () => Browser.RunLink(Links["ahk v2 docs"]),
		"e", () => Browser.RunLink(Links["gogoanime"]),
		"f", () => Browser.RunLink(Links["skill factory"]),
		"g", () => Browser.RunLink(Links["my github"]),
		; "h", () => Browser.RunLink(Links["phind"]),
		"h", (*) => Horizon.winObj.App(),
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
		"s", () => Steam.winObj.App(), ;! run an app via a class!!! expand on this!!!
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
	static waitelement_delay := delay := 30000
	static delay := waitelement_delay
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

	static TripPurpose() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		Sleep(100)
		expRpt.WaitElement({Type: '50003 (ComboBox)', Name: "Trip Purpose", Value: "-- Select --", LocalizedType: "combo box", AutomationId: "TripPurpose"},delay).Highlight(100).Invoke()
		Send('{Down 3}')
		Sleep(100)
		Send('{Enter}')
	}
	static Car_Service(){
		FeeType()
		Sleep(100)
		SendEvent('{Down 3}')
		Sleep(100)
		SendEvent('{Enter}')
	}
	static SeatFee(){
		FeeType()
		Sleep(100)
		SendEvent('{Down 5}')
		Sleep(100)
		SendEvent('{Enter}')
	}
	
	static FeeType(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		Sleep(100)
		expRpt.WaitElement({Type: '50003 (ComboBox)', Name: "Fee Type", Value: "-- Select --", LocalizedType: "combo box", AutomationId: "TypeAirFees"}, delay).Highlight(100).Invoke()
	}
	static Save(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		Sleep(100)
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Save", LocalizedType: "button", AutomationId: "save-btn"}, delay).Highlight(100).Invoke()
		
	}
	static WorkClothes(){
		Misc()
		Sleep(300)
		OtherMisc()
	}
	static TravelAgency() {
		travel()
		Sleep(500)
		Travel_Agency()
	}
	static Travel_Agency(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Travel Agency Transaction Fees - TMC", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	static RentalCar() {
		travel()
		Sleep(300)
		Rental_Car()
	}
	static AirlineFee(){
		travel()
		Sleep(300)
		Airline_Fee()
	}
	static Airline_Fee(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Airline Fee", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	static Rental_Car(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Car Rental", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	static Office_Expenses(){
		OfficeExpenses()
		Sleep(500)
		OfficeSupplies()
	}
	static OfficeExpenses() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Office Expenses", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
	}
	static OfficeSupplies(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Office Supplies", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	static Misc() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Misc", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
	}
	static OtherMisc(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Other / Misc", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	static Hotel() {
		travel()
		Sleep(500)
		h_hotel()
	}
	static travel() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Travel", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
	}
	static h_hotel() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Hotel", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	static CarService() {
		travel()
		Sleep(500)
		Car_Service()

		static travel() {
			expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
			expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Travel", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
		}
		static Car_Service() {
			expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
			expRpt.WaitElement({Type: '50000 (Button)', Name: "Car Service", LocalizedType: "button"}, delay).Highlight(100).Invoke()
		}
	}
	static CompanyCar(){
		Company_Car()
		Sleep(500)
		Fuel()
	}
	static Company_Car() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Company Car", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
	}
	static Fuel(){
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50000 (Button)', Name: "Fuel / Company Car", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	; static empMeal() {
	; 	; static Meal() {
	; 	; 	expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
	; 	; 	expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Meals / Entertainment", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
	; 	; }
	; }
	static EmployeeMeal(){
		Meals_Entertainment()
		Sleep(200)
		Employee_Meals()
	}
	static Meals_Entertainment() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: '50011 (MenuItem)', Name: "Meals / Entertainment", LocalizedType: "menu item"}, delay).Highlight(100).Invoke()
	}
	static Employee_Meals() {
		expRpt := UIA.ElementFromChromium('Chrome River - Google Chrome')
		expRpt.WaitElement({Type: 'Type: 50000 (Button)', Name: "Employee Meals", LocalizedType: "button"}, delay).Highlight(100).Invoke()
	}
	; --------------------------------------------------------------------------------
}

; HotIfWinActive('Chrome River - Google Chrome')
	; Hotstring("EndChars", " ")
; Hotstring(':*:meal', 'Business Travel - Meal - ')
; :*:meal::'Business Travel - Meal - '
; HotIfWinActive()
; ^+#l::HznAutoComplete()

hAutoComplete() {
	; SetCapsLockState("Off")
	acInfos := Infos('AutoComplete enabled'
					'Press "Shift+{Enter}" to activate'
				)
	; acInfos := Infos('Press "ctrl + a" to activate, or press "Shift+Enter"')
	; Hotkey(" ", (*) => createGUI())
	; Hotkey("^a", (*) => createGUI())
	Hotkey('+Enter', (*) => createGUI() )
	; createGUI()
	createGUI() {
		initQuery := "Recommendation Library"
		; initQuery := ""
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
		QSGui.Show( "AutoSize NA")
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