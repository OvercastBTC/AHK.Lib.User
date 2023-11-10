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
#Include <Notes\Vim>
#Include <Notes\Code>
#Include <Notes\Git>
#Include <Notes\Info>
#Include <Notes\Long>
#Include <Notes\Math>
#Include <Notes\Rust>
#Include <Notes\Tech>
#Include <Notes\Terminal>
#Include <Notes\Vim>
#Include <Utils\GetWeather>

#h:: {
	sValidKeys := Registers.ValidRegisters "[]\{}|-=_+;:'`",<.>/?"
	try key := Registers.ValidateKey(GetInput("L1", "{Esc}").Input, sValidKeys)
	catch UnsetItemError {
		Registers.CancelAction()
		return
	}

	static _ViewNote() {
		if !input := CleanInputBox().WaitForInput()
			return
		; note := Environment.Notes.Choose(input)
		note := Environment.Notes.Choose(input)
		if !note
			return
		A_Clipboard := note
		Infos(note)
	}

	static _ShowInInfo() {
		if !input := CleanInputBox().WaitForInput()
			return
		Infos(input)
	}

	static actions := Map(

		"m", () => Browser.RunLink(Links["gmail"]),
		"n", () => Browser.RunLink(Links["monkeytype"]),
		"g", () => Browser.RunLink(Links["my github"]),
		"f", () => Browser.RunLink(Links["skill factory"]),
		"x", () => Browser.RunLink(Links["regex"]),
		"w", () => Browser.RunLink(Links["wildberries"]),
		"d", () => DS4.winObj.App(),
		"s", () => Steam.winObj.App(),
		"a", () => Browser.RunLink(Links["ahk v2 docs"]),
		"r", () => Browser.RunLink(Links["reddit"]),
		"T", () => Browser.RunLink(Links["twitch"]),
		"h", () => Browser.RunLink(Links["phind"]),
		"j", () => EmojiSearch(CleanInputBox().WaitForInput()),
		"c", () => Infos(A_Clipboard),
		"k", KeyCodeGetter,
		'o', _ViewNote,
		"i", _ShowInInfo,
		"v", () => Browser.RunLink(Links["vk"]),
		"t", () => Browser.RunLink(Links["mastodon"]),
		"e", () => Browser.RunLink(Links["gogoanime"]),
		"u", () => Infos(GetWeather()),

	)
	if key
		try actions[key].Call()
}