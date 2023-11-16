;No dependencies

class Paths {

	static User         := "C:\Users\" A_UserName
	static AppData 		:= Paths.User "\AppData"
	static LocalAppData := Paths.AppData "\Local"
	static AppDataProgs := Paths.LocalAppData "\Programs"
	static RoamingAppData := Paths.AppData "\Roaming"
	static System32     := "C:\Windows\System32"
	static OneDrive 	:= Paths.User . '\OneDrive - FM Global'
	static Startup 		:= Paths.RoamingAppData . '\Microsoft\Windows\Start Menu\Programs\Startup'

	static StandardAhkLibLocation := A_MyDocuments '\AutoHotkey\Lib'
	static Prog  := Paths.AppDataProgs '\AutoHotkey\v2'
	static v2Prog  := Paths.AppDataProgs '\AutoHotkey\v2'
	static v1Prog  := Paths.AppDataProgs '\AutoHotkey'
	static v2Proj  := Paths.v2Prog '\AHK.Projects.v2'
	static v2Lib := Paths.v2Prog '\Lib'

	static lnchr := Paths.v2Proj '\LNCHR'

	static Lib   := Paths.StandardAhkLibLocation

	static Main  := Paths.OneDrive . 'AHK.Main'
	static Music := Paths.OneDrive . '\Music'
	static Shows := Paths.Lib 'App\Shows.ahk'
	static Info  := Paths.Lib '\Tools\Info.ahk'
	static Test  := "C:\Programming\test"
	static Reg   := Paths.Lib '\Abstractions\Registers.ahk'
	
	static folder := Map(
		'Notes',  (Paths.Lib '\Notes'),
		'Links' , (Paths.Lib '\Links'),
		'RecLibs',(Paths.Lib '\RecLibs'),
		'Common', (Paths.v2Lib 'Common_'),
	)

	static Files     := Paths.Main "\Files"
	static Tools     := Paths.Main "\Tools"

	static Pictures     := Paths.OneDrive "\Pictures"
	static Content      := Paths.Pictures "\Content"
	static OnePiece     := Paths.Pictures "\Content\One piece"
	static VideoTools   := Paths.Pictures "\Tools"
	static ScreenVideos := Paths.Pictures "\Screenvideos"
	static Screenshots  := Paths.Pictures '\Screenshots'
	static Downloads 	:= Paths.OneDrive '\Downloads'
	; static Downloads 	:= this.OneDrive 'Downloads'
	; Paths.OneDrive '\Downloads'
	static Tree         := Paths.Pictures "\Tree"
	static Memes        := Paths.Pictures "\Tree\Memes"
	static Emoji        := Paths.Pictures "\Tree\Emojis"
	static Other        := Paths.Pictures "\Tree\Other"
	static Logos        := Paths.Pictures "\Tree\Logos"
	static Themes       := Paths.Pictures "\Tree\Themes"

	static Audio := Paths.User . '\Music'
	static Sounds := Paths.Audio "\Sounds"

	static VsCodeExtensions := "C:\Users\" A_UserName "\.vscode-insiders\extensions"
	static SavedScreenshots := Paths.LocalAppData "\Packages\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\TempState\ScreenClip"

	static Ptf := Map(
		"playlist-sorter", Paths.Files "\Innit\playlist-sorter.txt",
		"test-state",      Paths.Files "\Innit\test-state.txt",
		"time-agent",      Paths.Files "\Innit\time-agent.txt",

		"BlankPic", Paths.Files "\img\BlankPic.png",

		"Hub", Paths.Main "\Hub.ahk",

		"Tests", Paths.Tools "\Tests.ahk",
		"Timer", Paths.Tools "\Timer.ahk",

		"AhkTest", Paths.Test "\AhkTest.ahk",

		"vine boom",         Paths.Sounds "\vine boom.wav",
		"faded than a hoe",  Paths.Sounds "\faded than a hoe.wav",
		"heheheha",          Paths.Sounds "\heheheha.wav",
		"shall we",          Paths.Sounds "\shall we.wav",
		"slip and crash",    Paths.Sounds "\slip and crash.wav",
		"cartoon running",   Paths.Sounds "\cartoon running.wav",
		"rizz",              Paths.Sounds "\rizz.wav",
		"bruh sound effect", Paths.Sounds "\bruh sound effect.wav",
		"cartoon",           Paths.Sounds "\cartoon.wav",
		"hohoho",            Paths.Sounds "\hohoho.wav",
		"bing chilling 1",   Paths.Sounds "\bing chilling 1.wav",
		"bing chilling 2",   Paths.Sounds "\bing chilling 2.wav",
		"oh fr on god",      Paths.Sounds "\oh fr on god.wav",
		"sus",               Paths.Sounds "\sus.wav",
		"i just farted",     Paths.Sounds "\i just farted.wav",
		"ting",              Paths.Sounds "\ting.wav",
		"shutter",           Paths.Sounds "\shutter.wav",
		"was that his cock", Paths.Sounds "\was that his cock.wav",
		"cyberpunk",         Paths.Sounds "\cyberpunk.wav",
		"better call saul",  Paths.Sounds "\better call saul short.wav",

		"Discovery log", Paths.Music "\Discovery log.txt",
		"Unfinished",    Paths.Music "\Unfinished.txt",
		"Rappers",       Paths.Music "\Rappers.txt",
		"Artists",       Paths.Music "\Favorites.md",

		"Shows",    Paths.Shows "\Shows.jsonc",
		"Consumed", Paths.Shows "\Consumed.md",

		"Diary",     Paths.Info "\diary.md",
		"Events",    Paths.Info "\events.jsonc",
		"Birthdays", Paths.Info "\birthdays.jsonc",

		"femboy",       Paths.Memes "\femboy.png",
		"writing fire", Paths.Memes "\writing fire.jpg",
		"urethra",      Paths.Memes "\urethra.jpg",
		"welp",         Paths.Memes "\welp.jpg",
		"how did we get here", Paths.Memes "\how did we get here.jpeg",
		"do you have the slightest idea how little that narrows it down", Paths.Memes "\do you have the slightest idea how little that narrows it down.png",

	)

	static Apps := Map(

		"Sound mixer",       "SndVol.exe",
		"Slide to shutdown", "SlideToShutDown.exe",

	)
}
