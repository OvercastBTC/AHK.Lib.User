;No dependencies

class Paths {

	static User				:= "C:\Users\" A_UserName
	static OneDrive			:= Paths.User "\OneDrive - FM Global"
	static LocalAppData		:= Paths.User "\AppData\Local"
	static AHKProg			:= A_ProgramFiles "\AutoHotkey"
	static System32			:= "C:\Windows\System32"
	static Startup			:= A_Startup
	static Script			:= A_ScriptName
	static ScriptDir		:= A_ScriptDir
	static ShortCut			:= ".lnk"
	static Lib				:= "\Lib"
	static MyDocs			:= A_MyDocuments
	static AHK				:= "\AutoHotkey"
	static AHKMain			:= "\AHK.Main"
	static v1				:= ".v1"
	static v2				:= ".v2"
	static Proj				:= "\AHK.Projects"

	static LibLocal			:= Paths.LocalAppData.AHK.Lib
	static LibMain			:= Paths.MyDocs.AHK.Lib
	static LibMainv1		:= Paths.LibMain.Lib.v1
	static LibMainv2		:= Paths.LibMain.Lib.v2
	static LibScript		:= Paths.ScriptDir.Lib
	static LibProjv1		:= Paths.ProjMainv1.Lib
	static LibProjv2		:= Paths.ProjMainv2.Lib

	static ProjMain  		:= Paths.OneDrive.AHKMain
	static ProjMainv1  		:= Paths.ProjMain.Proj.v1
	static ProjMainv2  		:= Paths.ProjMain.Proj.v2
	; static Music 			:= "C:\Programming\music"
	; static Shows 			:= "C:\Programming\shows"
	; static Info  			:= "C:\Programming\info"
	; static Test  			:= "C:\Programming\test"
	; static Reg   			:= "C:\Programming\registers"

	; static Files    		:= Paths.Main "\Files"
	static Tools     		:= Paths.LibMain "\Tools"

	static Pictures     	:= "C:\Pictures"
	static Content      	:= Paths.Pictures "\Content"
	static OnePiece     	:= Paths.Pictures "\Content\One piece"
	static VideoTools   	:= Paths.Pictures "\Tools"
	static ScreenVideos 	:= Paths.Pictures "\Screenvideos"
	static Downloaded   	:= Paths.Pictures "\Downloaded"
	static Tree         	:= Paths.Pictures "\Tree"
	static Memes        	:= Paths.Pictures "\Tree\Memes"
	static Emoji        	:= Paths.Pictures "\Tree\Emojis"
	static Other        	:= Paths.Pictures "\Tree\Other"
	static Logos        	:= Paths.Pictures "\Tree\Logos"
	static Themes       	:= Paths.Pictures "\Tree\Themes"

	static Audio 			:= "C:\Audio"
	static Sounds 			:= Paths.Audio "\Sounds"

	static VsCodeExtensions := Paths.User "\.vscode-insiders\extensions"
	static SavedScreenshots := Paths.LocalAppData "\Packages\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\TempState\ScreenClip"

	static Ptf := Map(
		; "playlist-sorter", Paths.Files "\Innit\playlist-sorter.txt",
		; "test-state",      Paths.Files "\Innit\test-state.txt",
		; "time-agent",      Paths.Files "\Innit\time-agent.txt",

		; "BlankPic", Paths.Files "\img\BlankPic.png",

		; "Hub",	Paths.Main "\Hub.ahk",

		"Tests",				Paths.Tools "\Tests.ahk",
		"Timer",				Paths.Tools "\Timer.ahk",

		"AhkTest",				Paths.Test "\AhkTest.ahk",

		"vine boom",        	Paths.Sounds "\vine boom.wav",
		"faded than a hoe", 	Paths.Sounds "\faded than a hoe.wav",
		"heheheha",         	Paths.Sounds "\heheheha.wav",
		"shall we",         	Paths.Sounds "\shall we.wav",
		"slip and crash",   	Paths.Sounds "\slip and crash.wav",
		"cartoon running",  	Paths.Sounds "\cartoon running.wav",
		"rizz",             	Paths.Sounds "\rizz.wav",
		"bruh sound effect", 	Paths.Sounds "\bruh sound effect.wav",
		"cartoon",           	Paths.Sounds "\cartoon.wav",
		"hohoho",            	Paths.Sounds "\hohoho.wav",
		"bing chilling 1",   	Paths.Sounds "\bing chilling 1.wav",
		"bing chilling 2",   	Paths.Sounds "\bing chilling 2.wav",
		"oh fr on god",      	Paths.Sounds "\oh fr on god.wav",
		"sus",               	Paths.Sounds "\sus.wav",
		"i just farted",     	Paths.Sounds "\i just farted.wav",
		"ting",              	Paths.Sounds "\ting.wav",
		"shutter",           	Paths.Sounds "\shutter.wav",
		"was that his cock", 	Paths.Sounds "\was that his cock.wav",
		"cyberpunk",         	Paths.Sounds "\cyberpunk.wav",
		"better call saul",  	Paths.Sounds "\better call saul short.wav",

		"Discovery log", 		Paths.Music "\Discovery log.txt",
		"Unfinished",    		Paths.Music "\Unfinished.txt",
		"Rappers",       		Paths.Music "\Rappers.txt",
		"Artists",       		Paths.Music "\Favorites.md",

		"Shows",    			Paths.Shows "\Shows.jsonc",
		"Consumed", 			Paths.Shows "\Consumed.md",

		"Diary",     			Paths.Info "\diary.md",
		"Events",    			Paths.Info "\events.jsonc",
		"Birthdays", 			Paths.Info "\birthdays.jsonc",

		"femboy",       		Paths.Memes "\femboy.png",
		"writing fire", 		Paths.Memes "\writing fire.jpg",
		"urethra",      		Paths.Memes "\urethra.jpg",
		"welp",         		Paths.Memes "\welp.jpg",
		"how did we get here", 	Paths.Memes "\how did we get here.jpeg",
		"do you have the slightest idea how little that narrows it down", Paths.Memes "\do you have the slightest idea how little that narrows it down.png",

	)

	static Apps := Map(

		"AHK Script",       		"AHK Script.v2.ahk",
		"Hzn+",						"HznPlus.ahk",
		"Detect_ActiveProcess v2", 	"A_Process.v2.ahk",

	)
}
