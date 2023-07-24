#Include <Utils\Win>

class Autohotkey {

	static path := "C:\Users\" . A_UserName . "\AppData\Local\Programs\AutoHotkey"
	static currVersion := this.path "\v2"
	static processExe := A_AhkPath
	static exeTitle := "ahk_exe " this.processExe

}