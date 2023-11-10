#Include <Utils\Win>

class Autohotkey {

	static path := A_ProgramFiles "\AutoHotkey"
	static currVersion := this.path "\v2"
	static processExe := A_AhkPath
	static exeTitle := "ahk_exe " this.processExe

}
class LocalUser extends AutoHotkey {

	static path := A_AppData '..\Local' "\AutoHotkey"
	static currVersion := this.path "\v2"
	static processExe := A_AhkPath
	static exeTitle := "ahk_exe " this.processExe

}