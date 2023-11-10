#Include <Utils\Win>

class Autohotkey {

	static path := A_ProgramFiles "\AutoHotkey"
	static currVersion := this.path "\v2"
	static processExe := A_AhkPath
	static exeTitle := "ahk_exe " this.processExe

}
class LocalUser extends AutoHotkey {

	static path := RegExReplace(A_AppData, 'i)\Roaming', '\Local\Programs\' . "\AutoHotkey")
	static currVersion := this.path "\v2"
	static processExe := A_AhkPath
	static exeTitle := "ahk_exe " this.processExe

}