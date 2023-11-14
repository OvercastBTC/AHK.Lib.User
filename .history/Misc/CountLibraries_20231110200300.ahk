#Include <Environment>
#Include <App\Browser>
#Include <App\Autohotkey>

#Include <Paths>


CountLibraries() {
	libraries := 0
	loop files Paths.Lib "\*.ahk", "R" {
		libraries++
	}
	return libraries
}