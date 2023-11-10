#Include <App\Autohotkey>
#Include <Abstractions\Base>

#HotIf WinActive(AutoHotkey_ProgFiles.exeTitle)
^BackSpace::DeleteWord()
#HotIf
