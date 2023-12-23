#Requires AutoHotkey v2+
#Include <Directives\__AE.v2>

; --------------------------------------------------------------------------------
/**
 * function...: Include Library
 */

#Include <App\Git>
#Include <System\Web>

; --------------------------------------------------------------------------------
;---------------------------------------------------------------------------
;                              CHANGELOG
;---------------------------------------------------------------------------

; v1.0 - New script.

;---------------------------------------------------------------------------
;                               SCRIPT
;---------------------------------------------------------------------------


class update_script {
	static bdrive := '\\corp\data\'
	; static b_drive := '\\corp\data\San Francisco\Engineering\AutoHotkey'
	static b_drive := this.bdrive 'San Francisco\Engineering\AutoHotkey'
	static Input_File := this.b_drive '\Install.zip' ; Set location of the zip file.
	static Input_HznPlus_File := this.b_drive '\HznPlus.v2.zip' ; Set location of the zip file.
	; Set location of the zip file.
	static Input_Lib_File := this.b_drive '\Lib.zip' 
	
	; Output_Lib_File := '`'C:\Users\' A_UserName '\OneDrive - FM Global\Documents\Autohotkey\Lib`''
	static Output_Lib_File := A_MyDocuments '\AutoHotkey\Lib'
	
	class FMCounter {
		__New() {
			this.interval := 1000
			this.count := 0
			; Tick() has an implicit parameter "this" which is a reference to
			; the object, so we need to create a function which encapsulates
			; "this" and the method to call:
			this.timer := ObjBindMethod(this, "Tick")
		}
		Start() {
			SetTimer( this.timer, this.interval)
			Infos("Counter started", 5000)
		}
		Stop() {
			; To turn off the timer, we must pass the same object as before:
			SetTimer( this.timer, 0)
			Infos("Counter stopped at " this.count)
		}
		; In this example, the timer calls this method:
		Tick() {
			++this.count
			Infos(++this.count)
		}
		MyCallBack() {
			SetTimer( this.timer, 1000)
		}
	}

	static _bdrive(){
		; static b_drive := '\\corp\data\San Francisco\Engineering\AutoHotkey'
		bExist := DirExist(this.b_drive)
		bStatus := (bExist = 'D') ? 'Ready' : 'Offline'
		Infos(bStatus)
		return bStatus
	}

	static bDriveStatus() {
		bStatus := DriveGetStatus(this.b_drive)
		Infos(bStatus)
		return bStatus
	}

	; B_Drive_File_Check := StrSplit(Input_File, '`'', A_Space)
	static internetOn(link?) {
		link := 'https://www.autohotkey.com/'
		HTTP := ComObject("WinHttp.WinHttpRequest.5.1")
		HTTP.Open("GET", link, true)
		HTTP.Send()
		try HTTP.WaitForResponse()
		catch Any {
			return ""
		}
		return HTTP.ResponseText
	}

	; Check to see if the computer can see the file on the B: Drive
	static updateLib() {

		checkConnection()

		checkConnection(){
			((this.bDriveStatus() != 'Ready') || (this.internetOn() = '')) ? bDrive_NotConnected() : bDrive_Connected()
		}
		
		bDrive_NotConnected(fCounter := 5, fC_units:='sec', sCounter:= 2, sC_units := 'sec') {
			cSec := 1000
			cMin := (cSec * 60)
			cHrs := (cMin * 60)
			cDays := (cHrs * 24)
			;? fMultiply = first multiplier => if = sec|min|hr|[days], then multiply (fCounter * fMultiply)
			fMultiply := (fC_units = 'sec') ? cSec : (fC_units = 'min') ? cMin : (fC_units = 'hr') ? cHrs : cDays
			;? sMultiply = second multiplier => if = sec|min|hr|[days], then multiply (sCounter * sMultiply)
			sMultiply := (sC_units = 'sec') ? cSec : (sC_units = 'min') ? cMin : (sC_units = 'hr') ? cHrs : cDays
			fCounter := (fCounter * fMultiply)
			sCounter := (sCounter * sMultiply)
			counter := this.FMCounter()
			counter.Start()
			Sleep(fCounter)
			counter.Stop()
			Sleep(sCounter)
		}

		bDrive_Connected() {
			; fix ...: suggest copying Path.ahk to MyDocuments\AutoHotkey\Lib folder for easier use of paths. Then move it to the "Main" Lib folder, and create a blank one for them to use.

			; static mainLib := 'C:\Users\' A_UserName '\AppData\Local\Programs\AutoHotkey\v2\Lib\TestingStuff'
			; static mainLib := 'C:\Users\' A_UserName '\AppData\Local\Programs\AutoHotkey\v2\Lib'
			static persRepo := A_MyDocuments '\AutoHotkey'
			; static persLib := this.persRepo '\Lib'
			static persLib := persRepo '\Lib\TestingStuff'
			static mainLib := persLib '\TestingStuff'
			If !DirExist(mainLib) {
				DirCreate(mainLib)
			}
			If !DirExist(persRepo) {
				DirCreate(persRepo)
			}
			If !DirExist(persLib) {
				DirCreate(persLib)
			}
			; Infos('Does the directory exist??? [' DirExist(mainLib) ']')
			; MsgBox("Please select the folder destination that you want to download the script to.")
			infos(
				'The main library files will be installed to ' 
				this.mainLib '(not intended for editing).'
				'`n'
				'Your main scripts for running and person editing/additions is located in: `n'
				'[' A_MyDocuments '\AutoHotkey]'
				'`n'
				'Note: This also includes a personal library (Lib) folder where AutoHotkey will automatically look for library files.'
			)
		
			; Get_Output_Folder := FileSelect("D", , "Select a folder") ; Have the user select the desired folder to expand the zip file.
			Get_Output_Folder := this.mainLib
	
			PS_Output_Folder := "'" Get_Output_Folder "'"
	
			; Check_Overwrite1 := Get_Output_Folder "\Personal_Information.V2.ahk"
			Check_Overwrite1 := this.persLib "\Personal_Information.V2.ahk"
			
			; Check_Overwrite2 := Get_Output_Folder "\Personalized_Scripts.V2.ahk"
			Check_Overwrite2 := this.persRepo "\Personalized_Scripts.V2.ahk"
			
			; Check_Overwrite3 := Get_Output_Folder "\AutoCorrect-Personal-Additions.V2.ahk"
			Check_Overwrite3 := this.persLib "\AutoCorrect-Personal-Additions.V2.ahk"
			
			if(FileExist(Check_Overwrite1) || FileExist(Check_Overwrite2) || FileExist(Check_Overwrite3)) {
					myGui := Gui(,"Personal Files Already in Selected Output Folder",)
					myGui.Opt("AlwaysOnTop")
					myGui.SetFont("s12")
					myGui.Add("Text",, "This folder already contains files, which may contain personal customizations.`n`nDo you want to overwrite these files?")
					myGui.Add("Button","x15 +default", "Overwrite Files").OnEvent("Click", Clickedoverwrite)
					myGui.Add("Button","x+m", "Do Not Overwrite").OnEvent("Click", Clickeddonotoverwrite)
					myGui.Show("w500")
					
					ClickedOverwrite(*) {
						myGui.Destroy()
						;? Open Powershell and expand the zip file to the selected folder. "-Force" will overwrite files with name conflicts. Powershell will close after successful operation.
						RunWait("Powershell.exe -WindowStyle hidden -Command Expand-Archive -LiteralPath " this.Input_File " -DestinationPath " PS_Output_Folder " -Force") 
					}
	
					ClickedDoNotOverwrite(*) {
						myGui.Destroy()
						;? Open Powershell and expand the zip file to the selected folder. Powershell will close after successful operation.
						RunWait("Powershell.exe -WindowStyle hidden -Command Expand-Archive -LiteralPath " this.Input_File " -DestinationPath " PS_Output_Folder "") 
					}
			} else {
				;? Open Powershell and expand the zip file to the selected folder. "-Force" will overwrite files with name conflicts. Powershell will close after successful operation.
				RunWait("Powershell.exe -WindowStyle hidden -Command Expand-Archive -LiteralPath " this.Input_File " -DestinationPath " PS_Output_Folder " -Force")
			}
			
			RunWait("PowerShell.exe -WindowStyle hidden -Command Expand-Archive -LiteralPath " this.Input_Lib_File " -DestinationPath " this.Output_Lib_File " -Force") ; Open Powershell and expand the zip file to the selected folder. "-Force" will overwrite files with name conflicts. Powershell will close after successful operation.
			
			; HrznGui := Gui(,"Install Horizon Plus",)
			HrznGui := Gui('+AlwaysOnTop',"Install Horizon Plus",)
			; HrznGui.Opt("AlwaysOnTop")
			HrznGui.SetFont("s12")
			HrznGui.AddText(, "Are you a Horizon User?")
			HrznGui.AddButton("x15 +default", "Install Horizon Plus Utility").OnEvent("Click", ClickedInstall)
			HrznGui.AddButton("x+m", "Cancel").OnEvent("Click", ClickedCancel)
			; HrznGui.Show("w500")
			; HrznGui.Show('AutoSize')
			HrznGui.Show()
			
			ClickedInstall(*)
			{
				HrznGui.Destroy()
				RunWait("Powershell.exe -WindowStyle hidden -Command Expand-Archive -LiteralPath " this.Input_HznPlus_File " -DestinationPath " PS_Output_Folder " -Force") ; Open Powershell and expand the zip file to the selected folder. "-Force" will overwrite files with name conflicts. Powershell will close after successful operation.
			}
			
			ClickedCancel(*)
			{
				HrznGui.Destroy()
			}
		}
	}
	; --------------------------------------------------------------------------------

	static _Install_Git(){
		Git_Link := 'https://github.com/git-for-windows/git-for-windows.github.io/blob/main/latest-64-bit-installer.url'

		if GetHtml(Git_Link) {
			new_Link := git_InstallAHKLibrary(Git_Link)
		}
	}
	static _Install_VSCode(){
		vscode_link := 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user'

		if GetHtml(vscode_link) {
			;fix ..: after verifying internet connectivity, perform a silent install of vscode
		}
	}
}
; --------------------------------------------------------------------------------
; end of the class