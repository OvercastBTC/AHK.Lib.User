;---------------------------------------------------------------------------
;                           Autoexecute Section
;---------------------------------------------------------------------------
#Include AE.V2.ahk
#Requires AutoHotkey v2

;---------------------------------------------------------------------------
;                        List of Files to Include/Run
;---------------------------------------------------------------------------
#Include Startup.V2.ahk
Startup.Check_Startup_Status()

SetTimer(update, 900000) ; Download Lib Folder for updates. attempt every 15 minutes until complete.

#Include multi_clipboard.v2.ahk

;@Ahk2Exe-IgnoreBegin
Run("C:\Users\" A_UserName "\OneDrive - FM Global\Documents\Autohotkey\Lib\RecLibHints.ahk")
#Include %A_ScriptDir%\Personal_Information.V2.ahk
#Include %A_ScriptDir%\Personalized_Scripts.V2.ahk
;@Ahk2Exe-IgnoreEnd

#Include AutoCorrect.V2.ahk
#Include Common Abbrevations.V2.ahk
#Include FMDS Titles.V2.ahk
#Include Misc Scripts.V2.ahk



update()
{
        
    ;MsgBox("Autoupdate called")

    Input_Lib_File := '`'B:\San Francisco\Engineering\AutoHotkey\Lib.zip`'' ; Set location of the zip file.

    Output_Lib_File := '`'C:\Users\' A_UserName '\OneDrive - FM Global\Documents\Autohotkey\Lib`''

    B_Drive_File_Check := StrSplit(Input_Lib_File, '`'', A_Space)
    
    if(!FileExist(B_Drive_File_Check[2])) ; Check to see if the computer can see the file on the B: Drive
        {
            ; B: Drive is not connected. Do nothing.
            ;MsgBox("B: Drive not found.")
            return
        }

    else{

            RunWait("PowerShell.exe -WindowStyle hidden -Command Expand-Archive -LiteralPath " Input_Lib_File " -DestinationPath " Output_Lib_File " -Force") ; Open Powershell and expand the zip file to the selected folder. "-Force" will overwrite files with name conflicts. Powershell will close after successful operation.
            SetTimer( , 0) ; update completed, turn off the timer.
            ;MsgBox("Update complete. Timer set to 0")
            return

        }
}