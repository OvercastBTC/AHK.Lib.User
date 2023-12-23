/************************************************************************
 * @description The Updater
 * @file FM Updater
 * @author OvercastBTC
 * @date 2023/12/09
 * @version 0.0.1
 ***********************************************************************/

_bdrive(){
	; static b_drive := '\\corp\data\San Francisco\Engineering\AutoHotkey'
	static b_drive := '\\corp\data'
	; static b_drive := 'https://fmglobal.sharepoint.com/:u:/r/teams/AutoHotKeyUserGroup/Shared%20Documents/General/'
	; bStatus := DirExist(b_drive)
	bExist := DirExist(b_drive)
	; bStatus := GetHtml(b_drive)
	bStatus := (bExist = 'D') ? 'Ready' : 'Offline'
	Infos(bStatus, 5000)
	return bStatus
}

_Install_Git(){
	Git_Link := 'https://github.com/git-for-windows/git-for-windows.github.io/blob/main/latest-64-bit-installer.url'

	if GetHtml(Git_Link) {
		new_Link := git_InstallAHKLibrary(Git_Link)
	}
	;! Custom Setup & Silent install
	; https://github.com/git-for-windows/git/wiki/Silent-or-Unattended-Installation
}
_Install_VSCode(){
	vscode_link := 'https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user'

	if GetHtml(vscode_link) {
		;fix ..: after verifying internet connectivity, perform a silent install of vscode
	}
}
bDriveStatus() {
	; drive := '\\corp\data\'
	drive := 'https://fmglobal.sharepoint.com/:u:/r/teams/AutoHotKeyUserGroup/Shared%20Documents/General/'
	bStatus := DriveGetStatus(drive)
	Infos(bStatus)
	return bStatus
}
_is_bdrive(){
	; static b_drive := '\\corp\data\San Francisco\Engineering\AutoHotkey'
	static b_drive := '\\corp\data\', driveType:='network'
	drive := ''
	tDrive := ''
	is_dir := DriveGetList()
	dArray := StrSplit(is_dir)
	try {
		for drive in dArray {
			drive := drive ':\'
			dType := DriveGetType(drive)
			tDrive .= '[' drive '] | ' dType ' | ' DriveGetStatus(drive) '`n'
			if (drive == 'B:\') || ((dType == driveType) && (DriveGetStatus(b_drive) == 'Ready')){
				bStatus := DriveGetStatus(drive)
				Infos('[B:\] ' bStatus)
				return bStatus
			}
		}
	}
	catch {
		bStatus := DriveGetStatus(b_drive)
		Infos('[B]' bStatus)
		return bStatus
	}
}

bDrive_NotConnected(fCounter := 5, fC_units:='sec', sCounter:= 2, sC_units := 'sec') {
	static cSec := 1000
	static cMin := (cSec * 60)
	static cHrs := (cMin * 60)
	static cDays := (cHrs * 24)
	; --------------------------------------------------------------------------------
	;? fMultiply = first multiplier => if = sec|min|hr|[days], then multiply (fCounter * fMultiply)
	fMultiply := (fC_units = 'sec') ? cSec : (fC_units = 'min') ? cMin : (fC_units = 'hr') ? cHrs : cDays
	;? sMultiply = second multiplier => if = sec|min|hr|[days], then multiply (sCounter * sMultiply)
	sMultiply := (sC_units = 'sec') ? cSec : (sC_units = 'min') ? cMin : (sC_units = 'hr') ? cHrs : cDays
	; --------------------------------------------------------------------------------
	fCounter := (fCounter * fMultiply)
	sCounter := (sCounter * sMultiply)
	counter := FMCounter()
	counter.Start()
	Sleep(fCounter)
	counter.Stop()
	Sleep(sCounter)
	return
}

; --------------------------------------------------------------------------------
_CheckUpdate() {
	static b_drive 		:= '\\corp\data\San Francisco\Engineering\AutoHotkey'
	static uFolder 		:= '\CheckUpdate'
	static uFile 		:= 'ScriptVersionMap.ahk'
	static diff 		:= ''
	static Lib 			:= Paths.Lib
	static lib_uPath 	:= Lib uFolder '\' uFile
	static vNeedle 		:= "i)m_ver := '\d.\d.\d'"
	static b_uPath 	:= b_drive uFolder '\' uFile
	;? set up the variables to simplify things visually, and setup a map() for simpler updating
	static m:='major', mi:='minor', p:='patch'
	vArray := []
	
	; open_file := FileOpen(lib_uPath,'rw','UTF-8'), (open_file ~= vNeedle) ? 'true' : 'false' 
	rFile := FileRead(lib_uPath,'UTF-8') ;? can do a regex match from this one
	vFLocal := vF_Check(rFile)
	; open_file := FileOpen(b_uPath,'rw','UTF-8'), (open_file ~= vNeedle) ? 'true' : 'false'
	rFile := FileRead(b_uPath,'UTF-8') ;? can do a regex match from this one
	vFRemote := vF_Check(rFile)
	; Infos(; 	vFRemote[m] ' : ' vFLocal[m] '`n' vFRemote[mi] ' : ' vFLocal[mi] '`n' vFRemote[p] ' : ' vFLocal[p])
	; --------------------------------------------------------------------------
	vArray := []
	major := (vFRemote[m] = vFLocal[m]) ? true : false
	minor := (vFRemote[mi] = vFLocal[mi]) ? true : false
	patch := (vFRemote[p] = vFLocal[p]) ? true : false
	; infos(major '`n' minor '`n' patch)
	vArray.Push(major, minor, patch)
	; --------------------------------------------------------------------------
	;? Use an array, combined with a switch - case, to simplify if/else if/else 
	vVer := ''
	for each, vVer in vArray {
		switch vVer {
			case  true:
				if ((major = true) && (minor = true) && (patch = true)) {
					Infos(
						'No need to update at this time.`n'
						'Local Version: ' vFLocal[m] '.' vFLocal[mi] '.' vFLocal[p] '`n'
						'Remote Version: ' vFRemote[m] '.' vFRemote[mi] '.' vFRemote[p], 10000)
					return
				}
			case false:
				if (!(major = true) || !(minor != true) || !(patch != true)) {
					Infos('false')
					; fix ..: if false, download crap
				}
		}
		; infos(major '`n' minor '`n' patch) ;? validation
	}
	; --------------------------------------------------------------------------
}	
vF_Check(cFile){
	static m:='major', mi:='minor', p:='patch'
	version_map := Map(), vNeedle := "'\d.\d.\d'", ver := ''
	;! requires FileRead() for regexmatch; read file if not already done
	; rFile := FileRead(cFile,'UTF-8') 
	RegExMatch(cFile, vNeedle, &ver_match)
	;? Remove the '' and . ; place each ver in a variable
	for each, ver in ver_match {
		version := StrReplace(ver, "'", ''), vers := StrSplit(version, '.',,3)
		major := vers[1], minor := vers[2], patch := vers[3]
	}
	; Infos(ver '`n' version '`n' major '`n' minor '`n' patch) ;? validation
	;? create a map() for the variable values
	version_map.Set(m, major, mi, minor, p, patch)
	;* return the map() to the calling function
	return version_map
}

uFile(m, mi, p) {
	static b_drive 		:= '\\corp\data\San Francisco\Engineering\AutoHotkey'
	static uFolder 		:= '\CheckUpdate'
	static uFile 		:= 'ScriptVersionMap.ahk'
	static diff 		:= ''
	static Lib 			:= Paths.Lib
	static bFile 		:= b_drive uFolder '\' uFile
	static lFile 		:= Lib uFolder '\' uFile
	static fNeedle := "([0-9]\.[0-9]\.[0-9])"
	; static m:='major', mi:='minor', p:='patch' ;? vMap.Set(m, major, mi, minor, p, patch)
	; m:='', mi:='', p:=''
	vMap := Map()
	ver := ''
	folder_array := [bFile, lFile]
	aFile := []
	aMatch := []
	mArray := []
	bArray := []
	lArray := []
	fLine := ''
	f := ''
	match:=''
	aLine := ''
	; --------------------------------------------------------------------------------
	bfo := FileOpen(bFile, 'rw', 'UTF-8')
	lfo := FileOpen(lFile, 'rw', 'UTF-8')
	bArray := vReadFile(bFile)
	lArray := vReadFile(lFile)
	; bArray := vReadFile(bfo)
	; lArray := vReadFile(lfo)
	vReadFile(vFile){
		; Infos(vFile)
		; rFile := FileRead(vFile,'UTF-8')
		Sleep(300)
		loop read vFile {
			aFile.Push(A_LoopReadLine)
			fLine .= A_LoopReadLine . '`n'
		}
		return aFile
	}
	; --------------------------------------------------------------------------------
	new_bFile := vMatchLine(bArray, bFile)
	vMatchLine(vArray, file_Open) {
		Infos(A_Index FileRead(file_open))
		RegExMatch(file_Open, fNeedle, &vMatch_Array)
		for each, value in vMatch_Array {
			Infos(value)
		}
		for each, aLine in aFile {
			; Infos(aLine '`n' fNeedle)
				if (aLine ~= fNeedle) {
					Infos(aLine)
					version := StrReplace(aLine, "'", ''), vers := StrSplit(version, '.',,3)
					major := vers[1], minor := vers[2], patch := vers[3]
				} 
			; }
			; str_match := StrSplit(aLine,')','i ) "')
			; rMatch := str_match[2]
			; Infos(rMatch)
			; ver_match := mArray[rMatch]
			; ;? Remove the '' and . ; place each ver in a variable
			; for each, ver in ver_match {
			; 	version := StrReplace(ver, "'", ''), vers := StrSplit(version, '.',,3)
			; 	major := vers[1], minor := vers[2], patch := vers[3]
			; }
			; Infos(ver '`n' version '`n' major '`n' minor '`n' patch) ;? validation
			;? create a map() for the variable values
			; vMap.Set(m, major, mi, minor, p, patch)
			nM  := RegExReplace(major, '([0-9]{1})', m)
			nMi := RegExReplace(minor, '([0-9]{1})', mi)
			nP  := RegExReplace(patch, '([0-9]{1})', p)
			aLine := "'" nM '.' nMi '.' nP "'"
			aFile.RemoveAt(A_Index)
			aFile.InsertAt(A_Index, aLine)
		}
		; --------------------------------------------------------------------------------
		; --------------------------------------------------------------------------------
		/**
		 * function: write each value in the array to @param f (string variable) 
		 */
		f .= aLine . '`n'
		; --------------------------------------------------------------------------------
		rf := file_Open.Write(f)
		Infos(f)
		file_Open := ''
		return rf
	}
	return 0
}