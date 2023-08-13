﻿; ======================================================================================================================
; Function:       Creates AHK #Include files for images providing a function which will internally create a bitmap/icon.
; AHK version:    2.0 (U32/U64)
; Script version: 2.0.00.00/2023-07-31/iPhilip - converted script to AutoHotkey v2
;                 1.0.00.02/2013-06-02/just me - added support for icons (HICON)
;                 1.0.00.01/2013-05-18/just me - fixed bug producing invalid function names
;                 1.0.00.00/2013-04-30/just me
; Credits:        Bitmap creation is based on "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN ->
;                 http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
; ======================================================================================================================
; This software is provided 'as-is', without any express or implied warranty.
; In no event will the authors be held liable for any damages arising from the use of this software.
; ======================================================================================================================
#Requires AutoHotkey v2.0
; ======================================================================================================================
; Some strings used to generate the #Include file
; ======================================================================================================================
Header1 := "
(
; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
)"
; ----------------------------------------------------------------------------------------------------------------------
Header2 := "
(
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
)"
; ----------------------------------------------------------------------------------------------------------------------
Footer1 := "
(
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", StrPtr(B64), "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", &DecLen := 0, "Ptr", 0, "Ptr", 0)
   Return False
Dec := Buffer(DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", StrPtr(B64), "UInt", 0, "UInt", 0x01, "Ptr", Dec, "UIntP", &DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream := ComValue(13, 0))
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
SI := Buffer(8 + 2 * A_PtrSize, 0), NumPut("UInt", 1, SI, 0)
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", &pToken := 0, "Ptr", SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", &pBitmap := 0)
)"
; ----------------------------------------------------------------------------------------------------------------------
Footer2 := "
(
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
Return hBitmap
}
)"
; ======================================================================================================================
; Read INI file
; ======================================================================================================================
IniFile := A_ScriptFullPath . ":INI"
ImgDir := OutDir := A_ScriptDir
Value := IniRead(IniFile, "Image2Include", "ImgDir", 0)
If (Value) && DirExist(Value)
   ImgDir := Value
Value := IniRead(IniFile, "Image2Include", "OutDir", 0)
If (Value) && DirExist(Value)
   OutDir := Value
ImgFile := ""
; ======================================================================================================================
; Select an image file
; ======================================================================================================================
MainGui := Gui(, "Convert Image to #Include File")
MainGui.MarginX := MainGui.MarginY := 20
MainGui.Add("Text", "xm cNavy", "Select Image File:")
MainGui.Add("Edit", "xm y+5 w400 vImgFile +ReadOnly")
MainGui.Add("Button", "x+10 yp hp vBtnFile", "...").OnEvent("Click", SelectFile)
MainGui.Add("Text", "xm y+5 cNavy", "Select Output Directory:")
MainGui.Add("Edit", "xm y+5 w400 vOutDir +ReadOnly", OutDir)
MainGui.Add("Button", "x+10 yp hp vBtnFolder", "...").OnEvent("Click", SelectFolder)
MainGui["BtnFolder"].GetPos(&P1X, , &P1W)
MainGui.Add("Text", "xm y+5 cNavy", "Script File:")
MainGui.Add("Edit", "xm y+5 w400 vOutFile +ReadOnly")
MainGui.Add("CheckBox", "xm vReturnHICON", "Return HICON handle")
MainGui.Add("Button", "xm vP2", "Convert Image").OnEvent("Click", Convert)
MainGui["P2"].GetPos(, , &P2W)
MainGui.Add("Button", "x" . (P1X + P1W - P2W) . " yp wp vBtnShow Disabled", "Show Script").OnEvent("Click", Show)
MainGui.Add("StatusBar", "vStatusBar")
MainGui["BtnFile"].Focus()
MainGui.Show()
MainGui.OnEvent('Close', (*) => ExitApp())
; Prepare the GUI which will show the generated script on demand
ScriptGui := Gui("+Owner" MainGui.Hwnd)
ScriptGui.MarginX := ScriptGui.MarginY := 0
ScriptGui.SetFont(, "Segoe UI")
ScriptGui.Add("Edit", "x0 y0 w" . Round(A_ScreenWidth * 0.8) . " h" . Round(A_ScreenHeight * 0.8) . " vEdit HScroll")
ScriptGui.Add("StatusBar", "vStatusBar")
; ----------------------------------------------------------------------------------------------------------------------
SelectFile(*) {
   global ImgDir, ImgFile, ImgName
   MainGui.Opt("+OwnDialogs")
   ImgFile := FileSelect(1, ImgDir, "Select a Picture", "Img (*.bmp; *.emf; *.exif; *.gif; *.ico; *.jpg; *.png; *.tif; *.wmf)")
   If !(ImgFile)
      Return
   MainGui["ImgFile"].Value := ImgFile
   SplitPath ImgFile, , &ImgDir, &ImgExt, &ImgName
   ImgName := RegExReplace(ImgName, "[\W]")
   ImgName .= "_" . ImgExt
   MainGui["OutFile"].Value := OutDir "\Create_" ImgName ".ahk"
   IniWrite ImgDir, IniFile, "Image2Include", "ImgDir"
}
; ----------------------------------------------------------------------------------------------------------------------
SelectFolder(*) {
   global OutDir, ImgFile
   MainGui.Opt("+OwnDialogs")
   OutDir := DirSelect("*" OutDir, 1, "Select a folder to store the scripts:")
   If !(OutDir)
      Return
   MainGui["OutDir"].Value := OutDir
   ImgFile := MainGui["ImgFile"].Value
   If (ImgFile != "") {
      MainGui["OutFile"].Value := OutDir "\Create_" ImgName ".ahk"
   }
   IniWrite OutDir, IniFile, "Image2Include", "OutDir"
}
; ----------------------------------------------------------------------------------------------------------------------
Convert(*) {
   global OutFile
   MainGui["StatusBar"].SetText("")
   MainGuiCtrlContents := MainGui.Submit(false)
   OutFile := MainGuiCtrlContents.OutFile
   If !FileExist(ImgFile) {
      MainGui["BtnFile"].Focus()
      Return
   }
   If !DirExist(OutDir) {
      MainGui["BtnFolder"].Focus()
      Return
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; Read the image file
   ; -------------------------------------------------------------------------------------------------------------------
   File := FileOpen(ImgFile, "r")
   BinLen := File.Length
   Bin := Buffer(BinLen)
   File.RawRead(Bin, BinLen)
   File.Close()
   ; -------------------------------------------------------------------------------------------------------------------
   ; Encode the image file
   ; -------------------------------------------------------------------------------------------------------------------
   DllCall("Crypt32.dll\CryptBinaryToString", "Ptr", Bin, "UInt", BinLen, "UInt", 0x01, "Ptr", 0, "UIntP", &B64Len := 0)
   B64 := Buffer(B64Len * 2, 0)
   DllCall("Crypt32.dll\CryptBinaryToString", "Ptr", Bin, "UInt", BinLen, "UInt", 0x01, "Ptr", B64, "UIntP", &B64Len)
   Bin := ""
   B64 := RegExReplace(StrGet(B64), "\r\n")
   B64Len := StrLen(B64)
   ; -------------------------------------------------------------------------------------------------------------------
   ; Write the AHK script
   ; -------------------------------------------------------------------------------------------------------------------
   PartLength := 16000
   CharsRead  := 1
   File := FileOpen(OutFile, "w`n", "UTF-8")
   File.WriteLine(Header1 . "`nCreate_" . ImgName . "(NewHandle := False) {`nStatic hBitmap := 0`n" Header2)
   Part := 'B64 := "'
   While (CharsRead < B64Len) {
      File.WriteLine(Part . SubStr(B64, CharsRead, PartLength) . '"')
      CharsRead += PartLength
      If (CharsRead < B64Len)
         Part := 'B64 .= "'
   }
   File.WriteLine(Footer1)
   If (MainGuiCtrlContents.ReturnHICON)
      File.WriteLine('DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", &hBitmap := 0)')
   Else
      File.WriteLine('DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", &hBitmap := 0, "UInt", 0)')
   File.WriteLine(Footer2)
   File.Close()
   MainGui["BtnShow"].Enabled := true
   SplitPath OutFile, &OutFileName
   MainGui["StatusBar"].SetText("  File " . OutFileName . " successfully created!")
}
; ======================================================================================================================
; Show the generated script
; ======================================================================================================================
Show(*) {
   ; -------------------------------------------------------------------------------------------------------------------
   ; Reread the AHK script and show it
   ; -------------------------------------------------------------------------------------------------------------------
   B64 := FileRead(OutFile)
   If !(B64)
      Return
   Size := FileGetSize(OutFile)
   ScriptGui["Edit"].Value := B64
   Lines := EditGetLineCount(ScriptGui["Edit"])
   ScriptGui["StatusBar"].SetText("  Size: " . Size . " bytes - Lines: " . Lines)
   ScriptGui.Title := OutFile
   ScriptGui.Show()
}