#Requires AutoHotkey v2.0-beta
#SingleInstance force
oMap := Map()
oMap.CaseSense := "Off"
oMap.Set("Red", "ff0000", "Green", "00ff00", "Deep Blue", "0000ff")

oMap.DefineProp('__get', { call: oMap.Get }) ; Trick of thqby

MsgBox('oMap["Red"]:' oMap["Red"] "`noMap.Red:" oMap.red)