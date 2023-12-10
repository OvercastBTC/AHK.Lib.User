#Requires AutoHotkey v2.0
#Include <Directives\__AE.v2>
#Include <Tools\Info>

oMap := Map()
oMap.CaseSense := "Off"
oMap.Set("Red", "ff0000", "Green", "00ff00", "Deep Blue", "0000ff")

oMap.DefineProp('__get', { call: oMap.Get }) ; Trick of thqby

MsgBox('oMap["Red"]: ' oMap["Red"] "`noMap.Red: " oMap.red)