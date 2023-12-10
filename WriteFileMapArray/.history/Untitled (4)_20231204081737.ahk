#Requires AutoHotkey v2.0
#Include <Directives\__AE.v2>
#Include <Tools\Info>
#Include <__setup>
#Include <.history\Directives\_setup_20231106202943>
	
	; oMap := Map()
	oMap := Setups.
	oMap.CaseSense := "Off"
	oMap.Set("Red", "ff0000", "Green", "00ff00", "Deep Blue", "0000ff")

	oMap.DefineProp("__get", { Call: Get })

	oMap.Default := "Def"

	Get(this, Key, Params) {

		; in this way Return the value of MapObj.Default, if defined:
		try {
			if !Params.Length
				return this[key] ;Or this.Get(Key)
		
			return this[key][Params*]
			
		} catch UnsetItemError
			throw UnsetItemError(Key)
	}

	; MsgBox(oMap.asli)
	; Infos(oMap.asli)
	Infos(oMap.red)