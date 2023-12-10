	oMap := Map()
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

	MsgBox(oMap.asli)