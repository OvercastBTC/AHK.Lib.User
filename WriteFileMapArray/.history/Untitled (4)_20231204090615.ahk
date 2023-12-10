#Requires AutoHotkey v2.0
#Include <Directives\__AE.v2>
#Include <Tools\Info>
#Include <Directives\_user_info>
#Include <WriteFileMapArray\Object2Str>
	
	oMap := Map()
	oMap.CaseSense := 'Off'
	; oMap := user_info.info_map
	oMap := user_info.info_arr
	; oMap.Set("Red", "ff0000", "Green", "00ff00", "Deep Blue", "0000ff")
	for key, value in user_info.info_map{
		newKey := StrReplace(key,A_Space, '_')
		oMap.InsertAt(A_Index, newKey)
	}

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
	Infos(oMap.manager)
	; Infos(Object2Str(oMap))