#Persistent
#SingleInstance ignore
#Include C:\Users\T04006\Desktop\Downloads\Software\AutoHotkey12204\About\About.ahk
#Include %A_ScriptDIR%\MostrarGUI.ahk
#Include %A_ScriptDIR%\MostrarGUIImatge.ahk
#Include %A_ScriptDIR%\CanviHK.ahk

shell := ComObjCreate("Shell.Application")
fitxer_ids := "series.txt"
fitxer_noms := "noms.txt"
fitxer_cfg := "config.ini"
tIDs := Object()	;passarà a ser tIDs, relació ID-nom_serie_api (fitxer series.txt)
tNoms := Object()	;relació nom_en_brut-ID (fitxer noms.txt)
t_epis := Object() 	;taula temporal per no buscar un episodi ja buscat
tPaths := Object()	;taula temporal de fitxers d'entrada
tFixes := Object()	;taula temporal de fitxers de sortida
tErrors := Object()	;taula temporal de fitxers erronis

Menu, Tray, NoStandard
Menu, Tray, Add, Canviar hotkey, CanviarHK
Menu, Tray, Add
if (PermetAbout())
	Menu, Tray, Add, About, About
Menu, Tray, Add, Sortir, Sortir
Menu, Tray, Tip, FixName12

CarregaFitxer(tIDs, fitxer_ids)
CarregaFitxer(tNoms, fitxer_noms)
CarregaINI()
Hotkey, IfWinActive, ahk_class CabinetWClass
Hotkey, %hk%, Accio
return

Accio()
{
	global shell, tPaths, tFixes, tErrors
	
	Hotkey, %A_ThisHotkey%, Off
	winget, hwnd, id
	for wnd in shell.Windows
	{
		if (wnd.hwnd = hwnd)
		{
			exp_hwnd := hwnd
			sel := wnd.Document.SelectedItems
			i := 0
			toSort := ""
			for item in sel 
				toSort .= item.path . "`n"
			break
		}
	}
	Sort, toSort
	tPaths := Object()
	i := 0
	Loop, parse, toSort,`n
	{
		if (A_LoopField != "")
		{
			i++
			tPaths[i] := A_LoopField
		}
	}
	if (i > 0)
	{
		tFixes := Object()
		tErrors := Object()
		valids := 0
		novalids := 0
		SetCursors(32650)
		Loop %i%
		{
			origen := tPaths[A_Index]
			SplitPath, origen, filein, path
			fixname := Avalua(filein)
			if (SubStr(fixname, 1, 6) != "ERROR.")
			{
				valids++
				tFixes[origen] := path . "\" . fixname
			}
			else
			{
				novalids++
				tErrors[origen] := SubStr(fixname, 7)
			}
		}
		RestoreCursors()
		MostrarGUI(valids, novalids)
	}
	Hotkey, %A_ThisHotkey%, On
}

Avalua(fitxer)
{
	global tNoms, tIDs
	
	regexp1 := "i)\.(s\d{2}e\d{2})"
	regexp2 := "i)(\.\d{3,4}\.)"
	regexp3 := "i) - (\d{2}x\d{2}) -"
	regexp4 := "i)(\.\d{1,2}x\d{2}\.)"
	
	SplitPath, fitxer,,, ext

	pos := RegExMatch(fitxer, regexp1, match)
	if (pos > 0) ;estem a sXXeYY
	{
		nom_serie := Substr(fitxer, 1, pos - 1)
		temp := Substr(match1, 2, 2)
		epi := Substr(match1, 5, 2)
	}
	else
	{
		pos := RegExMatch(fitxer, regexp2, match)
		if (pos > 0) ; estem a XXYY
		{
			;mirem si es torna a complir en la resta, si és així ens quedem amb la darrera
			final := false
			while(!final)
			{
				posant := pos
				pos := RegExMatch(fitxer, regexp2, match, pos + StrLen(match1) - 1)
				if (match1 = "")
					final := true
			}
			pos := posant			
			nom_serie := Substr(fitxer, 1, pos - 1)
			tot := Substr(fitxer, pos + 1, 4)
			if tot is integer
			{
				temp := Substr(tot, 1, 2)
				epi := Substr(tot, 3, 2)
			}
			else
			{
				temp := Substr(tot, 1, 1)
				epi := Substr(tot, 2, 2)
			}
		}
		else
		{
			pos := RegExMatch(fitxer, regexp3)
			if (pos > 0) ;estem a XXxYY
			{
				nom_serie := Substr(fitxer, 1, pos - 1)
				temp := Substr(fitxer, pos + 3, 2)
				epi := Substr(fitxer, pos + 6, 2)
			}
			else
			{
				pos := RegExMatch(fitxer, regexp4, match)
				if (pos > 0) ;estem a .XXxYY.
				{
					nom_serie := Substr(fitxer, 1, pos - 1)
					posx := InStr(match1, "x")
					if (posx = 4) ;la temporada és de dos dígits
					{
						temp := Substr(match1, 2, 2)
						epi := Substr(match1, 5, 2)
					}
					else ;la temporada és d'un dígit
					{
						temp := Substr(match1, 2, 1)
						epi := Substr(match1, 4, 2)
					}
				}
				else
					return "ERROR.No s'ha trobat coincidència (.sAAeBB. | .AABB. | - AAxBB - | .AAxBB.)"
			}
		}
	}
	
	if (Substr(temp,1, 1) = "0")
		temp := Substr(temp, 2)
	
	id := tNoms[nom_serie]
	if (id != "")
		nom_serie := tIDs[id]
		
	nom_epi := BuscaAPITVMaze(id, temp, epi, nom_serie)
	if (SubStr(nom_epi, 1, 6) = "ERROR.")
		return %nom_epi%	

	valor = %nom_serie% %temp%x%epi% %nom_epi%.%ext%
	return %valor%
}

BuscaAPITVmaze(id, se, ep, ByRef nom)
{
	global tIDs, tNoms, t_epis
	
	nom_ep := ""
	
	if (id = "")
	{
		nom_tractat := TractaNom(nom)
		
		url_id := "https://api.tvmaze.com/singlesearch/shows?q=" . nom_tractat
		
		headers := {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:51.0) Gecko/20100101 Firefox/51.0", "Access-Control-Allow-Origin": "*"}
		result := ObtePagina(url_id, headers, "t")
			 
		serie_ID := GetJSON(result, "id")
		if (serie_ID != "")
		{
			nom_bo := GetJSON(result, "name")
			StringUpper, nom_bo, nom_bo, T
			;si ha trobat la sèrie tractant el nom es guarda la relació nom_en_brut-id per properes execucions, i també el nom_bo
			tNoms[nom] := serie_ID
			tNoms[nom_bo] := serie_ID
			;ens guardem (sobreescrivint, si ja existeix) la relació id-nom_serie per properes execucions
			tIDs[serie_id] := nom_bo
			nom := nom_bo
		}
		else
		{
			InputBox, nom_usu, Fixname, La sèrie %nom% no es troba.`nIntrodueix el nom en versió original per seguir buscant:
			ocu := 1
			url_id := "https://api.tvmaze.com/search/shows?q=" . nom_usu
			result := ObtePagina(url_id, headers, "t")
			Loop
			{
				imatge := GetJSON(result, "medium", ocu)
				if (imatge = "")
					break
				else
				{
					es_serie := MostraOpcions(imatge)
					if (es_serie)
					{
						serie_id := GetJSON(result, "id", ocu)
						if (serie_ID != "")
						{
							nom_bo := GetJSON(result, "name")
							StringUpper, nom_bo, nom_bo, T
							;si ha trobat la sèrie tractant el nom es guarda la relació nom_en_brut-id per properes execucions, i també el nom_bo
							tNoms[nom] := serie_ID
							tNoms[nom_bo] := serie_ID
							;ens guardem (sobreescrivint, si ja existeix) la relació id-nom_serie per properes execucions
							tIDs[serie_id] := nom_bo
							nom := nom_bo
						}
						break
					}
					else
						ocu++
				}
			}
			if (serie_id = "")
			{
				nom_ep := "ERROR.Sèrie desconeguda"
				headers := ""
				return %nom_ep%
			}
		}
	}
	else
		serie_id := id
		
	val = %serie_id%º%se%º%ep%
	value := t_epis[val]
	if (value = "")
	{
		url_nom := "https://api.tvmaze.com/shows/" . serie_ID . "/episodebynumber?season=" . se . "&number=" . ep
		
		result := ObtePagina(url_nom, headers, "t")
			
		message := 	GetJSON(result, "message")
		if (message != "Page not found." and message != "Unknown episode")
		{
			nom_ep := GetJSON(result, "name")
			if (nom_ep != "")
			{
				StringUpper, nom_ep, nom_ep, T
				nom_ep := DepuraNom(nom_ep)
				val = %serie_id%º%se%º%ep%
				t_epis[val] := nom_ep
			}
			else
				nom_ep := "ERROR.Episodi no trobat"
			
		}
		else	
			nom_ep := "ERROR.Episodi desconegut"
	}
	else
		nom_ep := t_epis[val]

	headers := ""
	return nom_ep
}

TractaNom(nom)
{
	nom := StrReplace(nom, ".", A_Space)
	nom := StrReplace(nom, "S H I E L D", "S.H.I.E.L.D.")
	nom := StrReplace(nom, "(", "")
	nom := StrReplace(nom, ")", "")
	warray := StrSplit(nom, " ")
	firstw := warray[1]
	if firstw in Marvels,Marvel's,DCs,DC's
		nom := StrReplace(nom, firstw . " ")
	lastw := warray[warray.MaxIndex()]
	if lastw in US,UK,AUS
		nom := StrReplace(nom, " " . lastw, "")
	if lastw is integer
		nom := StrReplace(nom, " " . lastw, "")
	nom := StrReplace(nom, "S.h.i.e.l.d.", "S.H.I.E.L.D.")
	nom := Trim(nom)
	return nom
}

GetJSON(txt, clau, occurrence := 1)
{

	clau := Chr(34) . clau . Chr(34) . ":"
	pos_ini := InStr(txt, clau,,, occurrence)
	if (pos_ini = 0)
		return ""
	pos_ini += StrLen(clau)
	txt := SubStr(txt, pos_ini)
	pos_fin := InStr(txt, ",""")
	pos_fin2 := InStr(txt, "}")
	if ((pos_fin > pos_fin2) or (pos_fin = 0))
		pos_fin := pos_fin2
	val := Substr(txt, 1, pos_fin - 1)
	val := Trim(val, """")
	return %val%
}

DepuraNom(nom)
{
	forbidden = \,/,?,*,"",<,>,|,:
	Loop, parse, forbidden, `,
	{
		if InStr(nom, A_LoopField)
		{
			StringReplace, nom, nom, %A_LoopField%, _, All
			trobat := true
		}
	}
	if trobat
		nom .= "[warning]"
	return nom
}

ObtePagina(url, h, t_b := "t")
{
	static req := ComObjCreate("Msxml2.XMLHTTP")
	
	req.Open("GET", url, false)

	for key,value in h
		req.SetRequestHeader(key, value)

	req.Send()
	
	while (req.readyState != 4)
		sleep, 10
	
	if (t_b = "t")
		return req.ResponseText
	else
		return req.ResponseBody
}

CanviarHK()
{
	global hk, esperant, fitxer_cfg

	hk_ant := hk
	Hotkey, %hk%,, off
	CanviHK()
	while (esperant)
		Sleep, 10
	Hotkey, IfWinActive, ahk_class CabinetWClass
	Hotkey, %hk%, Accio, on
	if (hk_ant != hk)
		IniWrite, %hk%, %fitxer_cfg%, Parametres, hotkey
}

SetCursors(cursor)
{
	CursorHandle := DllCall( "LoadCursor", Uint,0, Int, cursor)
	Cursors = 32512,32513,32514,32515,32516,32640,32641,32642,32643,32644,32645,32646,32648,32649,32650,32651
	Loop, Parse, Cursors, `,
		DllCall( "SetSystemCursor", Uint,CursorHandle, Int, A_Loopfield)
}

RestoreCursors() 
{
	SPI_SETCURSORS := 0x57
	DllCall( "SystemParametersInfo", UInt,SPI_SETCURSORS, UInt,0, UInt,0, UInt,0 )
}

CarregaFitxer(ByRef taula, fitxer)
{
	ifExist, %fitxer%
	{
		Loop, read, %fitxer%
		{
			Loop, parse, A_LoopReadLine, `t
			{
				if (A_Index = 1)
					key := A_LoopField
				else
				{
					value := A_LoopField
					taula[key] := value
				}
			}
		}
	}
}

DesaFitxer(taula, fitxer)
{
	FileDelete, %fitxer%
	for key, value in taula
		FileAppend, %key%`t%value%`n, %fitxer%
}

CarregaINI()
{
	global hk, fitxer_cfg
	
	IfExist, %fitxer_cfg%
		IniRead, hk, %fitxer_cfg%, Parametres, hotkey
	else
		hk = #f ;per defecte el hotkey és win+f
}

ShowArrAso(arr)
{
	str := "["
	for k, v in arr
	{
		str .= "(" . k . "," . v . "),"
	}
	str := RTrim(str,",")
	str .= "]"
	return str
}

Sortir:
	DesaFitxer(tIDs, fitxer_ids)
	DesaFitxer(tNoms, fitxer_noms)
	tIDs := ""
	tNoms := ""
	t_epis := ""
	shell := ""
	tPaths := ""
	tFixes := ""
	tErrors := ""
	ExitApp
Return
