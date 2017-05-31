#Persistent
#SingleInstance ignore
#Include %A_ScriptDIR%\MostrarGUI.ahk
#Include %A_ScriptDIR%\CanviHK.ahk

shell := ComObjCreate("Shell.Application")
fitxer_ids := "series.txt"
fitxer_cfg := "config.ini"
t_IDs := Object()
t_epis := Object()
tPaths := Object()
tFixes := Object()
tNoDir := Object()
tErrors := Object()

Menu, Tray, NoStandard
Menu, Tray, Add, Canviar hotkey, CanviarHK
Menu, Tray, Add
Menu, Tray, Add, Sortir, Sortir
Menu, Tray, Tip, FixName

CarregaIDs()
CarregaINI()
Hotkey, IfWinActive, ahk_class CabinetWClass
Hotkey, %hk%, Accio
return

Accio()
{
	global shell, tPaths, tFixes, tNoDir, tErrors
	
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
		tNoDir := Object()
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
	nom_serie := RealName(nom_serie)
	nom_epi := BuscaAPITVMaze(nom_serie, temp, epi)
	if (SubStr(nom_epi, 1, 6) = "ERROR.")
		return %nom_epi%
	valor = %nom_serie% %temp%x%epi% %nom_epi%.%ext%
	return %valor%
}

BuscaAPITVmaze(ByRef nom, se, ep)
{
	global t_IDs, t_epis
	
	nom_ep := ""
	
	value := t_IDs[nom]
	if (value = "")
	{
		url_id := "https://api.tvmaze.com/singlesearch/shows?q=" . nom
		
		headers := {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:51.0) Gecko/20100101 Firefox/51.0", "Access-Control-Allow-Origin": "*"}
		result := ObtePagina(url_id, headers, "t")
			 
		serie_ID := GetJSON(result, """id"":")
		nom := GetJSON(result, """name"":")
		StringUpper, nom, nom, T
		;si ha trobat la sèrie guarda l'ID per futurs usos
		if (serie_ID != "")
			t_IDs[nom] := serie_ID
	}
	else
		serie_ID := t_IDs[nom]
	
	if (serie_ID != "")
	{
		val = %serie_id%º%se%º%ep%
		value := t_epis[val]
		if (value = "")
		{
			url_nom := "https://api.tvmaze.com/shows/" . serie_ID . "/episodebynumber?season=" . se . "&number=" . ep
			
			result := ObtePagina(url_nom, headers, "t")
				
			message := 	GetJSON(result, """message"":")
			if (message != "Page not found." and message != "Unknown episode")
			{
				nom_ep := GetJSON(result, """name"":")
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
	}
	else
		nom_ep := "ERROR.Sèrie desconeguda"
	headers := ""
	return nom_ep
}

CarregaIDs()
{
	global fitxer_ids, t_IDs
	
	ifExist, %fitxer_ids%
	{
		Loop, read, %fitxer_ids%
		{
			Loop, parse, A_LoopReadLine, `t
			{
				if (A_Index = 1)
					key := A_LoopField
				else
				{
					value := A_LoopField
					t_IDs[key] := value
				}
			}
		}
	}
}

DesaIDs()
{
	global fitxer_ids, t_IDs
	
	FileDelete, %fitxer_ids%
	for key, value in t_IDs
		FileAppend, %key%`t%value%`n, %fitxer_ids%
}

GetJSON(txt, clau)
{
	pos_ini := InStr(txt, clau)
	pos_ini += StrLen(clau)
	txt := SubStr(txt, pos_ini)
	pos_fin := InStr(txt, ",")
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

RealName(nom)
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

CarregaINI()
{
	global hk, fitxer_cfg
	
	IfExist, %fitxer_cfg%
		IniRead, hk, %fitxer_cfg%, Parametres, hotkey
	else
		hk = #f ;per defecte el hotkey és win+f
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

Sortir:
	DesaIDs()
	t_IDs := ""
	t_epis := ""
	shell := ""
	tPaths := ""
	tFixes := ""
	tNoDir := ""
	tErrors := ""
	ExitApp
Return
