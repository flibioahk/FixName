CanviHK()
{
	global hk, RadioB, Tecla, esperant

	esperant := true
	Gui, New, -MaximizeBox -MinimizeBox, Canviar hotkey	
	Y := 0
	layers := ["control", "alt", "shift", "windows"]
	SplitHK(hk, layer1, layer2)
	Loop, 4
	{
		Y := (A_Index * 20)
		literal := layers[A_Index]
		text = x20 y%Y% w62 h20
		if (A_Index = 1)
			text = %text% vRadioB
		if (layer1 = A_Index)
			text = %text% Checked
		Gui, Add, Radio, %text%, %literal%
	}
	Gui, Font, s12
	Gui, Add, Text, x102 y50 w10 h20 +0x200, +
	Gui, Font
	Gui, Add, Text, x132 y50 w68 h20 +0x1000 +0x200 vTecla Limit1 gHK_EsperaTecla, %layer2%
	Gui, Add, Button, x20 y120 w80 h20 gHK_OK, OK
	Gui, Add, Button, x120 y120 w80 h20 gHK_Sortir, Sortir
	Gui, Show, w220 h160
	Return
}

HK_EsperaTecla()
{
	global Tecla
	
	GuiControl, +0x6, Tecla
	Input, tec, L1 E, {F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{Capslock}{Numlock}{PrintScreen}{Pause}

	if (tec = "")
	{
		If InStr(ErrorLevel, "EndKey:")
			tec := SubStr(ErrorLevel, 8)
	}

	tec := GetKeyName(tec)
	GuiControl, -0x6, Tecla
	GuiControl,, Tecla, %tec%
}

HK_OK()
{
	global hk, RadioB, Tecla, esperant
	
	lay_mini := ["^", "!", "+", "#"]
	Gui, submit, nohide
	h := lay_mini[RadioB]
	GuiControlGet, k,, Tecla
	hk := h . k
	Gui, Destroy
	esperant := false
}

HK_Sortir()
{
	global esperant
	
	Gui, Destroy
	esperant := false
}

SplitHK(hk, ByRef l1, ByRef l2)
{
	eles := {"^": 1, "!": 2, "+": 3, "#": 4}
	
	h := SubStr(hk, 1, 1)
	l1 := eles[h]
	l2 := SubStr(hk, 2)
}