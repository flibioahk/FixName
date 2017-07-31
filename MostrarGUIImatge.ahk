MostraOpcions(URL)
{
	static WB
	global op_res := 2
	
	Gui, New, -Caption +Border
	Gui, Add, ActiveX, x0 y0 w250 h340 vWB, Shell.Explorer
	Gui, Add, Text,x76, És aquesta serie?
	Gui, Add, Button, w40,Si
	Gui, Add, Button, w40 xp+40, No
	WB.Navigate(URL)
	Gui, Show, w232
	
	while (op_res = 2)
	{
		Sleep 10
	}
	
	Gui, Destroy
	return %op_res%
}

ButtonSi()
{
	global op_res := 1
}

ButtonNo()
{
	global op_res := 0
}
