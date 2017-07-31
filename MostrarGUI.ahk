#Include Class_LV_Colors.ahk

tNoDir := Object()

MostrarGUI(num_elem, num_errors := 0)
{
	global tFixes, tNoDir, tErrors, CLV1, CLV2, HLV1, exp_hwnd
	
	Gui, -Resize -MaximizeBox -MinimizeBox +Owner%exp_hwnd%
	Gui, Margin, 10, 10
	if (num_elem > 0)
	{
		avis = 
(	
Es procedirà a modificar el nom dels següents fitxers. S'ha de tenir en compte que:
  - Si el nom destí té algun caràcter no permés (\/?*""<>|:) es substituirà
    per '_' i s'afegira [warning] al final del nom.
  - Només es processaran els fitxers marcats.
)
		min_elem := 7
		if (num_elem > min_elem)
		{
			lv_elems := min_elem
			pixw := 339
		}
		else
		{
			lv_elems := num_elem
			pixw := 348
		}
	
		Gui, Add, Text, xm, %avis%
		Gui, Add, ListView, xm r%lv_elems% w700 Grid Checked -Multi NoSortHDR Section ReadOnly hwndHLV1 LV0x400 LV0x100, Nom Antic|Nom Nou

		CLV1 := New LV_Colors(HLV1)
		CLV1.Critical := 100

		LV_ModifyCol(1, pixw)
		LV_ModifyCol(2, pixw)

		i:= 1
		for key,value in tFixes
		{
			SplitPath, key, outfilek
			SplitPath, value, outfilev
		
			tNoDir[outfilek] := key
			LV_Add("Check", outfilek, outfilev)
			if (InStr(outfilev, "[warning]"))
				CLV1.Cell(i, 2,, 0xFFBB00)
			else
				CLV1.Cell(i, 2,, 0x00BF00)
			i++
		}
	}
	if (num_errors > 0)
	{
		min_elem := 7
		if (num_errors > min_elem)
		{
			lv_elems := min_elem
			pixw := 339
		}
		else
		{
			lv_elems := num_errors
			pixw := 348
		}

	
		Gui, Add, Text,, Llista d'errors:
		Gui, Add, ListView, xm r%lv_elems% w700 Grid -Multi NoSortHDR Section ReadOnly hwndHLV2 LV0x400 LV0x100, Fitxer|Error
		
		CLV2 := New LV_Colors(HLV2)
		CLV2.Critical := 100
		
		LV_ModifyCol(1, pixw)
		LV_ModifyCol(2, pixw)
		
		i := 1
		for key,value in tErrors
		{
			SplitPath, key, outfilek
			LV_Add("", outfilek, value)
			CLV2.Cell(i, 2,, 0xFF0000)
			i++
		}
	}
	if (num_elem = 0)
		Gui, Add, Button, W80 H20 xs+310 gAcabarExecucio, Sortir
	else
	{
		Gui, Add, Button, W80 H20 xs+200, Processar
		Gui, Add, Button, W80 H20 yp xs+420 gAcabarExecucio, Sortir
	}
	Gui, Show,, FixName
	WinSet, Redraw,, ahk_id %HLV1%
	WinSet, Redraw,, ahk_id %HLV2%
}

ButtonProcessar()
{
	global HLV1, tFixes, tNoDir
	
	fila := 0  ; This causes the first loop iteration to start the search at the top of the list.
	Gui, ListView, %HLV1%
	Loop
	{
		fila := LV_GetNext(fila, "C")  ; Resume the search at the row after that found by the previous iteration.
		if not fila
			break
		LV_GetText(key, fila)
		antic := tNoDir[key]
		nou := tFixes[antic]
		FileMove, %antic%, %nou%
	}
	AcabarExecucio()
}

GuiClose()
{
	AcabarExecucio()
}

AcabarExecucio()
{
	global tPaths, tFixes, tErrors, CLV1, CLV2, tIDs, fitxer_ids, tNoms, fitxer_noms
	
	Gui, Destroy
	DesaFitxer(tIDs, fitxer_ids)
	DesaFitxer(tNoms, fitxer_noms)
	tPaths := ""
	tFixes := ""
	tErrors := ""
	CLV1 := ""
	CLV2 := ""
}

