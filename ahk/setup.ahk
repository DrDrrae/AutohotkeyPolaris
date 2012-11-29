#NoEnv
#SingleInstance force
keyDelay := 25
i := 0
OutputVar := ""
ItemRecordIterations := 10
controlDelay := 40
directReaderOn := ""
checkDirectReader := ""
StartTime := ""
EndTime := ""
SetupMsgBox := ""
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%

_Library=Arbutus
settings = z:/settings.ini


IniRead, CCCN, %settings%, %_Library%, CatalogChangesCardNumber, %A_Space%
IniRead, HSCN, %settings%, %_Library%, HoldShelfCardNumber, %A_Space%
IniRead, PMCN, %settings%, %_Library%, ProblemMaterialCardNumber, %A_Space%

SetupRetry:
Gui, Setup:Add, Text, x12 y10 w380 h30 , Setup the various cards for %_Library%
Gui, Setup:Add, Text, x12 y33 w380 h20 , Catalog Changes:
Gui, Setup:Add, Edit, x100 y33 w292 h20 Limit14 vCCCN Number, %CCCN%
Gui, Setup:Add, Text, x12 y58 w380 h20 , Hold Shelf:
Gui, Setup:Add, Edit, x100 y58 w292 h20 Limit14 vHSCN Number, %HSCN%
Gui, Setup:Add, Text, x12 y83 w380 h20 , Problem Material:
Gui, Setup:Add, Edit, x100 y83 w292 h20 Limit14 vPMCN Number, %PMCN%
Gui, Setup:Add, Button, x12 y110 w90 h30 Default gSetupButtonOK, &OK
Gui, Setup:Add, Button, x157 y110 w90 h30 Disabled gSetupButtonDone , &Done
Gui, Setup:Add, Button, x302 y110 w90 h30 gSetupButtonCancel, &Cancel
Gui, Setup:Show, x127 y87 h160 w404, Pick a Library
Return


SetupButtonCancel:
SetupGuiClose:
SetupGuiEscape:
	Gui, Setup:Destroy
	MsgBox, 4, Setup Canceled. , Setup Canceled.  Do you wish to continue?
	IfMsgBox Yes
	{
		Goto, LibraryButtonOk
	}
	else
	{
		ExitApp
	}
	
SetupButtonOK:
SetupButtonDone:
	Gui, Setup:Submit
	Gui, Setup:Destroy
		StringLen, CCCNLength, CCCN
		StringMid, CCCNFirst, CCCN, 1 ,5
		StringLen, HSCNLength, HSCN
		StringMid, HSCNFirst, HSCN, 1 ,5
		StringLen, PMCNLength, PMCN
		StringMid, PMCNFirst, PMCN, 1 ,5
		
		if (PMCNLength != 14 or  PMCNFirst != 21183)
		{
			SetupMsgBox = Problem Material card number incorrect.`r`n
		}
		if (HSCNLength != 14 or  HSCNFirst != 21183)
		{
			SetupMsgBox = Hold Shelf card number incorrect.`r`n%SetupMsgBox%
		}
		if (CCCNLength != 14 or  CCCNFirst != 21183)
		{
			SetupMsgBox = Catalog Changes card number incorrect.`r`n%SetupMsgBox%
		}
		if (SetupMsgBox != "")
		{
			MsgBox, %SetupMsgBox%
			SetupMsgBox := ""
			Goto, SetupRetry
		}

	IniWrite, %CCCN%, %settings%, %_Library%, CatalogChangesCardNumber
	IniWrite, %HSCN%, %settings%, %_Library%, HoldShelfCardNumber
	IniWrite, %PMCN%, %settings%, %_Library%, ProblemMaterialCardNumber

LibraryButtonOk:
	MsgBox Library Button Okay
	ExitApp