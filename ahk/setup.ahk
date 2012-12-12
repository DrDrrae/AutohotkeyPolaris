#NoEnv
#SingleInstance force
keyDelay = 25
i = 0
OutputVar := ""
ItemRecordIterations = 10
controlDelay = 40
directReaderOn := ""
checkDirectReader := ""
StartTime := ""
EndTime := ""
SetupMsgBox := ""
checkboxEnabled := ""
checkboxDelete := ""
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%

checkINI(CCCN,HSCN,PMCN)
{
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
		Return, false
	}
	Return, true
}

;*** Files
;* Settings
;IfInString, var, SearchString
IfInString, A_ScriptDir, Z:\
{
	settings = Z:\settings.ini
	log = Z:\log.txt
}
else
{
	settings = ./settings.ini
	log = ./log.txt
}

Setup:

IniRead, librariesSetup, %settings%
StringReplace, librariesSetup, librariesSetup, Default`n
Sort, librariesSetup, CL

	Gui, Setup:Add, Text, x12 y10 w380 h30 , Enable or disable, add or remove libraries.
	Gui, Setup:Add, Text, x12 y33 w80 h20 , Enabled:
	Gui, Setup:Add, Text, x175 y33 w80 h20 , Delete:
	Gui, Setup:Add, Text, x57 y33 w80 h20 , Default:
	
	;Loop, Parse, InputVar, Delimiters
	y1=58
	y2=60
	height1 = 120
	height2 = 70
	firstCheckedVariable = vRadioLibraryGroup
	librariesDefaultArray := ""
	ArrayCount = 1
	librariesArray := []
	librariesArrayNS := []
	IniRead, librariesDefault, %settings%, Default, Library
	Loop, Parse, librariesSetup, `n
	{
		StringReplace, A_LoopFieldNoSpace, A_LoopField, %A_SPACE%,, All
		if (A_Index != 1)
		{
			firstCheckedVariable := ""
		}
		if (librariesDefault = A_LoopField)
		{
			Gui, Setup:Add, Radio, x57 y%y1% w110 h20 Checked %firstCheckedVariable%, %A_LoopField%
		}
		else
		{
			Gui, Setup:Add, Radio, x57 y%y1% w110 h20 %firstCheckedVariable%, %A_LoopField%
		}
		
		librariesArray.Insert(A_LoopField)
		librariesArrayNS.Insert(A_LoopFieldNoSpace)
		y1+=20
	}
	y1=58
	Loop, Parse, librariesSetup, `n
	{
		StringReplace, A_LoopFieldNoSpace, A_LoopField, %A_SPACE%,, All
		IniRead, librariesEnabled, %settings%, %A_LoopField%, Enabled
		if (librariesEnabled = 1)
		{
			Gui, Setup:Add, Checkbox, x37 y%y1% w20 h20 Checked -wrap vEnabled%A_LoopFieldNoSpace%,
		}
		else
		{
			Gui, Setup:Add, Checkbox, x37 y%y1% w20 h20 -wrap vEnabled%A_LoopFieldNoSpace%,
		}
		Gui, Setup:Add, Checkbox, x175 y%y1% w110 h20 -wrap vDelete%A_LoopFieldNoSpace%,
		
		y1+=20
		y2+=20
		height1+=20
		height2+=20
	}
	
	Gui, Setup:Add, Button, x12 y%height2% w90 h30 Default gSetupButtonOK, &OK
	Gui, Setup:Add, Button, x108 y%height2% w90 h30 gSetupButtonDone , &Done
	Gui, Setup:Add, Button, x202 y%height2% w90 h30 gSetupButtonAdd , &Add
	Gui, Setup:Add, Button, x297 y%height2% w90 h30 gSetupButtonCancel, &Cancel
	Gui, Setup:Show, x127 y87 h%height1% w404, Pick a Library
	Return
	
	SetupButtonCancel:
	SetupGuiClose:
	SetupGuiEscape:
		ExitApp
	
	SetupButtonAdd:
		Gui, Setup:Destroy
		
		Gui, SetupAdd:Add, Text, x12 y10 w380 h30 , Add a new library and its various cards.
		Gui, SetupAdd:Add, Text, x12 y33 w100 h20 , Library Name:
		Gui, SetupAdd:Add, Edit, x100 y33 w292 h20 Limit14 vLibraryName,
		Gui, SetupAdd:Add, Text, x12 y58 w100 h20 , Catalog Changes:
		Gui, SetupAdd:Add, Edit, x100 y58 w292 h20 Limit14 vCCCN Number,
		Gui, SetupAdd:Add, Text, x12 y83 w100 h20 , Hold Shelf:
		Gui, SetupAdd:Add, Edit, x100 y83 w292 h20 Limit14 vHSCN Number,
		Gui, SetupAdd:Add, Text, x12 y108 w100 h20 , Problem Material:
		Gui, SetupAdd:Add, Edit, x100 y108 w292 h20 Limit14 vPMCN Number,
		Gui, SetupAdd:Add, Checkbox, x12 y133 w100 h20 vLibraryEnabled, Enable?
		Gui, SetupAdd:Add, Button, x12 y158 w90 h30 Default gSetupAddButtonOK, &OK
		Gui, SetupAdd:Add, Button, x157 y158 w90 h30 Disabled gSetupAddButtonDone , &Done
		Gui, SetupAdd:Add, Button, x302 y158 w90 h30 gSetupAddButtonCancel, &Cancel
		Gui, SetupAdd:Show, x127 y87 h193 w404, Pick a Library
		Return

		SetupAddButtonDone:
		SetupAddButtonCancel:
		SetupAddGuiClose:
		SetupAddGuiEscape:
			Gui, SetupAdd:Destroy
			Goto, Setup
			
		SetupAddButtonOK:
			Gui, SetupAdd:Submit
			Gui, SetupAdd:Destroy
			MsgBox, Library: %LibraryName%`nCCCN: %CCCN%`nHSCN: %HSCN%`nPMCN: %PMCN%`nEnabled: %LibraryEnabled%
			
			checkINIReturn := checkINI(CCCN,HSCN,PMCN)
		
			if (checkINIReturn = true)
			{
				IniWrite, %CCCN%, %settings%, %LibraryName%, CatalogChangesCardNumber
				IniWrite, %HSCN%, %settings%, %LibraryName%, HoldShelfCardNumber
				IniWrite, %PMCN%, %settings%, %LibraryName%, ProblemMaterialCardNumber
				IniWrite, %LibraryEnabled%, %settings%, %LibraryName%, Enabled
				Goto, Setup
			}
			else
			{
				Goto, SetupButtonAdd
			}
			
			;IniWrite,
			;Goto, Setup
		
	SetupButtonOK:
		StringTrimLeft, checkboxEnabled, checkboxEnabled, 1
		StringTrimLeft, checkboxDelete, checkboxDelete, 1
		Gui, Setup:Submit
		Gui, Setup:Destroy
		;Loop % librariesDefaultArray.MaxIndex()
		;{
		;	MsgBox % librariesDefaultArray[A_Index]
		;}
		;MsgBox % librariesDefaultArray[RadioLibraryGroup]
		defaultLibrary := librariesArray[RadioLibraryGroup]
		IniWrite, %defaultLibrary%, %settings%, Default, Library

		Loop % librariesArrayNS.MaxIndex()
		{
			element := librariesArrayNS[A_Index]
			elementEnabled = Enabled%element%
			elementEnabledvalue := %elementEnabled%
			element2 := librariesArray[A_Index]
			;MsgBox, %elementvalue%
			IniWrite, %elementEnabledvalue%, %settings%, %element2%, Enabled
			
			elementDelete = Delete%element%
			elementDeletevalue := %elementDelete%
			if (elementDeleteValue = 1)
			{
				elementSpace := librariesArray[A_Index]
				MsgBox, 260, Delete %elementSpace%, Are you sure you want to delete %elementSpace%?
				ifMsgBox Yes
				{
					IniDelete, %settings%, %elementSpace%
				}
			}
		}
		Goto, Setup

SetupButtonDone:
		if (A_IsCompiled)
			Run, PolarisSetupAHK.exe
		else
			Run, Z:\AutoHotkey\AutoHotkeyU64.exe Z:\ahk\Polaris.ahk
		ExitApp