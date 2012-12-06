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


IniRead, librariesSetup, %settings%
StringReplace, librariesSetup, librariesSetup, Default`n

	Gui, Setup:Add, Text, x12 y10 w380 h30 , Enable or disable, add or remove libraries.
	Gui, Setup:Add, Text, x12 y33 w380 h20 , Enabled:
	Gui, Setup:Add, Text, x57 y33 w380 h20 , Delete:
	
	;Loop, Parse, InputVar, Delimiters
	y1=58
	y2=60
	height1 = 120
	height2 = 70
	Loop, Parse, librariesSetup, `n
	{
		StringReplace, A_LoopFieldNoSpace, A_LoopField, %A_SPACE%,, All
		IniRead, librariesEnabled, %settings%, %A_LoopField%, Enabled
		if (librariesEnabled = 1)
		{
			Gui, Setup:Add, Checkbox, x12 y%y1% w20 h20 Checked vEnabled%A_LoopFieldNoSpace%,
		}
		else
		{
			Gui, Setup:Add, Checkbox, x12 y%y1% w20 h20 vEnabled%A_LoopFieldNoSpace%,
		}
		Gui, Setup:Add, Checkbox, x57 y%y1% w20 h20 vDelete%A_LoopFieldNoSpace%,
		Gui, Setup:Add, Text, x85 y%y2% w200 h20 , %A_LoopField%
		checkboxEnabled = %checkboxEnabled%,%A_LoopFieldNoSpace%
		checkboxDelete = %checkboxDelete%,%A_LoopFieldNoSpace%
		y1+=20
		y2+=20
		height1+=20
		height2+=20
	}
	;Gui, Setup:Add, Checkbox, x12 y58 w20 h20 vEnabledArbutus,
	;Gui, Setup:Add, Checkbox, x57 y58 w20 h20 vDeleteArbutus,
	;Gui, Setup:Add, Text, x85 y60 w200 h20 , Arbutus
	;Gui, Setup:Add, Checkbox, x12 y78 w20 h20 vEnabledCatonsville,
	;Gui, Setup:Add, Checkbox, x57 y78 w20 h20 vDeleteCatonsville,
	;Gui, Setup:Add, Text, x85 y80 w200 h20 , Catonsville
	
	Gui, Setup:Add, Button, x12 y%height2% w90 h30 Default gSetupButtonOK, &OK
	Gui, Setup:Add, Button, x157 y%height2% w90 h30 Disabled gSetupButtonDone , &Done
	Gui, Setup:Add, Button, x302 y%height2% w90 h30 gSetupButtonCancel, &Cancel
	Gui, Setup:Show, x127 y87 h%height1% w404, Pick a Library
	Return
	
	SetupButtonDone:
	SetupButtonCancel:
		ExitApp
	
	SetupButtonOK:
		StringTrimLeft, checkboxEnabled, checkboxEnabled, 1
		StringTrimLeft, checkboxDelete, checkboxDelete, 1
		Gui, Setup:Destroy
		MsgBox, %checkboxEnabled%
		MsgBox, %checkboxDelete%
		ExitApp