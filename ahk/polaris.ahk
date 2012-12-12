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
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%


IfWinNotExist, Polaris ILS
{
	MsgBox, Polaris ILS isn't running.
	ExitApp
}

;********************************
;*
;* Fuctions
;*
;********************************

;********
;* Turns DirectReader on or off
;* 0 for OFF, 1 for ON, 2 exits Direct Reader
;********
controlDirectReader(a=0)
{
	if (a = 0)
	{
		b=OFF
	}
	else if (a = 1)
	{
		b=ON
	}

	IfWinExist, CircControl-DirectReader
	{
		WinGet MMX, MinMax, CircControl-DirectReader
		if (MMX = -1)
		{
			WinRestore, CircControl-DirectReader
		}
		if (a = 2)
		{
			WinActivate, CircControl-DirectReader
			Sleep, 250
			Send !x
			;return
		}
		else
		{
			; Try to control DirectReader up to 5 times with a 100ms delay between tries.
			; Sometimes it doesn't change on the first try.
			Loop
			{
				if (a_index > 5)
				{
					StringLower, b_lower, b
					MsgBox, Couldn't turn %b_lower% Directreader.  Please turn it off manually.
					break  ; Terminate the loop
				}
				ControlGet, checkDirectReader, Enabled, , %b%, CircControl-DirectReader
				if (checkDirectReader = 1)
				{
					SetControlDelay -1
					ControlClick, %b%, CircControl-DirectReader,,,, NA
					Sleep, 100
				}
				else
				{
					break
				}
			}
			SetControlDelay, 40
		}
	}
}

startHotkey(sl=true)
{
	KeyWait, Ctrl
	KeyWait, Alt
	KeyWait, LWin
	KeyWait, RWin
	if (sl = true)
	{
		scrollLockState := GetKeyState("Scrolllock", "T")
		if (scrollLockState != true)
		{
			Exit
		}
		else
		{
			Return, true
		}
	}
	else
	{
		Return, true
	}
}

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

breakScript(type = 0)
{
	if (type = npu)
		title = Item Records - Barcode Find Tool
	else if (type = missing)
		title = Request Manager - Hold Requests
	else if (type if integer)
		title = ahk_id %active_id%

	IfWinExist, %title%
	{
		WinClose
	}

	controlDirectReader(1)
	Progress, Off
	BlockInput, Off
	Gui, Holds:Destroy
	Gui, HldShlf:Destroy
	Gui, PrblmMtrl:Destroy
	Exit
}

log(text)
{
	FileAppend, %A_Now% - %text%`n, %log%
}

db_startTime()
{
	;global StartTime
	StartTime := A_TickCount

	return %StartTime%
}
db_endTime(StartTime)
{
	;global StartTime
	EndTime := A_TickCount
	ElapsedTime := (EndTime - StartTime)/1000
	MsgBox, Start Time: %StartTime%`nEnd Time: %EndTime%`n`nTotal Time: %ElapsedTime% seconds
}


;********************************
;*
;* Auto Execute
;*
;********************************

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

;********************************
;*
;* Upgrades
;* Probably should be put into another script
;* To be used whenever the format of the
;* INI files changes significantly and can't
;* be reasonbly worked around.  Usually
;* resulting a a complete refresh of the INI
;* file.
;*
;********************************
IniRead, Version, %settings%, Version, Version, 0
;* First version.  Format changed significantly.
;* Easiest to purge the old file and start anew.
if (Version > 0.01)
{
	FileInstall, Z:\settingsDefault.ini, %settings%, 1
}

FileInstall, Z:\settingsDefault.ini, %settings%, 0
Sleep, 500

IniRead, Libraries, %settings%
IniRead, DefaultLibrary, %settings%, Default, Library
StringReplace, Libraries, Libraries, Default`n
;Loop, Parse, InputVar [, Delimiters, OmitChars] 
libEnabled = 0
libraryCount = 0
Loop, Parse, Libraries, `n
{
	IniRead, libEnabled, %settings%, %A_LoopField%, Enabled, 0
	if (libEnabled = 1)
	{
		if (libraryCount > 0)
		{
			Libraries2 = %Libraries2%|%A_LoopField%
		}
		else
		{
			Libraries2 = %A_LoopField%	
		}
		libraryCount++
	}
}
Sort Libraries2, CL D|

libraryCountLabel:
if (libraryCount < 1)
{
	MsgBox, No Libraries Enabled.
	Goto, SetupRetry
}
	
if (libraryCount > 1)
{
	;if (InStr(Libraries2,DefaultLibrary)>0)
	ifInString, Libraries2, %DefaultLibrary%
	{
		StringReplace, Libraries2, Libraries2, %DefaultLibrary%, %DefaultLibrary%|
		DefaultLibrary2=%DefaultLibrary%||
		ifNotInString, Libraries2, %DefaultLibrary2%
			StringReplace, Libraries2, Libraries2, %DefaultLibrary%, %DefaultLibrary%|
	}
	else
	{
		StringReplace, Libraries2, Libraries2, |, ||
	}
	Gui, Library:Add, Text, x12 y10 w380 h30 , Pick a Baltimore County Public Library Branch
	Gui, Library:Add, Text, x12 y35 w380 h30 , The branches are pulled from the settings.ini file.
	Gui, Library:Add, Text, x12 y65 w380 h20 , Library:
	Gui, Library:Add, DropDownList, x12 y80 w320 Sort vLibrary, %Libraries2%
	Gui, Library:Add, Button, x340 y80 w40 h20 gLibraryButtonEdit, &Edit
	Gui, Library:Add, Button, x12 y110 w90 h30 Default gLibraryButtonOK, &OK
	Gui, Library:Add, Button, x157 y110 w90 h30 gLibraryButtonSetup , &Setup
	Gui, Library:Add, Button, x302 y110 w90 h30 gLibraryButtonCancel, &Cancel
	Gui, Library:Show, x127 y87 h160 w404, Pick a Library
	Return
}
else
{
	Library = %Libraries2%
	;MsgBox, Library: %Library%
	Goto, LibrarySkip
}

LibraryGuiClose:
LibraryGuiEscape:
LibraryButtonCancel:
	Gui, Library:Destroy
	ExitApp
	
LibraryButtonSetup:
	SetupRetry:
	if (A_IsCompiled)
		Run, PolarisSetupAHK.exe
	else
		Run, Z:\AutoHotkey\AutoHotkeyU64.exe Z:\ahk\Setup.ahk
	ExitApp

LibraryButtonEdit:
	Gui, Library:Submit
	Gui, Library:Destroy
	IniRead, CCCN, %settings%, %Library%, CatalogChangesCardNumber, %A_Space%
	IniRead, HSCN, %settings%, %Library%, HoldShelfCardNumber, %A_Space%
	IniRead, PMCN, %settings%, %Library%, ProblemMaterialCardNumber, %A_Space%

	EditRetry:
	Gui, Edit:Add, Text, x12 y10 w380 h30 , Edit the various cards for %Library%
	Gui, Edit:Add, Text, x12 y33 w380 h20 , Catalog Changes:
	Gui, Edit:Add, Edit, x100 y33 w292 h20 Limit14 vCCCN Number, %CCCN%
	Gui, Edit:Add, Text, x12 y58 w380 h20 , Hold Shelf:
	Gui, Edit:Add, Edit, x100 y58 w292 h20 Limit14 vHSCN Number, %HSCN%
	Gui, Edit:Add, Text, x12 y83 w380 h20 , Problem Material:
	Gui, Edit:Add, Edit, x100 y83 w292 h20 Limit14 vPMCN Number, %PMCN%
	Gui, Edit:Add, Button, x12 y110 w90 h30 Default gEditButtonOK, &OK
	Gui, Edit:Add, Button, x157 y110 w90 h30 Disabled gEditButtonDone , &Done
	Gui, Edit:Add, Button, x302 y110 w90 h30 gEditButtonCancel, &Cancel
	Gui, Edit:Show, x127 y87 h160 w404, Pick a Library
	Return


	EditButtonCancel:
	EditGuiClose:
	EditGuiEscape:
		Gui, Edit:Destroy
		Goto, libraryCountLabel
		
	EditButtonOK:
	EditButtonDone:
		Gui, Edit:Submit
		Gui, Edit:Destroy
		
		checkINIReturn := checkINI(CCCN,HSCN,PMCN)
		
		if (checkINIReturn = true)
		{
			IniWrite, %CCCN%, %settings%, %Library%, CatalogChangesCardNumber
			IniWrite, %HSCN%, %settings%, %Library%, HoldShelfCardNumber
			IniWrite, %PMCN%, %settings%, %Library%, ProblemMaterialCardNumber
			Goto, LibrarySkip
		}
		else
		{
			GoTo, EditRetry
		}

LibraryButtonOK:
	Gui, Library:Submit
	Gui, Library:Destroy


LibrarySkip:
	
	; @todone: setup.
	IniRead, CCCN, %settings%, %Library%, CatalogChangesCardNumber, %A_Space%
	IniRead, HSCN, %settings%, %Library%, HoldShelfCardNumber, %A_Space%
	IniRead, PMCN, %settings%, %Library%, ProblemMaterialCardNumber, %A_Space%

		checkINIReturn := checkINI(CCCN,HSCN,PMCN)
		
		if (checkINIReturn = false)
		{
			GoTo, SetupRetry
		}
	
	
	
;********************************
;*
;* # Windows Key
;* ! Alt
;* ^ Control
;* + Shift
;*
;********************************



;********************************
;
; HotKeys
;
;********************************

;** Win + F2
;** Check In - Bulk Mode
#F2::
	startHotkey(false)
	SetTitleMatchMode, RegEx
	WinWait, Check In - .* - Polaris, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		WinActivate, Polaris ILS
		Sleep 50
		Send {F2}
	}
	else
		WinActivate, Check In - .* - Polaris
	Sleep 50
	Send !vb
	SetTitleMatchMode, 1
Return

;** Win + F3
;** Check Out - Bulk Mode
#F3::
	startHotkey(false)
	WinWait, Check Out - 0 - Normal - Polaris, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		WinActivate, Polaris ILS
		Sleep 50
		Send {F3}
	}
	else
		WinActivate, Check Out - 0 - Normal - Polaris
Return

;** Win + F4
;** Request Manager
#F4::
	startHotkey(false)
	WinWait, Request Manager - Hold Requests, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		WinActivate, Polaris ILS
		Sleep 50
		Send !rr
	}
	else
	{
		WinClose, Request Manager - Hold Requests
		WinActivate, Polaris ILS
		Sleep 50
		Send !rr
	}
Return

;** Win + F6
;** Patron Status
#F6::
	startHotkey(false)
	WinActivate, Polaris ILS
	Sleep 50
	Send {F6}
Return

;** Win + F9
;** Patron Status
#F9::
	startHotkey(false)
	WinActivate, Polaris ILS
	Sleep 50
	Send {F9}
Return

;** Control + F9
;** Item Records
^F9::
	startHotkey(false)
	WinActivate, Polaris ILS
	Sleep 50
	Send ^!{F9}
Return

;********************************
;
; Turn ITG CircControl on or off
; Control + Alt + o
;
;********************************
^!o::
{
	startHotkey()
	IfWinExist, CircControl-DirectReader
	{
		ControlGet, checkDirectReader, Enabled, , ON, CircControl-DirectReader
		if (checkDirectReader = 1)
		{
			;MsgBox, DirectReader off
			controlDirectReader(1)
		}
		else
		{
			;MsgBox, DirectReader on
			controlDirectReader(0)
		}
		return
	}
	return
}

;********************************
;
; Close ITG CircControl, Open ITG TagFast
; or vice versa
; Control + Alt + d
;
;********************************
^!d::
{
	startHotkey()
	start_time := db_startTime()
	IfWinExist, CircControl-DirectReader
	{
		BlockInput On
		Progress, m2 b fs28 zh0, Switching to ITG TagFast, Switching to ITG TagFast.
		controlDirectReader(2)
		WinWaitClose, CircControl-DirectReader, , 5
		Sleep, 500
		Run, VernTag.exe, C:\Program Files (x86)\Integrated Technology Group\Apex TagFast\
		WinWait, ITG TagFast, , 10
		if ErrorLevel
		{
			MsgBox, ITG TagFast didn't launch.
			breakscript()
		}
	}
	else IfWinExist, ITG TagFast
	{
		BlockInput On
		Progress, m2 b fs28 zh0, Switching to ITG CircControl, Switching to ITG CircControl.
		WinActivate
		Sleep, 250
		IfWinActive, ITG TagFast, Primary Item Identification
		{
			Send !{F4}
			Sleep, 250
		}
		Send !{F4}
		WinWaitClose, ITG TagFast, , 5
		Sleep, 500
		Run, VSIPStaff.exe, C:\Program Files (x86)\Integrated Technology Group\Apex CircControl\
		WinWait, CircControl-DirectReader, , 10
		if ErrorLevel
		{
			MsgBox, ITG CircControl didn't launch.
			breakscript()
		}
	}
	else
	{
		Progress, m2 b fs28 zh0, Starting ITG CircControl, Starting ITG CircControl.
		BlockInput, On
		Run, VSIPStaff.exe, C:\Program Files (x86)\Integrated Technology Group\Apex CircControl\
		WinWait, CircControl-DirectReader, , 10
	}
	Progress, Off
	BlockInput Off
	db_endTime(start_time)
	Return
}

;********************************
;
; Not Picked Up Holds
; Control + Alt + n
;
;********************************

^!n::
{
	startHotkey()
	;MsgBox, 1 , Not Picked Up Holds, Automatically place the $1.00 NPU fee on the patron's card.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	;IfMsgBox Cancel
	;{
		; Turn DirectReader On
		;controlDirectReader(1)
		; Exit
		;Exit
	;}
	
	itemsprocessed = 0

	; Turn Direct Reader Off
	controlDirectReader(0)

	NPULabel:

	; If Item Records window isn't open, open it
	SetTitleMatchMode, RegEx
	WinWait, Item Records - .* Find Tool, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		WinActivate, Polaris ILS
		Sleep 50
		Send ^!{F9}
		WinGet, active_id, ID, Item Records - Barcode Find Tool
	}

	; Activate Item Records window and focus on the barcode input box
	; @todo: make sure Item Records is expecting item barcode
	; 	doesn't work too well

	WinActivate, Item Records - .* Find Tool
	;ControlFocus, ComboBox2, Item Records - .* Find Tool
	;Sleep, 200
	;Send ba
	ControlFocus, Edit3, Item Records - .* Find Tool
	SetTitleMatchMode, 1

	; Display InputBox
	;********************************
	; GUI - Holds
	;********************************
	Gui, Holds:Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically place the $1.00 "Not Picked Up" fee on the patron's card.
	Gui, Holds:Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, Holds:Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, Holds:Add, Button, x12 y110 w90 h30 Default gHldButtonOK, &OK
	Gui, Holds:Add, Button, x157 y110 w90 h30 gHldButtonDone , &Done
	Gui, Holds:Add, Button, x302 y110 w90 h30 gHldButtonCancel, &Cancel
	Gui, Holds:Show, x127 y87 h160 w404, Not Picked Up
	Return

	; If the GUI is closed, escaped (escape key),
	; or the Done or Cancel button is pressed,
	; break out of the script.
	HoldsGuiClose:
	HoldsGuiEscape:
	HldButtonDone:
	HldButtonCancel:
		Gui, Holds:Destroy
		if (itemsprocessed > 0)
		{
			SetTitleMatchMode, RegEx
			WinWait, Check In - .* - Polaris, ,0
			if ErrorLevel
			{
				ErrorLevel = 0
				WinActivate, Polaris ILS
				Sleep 50
				Send {F2}
			}
			else
				WinActivate, Check In - .* - Polaris
			Sleep 50
			Send !vb
			SetTitleMatchMode, 1
		}
		breakScript(%active_id%)

	; When OK button is clicked
	HldButtonOK:
		;startTime()
		Gui, Holds:Submit  ; Save the input from the user to each control's associated variable.
		Gui, Holds:Destroy
		; Check length of input and the first five digits
		; Should be 14 numbers long and start with 31183
		; If not, alert the user and retry.
		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183)
		{
			BlockInput, On
			ControlFocus, Edit3, Item Records - Barcode Find Tool
			SetkeyDelay, 0
			Send %ItemBarcode%{Enter}
			SetkeyDelay, %keyDelay%
			; Waits x times every 250ms (one quarter second)
			; ItemRecordIterations = 10 (at top of script)
			; until the item gets returned by Polaris.
			; If item is never returned either due to lag
			; or the item simply doesn't exist, alert the
			; user and break the script.
			i = 0
			retryNPU:
			ControlGet, OutputVar, List, Selected Col9, SysListView321, Item Records
				if (ErrorLevel or OutputVar = "")
				{
					if (i < ItemRecordIterations)
					{
						i++
						Sleep, 250
						Goto, retryNPU
					}
					else
					{
						textToLog = Error Level: %ErrorLevel% Output Var: %OutputVar% i: %i%
						log(textToLog)
						MsgBox Error Level: %ErrorLevel%`r`nOutput Var: %OutputVar%`r`ni: %i%
						breakScript()
					}
				}
				if (%ItemBarcode% != %OutputVar%)
				{
					textToLog = Barcode doesn't match or item doesn't exist %ItemBarcode%
					log(textToLog)
					MsgBox, Barcode doesn't match or item doesn't exist.
					Goto, NPULabel
				}
			Sleep, 100
			Send {AppsKey}le
			; Waits up to three (3) seconds for the Patron Status window to launch and be active
			; If not, alerts the user and breaks the script
			Sleep, 100
			WinWaitActive, Patron Status, ,3
			if ErrorLevel
			{
				ErrorLevel = 0
				MsgBox, Patron Status windows didn't launch.
				breakScript()
			}
			Send {Alt}va^c
			; Waits up to three (3) seconds for the Charge window to launch and be active
			; If not, alerts the user and breaks the script
			WinWaitActive, Charge, ,3
			if ErrorLevel
			{
				ErrorLevel = 0
				MsgBox, Charge window didn't launch.
				breakScript()
			}
			; Sets input focus to the Charge window
			; and adds all the pertinant information
			; ($1.00, barcode, NPU type)
			; Doesn't currently finish the transaction
			; to make sure everything works correctly.
			ControlFocus, Edit1, Charge
			Send 1{Tab}n{tab}
			SetkeyDelay, 0
			Send %ItemBarcode%
			SetkeyDelay, %keyDelay%
			Send {Tab 3}
			itemsprocessed++
			;Send %ItemBarcode%{Enter}{AppsKey}le{Alt}va^c1{Tab}n{tab}%ItemBarcode%{Tab 3}
			;Send %ItemBarcode%{Enter}{AppsKey}le{Alt}va^c1{Tab}n{tab}%ItemBarcode%{Tab 3}{Enter}
			BlockInput, Off
			;endTime()
			; Wait until the Charge and Patron Status window is closed, then restart the script.
			WinWaitClose, Charge
			WinWaitClose, Patron Status
			Goto, NPULabel
		}
		else
		{
			; Input isn't correct.  Either the input isn't the correct length or it doesn't start with the right digits.
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
			Goto, NPULabel
		}
	Return
}

;********************************
;
; Unavailable
; Control + Alt + u
;
;********************************

^!u::
{
	startHotkey()
	;MsgBox, 1 , Unavailable, Automatically mark items as unavailable.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	;IfMsgBox Cancel
	;{
		; Turn DirectReader On
		;controlDirectReader(1)
		; Exit
		;Exit
	;}

	; Turn Direct Reader Off
	controlDirectReader(0)

	UnvlblLabel:

	; If Item Records window isn't open, open it
	SetTitleMatchMode, RegEx
	WinWait, Item Records - .* Find Tool, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		WinActivate, Polaris ILS
		Sleep 50
		Send ^!{F9}
		WinGet, active_id, ID, Item Records - Barcode Find Tool
	}

	; Activate Item Records window and focus on the barcode input box
	; @todo: make sure Item Records is expecting item barcode
	; 	doesn't work too well

	WinActivate, Item Records - .* Find Tool
	;ControlFocus, ComboBox2, Item Records - .* Find Tool
	;Sleep, 200
	;Send ba
	ControlFocus, Edit3, Item Records - .* Find Tool
	SetTitleMatchMode, 1

	; Display InputBox
	;********************************
	; GUI - Unvlbl
	;********************************
	Gui, Unvlbl:Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically mark the item as unavailable.
	Gui, Unvlbl:Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, Unvlbl:Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, Unvlbl:Add, Button, x12 y110 w90 h30 Default gUnvlblButtonOK, &OK
	Gui, Unvlbl:Add, Button, x157 y110 w90 h30 gUnvlblButtonDone , &Done
	Gui, Unvlbl:Add, Button, x302 y110 w90 h30 gUnvlblButtonCancel, &Cancel
	Gui, Unvlbl:Show, x127 y87 h160 w404, Not Picked Up
	Return

	; If the GUI is closed, escaped (escape key),
	; or the Done or Cancel button is pressed,
	; break out of the script.
	UnvlblGuiClose:
	UnvlblGuiEscape:
	UnvlblButtonDone:
	UnvlblButtonCancel:
		Gui, Unvlbl:Destroy
		breakScript(%active_id%)

	; When OK button is clicked
	UnvlblButtonOK:
		;startTime()
		Gui, Unvlbl:Submit  ; Save the input from the user to each control's associated variable.
		Gui, Unvlbl:Destroy
		; Check length of input and the first five digits
		; Should be 14 numbers long and start with 31183
		; If not, alert the user and retry.
		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183)
		{
			BlockInput, On
			ControlFocus, Edit3, Item Records - Barcode Find Tool
			SetkeyDelay, 0
			Send %ItemBarcode%{Enter}
			SetkeyDelay, %keyDelay%
			; Waits x times every 250ms (one quarter second)
			; ItemRecordIterations = 10 (at top of script)
			; until the item gets returned by Polaris.
			; If item is never returned either due to lag
			; or the item simply doesn't exist, alert the
			; user and break the script.
			i = 0
			retryUnvlbl:
			ControlGet, OutputVar, List, Selected Col9, SysListView321, Item Records - Barcode Find Tool
				if (ErrorLevel or OutputVar = "")
				{
					if (i < ItemRecordIterations)
					{
						i++
						Sleep, 250
						Goto, retryUnvlbl
					}
					else
					{
						textToLog = Error Level: %ErrorLevel% Output Var: %OutputVar% i: %i%
						log(textToLog)
						MsgBox Error Level: %ErrorLevel%`r`nOutput Var: %OutputVar%`r`ni: %i%
						breakScript()
					}
				}
				if (%ItemBarcode% != %OutputVar%)
				{
					textToLog = Barcode doesn't match or item doesn't exist %ItemBarcode%
					log(textToLog)
					MsgBox, Barcode doesn't match or item doesn't exist.
					Goto, UnvlblLabel
				}
			Sleep, 100
			Send {Enter}
			; Waits up to three (3) seconds for the Patron Status window to launch and be active
			; If not, alerts the user and breaks the script
			SetTitleMatchMode, RegEx
			Sleep, 100
			WinWaitActive, Item Record .* - Circulation -  Polaris, ,3
			if ErrorLevel
			{
				ErrorLevel = 0
				MsgBox, Patron Status windows didn't launch.
				breakScript()
			}
			ControlFocus, ComboBox5, Item Record .* - Circulation -  Polaris
			Sleep, 100
			Send u
			Send ^s
			;Send !{F4}

			; Wait until the Item Record window is closed, then restart the script.
			WinWaitClose, Item Record .* - Circulation -  Polaris
			SetTitleMatchMode, 1
			Goto, UnvlblLabel
		}
		else
		{
			; Input isn't correct.  Either the input isn't the correct length or it doesn't start with the right digits.
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
			Goto, UnvlblLabel
		}
	Return
}


;********************************
;
; Request Manager - Missing
; Control + Alt + m
;
;********************************

^!m::
{
	startHotkey()
	;MsgBox, 1 , Missing Request Manager Items, Mark Request Manager items as missing.`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	;IfMsgBox Cancel
	;{
		; Turn DirectReader On
		;controlDirectReader(1)
		; Exit
		;Exit
	;}

	WinWait, Request Manager - Hold Requests, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		MsgBox, Request Manager window isn't open.
		WinActivate, Polaris ILS
		Sleep 50
		Send {Alt}rr
		Sleep, 100
		;Exit
	}

	BlockInput, On

	ControlGet, OutputVar, List, Col9 , SysListView321, Request Manager - Hold Requests

	if (OutputVar == "")
	{
		MsgBox, There are no items to be marked as missing.
		breakScript()
	}

	MsgBox, 260, Mark items as missing?, The following barcodes will be marked as missing.  Do you wish to continue?`n`n%OutputVar%
	IfMsgBox No
	{
		Exit
	}

	MissingLabel:
	Loop, Parse, OutputVar, `n  ; Rows are delimited by linefeeds (`n).
	{

		ItemBarcode = %A_LoopField%
		; If Item Records window isn't open, open it
		WinWait, Item Records - Barcode Find Tool, ,0
		if ErrorLevel
		{
			ErrorLevel = 0
			WinActivate, Polaris ILS
			Sleep 50
			Send ^!{F9}
		}

		; Activate Item Records window and focus on the barcode input box
		; @todo: make sure Item Records is expecting item barcode
		WinActivate, Item Records - Barcode Find Tool
		ControlFocus, Edit3, Item Records - Barcode Find Tool

		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183)
		{
			BlockInput, On
			ControlFocus, Edit3, Item Records - Barcode Find Tool
			SetkeyDelay, 0
			Send %ItemBarcode%{Enter}
			SetkeyDelay, %keyDelay%
			; Waits x times every 250ms (one quarter second)
			; ItemRecordIterations = 10 (at top of script)
			; until the item gets returned by Polaris
			retryMissing:
			ControlGet, OutputVar, List, Selected Col9, SysListView321, Item Records
				if (ErrorLevel or OutputVar = "")
				{
					if (i < ItemRecordIterations)
					{
						i++
						Sleep, 250
						Goto, retryMissing
					}
					else
					{
						MsgBox Error Level: %ErrorLevel%`r`nOutput Var: %OutputVar%`r`ni: %i%
						breakScript()
					}
				}
				if (%ItemBarcode% != %OutputVar%)
				{
					MsgBox, Barcode doesn't match or item doesn't exist.
					Goto, MissingLabel
				}
			Sleep, 100

			Send {Enter}

			WinWaitActive, Item Record, Barcode:
			Sleep, 100
			ControlFocus, ComboBox5, Item Record, Barcode:

			Send m
			;Send ^s ;Save changed information
			;Send !{F4} ;Close window
		}

		WinWaitClose, Item Record, Barcode:
	}
	Return
}

;********************************
;
; Check Item out to the Hold Shelf
; Control + Alt + h
;
;********************************

^!h::
{
	startHotkey()
	IniRead, HSCN, %settings%, %Library%, HoldShelfCardNumber, %A_Space%
	if (HSCN="")
	{
		;********************************
		; Hold Shelf GUI
		;********************************
		Gui, HldCrd:Add, Text, x12 y10 w380 h30 , Hold Shelf Card Number
		Gui, HldCrd:Add, Text, x12 y50 w380 h20 , Card Number:
		Gui, HldCrd:Add, Edit, x12 y80 w380 h20 Limit14 vCardNumber Number
		Gui, HldCrd:Add, Button, x12 y110 w90 h30 Default gHldCrdButtonOK, &OK
		Gui, HldCrd:Add, Button, x157 y110 w90 h30 Disabled gHldCrdButtonDone , &Done
		Gui, HldCrd:Add, Button, x302 y110 w90 h30 gHldCrdButtonCancel, &Cancel
		Gui, HldCrd:Show, x127 y87 h160 w404, Hold Shelf Card Number
		Return
	}
	else
	{
		GoTo, HoldShelfContinue
	}

	HldCrdGuiClose:
	HldCrdGuiEscape:
	HldCrdButtonCancel:
	HldCrdButtonDone:
		Gui, HldCrd:Destroy
		breakScript()

	HldCrdButtonOK:
		Gui, HldCrd:Submit
		Gui, HldCrd:Destroy
		;IniWrite, Value, Filename, Section, Key 
		IniWrite, %CardNumber%, %settings%, %Library%, HoldShelfCardNumber
		HSCN = %CardNumber%

	HoldShelfContinue:
	ControlDirectReader(0)

	;MsgBox, 1 , Check Item Out to the Hold Shelf, Automatically check item out to the hold shelf.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	;IfMsgBox Cancel
	;{
		; Turn DirectReader On
		;controlDirectReader(1)
		; Exit
		;Exit
	;}

	HoldShelfLabel:
	;********************************
	; Hold Shelf GUI
	;********************************
	Gui, HldShlf:Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically check the item out to the hold shelf.
	Gui, HldShlf:Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, HldShlf:Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, HldShlf:Add, Button, x12 y110 w90 h30 Default gHldShlfButtonOK, &OK
	Gui, HldShlf:Add, Button, x157 y110 w90 h30 gHldShlfButtonDone , &Done
	Gui, HldShlf:Add, Button, x302 y110 w90 h30 gHldShlfButtonCancel, &Cancel
	Gui, HldShlf:Show, x127 y87 h160 w404, Hold Shelf
	Return

	HldShlfGuiClose:
	HldShlfGuiEscape:
	HldShlfButtonDone:
	HldShlfButtonCancel:
		; Check if any items have been checked out to
		; the problem material card.  If any has, finish
		; the transaction, if not, just quit.
		Gui, HldShlf:Destroy
		;MsgBox Close/Done/Cancel.

		SetTitleMatchMode, Slow
		IfWinExist, Check Out, %HSCN%
		{
			ControlGet, OutputVar, List, , SysListView323, Check Out
			;MsgBox, %OutputVar%
			if (OutputVar != "")
			{
				WinActivate, CheckOut
				ControlFocus, Edit3, Check Out
				Send {Enter}
				;Polaris
				;ahk_class #32770
				WinWaitActive, Polaris, Do you want to print a check-out receipt?, 3
				Send n
			}
			Send !{F4}
		}
		SetTitleMatchMode, Fast

		breakScript()

	HldShlfButtonOK:
		; Check out to problem material and return for more.
		Gui, HldShlf:Submit  ; Save the input from the user to each control's associated variable.
		Gui, HldShlf:Destroy

		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183) {
			BlockInput, On

			WinWait, Check Out, ,1
			if ErrorLevel
			{
				ErrorLevel = 0
				WinActivate, Polaris ILS
				Sleep 50
				Send {F3}
				WinGet, active_id, ID, Check Out - 
				;Exit
				;MsgBox, Check Out window isn't open.
			}

			Sleep, 150

			SetTitleMatchMode, Slow
			IfWinNotExist, Check Out, %HSCN%
			{
				SetTitleMatchMode, Fast

				WinActivate
				ControlFocus, Edit1, Check Out
				SetKeyDelay, 0
				Send %HSCN%{Enter}
				SetKeyDelay, %keydelay%
				Sleep, 300
				IfWinActive, Patron Blocks
				{
					Send y
				}
				SetKeyDelay, 0
				Send %ItemBarcode%{Enter}
				SetKeyDelay, %keydelay%
			}
			else
			{
				SetTitleMatchMode, Fast

				WinActivate
				ControlFocus, Edit3, Check Out
				SetKeyDelay, 0
				Send %ItemBarcode%{Enter}
				SetKeyDelay, %keydelay%
			}
		}
		else
		{
			textToLog = Input (%ItemBarcode%) doesn't look correct.  Please try again. ItemBarcodeLength: %ItemBarcodeLength% (Should be 14) ItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)
			log(textToLog)
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
		}

		Sleep, 100
		SetTitleMatchMode, Slow
		IfWinExist, Polaris, Do you want to renew?
		{
			Send n
			MsgBox, Item already checked out to the hold shelf.
		}
		SetTitleMatchMode, Fast


		BlockInput, Off
		Goto, HoldShelfLabel
	Return
}

;********************************
;
; Check Item out to the Problem Material
; Control + Alt + p
;
;********************************

^!p::
{
	startHotkey()
	IniRead, PMCN, %settings%, %Library%, ProblemMaterialCardNumber, %A_Space%
	if (PMCN="")
	{
		;********************************
		; No Problem Material Card GUI
		;********************************
		Gui, PrblmCrd:Add, Text, x12 y10 w380 h30 , Problem Material Card Number
		Gui, PrblmCrd:Add, Text, x12 y30 w380 h30 , There is no card number associated with the problem material card for %library%.
		Gui, PrblmCrd:Add, Text, x12 y65 w380 h20 , Card Number:
		Gui, PrblmCrd:Add, Edit, x12 y80 w380 h20 Limit14 vCardNumber Number
		Gui, PrblmCrd:Add, Button, x12 y110 w90 h30 Default gPrblmCrdButtonOK, &OK
		Gui, PrblmCrd:Add, Button, x157 y110 w90 h30 Disabled gPrblmCrdButtonDone , &Done
		Gui, PrblmCrd:Add, Button, x302 y110 w90 h30 gPrblmCrdButtonCancel, &Cancel
		Gui, PrblmCrd:Show, x127 y87 h160 w404, Problem Material Card Number
		Return
	}
	else
	{
		GoTo, ProblemMaterialContinue
	}

	PrblmCrdGuiClose:
	PrblmCrdGuiEscape:
	PrblmCrdButtonCancel:
	PrblmCrdButtonDone:
		Gui, PrblmCrd:Destroy
		breakScript()

	PrblmCrdButtonOK:
		Gui, PrblmCrd:Submit
		Gui, PrblmCrd:Destroy
		;IniWrite, Value, Filename, Section, Key 
		IniWrite, %CardNumber%, %settings%, %Library%, ProblemMaterialCardNumber
		PMCN = %CardNumber%

	ProblemMaterialContinue:
	ControlDirectReader(0)

	;MsgBox, 1 , Check Item Out to the Problem Material Card, Automatically check item out to the problem material card.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	;IfMsgBox Cancel
	;{
		; Turn DirectReader On
		;controlDirectReader(1)
		; Exit
		;Exit
	;}

	ProblemMaterialLabel:
	;********************************
	; Problem Material GUI
	;********************************
	Gui, PrblmMtrl:Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically check the item out to the problem material card.
	Gui, PrblmMtrl:Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, PrblmMtrl:Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, PrblmMtrl:Add, Button, x12 y110 w90 h30 Default gPrblmMtrlButtonOK, &OK
	Gui, PrblmMtrl:Add, Button, x157 y110 w90 h30 gPrblmMtrlButtonDone , &Done
	Gui, PrblmMtrl:Add, Button, x302 y110 w90 h30 gPrblmMtrlButtonCancel, &Cancel
	Gui, PrblmMtrl:Show, x127 y87 h160 w404, Problem Shelf
	Return

	PrblmMtrlGuiClose:
	PrblmMtrlGuiEscape:
	PrblmMtrlButtonDone:
	PrblmMtrlButtonCancel:
		; Check if any items have been checked out to
		; the problem material card.  If any has, finish
		; the transaction, if not, just quit.
		Gui, PrblmMtrl:Destroy

		SetTitleMatchMode, Slow
		IfWinExist, Check Out, %PMCN%
		{
			ControlGet, OutputVar, List, , SysListView323, Check Out
			if (OutputVar != "")
			{
				WinActivate, CheckOut
				ControlFocus, Edit3, Check Out
				Send {Enter}
				WinWaitActive, Polaris, Do you want to print a check-out receipt?, 3
				Send n
			}
			Send !{F4}
		}
		SetTitleMatchMode, Fast

		breakScript()

	PrblmMtrlButtonOK:
		; Check out to problem material and return for more.
		Gui, PrblmMtrl:Submit  ; Save the input from the user to each control's associated variable.
		Gui, PrblmMtrl:Destroy

		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183) {
			BlockInput, On

			WinWait, Check Out, ,1
			if ErrorLevel
			{
				ErrorLevel = 0
				WinActivate, Polaris ILS
				Sleep 50
				Send {F3}
				WinGet, active_id, ID, Check Out - 
				;Exit
				;MsgBox, Check Out window isn't open.
			}

			Sleep, 150

			SetTitleMatchMode, Slow
			IfWinNotExist, Check Out, %PMCN%
			{
				SetTitleMatchMode, Fast

				WinActivate
				ControlFocus, Edit1, Check Out
				SetKeyDelay, 0
				Send %PMCN%{Enter}
				SetKeyDelay, %keydelay%
				Sleep, 300
				IfWinActive, Patron Blocks
				{
					Send y
				}
				SetKeyDelay, 0
				Send %ItemBarcode%{Enter}
				SetKeyDelay, %keydelay%
			}
			else
			{
				SetTitleMatchMode, Fast

				WinActivate
				ControlFocus, Edit3, Check Out
				SetKeyDelay, 0
				Send %ItemBarcode%{Enter}
				SetKeyDelay, %keydelay%
			}
		}
		else
		{
			textToLog = Input (%ItemBarcode%) doesn't look correct.  Please try again. ItemBarcodeLength: %ItemBarcodeLength% (Should be 14) ItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)
			log(textToLog)
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
		}

		Sleep, 100
		SetTitleMatchMode, Slow
		IfWinExist, Polaris, Do you want to renew?
		{
			Send n
			MsgBox, Item already checked out to the hold shelf.
		}
		SetTitleMatchMode, Fast


		BlockInput, Off
		Goto, ProblemMaterialLabel
	Return
}