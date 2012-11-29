#NoEnv
#SingleInstance force
keyDelay = 25
i = 0
OutputVar = ""
ItemRecordIterations = 10
controlDelay = 40
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%


;ControlGet, directReaderOn, Enabled, , OFF, CircControl-DirectReader
;MsgBox, %directReaderOn%
;Exit


;********************************
;*
;* # Windows Key
;* ! Alt
;* ^ Control
;* + Shift
;*
;********************************

;********************************
;*
;* Fuctions
;*
;********************************

;********
;* Turns DirectReader on or off
;* 0 for OFF, 1 for ON
;********
controlDirectReader(a=0)
{
	if(a=0)
		b=OFF
	else
		b=ON

	IfWinExist, CircControl-DirectReader
	{
		; Try to control DirectReader up to 5 times with a 100ms delay between tries.
		; Sometimes it doesn't turn change on the first try.
		Loop
		{
			if a_index > 5
			{
				StringLower, b_lower, b
				MsgBox, Couldn't turn %b_lower% Directreader.  Please turn it off manually.
				break  ; Terminate the loop
			}
			ControlGet, directReaderOn, Enabled, , %b%, CircControl-DirectReader
			if (directReaderOn = 1)
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

breakScript(type = 0)
{
	if (type = 0)
		Goto, skip
	else if (type = npu)
		title = Item Records - Barcode Find Tool
	else if (type = missing)
		title = Request Manager - Hold Requests
	else if (type if integer)
		title = ahk_id %active_id%

	IfWinExist, %title%
	{
		WinClose
	}
	skip:
	controlDirectReader(1)
	Exit
}

db_startTime()
{
	StartTime := A_TickCount
}
db_endTime()
{
	EndTime := A_TickCount
	ElapsedTime := (EndTime - StartTime)/1000
	MsgBox, Start Time: %StartTime%`nEnd Time: %EndTime%`n`nTotal Time: %ElapsedTime% seconds
}

;********************************
;
; Not Picked Up Holds
; Control + Alt + n
;
;********************************


IfMsgBox Cancel
{
	; Turn DirectReader On
	controlDirectReader(1)
	; Exit App
	ExitApp
}

^!n::
{
	MsgBox, 1 , Not Picked Up Holds, Automatically place the $1.00 NPU fee on the patron's card.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
	
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
		Send ^{F9}
		WinGet, active_id, ID, Item Records - Barcode Find Tool
	}

	; Activate Item Records window and focus on the barcode input box
	; @todo: make sure Item Records is expecting item barcode
	
	WinActivate, Item Records - .* Find Tool
	ControlFocus, ComboBox2, Item Records - .* Find Tool
	Send ba
	ControlFocus, Edit3, Item Records - .* Find Tool
	SetTitleMatchMode, 1

	; Display InputBox
	; @todo: Use GUI (see problemmaterial.ahk)
	; will require not insignificant rewriting of script
	;********************************
	; GUI
	;********************************
	Gui, Add, Text, x12 y10 w380 h30 , Scan the item barcode here.  It will automatically place the $1.00 "Not Picked Up" fee on the patron's card.
	Gui, Add, Text, x12 y50 w380 h20 , Item Barcode:
	Gui, Add, Edit, x12 y80 w380 h20 Limit14 vItemBarcode Number
	Gui, Add, Button, x12 y110 w90 h30 Default gButtonOK, &OK
	Gui, Add, Button, x157 y110 w90 h30 gButtonDone , &Done
	Gui, Add, Button, x302 y110 w90 h30 gButtonCancel, &Cancel
	Gui, Show, x127 y87 h160 w404, Not Picked Up
	Return

	; When OK button is clicked
	ButtonOK:
		;startTime()
		Gui, Submit  ; Save the input from the user to each control's associated variable.
		Gui, Destroy
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
			; until the item gets returned by Polaris
			retry:
			ControlGet, OutputVar, List, Selected Col9, SysListView321, Item Records
				if (ErrorLevel or OutputVar = "")
				{
					if (i < ItemRecordIterations)
					{
						i++
						Sleep, 250
						Goto, retry
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
					Goto, NPULabel
				}
			Sleep, 100
			Send {AppsKey}le
			; Waits up to three (3) seconds for the Patron Status window to launch and be active
			; If not, alerts the user and breaks the script
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
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
			Goto, NPULabel
		}

	; If the QUI is closed, escaped (escape key),
	; or the Done or Cancel button is pressed,
	; break out of the script.
	GuiClose:
	GuiEscape:
	ButtonDone:
	ButtonCancel:
		Gui, Destroy
		breakScript(%active_id%)
}
