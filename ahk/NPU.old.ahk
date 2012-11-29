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
		SetControlDelay -1
		ControlClick, %b%, CircControl-DirectReader
		SetControlDelay, 40
	}
}

breakScript()
{
	controlDirectReader(1)
	Exit
}

;********************************
;
; Not Picked Up Holds
; Control + Alt + n
;
;********************************

MsgBox, 1 , Not Picked Up Holds, Automatically place the $1.00 NPU fee on the patron's card.`n`nIMPORTANT: MAKE SURE DIRECTREADER IS OFF!!`n`nPress OK to continue.`n`nThis message box will timeout in 7 seconds., 7
IfMsgBox Cancel
{
	; Turn DirectReader On
	controlDirectReader(1)
	; Exit App
	ExitApp
}

^!n::
{

	; Turn Direct Reader Off
	controlDirectReader(0)
	
	NPULabel:
	
	; If Item Records window isn't open, open it
	WinWait, Item Records - Barcode Find Tool, ,0
	if ErrorLevel
	{
		ErrorLevel = 0
		WinActivate, Polaris ILS
		Sleep 50
		Send ^{F9}
	}
	
	; Activate Item Records window and focus on the barcode input box
	; @todo: make sure Item Records is expecting item barcode
	WinActivate, Item Records - Barcode Find Tool
	ControlFocus, Edit3, Item Records - Barcode Find Tool
	
	; Display InputBox
	; @todo: Use GUI (see problemmaterial.ahk)
	; will require not insignificant rewriting of script
	InputBox, ItemBarcode, Item Barcode, Scan the item barcode here.  It will automatically place the $1.00 "Not Picked Up" fee on the patron's card.

	if ErrorLevel
	{
		ErrorLevel = 0
		breakScript()
	}
	else
	{
		;StartTime := A_TickCount
		StringLen, ItemBarcodeLength, ItemBarcode
		StringMid, ItemBarcodeFirst, ItemBarcode, 1 ,5
		if (ItemBarcodeLength = 14 and ItemBarcodeFirst = 31183) {
			BlockInput, On
			ControlFocus, Edit3, Item Records - Barcode Find Tool
			SetkeyDelay, 0
			Send %ItemBarcode%{Enter}
			SetkeyDelay, %keyDelay%
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
			WinWaitActive, Patron Status, ,3
			if ErrorLevel
			{
				ErrorLevel = 0
				MsgBox, Patron Status windows didn't launch.
				breakScript()
			}
			Send {Alt}va^c
			WinWaitActive, Charge, ,3
			if ErrorLevel
			{
				ErrorLevel = 0
				MsgBox, Charge window didn't launch.
				breakScript()
			}
			ControlFocus, Edit1, Charge
			Send 1{Tab}n{tab}
			SetkeyDelay, 0
			Send %ItemBarcode%
			SetkeyDelay, %keyDelay%
			Send {Tab 3}
			;Send %ItemBarcode%{Enter}{AppsKey}le{Alt}va^c1{Tab}n{tab}%ItemBarcode%{Tab 3}
			;Send %ItemBarcode%{Enter}{AppsKey}le{Alt}va^c1{Tab}n{tab}%ItemBarcode%{Tab 3}{Enter}
			BlockInput, Off
			;EndTime := A_TickCount
			;ElapsedTime := (EndTime - StartTime)/1000
			;MsgBox, Start Time: %StartTime%`nEnd Time: %EndTime%`n`nTotal Time: %ElapsedTime% seconds
			WinWaitClose, Charge
			WinWaitClose, Patron Status
			Goto, NPULabel
		}
		else
		{
			MsgBox, Input (%ItemBarcode%) doesn't look correct.  Please try again.`n`nItemBarcodeLength: %ItemBarcodeLength% (Should be 14)`nItemBarcodeFirst: %ItemBarcodeFirst% (Should be 31183)`n`nCheck the input and try again.
			Goto, NPULabel
		}
	}
	breakScript()
}

